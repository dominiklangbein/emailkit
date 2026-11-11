import type {
  Attachment,
  Domain,
  DomainIdentifier,
  Mailbox,
  SyncResult,
  Webhook,
  WebhookRequest,
  WebhookResponse,
} from "emailkit"
import { EmailKitSyncError } from "emailkit"

import { emailkit } from "../emailkit"
import {
  configuredMailboxes,
  getSandboxDriverInfo,
  hasSandboxDriver,
} from "./drivers"
import {
  deletePersistedMailbox,
  listPersistedMailboxes,
} from "./persistence/mailboxes"
import {
  deleteWebhookByProviderId,
  findWebhookByRowId,
  listExpiringWebhooks,
  listWebhooks,
  upsertWebhook,
  type SandboxWebhook,
} from "./persistence/webhooks"
import {
  recordSandboxEvent,
  rememberDomain,
  rememberMailbox,
  runSandboxTrace,
} from "./store"
import type {
  AttachmentInput,
  DomainActionInput,
  MailboxActionInput,
  ProviderFetchInput,
  SandboxSyncResult,
  SendSandboxEmailInput,
  SandboxWebhookView,
  SyncActionInput,
  WebhookActionInput,
} from "./types"

class SandboxHttpError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly body?: Record<string, unknown>
  ) {
    super(message)
  }
}

const requireEmailKit = () => {
  if (!emailkit) {
    throw new SandboxHttpError(
      "No EmailKit drivers are configured. Add at least one provider API key.",
      503
    )
  }
  return emailkit
}

const requireDriver = (emailDriver: string) => {
  if (!hasSandboxDriver(emailDriver))
    throw new Error(`Unknown email driver: ${emailDriver}`)
  return emailDriver
}

const requireReadyDriver = (emailDriver: string) => {
  requireDriver(emailDriver)
  const info = getSandboxDriverInfo(emailDriver)
  if (!info?.ready) {
    throw new SandboxHttpError(
      `${info?.label ?? emailDriver} is missing required environment variables.`,
      503,
      {
        missingRequiredEnv: info?.missingRequiredEnv ?? [],
      }
    )
  }
  return emailDriver
}

const jsonHeaders = { "content-type": "application/json" }

export const toErrorMessage = (error: unknown) =>
  error instanceof Error ? error.message : "Unknown sandbox error"

const statusFromError = (error: unknown) => {
  if (error instanceof SandboxHttpError) return error.status
  if (error && typeof error === "object") {
    const status = (error as { status?: unknown; httpStatus?: unknown }).status
    const httpStatus = (error as { status?: unknown; httpStatus?: unknown })
      .httpStatus
    if (typeof status === "number") return status
    if (typeof httpStatus === "number") return httpStatus
  }
  return 400
}

export const jsonOk = (body: Record<string, unknown> = {}) =>
  Response.json({ ok: true, ...body }, { headers: jsonHeaders })

export const jsonError = (error: unknown, status = statusFromError(error)) =>
  Response.json(
    {
      ok: false,
      error: toErrorMessage(error),
      ...(error instanceof SandboxHttpError ? error.body : {}),
    },
    { status, headers: jsonHeaders }
  )

export const sendSandboxEmail = async (input: SendSandboxEmailInput) => {
  const client = requireEmailKit()
  const emailDriver = requireReadyDriver(input.emailDriver)

  const message = {
    from: { email: input.fromEmail, name: input.fromName || undefined },
    to: { email: input.toEmail },
    cc: input.ccEmail ? { email: input.ccEmail } : undefined,
    bcc: input.bccEmail ? { email: input.bccEmail } : undefined,
    reply:
      input.replyToEmail || input.inReplyToMessageId
        ? {
            addresses: input.replyToEmail
              ? [{ email: input.replyToEmail }]
              : undefined,
            messageId: input.inReplyToMessageId,
          }
        : undefined,
    subject: input.subject,
    text: input.text || undefined,
    html: input.html || undefined,
    sendAt: input.sendAt ? new Date(input.sendAt) : undefined,
    templateId: input.templateId || undefined,
    templateData: input.templateData,
    track: input.track,
    unsubscribe: input.unsubscribe,
    idempotencyKey: input.idempotencyKey,
    tenantId: input.tenantId,
    tags: input.tags,
    metadata: input.metadata,
    headers: input.headers,
    provider: input.provider,
  }

  const result = await client.sendEmail(message)
  recordSandboxEvent({
    driver: emailDriver,
    category: "send",
    kind: "email.send",
    summary: `Sent: ${input.subject}`,
    details: { ...message, result },
  })
  return result
}

type DomainsMethods = {
  list: (input: Record<string, unknown>) => Promise<Domain[]>
  create: (input: Record<string, unknown>) => Promise<Domain>
  ensure: (input: Record<string, unknown>) => Promise<unknown>
  get: (input: Record<string, unknown>) => Promise<Domain>
  update: (input: Record<string, unknown>, patch: unknown) => Promise<Domain>
  verify: (input: Record<string, unknown>) => Promise<unknown>
  delete: (input: Record<string, unknown>) => Promise<unknown>
}

export const handleDomainAction = async (input: DomainActionInput) => {
  const client = requireEmailKit()
  const emailDriver = requireReadyDriver(input.emailDriver)
  // readyEmailDrivers is widened to EmailDriver[] which collapses the conditional
  // DomainsFacade method types. Driver readiness/capabilities are validated at runtime.
  const domains = client.domains as unknown as DomainsMethods

  switch (input.action) {
    case "list": {
      const result = await domains.list({ ...input.options, emailDriver })
      result.forEach((domain) => rememberDomain(emailDriver, domain))
      return { domains: result }
    }
    case "create": {
      const domain = await domains.create({ ...input.input, emailDriver })
      return { domain }
    }
    case "ensure": {
      const result = await domains.ensure({ ...input.input, emailDriver })
      return { result }
    }
    case "get": {
      const domain = await domains.get({ ...input.identifier, emailDriver })
      rememberDomain(emailDriver, domain)
      return { domain }
    }
    case "update": {
      const domain = await domains.update(
        { ...input.identifier, emailDriver },
        input.patch
      )
      rememberDomain(emailDriver, domain)
      return { domain }
    }
    case "verify": {
      const verification = await domains.verify({
        ...input.identifier,
        emailDriver,
      })
      return { verification }
    }
    case "delete": {
      const result = await domains.delete({ ...input.identifier, emailDriver })
      return { result }
    }
  }
}

export const handleMailboxAction = async (input: MailboxActionInput) => {
  const client = requireEmailKit()
  const emailDriver = requireReadyDriver(input.emailDriver)

  const caps = getSandboxDriverInfo(emailDriver)?.capabilities ?? {}

  switch (input.action) {
    case "list": {
      const driverMailboxes = caps.mailboxList
        ? await client.mailboxes.list({ ...input.options, emailDriver })
        : configuredMailboxes(emailDriver)
      driverMailboxes.forEach((mailbox) =>
        rememberMailbox(emailDriver, mailbox)
      )

      const persisted = await listPersistedMailboxes(emailDriver)
      const seen = new Set<string>()
      const mailboxes: Mailbox[] = []
      for (const mailbox of driverMailboxes) {
        const key = mailbox.email.toLowerCase()
        if (seen.has(key)) continue
        seen.add(key)
        mailboxes.push(mailbox)
      }
      for (const entry of persisted) {
        const key = entry.mailbox.email.toLowerCase()
        if (seen.has(key)) continue
        seen.add(key)
        mailboxes.push(entry.mailbox)
        rememberMailbox(emailDriver, entry.mailbox, entry.auth)
      }
      return { mailboxes }
    }
    case "connect": {
      const result = await client.mailboxes.connect({
        ...input.input,
        emailDriver,
      })
      if (result.mailbox) rememberMailbox(emailDriver, result.mailbox)
      return { result }
    }
    case "create": {
      const mailbox = await client.mailboxes.create({
        ...input.input,
        emailDriver,
      })
      rememberMailbox(emailDriver, mailbox)
      return { mailbox }
    }
    case "get": {
      const mailbox = await client.mailboxes.get({
        idOrEmail: input.idOrEmail,
        emailDriver,
      })
      rememberMailbox(emailDriver, mailbox)
      return { mailbox }
    }
    case "delete": {
      const driverResult = caps.mailboxDelete
        ? await client.mailboxes.delete({
            idOrEmail: input.idOrEmail,
            emailDriver,
          })
        : undefined
      const removed = await deletePersistedMailbox(emailDriver, input.idOrEmail)
      return { result: driverResult ?? { ok: removed } }
    }
  }
}

export const runProviderFetch = async ({
  emailDriver,
  path,
  init,
}: ProviderFetchInput) => {
  requireReadyDriver(emailDriver)
  const response = await requireEmailKit().providerFetch(path, {
    ...init,
    emailDriver,
  })
  const contentType = response.headers.get("content-type") ?? ""
  const body = contentType.includes("application/json")
    ? await response.json()
    : await response.text()
  const result = {
    ok: response.ok,
    status: response.status,
    statusText: response.statusText,
    headers: Object.fromEntries(response.headers),
    body,
  }
  recordSandboxEvent({
    driver: emailDriver,
    category: "tool",
    kind: "provider.fetch",
    summary: `${init?.method ?? "GET"} ${path}`,
    details: result,
  })
  return result
}

export const getAttachmentContent = async ({
  emailDriver,
  attachment,
}: AttachmentInput) => {
  requireReadyDriver(emailDriver)
  const content = await requireEmailKit().attachments.getContent(
    attachment as Attachment,
    { emailDriver }
  )
  const result =
    typeof content === "string"
      ? { type: "string", content }
      : {
          type: "bytes",
          byteLength: content.byteLength,
          preview: Array.from(content.slice(0, 80)),
        }

  recordSandboxEvent({
    driver: emailDriver,
    category: "tool",
    kind: "attachment.getContent",
    summary: `Fetched attachment: ${attachment.filename}`,
    details: result,
  })
  return result
}

export const handleSandboxWebhook = async (
  emailDriver: string,
  request: WebhookRequest
): Promise<WebhookResponse> => {
  const info = getSandboxDriverInfo(emailDriver)
  if (!info) {
    return {
      status: 404,
      body: { error: `Unknown email driver: ${emailDriver}` },
    }
  }
  if (!info.ready) {
    return {
      status: 503,
      body: {
        error: `${info.label} is not configured`,
        missingRequiredEnv: info.missingRequiredEnv,
      },
    }
  }
  const handler = requireEmailKit().handler()
  const requestForDriver = {
    ...request,
    query: { ...request.query, emailDriver },
    headers: { ...request.headers, "x-emailkit-driver": emailDriver },
  }

  return runSandboxTrace(emailDriver, async () => {
    recordSandboxEvent({
      driver: emailDriver,
      category: "webhook",
      kind: "webhook.request",
      summary: `${emailDriver} webhook request`,
      details: requestForDriver,
    })
    const response = await handler(requestForDriver)
    recordSandboxEvent({
      driver: emailDriver,
      category: "webhook",
      kind: "webhook.response",
      summary: `${emailDriver} webhook response ${response.status}`,
      details: response,
    })
    return response
  })
}

const toWebhookView = (webhook: SandboxWebhook): SandboxWebhookView => ({
  rowId: webhook.rowId,
  emailDriver: webhook.emailDriver ?? "",
  scope: webhook.scope,
  providerId: webhook.providerId,
  url: webhook.url,
  events: webhook.events,
  status: webhook.status,
  expiresAt: webhook.expiresAt?.toISOString(),
  renewAfter: webhook.renewAfter?.toISOString(),
  mailboxEmail: webhook.mailboxEmail,
  domain: webhook.domain,
  createdAt: webhook.createdAt?.toISOString(),
  updatedAt: webhook.updatedAt?.toISOString(),
})

const requireWebhookRow = async (rowId: string, emailDriver: string) => {
  const webhook = await findWebhookByRowId(rowId)
  if (!webhook) throw new SandboxHttpError(`Webhook ${rowId} not found`, 404)
  if (webhook.emailDriver !== emailDriver) {
    throw new SandboxHttpError(
      `Webhook ${rowId} belongs to a different driver`,
      400
    )
  }
  return webhook
}

const reattachContext = (
  webhook: SandboxWebhook,
  mailboxAuth?: unknown
): Record<string, unknown> | undefined => {
  switch (webhook.scope) {
    case "mailbox":
      if (!webhook.mailboxEmail) return undefined
      return {
        mailbox: { email: webhook.mailboxEmail },
        ...(mailboxAuth ? { auth: mailboxAuth } : {}),
      }
    case "domain":
      if (!webhook.domain) return undefined
      return { domain: webhook.domain }
    default:
      return undefined
  }
}

const findMailboxAuth = async (emailDriver: string, email: string) => {
  const mailboxes = await listPersistedMailboxes(emailDriver)
  return mailboxes.find(
    (entry) => entry.mailbox.email.toLowerCase() === email.toLowerCase()
  )?.auth
}

const setupWebhook = async (
  client: NonNullable<typeof emailkit>,
  emailDriver: string,
  input: Extract<WebhookActionInput, { action: "setup" }>
) => {
  const url = input.url?.trim()
  const events = input.events

  if (input.target.scope === "account") {
    const accountWebhooks = (
      client as unknown as {
        webhooks?: {
          setup: (
            input: Record<string, unknown>
          ) => Promise<{
            webhook?: Webhook
            deleted?: boolean
            context?: unknown
            raw?: unknown
          }>
        }
      }
    ).webhooks
    if (!accountWebhooks)
      throw new SandboxHttpError("Driver does not expose account webhooks", 400)
    return accountWebhooks.setup({
      emailDriver,
      ...(url ? { url } : {}),
      events,
    })
  }

  if (input.target.scope === "mailbox") {
    const auth = await findMailboxAuth(emailDriver, input.target.mailboxEmail)
    const mailboxWebhooks = (
      client.mailboxes as unknown as {
        webhooks?: {
          setup: (
            input: Record<string, unknown>
          ) => Promise<{
            webhook?: Webhook
            deleted?: boolean
            context?: unknown
            raw?: unknown
          }>
        }
      }
    ).webhooks
    if (!mailboxWebhooks)
      throw new SandboxHttpError("Driver does not expose mailbox webhooks", 400)
    return mailboxWebhooks.setup({
      emailDriver,
      ...(url ? { url } : {}),
      events,
      mailbox: { email: input.target.mailboxEmail },
      ...(auth ? { auth } : {}),
    })
  }

  const domainWebhooks = (
    client.domains as unknown as {
      webhooks?: {
        setup: (
          input: Record<string, unknown>
        ) => Promise<{
          webhook?: Webhook
          deleted?: boolean
          context?: unknown
          raw?: unknown
        }>
      }
    }
  ).webhooks
  if (!domainWebhooks)
    throw new SandboxHttpError("Driver does not expose domain webhooks", 400)
  return domainWebhooks.setup({
    emailDriver,
    ...(url ? { url } : {}),
    events,
    domain: input.target.domain,
  })
}

const refreshWebhook = async (
  client: NonNullable<typeof emailkit>,
  emailDriver: string,
  webhook: SandboxWebhook
) => {
  if (!webhook.providerId) {
    throw new SandboxHttpError(
      "Webhook has no providerId; refresh requires the provider id",
      400
    )
  }
  const ref = { webhookId: webhook.providerId, providerId: webhook.providerId }

  if (webhook.scope === "account") {
    const accountWebhooks = (
      client as unknown as {
        webhooks?: {
          refresh: (
            input: Record<string, unknown>
          ) => Promise<{
            webhook?: Webhook
            deleted?: boolean
            context?: unknown
            raw?: unknown
          }>
        }
      }
    ).webhooks
    if (!accountWebhooks)
      throw new SandboxHttpError("Driver does not expose account webhooks", 400)
    return accountWebhooks.refresh({ emailDriver, ...ref })
  }

  if (webhook.scope === "mailbox") {
    const auth = webhook.mailboxEmail
      ? await findMailboxAuth(emailDriver, webhook.mailboxEmail)
      : undefined
    const target = reattachContext(webhook, auth)
    if (!target)
      throw new SandboxHttpError(
        "Mailbox-scoped webhook is missing its mailbox",
        400
      )
    const mailboxWebhooks = (
      client.mailboxes as unknown as {
        webhooks?: {
          refresh: (
            input: Record<string, unknown>
          ) => Promise<{
            webhook?: Webhook
            deleted?: boolean
            context?: unknown
            raw?: unknown
          }>
        }
      }
    ).webhooks
    if (!mailboxWebhooks)
      throw new SandboxHttpError("Driver does not expose mailbox webhooks", 400)
    return mailboxWebhooks.refresh({ emailDriver, ...ref, ...target })
  }

  const target = reattachContext(webhook)
  if (!target)
    throw new SandboxHttpError(
      "Domain-scoped webhook is missing its domain",
      400
    )
  const domainWebhooks = (
    client.domains as unknown as {
      webhooks?: {
        refresh: (
          input: Record<string, unknown>
        ) => Promise<{
          webhook?: Webhook
          deleted?: boolean
          context?: unknown
          raw?: unknown
        }>
      }
    }
  ).webhooks
  if (!domainWebhooks)
    throw new SandboxHttpError("Driver does not expose domain webhooks", 400)
  return domainWebhooks.refresh({ emailDriver, ...ref, ...target })
}

const deleteWebhook = async (
  client: NonNullable<typeof emailkit>,
  emailDriver: string,
  webhook: SandboxWebhook
) => {
  if (!webhook.providerId) {
    await deleteWebhookByProviderId({
      emailDriver,
      scope: webhook.scope,
      providerId: webhook.id,
    })
    return { deleted: true }
  }
  const ref = { webhookId: webhook.providerId, providerId: webhook.providerId }

  if (webhook.scope === "account") {
    const accountWebhooks = (
      client as unknown as {
        webhooks?: {
          delete: (
            input: Record<string, unknown>
          ) => Promise<{
            webhook?: Webhook
            deleted?: boolean
            context?: unknown
            raw?: unknown
          }>
        }
      }
    ).webhooks
    if (!accountWebhooks)
      throw new SandboxHttpError("Driver does not expose account webhooks", 400)
    return accountWebhooks.delete({ emailDriver, ...ref })
  }

  if (webhook.scope === "mailbox") {
    const auth = webhook.mailboxEmail
      ? await findMailboxAuth(emailDriver, webhook.mailboxEmail)
      : undefined
    const target = reattachContext(webhook, auth)
    const mailboxWebhooks = (
      client.mailboxes as unknown as {
        webhooks?: {
          delete: (
            input: Record<string, unknown>
          ) => Promise<{
            webhook?: Webhook
            deleted?: boolean
            context?: unknown
            raw?: unknown
          }>
        }
      }
    ).webhooks
    if (!mailboxWebhooks)
      throw new SandboxHttpError("Driver does not expose mailbox webhooks", 400)
    return mailboxWebhooks.delete({ emailDriver, ...ref, ...(target ?? {}) })
  }

  const target = reattachContext(webhook)
  const domainWebhooks = (
    client.domains as unknown as {
      webhooks?: {
        delete: (
          input: Record<string, unknown>
        ) => Promise<{
          webhook?: Webhook
          deleted?: boolean
          context?: unknown
          raw?: unknown
        }>
      }
    }
  ).webhooks
  if (!domainWebhooks)
    throw new SandboxHttpError("Driver does not expose domain webhooks", 400)
  return domainWebhooks.delete({ emailDriver, ...ref, ...(target ?? {}) })
}

export const handleWebhookAction = async (input: WebhookActionInput) => {
  const client = requireEmailKit()
  const emailDriver = requireReadyDriver(input.emailDriver)

  switch (input.action) {
    case "list": {
      const webhooks = await listWebhooks(emailDriver)
      return { webhooks: webhooks.map(toWebhookView) }
    }
    case "setup": {
      const result = await setupWebhook(client, emailDriver, input)
      if (result?.webhook) {
        await upsertWebhook({
          emailDriver,
          webhook: { ...result.webhook, emailDriver },
          mailboxEmail:
            input.target.scope === "mailbox"
              ? input.target.mailboxEmail
              : undefined,
          domain:
            input.target.scope === "domain" ? input.target.domain : undefined,
        })
      }
      return {
        result,
        webhooks: (await listWebhooks(emailDriver)).map(toWebhookView),
      }
    }
    case "refresh": {
      const webhook = await requireWebhookRow(input.rowId, emailDriver)
      const result = await refreshWebhook(client, emailDriver, webhook)
      if (result?.webhook) {
        await upsertWebhook({
          emailDriver,
          webhook: { ...result.webhook, emailDriver },
          mailboxEmail: webhook.mailboxEmail,
          domain: webhook.domain,
        })
      }
      return {
        result,
        webhooks: (await listWebhooks(emailDriver)).map(toWebhookView),
      }
    }
    case "delete": {
      const webhook = await requireWebhookRow(input.rowId, emailDriver)
      const result = await deleteWebhook(client, emailDriver, webhook)
      if (result?.deleted && webhook.providerId) {
        await deleteWebhookByProviderId({
          emailDriver,
          scope: webhook.scope,
          providerId: webhook.providerId,
        })
      }
      return {
        result,
        webhooks: (await listWebhooks(emailDriver)).map(toWebhookView),
      }
    }
    case "renew-expiring": {
      const expiring = await listExpiringWebhooks(emailDriver)
      const results: { rowId: string; ok: boolean; error?: string }[] = []
      for (const webhook of expiring) {
        try {
          const result = await refreshWebhook(client, emailDriver, webhook)
          if (result?.webhook) {
            await upsertWebhook({
              emailDriver,
              webhook: { ...result.webhook, emailDriver },
              mailboxEmail: webhook.mailboxEmail,
              domain: webhook.domain,
            })
          }
          results.push({ rowId: webhook.rowId, ok: true })
        } catch (error) {
          results.push({
            rowId: webhook.rowId,
            ok: false,
            error: toErrorMessage(error),
          })
        }
      }
      return {
        results,
        webhooks: (await listWebhooks(emailDriver)).map(toWebhookView),
      }
    }
  }
}

type SyncMethod = (input: Record<string, unknown>) => Promise<SyncResult>

const runSync = async (
  client: NonNullable<typeof emailkit>,
  emailDriver: string,
  input: SyncActionInput,
  base: Record<string, unknown>
): Promise<SyncResult> => {
  if (input.target.scope === "account") {
    const accountSync = (client as unknown as { sync?: SyncMethod }).sync
    if (!accountSync)
      throw new SandboxHttpError("Driver does not support account sync", 400)
    return accountSync(base)
  }

  if (input.target.scope === "mailbox") {
    const mailboxSync = (client.mailboxes as unknown as { sync?: SyncMethod })
      .sync
    if (!mailboxSync)
      throw new SandboxHttpError("Driver does not support mailbox sync", 400)
    const auth = await findMailboxAuth(emailDriver, input.target.mailboxEmail)
    return mailboxSync({
      ...base,
      mailbox: { email: input.target.mailboxEmail },
      ...(auth ? { auth } : {}),
    })
  }

  const domainSync = (client.domains as unknown as { sync?: SyncMethod }).sync
  if (!domainSync)
    throw new SandboxHttpError("Driver does not support domain sync", 400)
  return domainSync({ ...base, domain: input.target.domain })
}

export const handleSyncAction = async (
  input: SyncActionInput
): Promise<SandboxSyncResult> => {
  const client = requireEmailKit()
  const emailDriver = requireReadyDriver(input.emailDriver)
  const scope = input.target.scope

  const since = new Date(input.since)
  if (Number.isNaN(since.getTime())) throw new Error("Invalid 'since' timestamp.")
  const until = input.until ? new Date(input.until) : undefined
  if (until && Number.isNaN(until.getTime()))
    throw new Error("Invalid 'until' timestamp.")

  const base: Record<string, unknown> = {
    emailDriver,
    since,
    ...(until ? { until } : {}),
    ...(input.context !== undefined ? { context: input.context } : {}),
  }

  recordSandboxEvent({
    driver: emailDriver,
    category: "system",
    kind: "sync.start",
    summary: `Sync started: ${scope} since ${since.toISOString()}`,
    details: { ...base, target: input.target },
  })

  try {
    const result = await runSync(client, emailDriver, input, base)
    recordSandboxEvent({
      driver: emailDriver,
      category: "system",
      kind: "sync.complete",
      summary: `Sync complete: ${result.dispatched} event${result.dispatched === 1 ? "" : "s"} replayed`,
      details: result,
    })
    return {
      ok: true,
      scope,
      dispatched: result.dispatched,
      syncedFrom: result.syncedFrom?.toISOString(),
    }
  } catch (error) {
    if (error instanceof EmailKitSyncError) {
      recordSandboxEvent({
        driver: emailDriver,
        category: "system",
        kind: "sync.error",
        summary: `Sync failed after ${error.dispatched} dispatched`,
        details: {
          dispatched: error.dispatched,
          lastEventTimestamp: error.lastEventTimestamp,
          message: error.message,
        },
      })
      return {
        ok: false,
        scope,
        dispatched: error.dispatched,
        lastEventTimestamp: error.lastEventTimestamp?.toISOString(),
        error: error.message,
      }
    }
    throw error
  }
}

export const parseDomainIdentifier = (
  body: Record<string, unknown>
): DomainIdentifier => {
  const domain =
    typeof body.domain === "string" && body.domain.trim()
      ? body.domain.trim()
      : undefined
  const domainId =
    typeof body.domainId === "string" && body.domainId.trim()
      ? body.domainId.trim()
      : undefined
  if (domain && domainId) return { domain, domainId }
  if (domain) return { domain }
  if (domainId) return { domainId }
  throw new Error("domain or domainId is required.")
}
