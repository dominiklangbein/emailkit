import {
  EmailKit,
  type EmailAddress,
  type EmailDriverTuple,
  type EmailKitHooks,
} from "emailkit"

import {
  configuredMailboxes,
  defaultEmailDriver,
  readyEmailDrivers,
  type SandboxEmailDriver,
} from "./sandbox/drivers"
import {
  findPersistedMailbox,
  persistMailbox,
} from "./sandbox/persistence/mailboxes"
import {
  deleteWebhookByProviderId,
  persistWebhookFromHook,
} from "./sandbox/persistence/webhooks"
import {
  findRememberedDomain,
  findRememberedMailbox,
  recordSandboxEvent,
  rememberDomain,
  rememberMailbox,
} from "./sandbox/store"

const resolveSender = async ({ from }: { from: EmailAddress }) => {
  const email = from.email.toLowerCase()
  const remembered = findRememberedMailbox(email)
  if (remembered) {
    return {
      emailDriver: remembered.emailDriver,
      mailbox: remembered.mailbox,
      ...(remembered.auth !== undefined ? { auth: remembered.auth } : {}),
    }
  }

  const persisted = await findPersistedMailbox(email)
  if (persisted) {
    rememberMailbox(persisted.emailDriver, persisted.mailbox, persisted.auth)
    return {
      emailDriver: persisted.emailDriver,
      mailbox: persisted.mailbox,
      ...(persisted.auth !== undefined ? { auth: persisted.auth } : {}),
    }
  }

  const appDriver = readyEmailDrivers.find((driver) =>
    configuredMailboxes(driver.id).some(
      (mailbox) => mailbox.email.toLowerCase() === email
    )
  )
  if (appDriver) return { emailDriver: appDriver.id }

  const domainName = email.split("@")[1]
  const domain = domainName ? findRememberedDomain(domainName) : undefined
  return { emailDriver: domain?.emailDriver ?? defaultEmailDriver() }
}

const hooks: EmailKitHooks = {
  domain: {
    onCreated: async ({ emailDriver, domain }) => {
      rememberDomain(emailDriver, domain)
      recordSandboxEvent({
        driver: emailDriver,
        category: "domain",
        kind: "domain.created",
        summary: `Domain created: ${domain.domain}`,
        details: domain,
      })
    },
    onVerified: async ({ emailDriver, domain }) => {
      rememberDomain(emailDriver, domain)
      recordSandboxEvent({
        driver: emailDriver,
        category: "domain",
        kind: "domain.verified",
        summary: `Domain verified: ${domain.domain}`,
        details: domain,
      })
    },
    onDeleted: async ({ emailDriver, domain }) => {
      recordSandboxEvent({
        driver: emailDriver,
        category: "domain",
        kind: "domain.deleted",
        summary: `Domain deleted: ${domain.domain}`,
        details: domain,
      })
    },
  },
  mailbox: {
    onConnected: async ({ emailDriver, mailbox, auth }) => {
      rememberMailbox(emailDriver, mailbox, auth)
      await persistMailbox(emailDriver, mailbox, auth)
      recordSandboxEvent({
        driver: emailDriver,
        category: "mailbox",
        kind: "mailbox.connected",
        summary: `Mailbox connected: ${mailbox.email}`,
        details: { mailbox },
      })
    },
    onCreated: async ({ emailDriver, mailbox }) => {
      rememberMailbox(emailDriver, mailbox)
      recordSandboxEvent({
        driver: emailDriver,
        category: "mailbox",
        kind: "mailbox.created",
        summary: `Mailbox created: ${mailbox.email}`,
        details: mailbox,
      })
    },
    onDeleted: async ({ emailDriver, mailbox }) => {
      recordSandboxEvent({
        driver: emailDriver,
        category: "mailbox",
        kind: "mailbox.deleted",
        summary: `Mailbox deleted: ${mailbox.email}`,
        details: mailbox,
      })
    },
    onAuthUpdated: async ({ emailDriver, mailbox, auth }) => {
      if (!mailbox.email) {
        throw new Error(
          "Cannot update sandbox mailbox auth without a mailbox email."
        )
      }

      const existing =
        findRememberedMailbox(mailbox.email)?.mailbox ??
        (await findPersistedMailbox(mailbox.email, emailDriver))?.mailbox
      const updatedMailbox = {
        ...existing,
        ...mailbox,
        id: mailbox.id ?? existing?.id ?? mailbox.email,
        email: mailbox.email,
      }

      rememberMailbox(emailDriver, updatedMailbox, auth)
      await persistMailbox(emailDriver, updatedMailbox, auth)
      recordSandboxEvent({
        driver: emailDriver,
        category: "mailbox",
        kind: "mailbox.auth.updated",
        summary: `Mailbox auth updated: ${updatedMailbox.email}`,
        details: { mailbox: updatedMailbox },
      })
    },
  },
  email: {
    onAll: async ({ emailDriver, type, data }) => {
      if (type === "opened" || type === "clicked") return

      recordSandboxEvent({
        driver: emailDriver ?? "emailkit",
        category: "hook",
        kind: `email.${type}`,
        summary: `Email hook: ${type}`,
        details: data,
      })
    },
    onOpened: async (event) => {
      recordSandboxEvent({
        driver: event.emailDriver ?? "emailkit",
        category: "hook",
        kind: "email.opened",
        summary: "Email hook: opened",
        details: event,
      })
    },
    onClicked: async (event) => {
      recordSandboxEvent({
        driver: event.emailDriver ?? "emailkit",
        category: "hook",
        kind: "email.clicked",
        summary: "Email hook: clicked",
        details: event,
      })
    },
  },
  webhook: {
    onCreated: async (event) => {
      await persistWebhookFromHook(event)
      recordSandboxEvent({
        driver: event.emailDriver,
        category: "hook",
        kind: "webhook.created",
        summary: `Webhook created: ${event.scope} · ${event.webhook.url}`,
        details: event,
      })
    },
    onUpdated: async (event) => {
      await persistWebhookFromHook(event)
      recordSandboxEvent({
        driver: event.emailDriver,
        category: "hook",
        kind: "webhook.updated",
        summary: `Webhook updated: ${event.scope} · ${event.webhook.providerId ?? event.webhook.id}`,
        details: event,
      })
    },
    onDeleted: async (event) => {
      const providerId = event.webhook?.providerId ?? event.providerId
      if (providerId) {
        await deleteWebhookByProviderId({
          emailDriver: event.emailDriver,
          scope: event.scope,
          providerId,
        })
      }
      recordSandboxEvent({
        driver: event.emailDriver,
        category: "hook",
        kind: "webhook.deleted",
        summary: `Webhook deleted: ${event.scope}`,
        details: event,
      })
    },
    onActionRequired: async (event) => {
      recordSandboxEvent({
        driver: event.emailDriver,
        category: "hook",
        kind: "webhook.action_required",
        summary: `Webhook action required: ${event.scope} · ${event.reason}`,
        details: event,
      })
    },
    onSyncRequired: async (event) => {
      recordSandboxEvent({
        driver: event.emailDriver,
        category: "hook",
        kind: "webhook.sync_required",
        summary: `Webhook sync required: ${event.scope} · ${event.reason}`,
        details: event,
      })
    },
  },
}

const isEmailDriverTuple = (
  drivers: readonly SandboxEmailDriver[]
): drivers is EmailDriverTuple => drivers.length > 0

export const emailkit = !isEmailDriverTuple(readyEmailDrivers)
  ? null
  : EmailKit({
      emailDrivers: readyEmailDrivers,
      resolveEmailDriver: async (context) => {
        if (context.operation === "sendEmail") {
          return resolveSender(context.message)
        }
        return { emailDriver: defaultEmailDriver() }
      },
      hooks,
    })
