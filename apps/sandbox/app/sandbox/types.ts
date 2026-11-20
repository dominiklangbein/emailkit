import type {
  ConnectMailboxInput,
  CreateDomainInput,
  CreateMailboxInput,
  Domain,
  DomainIdentifier,
  DriverCapabilities,
  ListDomainsOptions,
  ListMailboxesOptions,
  Mailbox,
  ProviderFetchInit,
  UpdateDomainInput,
  WebhookEventSelection,
  WebhookScope,
} from "emailkit"

export type SandboxDriverId = string

export type SandboxDriverInfo = {
  id: SandboxDriverId
  label: string
  family: "resend" | "mailgun" | "aiinbx" | "outlook" | "gmail"
  publicWebhookUrl: string
  requiredEnv: string[]
  optionalEnv: string[]
  missingRequiredEnv: string[]
  missingOptionalEnv: string[]
  ready: boolean
  defaultFromEmail: string
  defaultToEmail: string
  capabilities: DriverCapabilities
}

export type SandboxEventCategory =
  | "send"
  | "webhook"
  | "hook"
  | "domain"
  | "mailbox"
  | "tool"
  | "system"

export type SandboxEvent = {
  id: string
  driver: SandboxDriverId
  category: SandboxEventCategory
  kind: string
  summary: string
  timestamp: string
  details?: unknown
}

export type SandboxTrace = {
  id: string
  driver: SandboxDriverId
  summary: string
  startedAt: string
  updatedAt: string
  events: SandboxEvent[]
  correlation: {
    eventId?: string
    messageId?: string
    providerId?: string
    recipient?: string
    subject?: string
    status?: string
  }
}

export type SandboxSnapshot = {
  drivers: SandboxDriverInfo[]
  traces: SandboxTrace[]
  stats: {
    traces: number
    events: number
    sends: number
    webhooks: number
    hooks: number
  }
}

export type SendSandboxEmailInput = {
  emailDriver: SandboxDriverId
  fromEmail: string
  fromName?: string
  toEmail: string
  ccEmail?: string
  bccEmail?: string
  replyToEmail?: string
  inReplyToMessageId?: string
  subject: string
  text?: string
  html?: string
  sendAt?: string
  templateId?: string
  templateData?: Record<string, unknown>
  track?: { opens?: boolean; clicks?: boolean }
  unsubscribe?: { global?: boolean; listId?: string; groupId?: string }
  idempotencyKey?: string
  tenantId?: string
  /** outlook-app: Microsoft tenant of the sending mailbox, passed as sender auth. */
  outlookTenant?: string
  tags?: string[]
  metadata?: Record<string, string>
  headers?: Record<string, string>
  provider?: Record<string, unknown>
}

export type DomainActionInput =
  | { action: "list"; emailDriver: string; options?: ListDomainsOptions }
  | {
      action: "create" | "ensure"
      emailDriver: string
      input: CreateDomainInput
    }
  | {
      action: "get" | "verify" | "delete"
      emailDriver: string
      identifier: DomainIdentifier
    }
  | {
      action: "update"
      emailDriver: string
      identifier: DomainIdentifier
      patch: UpdateDomainInput
    }

export type MailboxActionInput =
  | { action: "list"; emailDriver: string; options?: ListMailboxesOptions }
  | { action: "connect"; emailDriver: string; input?: ConnectMailboxInput }
  | { action: "create"; emailDriver: string; input: CreateMailboxInput }
  | { action: "get" | "delete"; emailDriver: string; idOrEmail: string }

export type ProviderFetchInput = {
  emailDriver: string
  path: string
  init?: ProviderFetchInit
}

export type AttachmentInput = {
  emailDriver: string
  attachment: {
    filename: string
    content?: string
    url?: string
    contentType?: string
    size?: number
    contentId?: string
    isInline?: boolean
  }
}

export type RememberedDomain = {
  emailDriver: string
  domain: Domain
}

export type RememberedMailbox = {
  emailDriver: string
  mailbox: Mailbox
  auth?: unknown
}

export type WebhookSetupTarget =
  | { scope: "account" }
  | { scope: "mailbox"; mailboxEmail: string }
  | { scope: "domain"; domain: string }

export type WebhookActionInput =
  | { action: "list"; emailDriver: string }
  | {
      action: "setup"
      emailDriver: string
      target: WebhookSetupTarget
      url?: string
      events?: WebhookEventSelection
    }
  | { action: "refresh" | "delete"; emailDriver: string; rowId: string }
  | { action: "renew-expiring"; emailDriver: string }

export type SandboxWebhookView = {
  rowId: string
  emailDriver: string
  scope: WebhookScope
  providerId?: string
  url: string
  events?: string[]
  status: string
  expiresAt?: string
  renewAfter?: string
  mailboxEmail?: string
  domain?: string
  createdAt?: string
  updatedAt?: string
}

export type SyncTarget =
  | { scope: "account" }
  | { scope: "mailbox"; mailboxEmail: string }
  | { scope: "domain"; domain: string }

export type SyncActionInput = {
  emailDriver: string
  target: SyncTarget
  since: string
  until?: string
  context?: unknown
}

export type SandboxSyncResult = {
  ok: boolean
  scope: WebhookScope
  dispatched: number
  syncedFrom?: string
  /** Set when the run failed partway; resume with `since: lastEventTimestamp`. */
  lastEventTimestamp?: string
  error?: string
}
