import {
  AIINBX_CAPABILITIES,
  AIInbxDriver,
  GMAIL_CAPABILITIES,
  GmailDriver,
  isGmailAuth,
  isOutlookAuth,
  MAILGUN_CAPABILITIES,
  MailgunDriver,
  OUTLOOK_APPLICATION_CAPABILITIES,
  OUTLOOK_CAPABILITIES,
  OutlookDriver,
  RESEND_CAPABILITIES,
  ResendDriver,
  type DriverCapabilities,
  type EmailDriverTuple,
  type Mailbox,
  type OutlookWebhookAuthResolver,
} from "emailkit"

import {
  findPersistedMailbox,
  findPersistedOutlookMailboxForWebhook,
  persistMailbox,
} from "./persistence/mailboxes"
import { findOutlookMailboxIdForSubscription } from "./persistence/webhooks"
import type { SandboxDriverInfo } from "./types"

export type SandboxEmailDriver = EmailDriverTuple[number]

type DriverDefinition = {
  id: string
  label: string
  family: SandboxDriverInfo["family"]
  capabilities: DriverCapabilities
  requiredEnv: string[]
  optionalEnv: string[]
  create: () => SandboxEmailDriver
}

const env = (key: string) => process.env[key] ?? ""

const listEnv = (key: string) =>
  env(key)
    .split(/[,\s]+/)
    .map((value) => value.trim())
    .filter(Boolean)

export const publicBaseUrl = () =>
  (env("PUBLIC_BASE_URL") || env("APP_URL")).replace(/\/+$/, "") ||
  "http://localhost:3210"

const outlookAutoSubscribeInbound = () =>
  env("OUTLOOK_AUTO_SUBSCRIBE_INBOUND") !== "false"

const outlookScopes = () => {
  const configured = listEnv("OUTLOOK_SCOPES")
  return configured.length
    ? configured
    : ["offline_access", "User.Read", "Mail.Send", "Mail.Read"]
}

const outlookProviderMailboxIdFromResource = (resource?: string) => {
  const match = resource?.match(/users\/([^/]+)/i)
  return match?.[1]
}

// Application mode has no connect flow and cannot list mailboxes without a
// directory permission, so the mailboxes granted via Exchange RBAC come from env.
export const configuredMailboxes = (emailDriver: string): Mailbox[] =>
  emailDriver === "outlook-app"
    ? listEnv("OUTLOOK_APP_MAILBOXES").map((email) => ({
        id: email,
        email,
        status: "connected",
      }))
    : []

export const emailkitPublicRoute = (emailDriver: string) =>
  `${publicBaseUrl()}/api/email/${encodeURIComponent(emailDriver)}`

const outlookWebhookAuthResolver: OutlookWebhookAuthResolver = async ({
  mailboxEmail,
  notification,
}) => {
  if (mailboxEmail) {
    const mailbox = await findPersistedMailbox(mailboxEmail, "outlook")
    const auth = mailbox?.auth
    if (isOutlookAuth(auth)) return auth
  }
  if (notification.subscriptionId) {
    const mailbox = await findOutlookMailboxIdForSubscription(
      notification.subscriptionId
    )
    const auth = mailbox?.auth
    if (isOutlookAuth(auth)) return auth
  }
  const persisted = await findPersistedOutlookMailboxForWebhook({
    subscriptionId: notification.subscriptionId,
    providerMailboxId: outlookProviderMailboxIdFromResource(
      notification.resource
    ),
  })
  const auth = persisted?.auth
  return isOutlookAuth(auth) ? auth : undefined
}

const definitions: DriverDefinition[] = [
  {
    id: "resend",
    label: "Resend",
    family: "resend",
    capabilities: RESEND_CAPABILITIES,
    requiredEnv: ["RESEND_API_KEY"],
    optionalEnv: [
      "RESEND_WEBHOOK_SECRET",
      "FROM_EMAIL_ADDRESS",
      "TO_EMAIL_ADDRESS",
    ],
    create: () =>
      ResendDriver({
        id: "resend",
        apiKey: env("RESEND_API_KEY"),
        webhookSecret: env("RESEND_WEBHOOK_SECRET") || undefined,
      }),
  },
  {
    id: "mailgun",
    label: "Mailgun",
    family: "mailgun",
    capabilities: MAILGUN_CAPABILITIES,
    requiredEnv: ["MAILGUN_API_KEY"],
    optionalEnv: [
      "MAILGUN_WEBHOOK_SIGNING_KEY",
      "MAILGUN_REGION",
      "FROM_EMAIL_ADDRESS",
      "TO_EMAIL_ADDRESS",
    ],
    create: () =>
      MailgunDriver({
        id: "mailgun",
        apiKey: env("MAILGUN_API_KEY"),
        region: env("MAILGUN_REGION") === "eu" ? "eu" : "us",
        webhookSigningKey: env("MAILGUN_WEBHOOK_SIGNING_KEY") || undefined,
      }),
  },
  {
    id: "aiinbx",
    label: "AIInbx",
    family: "aiinbx",
    capabilities: AIINBX_CAPABILITIES,
    requiredEnv: ["AI_INBX_API_KEY"],
    optionalEnv: ["AI_INBX_SECRET", "FROM_EMAIL_ADDRESS", "TO_EMAIL_ADDRESS"],
    create: () =>
      AIInbxDriver({
        id: "aiinbx",
        apiKey: env("AI_INBX_API_KEY"),
        webhookSecret: env("AI_INBX_SECRET") || undefined,
      }),
  },
  {
    id: "gmail",
    label: "Gmail",
    family: "gmail",
    capabilities: GMAIL_CAPABILITIES,
    requiredEnv: [
      "GMAIL_CLIENT_ID",
      "GMAIL_CLIENT_SECRET",
      "GMAIL_WEBHOOK_TOKEN",
      "EMAILKIT_SECRET",
    ],
    optionalEnv: [
      "GMAIL_PUBSUB_TOPIC",
      "GMAIL_AUTO_SUBSCRIBE_INBOUND",
      "GMAIL_SCOPES",
      "PUBLIC_BASE_URL",
      "APP_URL",
      "FROM_EMAIL_ADDRESS",
      "TO_EMAIL_ADDRESS",
    ],
    create: () =>
      GmailDriver({
        id: "gmail",
        clientId: env("GMAIL_CLIENT_ID"),
        clientSecret: env("GMAIL_CLIENT_SECRET"),
        pubsubTopic: env("GMAIL_PUBSUB_TOPIC") || undefined,
        scopes: listEnv("GMAIL_SCOPES").length
          ? listEnv("GMAIL_SCOPES")
          : undefined,
        autoSubscribeInbound: env("GMAIL_AUTO_SUBSCRIBE_INBOUND") !== "false",
        verificationToken: env("GMAIL_WEBHOOK_TOKEN"),
        webhookAuthResolver: async ({ mailboxEmail }) => {
          if (!mailboxEmail) return undefined
          const mailbox = await findPersistedMailbox(mailboxEmail, "gmail")
          const auth = mailbox?.auth
          return isGmailAuth(auth) ? auth : undefined
        },
        onAuthUpdated: async ({ auth, mailbox }) => {
          const email = mailbox?.email
          if (!email) return
          const existing = (await findPersistedMailbox(email, "gmail"))
            ?.mailbox
          await persistMailbox(
            "gmail",
            {
              ...existing,
              id: existing?.id ?? mailbox?.id ?? email,
              email,
              status: existing?.status ?? "connected",
            },
            auth
          )
        },
      })
  },
  {
    id: "outlook",
    label: "Outlook",
    family: "outlook",
    capabilities: OUTLOOK_CAPABILITIES,
    requiredEnv: [
      "OUTLOOK_CLIENT_ID",
      "OUTLOOK_CLIENT_SECRET",
      "EMAILKIT_SECRET",
    ],
    optionalEnv: [
      "OUTLOOK_TENANT",
      "OUTLOOK_AUTO_SUBSCRIBE_INBOUND",
      "OUTLOOK_WEBHOOK_CLIENT_STATE",
      "OUTLOOK_SCOPES",
      "PUBLIC_BASE_URL",
      "APP_URL",
      "FROM_EMAIL_ADDRESS",
      "TO_EMAIL_ADDRESS",
    ],
    create: () => {
      const autoSubscribeInbound = outlookAutoSubscribeInbound()
      const webhookClientState = env("OUTLOOK_WEBHOOK_CLIENT_STATE") || undefined
      // Without auto-subscribe the driver requires an explicit clientState to
      // verify notifications; degrade to send-only instead of throwing at
      // module load and taking every sandbox route down.
      const inboundReady = autoSubscribeInbound || Boolean(webhookClientState)
      if (!inboundReady) {
        console.warn(
          "[sandbox] Outlook inbound disabled: set OUTLOOK_WEBHOOK_CLIENT_STATE or enable OUTLOOK_AUTO_SUBSCRIBE_INBOUND"
        )
      }
      return OutlookDriver({
        id: "outlook",
        clientId: env("OUTLOOK_CLIENT_ID"),
        clientSecret: env("OUTLOOK_CLIENT_SECRET"),
        tenant: env("OUTLOOK_TENANT") || "common",
        scopes: outlookScopes(),
        autoSubscribeInbound,
        webhookClientState,
        ...(inboundReady
          ? {
              webhookAuthResolver: outlookWebhookAuthResolver,
            }
          : {}),
      })
    },
  },
  {
    id: "outlook-app",
    label: "Outlook (app)",
    family: "outlook",
    capabilities: OUTLOOK_APPLICATION_CAPABILITIES,
    requiredEnv: [
      "OUTLOOK_CLIENT_ID",
      "OUTLOOK_CLIENT_SECRET",
      "OUTLOOK_APP_TENANT",
      "OUTLOOK_APP_MAILBOXES",
    ],
    optionalEnv: [
      "OUTLOOK_WEBHOOK_CLIENT_STATE",
      "PUBLIC_BASE_URL",
      "APP_URL",
      "TO_EMAIL_ADDRESS",
    ],
    create: () =>
      OutlookDriver({
        id: "outlook-app",
        authMode: "application",
        clientId: env("OUTLOOK_CLIENT_ID"),
        clientSecret: env("OUTLOOK_CLIENT_SECRET"),
        tenant: env("OUTLOOK_APP_TENANT"),
        webhookClientState: env("OUTLOOK_WEBHOOK_CLIENT_STATE") || undefined,
      }),
  },
]

const toInfo = (definition: DriverDefinition): SandboxDriverInfo => {
  const missingRequiredEnv = definition.requiredEnv.filter((key) => !env(key))
  const missingOptionalEnv = definition.optionalEnv.filter((key) => !env(key))

  return {
    id: definition.id,
    label: definition.label,
    family: definition.family,
    publicWebhookUrl: emailkitPublicRoute(definition.id),
    requiredEnv: definition.requiredEnv,
    optionalEnv: definition.optionalEnv,
    missingRequiredEnv,
    missingOptionalEnv,
    ready: missingRequiredEnv.length === 0,
    defaultFromEmail: env("FROM_EMAIL_ADDRESS"),
    defaultToEmail: env("TO_EMAIL_ADDRESS"),
    capabilities: definition.capabilities,
  }
}

export const sandboxDrivers = definitions.map(toInfo)

export const readyEmailDrivers = definitions
  .filter(
    (definition) =>
      sandboxDrivers.find((driver) => driver.id === definition.id)?.ready
  )
  .map((definition): SandboxEmailDriver => definition.create())

export const hasSandboxDriver = (id: unknown): id is string =>
  typeof id === "string" && sandboxDrivers.some((driver) => driver.id === id)

export const getSandboxDriverInfo = (id: unknown) =>
  typeof id === "string"
    ? sandboxDrivers.find((driver) => driver.id === id)
    : undefined

export const defaultEmailDriver = () =>
  readyEmailDrivers[0]?.id ?? sandboxDrivers[0]?.id ?? "resend"
