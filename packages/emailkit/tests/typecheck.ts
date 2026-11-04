import {
  EmailKit,
  MAILGUN_CAPABILITIES,
  OutlookDriver,
  isGmailAuth,
  isOutlookAuth,
  type EmailDriver,
  type GmailMailboxAuth,
  type OutlookMailboxAuth,
} from "../src";
import type { WebhookLifecycleTarget, WebhookScope } from "../src";

declare const plainDriver: EmailDriver<{}, "plain-driver">;
declare const fetchDriver: EmailDriver<{ providerFetch: true }, "fetch-driver"
>;
declare const senderDriver: EmailDriver<{ senderAuth: true; senderMailbox: true }, "sender-driver"
>;
declare const accountWebhookSetupDriver: EmailDriver<{ webhooks: { account: { setup: true } } }, "account-setup-driver"
>;

const providerFetchClient = EmailKit({
  emailDrivers: [fetchDriver, plainDriver],
  resolveEmailDriver: () => "fetch-driver",
});
providerFetchClient.providerFetch("/v1/messages", {
  emailDriver: "fetch-driver",
});
providerFetchClient.providerFetch("/v1/messages", {
  // @ts-expect-error providerFetch cannot select drivers that do not declare providerFetch.
  emailDriver: "plain-driver",
});

const plainClient = EmailKit({ emailDrivers: [plainDriver] });
// @ts-expect-error providerFetch is absent without a supporting driver.
plainClient.providerFetch("/v1/messages");

EmailKit({
  emailDrivers: [senderDriver, plainDriver],
  resolveEmailDriver: () => ({
    emailDriver: "sender-driver" as const,
    auth: { accessToken: "access_123" },
    mailbox: { id: "mbx_123", email: "sender@example.com" },
  }),
});

EmailKit({
  emailDrivers: [senderDriver, plainDriver],
  // @ts-expect-error resolver sender.auth requires the selected driver to declare senderAuth.
  resolveEmailDriver: () => ({
    emailDriver: "plain-driver" as const,
    auth: { accessToken: "access_123" },
  }),
});

const senderClient = EmailKit({
  emailDrivers: [senderDriver, plainDriver],
  resolveEmailDriver: () => "sender-driver",
});
senderClient.sendEmail({
  from: { email: "sender@example.com" },
  to: { email: "recipient@example.com" },
  subject: "With sender auth",
  sender: {
    emailDriver: "sender-driver" as const,
    auth: { accessToken: "access_123" },
    mailbox: { id: "mbx_123", email: "sender@example.com" },
  },
});
senderClient.sendEmail({
  from: { email: "sender@example.com" },
  to: { email: "recipient@example.com" },
  subject: "Unsupported sender auth",
  sender: {
    emailDriver: "plain-driver" as const,
    // @ts-expect-error sender.auth requires the selected driver to declare senderAuth.
    auth: { accessToken: "access_123" },
  },
});

const accountWebhookSetupClient = EmailKit({
  emailDrivers: [accountWebhookSetupDriver],
});
accountWebhookSetupClient.webhooks.setup({
  url: "https://example.com/webhook",
  events: "all",
});
// @ts-expect-error refresh is absent unless a driver declares account webhook refresh support.
accountWebhookSetupClient.webhooks.refresh({ id: "webhook_123" });

declare const mailboxSyncDriver: EmailDriver<{ sync: { mailbox: true } }, "mailbox-sync-driver"
>;
declare const accountSyncDriver: EmailDriver<{ sync: { account: true } }, "account-sync-driver"
>;

const mailboxSyncClient = EmailKit({
  emailDrivers: [mailboxSyncDriver, plainDriver],
  resolveEmailDriver: () => "mailbox-sync-driver",
});
void mailboxSyncClient.mailboxes.sync({
  email: "support@example.com",
  since: new Date(),
  emailDriver: "mailbox-sync-driver",
});
void mailboxSyncClient.mailboxes.sync({
  email: "support@example.com",
  since: new Date(),
  // @ts-expect-error mailbox sync cannot select drivers that do not declare it.
  emailDriver: "plain-driver",
});
// @ts-expect-error domain sync is absent for mailbox-only sync drivers.
void mailboxSyncClient.domains.sync({
  domain: "example.com",
  since: new Date(),
});
// @ts-expect-error account sync is absent for mailbox-only sync drivers.
void mailboxSyncClient.sync({ since: new Date() });

const accountSyncClient = EmailKit({ emailDrivers: [accountSyncDriver] });
void accountSyncClient.sync({ since: new Date() });

// @ts-expect-error mailbox sync is absent without a supporting driver.
void plainClient.mailboxes.sync({
  email: "support@example.com",
  since: new Date(),
});

const webhookScope: WebhookScope = "account";
const webhookTarget: WebhookLifecycleTarget = { mailboxId: "mbx_123" };
// @ts-expect-error address is not a public webhook scope.
const addressScope: WebhookScope = "address";
// @ts-expect-error address is not a public webhook lifecycle target.
const addressTarget: WebhookLifecycleTarget = { address: "sender@example.com" };

declare const persistedAuth: unknown;
if (isGmailAuth(persistedAuth)) {
  const gmailAuth: GmailMailboxAuth = persistedAuth;
  const accessToken: string = persistedAuth.accessToken;
  void gmailAuth;
  void accessToken;
}
if (isOutlookAuth(persistedAuth)) {
  const outlookAuth: OutlookMailboxAuth = persistedAuth;
  const accessToken: string = persistedAuth.accessToken;
  void outlookAuth;
  void accessToken;
}

declare const mailgunSendDriver: EmailDriver<typeof MAILGUN_CAPABILITIES, "mailgun"
>;

const outlookDraftDriver = OutlookDriver({
  clientId: "client_123",
  clientSecret: "secret_123",
  sendEmailMode: "draft",
});
const outlookDefaultDriver = OutlookDriver({
  clientId: "client_123",
  clientSecret: "secret_123",
});
const outlookSendMailDriver = OutlookDriver({
  clientId: "client_123",
  clientSecret: "secret_123",
  sendEmailMode: "sendMail",
});

// Draft send mode advertises nativeReplyThreading: reply.messageId and
// reply.isReply compile.
const draftCapability: true =
  outlookDraftDriver.capabilities.nativeReplyThreading;
void draftCapability;
void outlookDraftDriver.sendEmail({
  from: { email: "support@example.com" },
  to: { email: "recipient@example.com" },
  subject: "Re: Hello",
  text: "Reply",
  reply: { messageId: "<previous@example.com>", isReply: true },
});
void outlookDraftDriver.sendEmail({
  from: { email: "support@example.com" },
  to: { email: "recipient@example.com" },
  subject: "Re: Hello",
  text: "Reply",
  reply: {
    messageId: "<previous@example.com>",
    // @ts-expect-error reply.references requires the replyHeaders capability.
    references: ["<previous@example.com>"],
  },
});
void outlookDraftDriver.sendEmail({
  from: { email: "support@example.com" },
  to: { email: "recipient@example.com" },
  subject: "Re: Hello",
  text: "Reply",
  reply: {
    // @ts-expect-error reply.threadId requires the replyThreadId capability.
    threadId: "conversation_123",
  },
});

// The default (sendMail) configuration does not advertise
// nativeReplyThreading, so reply.messageId is a compile error.
// @ts-expect-error nativeReplyThreading requires sendEmailMode: "draft".
void outlookDefaultDriver.capabilities.nativeReplyThreading;
void outlookDefaultDriver.sendEmail({
  from: { email: "support@example.com" },
  to: { email: "recipient@example.com" },
  subject: "Re: Hello",
  text: "Reply",
  reply: {
    addresses: [{ email: "support@example.com" }],
    // @ts-expect-error reply.messageId requires sendEmailMode: "draft".
    messageId: "<previous@example.com>",
  },
});
void outlookSendMailDriver.sendEmail({
  from: { email: "support@example.com" },
  to: { email: "recipient@example.com" },
  subject: "Re: Hello",
  text: "Reply",
  reply: {
    addresses: [{ email: "support@example.com" }],
    // @ts-expect-error reply.messageId requires sendEmailMode: "draft".
    messageId: "<previous@example.com>",
  },
});

// Mixed driver tuples keep working with mode-derived Outlook capabilities.
const outlookMixedClient = EmailKit({
  emailDrivers: [outlookDraftDriver, mailgunSendDriver],
  resolveEmailDriver: () => "outlook",
});
void outlookMixedClient;

// Application (Exchange RBAC) mode: app-only token, no mailbox connect, no
// per-send auth, a concrete tenant.
const outlookAppDriver = OutlookDriver({
  id: "outlook-app",
  clientId: "client_123",
  clientSecret: "secret_123",
  tenant: "tenant_123",
  authMode: "application",
  sendEmailMode: "draft",
});
const appThreading: true = outlookAppDriver.capabilities.nativeReplyThreading;
const appConnect: false = outlookAppDriver.capabilities.mailboxConnect;
const appSenderAuth: false = outlookAppDriver.capabilities.senderAuth;
const appDriverId: "outlook-app" = outlookAppDriver.id;
void appThreading;
void appConnect;
void appSenderAuth;
void appDriverId;
// @ts-expect-error application mode requires a tenant.
OutlookDriver({
  clientId: "client_123",
  clientSecret: "secret_123",
  authMode: "application",
});
// @ts-expect-error application mode uses its own token, not webhookAuth.
OutlookDriver({
  clientId: "client_123",
  clientSecret: "secret_123",
  tenant: "tenant_123",
  authMode: "application",
  webhookAuth: { accessToken: "access_123" },
});

const outlookAppClient = EmailKit({ emailDrivers: [outlookAppDriver] });
outlookAppClient.sendEmail({
  from: { email: "support@example.com" },
  to: { email: "recipient@example.com" },
  subject: "App-only send",
  text: "Hi",
  sender: {
    emailDriver: "outlook-app" as const,
    mailbox: { id: "user_guid", email: "support@example.com" },
  },
});
outlookAppClient.sendEmail({
  from: { email: "support@example.com" },
  to: { email: "recipient@example.com" },
  subject: "App-only send",
  text: "Hi",
  sender: {
    emailDriver: "outlook-app" as const,
    // @ts-expect-error application mode does not take per-send auth.
    auth: { accessToken: "access_123" },
  },
});
void outlookAppClient.mailboxes.webhooks.setup({
  emailDriver: "outlook-app",
  mailbox: { id: "user_guid", email: "support@example.com" },
  events: ["inbound"],
});
void outlookAppClient.mailboxes.webhooks.refresh({
  emailDriver: "outlook-app",
  providerId: "sub_123",
  mailbox: { id: "user_guid", email: "support@example.com" },
});
void outlookAppClient.mailboxes.sync({
  emailDriver: "outlook-app",
  mailbox: { id: "user_guid", email: "support@example.com" },
  since: new Date(),
});
// @ts-expect-error application mode cannot connect mailboxes through OAuth.
void outlookAppClient.mailboxes.connect({ email: "support@example.com" });

// replyHeaders drivers keep accepting RFC reply headers unchanged.
void mailgunSendDriver.sendEmail({
  from: { email: "support@example.com" },
  to: { email: "recipient@example.com" },
  subject: "Re: Hello",
  text: "Reply",
  reply: {
    messageId: "<previous@example.com>",
    references: ["<previous@example.com>"],
    isReply: true,
  },
});

void webhookScope;
void webhookTarget;
void addressScope;
void addressTarget;
