/**
 * Outlook/Microsoft Graph email driver.
 *
 * Implements delegated mailbox OAuth and sendMail through Microsoft Graph.
 * Documentation:
 * - https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-auth-code-flow
 * - https://learn.microsoft.com/graph/api/user-sendmail
 */

import { createHash } from "crypto";
import type {
  EmailDriver,
  EmailDriverConfig,
  EmailDriverOperationOptions,
  ProviderFetchInit,
  SendEmailOptions,
  SyncStream,
} from "../driver";
import type {
  Attachment,
  DriverCapabilities,
  EmailAddress,
  EmailMessage,
  InboundEmailEvent,
  Mailbox,
  MailboxConnectionResult,
  MailboxSyncInput,
  MailboxWebhookDeleteInput,
  MailboxWebhookDeleteResult,
  MailboxWebhookRefreshInput,
  MailboxWebhookRefreshResult,
  MailboxWebhookSetupInput,
  MailboxWebhookSetupResult,
  SendEmailResult,
  Webhook,
  WebhookEvent,
  WebhookEventResult,
  WebhookLifecycleDriverEvent,
  WebhookRecommendedAction,
  WebhookRequest,
  WebhookResponse,
} from "../types";
import { EmailKitError } from "../types";
import { bytesToBase64, stringToBase64 } from "../utils/base64";
import {
  type OAuthMailboxAuth,
  addNormalizedMailboxEmail,
  authorizationHeader,
  createOAuthMailboxKit,
  formatScopes,
  isOAuthMailboxAuth,
  mailboxProviderMetadata,
  nonEmptyString,
  normalizeBaseUrl,
  normalizedMailboxEmails,
  oauthErrorMessage,
  readJsonResponse,
  resolveProviderFetchUrl,
  resolvePublicWebhookUrl,
  sanitizeTokenResponse,
  uniqueScopes,
  unsupportedOAuthMailboxSendFields,
} from "../utils/oauth-mailbox";
import {
  parseWebhookBody,
  webhookEvents,
  webhookRenewAfter,
} from "../utils/webhook";
import {
  OAUTH_STATE_VERSION,
  createCodeChallenge,
  createCodeVerifier,
  createStateNonce,
  encodeOAuthState,
  type OAuthStatePayload,
} from "../utils/oauth-state";
import {
  buildReplyContext,
  hasReplyData,
  replyAddressesAsArray,
  resolveMessageReplyContext,
} from "../utils/reply";

/** Settings shared by delegated and application Outlook drivers. */
export interface OutlookDriverBaseConfig<
  TId extends string = "outlook",
  TSendEmailMode extends OutlookSendEmailMode = OutlookSendEmailMode,
> extends EmailDriverConfig {
  /**
   * EmailKit driver id. Override when configuring multiple Outlook drivers.
   */
  id?: TId;
  clientId: string;
  clientSecret: string;
  /**
   * Microsoft Entra tenant segment. Common values: "common", "organizations",
   * "consumers", or a tenant id/domain.
   */
  tenant?: string;
  /**
   * Optional Microsoft identity base URL.
   * Defaults to https://login.microsoftonline.com.
   */
  authBase?: string;
  /**
   * Optional Microsoft Graph base URL.
   * Defaults to https://graph.microsoft.com/v1.0.
   */
  graphBase?: string;
  /**
   * How outbound messages are submitted to Microsoft Graph.
   *
   * - "sendMail" (default): requires Mail.Send and returns Graph's request id
   *   as a send receipt because Graph returns 202 with no message object.
   * - "draft": creates a draft with immutable ids, sends it, and returns the
   *   Graph message id. This requires Mail.ReadWrite in addition to Mail.Send.
   *   Draft mode also enables native reply threading: when `reply.messageId`
   *   is set, the driver replies through Graph `createReply` so Exchange wires
   *   up conversation threading and reply headers itself. Graph cannot set
   *   `In-Reply-To`/`References` headers directly, so `reply.messageId`
   *   throws NOT_SUPPORTED in "sendMail" mode. This is enforced at the type
   *   level: the `nativeReplyThreading` capability (and with it
   *   `reply.messageId`) only exists on drivers constructed with
   *   `sendEmailMode: "draft"`.
   */
  sendEmailMode?: TSendEmailMode;
  /**
   * Automatically renew Microsoft Graph mailbox subscriptions when Graph sends
   * a lifecycle reauthorizationRequired notification. Delegated drivers need
   * webhookAuth or webhookAuthResolver to authenticate the renewal request;
   * application drivers use their own token.
   * When enabled, new subscriptions default lifecycleNotificationUrl to the
   * normal inbound notification URL unless explicitly configured.
   * Defaults to true.
   */
  autoRenewOnLifecycle?: boolean;
  /**
   * Microsoft Graph subscription resource. Defaults to all messages in the
   * connected user's mailbox.
   */
  inboundResource?: string;
  /**
   * Subscription lifetime in minutes. Outlook message subscriptions expire and
   * must be renewed before expiration.
   */
  inboundSubscriptionMinutes?: number;
  /**
   * Expected Microsoft Graph subscription clientState value(s); notifications
   * that do not carry a matching clientState are rejected with 401. When the
   * driver creates subscriptions itself (autoSubscribeInbound, or no
   * webhookAuthResolver) it defaults to a value derived from the client
   * credentials. Required when webhookAuthResolver is set without
   * autoSubscribeInbound, because subscriptions created outside this driver
   * must carry the same value.
   */
  webhookClientState?: string | string[];
}

/**
 * Delegated Outlook driver: each mailbox signs in through OAuth
 * (`mailboxes.connect`) and every operation runs on that user's token
 * against Graph `/me`.
 */
export interface OutlookDriverConfig<
  TId extends string = "outlook",
  TSendEmailMode extends OutlookSendEmailMode = OutlookSendEmailMode,
> extends OutlookDriverBaseConfig<TId, TSendEmailMode> {
  authMode?: "delegated";
  scopes?: string[];
  /**
   * Create a Microsoft Graph inbound-message subscription immediately after a
   * mailbox is connected. The created subscription is returned on mailbox.raw
   * as `inboundSubscription`. Defaults to false.
   */
  autoSubscribeInbound?: boolean;
  /**
   * Mailbox auth used to hydrate Microsoft Graph change notifications.
   * Microsoft Graph webhooks only include resource identifiers, so inbound
   * message normalization requires an access token with Mail.Read scope.
   */
  webhookAuth?: OutlookMailboxAuth;
  /**
   * Resolve mailbox auth for a specific Graph notification. Use this when one
   * webhook endpoint receives subscriptions for multiple mailboxes.
   */
  webhookAuthResolver?: OutlookWebhookAuthResolver;
}

/**
 * App-only Outlook driver for Exchange Online RBAC for Applications
 * (https://learn.microsoft.com/exchange/permissions-exo/application-rbac).
 * The driver gets its own client-credentials token and addresses each mailbox
 * as Graph `users/{id}`, so operations take `mailbox` (Graph user id or UPN in
 * `id`, or `email`) and no per-mailbox auth. Exchange RBAC decides which
 * mailboxes the app can reach; others fail with Graph 403. `inboundResource`
 * is written relative to `me`, which is replaced with the addressed mailbox.
 */
export interface OutlookApplicationDriverConfig<
  TId extends string = "outlook",
  TSendEmailMode extends OutlookSendEmailMode = OutlookSendEmailMode,
> extends OutlookDriverBaseConfig<TId, TSendEmailMode> {
  authMode: "application";
  /** Tenant id or verified domain. "common", "organizations" and "consumers" are rejected. */
  tenant: string;
  scopes?: never;
  autoSubscribeInbound?: never;
  webhookAuth?: never;
  webhookAuthResolver?: never;
}

export type OutlookAuthMode = "delegated" | "application";

export type OutlookMailboxAuth = OAuthMailboxAuth;

export type OutlookSendEmailMode = "sendMail" | "draft";

export interface OutlookSendEmailResult extends SendEmailResult {
  /**
   * Microsoft Graph request/receipt id for the send operation. When
   * `messageIdKind` is "sendReceipt", this is also the required normalized
   * `messageId` fallback because Graph sendMail does not return a message.
   */
  receiptId?: string;
  raw?: {
    skippedHeaders?: string[];
    messageIdKind?: "graphMessageId" | "sendReceipt";
    replySourceLookup?: unknown;
    createDraft?: unknown;
    updateDraft?: unknown;
    sendDraft?: unknown;
    sendMail?: unknown;
  };
}

export interface OutlookInboundSubscription {
  id: string;
  resource?: string;
  changeType?: string;
  notificationUrl?: string;
  lifecycleNotificationUrl?: string;
  expirationDateTime?: string;
  clientState?: string;
  raw?: unknown;
}

export interface OutlookCapabilities extends DriverCapabilities {
  cc: true;
  bcc: true;
  replyTo: true;
  attachments: true;
  customHeaders: true;
  providerFetch: true;
  senderAuth: true;
  senderMailbox: true;
  requiresSecret: true;
  mailboxConnect: true;
  webhooks: {
    mailbox: true;
  };
  sync: {
    mailbox: true;
  };
  publicRoutes: {
    webhook: true;
    lifecycleWebhook: true;
    connectCallback: true;
    connectLanding: true;
  };
}

/**
 * Capabilities of an Outlook driver configured with `sendEmailMode: "draft"`.
 * Only the draft send flow can thread replies natively through Graph
 * `createReply`, so `nativeReplyThreading` exists exclusively here.
 */
export interface OutlookDraftCapabilities extends OutlookCapabilities {
  nativeReplyThreading: true;
}

export const OUTLOOK_CAPABILITIES = {
  cc: true,
  bcc: true,
  replyTo: true,
  attachments: true,
  customHeaders: true,
  providerFetch: true,
  senderAuth: true,
  senderMailbox: true,
  requiresSecret: true,
  mailboxConnect: true,
  webhooks: {
    mailbox: true,
  },
  sync: {
    mailbox: true,
  },
  publicRoutes: {
    webhook: true,
    lifecycleWebhook: true,
    connectCallback: true,
    connectLanding: true,
  },
} as const satisfies OutlookCapabilities;

export const OUTLOOK_DRAFT_CAPABILITIES = {
  ...OUTLOOK_CAPABILITIES,
  nativeReplyThreading: true,
} as const satisfies OutlookDraftCapabilities;

/**
 * Capabilities advertised by an Outlook driver for a given configured
 * `sendEmailMode`. Mirrors the runtime capabilities object the factory builds.
 */
export type OutlookCapabilitiesForSendMode<
  TSendEmailMode extends OutlookSendEmailMode,
> = TSendEmailMode extends "draft"
  ? typeof OUTLOOK_DRAFT_CAPABILITIES
  : typeof OUTLOOK_CAPABILITIES;

/**
 * Capabilities of an application-mode Outlook driver. The driver
 * authenticates as the app, so there is no OAuth mailbox connect, no
 * per-send `sender.auth` and no EmailKit secret for signed OAuth state.
 */
export const OUTLOOK_APPLICATION_CAPABILITIES = {
  ...OUTLOOK_CAPABILITIES,
  senderAuth: false,
  requiresSecret: false,
  mailboxConnect: false,
  publicRoutes: {
    webhook: true,
    lifecycleWebhook: true,
  },
} as const satisfies DriverCapabilities;

export const OUTLOOK_APPLICATION_DRAFT_CAPABILITIES = {
  ...OUTLOOK_APPLICATION_CAPABILITIES,
  nativeReplyThreading: true,
} as const satisfies DriverCapabilities;

/** Application-mode counterpart of {@link OutlookCapabilitiesForSendMode}. */
export type OutlookApplicationCapabilitiesForSendMode<
  TSendEmailMode extends OutlookSendEmailMode,
> = TSendEmailMode extends "draft"
  ? typeof OUTLOOK_APPLICATION_DRAFT_CAPABILITIES
  : typeof OUTLOOK_APPLICATION_CAPABILITIES;

const OUTLOOK_CAPABILITIES_BY_MODE = {
  delegated: {
    sendMail: OUTLOOK_CAPABILITIES,
    draft: OUTLOOK_DRAFT_CAPABILITIES,
  },
  application: {
    sendMail: OUTLOOK_APPLICATION_CAPABILITIES,
    draft: OUTLOOK_APPLICATION_DRAFT_CAPABILITIES,
  },
} as const satisfies Record<
  OutlookAuthMode,
  Record<OutlookSendEmailMode, DriverCapabilities>
>;

const SHARED_TENANT_SEGMENTS = /^(common|organizations|consumers)$/i;

type AnyOutlookDriverConfig =
  | OutlookDriverConfig<string>
  | OutlookApplicationDriverConfig<string>;

const DEFAULT_SCOPES = [
  "offline_access",
  "User.Read",
  "Mail.Send",
  "Mail.Read",
];
const PROVIDER = "outlook";
const DEFAULT_INBOUND_RESOURCE = "me/messages";
const DEFAULT_INBOUND_SUBSCRIPTION_MINUTES = 60 * 24 * 3;
const WEBHOOK_RENEWAL_BUFFER_MS = 60 * 60 * 1000;
const IMMUTABLE_ID_PREFER_HEADER = 'IdType="ImmutableId"';

interface OutlookStatePayload extends OAuthStatePayload {
  provider: typeof PROVIDER;
  webhookUrl?: string;
  lifecycleNotificationUrl?: string;
}

interface MicrosoftUserResponse {
  id?: string;
  displayName?: string;
  mail?: string | null;
  userPrincipalName?: string;
  [key: string]: unknown;
}

interface MicrosoftSubscriptionResponse {
  id?: string;
  resource?: string;
  changeType?: string;
  notificationUrl?: string;
  lifecycleNotificationUrl?: string;
  expirationDateTime?: string;
  clientState?: string;
  [key: string]: unknown;
}

type GraphRecipient = {
  emailAddress: {
    address: string;
    name?: string;
  };
};

interface GraphWebhookPayload {
  value: GraphChangeNotification[];
  [key: string]: unknown;
}

interface GraphChangeNotification {
  subscriptionId?: string;
  changeType?: string;
  resource?: string;
  resourceData?: {
    id?: string;
    "@odata.id"?: string;
    "@odata.type"?: string;
    [key: string]: unknown;
  };
  clientState?: string;
  tenantId?: string;
  subscriptionExpirationDateTime?: string;
  lifecycleEvent?: string;
  [key: string]: unknown;
}

export interface OutlookWebhookAuthResolverContext {
  notification: GraphChangeNotification;
  request: WebhookRequest;
  query?: Record<string, string>;
  subscriptionId?: string;
  clientState?: string;
  mailboxEmail?: string;
  mailboxId?: string;
  messageId?: string;
}

export type OutlookWebhookAuthResolver = (
  context: OutlookWebhookAuthResolverContext,
) =>
  | OutlookMailboxAuth
  | undefined
  | null
  | Promise<OutlookMailboxAuth | undefined | null>;

interface GraphEmailAddress {
  emailAddress?: {
    address?: string;
    name?: string | null;
  };
}

interface GraphMessage {
  id?: string;
  internetMessageId?: string;
  subject?: string | null;
  body?: {
    contentType?: string;
    content?: string | null;
  };
  bodyPreview?: string | null;
  from?: GraphEmailAddress | null;
  toRecipients?: GraphEmailAddress[];
  ccRecipients?: GraphEmailAddress[];
  bccRecipients?: GraphEmailAddress[];
  replyTo?: GraphEmailAddress[];
  receivedDateTime?: string;
  sentDateTime?: string;
  conversationId?: string;
  internetMessageHeaders?: Array<{
    name?: string;
    value?: string;
  }>;
  attachments?: GraphAttachment[];
  [key: string]: unknown;
}

interface GraphAttachment {
  id?: string;
  name?: string;
  contentType?: string | null;
  size?: number;
  isInline?: boolean;
  contentId?: string | null;
  [key: string]: unknown;
}

interface OutlookAttachmentProviderMetadata {
  notification?: GraphChangeNotification;
  query?: Record<string, string>;
  mailboxEmail?: string;
  mailboxId?: string;
  messageId?: string;
}

const resolveWebhookSetupUrl = (input: { url?: string }): string => {
  const url = nonEmptyString(input.url);
  if (!url) {
    throw new EmailKitError(
      "Outlook mailbox webhook setup requires url from EmailKit public routes",
      PROVIDER,
      "MISSING_WEBHOOK_URL",
    );
  }
  return url;
};

const resolveLifecycleNotificationUrl = (
  notificationUrl: string,
  provider?: Record<string, unknown>,
  options?: EmailDriverOperationOptions,
  autoRenewOnLifecycle = true,
): string | undefined => {
  const publicRoutes = options?.publicRoutes;
  return (
    nonEmptyString(provider?.lifecycleNotificationUrl) ||
    nonEmptyString(publicRoutes?.webhook?.lifecycleNotificationUrl) ||
    (autoRenewOnLifecycle ? notificationUrl : undefined)
  );
};

const scopesForMailboxConnect = (
  scopes: string[],
  config: AnyOutlookDriverConfig,
): string[] => {
  const requiredScopes = [...scopes];
  if (config.autoSubscribeInbound) requiredScopes.push("Mail.Read");
  if (config.sendEmailMode === "draft") {
    requiredScopes.push("Mail.ReadWrite");
  }
  return uniqueScopes(requiredScopes);
};

const resolveSendEmailMode = (
  value: unknown,
  source: string,
): OutlookSendEmailMode | undefined => {
  if (value === undefined) return undefined;
  if (value === "sendMail" || value === "draft") return value;
  throw new EmailKitError(
    `Outlook ${source} must be "sendMail" or "draft"`,
    PROVIDER,
    "INVALID_PROVIDER_OPTION",
  );
};

const defaultWebhookClientState = (
  config: AnyOutlookDriverConfig,
): string =>
  `emailkit:${createHash("sha256")
    .update(
      `${config.clientId}:${config.clientSecret}${
        config.authMode === "application" ? `:${config.tenant}` : ""
      }`,
    )
    .digest("hex")
    .slice(0, 32)}`;

const webhookClientStates = (config: AnyOutlookDriverConfig): string[] => {
  if (Array.isArray(config.webhookClientState))
    return config.webhookClientState;
  if (config.webhookClientState) return [config.webhookClientState];
  return config.webhookAuthResolver && !config.autoSubscribeInbound
    ? []
    : [defaultWebhookClientState(config)];
};

const subscriptionClientState = (config: AnyOutlookDriverConfig): string =>
  webhookClientStates(config)[0] || defaultWebhookClientState(config);

const autoRenewOnLifecycleEnabled = (
  config: AnyOutlookDriverConfig,
): boolean => config.autoRenewOnLifecycle !== false;

const addressToRecipient = (address: EmailAddress): GraphRecipient => ({
  emailAddress: {
    address: address.email,
    ...(address.name ? { name: address.name } : {}),
  },
});

const addressesToRecipients = (
  addresses: EmailAddress | EmailAddress[] | undefined,
): GraphRecipient[] | undefined => {
  if (!addresses) return undefined;
  const list = Array.isArray(addresses) ? addresses : [addresses];
  return list.map(addressToRecipient);
};

const graphAddressToEmailAddress = (
  recipient: GraphEmailAddress | null | undefined,
): EmailAddress | undefined => {
  const address = recipient?.emailAddress?.address?.trim();
  if (!address) return undefined;
  const name = recipient?.emailAddress?.name?.trim();
  return {
    email: address,
    ...(name ? { name } : {}),
  };
};

const graphRecipientsToEmailAddresses = (
  recipients: GraphEmailAddress[] | undefined,
): EmailAddress[] => {
  if (!recipients) return [];
  return recipients
    .map(graphAddressToEmailAddress)
    .filter((address): address is EmailAddress => Boolean(address));
};

const attachmentToGraph = (attachment: Attachment): Record<string, unknown> => {
  if (attachment.content === undefined) {
    throw new EmailKitError(
      `Outlook attachment ${attachment.filename} must include content; URL-only attachments are not supported by sendMail`,
      PROVIDER,
      "INVALID_ATTACHMENT",
    );
  }

  return {
    "@odata.type": "#microsoft.graph.fileAttachment",
    name: attachment.filename,
    contentType: attachment.contentType || "application/octet-stream",
    contentBytes:
      typeof attachment.content === "string"
        ? stringToBase64(attachment.content)
        : bytesToBase64(attachment.content),
    ...(attachment.isInline ? { isInline: true } : {}),
    ...(attachment.contentId ? { contentId: attachment.contentId } : {}),
  };
};

const angleBracketedMessageId = (messageId: string): string =>
  `<${messageId.trim().replace(/^</, "").replace(/>$/, "")}>`;

const escapeODataStringLiteral = (value: string): string =>
  value.replace(/'/g, "''");

export const isOutlookAuth = (auth: unknown): auth is OutlookMailboxAuth =>
  isOAuthMailboxAuth(auth);

const configuredWebhookAuth = async (
  config: AnyOutlookDriverConfig,
  request: WebhookRequest,
  notification: GraphChangeNotification,
): Promise<OutlookMailboxAuth | undefined> => {
  if (config.webhookAuthResolver) {
    const metadata = notification as GraphChangeNotification & {
      mailboxEmail?: string;
      mailboxId?: string;
      messageId?: string;
    };
    const resolved = await config.webhookAuthResolver({
      notification,
      request,
      query: request.query,
      subscriptionId: notification.subscriptionId,
      clientState: notification.clientState,
      mailboxEmail: metadata.mailboxEmail,
      mailboxId: metadata.mailboxId,
      messageId: metadata.messageId,
    });
    if (isOutlookAuth(resolved)) return resolved;
  }
  if (isOutlookAuth(config.webhookAuth)) return config.webhookAuth;
  return undefined;
};

const staticWebhookAuthFromConfig = (
  config: AnyOutlookDriverConfig,
): OutlookMailboxAuth | undefined => {
  if (isOutlookAuth(config.webhookAuth)) return config.webhookAuth;
  return undefined;
};

const normalizeInboundSubscription = (
  subscription: OutlookInboundSubscription,
  input: {
    driverId: string;
    url: string;
    events?: Webhook["events"];
  },
): Webhook => {
  const expiresAt = subscription.expirationDateTime
    ? new Date(subscription.expirationDateTime)
    : undefined;
  const validExpiresAt =
    expiresAt && !Number.isNaN(expiresAt.getTime()) ? expiresAt : undefined;

  return {
    id: subscription.id,
    emailDriver: input.driverId,
    scope: "mailbox",
    url: subscription.notificationUrl || input.url,
    events: input.events,
    status: "active",
    providerId: subscription.id,
    expiresAt: validExpiresAt,
    renewAfter: webhookRenewAfter(validExpiresAt, WEBHOOK_RENEWAL_BUFFER_MS),
    raw: subscription.raw || subscription,
  };
};

const webhookSubscriptionId = (
  input: MailboxWebhookRefreshInput | MailboxWebhookDeleteInput,
): string => {
  const id =
    ("providerId" in input ? input.providerId : undefined) ||
    input.webhook?.providerId ||
    input.webhookId ||
    input.webhook?.id;
  if (!id) {
    throw new EmailKitError(
      "Outlook mailbox webhook operation requires a Graph subscription id",
      PROVIDER,
      "MISSING_WEBHOOK_ID",
    );
  }
  return id;
};

const isGraphWebhookPayload = (
  payload: unknown,
): payload is GraphWebhookPayload =>
  typeof payload === "object" &&
  payload !== null &&
  Array.isArray((payload as GraphWebhookPayload).value);

const isMessageNotification = (
  notification: GraphChangeNotification,
): boolean => {
  const resource =
    notification.resource || notification.resourceData?.["@odata.id"];
  const odataType = notification.resourceData?.["@odata.type"];
  return Boolean(
    (resource && /(^|\/)messages(\/|$|\()/i.test(resource)) ||
      odataType === "#Microsoft.Graph.Message" ||
      odataType === "#microsoft.graph.message",
  );
};

const isCreatedMessageNotification = (
  notification: GraphChangeNotification,
): boolean => {
  if (isLifecycleNotification(notification)) return false;
  const changeType = notification.changeType?.toLowerCase();
  return (
    isMessageNotification(notification) &&
    (!changeType || changeType === "created")
  );
};

const graphLifecycleEventName = (
  notification: GraphChangeNotification,
): string | undefined => notification.lifecycleEvent?.toLowerCase();

const isLifecycleNotification = (
  notification: GraphChangeNotification,
): boolean => {
  const lifecycleEvent = graphLifecycleEventName(notification);
  return (
    lifecycleEvent === "reauthorizationrequired" ||
    lifecycleEvent === "subscriptionremoved" ||
    lifecycleEvent === "missed"
  );
};

const lifecycleTargetFromNotification = (
  notification: GraphChangeNotification,
  request: WebhookRequest,
) => {
  const mailboxEmail = request.query?.mailboxEmail || request.query?.email;
  const explicitMailboxId = request.query?.mailboxId;
  const resource =
    notification.resource || notification.resourceData?.["@odata.id"];
  const resourceMailboxId = resource?.match(/(?:^|\/)users\/([^/]+)/i)?.[1];
  const mailboxId = explicitMailboxId || resourceMailboxId;
  if (!mailboxEmail && !mailboxId) return undefined;

  return {
    ...(mailboxEmail ? { mailboxEmail } : {}),
    ...(mailboxId ? { mailboxId: decodeURIComponent(mailboxId) } : {}),
  };
};

const lifecycleRecommendedActions = (
  notification: GraphChangeNotification,
): WebhookRecommendedAction[] => {
  switch (graphLifecycleEventName(notification)) {
    case "reauthorizationrequired":
      return notification.subscriptionId ? ["renew"] : ["reauthorize"];
    case "subscriptionremoved":
      return ["delete_local", "recreate"];
    case "missed":
      return ["sync"];
    default:
      return ["inspect"];
  }
};

const transformGraphLifecycleNotification = (
  notification: GraphChangeNotification,
  request: WebhookRequest,
  driverId: string,
): WebhookLifecycleDriverEvent | undefined => {
  const lifecycleEvent = graphLifecycleEventName(notification);
  if (!lifecycleEvent) return undefined;

  const subscriptionId = notification.subscriptionId;
  const target = lifecycleTargetFromNotification(notification, request);
  const base = {
    emailDriver: driverId,
    source: "provider" as const,
    scope: "mailbox" as const,
    ...(subscriptionId
      ? {
          providerId: subscriptionId,
          subscriptionId,
          webhookId: subscriptionId,
        }
      : {}),
    ...(target ? { target } : {}),
    recommendedActions: lifecycleRecommendedActions(notification),
    receivedAt: new Date(),
    raw: notification,
  };

  switch (lifecycleEvent) {
    case "reauthorizationrequired":
      return {
        type: "webhook.lifecycle",
        data: {
          ...base,
          action: "action_required",
          reason: "reauthorization_required",
        },
      };
    case "subscriptionremoved":
      return {
        type: "webhook.lifecycle",
        data: {
          ...base,
          action: "deleted",
          reason: "subscription_removed",
          status: "deleted",
        },
      };
    case "missed":
      return {
        type: "webhook.lifecycle",
        data: {
          ...base,
          action: "sync_required",
          reason: "notifications_missed",
        },
      };
  }
};

const verifyWebhookClientState = (
  config: AnyOutlookDriverConfig,
  notification: GraphChangeNotification,
): void => {
  const allowed = webhookClientStates(config);
  if (allowed.length === 0) {
    // Normally unreachable: OutlookDriver() rejects this configuration at
    // construction. Guards dynamic config mutation after construction.
    throw new EmailKitError(
      "Outlook webhook rejected: no clientState configured. Set webhookClientState in the driver config and include the same value in your Graph subscriptions.",
      PROVIDER,
      "MISSING_WEBHOOK_CLIENT_STATE",
      401,
    );
  }

  if (notification.clientState && allowed.includes(notification.clientState)) {
    return;
  }

  throw new EmailKitError(
    "Invalid Outlook webhook clientState",
    PROVIDER,
    "INVALID_WEBHOOK_CLIENT_STATE",
    401,
  );
};

const outlookProviderMetadata = (
  provider: Record<string, unknown> | undefined,
): OutlookAttachmentProviderMetadata | undefined => {
  const outlook = provider?.outlook;
  if (!outlook || typeof outlook !== "object") return undefined;
  return outlook as OutlookAttachmentProviderMetadata;
};

// Graph ids are GUIDs and URL-safe base64. Excluding ".", "%" and "\" keeps URL
// normalization from turning a segment such as ".." into a folder path.
const SINGLE_MESSAGE_RESOURCE =
  /^\/?(?:me|users\/[\w@-]+)\/messages\/[\w=-]+$/i;

const notificationResourcePath = (
  notification: GraphChangeNotification,
): string | undefined => {
  const resource =
    notification.resourceData?.["@odata.id"] || notification.resource;
  if (resource) {
    if (!SINGLE_MESSAGE_RESOURCE.test(resource)) {
      throw new EmailKitError(
        "Invalid Microsoft Graph message resource URL",
        PROVIDER,
        "INVALID_WEBHOOK_RESOURCE",
        400,
      );
    }
    return resource.replace(/^\//, "");
  }
  const messageId = notification.resourceData?.id;
  return messageId ? `me/messages/${encodeURIComponent(messageId)}` : undefined;
};

const applyGraphMessageSelection = (url: URL): void => {
  url.searchParams.set(
    "$select",
    [
      "id",
      "internetMessageId",
      "subject",
      "body",
      "bodyPreview",
      "from",
      "toRecipients",
      "ccRecipients",
      "bccRecipients",
      "replyTo",
      "receivedDateTime",
      "sentDateTime",
      "conversationId",
      "internetMessageHeaders",
      "attachments",
    ].join(","),
  );
  url.searchParams.set(
    "$expand",
    "attachments($select=id,name,contentType,size,isInline)",
  );
};

const graphMessageUrl = (graphBase: string, resourcePath: string): string => {
  const url = new URL(`${graphBase}/${resourcePath}`);
  applyGraphMessageSelection(url);
  return url.toString();
};

const resolveGraphUrl = (
  graphBase: string,
  path: string | URL,
  init?: ProviderFetchInit,
): URL =>
  resolveProviderFetchUrl(graphBase, path, init, {
    provider: PROVIDER,
    message: "Invalid Microsoft Graph providerFetch URL",
  });

const normalizeHeaders = (
  headers: GraphMessage["internetMessageHeaders"],
): Record<string, string> => {
  const normalized: Record<string, string> = {};
  for (const header of headers || []) {
    if (header.name && typeof header.value === "string") {
      normalized[header.name.toLowerCase()] = header.value;
    }
  }
  return normalized;
};

// Extends the shared helper with the Graph /me identities stashed on
// mailbox.raw.user at connect time.
const outlookMailboxEmails = (input: MailboxSyncInput): Set<string> => {
  const emails = normalizedMailboxEmails(input);

  if (
    input.mailbox &&
    typeof input.mailbox === "object" &&
    !Array.isArray(input.mailbox)
  ) {
    const raw = "raw" in input.mailbox ? input.mailbox.raw : undefined;
    if (raw && typeof raw === "object" && !Array.isArray(raw)) {
      const user = (raw as { user?: unknown }).user;
      if (user && typeof user === "object" && !Array.isArray(user)) {
        addNormalizedMailboxEmail(emails, (user as { mail?: unknown }).mail);
        addNormalizedMailboxEmail(
          emails,
          (user as { userPrincipalName?: unknown }).userPrincipalName,
        );
      }
    }
  }

  return emails;
};

const isMessageFromMailbox = (
  message: GraphMessage,
  mailboxEmails: Set<string>,
): boolean => {
  if (mailboxEmails.size === 0) return false;
  const from = graphAddressToEmailAddress(message.from)?.email;
  return typeof from === "string" && mailboxEmails.has(from.toLowerCase());
};

const transformGraphMessage = (
  message: GraphMessage,
  notification: GraphChangeNotification | undefined,
  messageUrl: string,
  request?: WebhookRequest,
  attachmentProvider?: OutlookAttachmentProviderMetadata,
): InboundEmailEvent => {
  const headers = normalizeHeaders(message.internetMessageHeaders);
  const replyTo = graphRecipientsToEmailAddresses(message.replyTo);
  const references = headers.references
    ? headers.references.split(/\s+/).filter(Boolean)
    : undefined;
  const bodyContent = message.body?.content || undefined;
  const bodyType = message.body?.contentType?.toLowerCase();
  const timestamp =
    message.receivedDateTime ||
    message.sentDateTime ||
    new Date().toISOString();
  const attachmentBaseUrl = messageUrl.replace(/[?#].*$/, "");
  const attachments = (message.attachments || [])
    .filter((attachment) => attachment.id || attachment.name)
    .map(
      (attachment, index): Attachment => ({
        filename: attachment.name || `attachment-${index + 1}`,
        contentType: attachment.contentType || undefined,
        size: attachment.size,
        contentId: attachment.contentId || undefined,
        isInline: attachment.isInline || undefined,
        ...(attachment.id
          ? {
              url: `${attachmentBaseUrl}/attachments/${encodeURIComponent(
                attachment.id,
              )}/$value`,
            }
          : {}),
        provider: {
          outlook: {
            ...(notification ? { notification } : {}),
            ...(request?.query ? { query: request.query } : {}),
            ...(attachmentProvider
              ? {
                  ...attachmentProvider,
                  ...(message.id ? { messageId: message.id } : {}),
                }
              : {}),
          },
        },
      }),
    );

  return {
    schemaVersion: "1",
    eventId: notification
      ? notification.subscriptionId && message.id
        ? `${notification.subscriptionId}:${message.id}:${
            notification.changeType || "updated"
          }`
        : undefined
      : message.id
        ? `sync:${message.id}`
        : undefined,
    messageId: message.internetMessageId || message.id || "",
    providerId: message.id,
    from: graphAddressToEmailAddress(message.from) || { email: "" },
    to: graphRecipientsToEmailAddresses(message.toRecipients),
    cc: graphRecipientsToEmailAddresses(message.ccRecipients),
    bcc: graphRecipientsToEmailAddresses(message.bccRecipients),
    reply: buildReplyContext({
      addresses: replyTo,
      messageId: headers["in-reply-to"],
      references,
      threadId: message.conversationId,
    }),
    subject: message.subject || "",
    text: bodyType === "text" ? bodyContent : message.bodyPreview || undefined,
    html: bodyType === "html" ? bodyContent : undefined,
    attachments: attachments.length > 0 ? attachments : undefined,
    headers,
    timestamp: new Date(timestamp),
    raw: {
      ...(notification ? { notification } : {}),
      message,
    },
  };
};

const requestIdFrom = (response: Response): string | undefined =>
  response.headers.get("request-id") ||
  response.headers.get("x-ms-request-id") ||
  response.headers.get("client-request-id") ||
  undefined;

export function OutlookDriver<
  const TId extends string = "outlook",
  const TSendEmailMode extends OutlookSendEmailMode = "sendMail",
>(
  config: OutlookApplicationDriverConfig<TId, TSendEmailMode>,
): EmailDriver<OutlookApplicationCapabilitiesForSendMode<TSendEmailMode>, TId>;
export function OutlookDriver<
  const TId extends string = "outlook",
  const TSendEmailMode extends OutlookSendEmailMode = "sendMail",
>(
  config: OutlookDriverConfig<TId, TSendEmailMode>,
): EmailDriver<OutlookCapabilitiesForSendMode<TSendEmailMode>, TId>;
export function OutlookDriver(
  config: AnyOutlookDriverConfig,
  // The overloads above carry the per-mode capability types.
): EmailDriver<any, string> {
  const driverId = config.id || "outlook";
  const authMode: OutlookAuthMode = config.authMode || "delegated";
  const applicationAuthMode = authMode === "application";
  if (
    applicationAuthMode &&
    (!config.tenant || SHARED_TENANT_SEGMENTS.test(config.tenant))
  ) {
    throw new EmailKitError(
      'Outlook authMode "application" requires a concrete tenant id or domain, not "common", "organizations" or "consumers"',
      PROVIDER,
      "INVALID_CONFIG",
    );
  }
  if (webhookClientStates(config).length === 0) {
    // Only reachable with webhookAuthResolver and no autoSubscribeInbound
    // (or an explicitly empty webhookClientState array): subscriptions are
    // created outside this driver, so the derived default clientState cannot
    // be assumed and verification needs an explicit value.
    throw new EmailKitError(
      "Outlook webhook verification requires a clientState. Set webhookClientState in the driver config and include the same value in the Graph subscriptions you create.",
      PROVIDER,
      "MISSING_WEBHOOK_CLIENT_STATE",
    );
  }
  const tenant = encodeURIComponent(config.tenant || "common");
  const authBase = normalizeBaseUrl(
    config.authBase || "https://login.microsoftonline.com",
  );
  const graphBase = normalizeBaseUrl(
    config.graphBase || "https://graph.microsoft.com/v1.0",
  );
  const configuredSendEmailMode =
    resolveSendEmailMode(config.sendEmailMode, "sendEmailMode") || "sendMail";
  const capabilities =
    OUTLOOK_CAPABILITIES_BY_MODE[authMode][configuredSendEmailMode];
  const scopes =
    config.scopes && config.scopes.length > 0 ? config.scopes : DEFAULT_SCOPES;
  const mailboxScopes = scopesForMailboxConnect(scopes, config);

  const {
    requireSecret,
    resolveConnectCallbackUrl: resolveMailboxConnectCallbackUrl,
    parseCallbackRequest,
    decodeState,
    fetchToken,
    toAuth,
    resolveMailboxOperationAuth: resolveDelegatedMailboxAuth,
    clientCredentialsAuth,
  } = createOAuthMailboxKit<OutlookMailboxAuth>({
    provider: PROVIDER,
    label: "Outlook",
    issuer: "Microsoft",
    clientId: config.clientId,
    clientSecret: config.clientSecret,
    tokenEndpoint: `${authBase}/${tenant}/oauth2/v2.0/token`,
    // Microsoft refresh grants scope down to the requested scopes.
    refreshForm: (auth) => ({
      scope: formatScopes(
        auth.scopes && auth.scopes.length > 0 ? auth.scopes : mailboxScopes,
      ),
    }),
  });

  // Application mode authenticates as the app: every mailbox and webhook
  // operation uses the cached client-credentials token instead of user auth.
  const getApplicationAuth = clientCredentialsAuth(
    `${new URL(graphBase).origin}/.default`,
  );
  const resolveMailboxOperationAuth: typeof resolveDelegatedMailboxAuth =
    applicationAuthMode
      ? () => getApplicationAuth()
      : resolveDelegatedMailboxAuth;
  const webhookAuthFromConfig: typeof configuredWebhookAuth =
    applicationAuthMode ? () => getApplicationAuth() : configuredWebhookAuth;

  const graphMailboxPath = (mailbox?: {
    id?: string;
    email?: string;
  }): string => {
    if (!applicationAuthMode) return "me";
    const key = nonEmptyString(mailbox?.id) || nonEmptyString(mailbox?.email);
    if (!key) {
      throw new EmailKitError(
        "Outlook application auth requires a mailbox with a Graph user id, UPN or email",
        PROVIDER,
        "MISSING_MAILBOX",
      );
    }
    return `users/${encodeURIComponent(key)}`;
  };

  // A resource that does not start with `me` (e.g. an explicit users/{id}/...)
  // passes through untouched and never requires a mailbox.
  const mailboxResource = (
    resource: string,
    mailbox?: { id?: string; email?: string },
  ): string =>
    resource.replace(/^\/?me(?=\/|$)/i, () => graphMailboxPath(mailbox));

  const lookupNativeReplySource = async (
    authorization: string,
    mailboxPath: string,
    replyMessageId: string,
    signal?: AbortSignal,
  ): Promise<{ sourceId?: string; raw: unknown }> => {
    const internetMessageId = angleBracketedMessageId(replyMessageId);
    const filter = `internetMessageId eq '${escapeODataStringLiteral(
      internetMessageId,
    )}'`;
    const url = `${graphBase}/${mailboxPath}/messages?$filter=${encodeURIComponent(
      filter,
    )}&$select=id,internetMessageId`;

    const response = await fetch(url, {
      method: "GET",
      headers: {
        Authorization: authorization,
        Accept: "application/json",
        Prefer: IMMUTABLE_ID_PREFER_HEADER,
      },
      signal,
    });
    const body = await readJsonResponse(response);
    if (!response.ok) {
      throw new EmailKitError(
        oauthErrorMessage(
          body,
          "Microsoft Graph reply source message lookup failed",
        ),
        PROVIDER,
        undefined,
        response.status,
        undefined,
        body,
      );
    }

    // The internetMessageId filter is known to return stale or unrelated
    // hits, so only trust a result whose internetMessageId matches exactly.
    const value =
      typeof body === "object" && body !== null
        ? (body as { value?: unknown }).value
        : undefined;
    const match = Array.isArray(value)
      ? (value as GraphMessage[]).find(
          (candidate) =>
            typeof candidate?.id === "string" &&
            candidate.internetMessageId === internetMessageId,
        )
      : undefined;
    return { sourceId: match?.id, raw: body };
  };

  const deleteDraftMessage = async (
    authorization: string,
    mailboxPath: string,
    draftId: string,
  ): Promise<void> => {
    try {
      await fetch(`${graphBase}/${mailboxPath}/messages/${encodeURIComponent(draftId)}`, {
        method: "DELETE",
        headers: {
          Authorization: authorization,
          Prefer: IMMUTABLE_ID_PREFER_HEADER,
        },
      });
    } catch {
      // Best-effort cleanup; the caller rethrows the original send error.
    }
  };

  const createInboundSubscription = async (
    auth: OutlookMailboxAuth,
    input?: {
      notificationUrl?: string;
      events?: Webhook["events"];
      provider?: Record<string, unknown>;
      options?: EmailDriverOperationOptions;
      mailbox?: { id?: string; email?: string };
    },
  ): Promise<OutlookInboundSubscription> => {
    const provider = input?.provider || {};
    const minutes =
      typeof provider.subscriptionMinutes === "number" &&
      provider.subscriptionMinutes > 0
        ? provider.subscriptionMinutes
        : typeof config.inboundSubscriptionMinutes === "number" &&
            config.inboundSubscriptionMinutes > 0
          ? config.inboundSubscriptionMinutes
          : DEFAULT_INBOUND_SUBSCRIPTION_MINUTES;
    const resource = mailboxResource(
      typeof provider.resource === "string" && provider.resource
        ? provider.resource
        : config.inboundResource || DEFAULT_INBOUND_RESOURCE,
      input?.mailbox,
    );
    const changeType =
      typeof provider.changeType === "string" && provider.changeType
        ? provider.changeType
        : "created";
    const notificationUrl = input?.notificationUrl;
    if (!notificationUrl) {
      throw new EmailKitError(
        "Outlook inbound subscription creation requires a notificationUrl",
        PROVIDER,
        "MISSING_WEBHOOK_URL",
      );
    }
    const clientState =
      typeof provider.clientState === "string" && provider.clientState
        ? provider.clientState
        : subscriptionClientState(config);
    const lifecycleNotificationUrl = resolveLifecycleNotificationUrl(
      notificationUrl,
      provider,
      input?.options,
      autoRenewOnLifecycleEnabled(config),
    );
    const expirationDateTime =
      typeof provider.expirationDateTime === "string" &&
      provider.expirationDateTime
        ? provider.expirationDateTime
        : provider.expirationDateTime instanceof Date
          ? provider.expirationDateTime.toISOString()
          : new Date(Date.now() + minutes * 60_000).toISOString();

    const requestBody: Record<string, unknown> = {
      changeType,
      notificationUrl,
      resource,
      expirationDateTime,
      clientState,
      latestSupportedTlsVersion: "v1_2",
    };
    if (lifecycleNotificationUrl) {
      requestBody.lifecycleNotificationUrl = lifecycleNotificationUrl;
    }

    const response = await fetch(`${graphBase}/subscriptions`, {
      method: "POST",
      headers: {
        Authorization: authorizationHeader(auth),
        "Content-Type": "application/json",
      },
      body: JSON.stringify(requestBody),
    });
    const body = await readJsonResponse(response);
    if (!response.ok) {
      throw new EmailKitError(
        oauthErrorMessage(
          body,
          "Microsoft Graph inbound subscription creation failed",
        ),
        PROVIDER,
        undefined,
        response.status,
        undefined,
        body,
      );
    }

    const subscription = body as MicrosoftSubscriptionResponse;
    if (!subscription.id) {
      throw new EmailKitError(
        "Microsoft Graph subscription response did not include an id",
        PROVIDER,
        "INVALID_SUBSCRIPTION_RESPONSE",
        undefined,
        undefined,
        body,
      );
    }

    return {
      id: subscription.id,
      resource: subscription.resource,
      changeType: subscription.changeType,
      notificationUrl: subscription.notificationUrl,
      lifecycleNotificationUrl: subscription.lifecycleNotificationUrl,
      expirationDateTime: subscription.expirationDateTime,
      clientState: subscription.clientState,
      raw: subscription,
    };
  };

  const refreshInboundSubscription = async (
    auth: OutlookMailboxAuth,
    id: string,
    provider?: Record<string, unknown>,
  ): Promise<OutlookInboundSubscription> => {
    const minutes =
      typeof provider?.subscriptionMinutes === "number" &&
      provider.subscriptionMinutes > 0
        ? provider.subscriptionMinutes
        : typeof config.inboundSubscriptionMinutes === "number" &&
            config.inboundSubscriptionMinutes > 0
          ? config.inboundSubscriptionMinutes
          : DEFAULT_INBOUND_SUBSCRIPTION_MINUTES;
    const requestBody: Record<string, unknown> = {
      expirationDateTime:
        typeof provider?.expirationDateTime === "string" &&
        provider.expirationDateTime
          ? provider.expirationDateTime
          : provider?.expirationDateTime instanceof Date
            ? provider.expirationDateTime.toISOString()
            : new Date(Date.now() + minutes * 60_000).toISOString(),
    };
    if (
      typeof provider?.notificationUrl === "string" &&
      provider.notificationUrl
    ) {
      requestBody.notificationUrl = provider.notificationUrl;
    }

    const response = await fetch(
      `${graphBase}/subscriptions/${encodeURIComponent(id)}`,
      {
        method: "PATCH",
        headers: {
          Authorization: authorizationHeader(auth),
          "Content-Type": "application/json",
        },
        body: JSON.stringify(requestBody),
      },
    );
    const body = await readJsonResponse(response);
    if (!response.ok) {
      throw new EmailKitError(
        oauthErrorMessage(
          body,
          "Microsoft Graph inbound subscription renewal failed",
        ),
        PROVIDER,
        undefined,
        response.status,
        undefined,
        body,
      );
    }

    const subscription = body as MicrosoftSubscriptionResponse;
    if (!subscription.id) {
      throw new EmailKitError(
        "Microsoft Graph subscription response did not include an id",
        PROVIDER,
        "INVALID_SUBSCRIPTION_RESPONSE",
        undefined,
        undefined,
        body,
      );
    }

    return {
      id: subscription.id,
      resource: subscription.resource,
      changeType: subscription.changeType,
      notificationUrl: subscription.notificationUrl,
      lifecycleNotificationUrl: subscription.lifecycleNotificationUrl,
      expirationDateTime: subscription.expirationDateTime,
      clientState: subscription.clientState,
      raw: subscription,
    };
  };

  const deleteInboundSubscription = async (
    auth: OutlookMailboxAuth,
    id: string,
  ): Promise<{ deleted: boolean; raw: unknown }> => {
    const response = await fetch(
      `${graphBase}/subscriptions/${encodeURIComponent(id)}`,
      {
        method: "DELETE",
        headers: {
          Authorization: authorizationHeader(auth),
        },
      },
    );
    const body = await readJsonResponse(response);
    if (response.status === 404) {
      return { deleted: true, raw: body };
    }
    if (!response.ok) {
      throw new EmailKitError(
        oauthErrorMessage(
          body,
          "Microsoft Graph inbound subscription deletion failed",
        ),
        PROVIDER,
        undefined,
        response.status,
        undefined,
        body,
      );
    }
    return { deleted: true, raw: body };
  };

  const autoRenewLifecycleNotification = async (
    request: WebhookRequest,
    notification: GraphChangeNotification,
  ): Promise<WebhookLifecycleDriverEvent | undefined> => {
    if (!autoRenewOnLifecycleEnabled(config)) return undefined;
    if (graphLifecycleEventName(notification) !== "reauthorizationrequired") {
      return undefined;
    }
    if (!notification.subscriptionId) return undefined;

    const auth = await webhookAuthFromConfig(config, request, notification);
    if (!auth) return undefined;

    const inboundSubscription = await refreshInboundSubscription(
      auth,
      notification.subscriptionId,
    );
    const url = inboundSubscription.notificationUrl;
    if (!url) {
      throw new EmailKitError(
        "Outlook lifecycle renewal requires a notificationUrl from Microsoft Graph",
        PROVIDER,
        "MISSING_WEBHOOK_URL",
      );
    }
    const webhook = normalizeInboundSubscription(inboundSubscription, {
      driverId,
      url,
      events: ["inbound"],
    });

    return {
      type: "webhook.lifecycle",
      data: {
        emailDriver: driverId,
        action: "updated",
        source: "provider",
        reason: "renewed",
        recommendedActions: ["persist"],
        scope: "mailbox",
        webhook,
        webhookId: webhook.id,
        providerId: webhook.providerId,
        subscriptionId: inboundSubscription.id,
        ...(lifecycleTargetFromNotification(notification, request)
          ? { target: lifecycleTargetFromNotification(notification, request) }
          : {}),
        status: webhook.status,
        expiresAt: webhook.expiresAt,
        renewAfter: webhook.renewAfter,
        receivedAt: new Date(),
        raw: {
          notification,
          inboundSubscription,
        },
      },
    };
  };

  const setupMailboxWebhook = async (
    input: MailboxWebhookSetupInput,
    options?: EmailDriverOperationOptions,
  ): Promise<MailboxWebhookSetupResult> => {
    const auth = await resolveMailboxOperationAuth(
      "mailbox webhook operation",
      input,
      options,
    );
    const events = webhookEvents(input.events);
    const url = resolveWebhookSetupUrl(input);
    const { mailboxId, mailboxEmail } = mailboxProviderMetadata(input);
    const inboundSubscription = await createInboundSubscription(auth, {
      notificationUrl: url,
      events,
      provider: input.provider,
      ...(options ? { options } : {}),
      mailbox: { id: mailboxId, email: mailboxEmail },
    });
    const webhook = normalizeInboundSubscription(inboundSubscription, {
      driverId,
      url,
      events,
    });

    return {
      webhook,
      context: input.context ?? options?.context,
      raw: {
        inboundSubscription,
      },
    };
  };

  const resolveProviderFetchAuth = async (
    init?: ProviderFetchInit,
  ): Promise<OutlookMailboxAuth | undefined> => {
    if (applicationAuthMode) return getApplicationAuth();
    const metadata = outlookProviderMetadata(init?.provider);
    if (metadata?.notification) {
      const resolved = await webhookAuthFromConfig(
        config,
        {
          method: "GET",
          headers: {},
          body: null,
          ...(metadata.query ? { query: metadata.query } : {}),
        },
        metadata.notification,
      );
      if (resolved) return resolved;
    }
    if (metadata?.mailboxEmail || metadata?.mailboxId || metadata?.messageId) {
      const resolved = await webhookAuthFromConfig(
        config,
        {
          method: "GET",
          headers: {},
          body: null,
          ...(metadata.query ? { query: metadata.query } : {}),
        },
        {
          resource: metadata.mailboxId
            ? `users/${encodeURIComponent(metadata.mailboxId)}/messages/${
                metadata.messageId
                  ? encodeURIComponent(metadata.messageId)
                  : ""
              }`
            : metadata.messageId
              ? `me/messages/${encodeURIComponent(metadata.messageId)}`
              : undefined,
          resourceData: metadata.messageId ? { id: metadata.messageId } : undefined,
          mailboxEmail: metadata.mailboxEmail,
          mailboxId: metadata.mailboxId,
          messageId: metadata.messageId,
        },
      );
      if (resolved) return resolved;
    }

    return staticWebhookAuthFromConfig(config);
  };

  const oauthConnect = {
    handleCallback: async (
      request: WebhookRequest,
      options,
    ): Promise<MailboxConnectionResult> => {
      const secret = requireSecret(options?.secret, "callback handling");
      const { code, state: stateValue } = parseCallbackRequest(request);
      const state = decodeState<OutlookStatePayload>(stateValue, secret);
      const token = await fetchToken(
        new URLSearchParams({
          client_id: config.clientId,
          client_secret: config.clientSecret,
          grant_type: "authorization_code",
          code,
          redirect_uri: state.callbackUrl,
          scope: formatScopes(state.scopes),
          code_verifier: state.codeVerifier,
        }),
      );
      const auth = toAuth(token);

      const meResponse = await fetch(
        `${graphBase}/me?$select=id,displayName,mail,userPrincipalName`,
        {
          method: "GET",
          headers: {
            Authorization: authorizationHeader(auth),
          },
        },
      );
      const meBody = await readJsonResponse(meResponse);
      if (!meResponse.ok) {
        throw new EmailKitError(
          oauthErrorMessage(meBody, "Microsoft Graph /me request failed"),
          PROVIDER,
          undefined,
          meResponse.status,
          undefined,
          meBody,
        );
      }

      const user = meBody as MicrosoftUserResponse;
      const email = user.mail || user.userPrincipalName || state.email;
      if (!user.id || !email) {
        throw new EmailKitError(
          "Microsoft Graph /me response did not include a usable mailbox id/email",
          PROVIDER,
          "INVALID_ME_RESPONSE",
          undefined,
          undefined,
          meBody,
        );
      }

      const mailbox: Mailbox = {
        id: user.id,
        email,
        displayName: user.displayName,
        status: "connected",
        raw: user,
      };
      const autoSubscribeWebhookUrl =
        state.webhookUrl || resolvePublicWebhookUrl(options);
      const webhookSetup = config.autoSubscribeInbound
        ? await setupMailboxWebhook(
            {
              mailbox,
              auth,
              ...(autoSubscribeWebhookUrl
                ? { url: autoSubscribeWebhookUrl }
                : {}),
              events: ["inbound"],
              ...(state.lifecycleNotificationUrl
                ? {
                    provider: {
                      lifecycleNotificationUrl: state.lifecycleNotificationUrl,
                    },
                  }
                : {}),
              context: state.context,
            },
            { ...options, context: state.context },
          )
        : undefined;
      const inboundSubscription = (
        webhookSetup?.raw as
          | { inboundSubscription?: OutlookInboundSubscription }
          | undefined
      )?.inboundSubscription;
      if (inboundSubscription) {
        mailbox.raw = {
          user,
          inboundSubscription,
          webhook: webhookSetup?.webhook,
        };
      }

      return {
        mailbox,
        auth,
        ...(webhookSetup?.webhook ? { webhooks: [webhookSetup.webhook] } : {}),
        context: state.context,
        raw: {
          token: sanitizeTokenResponse(token),
          user,
          ...(inboundSubscription
            ? { inboundSubscription, webhook: webhookSetup?.webhook }
            : {}),
        },
      };
    },

    mailboxes: {
      connect: async (
        input = {},
        options,
      ): Promise<MailboxConnectionResult> => {
        const secret = requireSecret(options?.secret, "mailbox connect");
        const callbackUrl = resolveMailboxConnectCallbackUrl(input, options);
        const webhookUrl = resolvePublicWebhookUrl(options);
        const lifecycleNotificationUrl = resolveLifecycleNotificationUrl(
          webhookUrl || "",
          input.provider,
          options,
          false,
        );
        const requestedScopes = scopesForMailboxConnect(
          input.scopes && input.scopes.length > 0
            ? input.scopes
            : mailboxScopes,
          config,
        );
        const codeVerifier = createCodeVerifier();
        const codeChallenge = createCodeChallenge(codeVerifier);
        const state = encodeOAuthState(
          {
            v: OAUTH_STATE_VERSION,
            provider: PROVIDER,
            nonce: createStateNonce(),
            issuedAt: Date.now(),
            callbackUrl,
            ...(webhookUrl ? { webhookUrl } : {}),
            ...(lifecycleNotificationUrl ? { lifecycleNotificationUrl } : {}),
            scopes: requestedScopes,
            codeVerifier,
            email: input.email,
            context: input.context,
          },
          secret,
        );

        const authorizationUrl = new URL(
          `${authBase}/${tenant}/oauth2/v2.0/authorize`,
        );
        authorizationUrl.searchParams.set("client_id", config.clientId);
        authorizationUrl.searchParams.set("response_type", "code");
        authorizationUrl.searchParams.set("redirect_uri", callbackUrl);
        authorizationUrl.searchParams.set("response_mode", "query");
        authorizationUrl.searchParams.set(
          "scope",
          formatScopes(requestedScopes),
        );
        authorizationUrl.searchParams.set("state", state);
        authorizationUrl.searchParams.set("code_challenge", codeChallenge);
        authorizationUrl.searchParams.set("code_challenge_method", "S256");
        if (input.email) {
          authorizationUrl.searchParams.set("login_hint", input.email);
        }

        return {
          redirectUrl: authorizationUrl.toString(),
          state,
          context: input.context,
        };
      },
    },
  } satisfies Pick<EmailDriver, "handleCallback" | "mailboxes">;

  return {
    id: driverId,
    name: "outlook",
    capabilities,

    sendEmail: async (
      // The implementation accepts the widest (draft-mode) message shape; the
      // factory return type narrows reply.messageId away for sendMail mode.
      message: EmailMessage<typeof OUTLOOK_DRAFT_CAPABILITIES>,
      options?: SendEmailOptions,
    ): Promise<SendEmailResult> => {
      if (!message.html && !message.text) {
        throw new EmailKitError(
          "Outlook sendEmail requires either html or text content",
          PROVIDER,
          "MISSING_REQUIRED_FIELD",
        );
      }

      const unsupportedSendFields = unsupportedOAuthMailboxSendFields(message);
      if (unsupportedSendFields.length > 0) {
        throw new EmailKitError(
          `Outlook sendEmail does not support these EmailKit send fields: ${unsupportedSendFields.join(
            ", ",
          )}`,
          PROVIDER,
          "NOT_SUPPORTED",
        );
      }

      const senderMailboxEmail = options?.mailbox?.email?.toLowerCase();
      if (
        senderMailboxEmail &&
        message.from.email.toLowerCase() !== senderMailboxEmail
      ) {
        throw new EmailKitError(
          "Outlook sendEmail can only send from the authenticated mailbox. Shared mailbox/send-as is not supported by this driver.",
          PROVIDER,
          "NOT_SUPPORTED",
        );
      }

      const mailboxPath = graphMailboxPath(
        options?.mailbox ?? { email: message.from.email },
      );
      const auth = await resolveMailboxOperationAuth(
        "sendEmail",
        { auth: options?.auth, mailbox: options?.mailbox },
        options,
        options?.signal,
      );

      const reply = resolveMessageReplyContext(message);
      const unsupportedReplyFields = [
        ...(reply.references?.length ? ["reply.references"] : []),
        ...(reply.threadId ? ["reply.threadId"] : []),
      ];
      if (unsupportedReplyFields.length > 0) {
        throw new EmailKitError(
          `Outlook sendEmail does not support reply threading fields: ${unsupportedReplyFields.join(
            ", ",
          )}. Microsoft Graph derives the reference chain itself; use reply.messageId for native reply threading in draft send mode.`,
          PROVIDER,
          "NOT_SUPPORTED",
        );
      }
      const replyTo = hasReplyData(reply)
        ? addressesToRecipients(replyAddressesAsArray(reply))
        : undefined;
      const skippedHeaders = Object.keys(message.headers || {}).filter(
        (name) => !name.toLowerCase().startsWith("x-"),
      );
      if (skippedHeaders.length > 0) {
        throw new EmailKitError(
          `Outlook sendEmail only supports custom headers beginning with X-: ${skippedHeaders.join(
            ", ",
          )}`,
          PROVIDER,
          "NOT_SUPPORTED",
        );
      }
      const internetMessageHeaders = Object.entries(message.headers || {}).map(
        ([name, value]) => ({ name, value: String(value) }),
      );

      const graphMessage: Record<string, unknown> = {
        subject: message.subject,
        from: addressToRecipient(message.from),
        body: {
          contentType: message.html ? "HTML" : "Text",
          content: message.html || message.text || "",
        },
        toRecipients: addressesToRecipients(message.to) || [],
      };
      const ccRecipients = addressesToRecipients(message.cc);
      const bccRecipients = addressesToRecipients(message.bcc);
      if (ccRecipients?.length) graphMessage.ccRecipients = ccRecipients;
      if (bccRecipients?.length) graphMessage.bccRecipients = bccRecipients;
      if (replyTo?.length) graphMessage.replyTo = replyTo;
      if (internetMessageHeaders.length) {
        graphMessage.internetMessageHeaders = internetMessageHeaders;
      }
      const graphAttachments = message.attachments?.length
        ? message.attachments.map(attachmentToGraph)
        : undefined;
      if (graphAttachments) graphMessage.attachments = graphAttachments;

      const sendEmailMode =
        resolveSendEmailMode(
          message.provider?.sendEmailMode,
          "message.provider.sendEmailMode",
        ) || configuredSendEmailMode;

      if (reply.messageId && sendEmailMode !== "draft") {
        throw new EmailKitError(
          'Outlook native reply threading via reply.messageId requires draft send mode. Configure sendEmailMode: "draft" (Mail.ReadWrite) to send threaded replies.',
          PROVIDER,
          "NOT_SUPPORTED",
        );
      }

      // Graph createReply only documents comment/message.body for its JSON
      // message parameter; internetMessageHeaders are only documented as
      // creation-writable on plain POST /messages and are rejected on PATCH.
      // Refuse the combination instead of risking a 400 or a silent drop.
      if (reply.messageId && internetMessageHeaders.length) {
        throw new EmailKitError(
          "Outlook cannot combine custom headers with native reply threading: Microsoft Graph createReply does not accept internetMessageHeaders. Send without reply.messageId or without headers.",
          PROVIDER,
          "NOT_SUPPORTED",
        );
      }

      if (
        sendEmailMode === "draft" &&
        message.provider?.saveToSentItems === false
      ) {
        throw new EmailKitError(
          'Outlook draft send mode always saves to Sent Items. Use sendEmailMode: "sendMail" when provider.saveToSentItems is false.',
          PROVIDER,
          "NOT_SUPPORTED",
        );
      }

      if (sendEmailMode === "draft") {
        const authorization = authorizationHeader(auth);
        const replySource = reply.messageId
          ? await lookupNativeReplySource(
              authorization,
              mailboxPath,
              reply.messageId,
              options?.signal,
            )
          : undefined;
        const replySourceId = replySource?.sourceId;
        const replyThreading: SendEmailResult["replyThreading"] =
          reply.messageId ? (replySourceId ? "applied" : "skipped") : undefined;

        const createDraftResponse = await fetch(
          replySourceId
            ? `${graphBase}/${mailboxPath}/messages/${encodeURIComponent(
                replySourceId,
              )}/createReply`
            : `${graphBase}/${mailboxPath}/messages`,
          {
            method: "POST",
            headers: {
              Authorization: authorization,
              "Content-Type": "application/json",
              Prefer: IMMUTABLE_ID_PREFER_HEADER,
            },
            ...(replySourceId ? {} : { body: JSON.stringify(graphMessage) }),
            signal: options?.signal,
          },
        );
        const createDraftBody = await readJsonResponse(createDraftResponse);
        if (!createDraftResponse.ok) {
          throw new EmailKitError(
            oauthErrorMessage(
              createDraftBody,
              replySourceId
                ? "Microsoft Graph create reply draft failed"
                : "Microsoft Graph create draft failed",
            ),
            PROVIDER,
            undefined,
            createDraftResponse.status,
            undefined,
            createDraftBody,
          );
        }

        const draft =
          typeof createDraftBody === "object" && createDraftBody !== null
            ? (createDraftBody as { id?: unknown; conversationId?: unknown })
            : undefined;
        const draftId = typeof draft?.id === "string" ? draft.id : undefined;
        if (!draftId) {
          throw new EmailKitError(
            "Microsoft Graph create draft response did not include a message id",
            PROVIDER,
            "INVALID_SEND_RESPONSE",
            undefined,
            undefined,
            createDraftBody,
          );
        }
        const threadId =
          typeof draft?.conversationId === "string"
            ? draft.conversationId
            : undefined;

        // Draft cleanup must stay limited to failures that happen strictly
        // before the send attempt: once /send has been issued, an ambiguous
        // failure (timeout, abort, 5xx) may still have delivered the message,
        // and deleting the draft id would remove the Sent Items copy.
        let sendAttempted = false;
        try {
          let updateDraftBody: unknown;
          if (replySourceId) {
            const { attachments: _inlineAttachments, ...replyDraftMessage } =
              graphMessage;
            const updateDraftResponse = await fetch(
              `${graphBase}/${mailboxPath}/messages/${encodeURIComponent(draftId)}`,
              {
                method: "PATCH",
                headers: {
                  Authorization: authorization,
                  "Content-Type": "application/json",
                  Prefer: IMMUTABLE_ID_PREFER_HEADER,
                },
                body: JSON.stringify(replyDraftMessage),
                signal: options?.signal,
              },
            );
            updateDraftBody = await readJsonResponse(updateDraftResponse);
            if (!updateDraftResponse.ok) {
              throw new EmailKitError(
                oauthErrorMessage(
                  updateDraftBody,
                  "Microsoft Graph update reply draft failed",
                ),
                PROVIDER,
                undefined,
                updateDraftResponse.status,
                undefined,
                updateDraftBody,
              );
            }

            // Graph message PATCH does not accept an attachments property, so
            // reply drafts upload each attachment individually.
            for (const attachment of graphAttachments || []) {
              const attachmentResponse = await fetch(
                `${graphBase}/${mailboxPath}/messages/${encodeURIComponent(
                  draftId,
                )}/attachments`,
                {
                  method: "POST",
                  headers: {
                    Authorization: authorization,
                    "Content-Type": "application/json",
                    Prefer: IMMUTABLE_ID_PREFER_HEADER,
                  },
                  body: JSON.stringify(attachment),
                  signal: options?.signal,
                },
              );
              const attachmentBody = await readJsonResponse(attachmentResponse);
              if (!attachmentResponse.ok) {
                throw new EmailKitError(
                  oauthErrorMessage(
                    attachmentBody,
                    "Microsoft Graph reply draft attachment upload failed",
                  ),
                  PROVIDER,
                  undefined,
                  attachmentResponse.status,
                  undefined,
                  attachmentBody,
                );
              }
            }
          }

          sendAttempted = true;
          const sendDraftResponse = await fetch(
            `${graphBase}/${mailboxPath}/messages/${encodeURIComponent(draftId)}/send`,
            {
              method: "POST",
              headers: {
                Authorization: authorization,
                Prefer: IMMUTABLE_ID_PREFER_HEADER,
              },
              signal: options?.signal,
            },
          );
          const sendDraftBody = await readJsonResponse(sendDraftResponse);
          if (!sendDraftResponse.ok) {
            throw new EmailKitError(
              oauthErrorMessage(
                sendDraftBody,
                "Microsoft Graph send draft failed",
              ),
              PROVIDER,
              undefined,
              sendDraftResponse.status,
              undefined,
              sendDraftBody,
            );
          }

          const sendRequestId =
            requestIdFrom(sendDraftResponse) ||
            requestIdFrom(createDraftResponse);
          const result: OutlookSendEmailResult = {
            messageId: draftId,
            provider: driverId,
            requestId: sendRequestId,
            receiptId: sendRequestId,
            providerId: draftId,
            ...(threadId ? { threadId } : {}),
            ...(replyThreading ? { replyThreading } : {}),
            raw: {
              skippedHeaders: skippedHeaders.length
                ? skippedHeaders
                : undefined,
              messageIdKind: "graphMessageId",
              ...(replySource ? { replySourceLookup: replySource.raw } : {}),
              createDraft: createDraftBody,
              ...(updateDraftBody !== undefined
                ? { updateDraft: updateDraftBody }
                : {}),
              sendDraft: sendDraftBody,
            },
          };
          return result;
        } catch (error) {
          if (!sendAttempted) {
            await deleteDraftMessage(authorization, mailboxPath, draftId);
          }
          throw error;
        }
      }

      const requestBody: Record<string, unknown> = {
        message: graphMessage,
      };
      if (message.provider?.saveToSentItems === false) {
        requestBody.saveToSentItems = false;
      }

      const response = await fetch(`${graphBase}/${mailboxPath}/sendMail`, {
        method: "POST",
        headers: {
          Authorization: authorizationHeader(auth),
          "Content-Type": "application/json",
        },
        body: JSON.stringify(requestBody),
        signal: options?.signal,
      });
      const body = await readJsonResponse(response);
      if (!response.ok) {
        throw new EmailKitError(
          oauthErrorMessage(body, "Microsoft Graph sendMail failed"),
          PROVIDER,
          undefined,
          response.status,
          undefined,
          body,
        );
      }

      const requestId = requestIdFrom(response);
      const receiptId = requestId || `${Date.now()}@outlook.graph`;
      const result: OutlookSendEmailResult = {
        messageId: receiptId,
        provider: driverId,
        requestId,
        receiptId,
        raw: {
          skippedHeaders: skippedHeaders.length ? skippedHeaders : undefined,
          messageIdKind: "sendReceipt",
          sendMail: body,
        },
      };
      return result;
    },

    providerFetch: async (path, init) => {
      const auth = await resolveProviderFetchAuth(init);
      if (!auth) {
        throw new EmailKitError(
          "Outlook providerFetch requires webhookAuth or resolvable Outlook attachment metadata",
          PROVIDER,
          "MISSING_AUTH",
        );
      }

      const {
        searchParams: _ignored,
        headers: initHeaders,
        provider: _provider,
        ...restInit
      } = init ?? {};
      const headers = new Headers(initHeaders);
      if (!headers.has("Authorization")) {
        headers.set(
          "Authorization",
          authorizationHeader(auth),
        );
      }

      return fetch(resolveGraphUrl(graphBase, path, init), {
        ...restInit,
        headers,
      });
    },

    handleWebhook: async (
      request: WebhookRequest,
    ): Promise<WebhookEventResult> => {
      const payload = parseWebhookBody(request);
      if (!isGraphWebhookPayload(payload)) {
        return { type: "unknown", data: payload };
      }

      const lifecycleEvents: WebhookLifecycleDriverEvent[] = [];
      for (const notification of payload.value.filter(
        isLifecycleNotification,
      )) {
        verifyWebhookClientState(config, notification);
        lifecycleEvents.push(
          (await autoRenewLifecycleNotification(request, notification)) ||
            transformGraphLifecycleNotification(
              notification,
              request,
              driverId,
            )!,
        );
      }

      const notifications = payload.value.filter(isCreatedMessageNotification);
      if (notifications.length === 0 && lifecycleEvents.length === 0) {
        return { type: "unknown", data: payload };
      }

      const events: Array<WebhookEvent | WebhookLifecycleDriverEvent> = [
        ...lifecycleEvents,
      ];
      for (const notification of notifications) {
        verifyWebhookClientState(config, notification);
        const resourcePath = notificationResourcePath(notification);
        const auth = await webhookAuthFromConfig(config, request, notification);
        if (!resourcePath || !auth) continue;

        const messageUrl = graphMessageUrl(graphBase, resourcePath);
        const response = await fetch(messageUrl, {
          headers: {
            Authorization: authorizationHeader(auth),
            Accept: "application/json",
          },
        });
        const body = await readJsonResponse(response);
        if (!response.ok) {
          throw new EmailKitError(
            oauthErrorMessage(body, "Microsoft Graph message fetch failed"),
            PROVIDER,
            undefined,
            response.status,
            undefined,
            body,
          );
        }

        events.push({
          type: "inbound",
          data: transformGraphMessage(
            body as GraphMessage,
            notification,
            messageUrl,
            request,
          ),
        });
      }

      if (events.length === 0) return { type: "unknown", data: payload };
      return events.length === 1 ? events[0]! : events;
    },

    webhookResponse: async (
      request: WebhookRequest,
      _handled: boolean,
    ): Promise<WebhookResponse> => {
      const validationToken = request.query?.validationToken;
      if (validationToken) {
        return {
          status: 200,
          body: decodeURIComponent(validationToken),
          headers: { "Content-Type": "text/plain" },
        };
      }

      return {
        status: 202,
        body: { success: true },
      };
    },

    // The OAuth connect flow only exists for delegated mailboxes.
    ...(applicationAuthMode ? {} : oauthConnect),

    webhooks: {
      mailbox: {
        setup: setupMailboxWebhook,

        refresh: async (
          input: MailboxWebhookRefreshInput,
          options?: EmailDriverOperationOptions,
        ): Promise<MailboxWebhookRefreshResult> => {
          const auth = await resolveMailboxOperationAuth(
            "mailbox webhook operation",
            input,
            options,
          );
          const id = webhookSubscriptionId(input);
          const inboundSubscription = await refreshInboundSubscription(
            auth,
            id,
            input.provider,
          );
          const url =
            inboundSubscription.notificationUrl ||
            input.webhook?.url ||
            resolvePublicWebhookUrl(options);
          if (!url) {
            throw new EmailKitError(
              "Outlook mailbox webhook refresh requires an existing webhook url or Microsoft Graph notificationUrl",
              PROVIDER,
              "MISSING_WEBHOOK_URL",
            );
          }
          const webhook = normalizeInboundSubscription(inboundSubscription, {
            driverId,
            url,
            events: input.webhook?.events,
          });

          return {
            webhook,
            context: input.context ?? options?.context,
            raw: {
              inboundSubscription,
            },
          };
        },

        delete: async (
          input: MailboxWebhookDeleteInput,
          options?: EmailDriverOperationOptions,
        ): Promise<MailboxWebhookDeleteResult> => {
          const auth = await resolveMailboxOperationAuth(
            "mailbox webhook operation",
            input,
            options,
          );
          const id = webhookSubscriptionId(input);
          const result = await deleteInboundSubscription(auth, id);
          const existingWebhook = input.webhook;
          const url = existingWebhook?.url || resolvePublicWebhookUrl(options);
          if (!url) {
            throw new EmailKitError(
              "Outlook mailbox webhook delete requires an existing webhook url",
              PROVIDER,
              "MISSING_WEBHOOK_URL",
            );
          }

          return {
            deleted: result.deleted,
            webhook: {
              id: existingWebhook?.id || id,
              emailDriver: driverId,
              scope: "mailbox",
              url,
              events: existingWebhook?.events,
              status: "deleted",
              providerId: existingWebhook?.providerId || id,
              raw: existingWebhook?.raw,
            },
            context: input.context ?? options?.context,
            raw: result.raw,
          };
        },
      },
    },

    sync: {
      mailbox: async function* (
        input: MailboxSyncInput,
        options?: EmailDriverOperationOptions,
      ): SyncStream {
        const attachmentProvider = mailboxProviderMetadata(input);
        if (applicationAuthMode && !attachmentProvider.mailboxEmail) {
          throw new EmailKitError(
            "Outlook application sync requires the mailbox email to tell received mail from the mailbox's own sent mail",
            PROVIDER,
            "MISSING_MAILBOX",
          );
        }
        const resource = mailboxResource(
          config.inboundResource || DEFAULT_INBOUND_RESOURCE,
          {
            id: attachmentProvider.mailboxId,
            email: attachmentProvider.mailboxEmail,
          },
        );
        const auth = await resolveMailboxOperationAuth(
          "mailbox sync",
          input,
          options,
          input.signal,
        );
        const until = input.until || new Date();
        const mailboxEmails = outlookMailboxEmails(input);
        const listUrl = new URL(`${graphBase}/${resource}`);
        listUrl.searchParams.set(
          "$filter",
          `receivedDateTime ge ${input.since.toISOString()} and receivedDateTime lt ${until.toISOString()}`,
        );
        listUrl.searchParams.set("$orderby", "receivedDateTime asc");
        listUrl.searchParams.set("$top", "50");
        applyGraphMessageSelection(listUrl);

        let nextUrl: string | undefined = listUrl.toString();
        while (nextUrl) {
          const response = await fetch(nextUrl, {
            headers: {
              Authorization: authorizationHeader(auth),
              Accept: "application/json",
            },
            signal: input.signal,
          });
          const body = await readJsonResponse(response);
          if (!response.ok) {
            throw new EmailKitError(
              oauthErrorMessage(
                body,
                "Microsoft Graph message listing failed",
              ),
              PROVIDER,
              undefined,
              response.status,
              undefined,
              body,
            );
          }

          const page = body as {
            value?: GraphMessage[];
            "@odata.nextLink"?: string;
          };
          for (const message of page.value || []) {
            if (isMessageFromMailbox(message, mailboxEmails)) continue;
            yield {
              type: "inbound",
              data: transformGraphMessage(
                message,
                undefined,
                `${graphBase}/${resource}/${encodeURIComponent(
                  message.id || "",
                )}`,
                undefined,
                attachmentProvider,
              ),
            };
          }

          nextUrl = page["@odata.nextLink"];
        }

        return { syncedFrom: input.since };
      },
    },
  };
}
