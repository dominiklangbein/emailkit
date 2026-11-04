<p align="center">
  <img src="https://raw.githubusercontent.com/paukraft/emailkit/main/apps/docs/public/emailkit-logo.png" alt="emailkit" width="96" />
</p>

# emailkit

Unified email SDK with pluggable drivers (Gmail, Mailgun, Resend, AIInbx, Outlook) and optional Next.js helpers.

## Install

```bash
npm i emailkit
```

## Migrating from 3.x to 4.0

Unsubscribes have their own hook:

- Resend and Mailgun unsubscribe events now fire `hooks.email.onUnsubscribed`
  (status `"unsubscribed"`) instead of `onComplained`. An opt-out is not a spam
  report; handle both if you suppress on either.

### AIInbx platform 2.0

The AIInbx driver now targets the rewritten AIInbx platform (API v2). Driver
config is unchanged, but:

- Create a new webhook endpoint (`emailkit.webhooks.setup({ emailDriver: "aiinbx" })`)
  and use its signing secret — v2 signs with a single `AIInbx-Signature` header.
- Replies thread by `reply.threadId` only; `reply.messageId`/`references` are
  gone because AIInbx derives the RFC headers from the thread.
- `track` is gone — tracking is a domain setting (`domains.update`). Custom
  headers, `sendAt`, `idempotencyKey`, and `unsubscribe` are new.
- `sendEmail` returns the RFC Message-ID as `messageId` and the AIInbx email id
  (`eml_…`) as `providerId`; delivery events carry both, so
  `result.messageId === event.messageId` like every other driver.
- Since 4.2.0: `result.messageId` and outbound `event.messageId` are the
  RFC Message-ID again; use `providerId` for the `eml_` id.
- Since 4.3.0 (AIInbx 2.1): `reply.messageId`/`references` and per-send `track`
  are back; `reply.threadId` still replies on the thread. `sync` replays
  delivery, tracking, and mailbox events too, under their webhook `eventId` —
  replayed inbound no longer uses `${emailId}:received`. Verdict fields are
  `string | null` and `verdicts` is always set (null for connected mailboxes).
  `webhooks.setup` creates full-payload endpoints, so inbound needs no second
  API call; recreate older endpoints to get that. Connect `context` travels in
  mailbox metadata (500 characters) and reaches `onDeleted` as well.
- Gmail and Outlook inboxes connect through AIInbx's hosted OAuth:
  `emailkit.mailboxes.connect({ emailDriver: "aiinbx", ... })`; completion
  arrives by webhook on `hooks.mailbox.onConnected`. No `secret` required.
- Category, verdicts, body segments, and prepared attachment Markdown are typed
  under `getAIInbxInbound(event)` / `getAIInbxAttachment(attachment)`.

## Migrating from 2.x to 3.0

Webhook verification is stricter:

- Resend and AIInbx `verifyWebhook` now require `rawBody` (the unparsed
  request text) on the webhook request and throw `MISSING_RAW_BODY` without
  it. Custom adapters must capture the body as text before parsing; the
  built-in Next.js adapter already does.
- Mailgun and AIInbx reject webhooks whose signed timestamp is more than
  5 minutes from server time (replay protection).
- Outlook drivers using `webhookAuthResolver` without `autoSubscribeInbound`
  must set an explicit `webhookClientState` — construction throws otherwise,
  and Graph subscriptions created without that `clientState` must be
  recreated before notifications verify.

## Usage

```ts
import { EmailKit, MailgunDriver, ResendDriver } from "emailkit";

const emailkit = EmailKit({
  emailDrivers: [
    ResendDriver({
      id: "resend",
      apiKey: process.env.RESEND_API_KEY!,
      webhookSecret: process.env.RESEND_WEBHOOK_SECRET!,
    }),
    MailgunDriver({
      id: "mailgun",
      apiKey: process.env.MAILGUN_API_KEY!,
      webhookSigningKey: process.env.MAILGUN_WEBHOOK_SIGNING_KEY!,
    }),
  ],
  resolveEmailDriver: async (ctx) => {
    if (ctx.operation === "sendEmail") {
      return ctx.message.from.email.endsWith("@mg.example.com")
        ? { emailDriver: "mailgun" }
        : { emailDriver: "resend" };
    }

    if (
      "input" in ctx &&
      ctx.input &&
      "domain" in ctx.input &&
      ctx.input.domain?.endsWith("mg.example.com")
    ) {
      return { emailDriver: "mailgun" };
    }

    return "resend";
  },
  hooks: {
    email: {
      onInbound: async (event) => {
        console.log("Received inbound email:", event);
      },
      onOpened: async (event) => {
        console.log("Email opened:", event);
      },
    },
  },
});

await emailkit.sendEmail({
  from: { email: "sender@example.com", name: "Sender" },
  to: { email: "recipient@example.com" },
  subject: "Hello",
  text: "Hello world",
});
```

### EmailKit secret

Drivers that sign callback state or run OAuth mailbox connection flows need an EmailKit secret. EmailKit automatically reads `EMAILKIT_SECRET` when `secret` is omitted:

```sh
EMAILKIT_SECRET="replace-with-a-long-random-secret"
```

You can still pass `secret` explicitly, and it takes precedence over `EMAILKIT_SECRET`:

```ts
const emailkit = EmailKit({
  emailDrivers: [OutlookDriver({ clientId, clientSecret })],
  secret: process.env.AUTH_SECRET,
});
```

EmailKit also reads `PUBLIC_BASE_URL` automatically for public webhook and
callback URLs. With `PUBLIC_BASE_URL=https://app.example.com`, the default
driver route is `https://app.example.com/api/email/:emailDriverId`. Passing
`publicRoutes.baseUrl` or `publicRoutes.route` overrides those defaults.

### Mailboxes

Mailbox objects are durable identity only. Store provider OAuth material from
`hooks.mailbox.onConnected`, then pass it back as a separate `auth` value when
sending or managing mailbox-scoped webhooks:

```ts
import { EmailKit, OutlookDriver } from "emailkit";

const emailkit = EmailKit({
  secret: process.env.EMAILKIT_SECRET!,
  emailDrivers: [
    OutlookDriver({
      id: "outlook",
      clientId: process.env.OUTLOOK_CLIENT_ID!,
      clientSecret: process.env.OUTLOOK_CLIENT_SECRET!,
      webhookClientState: process.env.OUTLOOK_WEBHOOK_CLIENT_STATE!,
      webhookAuthResolver: async ({ mailbox }) => {
        return mailbox?.id ? await loadOutlookAuth(mailbox.id) : undefined;
      },
    }),
  ],
  hooks: {
    mailbox: {
      onConnected: async ({ mailbox, auth }) => {
        if (mailbox && auth) await saveOutlookAuth(mailbox.id, auth);
      },
    },
  },
});

await emailkit.mailboxes.connect({
  emailDriver: "outlook",
  email: "support@example.com",
});

const mailbox = await loadMailbox("support@example.com");
const auth = await loadOutlookAuth(mailbox.id);

await emailkit.sendEmail({
  from: { email: mailbox.email },
  to: { email: "recipient@example.com" },
  subject: "Hello",
  text: "Hello",
  sender: { emailDriver: "outlook", mailbox, auth },
});

await emailkit.mailboxes.webhooks.setup({
  emailDriver: "outlook",
  mailbox,
  auth,
  events: ["inbound"],
});
```

### Gmail

Gmail uses delegated Google OAuth mailbox auth. The same EmailKit mailbox,
send, webhook, sync, attachment, and `providerFetch` APIs work across Gmail and
other drivers; Gmail-specific setup is contained in the driver config.

Send-only Gmail needs a Google Cloud project with the Gmail API enabled, an
OAuth consent screen, and an OAuth client configured with your EmailKit
callback URL:

```sh
EMAILKIT_SECRET="replace-with-a-long-random-secret"
PUBLIC_BASE_URL="https://app.example.com"
GOOGLE_CLIENT_ID="..."
GOOGLE_CLIENT_SECRET="..."
```

With the default route template, register this OAuth redirect URI:

```txt
https://app.example.com/api/email/gmail
```

```ts
import { EmailKit, GmailDriver } from "emailkit";

export const emailkit = EmailKit({
  emailDrivers: [
    GmailDriver({
      id: "gmail",
      clientId: process.env.GOOGLE_CLIENT_ID!,
      clientSecret: process.env.GOOGLE_CLIENT_SECRET!,
    }),
  ],
  hooks: {
    mailbox: {
      onConnected: async ({ mailbox, auth }) => {
        if (mailbox && auth) await saveGmailAuth(mailbox.email, auth);
      },
      onAuthUpdated: async ({ mailbox, auth }) => {
        if (mailbox?.email) await saveGmailAuth(mailbox.email, auth);
      },
    },
  },
});
```

Connect a mailbox once, persist the returned auth in `onConnected`, then send
through the universal `sendEmail` API:

```ts
const { redirectUrl } = await emailkit.mailboxes.connect({
  emailDriver: "gmail",
  email: "support@example.com",
});

const mailbox = await loadMailbox("support@example.com");
const auth = await loadGmailAuth(mailbox.email);

await emailkit.sendEmail({
  from: { email: mailbox.email },
  to: { email: "recipient@example.com" },
  subject: "Hello from Gmail",
  text: "Same EmailKit API, Gmail underneath.",
  sender: { emailDriver: "gmail", mailbox, auth },
});
```

Inbound Gmail needs one extra Google Cloud path because Gmail publishes mailbox
changes through Cloud Pub/Sub:

1. Enable the Gmail API and Cloud Pub/Sub in the same Google Cloud project.
2. Create a topic, for example `projects/my-project/topics/emailkit-gmail`.
3. Grant `gmail-api-push@system.gserviceaccount.com` the Pub/Sub Publisher
   role on that topic.
4. Create a push subscription whose endpoint is your EmailKit webhook route
   with a token query parameter, for example
   `https://app.example.com/api/email/gmail?token=secret-token`.
5. Set the same topic and token on the driver:

```sh
GOOGLE_PUBSUB_TOPIC="projects/my-project/topics/emailkit-gmail"
GMAIL_WEBHOOK_TOKEN="secret-token"
```

```ts
GmailDriver({
  id: "gmail",
  clientId: process.env.GOOGLE_CLIENT_ID!,
  clientSecret: process.env.GOOGLE_CLIENT_SECRET!,
  pubsubTopic: process.env.GOOGLE_PUBSUB_TOPIC!,
  verificationToken: process.env.GMAIL_WEBHOOK_TOKEN!,
  autoSubscribeInbound: true,
  webhookAuthResolver: async ({ mailboxEmail }) => {
    return mailboxEmail ? await loadGmailAuth(mailboxEmail) : undefined;
  },
  onAuthUpdated: async ({ mailbox, auth }) => {
    if (mailbox?.email) await saveGmailAuth(mailbox.email, auth);
  },
});
```

Gmail push notifications only include the mailbox email and a history cursor.
EmailKit hydrates those notifications into normal `onInbound` events, but it
must be able to load and update the mailbox auth. Wire
`hooks.mailbox.onConnected`, `hooks.mailbox.onAuthUpdated`, and the Gmail
driver's `onAuthUpdated` to the same durable store. Gmail watches expire after
about 7 days; EmailKit renews active watches opportunistically and reports
`renewAfter` for scheduled refreshes.

### Outlook

Outlook uses delegated Microsoft Graph mailbox auth by default. For app-only
access limited by Exchange Online RBAC for Applications, set
`authMode: "application"` (see the Outlook provider docs). Register your EmailKit
route as a Web redirect URI in Microsoft Entra, then configure the driver:

```ts
import { EmailKit, OutlookDriver } from "emailkit";

export const emailkit = EmailKit({
  secret: process.env.EMAILKIT_SECRET!,
  emailDrivers: [
    OutlookDriver({
      id: "outlook",
      clientId: process.env.OUTLOOK_CLIENT_ID!,
      clientSecret: process.env.OUTLOOK_CLIENT_SECRET!,
      tenant: process.env.OUTLOOK_TENANT ?? "common",
      scopes: ["offline_access", "User.Read", "Mail.Send", "Mail.Read"],
      webhookClientState: process.env.OUTLOOK_WEBHOOK_CLIENT_STATE!,
    }),
  ],
  publicRoutes: {
    connectLandingRoutes: {
      success: "/settings/email",
      failure: "/settings/email",
    },
  },
});
```

The Outlook driver exposes only the send features it can map to Microsoft
Graph: CC, BCC, Reply-To addresses, attachments with content, and custom
headers that begin with `X-`. It does not expose templates, scheduling,
unsubscribe, tracking controls, tags, metadata, idempotency, or domain APIs.
Shared mailbox/send-as is not supported by the normalized driver; use
`providerFetch` for Graph-specific escape hatches.

Outlook webhook setup uses the `auth` you pass to create the Microsoft Graph
subscription, but live webhook parsing happens later in a separate request.
Configure `webhookAuthResolver` or `webhookAuth` so EmailKit can hydrate Graph
notifications when those webhooks arrive. Notifications without a matching
`clientState` are rejected with 401: the driver stamps a derived default onto
subscriptions it creates, and when you create subscriptions yourself
(`webhookAuthResolver` without `autoSubscribeInbound`) you must set
`webhookClientState` and include the same value in those subscriptions —
the driver throws at construction otherwise.

Reply threading is native (`nativeReplyThreading` capability): Microsoft Graph
cannot set `In-Reply-To`/`References` headers on outbound mail, so the driver
maps `reply.messageId` onto Graph `createReply` instead and lets Exchange wire
up conversation threading and reply headers itself. This requires
`sendEmailMode: "draft"` (Mail.ReadWrite), and the capability is derived from
the configuration: only a driver constructed with `sendEmailMode: "draft"`
advertises `nativeReplyThreading`, so `reply.messageId` is a compile error on
a driver using the default `"sendMail"` mode. A per-message
`provider.sendEmailMode: "sendMail"` override on a draft-configured driver
still throws `NOT_SUPPORTED` at runtime when combined with `reply.messageId`.
The driver looks the source message up by `internetMessageId` in the connected
mailbox, creates a reply draft, patches in the outgoing message, uploads
attachments, and sends it. The send result reports
`replyThreading: "applied"` on success, or `replyThreading: "skipped"` when
the source message was not found and the message was sent unthreaded instead
(an unthreaded reply beats a failed send). `reply.references` and
`reply.threadId` stay unsupported because Exchange derives the reference chain
itself. Custom `headers` cannot be combined with `reply.messageId`: Graph
`createReply` does not accept `internetMessageHeaders` and drafts cannot
receive them after creation, so the driver throws `NOT_SUPPORTED` instead of
risking a silently dropped header — send without `reply.messageId` or without
custom headers.

### Domains

Use domain names or provider ids. EmailKit resolves the provider-specific identifier for you:

```ts
const existing = await emailkit.domains.getOrNull({ domain: "mg.example.com" });

const { domain, created } = await emailkit.domains.ensure({
  emailDriver: "mailgun",
  domain: "mg.example.com",
});

console.log(domain.status, created);
```

`get()` stays strict and throws on missing domains. `getOrNull()` is the normalized lookup path when you want "find this domain if it exists".

Domain operation context is carried separately from the provider identifier and is surfaced in hooks:

```ts
await emailkit.domains.verify({
  domain: "mg.example.com",
  context: { tenantId: "tenant_123" },
});
```

Drivers can advertise method-level domain support:

```ts
const capabilities = {
  domains: {
    list: true,
    create: true,
    get: true,
    verify: true,
    delete: true,
    identifier: "domainId",
  },
} as const;
```

Unsupported domain methods are omitted from the typed facade where possible and throw `NOT_SUPPORTED` if called dynamically.

### Inbound attachments

For inbound email, treat `event.attachments` as metadata plus optional eager content. The stable retrieval path is `emailkit.attachments.getContent(...)`:

```ts
hooks: {
  email: {
    onInbound: async (event) => {
      for (const attachment of event.attachments ?? []) {
        const content = await emailkit.attachments.getContent(attachment);
        await saveAttachment(attachment.filename, content);
      }
    },
  },
}
```

If a driver can hydrate stored attachments before your hook runs, it will populate `attachment.content` eagerly. `getContent()` still works either way for drivers that declare `providerFetch`, so app code does not need provider-specific fetch/auth logic.

Inbound attachments are stamped with `attachment.emailDriver` by the EmailKit client. That means stored attachment records can be persisted and later passed back to `emailkit.attachments.getContent(attachment)` without also passing `{ emailDriver }`. If you do pass an explicit `{ emailDriver }`, it must match the attachment stamp.

### Sync (replay missed events)

If your server was down and webhooks were lost, ask the provider to replay the
window. Synced events are re-pulled from the provider and dispatched through
the same hooks as live webhooks (`onInbound`, `onDelivered`, ...), so nothing
else in your app changes. The only knob is `since`:

```ts
// Mailbox-scoped (Outlook): same mailbox + auth as other mailbox operations
const result = await emailkit.mailboxes.sync({
  emailDriver: "outlook",
  mailbox,
  auth,
  since: outageStartedAt,
});

// Domain-scoped (Mailgun): replays inbound and outbound tracking events
await emailkit.domains.sync({ emailDriver: "mailgun", domain: "mg.example.com", since });

// Account-scoped: Resend replays inbound email only; AIInbx replays inbound,
// delivery, tracking, and mailbox events
await emailkit.sync({ emailDriver: "resend", since });
await emailkit.sync({ emailDriver: "aiinbx", since });

console.log(result.dispatched, result.syncedFrom);
```

Semantics:

- Events are replayed oldest-first and are indistinguishable from live webhook
  events. Sync is at-least-once: handle inbound idempotently (upsert by
  `messageId`) and replays cost nothing.
- Event coverage is bounded by what the provider's API can list after the
  fact. Mailgun and AIInbx replay inbound and outbound tracking events; Outlook queries
  the mailbox message collection by default so moved or archived received mail
  can still replay, while skipping messages sent by the synced mailbox; Resend
  replays inbound email only, so `onDelivered`/`onOpened` webhooks
  missed during the outage are not recoverable there.
- `result.syncedFrom` is the earliest time the provider data actually covered.
  When provider retention cuts the window short (Mailgun stores events for a
  bounded window), `syncedFrom` is later than `since` — reported, never thrown.
  For Mailgun, set `eventsRetentionDays` in the driver config to match your
  plan (default 3, the documented minimum) so coverage isn't under-reported.
- Optional `until` bounds the window (exclusive), `signal` aborts long syncs,
  and `context` is surfaced on the `email.onAll` envelope for replayed events
  so you can suppress side effects like auto-replies during a replay.
- If a hook throws mid-replay, sync throws `EmailKitSyncError` carrying
  `dispatched` and `lastEventTimestamp`; resume with
  `since: error.lastEventTimestamp`.

Like every other facade, sync is typed by capability (`sync: { mailbox: true }`
etc.) — the method only exists on scopes your configured drivers support.

### Webhooks (framework-agnostic)

```ts
const handle = emailkit.handler();
const res = await handle({ method, headers, body, rawBody });
```

`rawBody` is the unparsed request text. Drivers that sign the request body
(Resend, AIInbx) require it and fail verification with a `MISSING_RAW_BODY`
error when a custom adapter omits it — capture the body as text before
parsing it.

Provider webhooks are verified before dispatch and rejected with 401 when
verification is not possible — there is no permissive fallback. Configure each
driver's webhook signing secret (`webhookSecret`, `webhookSigningKey`, Outlook
`webhookClientState`, or Gmail `verificationToken`) before accepting
production webhook traffic. Gmail additionally requires the Pub/Sub push
subscription endpoint URL to carry the same value as
`?token=<verificationToken>`; a Gmail driver configured for inbound webhooks
without a `verificationToken` throws at construction. Outlook derives a
default `clientState` for subscriptions the driver creates itself, so an
explicit `webhookClientState` is only required when subscriptions are created
outside the driver (`webhookAuthResolver` without `autoSubscribeInbound`) —
that misconfiguration also throws at construction. Drivers that sign the
request body (Resend, AIInbx) require `rawBody` on the webhook request; the
built-in Next.js adapter provides it automatically for every content type.
Mailgun and AIInbx additionally reject webhooks whose signed timestamp is
more than 5 minutes from server time (replay protection) — rejections are
logged so clock skew is distinguishable from a bad signature.

### Provider fetch helper

Call provider-specific endpoints while reusing the configured authentication and base URL. This advanced escape hatch is only surfaced when at least one configured driver declares `providerFetch`:

```ts
const response = await emailkit.providerFetch("/v4/domains", {
  emailDriver: "mailgun",
  method: "GET",
  searchParams: { limit: 25 },
});

if (!response.ok) throw new Error("Provider API failed");
const domains = await response.json();
```

The helper accepts any `fetch` options plus an optional `searchParams` object.
Use relative paths for provider APIs, or absolute URLs that stay inside the
provider API origin. Stored attachment downloads should go through
`emailkit.attachments.getContent(...)`; drivers avoid sending provider bearer
tokens to unrelated absolute URLs.

### Next.js adapter (optional)

Single-driver route:

```ts
import { createNextEmailKitHandler } from "emailkit/nextjs";

export const { GET, POST } = createNextEmailKitHandler(emailkit);
```

Multi-driver route, for example `app/api/email/[emailDriver]/route.ts`:

```ts
import { createNextEmailKitHandler } from "emailkit/nextjs";

export const { GET, POST } = createNextEmailKitHandler(emailkit, {
  emailDriver: async (_request, context) => {
    const params = await context.params;
    const emailDriver = params?.emailDriver;
    return typeof emailDriver === "string" ? emailDriver : undefined;
  },
});
```
