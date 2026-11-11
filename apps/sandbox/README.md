# EmailKit Sandbox

Local Next.js sandbox for exercising EmailKit drivers.

## Run

```bash
bun dev
```

The app runs on `http://localhost:3210`.

## Outlook

Required:

```bash
OUTLOOK_CLIENT_ID="..."
OUTLOOK_CLIENT_SECRET="..."
EMAILKIT_SECRET="replace-with-a-long-random-secret"
```

Optional:

```bash
DATABASE_URL="file:./dev.db"
OUTLOOK_TENANT="common"
PUBLIC_BASE_URL="https://your-public-url.example.com"
OUTLOOK_AUTO_SUBSCRIBE_INBOUND="true"
OUTLOOK_WEBHOOK_CLIENT_STATE="optional-shared-client-state"
OUTLOOK_SCOPES="offline_access User.Read Mail.Send Mail.Read"
FROM_EMAIL_ADDRESS="connected-mailbox@example.com"
TO_EMAIL_ADDRESS="recipient@example.com"
```

EmailKit reads `PUBLIC_BASE_URL` automatically and uses `/api/email/:emailDriverId` as the default public route. The sandbox UI also uses `PUBLIC_BASE_URL` for display; if it is omitted there, the UI falls back to `APP_URL`, then `http://localhost:3210`.

In Microsoft Entra, add `${PUBLIC_BASE_URL}/api/email/outlook` as a Web redirect URI for the app registration. Microsoft Graph inbound webhooks also require that URL to be publicly reachable HTTPS.

## Outlook application mode

The `outlook-app` driver uses the app's own token and the mailboxes an admin granted through Exchange Online RBAC (see the Outlook provider docs). There is no connect step. List the granted mailboxes instead:

```bash
OUTLOOK_CLIENT_ID="..."
OUTLOOK_CLIENT_SECRET="..."
OUTLOOK_APP_TENANT="contoso.onmicrosoft.com" # a tenant id or domain, not "common"
OUTLOOK_APP_MAILBOXES="support@contoso.com,sales@contoso.com"
```

These mailboxes appear in the sync and webhook pickers, and a send from one of them goes through `outlook-app`. Inbound webhooks arrive on `${PUBLIC_BASE_URL}/api/email/outlook-app`; no redirect URI is needed.
