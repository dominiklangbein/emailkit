import { afterEach, describe, expect, it, vi } from "vitest";

import {
  EmailKit,
  OutlookDriver,
  type OutlookApplicationDriverConfig,
  type MailboxSyncInput,
  type OutlookSendEmailMode,
  type SyncStream,
} from "../src";

afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

const TOKEN_URL =
  "https://login.microsoftonline.com/tenant_123/oauth2/v2.0/token";
const GRAPH = "https://graph.microsoft.com/v1.0";
const WEBHOOK_URL = "https://app.example.com/api/email/outlook";
const MAILBOX = { id: "user_guid", email: "support@example.com" };

const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json" },
  });

const tokenResponse = () =>
  json({ access_token: "app_token_1", expires_in: 3600, token_type: "Bearer" });

const createAppDriver = <
  const TSendEmailMode extends OutlookSendEmailMode = "draft",
>(
  overrides: Partial<
    OutlookApplicationDriverConfig<"outlook", TSendEmailMode>
  > = {},
) =>
  OutlookDriver({
    clientId: "client_123",
    clientSecret: "secret_123",
    tenant: "tenant_123",
    authMode: "application",
    sendEmailMode: "draft" as TSendEmailMode,
    ...overrides,
  });

const message = {
  from: { email: "support@example.com" },
  to: { email: "customer@example.com" },
  subject: "Hello",
  text: "Hi there",
};

const calls = (fetchMock: ReturnType<typeof vi.fn>) =>
  fetchMock.mock.calls.map(([url, init]) => ({
    url: String(url),
    method: (init as RequestInit | undefined)?.method ?? "GET",
    authorization: new Headers((init as RequestInit | undefined)?.headers).get(
      "Authorization",
    ),
    body: (init as RequestInit | undefined)?.body,
  }));

const drain = async (stream: SyncStream) => {
  const events = [];
  while (true) {
    const next = await stream.next();
    if (next.done) return events;
    events.push(next.value);
  }
};

describe("OutlookDriver application auth", () => {
  it.each([undefined, "common", "Organizations", "consumers"])(
    "rejects application mode without a concrete tenant (%s)",
    (tenant) => {
      expect(() =>
        OutlookDriver({
          clientId: "client_123",
          clientSecret: "secret_123",
          tenant: tenant as string,
          authMode: "application",
        }),
      ).toThrow(expect.objectContaining({ code: "INVALID_CONFIG" }));
    },
  );

  it("advertises no mailbox connect, sender auth or EmailKit secret", () => {
    const driver = createAppDriver();

    expect(driver.capabilities).toMatchObject({
      mailboxConnect: false,
      senderAuth: false,
      senderMailbox: true,
      requiresSecret: false,
      nativeReplyThreading: true,
      publicRoutes: { webhook: true, lifecycleWebhook: true },
    });
    expect(driver.mailboxes).toBeUndefined();
    expect(driver.handleCallback).toBeUndefined();
  });

  it("sends a draft from users/{id} with the application token", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(tokenResponse())
      .mockResolvedValueOnce(json({ id: "draft_1", conversationId: "conv_1" }))
      .mockResolvedValueOnce(new Response(null, { status: 202 }));
    vi.stubGlobal("fetch", fetchMock);

    const result = await createAppDriver().sendEmail(message, {
      mailbox: MAILBOX,
    });

    const [token, create, send] = calls(fetchMock);
    expect(token!.url).toBe(TOKEN_URL);
    expect(Object.fromEntries(token!.body as URLSearchParams)).toMatchObject({
      grant_type: "client_credentials",
      scope: "https://graph.microsoft.com/.default",
    });
    expect(create).toMatchObject({
      url: `${GRAPH}/users/user_guid/messages`,
      method: "POST",
      authorization: "Bearer app_token_1",
    });
    expect(send).toMatchObject({
      url: `${GRAPH}/users/user_guid/messages/draft_1/send`,
      method: "POST",
      authorization: "Bearer app_token_1",
    });
    expect(result.messageId).toBe("draft_1");
  });

  it("derives the token scope from a custom graphBase", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(tokenResponse())
      .mockResolvedValueOnce(json({ id: "draft_1" }))
      .mockResolvedValueOnce(new Response(null, { status: 202 }));
    vi.stubGlobal("fetch", fetchMock);

    await createAppDriver({
      graphBase: "https://graph.microsoft.us/v1.0",
    }).sendEmail(message, { mailbox: MAILBOX });

    const [token, create] = calls(fetchMock);
    expect(Object.fromEntries(token!.body as URLSearchParams).scope).toBe(
      "https://graph.microsoft.us/.default",
    );
    expect(create!.url).toBe(
      "https://graph.microsoft.us/v1.0/users/user_guid/messages",
    );
  });

  it("falls back to the encoded email when id is empty", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(tokenResponse())
      .mockResolvedValueOnce(json({ id: "draft_1" }))
      .mockResolvedValueOnce(new Response(null, { status: 202 }));
    vi.stubGlobal("fetch", fetchMock);

    await createAppDriver().sendEmail(
      { ...message, from: { email: "sales+eu@example.com" } },
      { mailbox: { id: "", email: "sales+eu@example.com" } },
    );

    expect(calls(fetchMock)[1]!.url).toBe(
      `${GRAPH}/users/sales%2Beu%40example.com/messages`,
    );
  });

  it("threads replies through the addressed mailbox", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(tokenResponse())
      .mockResolvedValueOnce(
        json({
          value: [{ id: "src_1", internetMessageId: "<orig@example.com>" }],
        }),
      )
      .mockResolvedValueOnce(json({ id: "draft_1" }))
      .mockResolvedValueOnce(json({ id: "draft_1" }))
      .mockResolvedValueOnce(new Response(null, { status: 202 }));
    vi.stubGlobal("fetch", fetchMock);

    const result = await createAppDriver().sendEmail(
      { ...message, reply: { messageId: "<orig@example.com>" } },
      { mailbox: MAILBOX },
    );

    const [, lookup, createReply, patch, send] = calls(fetchMock);
    expect(
      lookup!.url.startsWith(`${GRAPH}/users/user_guid/messages?$filter=`),
    ).toBe(true);
    expect(createReply!.url).toBe(
      `${GRAPH}/users/user_guid/messages/src_1/createReply`,
    );
    expect(patch).toMatchObject({
      url: `${GRAPH}/users/user_guid/messages/draft_1`,
      method: "PATCH",
    });
    expect(send!.url).toBe(`${GRAPH}/users/user_guid/messages/draft_1/send`);
    expect(result.replyThreading).toBe("applied");
  });

  it("deletes failed reply drafts in the addressed mailbox", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(tokenResponse())
      .mockResolvedValueOnce(
        json({
          value: [{ id: "src_1", internetMessageId: "<orig@example.com>" }],
        }),
      )
      .mockResolvedValueOnce(json({ id: "draft_1" }))
      .mockResolvedValueOnce(json({ error: { message: "Bad patch" } }, 400))
      .mockResolvedValueOnce(new Response(null, { status: 204 }));
    vi.stubGlobal("fetch", fetchMock);

    await expect(
      createAppDriver().sendEmail(
        { ...message, reply: { messageId: "<orig@example.com>" } },
        { mailbox: MAILBOX },
      ),
    ).rejects.toMatchObject({ httpStatus: 400 });

    expect(calls(fetchMock)[4]).toMatchObject({
      url: `${GRAPH}/users/user_guid/messages/draft_1`,
      method: "DELETE",
      authorization: "Bearer app_token_1",
    });
  });

  it("uses users/{id}/sendMail in sendMail mode", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(tokenResponse())
      .mockResolvedValueOnce(new Response(null, { status: 202 }));
    vi.stubGlobal("fetch", fetchMock);

    await createAppDriver({ sendEmailMode: "sendMail" }).sendEmail(message, {
      mailbox: MAILBOX,
    });

    expect(calls(fetchMock)[1]).toMatchObject({
      url: `${GRAPH}/users/user_guid/sendMail`,
      authorization: "Bearer app_token_1",
    });
  });

  it("surfaces Graph 403 for mailboxes outside the RBAC scope", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(tokenResponse())
      .mockResolvedValueOnce(
        json(
          {
            error: {
              code: "ErrorAccessDenied",
              message: "Access is denied. Check credentials and try again.",
            },
          },
          403,
        ),
      );
    vi.stubGlobal("fetch", fetchMock);

    await expect(
      createAppDriver().sendEmail(message, { mailbox: MAILBOX }),
    ).rejects.toMatchObject({
      httpStatus: 403,
      message: expect.stringContaining("Access is denied"),
    });
    expect(calls(fetchMock).some((call) => call.method === "DELETE")).toBe(
      false,
    );
  });

  it("requires a mailbox email to sync before requesting a token", async () => {
    const fetchMock = vi.fn();
    vi.stubGlobal("fetch", fetchMock);
    const driver = createAppDriver();

    for (const input of [
      // Deliberately invalid input: no mailbox, id or email.
      {},
      // An id alone cannot tell received mail from the mailbox's sent mail.
      { mailbox: { id: "user_guid" } },
    ]) {
      await expect(
        drain(
          driver.sync!.mailbox!({
            since: new Date("2026-01-01"),
            ...input,
          } as MailboxSyncInput),
        ),
      ).rejects.toMatchObject({ code: "MISSING_MAILBOX" });
    }
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("sends from the from address when no mailbox is given", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(tokenResponse())
      .mockResolvedValueOnce(json({ id: "draft_1" }))
      .mockResolvedValueOnce(new Response(null, { status: 202 }));
    vi.stubGlobal("fetch", fetchMock);

    await createAppDriver().sendEmail({
      ...message,
      from: { email: "sales+eu@example.com" },
    });

    expect(calls(fetchMock)[1]!.url).toBe(
      `${GRAPH}/users/sales%2Beu%40example.com/messages`,
    );
  });

  it("scopes configured resources to the addressed mailbox", async () => {
    const subscription = (resource: string) =>
      json({
        id: "sub_1",
        resource,
        changeType: "created",
        notificationUrl: WEBHOOK_URL,
        expirationDateTime: "2026-12-01T00:00:00Z",
      });
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(tokenResponse())
      .mockResolvedValueOnce(subscription("x"))
      .mockResolvedValueOnce(tokenResponse())
      .mockResolvedValueOnce(subscription("x"))
      .mockResolvedValueOnce(subscription("x"));
    vi.stubGlobal("fetch", fetchMock);
    const setup = (
      driver: ReturnType<typeof createAppDriver>,
      provider?: Record<string, unknown>,
    ) =>
      driver.webhooks!.mailbox!.setup!({
        mailbox: MAILBOX,
        url: WEBHOOK_URL,
        events: ["inbound"],
        ...(provider ? { provider } : {}),
      });

    await setup(
      createAppDriver({ inboundResource: "me/mailFolders('Inbox')/messages" }),
    );
    const defaultDriver = createAppDriver();
    await setup(defaultDriver);
    await setup(defaultDriver, { resource: "users/other_guid/messages" });

    const resources = calls(fetchMock)
      .filter((call) => call.url === `${GRAPH}/subscriptions`)
      .map((call) => JSON.parse(String(call.body)).resource);
    expect(resources).toEqual([
      "users/user_guid/mailFolders('Inbox')/messages",
      "users/user_guid/messages",
      "users/other_guid/messages",
    ]);
  });

  it("accepts the same mailbox inputs for webhook setup as for sync", async () => {
    const fetchMock = vi.fn(async (input: unknown) =>
      String(input) === TOKEN_URL
        ? tokenResponse()
        : json({
            id: "sub_1",
            notificationUrl: WEBHOOK_URL,
            expirationDateTime: "2026-12-01T00:00:00Z",
          }),
    );
    vi.stubGlobal("fetch", fetchMock);
    const driver = createAppDriver();
    const base = { url: WEBHOOK_URL, events: ["inbound" as const] };

    await driver.webhooks!.mailbox!.setup!({
      ...base,
      email: "support@contoso.com",
    });
    await driver.webhooks!.mailbox!.setup!({ ...base, mailboxId: "user_guid" });

    const resources = calls(fetchMock)
      .filter((call) => call.url === `${GRAPH}/subscriptions`)
      .map((call) => JSON.parse(String(call.body)).resource);
    expect(resources).toEqual([
      "users/support%40contoso.com/messages",
      "users/user_guid/messages",
    ]);
  });

  it("derives a per-tenant default clientState in application mode", async () => {
    const fetchMock = vi.fn(async (input: unknown) =>
      String(input).endsWith("/token")
        ? tokenResponse()
        : json({
            id: "sub_1",
            notificationUrl: WEBHOOK_URL,
            expirationDateTime: "2026-12-01T00:00:00Z",
          }),
    );
    vi.stubGlobal("fetch", fetchMock);

    for (const tenant of ["tenant_a", "tenant_b"]) {
      await createAppDriver({ tenant }).webhooks!.mailbox!.setup!({
        mailbox: MAILBOX,
        url: WEBHOOK_URL,
        events: ["inbound"],
      });
    }

    const [a, b] = calls(fetchMock)
      .filter((call) => call.url === `${GRAPH}/subscriptions`)
      .map((call) => JSON.parse(String(call.body)).clientState);
    expect(a).toMatch(/^emailkit:/);
    expect(b).toMatch(/^emailkit:/);
    expect(a).not.toBe(b);
  });

  it("syncs the addressed mailbox", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(tokenResponse())
      .mockResolvedValueOnce(json({ value: [] }));
    vi.stubGlobal("fetch", fetchMock);

    await drain(
      createAppDriver().sync!.mailbox!({
        mailbox: MAILBOX,
        since: new Date("2026-01-01"),
      }),
    );

    const list = calls(fetchMock)[1]!;
    expect(new URL(list.url).pathname).toBe("/v1.0/users/user_guid/messages");
    expect(list.authorization).toBe("Bearer app_token_1");
  });

  it("hydrates webhook notifications with the application token", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(tokenResponse())
      .mockResolvedValueOnce(
        json({
          id: "msg_1",
          internetMessageId: "<msg_1@example.com>",
          subject: "Need help",
          body: { contentType: "text", content: "Hello" },
          from: { emailAddress: { address: "customer@example.com" } },
          toRecipients: [{ emailAddress: { address: "support@example.com" } }],
          receivedDateTime: "2026-05-21T10:30:00Z",
        }),
      );
    vi.stubGlobal("fetch", fetchMock);

    const event = await createAppDriver({
      webhookClientState: "state_123",
    }).handleWebhook!({
      method: "POST",
      headers: {},
      body: {
        value: [
          {
            subscriptionId: "sub_1",
            clientState: "state_123",
            changeType: "created",
            resource:
              "Users/11111111-2222-3333-4444-555555555555/Messages/AAMkAGRj-T7_KzowPTAAA=",
            resourceData: {
              "@odata.type": "#Microsoft.Graph.Message",
              "@odata.id":
                "Users/11111111-2222-3333-4444-555555555555/Messages/AAMkAGRj-T7_KzowPTAAA=",
              id: "msg_1",
            },
          },
        ],
      },
    });

    const hydrate = calls(fetchMock)[1]!;
    expect(
      hydrate.url.startsWith(
        `${GRAPH}/Users/11111111-2222-3333-4444-555555555555/Messages/AAMkAGRj-T7_KzowPTAAA=?`,
      ),
    ).toBe(true);
    expect(hydrate.authorization).toBe("Bearer app_token_1");
    expect(event).toMatchObject({
      type: "inbound",
      data: { messageId: "<msg_1@example.com>" },
    });
  });

  it.each([
    "users/victim/messages",
    "users/victim/messages/msg_1/attachments",
    "users/victim/messages/msg_1?$select=body",
    "users/victim/messages/msg_1#fragment",
    "users/victim/messages/..",
    "users/victim/messages/%2e%2e",
    "users/victim/messages/msg_1\\..",
    "https://attacker.example/users/victim/messages/msg_1",
  ])("rejects non-message webhook resource %s", async (resource) => {
    const fetchMock = vi.fn();
    vi.stubGlobal("fetch", fetchMock);

    await expect(
      createAppDriver({ webhookClientState: "state_123" }).handleWebhook!({
        method: "POST",
        headers: {},
        body: {
          value: [
            {
              subscriptionId: "sub_1",
              clientState: "state_123",
              changeType: "created",
              resource,
              resourceData: { "@odata.type": "#Microsoft.Graph.Message" },
            },
          ],
        },
      }),
    ).rejects.toMatchObject({ code: "INVALID_WEBHOOK_RESOURCE" });
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("uses the application token for providerFetch", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(tokenResponse())
      .mockResolvedValueOnce(json({ value: [] }));
    vi.stubGlobal("fetch", fetchMock);

    await createAppDriver().providerFetch!("users/user_guid/mailFolders");

    expect(calls(fetchMock)[1]).toMatchObject({
      url: `${GRAPH}/users/user_guid/mailFolders`,
      authorization: "Bearer app_token_1",
    });
  });

  it("keeps /me in delegated mode when a mailbox is passed", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(json({ id: "draft_1" }))
      .mockResolvedValueOnce(new Response(null, { status: 202 }));
    vi.stubGlobal("fetch", fetchMock);

    await OutlookDriver({
      clientId: "client_123",
      clientSecret: "secret_123",
      tenant: "common",
      sendEmailMode: "draft",
    }).sendEmail(message, {
      auth: { accessToken: "access_123", expiresAt: Date.now() + 120_000 },
      mailbox: MAILBOX,
    });

    expect(calls(fetchMock)[0]!.url).toBe(`${GRAPH}/me/messages`);
  });

  it("runs send, webhook setup and inbound through the EmailKit client", async () => {
    const fetchMock = vi.fn(async (input: unknown, init?: RequestInit) => {
      const url = String(input);
      if (url === TOKEN_URL) return tokenResponse();
      if (
        url === `${GRAPH}/users/user_guid/messages` &&
        init?.method === "POST"
      )
        return json({ id: "draft_1" });
      if (url.endsWith("/draft_1/send"))
        return new Response(null, { status: 202 });
      if (url === `${GRAPH}/subscriptions`)
        return json({
          id: "sub_1",
          resource: "users/user_guid/messages",
          changeType: "created",
          notificationUrl: "https://app.example.com/api/email/outlook-app",
          expirationDateTime: "2026-12-01T00:00:00Z",
        });
      if (
        url.startsWith(
          `${GRAPH}/Users/11111111-2222-3333-4444-555555555555/Messages/AAMkAGRj-T7_KzowPTAAA=?`,
        )
      )
        return json({
          id: "msg_1",
          internetMessageId: "<msg_1@example.com>",
          subject: "Need help",
          body: { contentType: "text", content: "Hello" },
          from: { emailAddress: { address: "customer@example.com" } },
          toRecipients: [{ emailAddress: { address: "support@example.com" } }],
          receivedDateTime: "2026-05-21T10:30:00Z",
        });
      throw new Error(`Unexpected fetch ${init?.method ?? "GET"} ${url}`);
    });
    vi.stubGlobal("fetch", fetchMock);
    const onCreated = vi.fn();
    const onInbound = vi.fn();
    const emailkit = EmailKit({
      emailDrivers: [
        OutlookDriver({
          id: "outlook-app",
          clientId: "client_123",
          clientSecret: "secret_123",
          tenant: "tenant_123",
          authMode: "application",
          sendEmailMode: "draft",
          webhookClientState: "state_123",
        }),
        OutlookDriver({
          id: "outlook",
          clientId: "client_123",
          clientSecret: "secret_123",
        }),
      ],
      secret: "emailkit-secret",
      publicRoutes: { baseUrl: "https://app.example.com" },
      resolveEmailDriver: () => "outlook-app",
      hooks: { webhook: { onCreated }, email: { onInbound } },
    });

    await emailkit.sendEmail({
      ...message,
      sender: { emailDriver: "outlook-app", mailbox: MAILBOX },
    });
    await emailkit.mailboxes.webhooks.setup({
      emailDriver: "outlook-app",
      mailbox: MAILBOX,
      events: ["inbound"],
    });
    await emailkit.handler()({
      method: "POST",
      headers: {},
      query: { emailDriver: "outlook-app" },
      body: {
        value: [
          {
            subscriptionId: "sub_1",
            clientState: "state_123",
            changeType: "created",
            resource:
              "Users/11111111-2222-3333-4444-555555555555/Messages/AAMkAGRj-T7_KzowPTAAA=",
            resourceData: { id: "msg_1" },
          },
        ],
      },
    });

    const subscribe = calls(fetchMock).find(
      (call) => call.url === `${GRAPH}/subscriptions`,
    )!;
    expect(JSON.parse(String(subscribe.body))).toMatchObject({
      resource: "users/user_guid/messages",
      notificationUrl: "https://app.example.com/api/email/outlook-app",
    });
    expect(onCreated).toHaveBeenCalledTimes(1);
    expect(onInbound).toHaveBeenCalledWith(
      expect.objectContaining({
        emailDriver: "outlook-app",
        messageId: "<msg_1@example.com>",
      }),
    );
    expect(
      calls(fetchMock).filter((call) => call.url === TOKEN_URL),
    ).toHaveLength(1);
  });
});
