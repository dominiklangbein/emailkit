import { afterEach, describe, expect, it, vi } from "vitest";
import { createOAuthMailboxKit } from "../src/utils/oauth-mailbox";

const TOKEN_URL =
  "https://login.microsoftonline.com/tenant_123/oauth2/v2.0/token";
const SCOPE = "https://graph.microsoft.com/.default";

const tokenResponse = (accessToken: string) =>
  new Response(
    JSON.stringify({
      access_token: accessToken,
      expires_in: 3600,
      token_type: "Bearer",
    }),
    { status: 200, headers: { "Content-Type": "application/json" } },
  );

const createKit = () =>
  createOAuthMailboxKit({
    provider: "outlook",
    label: "Outlook",
    issuer: "Microsoft",
    clientId: "client_123",
    clientSecret: "secret_123",
    tokenEndpoint: TOKEN_URL,
  });

afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

describe("clientCredentialsAuth", () => {
  it("requests a client-credentials token for the given scope", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(tokenResponse("app_token_1"));
    vi.stubGlobal("fetch", fetchMock);

    const auth = await createKit().clientCredentialsAuth(SCOPE)();

    expect(auth.accessToken).toBe("app_token_1");
    const [url, init] = fetchMock.mock.calls[0];
    expect(url).toBe(TOKEN_URL);
    expect(Object.fromEntries(init.body as URLSearchParams)).toEqual({
      client_id: "client_123",
      client_secret: "secret_123",
      grant_type: "client_credentials",
      scope: SCOPE,
    });
  });

  it("reuses the cached token and refreshes inside the leeway", async () => {
    vi.useFakeTimers();
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(tokenResponse("app_token_1"))
      .mockResolvedValueOnce(tokenResponse("app_token_2"));
    vi.stubGlobal("fetch", fetchMock);
    const getAuth = createKit().clientCredentialsAuth(SCOPE);

    expect((await getAuth()).accessToken).toBe("app_token_1");
    expect((await getAuth()).accessToken).toBe("app_token_1");
    expect(fetchMock).toHaveBeenCalledTimes(1);

    vi.advanceTimersByTime((3600 - 30) * 1000);
    expect((await getAuth()).accessToken).toBe("app_token_2");
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it("shares one in-flight request", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(tokenResponse("app_token_1"));
    vi.stubGlobal("fetch", fetchMock);
    const getAuth = createKit().clientCredentialsAuth(SCOPE);

    const [a, b] = await Promise.all([getAuth(), getAuth()]);

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(a.accessToken).toBe("app_token_1");
    expect(b.accessToken).toBe("app_token_1");
  });

  it("propagates errors and retries on the next call", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            error: "invalid_client",
            error_description: "AADSTS7000215: Invalid client secret",
          }),
          { status: 401, headers: { "Content-Type": "application/json" } },
        ),
      )
      .mockResolvedValueOnce(tokenResponse("app_token_1"));
    vi.stubGlobal("fetch", fetchMock);
    const getAuth = createKit().clientCredentialsAuth(SCOPE);

    await expect(getAuth()).rejects.toMatchObject({
      httpStatus: 401,
      message: expect.stringContaining("AADSTS7000215"),
    });
    expect((await getAuth()).accessToken).toBe("app_token_1");
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });
});
