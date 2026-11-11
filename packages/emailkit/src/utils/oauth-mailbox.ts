/**
 * Shared machinery for OAuth-delegated mailbox drivers (Gmail, Outlook).
 *
 * The providers differ only in endpoints, refresh-form extras, and the extra
 * fields they persist on the auth blob. Everything else — secret validation,
 * public-route resolution, token fetch/refresh, auth type guards, response
 * parsing — is provider-independent and lives here.
 * `createOAuthMailboxKit()` closes over the provider parameters and returns
 * the token/auth helpers; the stateless helpers are exported directly.
 */

import type {
  EmailDriverOperationOptions,
  ProviderFetchInit,
} from "../driver";
import type {
  Mailbox,
  MailboxIdentity,
  MailboxSyncInput,
  MailboxWebhookTarget,
  WebhookRequest,
} from "../types";
import { EmailKitError } from "../types";
import { decodeOAuthState, type OAuthStatePayload } from "./oauth-state";

export const TOKEN_REFRESH_LEEWAY_MS = 60_000;
const CLIENT_CREDENTIALS_TIMEOUT_MS = 30_000;

export interface OAuthMailboxAuth {
  accessToken: string;
  refreshToken?: string;
  expiresAt?: number;
  scopes?: string[];
  tokenType?: string;
}

export interface OAuthTokenResponse {
  access_token?: string;
  refresh_token?: string;
  expires_in?: number;
  /** Microsoft only. */
  ext_expires_in?: number;
  scope?: string;
  token_type?: string;
  error?: string;
  error_description?: string;
  [key: string]: unknown;
}

export const nonEmptyString = (value: unknown): string | undefined =>
  typeof value === "string" && value.trim() ? value : undefined;

export const normalizeBaseUrl = (url: string): string =>
  url.replace(/\/+$/, "");

export const formatScopes = (scopes: string[]): string => scopes.join(" ");

// Case-insensitive dedupe (first occurrence's casing wins): Microsoft treats
// scopes case-insensitively, and Google scope URLs are canonically lowercase,
// so two entries differing only in case are never two distinct scopes.
export const uniqueScopes = (scopes: string[]): string[] => {
  const seen = new Set<string>();
  return scopes.filter((scope) => {
    const key = scope.toLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
};

// Some endpoints return 204/202 with a JSON content-type but an empty body
// (e.g. Gmail users.stop), so parse from text instead of trusting the header.
export const readJsonResponse = async (
  response: Response,
): Promise<unknown> => {
  const text = await response.text();
  if (!text) return undefined;
  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("application/json")) {
    try {
      return JSON.parse(text);
    } catch {
      return text;
    }
  }
  return text;
};

/**
 * Extract a human-readable message from Google/Microsoft error bodies: Graph
 * and Gmail API errors nest `{ error: { message } }`, OAuth token endpoints
 * return `{ error, error_description }`.
 */
export const oauthErrorMessage = (body: unknown, fallback: string): string => {
  if (typeof body === "object" && body !== null) {
    const record = body as Record<string, unknown>;
    const error = record.error;
    if (typeof error === "object" && error !== null) {
      const message = (error as Record<string, unknown>).message;
      if (typeof message === "string") return message;
    }
    if (typeof error === "string" && error) {
      const description = record.error_description;
      return typeof description === "string" && description
        ? `${error}: ${description}`
        : error;
    }
    if (typeof record.error_description === "string" && record.error_description) {
      return record.error_description;
    }
    if (typeof record.message === "string") return record.message;
  }
  if (typeof body === "string" && body) return body;
  return fallback;
};

/** Token-response echo safe to expose on result.raw (no token material). */
export const sanitizeTokenResponse = (token: OAuthTokenResponse) => ({
  expiresIn: token.expires_in,
  ...(token.ext_expires_in !== undefined
    ? { extExpiresIn: token.ext_expires_in }
    : {}),
  scopes:
    typeof token.scope === "string"
      ? token.scope.split(/\s+/).filter(Boolean)
      : undefined,
  tokenType: token.token_type,
});

export const authorizationHeader = (auth: OAuthMailboxAuth): string =>
  `${auth.tokenType || "Bearer"} ${auth.accessToken}`;

export const isOAuthMailboxAuth = (auth: unknown): auth is OAuthMailboxAuth =>
  typeof auth === "object" &&
  auth !== null &&
  typeof (auth as OAuthMailboxAuth).accessToken === "string";

const UNSUPPORTED_OAUTH_MAILBOX_SEND_FIELDS = [
  "track",
  "tags",
  "metadata",
  "sendAt",
  "templateId",
  "templateData",
  "sandbox",
  "idempotencyKey",
] as const;

export const unsupportedOAuthMailboxSendFields = (
  message: object,
): string[] => {
  const record = message as Record<string, unknown>;
  return UNSUPPORTED_OAUTH_MAILBOX_SEND_FIELDS.filter(
    (field) => record[field] !== undefined,
  );
};

export const addNormalizedMailboxEmail = (
  emails: Set<string>,
  value: unknown,
): void => {
  if (typeof value === "string" && value.trim()) {
    emails.add(value.trim().toLowerCase());
  }
};

export const normalizedMailboxEmails = (
  input: MailboxSyncInput,
): Set<string> => {
  const emails = new Set<string>();
  addNormalizedMailboxEmail(emails, input.email);
  if (
    input.mailbox &&
    typeof input.mailbox === "object" &&
    !Array.isArray(input.mailbox)
  ) {
    addNormalizedMailboxEmail(emails, input.mailbox.email);
  }
  return emails;
};

export interface SyncMailboxProviderMetadata {
  mailboxEmail?: string;
  mailboxId?: string;
}

export const mailboxProviderMetadata = (
  input: MailboxWebhookTarget,
): SyncMailboxProviderMetadata => {
  const mailbox =
    input.mailbox &&
    typeof input.mailbox === "object" &&
    !Array.isArray(input.mailbox)
      ? input.mailbox
      : undefined;
  return {
    mailboxEmail: nonEmptyString(input.email) || nonEmptyString(mailbox?.email),
    mailboxId: nonEmptyString(input.mailboxId) || nonEmptyString(mailbox?.id),
  };
};

export const resolvePublicWebhookUrl = (
  options?: EmailDriverOperationOptions,
): string | undefined =>
  nonEmptyString(options?.publicRoutes?.webhook?.url);

/**
 * Resolve and validate a providerFetch path against the provider API base:
 * relative paths resolve under the base, absolute URLs must stay on it.
 */
export const resolveProviderFetchUrl = (
  base: string,
  path: string | URL,
  init: ProviderFetchInit | undefined,
  { provider, message }: { provider: string; message: string },
): URL => {
  const value = path instanceof URL ? path.toString() : path;
  const url = /^https?:\/\//i.test(value)
    ? new URL(value)
    : new URL(value.replace(/^\/+/, ""), `${base}/`);
  const baseUrl = new URL(`${base}/`);
  if (
    url.origin !== baseUrl.origin ||
    !url.pathname.startsWith(baseUrl.pathname)
  ) {
    throw new EmailKitError(
      message,
      provider,
      "INVALID_PROVIDER_FETCH_URL",
      400,
    );
  }

  if (init?.searchParams) {
    const params =
      init.searchParams instanceof URLSearchParams
        ? Array.from(init.searchParams.entries())
        : Object.entries(init.searchParams).flatMap(([key, value]) => {
            if (value === undefined || value === null) return [];
            if (Array.isArray(value)) {
              return value
                .filter((entry) => entry !== undefined && entry !== null)
                .map((entry) => [key, String(entry)] as const);
            }
            return [[key, String(value)] as const];
          });

    for (const [key] of params) url.searchParams.delete(key);
    for (const [key, value] of params) url.searchParams.append(key, value);
  }

  return url;
};

export interface OAuthMailboxKitOptions<TAuth extends OAuthMailboxAuth> {
  /** EmailKitError provider id, e.g. "gmail". */
  provider: string;
  /** Driver label used in error messages, e.g. "Gmail". */
  label: string;
  /** Token issuer label used in error messages, e.g. "Google". */
  issuer: string;
  clientId: string;
  clientSecret: string;
  tokenEndpoint: string;
  /** Extra form fields for refresh_token grants (e.g. Microsoft `scope`). */
  refreshForm?: (auth: TAuth) => Record<string, string>;
  /**
   * Graft provider-specific auth fields (e.g. the Gmail history cursor) from
   * the previous auth blob onto a freshly minted one.
   */
  mergeAuth?: (auth: OAuthMailboxAuth, previous?: TAuth) => TAuth;
}

export const createOAuthMailboxKit = <TAuth extends OAuthMailboxAuth>({
  provider,
  label,
  issuer,
  clientId,
  clientSecret,
  tokenEndpoint,
  refreshForm,
  mergeAuth = (auth) => auth as TAuth,
}: OAuthMailboxKitOptions<TAuth>) => {
  const isAuth = (auth: unknown): auth is TAuth => isOAuthMailboxAuth(auth);

  const requireSecret = (
    secret: string | undefined,
    operation: string,
  ): string => {
    if (!secret) {
      throw new EmailKitError(
        `EmailKit secret is required for ${label} ${operation}`,
        provider,
        "MISSING_SECRET",
      );
    }
    return secret;
  };

  const resolveConnectCallbackUrl = (
    input: { callbackUrl?: unknown },
    options?: EmailDriverOperationOptions,
  ): string => {
    const callbackUrl =
      nonEmptyString(input.callbackUrl) ||
      nonEmptyString(options?.publicRoutes?.callback?.url);

    if (!callbackUrl) {
      throw new EmailKitError(
        `${label} mailbox connect requires callbackUrl from EmailKit public routes`,
        provider,
        "MISSING_CALLBACK_URL",
      );
    }

    return callbackUrl;
  };

  const parseCallbackRequest = (
    request: WebhookRequest,
  ): { code: string; state: string } => {
    const queryError = request.query?.error;
    if (queryError) {
      throw new EmailKitError(
        request.query?.error_description || queryError,
        provider,
        queryError,
      );
    }

    const code = request.query?.code;
    if (!code) {
      throw new EmailKitError(
        `Missing ${label} OAuth code`,
        provider,
        "MISSING_CODE",
      );
    }
    const state = request.query?.state;
    if (!state) {
      throw new EmailKitError(
        `Missing ${label} OAuth state`,
        provider,
        "MISSING_STATE",
      );
    }
    return { code, state };
  };

  const decodeState = <T extends OAuthStatePayload>(
    state: string,
    secret: string,
  ): T => decodeOAuthState<T>(state, secret, { provider, label });

  const fetchToken = async (
    form: URLSearchParams,
    signal?: AbortSignal,
  ): Promise<OAuthTokenResponse> => {
    const response = await fetch(tokenEndpoint, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: form,
      signal,
    });
    const body = await readJsonResponse(response);
    if (!response.ok) {
      throw new EmailKitError(
        oauthErrorMessage(body, `${issuer} token request failed`),
        provider,
        undefined,
        response.status,
        undefined,
        body,
      );
    }
    return body as OAuthTokenResponse;
  };

  const toAuth = (token: OAuthTokenResponse, previous?: TAuth): TAuth => {
    if (!token.access_token) {
      throw new EmailKitError(
        `${issuer} token response did not include an access token`,
        provider,
        "INVALID_TOKEN_RESPONSE",
        undefined,
        undefined,
        token,
      );
    }

    return mergeAuth(
      {
        accessToken: token.access_token,
        refreshToken: token.refresh_token || previous?.refreshToken,
        expiresAt:
          typeof token.expires_in === "number"
            ? Date.now() + token.expires_in * 1000
            : previous?.expiresAt,
        scopes:
          typeof token.scope === "string"
            ? token.scope.split(/\s+/).filter(Boolean)
            : previous?.scopes,
        tokenType: token.token_type || previous?.tokenType || "Bearer",
      },
      previous,
    );
  };

  const refreshAuth = async (
    auth: TAuth,
    signal?: AbortSignal,
  ): Promise<TAuth> => {
    if (!auth.refreshToken) {
      throw new EmailKitError(
        `${label} access token is expired and no refresh token was provided`,
        provider,
        "MISSING_REFRESH_TOKEN",
      );
    }

    const form = new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      grant_type: "refresh_token",
      refresh_token: auth.refreshToken,
      ...(refreshForm ? refreshForm(auth) : {}),
    });
    const token = await fetchToken(form, signal);
    return toAuth(token, auth);
  };

  /** Refresh when the token expires within the leeway; undefined when fresh. */
  const refreshAuthIfNeeded = async (
    auth: TAuth,
    signal?: AbortSignal,
  ): Promise<TAuth | undefined> =>
    typeof auth.expiresAt === "number" &&
    auth.expiresAt <= Date.now() + TOKEN_REFRESH_LEEWAY_MS
      ? refreshAuth(auth, signal)
      : undefined;

  const resolveMailboxOperationAuth = async (
    operation: string,
    input: { auth?: unknown; mailbox?: MailboxIdentity | Mailbox },
    options?: EmailDriverOperationOptions,
    signal?: AbortSignal,
  ): Promise<TAuth> => {
    const inputAuth = isAuth(input.auth) ? input.auth : undefined;
    const optionsAuth = isAuth(options?.auth) ? options.auth : undefined;
    let auth = inputAuth || optionsAuth;
    if (!auth) {
      throw new EmailKitError(
        `${label} ${operation} requires mailbox auth with an accessToken`,
        provider,
        "MISSING_AUTH",
      );
    }

    const refreshed = await refreshAuthIfNeeded(auth, signal);
    if (refreshed) {
      const mailbox = "mailbox" in input ? input.mailbox : options?.mailbox;
      await options?.onAuthUpdated?.({
        auth: refreshed,
        previousAuth: auth,
        ...(mailbox ? { mailbox } : {}),
        ...(options?.context !== undefined ? { context: options.context } : {}),
      });
      auth = refreshed;
    }

    return auth;
  };

  /**
   * App-only (client-credentials) token getter. The token is cached and
   * concurrent callers share one request, so the request is bounded by its own
   * timeout instead of any single caller's signal.
   */
  const clientCredentialsAuth = (scope: string): (() => Promise<TAuth>) => {
    const form = new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      grant_type: "client_credentials",
      scope,
    });
    let cached: TAuth | undefined;
    let inflight: Promise<TAuth> | undefined;
    return () => {
      if (
        cached?.expiresAt &&
        cached.expiresAt > Date.now() + TOKEN_REFRESH_LEEWAY_MS
      ) {
        return Promise.resolve(cached);
      }
      inflight ??= fetchToken(
        form,
        AbortSignal.timeout(CLIENT_CREDENTIALS_TIMEOUT_MS),
      )
        .then((token) => (cached = toAuth(token)))
        .finally(() => {
          inflight = undefined;
        });
      return inflight;
    };
  };

  return {
    isAuth,
    requireSecret,
    resolveConnectCallbackUrl,
    parseCallbackRequest,
    decodeState,
    fetchToken,
    toAuth,
    refreshAuth,
    refreshAuthIfNeeded,
    resolveMailboxOperationAuth,
    clientCredentialsAuth,
  };
};

export type OAuthMailboxKit<TAuth extends OAuthMailboxAuth> = ReturnType<
  typeof createOAuthMailboxKit<TAuth>
>;
