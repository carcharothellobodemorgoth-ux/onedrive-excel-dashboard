import NextAuth from "next-auth";
import MicrosoftEntraID from "next-auth/providers/microsoft-entra-id";

declare module "next-auth" {
  interface Session {
    accessToken?: string;
    error?: string;
  }
}

type AppJWT = {
  refreshToken?: string;
  error?: string;
  sub?: string;
};

const scopes = [
  "openid",
  "profile",
  "email",
  "offline_access",
  "User.Read",
  "Files.Read",
  "Files.Read.All",
].join(" ");

const msaIssuer =
  process.env.AUTH_MICROSOFT_ENTRA_ID_ISSUER ??
  "https://login.microsoftonline.com/9188040d-6c67-4c5b-b112-36a304b66dad/v2.0";

/** Per-instance cache so large Graph access tokens never enter the session cookie. */
const accessCache = new Map<
  string,
  { accessToken: string; expiresAtMs: number }
>();

function tokenEndpoint(): string {
  return `${msaIssuer.replace(/\/$/, "")}/oauth2/v2.0/token`;
}

async function exchangeRefreshToken(
  refreshToken: string,
): Promise<
  | { ok: true; accessToken: string; refreshToken?: string; expiresIn: number }
  | { ok: false; error: string }
> {
  const response = await fetch(tokenEndpoint(), {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: process.env.AUTH_MICROSOFT_ENTRA_ID_ID!,
      client_secret: process.env.AUTH_MICROSOFT_ENTRA_ID_SECRET!,
      grant_type: "refresh_token",
      refresh_token: refreshToken,
      scope: scopes,
    }),
    cache: "no-store",
  });

  const raw = await response.text();
  let refreshed: {
    access_token?: string;
    refresh_token?: string;
    expires_in?: number;
    error?: string;
    error_description?: string;
  } = {};
  try {
    refreshed = raw ? (JSON.parse(raw) as typeof refreshed) : {};
  } catch {
    return {
      ok: false,
      error: `Token endpoint non-JSON (${response.status}): ${raw.slice(0, 120)}`,
    };
  }

  if (!response.ok || !refreshed.access_token) {
    return {
      ok: false,
      error:
        refreshed.error_description ??
        refreshed.error ??
        `Refresh failed (${response.status})`,
    };
  }

  return {
    ok: true,
    accessToken: refreshed.access_token,
    refreshToken: refreshed.refresh_token,
    expiresIn: refreshed.expires_in ?? 3600,
  };
}

function cacheAccess(
  sub: string,
  accessToken: string,
  expiresInSec: number,
): void {
  accessCache.set(sub, {
    accessToken,
    expiresAtMs: Date.now() + Math.max(60, expiresInSec - 60) * 1000,
  });
}

function cachedAccess(sub: string | undefined): string | undefined {
  if (!sub) return undefined;
  const hit = accessCache.get(sub);
  if (!hit) return undefined;
  if (Date.now() >= hit.expiresAtMs) {
    accessCache.delete(sub);
    return undefined;
  }
  return hit.accessToken;
}

export const { handlers, auth, signIn, signOut } = NextAuth({
  providers: [
    MicrosoftEntraID({
      clientId: process.env.AUTH_MICROSOFT_ENTRA_ID_ID!,
      clientSecret: process.env.AUTH_MICROSOFT_ENTRA_ID_SECRET!,
      issuer: msaIssuer,
      authorization: {
        params: {
          scope: scopes,
        },
      },
    }),
  ],
  callbacks: {
    async jwt({ token, account }) {
      const t = token as typeof token & AppJWT;
      // Keep only refreshToken in the cookie (MSA access tokens are huge and
      // truncate Auth.js cookies → Graph "JWT is not well formed").
      if (account) {
        if (account.refresh_token) {
          t.refreshToken = account.refresh_token;
        }
        if (account.access_token && t.sub) {
          const expiresIn = account.expires_at
            ? Math.max(60, account.expires_at - Math.floor(Date.now() / 1000))
            : 3600;
          cacheAccess(t.sub, account.access_token, expiresIn);
        }
        t.error = undefined;
        return t;
      }

      if (t.sub && cachedAccess(t.sub)) {
        t.error = undefined;
        return t;
      }

      if (!t.refreshToken || typeof t.refreshToken !== "string") {
        return { ...t, error: "RefreshTokenMissing" };
      }

      const exchanged = await exchangeRefreshToken(t.refreshToken);
      if (!exchanged.ok) {
        return { ...t, error: `RefreshAccessTokenError:${exchanged.error}` };
      }

      if (t.sub) {
        cacheAccess(t.sub, exchanged.accessToken, exchanged.expiresIn);
      }
      if (exchanged.refreshToken) {
        t.refreshToken = exchanged.refreshToken;
      }
      t.error = undefined;
      return t;
    },
    async session({ session, token }) {
      const t = token as typeof token & AppJWT;
      session.error =
        typeof t.error === "string" ? t.error : undefined;
      const access = cachedAccess(
        typeof t.sub === "string" ? t.sub : undefined,
      );
      session.accessToken = access;
      return session;
    },
  },
  pages: {
    signIn: "/",
  },
  trustHost: true,
});
