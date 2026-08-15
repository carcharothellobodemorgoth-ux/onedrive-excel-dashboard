import NextAuth from "next-auth";
import MicrosoftEntraID from "next-auth/providers/microsoft-entra-id";

declare module "next-auth" {
  interface Session {
    accessToken?: string;
    error?: string;
  }
}

type TokenBag = {
  accessToken?: string;
  refreshToken?: string;
  expiresAt?: number;
  error?: string;
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
  // Personal MSA tenant GUID (not the "consumers" alias — OIDC issuer must match)
  "https://login.microsoftonline.com/9188040d-6c67-4c5b-b112-36a304b66dad/v2.0";

/** issuer ends with /v2.0; token endpoint is /{tenant}/oauth2/v2.0/token (not /v2.0/oauth2/...). */
function tokenUrlFromIssuer(issuer: string): string {
  const trimmed = issuer.replace(/\/$/, "");
  if (trimmed.endsWith("/v2.0")) {
    return `${trimmed.slice(0, -"/v2.0".length)}/oauth2/v2.0/token`;
  }
  return `${trimmed}/oauth2/v2.0/token`;
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
      const t = token as typeof token & TokenBag;

      if (account) {
        t.accessToken = account.access_token;
        t.refreshToken = account.refresh_token;
        t.expiresAt = account.expires_at;
        t.error = undefined;
        return t;
      }

      if (
        typeof t.expiresAt === "number" &&
        Date.now() < t.expiresAt * 1000 - 60_000 &&
        typeof t.accessToken === "string" &&
        t.accessToken.length > 20
      ) {
        return t;
      }

      if (!t.refreshToken || typeof t.refreshToken !== "string") {
        return {
          ...t,
          accessToken: undefined,
          error: "RefreshTokenMissing",
        };
      }

      try {
        const response = await fetch(tokenUrlFromIssuer(msaIssuer), {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: new URLSearchParams({
            client_id: process.env.AUTH_MICROSOFT_ENTRA_ID_ID!,
            client_secret: process.env.AUTH_MICROSOFT_ENTRA_ID_SECRET!,
            grant_type: "refresh_token",
            refresh_token: t.refreshToken,
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
            ...t,
            accessToken: undefined,
            error: `RefreshAccessTokenError:non-json:${response.status}`,
          };
        }

        if (!response.ok || !refreshed.access_token) {
          return {
            ...t,
            accessToken: undefined,
            error: `RefreshAccessTokenError:${
              refreshed.error_description ?? refreshed.error ?? response.status
            }`,
          };
        }

        return {
          ...t,
          accessToken: refreshed.access_token,
          refreshToken: refreshed.refresh_token ?? t.refreshToken,
          expiresAt: Math.floor(
            Date.now() / 1000 + (refreshed.expires_in ?? 3600),
          ),
          error: undefined,
        };
      } catch (e) {
        return {
          ...t,
          accessToken: undefined,
          error: `RefreshAccessTokenError:${
            e instanceof Error ? e.message : "unknown"
          }`,
        };
      }
    },
    async session({ session, token }) {
      const t = token as typeof token & TokenBag;
      session.accessToken =
        typeof t.accessToken === "string" ? t.accessToken : undefined;
      session.error = typeof t.error === "string" ? t.error : undefined;
      return session;
    },
  },
  pages: {
    signIn: "/",
  },
  trustHost: true,
});
