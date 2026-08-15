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

export const { handlers, auth, signIn, signOut } = NextAuth({
  providers: [
    MicrosoftEntraID({
      clientId: process.env.AUTH_MICROSOFT_ENTRA_ID_ID!,
      clientSecret: process.env.AUTH_MICROSOFT_ENTRA_ID_SECRET!,
      issuer:
        process.env.AUTH_MICROSOFT_ENTRA_ID_ISSUER ??
        // Personal MSA tenant GUID (not the "consumers" alias — OIDC issuer must match)
        "https://login.microsoftonline.com/9188040d-6c67-4c5b-b112-36a304b66dad/v2.0",
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
        return t;
      }

      if (
        typeof t.expiresAt === "number" &&
        Date.now() < t.expiresAt * 1000 - 60_000
      ) {
        return t;
      }

      if (!t.refreshToken || typeof t.refreshToken !== "string") {
        return { ...t, error: "RefreshTokenMissing" };
      }

      try {
        const issuer =
          process.env.AUTH_MICROSOFT_ENTRA_ID_ISSUER ??
          "https://login.microsoftonline.com/9188040d-6c67-4c5b-b112-36a304b66dad/v2.0";
        const tokenUrl = `${issuer.replace(/\/$/, "")}/oauth2/v2.0/token`;
        const response = await fetch(tokenUrl, {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: new URLSearchParams({
            client_id: process.env.AUTH_MICROSOFT_ENTRA_ID_ID!,
            client_secret: process.env.AUTH_MICROSOFT_ENTRA_ID_SECRET!,
            grant_type: "refresh_token",
            refresh_token: t.refreshToken,
            scope: scopes,
          }),
        });

        const refreshed = (await response.json()) as {
          access_token?: string;
          refresh_token?: string;
          expires_in?: number;
        };

        if (!response.ok || !refreshed.access_token) {
          return { ...t, error: "RefreshAccessTokenError" };
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
      } catch {
        return { ...t, error: "RefreshAccessTokenError" };
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
