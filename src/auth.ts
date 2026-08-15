import NextAuth from "next-auth";
import MicrosoftEntraID from "next-auth/providers/microsoft-entra-id";
import { getToken } from "next-auth/jwt";
import { cookies, headers } from "next/headers";

declare module "next-auth" {
  interface Session {
    error?: string;
  }
}

declare module "next-auth/jwt" {
  interface JWT {
    refreshToken?: string;
    error?: string;
  }
}

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

function tokenEndpoint(): string {
  return `${msaIssuer.replace(/\/$/, "")}/oauth2/v2.0/token`;
}

async function exchangeRefreshToken(
  refreshToken: string,
): Promise<
  | { ok: true; accessToken: string; refreshToken?: string }
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

  const refreshed = (await response.json()) as {
    access_token?: string;
    refresh_token?: string;
    error?: string;
    error_description?: string;
  };

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
  };
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
      // Never persist Graph access_token in the session cookie — MSA tokens are large
      // and truncate the Auth.js JWT (→ Graph IDX14100 "JWT is not well formed").
      if (account?.refresh_token) {
        token.refreshToken = account.refresh_token;
        token.error = undefined;
      }
      if ("accessToken" in token) {
        delete (token as Record<string, unknown>).accessToken;
      }
      if ("expiresAt" in token) {
        delete (token as Record<string, unknown>).expiresAt;
      }
      return token;
    },
    async session({ session, token }) {
      session.error =
        typeof token.error === "string" ? token.error : undefined;
      return session;
    },
  },
  pages: {
    signIn: "/",
  },
  trustHost: true,
});

/**
 * Fresh Microsoft Graph access token for the current user.
 * Uses only the refresh_token stored in the Auth.js JWT cookie.
 */
export async function getGraphAccessToken(): Promise<
  { ok: true; accessToken: string } | { ok: false; status: number; error: string }
> {
  const headerList = await headers();
  const cookieStore = await cookies();
  const cookieHeader = cookieStore
    .getAll()
    .map((c) => `${c.name}=${c.value}`)
    .join("; ");

  const jwt = await getToken({
    // Minimal req shape for next-auth/jwt in App Router
    req: {
      headers: {
        cookie: cookieHeader,
        ...Object.fromEntries(headerList.entries()),
      },
    } as Parameters<typeof getToken>[0]["req"],
    secret: process.env.AUTH_SECRET,
    secureCookie:
      process.env.AUTH_URL?.startsWith("https://") ||
      process.env.NODE_ENV === "production",
  });

  const refreshToken =
    typeof jwt?.refreshToken === "string" ? jwt.refreshToken : null;

  if (!refreshToken) {
    return {
      ok: false,
      status: 401,
      error:
        "Sesión sin permiso de OneDrive. Cerrá sesión y volvé a entrar con Microsoft.",
    };
  }

  try {
    const exchanged = await exchangeRefreshToken(refreshToken);
    if (!exchanged.ok) {
      return {
        ok: false,
        status: 401,
        error: `Sesión expirada (${exchanged.error}). Cerrá sesión y volvé a entrar.`,
      };
    }

    if (
      !exchanged.accessToken.includes(".") ||
      exchanged.accessToken.length < 20
    ) {
      return {
        ok: false,
        status: 401,
        error:
          "Token de Microsoft inválido. Cerrá sesión y volvé a entrar.",
      };
    }

    return { ok: true, accessToken: exchanged.accessToken };
  } catch (e) {
    return {
      ok: false,
      status: 401,
      error: e instanceof Error ? e.message : "No se pudo renovar el token",
    };
  }
}
