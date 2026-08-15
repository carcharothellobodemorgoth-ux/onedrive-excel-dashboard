import { auth } from "@/auth";
import { NextResponse } from "next/server";

export async function GET(request: Request) {
  const session = await auth();
  const { searchParams } = new URL(request.url);
  const callbackUrl = searchParams.get("callbackUrl") ?? "/dashboard";

  if (session?.user) {
    return NextResponse.redirect(new URL(callbackUrl, request.url));
  }

  // Redirect to Auth.js Microsoft provider
  const signInUrl = new URL("/api/auth/signin/microsoft-entra-id", request.url);
  signInUrl.searchParams.set("callbackUrl", callbackUrl);
  return NextResponse.redirect(signInUrl);
}
