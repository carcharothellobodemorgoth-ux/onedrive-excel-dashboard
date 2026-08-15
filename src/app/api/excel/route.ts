import { auth } from "@/auth";
import { getWorksheetData, loadWorkbookSummary } from "@/lib/graph";
import { NextRequest, NextResponse } from "next/server";

async function requireGraphToken() {
  const session = await auth();
  if (!session?.user) {
    return {
      error: NextResponse.json({ error: "No autenticado" }, { status: 401 }),
    };
  }
  if (session.error || !session.accessToken) {
    return {
      error: NextResponse.json(
        {
          error:
            "Sesión de Microsoft expirada o incompleta. Cerrá sesión y volvé a entrar.",
          detail: session.error ?? "missing_access_token",
          code: "reauth_required",
        },
        { status: 401 },
      ),
    };
  }
  if (
    !session.accessToken.includes(".") ||
    session.accessToken.length < 20
  ) {
    return {
      error: NextResponse.json(
        {
          error:
            "Token de Microsoft inválido. Cerrá sesión y volvé a entrar.",
          code: "reauth_required",
        },
        { status: 401 },
      ),
    };
  }
  return { accessToken: session.accessToken };
}

export async function GET(request: NextRequest) {
  const gate = await requireGraphToken();
  if ("error" in gate) return gate.error;

  const worksheetId = request.nextUrl.searchParams.get("worksheetId");
  const itemIdParam = request.nextUrl.searchParams.get("itemId");
  const driveIdParam = request.nextUrl.searchParams.get("driveId");
  const shareUrl = request.nextUrl.searchParams.get("shareUrl");

  try {
    if (!worksheetId) {
      const summary = await loadWorkbookSummary(
        gate.accessToken,
        itemIdParam,
        driveIdParam,
        shareUrl,
      );
      if (!summary.ok) {
        return NextResponse.json(summary, { status: 409 });
      }
      return NextResponse.json(summary);
    }

    let itemId = itemIdParam;
    let driveId = driveIdParam;
    if (!itemId || !driveId) {
      const summary = await loadWorkbookSummary(
        gate.accessToken,
        null,
        null,
        shareUrl,
      );
      if (!summary.ok) {
        return NextResponse.json(summary, { status: 409 });
      }
      itemId = summary.itemId;
      driveId = summary.driveId;
    }

    const sheet = await getWorksheetData(
      gate.accessToken,
      driveId,
      itemId,
      worksheetId,
    );
    return NextResponse.json(sheet);
  } catch (err) {
    const message = err instanceof Error ? err.message : "Error Graph";
    const reauth =
      /InvalidAuthenticationToken|JWT is not well formed|401/i.test(message);
    return NextResponse.json(
      {
        error: reauth
          ? "Sesión de Microsoft inválida. Cerrá sesión y volvé a entrar."
          : message,
        code: reauth ? "reauth_required" : undefined,
      },
      { status: reauth ? 401 : 502 },
    );
  }
}

export async function POST(request: NextRequest) {
  const gate = await requireGraphToken();
  if ("error" in gate) return gate.error;

  try {
    const body = (await request.json()) as { shareUrl?: string };
    const summary = await loadWorkbookSummary(
      gate.accessToken,
      null,
      null,
      body.shareUrl ?? null,
    );
    if (!summary.ok) {
      return NextResponse.json(summary, { status: 409 });
    }
    return NextResponse.json(summary);
  } catch (err) {
    const message = err instanceof Error ? err.message : "Error Graph";
    const reauth =
      /InvalidAuthenticationToken|JWT is not well formed|401/i.test(message);
    return NextResponse.json(
      {
        error: reauth
          ? "Sesión de Microsoft inválida. Cerrá sesión y volvé a entrar."
          : message,
        code: reauth ? "reauth_required" : undefined,
      },
      { status: reauth ? 401 : 502 },
    );
  }
}
