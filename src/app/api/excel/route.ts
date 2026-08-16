import { auth } from "@/auth";
import { getWorksheetData, loadWorkbookSummary } from "@/lib/graph";
import { NextRequest, NextResponse } from "next/server";

export const dynamic = "force-dynamic";
export const revalidate = 0;

const NO_STORE = {
  "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
  Pragma: "no-cache",
};

function json(data: unknown, init?: { status?: number }) {
  return NextResponse.json(data, {
    status: init?.status,
    headers: NO_STORE,
  });
}

async function requireGraphToken() {
  const session = await auth();
  if (!session?.user) {
    return {
      error: json({ error: "No autenticado" }, { status: 401 }),
    };
  }
  if (session.error || !session.accessToken) {
    return {
      error: json(
        {
          error:
            "Sesión de Microsoft expirada. Cerrá sesión y volvé a entrar.",
          detail: session.error ?? "missing_access_token",
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
        return json(summary, { status: 409 });
      }
      return json(summary);
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
        return json(summary, { status: 409 });
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
    return json(sheet);
  } catch (err) {
    const message = err instanceof Error ? err.message : "Error Graph";
    const reauth =
      /InvalidAuthenticationToken|JWT is not well formed|401/i.test(message);
    return json(
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
      return json(summary, { status: 409 });
    }
    return json(summary);
  } catch (err) {
    const message = err instanceof Error ? err.message : "Error Graph";
    const reauth =
      /InvalidAuthenticationToken|JWT is not well formed|401/i.test(message);
    return json(
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
