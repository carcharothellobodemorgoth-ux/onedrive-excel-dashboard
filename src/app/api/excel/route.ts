import { auth } from "@/auth";
import { getWorksheetData, loadWorkbookSummary } from "@/lib/graph";
import { NextRequest, NextResponse } from "next/server";

export async function GET(request: NextRequest) {
  const session = await auth();
  if (!session?.accessToken) {
    return NextResponse.json({ error: "No autenticado" }, { status: 401 });
  }
  if (session.error) {
    return NextResponse.json(
      { error: "Sesión expirada", detail: session.error },
      { status: 401 },
    );
  }

  const worksheetId = request.nextUrl.searchParams.get("worksheetId");
  const itemIdParam = request.nextUrl.searchParams.get("itemId");
  const driveIdParam = request.nextUrl.searchParams.get("driveId");

  try {
    if (!worksheetId) {
      const summary = await loadWorkbookSummary(
        session.accessToken,
        itemIdParam,
        driveIdParam,
      );
      if (!summary.ok) {
        return NextResponse.json(summary, { status: 409 });
      }
      return NextResponse.json(summary);
    }

    let itemId = itemIdParam;
    let driveId = driveIdParam;
    if (!itemId || !driveId) {
      const summary = await loadWorkbookSummary(session.accessToken);
      if (!summary.ok) {
        return NextResponse.json(summary, { status: 409 });
      }
      itemId = summary.itemId;
      driveId = summary.driveId;
    }

    const sheet = await getWorksheetData(
      session.accessToken,
      driveId,
      itemId,
      worksheetId,
    );
    return NextResponse.json(sheet);
  } catch (err) {
    const message = err instanceof Error ? err.message : "Error Graph";
    return NextResponse.json({ error: message }, { status: 502 });
  }
}
