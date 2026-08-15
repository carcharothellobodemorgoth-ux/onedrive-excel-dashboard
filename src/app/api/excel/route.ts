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

  const itemId = process.env.EXCEL_DRIVE_ITEM_ID;
  if (!itemId) {
    return NextResponse.json(
      { error: "EXCEL_DRIVE_ITEM_ID no configurado" },
      { status: 500 },
    );
  }

  const worksheetId = request.nextUrl.searchParams.get("worksheetId");

  try {
    if (!worksheetId) {
      const summary = await loadWorkbookSummary(session.accessToken, itemId);
      return NextResponse.json(summary);
    }

    const sheet = await getWorksheetData(
      session.accessToken,
      itemId,
      worksheetId,
    );
    return NextResponse.json(sheet);
  } catch (err) {
    const message = err instanceof Error ? err.message : "Error Graph";
    return NextResponse.json({ error: message }, { status: 502 });
  }
}
