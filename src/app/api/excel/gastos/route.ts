import { auth } from "@/auth";
import {
  appendGastoVarios,
  findWorksheetByName,
  getWorksheetData,
  loadWorkbookSummary,
  sumGastosVariosByQuincena,
} from "@/lib/graph";
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

function graphErrorResponse(err: unknown) {
  const message = err instanceof Error ? err.message : "Error Graph";
  const reauth =
    /InvalidAuthenticationToken|JWT is not well formed|401|Access Denied|Forbidden|403/i.test(
      message,
    );
  const needWrite =
    /Access Denied|Forbidden|403|insufficient|Files\.ReadWrite/i.test(message);
  return NextResponse.json(
    {
      error: reauth
        ? needWrite
          ? "Faltan permisos de escritura o la sesión es vieja. Cerrá sesión, volvé a entrar y aceptá Files.ReadWrite."
          : "Sesión de Microsoft inválida. Cerrá sesión y volvé a entrar."
        : message,
      code: reauth ? "reauth_required" : undefined,
    },
    { status: reauth ? 401 : 502 },
  );
}

/** GET: sums by quincena (+ optional sheet matrix). */
export async function GET(request: NextRequest) {
  const gate = await requireGraphToken();
  if ("error" in gate) return gate.error;

  const itemId = request.nextUrl.searchParams.get("itemId");
  const driveId = request.nextUrl.searchParams.get("driveId");

  try {
    let resolvedItemId = itemId;
    let resolvedDriveId = driveId;
    if (!resolvedItemId || !resolvedDriveId) {
      const summary = await loadWorkbookSummary(gate.accessToken, null, null);
      if (!summary.ok) {
        return NextResponse.json(summary, { status: 409 });
      }
      resolvedItemId = summary.itemId;
      resolvedDriveId = summary.driveId;
    }

    const existing = await findWorksheetByName(
      gate.accessToken,
      resolvedDriveId,
      resolvedItemId,
      "Gastos Varios",
    );

    if (!existing) {
      return NextResponse.json({
        ok: true,
        worksheet: null,
        byQuincena: Object.fromEntries(
          Array.from({ length: 24 }, (_, i) => [i + 1, 0]),
        ),
        rows: [],
        itemId: resolvedItemId,
        driveId: resolvedDriveId,
      });
    }

    const sheet = await getWorksheetData(
      gate.accessToken,
      resolvedDriveId,
      resolvedItemId,
      existing.id,
    );
    const byQuincena = sumGastosVariosByQuincena(sheet.rows);

    return NextResponse.json({
      ok: true,
      worksheet: existing,
      byQuincena,
      rows: sheet.rows,
      itemId: resolvedItemId,
      driveId: resolvedDriveId,
    });
  } catch (err) {
    return graphErrorResponse(err);
  }
}

/** POST: append one gasto { description, category, amount, quincena }. */
export async function POST(request: NextRequest) {
  const gate = await requireGraphToken();
  if ("error" in gate) return gate.error;

  try {
    const body = (await request.json()) as {
      description?: string;
      category?: string;
      amount?: number | string;
      quincena?: number | string;
      itemId?: string;
      driveId?: string;
    };

    const description = String(body.description ?? "").trim();
    const category = String(body.category ?? "").trim();
    const amount =
      typeof body.amount === "number"
        ? body.amount
        : Number(String(body.amount ?? "").replace(",", "."));
    const quincena =
      typeof body.quincena === "number"
        ? body.quincena
        : Number(body.quincena);

    if (!description) {
      return NextResponse.json(
        { error: "La descripción es obligatoria" },
        { status: 400 },
      );
    }
    if (!category) {
      return NextResponse.json(
        { error: "La categoría es obligatoria" },
        { status: 400 },
      );
    }
    if (!Number.isFinite(amount) || amount === 0) {
      return NextResponse.json(
        { error: "El monto debe ser un número distinto de 0" },
        { status: 400 },
      );
    }
    if (!Number.isFinite(quincena) || quincena < 1 || quincena > 24) {
      return NextResponse.json(
        { error: "Quincena inválida (1–24)" },
        { status: 400 },
      );
    }

    let itemId = body.itemId ?? null;
    let driveId = body.driveId ?? null;
    if (!itemId || !driveId) {
      const summary = await loadWorkbookSummary(gate.accessToken, null, null);
      if (!summary.ok) {
        return NextResponse.json(summary, { status: 409 });
      }
      itemId = summary.itemId;
      driveId = summary.driveId;
    }

    const result = await appendGastoVarios(
      gate.accessToken,
      driveId,
      itemId,
      {
        description,
        category,
        amount: Math.abs(amount),
        quincena,
      },
    );

    return NextResponse.json({
      ok: true,
      worksheet: result.worksheet,
      row: result.row,
      itemId,
      driveId,
    });
  } catch (err) {
    return graphErrorResponse(err);
  }
}
