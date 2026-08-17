import { auth } from "@/auth";
import {
  appendGastoVarios,
  deleteGastoVarios,
  findWorksheetByName,
  getWorksheetData,
  loadWorkbookSummary,
  sumGastosVariosByQuincena,
  updateGastoVarios,
} from "@/lib/graph";
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

function graphErrorResponse(err: unknown) {
  const message = err instanceof Error ? err.message : "Error Graph";
  const reauth =
    /InvalidAuthenticationToken|JWT is not well formed|401|Access Denied|Forbidden|403/i.test(
      message,
    );
  const needWrite =
    /Access Denied|Forbidden|403|insufficient|Files\.ReadWrite/i.test(message);
  return json(
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
        return json(summary, { status: 409 });
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
      return json({
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

    return json({
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
      return json(
        { error: "La descripción es obligatoria" },
        { status: 400 },
      );
    }
    if (!category) {
      return json(
        { error: "La categoría es obligatoria" },
        { status: 400 },
      );
    }
    if (!Number.isFinite(amount) || amount === 0) {
      return json(
        { error: "El monto debe ser un número distinto de 0" },
        { status: 400 },
      );
    }
    if (!Number.isFinite(quincena) || quincena < 1 || quincena > 24) {
      return json(
        { error: "Quincena inválida (1–24)" },
        { status: 400 },
      );
    }

    let itemId = body.itemId ?? null;
    let driveId = body.driveId ?? null;
    if (!itemId || !driveId) {
      const summary = await loadWorkbookSummary(gate.accessToken, null, null);
      if (!summary.ok) {
        return json(summary, { status: 409 });
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

    return json({
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

type GastoBody = {
  description?: string;
  category?: string;
  amount?: number | string;
  quincena?: number | string;
  previousQuincena?: number | string;
  row?: number | string;
  itemId?: string;
  driveId?: string;
};

function parseGastoFields(body: GastoBody) {
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
  return { description, category, amount, quincena };
}

function gastoFieldError(fields: ReturnType<typeof parseGastoFields>) {
  if (!fields.description) {
    return json({ error: "La descripción es obligatoria" }, { status: 400 });
  }
  if (!fields.category) {
    return json({ error: "La categoría es obligatoria" }, { status: 400 });
  }
  if (!Number.isFinite(fields.amount) || fields.amount === 0) {
    return json(
      { error: "El monto debe ser un número distinto de 0" },
      { status: 400 },
    );
  }
  if (
    !Number.isFinite(fields.quincena) ||
    fields.quincena < 1 ||
    fields.quincena > 24
  ) {
    return json({ error: "Quincena inválida (1–24)" }, { status: 400 });
  }
  return null;
}

async function resolveWorkbookIds(
  accessToken: string,
  itemId: string | null | undefined,
  driveId: string | null | undefined,
) {
  if (itemId && driveId) {
    return { itemId, driveId };
  }
  const summary = await loadWorkbookSummary(accessToken, null, null);
  if (!summary.ok) {
    return { error: json(summary, { status: 409 }) };
  }
  return { itemId: summary.itemId, driveId: summary.driveId };
}

/** PATCH: update one gasto { row, description, category, amount, quincena }. */
export async function PATCH(request: NextRequest) {
  const gate = await requireGraphToken();
  if ("error" in gate) return gate.error;

  try {
    const body = (await request.json()) as GastoBody;
    const fields = parseGastoFields(body);
    const invalid = gastoFieldError(fields);
    if (invalid) return invalid;

    const row =
      typeof body.row === "number" ? body.row : Number(body.row);
    if (!Number.isFinite(row) || row < 2) {
      return json({ error: "Fila inválida" }, { status: 400 });
    }
    const previousQuincena =
      body.previousQuincena === undefined || body.previousQuincena === ""
        ? fields.quincena
        : typeof body.previousQuincena === "number"
          ? body.previousQuincena
          : Number(body.previousQuincena);

    const ids = await resolveWorkbookIds(
      gate.accessToken,
      body.itemId,
      body.driveId,
    );
    if ("error" in ids) return ids.error;

    const result = await updateGastoVarios(
      gate.accessToken,
      ids.driveId,
      ids.itemId,
      {
        row,
        description: fields.description,
        category: fields.category,
        amount: Math.abs(fields.amount),
        quincena: fields.quincena,
        previousQuincena: Number.isFinite(previousQuincena)
          ? previousQuincena
          : fields.quincena,
      },
    );

    return json({
      ok: true,
      worksheet: result.worksheet,
      row: result.row,
      itemId: ids.itemId,
      driveId: ids.driveId,
    });
  } catch (err) {
    return graphErrorResponse(err);
  }
}

/** DELETE: remove one gasto (?row=&quincena=). */
export async function DELETE(request: NextRequest) {
  const gate = await requireGraphToken();
  if ("error" in gate) return gate.error;

  try {
    const sp = request.nextUrl.searchParams;
    const row = Number(sp.get("row"));
    const quincenaRaw = sp.get("quincena");
    const quincena =
      quincenaRaw === null || quincenaRaw === ""
        ? undefined
        : Number(quincenaRaw);

    if (!Number.isFinite(row) || row < 2) {
      return json({ error: "Fila inválida" }, { status: 400 });
    }
    if (
      quincena !== undefined &&
      (!Number.isFinite(quincena) || quincena < 1 || quincena > 24)
    ) {
      return json({ error: "Quincena inválida (1–24)" }, { status: 400 });
    }

    const ids = await resolveWorkbookIds(
      gate.accessToken,
      sp.get("itemId"),
      sp.get("driveId"),
    );
    if ("error" in ids) return ids.error;

    const result = await deleteGastoVarios(
      gate.accessToken,
      ids.driveId,
      ids.itemId,
      row,
      quincena,
    );

    return json({
      ok: true,
      deleted: result.deleted,
      worksheet: result.worksheet,
      itemId: ids.itemId,
      driveId: ids.driveId,
    });
  } catch (err) {
    return graphErrorResponse(err);
  }
}
