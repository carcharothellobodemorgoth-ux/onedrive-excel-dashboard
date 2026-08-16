const GRAPH = "https://graph.microsoft.com/v1.0";

export type DriveItem = {
  id: string;
  name: string;
  webUrl?: string;
  lastModifiedDateTime?: string;
  parentReference?: { driveId?: string; path?: string };
};

export type ResolvedWorkbook = {
  itemId: string;
  driveId: string;
  item: DriveItem;
  strategy: string;
};

export type Worksheet = {
  id: string;
  name: string;
  position: number;
};

export type SheetData = {
  worksheet: Worksheet;
  headers: string[];
  rows: (string | number | boolean | null)[][];
};

export type ResolveResult =
  | ({ ok: true } & ResolvedWorkbook)
  | { ok: false; candidates: DriveItem[]; message: string };

async function graphFetch<T>(
  accessToken: string,
  path: string,
  init?: RequestInit,
): Promise<T> {
  const res = await fetch(`${GRAPH}${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/json",
      ...(init?.headers ?? {}),
    },
    cache: "no-store",
  });

  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Graph ${res.status}: ${body.slice(0, 500)}`);
  }

  if (res.status === 204) {
    return undefined as T;
  }
  const text = await res.text();
  if (!text) return undefined as T;
  return JSON.parse(text) as T;
}

export const GASTOS_VARIOS_SHEET = "Gastos Varios";
/** A=descripción, B=categoría, C+=quincenas (misma convención que 20262027). */
export const GASTOS_VARIOS_DATA_START = 2;
export const GASTOS_VARIOS_QUINCENAS = 24;

function colLetter(index0: number): string {
  let n = index0 + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

async function withWorkbookSession<T>(
  accessToken: string,
  driveId: string,
  itemId: string,
  fn: (sessionHeaders: Record<string, string>) => Promise<T>,
): Promise<T> {
  const session = await graphFetch<{ id: string }>(
    accessToken,
    itemPath(driveId, itemId, "/workbook/createSession"),
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ persistChanges: true }),
    },
  );
  const sessionHeaders = { "workbook-session-id": session.id };
  try {
    return await fn(sessionHeaders);
  } finally {
    try {
      await graphFetch(
        accessToken,
        itemPath(driveId, itemId, "/workbook/closeSession"),
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            ...sessionHeaders,
          },
          body: JSON.stringify({}),
        },
      );
    } catch {
      /* ignore close errors */
    }
  }
}

async function graphFetchOk<T>(
  accessToken: string,
  path: string,
): Promise<T | null> {
  const res = await fetch(`${GRAPH}${path}`, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/json",
    },
    cache: "no-store",
  });
  if (!res.ok) return null;
  return res.json() as Promise<T>;
}

function encodePath(path: string): string {
  // Graph path addressing: /me/drive/root:/folder/file.xlsx
  return path
    .split("/")
    .filter(Boolean)
    .map((p) => encodeURIComponent(p))
    .join("/");
}

function itemPath(driveId: string, itemId: string, suffix = ""): string {
  return `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}${suffix}`;
}

function asResolved(item: DriveItem, strategy: string): ResolvedWorkbook | null {
  const driveId = item.parentReference?.driveId;
  if (!item.id || !driveId) return null;
  return { itemId: item.id, driveId, item, strategy };
}

async function getPrimaryDriveId(accessToken: string): Promise<string | null> {
  const drive = await graphFetchOk<{ id: string }>(accessToken, `/me/drive?$select=id`);
  return drive?.id ?? null;
}

async function tryPath(
  accessToken: string,
  relativePath: string,
  strategy: string,
): Promise<ResolvedWorkbook | null> {
  const encoded = encodePath(relativePath);
  const item = await graphFetchOk<DriveItem>(
    accessToken,
    `/me/drive/root:/${encoded}?$select=id,name,webUrl,lastModifiedDateTime,parentReference`,
  );
  if (!item) return null;
  const resolved = asResolved(item, strategy);
  if (resolved) return resolved;

  // Fallback: attach primary drive id if parentReference missing
  const driveId = await getPrimaryDriveId(accessToken);
  if (!driveId || !item.id) return null;
  return {
    itemId: item.id,
    driveId,
    item: { ...item, parentReference: { driveId } },
    strategy: `${strategy}+me-drive`,
  };
}

async function tryDriveItem(
  accessToken: string,
  driveId: string,
  itemId: string,
  strategy: string,
): Promise<ResolvedWorkbook | null> {
  const item = await graphFetchOk<DriveItem>(
    accessToken,
    itemPath(
      driveId,
      itemId,
      "?$select=id,name,webUrl,lastModifiedDateTime,parentReference",
    ),
  );
  if (!item?.id) return null;
  const resolved = asResolved(item, strategy);
  if (resolved) return resolved;
  return {
    itemId: item.id,
    driveId,
    item: { ...item, parentReference: { ...(item.parentReference ?? {}), driveId } },
    strategy,
  };
}

function isExcelName(name?: string): boolean {
  return /\.(xlsx|xlsm|xls)$/i.test(name ?? "") || /proyecci[oó]n/i.test(name ?? "");
}

type SharedDriveItem = DriveItem & {
  remoteItem?: {
    id?: string;
    name?: string;
    webUrl?: string;
    lastModifiedDateTime?: string;
    file?: unknown;
    folder?: unknown;
    parentReference?: { driveId?: string; path?: string };
  };
};

/** sharedWithMe returns a local stub; workbook APIs need remoteItem id + driveId. */
function normalizeSharedItem(raw: SharedDriveItem): DriveItem | null {
  const remote = raw.remoteItem;
  if (remote?.id && remote.parentReference?.driveId) {
    return {
      id: remote.id,
      name: remote.name ?? raw.name,
      webUrl: remote.webUrl ?? raw.webUrl,
      lastModifiedDateTime:
        remote.lastModifiedDateTime ?? raw.lastModifiedDateTime,
      parentReference: remote.parentReference,
    };
  }
  if (raw.id && raw.parentReference?.driveId) {
    return raw;
  }
  return null;
}

async function collectSharedExcel(
  accessToken: string,
): Promise<DriveItem[]> {
  const out: DriveItem[] = [];
  const paths = [
    `/me/drive/sharedWithMe?$top=100&$select=id,name,webUrl,lastModifiedDateTime,remoteItem,file,folder`,
    `/me/drive/sharedWithMe?allowexternal=true&$top=100`,
  ];
  for (const path of paths) {
    const data = await graphFetchOk<{ value: SharedDriveItem[] }>(
      accessToken,
      path,
    );
    for (const raw of data?.value ?? []) {
      const name = raw.remoteItem?.name ?? raw.name;
      if (!isExcelName(name) && !isExcelName(raw.name)) continue;
      const normalized = normalizeSharedItem(raw);
      if (normalized) out.push(normalized);
    }
  }
  return out;
}

async function collectExcelCandidates(
  accessToken: string,
): Promise<DriveItem[]> {
  const folder = process.env.EXCEL_FOLDER?.trim() || "Proyecciones";
  const hint = process.env.EXCEL_FILE_NAME?.trim() || "PROYECCIÓN 26-27";
  const pools: DriveItem[] = [];
  const paths = [
    `/me/drive/root:/${encodePath(folder)}:/children?$top=200&$select=id,name,webUrl,lastModifiedDateTime,file,parentReference`,
    `/me/drive/root/search(q='${hint.replace(/'/g, "''")}')?$top=25&$select=id,name,webUrl,lastModifiedDateTime,file,parentReference`,
    `/me/drive/recent?$top=50&$select=id,name,webUrl,lastModifiedDateTime,file,parentReference`,
    `/me/drive/root/search(q='.xlsx')?$top=50&$select=id,name,webUrl,lastModifiedDateTime,file,parentReference`,
  ];

  for (const path of paths) {
    const data = await graphFetchOk<{ value: DriveItem[] }>(accessToken, path);
    if (data?.value) pools.push(...data.value);
  }

  pools.push(...(await collectSharedExcel(accessToken)));

  const byId = new Map<string, DriveItem>();
  for (const it of pools) {
    if (it?.id && isExcelName(it.name)) byId.set(it.id, it);
  }
  return [...byId.values()].sort((a, b) =>
    (b.lastModifiedDateTime ?? "").localeCompare(a.lastModifiedDateTime ?? ""),
  );
}

function toShareToken(url: string): string {
  const base64 = Buffer.from(url.trim(), "utf8")
    .toString("base64")
    .replace(/=+$/g, "")
    .replace(/\+/g, "-")
    .replace(/\//g, "_");
  return `u!${base64}`;
}

/**
 * Resolve a sharing / open link to driveId+itemId.
 * Works for "anyone with the link" shares that never appear in sharedWithMe.
 */
export async function resolveFromShareUrl(
  accessToken: string,
  shareUrl: string,
): Promise<ResolveResult> {
  const trimmed = shareUrl.trim();
  if (!trimmed) {
    return {
      ok: false,
      candidates: [],
      message: "Pegá un link de Compartir / OneDrive.",
    };
  }

  const urls = Array.from(
    new Set([
      trimmed,
      // common variants
      trimmed.split("#")[0],
      trimmed.replace("/edit.aspx", "/redir.aspx"),
    ]),
  );

  let lastError = "";
  for (const url of urls) {
    const token = toShareToken(url);
    const item = await graphFetchOk<DriveItem>(
      accessToken,
      `/shares/${token}/driveItem?$select=id,name,webUrl,lastModifiedDateTime,parentReference`,
    );
    if (item?.id && item.parentReference?.driveId) {
      return {
        ok: true,
        itemId: item.id,
        driveId: item.parentReference.driveId,
        item,
        strategy: "share-url",
      };
    }

    // Some shares need an extra hop
    const shareMeta = await graphFetchOk<{
      error?: { message?: string };
    }>(accessToken, `/shares/${token}`);
    if (shareMeta && "error" in (shareMeta as object)) {
      lastError = JSON.stringify(shareMeta).slice(0, 200);
    }

    const encodedPath = `/shares/${token}/driveItem`;
    try {
      const forced = await graphFetch<DriveItem>(accessToken, `${encodedPath}?$select=id,name,webUrl,lastModifiedDateTime,parentReference`);
      if (forced.id && forced.parentReference?.driveId) {
        return {
          ok: true,
          itemId: forced.id,
          driveId: forced.parentReference.driveId,
          item: forced,
          strategy: "share-url-forced",
        };
      }
    } catch (e) {
      lastError = e instanceof Error ? e.message : String(e);
    }
  }

  return {
    ok: false,
    candidates: [],
    message: `No pude abrir ese link con Graph. En Excel: Compartir → Copiar vínculo (no la URL del navegador doc.aspx). Detalle: ${lastError.slice(0, 180)}`,
  };
}

export async function resolveExcelItemId(
  accessToken: string,
  preferredItemId?: string | null,
  preferredDriveId?: string | null,
): Promise<ResolveResult> {
  const folder = process.env.EXCEL_FOLDER?.trim() || "Proyecciones";
  const fileName =
    process.env.EXCEL_FILE_NAME?.trim() || "PROYECCIÓN 26-27";
  const configuredPath =
    process.env.EXCEL_FILE_PATH?.trim() ||
    `${folder}/${fileName.replace(/\.xlsx$/i, "")}.xlsx`;
  const needle = fileName.replace(/\.xlsx$/i, "").toLowerCase();
  const configuredShare =
    process.env.EXCEL_SHARE_URL?.trim() ||
    "https://1drv.ms/x/c/d883d00740abbada/IQDVZuw8RO2ZR4CJqAH8k6nyAStmn-NqhF2_OUAcEu3D3_M?e=76bLdx";

  // 0) Preferred selection from UI (may be a shared remote item)
  if (preferredItemId && preferredDriveId) {
    const hit = await tryDriveItem(
      accessToken,
      preferredDriveId,
      preferredItemId,
      "user-selected",
    );
    if (hit) return { ok: true, ...hit };
  }

  // 1) Configured share link (works for "anyone with the link")
  const fromShare = await resolveFromShareUrl(accessToken, configuredShare);
  if (fromShare.ok) return fromShare;

  // 2) Shared with me
  const shared = await collectSharedExcel(accessToken);
  const sharedHit = shared.find((c) =>
    (c.name ?? "").toLowerCase().includes(needle),
  ) ?? shared.find((c) =>
    /26[\s_-]?27/.test(c.name ?? "") && /proyecci/i.test(c.name ?? ""),
  );
  if (sharedHit?.parentReference?.driveId) {
    const hit = await tryDriveItem(
      accessToken,
      sharedHit.parentReference.driveId,
      sharedHit.id,
      "sharedWithMe",
    );
    if (hit) return { ok: true, ...hit };
    return {
      ok: true,
      itemId: sharedHit.id,
      driveId: sharedHit.parentReference.driveId,
      item: sharedHit,
      strategy: "sharedWithMe-direct",
    };
  }

  // 3) Explicit path on own drive
  const pathVariants = Array.from(
    new Set([
      configuredPath,
      `${folder}/${fileName}`,
      `${folder}/${fileName}.xlsx`,
      `${folder}/PROYECCIÓN 26-27.xlsx`,
      `${folder}/PROYECCION 26-27.xlsx`,
      `Proyecciones/PROYECCIÓN 26-27.xlsx`,
    ]),
  );

  for (const p of pathVariants) {
    const hit = await tryPath(accessToken, p, `path:${p}`);
    if (hit) return { ok: true, ...hit };
  }

  // 4) Own drive + shared candidates — pick by name
  const candidates = await collectExcelCandidates(accessToken);
  const named =
    candidates.find((c) => (c.name ?? "").toLowerCase().includes(needle)) ??
    candidates.find(
      (c) =>
        /26[\s_-]?27/.test(c.name ?? "") && /proyecci/i.test(c.name ?? ""),
    );
  if (named?.parentReference?.driveId) {
    return {
      ok: true,
      itemId: named.id,
      driveId: named.parentReference.driveId,
      item: named,
      strategy: "name-match",
    };
  }

  return {
    ok: false,
    candidates: [...shared, ...candidates].filter(
      (c, i, arr) => arr.findIndex((x) => x.id === c.id) === i,
    ),
    message:
      fromShare.message ||
      "No encontré PROYECCIÓN 26-27. Pegá el link 1drv.ms en el campo de share o elegí un archivo de la lista.",
  };
}

export async function listWorksheets(
  accessToken: string,
  driveId: string,
  itemId: string,
): Promise<Worksheet[]> {
  const data = await graphFetch<{ value: Worksheet[] }>(
    accessToken,
    itemPath(
      driveId,
      itemId,
      "/workbook/worksheets?$select=id,name,position&$orderby=position",
    ),
  );
  return data.value ?? [];
}

type UsedRange = {
  values?: unknown[][];
};

function cellValue(raw: unknown): string | number | boolean | null {
  if (raw === null || raw === undefined || raw === "") return null;
  if (typeof raw === "number" || typeof raw === "boolean") return raw;
  if (typeof raw === "string") {
    const trimmed = raw.trim();
    if (trimmed === "") return null;
    const asNum = Number(trimmed.replace(/,/g, ""));
    if (trimmed !== "" && !Number.isNaN(asNum) && /^-?\d/.test(trimmed)) {
      return asNum;
    }
    return trimmed;
  }
  return String(raw);
}

export async function getWorksheetData(
  accessToken: string,
  driveId: string,
  itemId: string,
  worksheetId: string,
): Promise<SheetData> {
  const worksheet = await graphFetch<Worksheet>(
    accessToken,
    itemPath(
      driveId,
      itemId,
      `/workbook/worksheets/${encodeURIComponent(worksheetId)}?$select=id,name,position`,
    ),
  );

  const range = await graphFetch<UsedRange>(
    accessToken,
    itemPath(
      driveId,
      itemId,
      `/workbook/worksheets/${encodeURIComponent(worksheetId)}/usedRange(valuesOnly=true)?$select=values`,
    ),
  );

  const values = range.values ?? [];
  if (values.length === 0) {
    return { worksheet, headers: [], rows: [] };
  }

  // Full matrix: Excel row 1 = rows[0]. Column A = labels; other cols = quincenas.
  // Do NOT treat the first row as headers (rows 1–3 are ingresos).
  const width = Math.max(...values.map((r) => r.length), 1);
  const headers = Array.from({ length: width }, (_, i) =>
    i === 0 ? "Imputación" : `Q${i}`,
  );
  const rows = values.map((row) =>
    headers.map((_, i) => cellValue(row[i])),
  );

  return { worksheet, headers, rows };
}

export async function loadWorkbookSummary(
  accessToken: string,
  preferredItemId?: string | null,
  preferredDriveId?: string | null,
  shareUrl?: string | null,
): Promise<
  | ({
      ok: true;
      worksheets: Worksheet[];
    } & ResolvedWorkbook)
  | { ok: false; candidates: DriveItem[]; message: string }
> {
  if (shareUrl?.trim()) {
    const fromShare = await resolveFromShareUrl(accessToken, shareUrl);
    if (!fromShare.ok) return fromShare;
    const worksheets = await listWorksheets(
      accessToken,
      fromShare.driveId,
      fromShare.itemId,
    );
    return { ...fromShare, worksheets };
  }

  const resolved = await resolveExcelItemId(
    accessToken,
    preferredItemId,
    preferredDriveId,
  );
  if (!resolved.ok) return resolved;

  const worksheets = await listWorksheets(
    accessToken,
    resolved.driveId,
    resolved.itemId,
  );
  return { ...resolved, worksheets };
}

export async function findWorksheetByName(
  accessToken: string,
  driveId: string,
  itemId: string,
  name: string,
): Promise<Worksheet | null> {
  const sheets = await listWorksheets(accessToken, driveId, itemId);
  return (
    sheets.find((w) => w.name.trim().toLowerCase() === name.trim().toLowerCase()) ??
    null
  );
}

async function writeRangeValues(
  accessToken: string,
  driveId: string,
  itemId: string,
  worksheetId: string,
  address: string,
  values: (string | number | boolean | null)[][],
  sessionHeaders?: Record<string, string>,
): Promise<void> {
  await graphFetch(
    accessToken,
    itemPath(
      driveId,
      itemId,
      `/workbook/worksheets/${encodeURIComponent(worksheetId)}/range(address='${address}')`,
    ),
    {
      method: "PATCH",
      headers: {
        "Content-Type": "application/json",
        ...(sessionHeaders ?? {}),
      },
      body: JSON.stringify({ values }),
    },
  );
}

/**
 * Ensures the "Gastos Varios" sheet exists (creates + header row if missing).
 * Layout: A=Descripción, B=Categoría, C+=quincenas 1..24.
 */
export async function ensureGastosVariosSheet(
  accessToken: string,
  driveId: string,
  itemId: string,
): Promise<Worksheet> {
  const existing = await findWorksheetByName(
    accessToken,
    driveId,
    itemId,
    GASTOS_VARIOS_SHEET,
  );
  if (existing) return existing;

  return withWorkbookSession(accessToken, driveId, itemId, async (sessionHeaders) => {
    const created = await graphFetch<Worksheet>(
      accessToken,
      itemPath(driveId, itemId, "/workbook/worksheets"),
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          ...sessionHeaders,
        },
        body: JSON.stringify({ name: GASTOS_VARIOS_SHEET }),
      },
    );

    const lastCol = colLetter(
      GASTOS_VARIOS_DATA_START + GASTOS_VARIOS_QUINCENAS - 1,
    );
    const header: (string | number | null)[] = ["Descripción", "Categoría"];
    for (let q = 1; q <= GASTOS_VARIOS_QUINCENAS; q++) {
      header.push(`Q${q}`);
    }
    await writeRangeValues(
      accessToken,
      driveId,
      itemId,
      created.id,
      `A1:${lastCol}1`,
      [header],
      sessionHeaders,
    );
    return created;
  });
}

export type AppendGastoInput = {
  description: string;
  category: string;
  amount: number;
  /** Quincena 1-based (1..24) */
  quincena: number;
};

/**
 * Appends one expense row to "Gastos Varios" (creates the sheet if needed).
 */
export async function appendGastoVarios(
  accessToken: string,
  driveId: string,
  itemId: string,
  input: AppendGastoInput,
): Promise<{ worksheet: Worksheet; row: number }> {
  const worksheet = await ensureGastosVariosSheet(accessToken, driveId, itemId);
  const q = Math.min(
    Math.max(Math.floor(input.quincena), 1),
    GASTOS_VARIOS_QUINCENAS,
  );
  const amountCol = GASTOS_VARIOS_DATA_START + q - 1;
  const width = GASTOS_VARIOS_DATA_START + GASTOS_VARIOS_QUINCENAS;
  const lastColLetter = colLetter(width - 1);

  return withWorkbookSession(accessToken, driveId, itemId, async (sessionHeaders) => {
    const range = await graphFetch<UsedRange>(
      accessToken,
      itemPath(
        driveId,
        itemId,
        `/workbook/worksheets/${encodeURIComponent(worksheet.id)}/usedRange(valuesOnly=true)?$select=values`,
      ),
      { headers: sessionHeaders },
    );
    const used = range.values ?? [];
    const nextRow = Math.max(used.length + 1, 2); // row 1 = header

    const rowValues: (string | number | null)[] = Array.from(
      { length: width },
      () => null,
    );
    rowValues[0] = input.description.trim();
    rowValues[1] = input.category.trim();
    rowValues[amountCol] = Math.abs(input.amount);

    await writeRangeValues(
      accessToken,
      driveId,
      itemId,
      worksheet.id,
      `A${nextRow}:${lastColLetter}${nextRow}`,
      [rowValues],
      sessionHeaders,
    );

    return { worksheet, row: nextRow };
  });
}

/**
 * True when column A looks like a real expense line (not headers / balance row).
 * User layout: rows 1–2 headers, row 3 "Lo que queda", row 4+ expenses.
 */
export function isGastosVariosExpenseRow(
  row: (string | number | boolean | null)[] | undefined,
): boolean {
  if (!row) return false;
  const label = String(row[0] ?? "")
    .trim()
    .toLowerCase();
  if (!label) return false;
  if (
    /^(descripci[oó]n|categoria|categoría|lo que queda|resta|total|mes|agosto|septiembre|octubre|noviembre|diciembre|enero|febrero|marzo|abril|mayo|junio|julio|1era|2da|1ª|2ª|q\d+)$/i.test(
      label,
    )
  ) {
    return false;
  }
  // Month-only header cells sometimes land in col A when merged oddly
  if (/^(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)\b/i.test(label)) {
    return false;
  }
  return true;
}

function cellToAbsNumber(cell: string | number | boolean | null | undefined): number {
  if (typeof cell === "number" && Number.isFinite(cell)) return Math.abs(cell);
  if (typeof cell === "string" && cell.trim()) {
    const cleaned = cell
      .replace(/\$/g, "")
      .replace(/\s/g, "")
      .replace(/\./g, "")
      .replace(",", ".");
    const parsed = Number(cleaned);
    if (Number.isFinite(parsed)) return Math.abs(parsed);
  }
  return 0;
}

/**
 * Sum absolute amounts per quincena from a Gastos Varios matrix (full usedRange).
 * Only counts expense rows — skips headers and the "Lo que queda" mirror row.
 */
export function sumGastosVariosByQuincena(
  rows: (string | number | boolean | null)[][],
): Record<number, number> {
  const out: Record<number, number> = {};
  for (let q = 1; q <= GASTOS_VARIOS_QUINCENAS; q++) out[q] = 0;
  if (!rows.length) return out;

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (!isGastosVariosExpenseRow(row)) continue;
    for (let q = 1; q <= GASTOS_VARIOS_QUINCENAS; q++) {
      const n = cellToAbsNumber(row[GASTOS_VARIOS_DATA_START + q - 1]);
      if (n !== 0) out[q] = (out[q] ?? 0) + n;
    }
  }
  return out;
}
