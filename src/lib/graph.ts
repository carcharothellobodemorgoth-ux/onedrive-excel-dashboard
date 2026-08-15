const GRAPH = "https://graph.microsoft.com/v1.0";

export type DriveItem = {
  id: string;
  name: string;
  webUrl?: string;
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

  return res.json() as Promise<T>;
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

function residCandidates(resid: string, cid?: string): string[] {
  const dashed = resid.toLowerCase();
  const compact = dashed.replace(/-/g, "");
  const compactUpper = compact.toUpperCase();
  const out = new Set<string>([dashed, compact, compactUpper, resid]);

  if (cid) {
    const c = cid.toLowerCase();
    out.add(`${c}!${dashed}`);
    out.add(`${c}!${compact}`);
    out.add(`${c}!${compactUpper}`);
  }

  return [...out];
}

/**
 * OneDrive personal "resid" from the browser URL is NOT a Graph item id.
 * Resolve via cid/resid variants, search, and recent files.
 */
export async function resolveExcelItemId(
  accessToken: string,
): Promise<{ itemId: string; item: DriveItem; strategy: string }> {
  const configured = process.env.EXCEL_DRIVE_ITEM_ID?.trim();
  const cid =
    process.env.EXCEL_ONEDRIVE_CID?.trim() || "d883d00740abbada";
  const resid =
    process.env.EXCEL_ONEDRIVE_RESID?.trim() ||
    configured ||
    "3cec66d5-ed44-4799-8089-a801fc93a9f2";
  const fileNameHint = process.env.EXCEL_FILE_NAME?.trim().toLowerCase();

  const tryItem = async (
    itemId: string,
    strategy: string,
  ): Promise<{ itemId: string; item: DriveItem; strategy: string } | null> => {
    const item = await graphFetchOk<DriveItem>(
      accessToken,
      `/me/drive/items/${encodeURIComponent(itemId)}?$select=id,name,webUrl,file`,
    );
    if (item?.id) return { itemId: item.id, item, strategy };

    if (cid) {
      const viaDrive = await graphFetchOk<DriveItem>(
        accessToken,
        `/me/drives/${encodeURIComponent(cid)}/items/${encodeURIComponent(itemId)}?$select=id,name,webUrl,file`,
      );
      if (viaDrive?.id) {
        return { itemId: viaDrive.id, item: viaDrive, strategy: `${strategy}+drive` };
      }
    }
    return null;
  };

  if (configured) {
    const hit = await tryItem(configured, "EXCEL_DRIVE_ITEM_ID");
    if (hit) return hit;
  }

  for (const candidate of residCandidates(resid, cid)) {
    const hit = await tryItem(candidate, `resid:${candidate}`);
    if (hit) return hit;
  }

  const pools: DriveItem[] = [];

  const recent = await graphFetchOk<{ value: DriveItem[] }>(
    accessToken,
    `/me/drive/recent?$top=50&$select=id,name,webUrl,file`,
  );
  if (recent?.value) pools.push(...recent.value);

  const search = await graphFetchOk<{ value: DriveItem[] }>(
    accessToken,
    `/me/drive/root/search(q='.xlsx')?$top=50&$select=id,name,webUrl,file`,
  );
  if (search?.value) pools.push(...search.value);

  const searchXls = await graphFetchOk<{ value: DriveItem[] }>(
    accessToken,
    `/me/drive/root/search(q='xlsx')?$top=50&$select=id,name,webUrl,file`,
  );
  if (searchXls?.value) pools.push(...searchXls.value);

  const byId = new Map<string, DriveItem>();
  for (const it of pools) {
    if (it?.id) byId.set(it.id, it);
  }

  const residNeedle = resid.replace(/-/g, "").toLowerCase();
  for (const it of byId.values()) {
    const url = (it.webUrl ?? "").toLowerCase();
    const name = (it.name ?? "").toLowerCase();
    if (
      url.includes(resid.toLowerCase()) ||
      url.includes(residNeedle) ||
      (fileNameHint && name.includes(fileNameHint))
    ) {
      return { itemId: it.id, item: it, strategy: "search/recent-match" };
    }
  }

  const excelLike = [...byId.values()].filter((it) =>
    /\.(xlsx|xlsm|xls)$/i.test(it.name ?? ""),
  );

  if (excelLike.length === 1) {
    return {
      itemId: excelLike[0].id,
      item: excelLike[0],
      strategy: "single-excel-in-recent-search",
    };
  }

  const names = excelLike
    .slice(0, 12)
    .map((f) => f.name)
    .join(", ");
  throw new Error(
    `No se pudo resolver la Excel (resid inválido para Graph). Archivos .xlsx vistos: ${names || "(ninguno)"}. Definí EXCEL_DRIVE_ITEM_ID con el id real de Graph o EXCEL_FILE_NAME.`,
  );
}

export async function getDriveItem(
  accessToken: string,
  itemId: string,
): Promise<DriveItem> {
  return graphFetch<DriveItem>(
    accessToken,
    `/me/drive/items/${encodeURIComponent(itemId)}?$select=id,name,webUrl`,
  );
}

export async function listWorksheets(
  accessToken: string,
  itemId: string,
): Promise<Worksheet[]> {
  const data = await graphFetch<{ value: Worksheet[] }>(
    accessToken,
    `/me/drive/items/${encodeURIComponent(itemId)}/workbook/worksheets?$select=id,name,position&$orderby=position`,
  );
  return data.value ?? [];
}

type UsedRange = {
  values?: unknown[][];
  text?: string[][];
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
  itemId: string,
  worksheetId: string,
): Promise<SheetData> {
  const worksheet = await graphFetch<Worksheet>(
    accessToken,
    `/me/drive/items/${encodeURIComponent(itemId)}/workbook/worksheets/${encodeURIComponent(worksheetId)}?$select=id,name,position`,
  );

  const range = await graphFetch<UsedRange>(
    accessToken,
    `/me/drive/items/${encodeURIComponent(itemId)}/workbook/worksheets/${encodeURIComponent(worksheetId)}/usedRange(valuesOnly=true)?$select=values`,
  );

  const values = range.values ?? [];
  if (values.length === 0) {
    return { worksheet, headers: [], rows: [] };
  }

  const headerRow = values[0] ?? [];
  const headers = headerRow.map((h, i) => {
    const label = cellValue(h);
    return label === null || label === "" ? `Columna ${i + 1}` : String(label);
  });

  const rows = values.slice(1).map((row) =>
    headers.map((_, i) => cellValue(row[i])),
  );

  return { worksheet, headers, rows };
}

export async function loadWorkbookSummary(
  accessToken: string,
): Promise<{ item: DriveItem; worksheets: Worksheet[]; itemId: string; strategy: string }> {
  const resolved = await resolveExcelItemId(accessToken);
  const worksheets = await listWorksheets(accessToken, resolved.itemId);
  return {
    item: resolved.item,
    worksheets,
    itemId: resolved.itemId,
    strategy: resolved.strategy,
  };
}
