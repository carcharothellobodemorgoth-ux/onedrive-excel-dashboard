const GRAPH = "https://graph.microsoft.com/v1.0";

export type DriveItem = {
  id: string;
  name: string;
  webUrl?: string;
  lastModifiedDateTime?: string;
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
  | { ok: true; itemId: string; item: DriveItem; strategy: string }
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

function toShareToken(url: string): string {
  const base64 = Buffer.from(url, "utf8")
    .toString("base64")
    .replace(/=+$/g, "")
    .replace(/\+/g, "-")
    .replace(/\//g, "_");
  return `u!${base64}`;
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

function isExcelName(name?: string): boolean {
  return /\.(xlsx|xlsm|xls)$/i.test(name ?? "");
}

async function collectExcelCandidates(
  accessToken: string,
): Promise<DriveItem[]> {
  const pools: DriveItem[] = [];
  const paths = [
    `/me/drive/recent?$top=50&$select=id,name,webUrl,lastModifiedDateTime,file`,
    `/me/drive/root/search(q='.xlsx')?$top=50&$select=id,name,webUrl,lastModifiedDateTime,file`,
    `/me/drive/root/search(q='xlsx')?$top=50&$select=id,name,webUrl,lastModifiedDateTime,file`,
    `/me/drive/root/children?$top=200&$select=id,name,webUrl,lastModifiedDateTime,file,folder`,
  ];

  const hint = process.env.EXCEL_FILE_NAME?.trim();
  if (hint) {
    paths.unshift(
      `/me/drive/root/search(q='${hint.replace(/'/g, "''")}')?$top=25&$select=id,name,webUrl,lastModifiedDateTime,file`,
    );
  }

  for (const path of paths) {
    const data = await graphFetchOk<{ value: DriveItem[] }>(accessToken, path);
    if (data?.value) pools.push(...data.value);
  }

  const byId = new Map<string, DriveItem>();
  for (const it of pools) {
    if (it?.id && isExcelName(it.name)) byId.set(it.id, it);
  }
  return [...byId.values()].sort((a, b) =>
    (b.lastModifiedDateTime ?? "").localeCompare(a.lastModifiedDateTime ?? ""),
  );
}

export async function resolveExcelItemId(
  accessToken: string,
  preferredItemId?: string | null,
): Promise<ResolveResult> {
  const configured = process.env.EXCEL_DRIVE_ITEM_ID?.trim();
  const cid =
    process.env.EXCEL_ONEDRIVE_CID?.trim() || "d883d00740abbada";
  const resid =
    process.env.EXCEL_ONEDRIVE_RESID?.trim() ||
    "3cec66d5-ed44-4799-8089-a801fc93a9f2";
  const fileNameHint = process.env.EXCEL_FILE_NAME?.trim().toLowerCase();
  const shareUrl =
    process.env.EXCEL_SHARE_URL?.trim() ||
    `https://onedrive.live.com/personal/${cid}/_layouts/15/doc.aspx?resid=${resid}&cid=${cid}`;

  const tryItem = async (
    itemId: string,
    strategy: string,
  ): Promise<ResolveResult | null> => {
    const item = await graphFetchOk<DriveItem>(
      accessToken,
      `/me/drive/items/${encodeURIComponent(itemId)}?$select=id,name,webUrl,lastModifiedDateTime`,
    );
    if (item?.id) return { ok: true, itemId: item.id, item, strategy };

    if (cid) {
      const viaDrive = await graphFetchOk<DriveItem>(
        accessToken,
        `/me/drives/${encodeURIComponent(cid)}/items/${encodeURIComponent(itemId)}?$select=id,name,webUrl,lastModifiedDateTime`,
      );
      if (viaDrive?.id) {
        return {
          ok: true,
          itemId: viaDrive.id,
          item: viaDrive,
          strategy: `${strategy}+drive`,
        };
      }
    }
    return null;
  };

  if (preferredItemId) {
    const hit = await tryItem(preferredItemId, "user-selected");
    if (hit) return hit;
  }

  if (configured) {
    const hit = await tryItem(configured, "EXCEL_DRIVE_ITEM_ID");
    if (hit) return hit;
  }

  // Share / open link → driveItem
  for (const url of [
    shareUrl,
    `https://onedrive.live.com/edit.aspx?resid=${resid}&cid=${cid}`,
    `https://onedrive.live.com/?cid=${cid}&id=${resid}`,
  ]) {
    const share = await graphFetchOk<{ id?: string; name?: string; webUrl?: string }>(
      accessToken,
      `/shares/${toShareToken(url)}/driveItem?$select=id,name,webUrl,lastModifiedDateTime`,
    );
    if (share?.id) {
      return {
        ok: true,
        itemId: share.id,
        item: share as DriveItem,
        strategy: `share:${url.slice(0, 48)}`,
      };
    }
  }

  for (const candidate of residCandidates(resid, cid)) {
    const hit = await tryItem(candidate, `resid:${candidate}`);
    if (hit) return hit;
  }

  const candidates = await collectExcelCandidates(accessToken);
  const residNeedle = resid.replace(/-/g, "").toLowerCase();

  for (const it of candidates) {
    const url = (it.webUrl ?? "").toLowerCase();
    const name = (it.name ?? "").toLowerCase();
    if (
      url.includes(resid.toLowerCase()) ||
      url.includes(residNeedle) ||
      (fileNameHint && name.includes(fileNameHint))
    ) {
      return {
        ok: true,
        itemId: it.id,
        item: it,
        strategy: "candidate-match",
      };
    }
  }

  if (fileNameHint) {
    const byName = candidates.find((c) =>
      (c.name ?? "").toLowerCase().includes(fileNameHint),
    );
    if (byName) {
      return {
        ok: true,
        itemId: byName.id,
        item: byName,
        strategy: "EXCEL_FILE_NAME",
      };
    }
  }

  if (candidates.length === 1) {
    return {
      ok: true,
      itemId: candidates[0].id,
      item: candidates[0],
      strategy: "single-excel",
    };
  }

  return {
    ok: false,
    candidates,
    message:
      "No pude mapear el link de OneDrive a un archivo Graph. Elegí la planilla de la lista.",
  };
}

export async function getDriveItem(
  accessToken: string,
  itemId: string,
): Promise<DriveItem> {
  return graphFetch<DriveItem>(
    accessToken,
    `/me/drive/items/${encodeURIComponent(itemId)}?$select=id,name,webUrl,lastModifiedDateTime`,
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
  preferredItemId?: string | null,
): Promise<
  | {
      ok: true;
      item: DriveItem;
      worksheets: Worksheet[];
      itemId: string;
      strategy: string;
    }
  | { ok: false; candidates: DriveItem[]; message: string }
> {
  const resolved = await resolveExcelItemId(accessToken, preferredItemId);
  if (!resolved.ok) return resolved;

  const worksheets = await listWorksheets(accessToken, resolved.itemId);
  return {
    ok: true,
    item: resolved.item,
    worksheets,
    itemId: resolved.itemId,
    strategy: resolved.strategy,
  };
}
