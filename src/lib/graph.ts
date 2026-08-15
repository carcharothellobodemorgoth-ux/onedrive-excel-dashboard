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

  // 1) Shared with me FIRST (common for personal OneDrive links you can open but don't own)
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
    // remote ids from sharedWithMe are often already valid
    return {
      ok: true,
      itemId: sharedHit.id,
      driveId: sharedHit.parentReference.driveId,
      item: sharedHit,
      strategy: "sharedWithMe-direct",
    };
  }

  // 2) Explicit path on own drive
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

  // 3) Own drive + shared candidates — pick by name
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
      "No encontré PROYECCIÓN 26-27 en tu OneDrive ni en Compartido conmigo. Si es de otra persona, pedile que te la comparta (puede editar) o elegí un archivo de la lista. También cerrá sesión y volvé a entrar para aceptar el permiso Files.Read.All.",
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
