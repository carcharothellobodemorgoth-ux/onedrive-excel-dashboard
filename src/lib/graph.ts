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

export async function getDriveItem(
  accessToken: string,
  itemId: string,
): Promise<DriveItem> {
  return graphFetch<DriveItem>(
    accessToken,
    `/me/drive/items/${itemId}?$select=id,name,webUrl`,
  );
}

export async function listWorksheets(
  accessToken: string,
  itemId: string,
): Promise<Worksheet[]> {
  const data = await graphFetch<{ value: Worksheet[] }>(
    accessToken,
    `/me/drive/items/${itemId}/workbook/worksheets?$select=id,name,position&$orderby=position`,
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
    `/me/drive/items/${itemId}/workbook/worksheets/${worksheetId}?$select=id,name,position`,
  );

  const range = await graphFetch<UsedRange>(
    accessToken,
    `/me/drive/items/${itemId}/workbook/worksheets/${worksheetId}/usedRange(valuesOnly=true)?$select=values`,
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
  itemId: string,
): Promise<{ item: DriveItem; worksheets: Worksheet[] }> {
  const [item, worksheets] = await Promise.all([
    getDriveItem(accessToken, itemId),
    listWorksheets(accessToken, itemId),
  ]);
  return { item, worksheets };
}
