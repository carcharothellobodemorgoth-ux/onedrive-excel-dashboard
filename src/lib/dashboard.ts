export type Kpi = {
  id: string;
  label: string;
  value: string;
  hint: string;
};

export type ChartSeries = {
  name: string;
  data: { label: string; value: number }[];
};

export type ProyeccionView = {
  kpis: Kpi[];
  charts: ChartSeries[];
  periodLabel: string;
  expenseRows: { label: string; value: number }[];
  incomeRows: { label: string; value: number }[];
};

type Cell = string | number | boolean | null;

function isNumber(v: Cell): v is number {
  return typeof v === "number" && Number.isFinite(v);
}

function toNumber(v: Cell): number | null {
  if (isNumber(v)) return v;
  if (typeof v !== "string") return null;
  const cleaned = v
    .replace(/\$/g, "")
    .replace(/\s/g, "")
    .replace(/\./g, "")
    .replace(",", ".");
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function formatMoney(n: number): string {
  return new Intl.NumberFormat("es-AR", {
    style: "currency",
    currency: "ARS",
    maximumFractionDigits: 2,
  }).format(n);
}

function labelAt(rows: Cell[][], excelRow: number): string {
  const v = rows[excelRow - 1]?.[0];
  if (v === null || v === undefined || v === "") return `Fila ${excelRow}`;
  return String(v).trim();
}

function valueAt(rows: Cell[][], excelRow: number, col: number): number {
  return toNumber(rows[excelRow - 1]?.[col] ?? null) ?? 0;
}

function sumRows(
  rows: Cell[][],
  excelFrom: number,
  excelTo: number,
  col: number,
  skipLabels: RegExp[] = [],
): number {
  let sum = 0;
  for (let r = excelFrom; r <= excelTo; r++) {
    const label = labelAt(rows, r);
    if (!label || skipLabels.some((re) => re.test(label))) continue;
    sum += valueAt(rows, r, col);
  }
  return sum;
}

function listRows(
  rows: Cell[][],
  excelFrom: number,
  excelTo: number,
  col: number,
  skipLabels: RegExp[] = [],
): { label: string; value: number }[] {
  const out: { label: string; value: number }[] = [];
  for (let r = excelFrom; r <= excelTo; r++) {
    const label = labelAt(rows, r);
    if (!label || skipLabels.some((re) => re.test(label))) continue;
    const value = valueAt(rows, r, col);
    if (value === 0) continue;
    out.push({ label, value });
  }
  return out;
}

/** Rightmost quincena column that has meaningful numbers (prefer balance row 44). */
export function detectCurrentQuincenaCol(rows: Cell[][]): number {
  if (rows.length === 0) return 1;
  const width = Math.max(...rows.map((r) => r.length), 1);

  for (let col = width - 1; col >= 1; col--) {
    const balance = toNumber(rows[43]?.[col] ?? null); // Excel row 44
    if (balance !== null) return col;
  }

  for (let col = width - 1; col >= 1; col--) {
    let hits = 0;
    for (const excelRow of [1, 2, 3, 35, 44]) {
      if (toNumber(rows[excelRow - 1]?.[col] ?? null) !== null) hits++;
    }
    if (hits >= 1) return col;
  }

  return 1;
}

function quincenaLabel(col: number): string {
  const monthIndex = Math.floor((col - 1) / 2); // 0-based month block
  const half = (col - 1) % 2 === 0 ? "1ª" : "2ª";
  return `${half} quincena · mes ${monthIndex + 1}`;
}

const SKIP = [/^$/i, /^-+$/];

/**
 * Interpreta la hoja PROYECCIÓN 26-27 / 20262027 según la estructura acordada.
 * rows[0] = Excel fila 1 (ya NO se descarta como header).
 */
export function buildProyeccionView(rows: Cell[][]): ProyeccionView {
  if (rows.length < 44) {
    return {
      kpis: [
        {
          id: "err",
          label: "Estructura",
          value: "—",
          hint: `Se esperaban ≥44 filas; hay ${rows.length}`,
        },
      ],
      charts: [],
      periodLabel: "—",
      expenseRows: [],
      incomeRows: [],
    };
  }

  const col = detectCurrentQuincenaCol(rows);
  const periodLabel = quincenaLabel(col);

  const ingresos = sumRows(rows, 1, 3, col, SKIP);
  // Gastos principales: después de ingresos hasta antes del neteo (fila 35)
  const gastos = sumRows(rows, 4, 34, col, [
    ...SKIP,
    /^resta$/i,
    /^lo que queda$/i,
    /^transf/i,
  ]);
  const neteo = valueAt(rows, 35, col);
  const gastosPost = sumRows(rows, 37, 41, col, SKIP);
  const balanceFinal = valueAt(rows, 44, col);

  const incomeRows = listRows(rows, 1, 3, col, SKIP);
  const expenseRows = [
    ...listRows(rows, 4, 34, col, [
      ...SKIP,
      /^resta$/i,
      /^lo que queda$/i,
      /^transf/i,
    ]),
    ...listRows(rows, 37, 41, col, SKIP),
  ]
    .map((r) => ({ ...r, value: Math.abs(r.value) }))
    .sort((a, b) => b.value - a.value);

  const kpis: Kpi[] = [
    {
      id: "ingresos",
      label: "Ingresos",
      value: formatMoney(ingresos),
      hint: `Filas 1–3 · ${periodLabel}`,
    },
    {
      id: "gastos",
      label: "Gastos",
      value: formatMoney(gastos),
      hint: "Bloque de gastos (antes del neteo)",
    },
    {
      id: "neteo",
      label: labelAt(rows, 35) || "Neteo",
      value: formatMoney(neteo),
      hint: "Fila 35",
    },
    {
      id: "gastos-post",
      label: "Gastos post-balance",
      value: formatMoney(gastosPost),
      hint: "Filas 37–41",
    },
    {
      id: "final",
      label: labelAt(rows, 44) || "Balance final",
      value: formatMoney(balanceFinal),
      hint: "Fila 44 · lo que queda",
    },
  ];

  const charts: ChartSeries[] = [];

  if (expenseRows.length > 0) {
    charts.push({
      name: `Gastos · ${periodLabel}`,
      data: expenseRows.slice(0, 12).map((r) => ({
        label: r.label.slice(0, 28),
        value: r.value,
      })),
    });
  }

  // Evolución balance final (fila 44) por quincena
  const width = Math.max(...rows.map((r) => r.length), 1);
  const balanceSeries: { label: string; value: number }[] = [];
  for (let c = 1; c < width; c++) {
    const v = toNumber(rows[43]?.[c] ?? null);
    if (v === null) continue;
    balanceSeries.push({ label: `Q${c}`, value: v });
  }
  if (balanceSeries.length > 1) {
    charts.push({
      name: `${labelAt(rows, 44) || "Balance final"} · historial quincenas`,
      data: balanceSeries.slice(-12),
    });
  }

  charts.push({
    name: `Ingresos vs gastos · ${periodLabel}`,
    data: [
      { label: "Ingresos", value: ingresos },
      { label: "Gastos", value: Math.abs(gastos) },
      { label: "Post", value: Math.abs(gastosPost) },
    ],
  });

  return { kpis, charts, periodLabel, expenseRows, incomeRows };
}

/** @deprecated legacy auto-detect — kept unused */
export function buildKpis(): Kpi[] {
  return [];
}
export function buildCharts(): ChartSeries[] {
  return [];
}
