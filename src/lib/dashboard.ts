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
  todayLabel: string;
  inCycle: boolean;
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

const MONTHS_ES = [
  "ene",
  "feb",
  "mar",
  "abr",
  "may",
  "jun",
  "jul",
  "ago",
  "sep",
  "oct",
  "nov",
  "dic",
] as const;

/**
 * Ciclo de la planilla: primer mes = ago-2026 … jul-2027 (12 meses × 2 quincenas).
 * La quincena "actual" se elige con la fecha de hoy del dispositivo:
 *   días 1–15 → 1ª quincena · días 16–fin → 2ª quincena
 * Override: NEXT_PUBLIC_EXCEL_CYCLE_START=YYYY-MM
 */
function cycleStart(): { year: number; month: number } {
  const raw =
    (typeof process !== "undefined" &&
      (process.env.NEXT_PUBLIC_EXCEL_CYCLE_START ||
        process.env.EXCEL_CYCLE_START)?.trim()) ||
    "";
  if (raw && /^\d{4}-\d{2}$/.test(raw)) {
    const [y, m] = raw.split("-").map(Number);
    return { year: y, month: m };
  }
  return { year: 2026, month: 8 };
}

function monthsBetween(fromY: number, fromM: number, toY: number, toM: number): number {
  return (toY - fromY) * 12 + (toM - fromM);
}

function formatToday(d: Date): string {
  return d.toLocaleDateString("es-AR", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
}

export type PeriodResolution = {
  /** 1-based Excel quincena column, or null if today is outside the cycle */
  col: number | null;
  periodLabel: string;
  todayLabel: string;
  inCycle: boolean;
  half: 1 | 2;
};

/** Resolve which quincena column matches "today". Never assumes cycle-start === current month. */
export function resolvePeriodFromDate(
  rows: Cell[][],
  now: Date = new Date(),
): PeriodResolution {
  const width = Math.max(1, ...rows.map((r) => r.length));
  const maxCol = Math.max(1, width - 1);
  const start = cycleStart();
  const y = now.getFullYear();
  const m = now.getMonth() + 1;
  const day = now.getDate();
  const half: 1 | 2 = day <= 15 ? 1 : 2;
  const todayLabel = formatToday(now);

  const monthOffset = monthsBetween(start.year, start.month, y, m);

  if (monthOffset < 0 || monthOffset > 11) {
    const end = periodMeta(24); // last month of a full 12-month cycle
    return {
      col: null,
      inCycle: false,
      half,
      todayLabel,
      periodLabel: `Hoy ${todayLabel} está fuera del ciclo ${MONTHS_ES[start.month - 1]} ${start.year} → ${end.monthLabel}`,
    };
  }

  const col = Math.min(monthOffset * 2 + (half - 1) + 1, maxCol);
  const { halfLabel, monthLabel } = periodMeta(col);
  return {
    col,
    inCycle: true,
    half,
    todayLabel,
    periodLabel: `${halfLabel} quincena · ${monthLabel}`,
  };
}

/** @deprecated use resolvePeriodFromDate */
export function detectCurrentQuincenaCol(
  rows: Cell[][],
  now: Date = new Date(),
): number {
  return resolvePeriodFromDate(rows, now).col ?? 1;
}

function periodMeta(col: number): {
  halfLabel: string;
  monthLabel: string;
  year: number;
  monthIndex: number;
} {
  const start = cycleStart();
  const monthIndex = Math.floor((col - 1) / 2);
  const halfLabel = (col - 1) % 2 === 0 ? "1ª" : "2ª";

  const absMonth = start.month - 1 + monthIndex;
  const year = start.year + Math.floor(absMonth / 12);
  const calendarMonth = ((absMonth % 12) + 12) % 12;
  const monthLabel = `${MONTHS_ES[calendarMonth]} ${year}`;

  return { halfLabel, monthLabel, year, monthIndex };
}

function quincenaLabel(col: number): string {
  const { halfLabel, monthLabel } = periodMeta(col);
  return `${halfLabel} quincena · ${monthLabel}`;
}

const SKIP = [/^$/i, /^-+$/];

/**
 * Interpreta la hoja PROYECCIÓN 26-27 / 20262027 según la estructura acordada.
 * rows[0] = Excel fila 1 (ya NO se descarta como header).
 */
export function buildProyeccionView(
  rows: Cell[][],
  now: Date = new Date(),
): ProyeccionView {
  const period = resolvePeriodFromDate(rows, now);

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
      periodLabel: period.periodLabel,
      todayLabel: period.todayLabel,
      inCycle: period.inCycle,
      expenseRows: [],
      incomeRows: [],
    };
  }

  if (!period.inCycle || period.col === null) {
    return {
      kpis: [
        {
          id: "out",
          label: "Periodo",
          value: "Fuera de ciclo",
          hint: period.periodLabel,
        },
      ],
      charts: [],
      periodLabel: period.periodLabel,
      todayLabel: period.todayLabel,
      inCycle: false,
      expenseRows: [],
      incomeRows: [],
    };
  }

  const col = period.col;
  const periodLabel = period.periodLabel;

  const ingresos = sumRows(rows, 1, 3, col, SKIP);
  // Gastos principales: después de ingresos hasta antes del neteo (fila 35)
  const gastos = sumRows(rows, 4, 34, col, [
    ...SKIP,
    /^resta$/i,
    /^lo que queda$/i,
    /^transf/i,
  ]);
  const tarjetas = sumRows(rows, 21, 28, col, SKIP);
  const neteo = valueAt(rows, 35, col);
  const gastosPost = sumRows(rows, 37, 41, col, SKIP);
  const balanceFinal = valueAt(rows, 44, col);

  const incomeRows = listRows(rows, 1, 3, col, SKIP);
  const cardRows = listRows(rows, 21, 28, col, SKIP).map((r) => ({
    ...r,
    value: Math.abs(r.value),
  }));
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
      id: "tarjetas",
      label: "Gasto en tarjeta",
      value: formatMoney(Math.abs(tarjetas)),
      hint: `Filas 21–28 · ${periodLabel}`,
    },
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

  if (cardRows.length > 0) {
    charts.push({
      name: `Tarjetas · ${periodLabel}`,
      data: cardRows
        .sort((a, b) => b.value - a.value)
        .map((r) => ({
          label: r.label.slice(0, 28),
          value: r.value,
        })),
    });
  }

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
    const { halfLabel, monthLabel } = periodMeta(c);
    balanceSeries.push({
      label: `${halfLabel} ${monthLabel}`,
      value: v,
    });
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

  return {
    kpis,
    charts,
    periodLabel,
    todayLabel: period.todayLabel,
    inCycle: true,
    expenseRows,
    incomeRows,
  };
}

/** @deprecated legacy auto-detect — kept unused */
export function buildKpis(): Kpi[] {
  return [];
}
export function buildCharts(): ChartSeries[] {
  return [];
}
