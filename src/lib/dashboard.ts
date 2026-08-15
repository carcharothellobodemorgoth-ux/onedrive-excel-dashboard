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

export type PeriodMode = "quincena" | "mes" | "custom";

export type PeriodSelection =
  | { mode: "quincena"; col: number | "actual" }
  | { mode: "mes"; monthIndex: number | "actual" }
  | { mode: "custom"; fromCol: number; toCol: number };

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

function sumRowsCols(
  rows: Cell[][],
  excelFrom: number,
  excelTo: number,
  cols: number[],
  skipLabels: RegExp[] = [],
): number {
  let sum = 0;
  for (const col of cols) {
    for (let r = excelFrom; r <= excelTo; r++) {
      const label = labelAt(rows, r);
      if (!label || skipLabels.some((re) => re.test(label))) continue;
      sum += valueAt(rows, r, col);
    }
  }
  return sum;
}

function listRowsCols(
  rows: Cell[][],
  excelFrom: number,
  excelTo: number,
  cols: number[],
  skipLabels: RegExp[] = [],
): { label: string; value: number }[] {
  const map = new Map<string, number>();
  for (const col of cols) {
    for (let r = excelFrom; r <= excelTo; r++) {
      const label = labelAt(rows, r);
      if (!label || skipLabels.some((re) => re.test(label))) continue;
      const value = valueAt(rows, r, col);
      if (value === 0) continue;
      map.set(label, (map.get(label) ?? 0) + value);
    }
  }
  return [...map.entries()].map(([label, value]) => ({ label, value }));
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

export function periodMeta(col: number): {
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

export function quincenaLabel(col: number): string {
  const { halfLabel, monthLabel } = periodMeta(col);
  return `${halfLabel} quincena · ${monthLabel}`;
}

export function monthLabelFromIndex(monthIndex: number): string {
  return periodMeta(monthIndex * 2 + 1).monthLabel;
}

export type PeriodOption = { value: number; label: string };

export function listQuincenaOptions(maxCol = 24): PeriodOption[] {
  const n = Math.min(maxCol, 24);
  return Array.from({ length: n }, (_, i) => {
    const col = i + 1;
    return { value: col, label: quincenaLabel(col) };
  });
}

export function listMonthOptions(): PeriodOption[] {
  return Array.from({ length: 12 }, (_, monthIndex) => ({
    value: monthIndex,
    label: monthLabelFromIndex(monthIndex),
  }));
}

export type PeriodResolution = {
  col: number | null;
  periodLabel: string;
  todayLabel: string;
  inCycle: boolean;
  half: 1 | 2;
  monthIndex: number | null;
};

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
    const end = periodMeta(24);
    return {
      col: null,
      inCycle: false,
      half,
      monthIndex: null,
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
    monthIndex: monthOffset,
    todayLabel,
    periodLabel: `${halfLabel} quincena · ${monthLabel}`,
  };
}

export function detectCurrentQuincenaCol(
  rows: Cell[][],
  now: Date = new Date(),
): number {
  return resolvePeriodFromDate(rows, now).col ?? 1;
}

export function maxDataCol(rows: Cell[][]): number {
  const width = Math.max(1, ...rows.map((r) => r.length));
  return Math.max(1, Math.min(24, width - 1));
}

export function resolveSelectionCols(
  rows: Cell[][],
  selection: PeriodSelection,
  now: Date = new Date(),
): { cols: number[]; periodLabel: string; todayLabel: string; inCycle: boolean } {
  const today = resolvePeriodFromDate(rows, now);
  const maxCol = maxDataCol(rows);

  if (selection.mode === "quincena") {
    if (selection.col === "actual") {
      if (!today.inCycle || today.col === null) {
        return {
          cols: [],
          periodLabel: today.periodLabel,
          todayLabel: today.todayLabel,
          inCycle: false,
        };
      }
      return {
        cols: [today.col],
        periodLabel: today.periodLabel,
        todayLabel: today.todayLabel,
        inCycle: true,
      };
    }
    const col = Math.min(Math.max(selection.col, 1), maxCol);
    return {
      cols: [col],
      periodLabel: quincenaLabel(col),
      todayLabel: today.todayLabel,
      inCycle: true,
    };
  }

  if (selection.mode === "mes") {
    let monthIndex: number;
    if (selection.monthIndex === "actual") {
      if (!today.inCycle || today.monthIndex === null) {
        return {
          cols: [],
          periodLabel: today.periodLabel,
          todayLabel: today.todayLabel,
          inCycle: false,
        };
      }
      monthIndex = today.monthIndex;
    } else {
      monthIndex = Math.min(Math.max(selection.monthIndex, 0), 11);
    }
    const c1 = monthIndex * 2 + 1;
    const c2 = monthIndex * 2 + 2;
    const cols = [c1, c2].filter((c) => c <= maxCol);
    return {
      cols,
      periodLabel: `Mes · ${monthLabelFromIndex(monthIndex)}`,
      todayLabel: today.todayLabel,
      inCycle: true,
    };
  }

  let from = Math.min(selection.fromCol, selection.toCol);
  let to = Math.max(selection.fromCol, selection.toCol);
  from = Math.min(Math.max(from, 1), maxCol);
  to = Math.min(Math.max(to, 1), maxCol);
  const cols = Array.from({ length: to - from + 1 }, (_, i) => from + i);
  const label =
    from === to
      ? quincenaLabel(from)
      : `${quincenaLabel(from)} → ${quincenaLabel(to)}`;
  return {
    cols,
    periodLabel: `Custom · ${label}`,
    todayLabel: today.todayLabel,
    inCycle: true,
  };
}

const SKIP = [/^$/i, /^-+$/];

export const DEFAULT_PERIOD_SELECTION: PeriodSelection = {
  mode: "quincena",
  col: "actual",
};

export function buildProyeccionView(
  rows: Cell[][],
  selection: PeriodSelection = DEFAULT_PERIOD_SELECTION,
  now: Date = new Date(),
): ProyeccionView {
  const resolved = resolveSelectionCols(rows, selection, now);

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
      periodLabel: resolved.periodLabel,
      todayLabel: resolved.todayLabel,
      inCycle: resolved.inCycle,
      expenseRows: [],
      incomeRows: [],
    };
  }

  if (!resolved.inCycle || resolved.cols.length === 0) {
    return {
      kpis: [
        {
          id: "out",
          label: "Periodo",
          value: "Fuera de ciclo",
          hint: resolved.periodLabel,
        },
      ],
      charts: [],
      periodLabel: resolved.periodLabel,
      todayLabel: resolved.todayLabel,
      inCycle: false,
      expenseRows: [],
      incomeRows: [],
    };
  }

  const cols = resolved.cols;
  const lastCol = cols[cols.length - 1];
  const periodLabel = resolved.periodLabel;

  const ingresos = sumRowsCols(rows, 1, 3, cols, SKIP);
  const gastos = sumRowsCols(rows, 4, 34, cols, [
    ...SKIP,
    /^resta$/i,
    /^lo que queda$/i,
    /^transf/i,
  ]);
  const tarjetas = sumRowsCols(rows, 21, 28, cols, SKIP);
  const gastosPost = sumRowsCols(rows, 37, 41, cols, SKIP);
  const neteo = valueAt(rows, 35, lastCol);
  const balanceFinal = valueAt(rows, 44, lastCol);

  const incomeRows = listRowsCols(rows, 1, 3, cols, SKIP);
  const expenseRows = [
    ...listRowsCols(rows, 4, 34, cols, [
      ...SKIP,
      /^resta$/i,
      /^lo que queda$/i,
      /^transf/i,
    ]),
    ...listRowsCols(rows, 37, 41, cols, SKIP),
  ]
    .map((r) => ({ ...r, value: Math.abs(r.value) }))
    .sort((a, b) => b.value - a.value);

  const cardRows = listRowsCols(rows, 21, 28, cols, SKIP)
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
      hint: `Fila 35 · fin del periodo (${quincenaLabel(lastCol)})`,
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
      hint: `Fila 44 · fin del periodo (${quincenaLabel(lastCol)})`,
    },
  ];

  const charts: ChartSeries[] = [];

  if (cardRows.length > 0) {
    charts.push({
      name: `Tarjetas · ${periodLabel}`,
      data: cardRows.slice(0, 12).map((r) => ({
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
    todayLabel: resolved.todayLabel,
    inCycle: true,
    expenseRows,
    incomeRows,
  };
}
