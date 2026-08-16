export type Kpi = {
  id: string;
  label: string;
  value: string;
  hint: string;
};

export type ChartSeries = {
  name: string;
  kind?: "bar" | "pie" | "stacked";
  /** Full-width in the charts grid */
  wide?: boolean;
  data: {
    label: string;
    value?: number;
    gastado?: number;
    queda?: number;
    /** Quincena index (1-based) for clickable history bars */
    col?: number;
  }[];
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

type SheetLayout = {
  /** Array index of first quincena column */
  dataStart: number;
  hasCategory: boolean;
};

const CATEGORY_COL = 1;

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

function rawCategoryAt(rows: Cell[][], excelRow: number): string {
  const v = rows[excelRow - 1]?.[CATEGORY_COL];
  if (v === null || v === undefined || v === "") return "";
  return String(v).trim();
}

function discoverCategories(
  rows: Cell[][],
  excelFrom: number,
  excelTo: number,
  skipLabels: RegExp[] = [],
): string[] {
  const cats = new Set<string>();
  for (let r = excelFrom; r <= excelTo; r++) {
    const label = labelAt(rows, r);
    if (!label || skipLabels.some((re) => re.test(label))) continue;
    // Skip placeholder labels like "Fila 12" from empty A
    if (/^Fila \d+$/i.test(label)) continue;
    const cat = rawCategoryAt(rows, r);
    if (cat) cats.add(cat);
  }
  return [...cats];
}

function listCategoryTotals(
  rows: Cell[][],
  layout: SheetLayout,
  excelFrom: number,
  excelTo: number,
  quincenas: number[],
  skipLabels: RegExp[] = [],
): { label: string; value: number }[] {
  const map = new Map<string, number>();
  for (const cat of discoverCategories(rows, excelFrom, excelTo, skipLabels)) {
    map.set(cat, 0);
  }

  for (const q of quincenas) {
    for (let r = excelFrom; r <= excelTo; r++) {
      const label = labelAt(rows, r);
      if (!label || skipLabels.some((re) => re.test(label))) continue;
      if (/^Fila \d+$/i.test(label)) continue;
      const value = Math.abs(valueAtQuincena(rows, layout, r, q));
      if (value === 0) continue;
      const cat = layout.hasCategory
        ? rawCategoryAt(rows, r) || "Sin categoría"
        : "Sin categoría";
      if (!map.has(cat)) map.set(cat, 0);
      map.set(cat, (map.get(cat) ?? 0) + value);
    }
  }

  return [...map.entries()]
    .map(([label, value]) => ({ label, value }))
    .sort((a, b) => b.value - a.value || a.label.localeCompare(b.label, "es"));
}

/**
 * Detect A=label, B=categoría, C+=quincenas vs legacy A=label, B+=quincenas.
 * If column B looks like text categories on expense rows → new layout.
 */
export function detectSheetLayout(rows: Cell[][]): SheetLayout {
  let textish = 0;
  let numeric = 0;
  for (let r = 3; r < Math.min(34, rows.length); r++) {
    const cell = rows[r]?.[CATEGORY_COL];
    if (cell === null || cell === undefined || cell === "") continue;
    if (toNumber(cell) !== null) numeric += 1;
    else if (typeof cell === "string" && cell.trim()) textish += 1;
  }
  if (textish >= 3 && textish >= numeric) {
    return { dataStart: 2, hasCategory: true };
  }
  return { dataStart: 1, hasCategory: false };
}

function sheetCol(layout: SheetLayout, quincena: number): number {
  return layout.dataStart + quincena - 1;
}

function valueAtQuincena(
  rows: Cell[][],
  layout: SheetLayout,
  excelRow: number,
  quincena: number,
): number {
  return (
    toNumber(rows[excelRow - 1]?.[sheetCol(layout, quincena)] ?? null) ?? 0
  );
}

function sumRowsCols(
  rows: Cell[][],
  layout: SheetLayout,
  excelFrom: number,
  excelTo: number,
  quincenas: number[],
  skipLabels: RegExp[] = [],
): number {
  let sum = 0;
  for (const q of quincenas) {
    for (let r = excelFrom; r <= excelTo; r++) {
      const label = labelAt(rows, r);
      if (!label || skipLabels.some((re) => re.test(label))) continue;
      sum += valueAtQuincena(rows, layout, r, q);
    }
  }
  return sum;
}

function listRowsCols(
  rows: Cell[][],
  layout: SheetLayout,
  excelFrom: number,
  excelTo: number,
  quincenas: number[],
  skipLabels: RegExp[] = [],
): { label: string; value: number }[] {
  const map = new Map<string, number>();
  for (const q of quincenas) {
    for (let r = excelFrom; r <= excelTo; r++) {
      const label = labelAt(rows, r);
      if (!label || skipLabels.some((re) => re.test(label))) continue;
      const value = valueAtQuincena(rows, layout, r, q);
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
  const maxCol = maxDataCol(rows);
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
  const layout = detectSheetLayout(rows);
  const width = Math.max(1, ...rows.map((r) => r.length));
  return Math.max(1, Math.min(24, width - layout.dataStart));
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
const SKIP_EXPENSE = [
  ...SKIP,
  /^resta$/i,
  /^lo que queda$/i,
  /^queda$/i,
  /^transf/i,
  /^categoria$/i,
];

type BalanceLayout = {
  incomeFrom: number;
  incomeTo: number;
  expenseFrom: number;
  expenseTo: number;
  neteoRow: number | null;
  postFrom: number | null;
  postTo: number | null;
  balanceRow: number | null;
  cardRows: number[];
};

function findExcelRowsByLabel(rows: Cell[][], re: RegExp): number[] {
  const out: number[] = [];
  for (let i = 0; i < rows.length; i++) {
    const label = String(rows[i]?.[0] ?? "").trim();
    if (label && re.test(label)) out.push(i + 1);
  }
  return out;
}

function findExcelRowByLabel(rows: Cell[][], re: RegExp): number | null {
  return findExcelRowsByLabel(rows, re)[0] ?? null;
}

/**
 * Detect income / expenses / resta / lo que queda / cards by column A labels
 * so the sheet can be compacted (e.g. all balance data in rows 1–37).
 */
export function detectBalanceLayout(rows: Cell[][]): BalanceLayout {
  const balanceRow =
    findExcelRowByLabel(rows, /^lo que queda$/i) ??
    findExcelRowByLabel(rows, /^queda$/i);
  const neteoRow = findExcelRowByLabel(rows, /^resta$/i);
  const cardRows = findExcelRowsByLabel(
    rows,
    /amex|platinum|palacio|tarjeta/i,
  );

  const incomeFrom = 1;
  let incomeTo = 3;
  // If neteo is early, keep income as first rows before expenses
  if (neteoRow !== null && neteoRow <= 4) {
    incomeTo = Math.max(1, neteoRow - 1);
  }

  const expenseFrom = incomeTo + 1;
  const expenseTo =
    neteoRow !== null
      ? Math.max(expenseFrom, neteoRow - 1)
      : balanceRow !== null
        ? Math.max(expenseFrom, balanceRow - 1)
        : Math.max(expenseFrom, rows.length);

  let postFrom: number | null = null;
  let postTo: number | null = null;
  if (neteoRow !== null && balanceRow !== null && balanceRow > neteoRow + 1) {
    postFrom = neteoRow + 1;
    postTo = balanceRow - 1;
  }

  return {
    incomeFrom,
    incomeTo,
    expenseFrom,
    expenseTo,
    neteoRow,
    postFrom,
    postTo,
    balanceRow,
    cardRows,
  };
}

export const DEFAULT_PERIOD_SELECTION: PeriodSelection = {
  mode: "quincena",
  col: "actual",
};

export function listExpenseCategories(rows: Cell[][]): string[] {
  const bal = detectBalanceLayout(rows);
  const end = bal.balanceRow
    ? bal.balanceRow - 1
    : Math.max(bal.expenseTo, bal.postTo ?? bal.expenseTo);
  return discoverCategories(rows, bal.expenseFrom, end, SKIP_EXPENSE).sort(
    (a, b) => a.localeCompare(b, "es"),
  );
}

export function buildProyeccionView(
  rows: Cell[][],
  selection: PeriodSelection = DEFAULT_PERIOD_SELECTION,
  now: Date = new Date(),
  extras?: { gastosVariosByQuincena?: Record<number, number> },
): ProyeccionView {
  const layout = detectSheetLayout(rows);
  const bal = detectBalanceLayout(rows);
  const resolved = resolveSelectionCols(rows, selection, now);
  const gv = extras?.gastosVariosByQuincena ?? {};
  const gvSum = (cols: number[]) =>
    cols.reduce((acc, q) => acc + (gv[q] ?? 0), 0);

  if (rows.length < 3 || bal.balanceRow === null) {
    return {
      kpis: [
        {
          id: "err",
          label: "Estructura",
          value: "—",
          hint:
            bal.balanceRow === null
              ? `No encontré la fila «Lo que queda» (hay ${rows.length} filas)`
              : `Se esperaban datos de balance; hay ${rows.length} filas`,
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

  const ingresos = sumRowsCols(
    rows,
    layout,
    bal.incomeFrom,
    bal.incomeTo,
    cols,
    SKIP,
  );
  const gastos = sumRowsCols(
    rows,
    layout,
    bal.expenseFrom,
    bal.expenseTo,
    cols,
    SKIP_EXPENSE,
  );
  const tarjetas =
    bal.cardRows.length > 0
      ? bal.cardRows.reduce(
          (acc, r) =>
            acc +
            cols.reduce(
              (s, q) => s + valueAtQuincena(rows, layout, r, q),
              0,
            ),
          0,
        )
      : 0;
  const gastosPost =
    bal.postFrom !== null && bal.postTo !== null && bal.postTo >= bal.postFrom
      ? sumRowsCols(rows, layout, bal.postFrom, bal.postTo, cols, SKIP)
      : 0;
  const neteo =
    bal.neteoRow !== null
      ? valueAtQuincena(rows, layout, bal.neteoRow, lastCol)
      : 0;
  const balanceSheet = valueAtQuincena(
    rows,
    layout,
    bal.balanceRow,
    lastCol,
  );
  const gastosVariosPeriodo = gvSum(cols);
  const balanceFinal = balanceSheet - gastosVariosPeriodo;

  const incomeRows = listRowsCols(
    rows,
    layout,
    bal.incomeFrom,
    bal.incomeTo,
    cols,
    SKIP,
  );
  const expenseRows = [
    ...listRowsCols(
      rows,
      layout,
      bal.expenseFrom,
      bal.expenseTo,
      cols,
      SKIP_EXPENSE,
    ),
    ...(bal.postFrom !== null && bal.postTo !== null
      ? listRowsCols(rows, layout, bal.postFrom, bal.postTo, cols, SKIP)
      : []),
  ]
    .map((r) => ({ ...r, value: Math.abs(r.value) }))
    .sort((a, b) => b.value - a.value);

  const categoryChart = (() => {
    const catMap = new Map<string, number>();
    for (const part of [
      ...listCategoryTotals(
        rows,
        layout,
        bal.expenseFrom,
        bal.expenseTo,
        cols,
        SKIP_EXPENSE,
      ),
      ...(bal.postFrom !== null && bal.postTo !== null
        ? listCategoryTotals(
            rows,
            layout,
            bal.postFrom,
            bal.postTo,
            cols,
            SKIP,
          )
        : []),
    ]) {
      catMap.set(part.label, (catMap.get(part.label) ?? 0) + part.value);
    }
    return [...catMap.entries()]
      .map(([label, value]) => ({ label, value }))
      .sort((a, b) => b.value - a.value || a.label.localeCompare(b.label, "es"));
  })();

  const cardHint =
    bal.cardRows.length > 0
      ? `Filas ${bal.cardRows.join(", ")} · ${periodLabel}`
      : periodLabel;

  const kpis: Kpi[] = [
    {
      id: "final",
      label: labelAt(rows, bal.balanceRow) || "Lo que queda",
      value: formatMoney(balanceFinal),
      hint: `Fila ${bal.balanceRow} − Gastos Varios · fin (${quincenaLabel(lastCol)})`,
    },
    {
      id: "gastos-varios",
      label: "Gastos varios",
      value: formatMoney(gastosVariosPeriodo),
      hint: `Hoja Gastos Varios · ${periodLabel}`,
    },
    {
      id: "tarjetas",
      label: "Gasto en tarjeta",
      value: formatMoney(Math.abs(tarjetas)),
      hint: cardHint,
    },
    {
      id: "ingresos",
      label: "Ingresos",
      value: formatMoney(ingresos),
      hint: `Filas ${bal.incomeFrom}–${bal.incomeTo} · ${periodLabel}`,
    },
    {
      id: "gastos",
      label: "Gastos",
      value: formatMoney(gastos),
      hint: `Filas ${bal.expenseFrom}–${bal.expenseTo} (antes del neteo)`,
    },
  ];

  const charts: ChartSeries[] = [];

  const maxQ = maxDataCol(rows);
  const historyStacked: {
    label: string;
    gastado: number;
    queda: number;
    col: number;
  }[] = [];
  const balanceIdx = bal.balanceRow - 1;
  for (let q = 1; q <= maxQ; q++) {
    const quedaRaw = toNumber(
      rows[balanceIdx]?.[sheetCol(layout, q)] ?? null,
    );
    const gastoPeriodo = sumRowsCols(
      rows,
      layout,
      bal.expenseFrom,
      bal.expenseTo,
      [q],
      SKIP_EXPENSE,
    );
    const gastoPost =
      bal.postFrom !== null && bal.postTo !== null
        ? sumRowsCols(rows, layout, bal.postFrom, bal.postTo, [q], SKIP)
        : 0;
    const varios = gv[q] ?? 0;
    const gastado = Math.abs(gastoPeriodo) + Math.abs(gastoPost) + varios;
    if (quedaRaw === null && gastado === 0) continue;
    const { halfLabel, monthLabel } = periodMeta(q);
    historyStacked.push({
      label: `${halfLabel} ${monthLabel}`,
      gastado,
      queda: Math.max(0, (quedaRaw ?? 0) - varios),
      col: q,
    });
  }
  if (historyStacked.length > 0) {
    charts.push({
      name: "Gastado vs lo que queda · todas las quincenas",
      kind: "stacked",
      wide: true,
      data: historyStacked,
    });
  }

  if (layout.hasCategory && categoryChart.length > 0) {
    charts.push({
      name: `Categorías · ${periodLabel}`,
      data: categoryChart,
    });
  }

  if (expenseRows.length > 0) {
    charts.push({
      name: `Gastos · ${periodLabel}`,
      kind: "pie",
      data: expenseRows.slice(0, 12).map((r) => ({
        label: r.label.slice(0, 28),
        value: r.value,
      })),
    });
  }

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
