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

function isNumber(v: string | number | boolean | null): v is number {
  return typeof v === "number" && Number.isFinite(v);
}

function formatNumber(n: number): string {
  return new Intl.NumberFormat("es", {
    maximumFractionDigits: 2,
  }).format(n);
}

/** Infer KPIs from numeric columns and a categorical first column when present. */
export function buildKpis(
  headers: string[],
  rows: (string | number | boolean | null)[][],
): Kpi[] {
  const kpis: Kpi[] = [
    {
      id: "rows",
      label: "Filas",
      value: formatNumber(rows.length),
      hint: "Registros con datos",
    },
  ];

  headers.forEach((header, col) => {
    const nums = rows.map((r) => r[col]).filter(isNumber);
    if (nums.length < Math.max(2, Math.floor(rows.length * 0.3))) return;

    const sum = nums.reduce((a, b) => a + b, 0);
    const avg = sum / nums.length;
    kpis.push({
      id: `sum-${col}`,
      label: `Suma · ${header}`,
      value: formatNumber(sum),
      hint: `${nums.length} valores numéricos`,
    });
    kpis.push({
      id: `avg-${col}`,
      label: `Promedio · ${header}`,
      value: formatNumber(avg),
      hint: header,
    });
  });

  return kpis.slice(0, 6);
}

/** Build up to two simple bar charts: category (col0) vs first numeric columns. */
export function buildCharts(
  headers: string[],
  rows: (string | number | boolean | null)[][],
): ChartSeries[] {
  if (headers.length < 2 || rows.length === 0) return [];

  const numericCols = headers
    .map((_, col) => col)
    .filter((col) => {
      const nums = rows.map((r) => r[col]).filter(isNumber);
      return nums.length >= Math.max(2, Math.floor(rows.length * 0.3));
    })
    .filter((col) => col !== 0)
    .slice(0, 2);

  if (numericCols.length === 0) return [];

  const labelCol = 0;
  const charts: ChartSeries[] = [];

  for (const col of numericCols) {
    const buckets = new Map<string, number>();
    for (const row of rows) {
      const rawLabel = row[labelCol];
      const label =
        rawLabel === null || rawLabel === ""
          ? "(vacío)"
          : String(rawLabel).slice(0, 40);
      const val = row[col];
      if (!isNumber(val)) continue;
      buckets.set(label, (buckets.get(label) ?? 0) + val);
    }

    const data = [...buckets.entries()]
      .map(([label, value]) => ({ label, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 12);

    if (data.length > 0) {
      charts.push({ name: headers[col], data });
    }
  }

  return charts;
}
