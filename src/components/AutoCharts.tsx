"use client";

import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Legend,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";

type ChartPoint = {
  label: string;
  value?: number;
  gastado?: number;
  queda?: number;
  col?: number;
};

type ChartSeries = {
  name: string;
  kind?: "bar" | "pie" | "stacked";
  wide?: boolean;
  data: ChartPoint[];
};

const PIE_COLORS = [
  "#34d399",
  "#6ee7b7",
  "#a7f3d0",
  "#2dd4bf",
  "#22d3ee",
  "#38bdf8",
  "#818cf8",
  "#c084fc",
  "#f472b6",
  "#fb7185",
  "#fb923c",
  "#fbbf24",
];

const tooltipStyle = {
  background: "#18181b",
  border: "1px solid #3f3f46",
  borderRadius: 8,
};

function money(n: number): string {
  return new Intl.NumberFormat("es-AR", {
    style: "currency",
    currency: "ARS",
    maximumFractionDigits: 0,
  }).format(n);
}

function colFromClick(raw: unknown): number | null {
  if (!raw || typeof raw !== "object") return null;
  const payload =
    "payload" in raw && raw.payload && typeof raw.payload === "object"
      ? (raw.payload as ChartPoint)
      : (raw as ChartPoint);
  return typeof payload.col === "number" ? payload.col : null;
}

export function AutoCharts({
  charts,
  selectedCols = [],
  onSelectCol,
}: {
  charts: ChartSeries[];
  selectedCols?: number[];
  onSelectCol?: (col: number) => void;
}) {
  if (charts.length === 0) return null;

  const selected = new Set(selectedCols);

  return (
    <div className="grid gap-4 lg:grid-cols-2">
      {charts.map((chart) => (
        <article
          key={chart.name}
          className={`rounded-2xl border border-white/10 bg-white/5 p-5 backdrop-blur ${
            chart.wide ? "lg:col-span-2" : ""
          }`}
        >
          <h3 className="mb-1 text-sm font-semibold text-white">{chart.name}</h3>
          {chart.kind === "stacked" && onSelectCol && (
            <p className="mb-3 text-xs text-zinc-400">
              Clic en una barra para seleccionar esa quincena
            </p>
          )}
          <div className={`w-full ${chart.wide ? "h-80" : "h-72"}`}>
            <ResponsiveContainer width="100%" height="100%">
              {chart.kind === "pie" ? (
                <PieChart>
                  <Pie
                    data={chart.data}
                    dataKey="value"
                    nameKey="label"
                    cx="50%"
                    cy="45%"
                    outerRadius={95}
                    innerRadius={42}
                    paddingAngle={1}
                    label={false}
                  >
                    {chart.data.map((_, i) => (
                      <Cell
                        key={`${chart.name}-${i}`}
                        fill={PIE_COLORS[i % PIE_COLORS.length]}
                        stroke="#09090b"
                        strokeWidth={1}
                      />
                    ))}
                  </Pie>
                  <Tooltip
                    formatter={(value) =>
                      money(typeof value === "number" ? value : Number(value))
                    }
                    contentStyle={tooltipStyle}
                  />
                  <Legend
                    verticalAlign="bottom"
                    height={56}
                    wrapperStyle={{ fontSize: 11, color: "#a1a1aa" }}
                  />
                </PieChart>
              ) : chart.kind === "stacked" ? (
                <BarChart
                  data={chart.data}
                  margin={{ left: 0, right: 8, bottom: 8 }}
                  onClick={(state) => {
                    const payload = (
                      state as { activePayload?: { payload?: ChartPoint }[] }
                    )?.activePayload?.[0];
                    const col = colFromClick(payload);
                    if (col != null) onSelectCol?.(col);
                  }}
                  style={{ cursor: onSelectCol ? "pointer" : undefined }}
                >
                  <CartesianGrid strokeDasharray="3 3" stroke="#3f3f46" />
                  <XAxis
                    dataKey="label"
                    tick={{ fill: "#a1a1aa", fontSize: 10 }}
                    interval={0}
                    angle={-35}
                    textAnchor="end"
                    height={70}
                  />
                  <YAxis tick={{ fill: "#a1a1aa", fontSize: 11 }} />
                  <Tooltip
                    formatter={(value, name) => [
                      money(typeof value === "number" ? value : Number(value)),
                      name === "gastado" ? "Gastado" : "Lo que queda",
                    ]}
                    contentStyle={tooltipStyle}
                  />
                  <Legend
                    wrapperStyle={{ fontSize: 12, color: "#a1a1aa" }}
                    formatter={(value) =>
                      value === "gastado" ? "Gastado" : "Lo que queda"
                    }
                  />
                  <Bar
                    dataKey="gastado"
                    name="gastado"
                    stackId="a"
                    cursor="pointer"
                    onClick={(data) => {
                      const col = colFromClick(data);
                      if (col != null) onSelectCol?.(col);
                    }}
                  >
                    {chart.data.map((entry, i) => {
                      const isSel =
                        typeof entry.col === "number" && selected.has(entry.col);
                      return (
                        <Cell
                          key={`g-${i}`}
                          fill={isSel ? "#fb7185" : "#f87171"}
                          stroke={isSel ? "#ffffff" : "transparent"}
                          strokeWidth={isSel ? 2 : 0}
                        />
                      );
                    })}
                  </Bar>
                  <Bar
                    dataKey="queda"
                    name="queda"
                    stackId="a"
                    cursor="pointer"
                    radius={[6, 6, 0, 0]}
                    onClick={(data) => {
                      const col = colFromClick(data);
                      if (col != null) onSelectCol?.(col);
                    }}
                  >
                    {chart.data.map((entry, i) => {
                      const isSel =
                        typeof entry.col === "number" && selected.has(entry.col);
                      return (
                        <Cell
                          key={`q-${i}`}
                          fill={isSel ? "#6ee7b7" : "#34d399"}
                          stroke={isSel ? "#ffffff" : "transparent"}
                          strokeWidth={isSel ? 2 : 0}
                        />
                      );
                    })}
                  </Bar>
                </BarChart>
              ) : (
                <BarChart data={chart.data} margin={{ left: 0, right: 8 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#3f3f46" />
                  <XAxis
                    dataKey="label"
                    tick={{ fill: "#a1a1aa", fontSize: 11 }}
                    interval={0}
                    angle={-25}
                    textAnchor="end"
                    height={60}
                  />
                  <YAxis tick={{ fill: "#a1a1aa", fontSize: 11 }} />
                  <Tooltip
                    formatter={(value) =>
                      money(typeof value === "number" ? value : Number(value))
                    }
                    contentStyle={tooltipStyle}
                  />
                  <Bar dataKey="value" fill="#34d399" radius={[6, 6, 0, 0]} />
                </BarChart>
              )}
            </ResponsiveContainer>
          </div>
        </article>
      ))}
    </div>
  );
}
