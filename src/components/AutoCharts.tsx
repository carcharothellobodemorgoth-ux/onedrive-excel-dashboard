"use client";

import {
  Bar,
  BarChart,
  CartesianGrid,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";

type ChartSeries = {
  name: string;
  data: { label: string; value: number }[];
};

export function AutoCharts({ charts }: { charts: ChartSeries[] }) {
  if (charts.length === 0) return null;

  return (
    <div className="grid gap-4 lg:grid-cols-2">
      {charts.map((chart) => (
        <article
          key={chart.name}
          className="rounded-2xl border border-white/10 bg-white/5 p-5 backdrop-blur"
        >
          <h3 className="mb-4 text-sm font-semibold text-white">{chart.name}</h3>
          <div className="h-64 w-full">
            <ResponsiveContainer width="100%" height="100%">
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
                  contentStyle={{
                    background: "#18181b",
                    border: "1px solid #3f3f46",
                    borderRadius: 8,
                  }}
                />
                <Bar dataKey="value" fill="#34d399" radius={[6, 6, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </article>
      ))}
    </div>
  );
}
