"use client";

type Kpi = {
  id: string;
  label: string;
  value: string;
  hint: string;
};

export function KpiCards({ kpis }: { kpis: Kpi[] }) {
  if (kpis.length === 0) return null;

  return (
    <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-6">
      {kpis.map((kpi) => {
        const pairRow = kpi.id === "ingresos" || kpi.id === "gastos";
        return (
          <article
            key={kpi.id}
            className={`rounded-2xl border border-white/10 bg-white/5 p-5 backdrop-blur ${
              pairRow ? "xl:col-span-3" : "xl:col-span-2"
            }`}
          >
            <p className="text-xs font-medium uppercase tracking-wide text-emerald-300/80">
              {kpi.label}
            </p>
            <p className="mt-2 text-3xl font-semibold tracking-tight text-white">
              {kpi.value}
            </p>
            <p className="mt-1 text-sm text-zinc-400">{kpi.hint}</p>
          </article>
        );
      })}
    </div>
  );
}
