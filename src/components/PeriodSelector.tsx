"use client";

import type { PeriodSelection } from "@/lib/dashboard";
import { listMonthOptions, listQuincenaOptions } from "@/lib/dashboard";

const selectClass =
  "rounded-xl border border-white/15 bg-zinc-950/60 px-3 py-2 text-sm text-white";

export function PeriodSelector({
  value,
  onChange,
  maxCol = 24,
}: {
  value: PeriodSelection;
  onChange: (next: PeriodSelection) => void;
  maxCol?: number;
}) {
  const quincenas = listQuincenaOptions(maxCol);
  const months = listMonthOptions();

  return (
    <section className="rounded-2xl border border-white/10 bg-white/5 p-4">
      <p className="mb-3 text-xs font-semibold uppercase tracking-wide text-zinc-400">
        Periodo
      </p>
      <div className="flex flex-wrap gap-2">
        {(
          [
            ["quincena", "Quincena"],
            ["mes", "Mes"],
            ["custom", "Custom"],
          ] as const
        ).map(([mode, label]) => {
          const active = value.mode === mode;
          return (
            <button
              key={mode}
              type="button"
              onClick={() => {
                if (mode === "quincena") {
                  onChange({ mode: "quincena", col: "actual" });
                } else if (mode === "mes") {
                  onChange({ mode: "mes", monthIndex: "actual" });
                } else {
                  onChange({ mode: "custom", fromCol: 1, toCol: Math.min(2, maxCol) });
                }
              }}
              className={`rounded-full px-4 py-2 text-sm font-medium transition ${
                active
                  ? "bg-white text-zinc-950"
                  : "bg-white/5 text-zinc-300 hover:bg-white/10"
              }`}
            >
              {label}
            </button>
          );
        })}
      </div>

      <div className="mt-4 flex flex-wrap items-center gap-3">
        {value.mode === "quincena" && (
          <select
            className={selectClass}
            value={value.col === "actual" ? "actual" : String(value.col)}
            onChange={(e) => {
              const v = e.target.value;
              onChange({
                mode: "quincena",
                col: v === "actual" ? "actual" : Number(v),
              });
            }}
          >
            <option value="actual">Actual (según hoy)</option>
            {quincenas.map((q) => (
              <option key={q.value} value={q.value}>
                {q.label}
              </option>
            ))}
          </select>
        )}

        {value.mode === "mes" && (
          <select
            className={selectClass}
            value={
              value.monthIndex === "actual"
                ? "actual"
                : String(value.monthIndex)
            }
            onChange={(e) => {
              const v = e.target.value;
              onChange({
                mode: "mes",
                monthIndex: v === "actual" ? "actual" : Number(v),
              });
            }}
          >
            <option value="actual">Actual (según hoy)</option>
            {months.map((m) => (
              <option key={m.value} value={m.value}>
                {m.label}
              </option>
            ))}
          </select>
        )}

        {value.mode === "custom" && (
          <>
            <label className="flex items-center gap-2 text-sm text-zinc-300">
              Desde
              <select
                className={selectClass}
                value={value.fromCol}
                onChange={(e) =>
                  onChange({
                    ...value,
                    fromCol: Number(e.target.value),
                  })
                }
              >
                {quincenas.map((q) => (
                  <option key={q.value} value={q.value}>
                    {q.label}
                  </option>
                ))}
              </select>
            </label>
            <label className="flex items-center gap-2 text-sm text-zinc-300">
              Hasta
              <select
                className={selectClass}
                value={value.toCol}
                onChange={(e) =>
                  onChange({
                    ...value,
                    toCol: Number(e.target.value),
                  })
                }
              >
                {quincenas.map((q) => (
                  <option key={q.value} value={q.value}>
                    {q.label}
                  </option>
                ))}
              </select>
            </label>
          </>
        )}
      </div>
    </section>
  );
}
