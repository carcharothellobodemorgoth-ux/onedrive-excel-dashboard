"use client";

import type { PeriodSelection } from "@/lib/dashboard";
import { listMonthOptions, listQuincenaOptions } from "@/lib/dashboard";

const selectClass =
  "min-w-0 max-w-[16rem] rounded-lg border border-white/15 bg-zinc-950/60 px-2.5 py-1.5 text-sm text-white";

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
    <div className="flex flex-wrap items-center gap-2">
      <span className="text-xs font-semibold uppercase tracking-wide text-zinc-500">
        Periodo
      </span>
      <div className="inline-flex rounded-lg border border-white/10 bg-white/5 p-0.5">
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
                  onChange({
                    mode: "custom",
                    fromCol: 1,
                    toCol: Math.min(2, maxCol),
                  });
                }
              }}
              className={`rounded-md px-2.5 py-1 text-xs font-medium transition ${
                active
                  ? "bg-white text-zinc-950"
                  : "text-zinc-400 hover:text-zinc-200"
              }`}
            >
              {label}
            </button>
          );
        })}
      </div>

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
          <select
            className={selectClass}
            aria-label="Desde"
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
                Desde · {q.label}
              </option>
            ))}
          </select>
          <select
            className={selectClass}
            aria-label="Hasta"
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
                Hasta · {q.label}
              </option>
            ))}
          </select>
        </>
      )}
    </div>
  );
}
