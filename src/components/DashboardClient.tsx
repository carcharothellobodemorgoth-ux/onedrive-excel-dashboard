"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import Link from "next/link";
import { AutoCharts } from "@/components/AutoCharts";
import { KpiCards } from "@/components/KpiCards";
import { PeriodSelector } from "@/components/PeriodSelector";
import { SheetTable } from "@/components/SheetTable";
import {
  buildProyeccionView,
  DEFAULT_PERIOD_SELECTION,
  maxDataCol,
  resolveSelectionCols,
  type PeriodSelection,
} from "@/lib/dashboard";

type Worksheet = { id: string; name: string; position: number };
type DriveItem = {
  id: string;
  name: string;
  webUrl?: string;
  lastModifiedDateTime?: string;
  parentReference?: { driveId?: string };
};
type SheetPayload = {
  worksheet: Worksheet;
  headers: string[];
  rows: (string | number | boolean | null)[][];
};

const ITEM_KEY = "excel-dashboard-item-id";
const DRIVE_KEY = "excel-dashboard-drive-id";
const PERIOD_KEY = "excel-dashboard-period";

function loadPeriodSelection(): PeriodSelection {
  if (typeof window === "undefined") return DEFAULT_PERIOD_SELECTION;
  try {
    const raw = localStorage.getItem(PERIOD_KEY);
    if (!raw) return DEFAULT_PERIOD_SELECTION;
    const parsed = JSON.parse(raw) as PeriodSelection;
    if (
      parsed?.mode === "quincena" ||
      parsed?.mode === "mes" ||
      parsed?.mode === "custom"
    ) {
      return parsed;
    }
  } catch {
    /* ignore */
  }
  return DEFAULT_PERIOD_SELECTION;
}

export function DashboardClient({
  userName,
}: {
  userName?: string | null;
}) {
  const [item, setItem] = useState<DriveItem | null>(null);
  const [itemId, setItemId] = useState<string | null>(null);
  const [driveId, setDriveId] = useState<string | null>(null);
  const [worksheets, setWorksheets] = useState<Worksheet[]>([]);
  const [candidates, setCandidates] = useState<DriveItem[]>([]);
  const [activeId, setActiveId] = useState<string | null>(null);
  const [sheet, setSheet] = useState<SheetPayload | null>(null);
  const [loading, setLoading] = useState(true);
  const [sheetLoading, setSheetLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [shareUrl, setShareUrl] = useState(
    "https://1drv.ms/x/c/d883d00740abbada/IQDVZuw8RO2ZR4CJqAH8k6nyAStmn-NqhF2_OUAcEu3D3_M?e=76bLdx",
  );
  const [shareBusy, setShareBusy] = useState(false);
  const [period, setPeriod] = useState<PeriodSelection>(DEFAULT_PERIOD_SELECTION);
  const [gastosVariosByQuincena, setGastosVariosByQuincena] = useState<
    Record<number, number>
  >({});

  useEffect(() => {
    setPeriod(loadPeriodSelection());
  }, []);

  const applySummary = useCallback((data: {
    item: DriveItem;
    itemId: string;
    driveId: string;
    worksheets: Worksheet[];
  }) => {
    setItem(data.item);
    setItemId(data.itemId);
    setDriveId(data.driveId);
    setCandidates([]);
    if (typeof window !== "undefined") {
      localStorage.setItem(ITEM_KEY, data.itemId);
      localStorage.setItem(DRIVE_KEY, data.driveId);
    }
    setWorksheets(data.worksheets ?? []);
    const preferred =
      (data.worksheets as Worksheet[] | undefined)?.find((w) =>
        /20262027/i.test(w.name),
      ) ?? data.worksheets?.[0];
    setActiveId(preferred?.id ?? null);
  }, []);

  const loadSummary = useCallback(
    async (preferred?: { itemId?: string | null; driveId?: string | null }) => {
      setLoading(true);
      setError(null);
      setCandidates([]);
      try {
        const storedItem =
          preferred?.itemId ??
          (typeof window !== "undefined" ? localStorage.getItem(ITEM_KEY) : null);
        const storedDrive =
          preferred?.driveId ??
          (typeof window !== "undefined"
            ? localStorage.getItem(DRIVE_KEY)
            : null);

        const qs = new URLSearchParams();
        if (storedItem && storedDrive) {
          qs.set("itemId", storedItem);
          qs.set("driveId", storedDrive);
        }
        const res = await fetch(
          `/api/excel${qs.toString() ? `?${qs}` : ""}`,
        );
        const raw = await res.text();
        const data = raw
          ? (JSON.parse(raw) as Record<string, unknown>)
          : {};

        if (res.status === 409 && Array.isArray(data.candidates)) {
          setCandidates(data.candidates as DriveItem[]);
          setError(
            typeof data.message === "string"
              ? data.message
              : "Elegí una planilla",
          );
          setItem(null);
          setItemId(null);
          setDriveId(null);
          setWorksheets([]);
          return;
        }

        if (!res.ok) {
          throw new Error(
            typeof data.error === "string"
              ? data.error
              : typeof data.message === "string"
                ? data.message
                : `Error HTTP ${res.status}`,
          );
        }
        applySummary(
          data as {
            item: DriveItem;
            itemId: string;
            driveId: string;
            worksheets: Worksheet[];
          },
        );
      } catch (e) {
        setError(e instanceof Error ? e.message : "Error desconocido");
      } finally {
        setLoading(false);
      }
    },
    [applySummary],
  );

  const openShareUrl = useCallback(async () => {
    if (!shareUrl.trim()) return;
    setShareBusy(true);
    setError(null);
    try {
      const res = await fetch("/api/excel", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ shareUrl: shareUrl.trim() }),
      });
      const raw = await res.text();
      const data = raw ? (JSON.parse(raw) as Record<string, unknown>) : {};
      if (!res.ok || data.ok === false) {
        throw new Error(
          typeof data.message === "string"
            ? data.message
            : typeof data.error === "string"
              ? data.error
              : `Error HTTP ${res.status}`,
        );
      }
      applySummary(
        data as {
          item: DriveItem;
          itemId: string;
          driveId: string;
          worksheets: Worksheet[];
        },
      );
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error desconocido");
    } finally {
      setShareBusy(false);
    }
  }, [shareUrl, applySummary]);

  const loadGastosVarios = useCallback(
    async (resolvedItemId: string, resolvedDriveId: string) => {
      try {
        const qs = new URLSearchParams({
          itemId: resolvedItemId,
          driveId: resolvedDriveId,
        });
        const res = await fetch(`/api/excel/gastos?${qs}`);
        const raw = await res.text();
        const data = raw ? (JSON.parse(raw) as Record<string, unknown>) : {};
        if (!res.ok) return;
        const by = data.byQuincena;
        if (by && typeof by === "object") {
          setGastosVariosByQuincena(by as Record<number, number>);
        }
      } catch {
        /* non-fatal: dashboard still works without varios */
      }
    },
    [],
  );

  const loadSheet = useCallback(
    async (
      worksheetId: string,
      resolvedItemId: string,
      resolvedDriveId: string,
    ) => {
      setSheetLoading(true);
      setError(null);
      try {
        const qs = new URLSearchParams({
          worksheetId,
          itemId: resolvedItemId,
          driveId: resolvedDriveId,
        });
        const res = await fetch(`/api/excel?${qs.toString()}`);
        const raw = await res.text();
        const data = raw ? (JSON.parse(raw) as Record<string, unknown>) : {};
        if (!res.ok) {
          throw new Error(
            typeof data.error === "string"
              ? data.error
              : `Error HTTP ${res.status}`,
          );
        }
        setSheet(data as SheetPayload);
        void loadGastosVarios(resolvedItemId, resolvedDriveId);
      } catch (e) {
        setError(e instanceof Error ? e.message : "Error desconocido");
        setSheet(null);
      } finally {
        setSheetLoading(false);
      }
    },
    [loadGastosVarios],
  );

  useEffect(() => {
    // Clear stale item-only cache from previous broken resolver
    if (typeof window !== "undefined") {
      const hasItem = localStorage.getItem(ITEM_KEY);
      const hasDrive = localStorage.getItem(DRIVE_KEY);
      if (hasItem && !hasDrive) {
        localStorage.removeItem(ITEM_KEY);
      }
    }
    void loadSummary({});
  }, [loadSummary]);

  useEffect(() => {
    if (activeId && itemId && driveId) {
      void loadSheet(activeId, itemId, driveId);
    }
  }, [activeId, itemId, driveId, loadSheet]);

  const view = useMemo(
    () =>
      sheet
        ? buildProyeccionView(sheet.rows, period, new Date(), {
            gastosVariosByQuincena,
          })
        : null,
    [sheet, period, gastosVariosByQuincena],
  );

  const periodMaxCol = sheet ? maxDataCol(sheet.rows) : 24;

  const selectedCols = useMemo(() => {
    if (!sheet) return [] as number[];
    return resolveSelectionCols(sheet.rows, period).cols;
  }, [sheet, period]);

  const applyPeriod = useCallback((next: PeriodSelection) => {
    setPeriod(next);
    if (typeof window !== "undefined") {
      localStorage.setItem(PERIOD_KEY, JSON.stringify(next));
    }
  }, []);

  const detailHeaders = ["Imputación", "Monto"];
  const incomeTableRows =
    view?.incomeRows.map((r) => [r.label, r.value] as (string | number)[]) ??
    [];
  const expenseTableRows =
    view?.expenseRows.map((r) => [r.label, r.value] as (string | number)[]) ??
    [];

  return (
    <div className="mx-auto flex w-full max-w-7xl flex-col gap-8 px-4 py-8 sm:px-6">
      <header className="flex flex-col gap-4 sm:flex-row sm:items-end sm:justify-between">
        <div>
          <p className="text-sm text-emerald-300/90">
            {userName ? `Hola, ${userName}` : "OneDrive Excel"}
          </p>
          <h1 className="mt-1 text-3xl font-semibold tracking-tight text-white">
            {item?.name ?? "Dashboard"}
          </h1>
          <p className="mt-1 text-sm text-zinc-400">
            Datos en vivo desde tu OneDrive vía Microsoft Graph
          </p>
        </div>
        <div className="flex flex-wrap gap-2">
          <Link
            href="/dashboard/gastos"
            className="rounded-full bg-emerald-500 px-4 py-2 text-sm font-semibold text-zinc-950 hover:bg-emerald-400"
          >
            + Cargar gasto
          </Link>
          <button
            type="button"
            onClick={() => {
              if (typeof window !== "undefined") {
                localStorage.removeItem(ITEM_KEY);
                localStorage.removeItem(DRIVE_KEY);
              }
              void loadSummary({ itemId: null, driveId: null });
            }}
            className="rounded-full border border-white/15 px-4 py-2 text-sm font-medium text-zinc-200 hover:bg-white/5"
          >
            Actualizar
          </button>
          <a
            href="/api/auth/signout"
            className="rounded-full border border-white/15 px-4 py-2 text-sm font-medium text-zinc-200 hover:bg-white/5"
          >
            Cerrar sesión
          </a>
        </div>
      </header>

      {error && candidates.length === 0 && (
        <div className="rounded-2xl border border-red-500/30 bg-red-500/10 px-4 py-3 text-sm text-red-200">
          <p>{error}</p>
          {/sesión|token|autentic|reauth|JWT|Microsoft/i.test(error) && (
            <p className="mt-2">
              <a
                href="/api/auth/signout"
                className="font-semibold text-white underline underline-offset-2 hover:text-emerald-200"
              >
                Cerrar sesión y volver a entrar
              </a>
            </p>
          )}
        </div>
      )}

      {candidates.length > 0 && (
        <section className="rounded-2xl border border-amber-400/30 bg-amber-400/10 p-5">
          <h2 className="text-lg font-semibold text-white">
            ¿Cuál es la planilla?
          </h2>
          <p className="mt-1 text-sm text-zinc-300">
            <strong>PROYECCIÓN 26-27</strong> no está en tu drive ni en
            “Compartido conmigo” (típico de links “cualquiera con el vínculo”).
            Pegá el link de <strong>Compartir → Copiar vínculo</strong> desde
            Excel:
          </p>
          <div className="mt-4 flex flex-col gap-2 sm:flex-row">
            <input
              value={shareUrl}
              onChange={(e) => setShareUrl(e.target.value)}
              placeholder="https://1drv.ms/... o link de Compartir"
              className="min-w-0 flex-1 rounded-xl border border-white/15 bg-zinc-950/60 px-3 py-2 text-sm text-white placeholder:text-zinc-500"
            />
            <button
              type="button"
              disabled={shareBusy || !shareUrl.trim()}
              onClick={() => void openShareUrl()}
              className="rounded-full bg-white px-4 py-2 text-sm font-semibold text-zinc-950 disabled:opacity-50"
            >
              {shareBusy ? "Abriendo…" : "Abrir link"}
            </button>
          </div>
          <p className="mt-3 text-xs text-zinc-400">
            La URL del navegador (`doc.aspx?resid=...`) suele fallar; el vínculo
            de Compartir sí funciona.
          </p>
          <ul className="mt-4 grid gap-2 sm:grid-cols-2">
            {candidates.map((c) => (
              <li key={c.id}>
                <button
                  type="button"
                  onClick={() => {
                    const d = c.parentReference?.driveId ?? null;
                    if (typeof window !== "undefined") {
                      localStorage.setItem(ITEM_KEY, c.id);
                      if (d) localStorage.setItem(DRIVE_KEY, d);
                    }
                    void loadSummary({ itemId: c.id, driveId: d });
                  }}
                  className="w-full rounded-xl border border-white/10 bg-zinc-950/40 px-4 py-3 text-left hover:border-emerald-400/40 hover:bg-emerald-500/10"
                >
                  <span className="block font-medium text-white">{c.name}</span>
                  <span className="mt-1 block truncate text-xs text-zinc-500">
                    {c.lastModifiedDateTime
                      ? new Date(c.lastModifiedDateTime).toLocaleString("es")
                      : c.id}
                  </span>
                </button>
              </li>
            ))}
          </ul>
        </section>
      )}

      {!loading && !item && candidates.length === 0 && (
        <section className="rounded-2xl border border-white/10 bg-white/5 p-5">
          <h2 className="text-lg font-semibold text-white">
            Abrir planilla por link
          </h2>
          <p className="mt-1 text-sm text-zinc-300">
            En Excel Online: <strong>Compartir → Copiar vínculo</strong> y
            pegalo acá.
          </p>
          <div className="mt-4 flex flex-col gap-2 sm:flex-row">
            <input
              value={shareUrl}
              onChange={(e) => setShareUrl(e.target.value)}
              placeholder="https://1drv.ms/..."
              className="min-w-0 flex-1 rounded-xl border border-white/15 bg-zinc-950/60 px-3 py-2 text-sm text-white placeholder:text-zinc-500"
            />
            <button
              type="button"
              disabled={shareBusy || !shareUrl.trim()}
              onClick={() => void openShareUrl()}
              className="rounded-full bg-emerald-500 px-4 py-2 text-sm font-semibold text-zinc-950 disabled:opacity-50"
            >
              {shareBusy ? "Abriendo…" : "Abrir link"}
            </button>
          </div>
        </section>
      )}

      {loading ? (
        <p className="text-zinc-400">Cargando workbook…</p>
      ) : (
        item && (
          <>
            <nav className="flex flex-wrap gap-2">
              {worksheets.map((ws) => (
                <button
                  key={ws.id}
                  type="button"
                  onClick={() => setActiveId(ws.id)}
                  className={`rounded-full px-4 py-2 text-sm font-medium transition ${
                    activeId === ws.id
                      ? "bg-white text-zinc-950"
                      : "bg-white/5 text-zinc-300 hover:bg-white/10"
                  }`}
                >
                  {ws.name}
                </button>
              ))}
            </nav>

            {sheetLoading && (
              <p className="text-sm text-zinc-400">Cargando hoja…</p>
            )}

            {sheet && !sheetLoading && view && (
              <div className="flex flex-col gap-6">
                <PeriodSelector
                  value={period}
                  maxCol={periodMaxCol}
                  onChange={applyPeriod}
                />
                <p className="text-sm text-emerald-300/90">
                  Hoy: <strong className="text-white">{view.todayLabel}</strong>
                  {" · "}
                  Periodo:{" "}
                  <strong className="text-white">{view.periodLabel}</strong>
                  {" · "}
                  hoja {sheet.worksheet.name}
                </p>
                {!view.inCycle && (
                  <div className="rounded-2xl border border-amber-400/30 bg-amber-400/10 px-4 py-3 text-sm text-amber-100">
                    La fecha de hoy no cae en el ciclo de la planilla
                    (ago-2026 → jul-2027). Elegí otra quincena/mes en el selector
                    o un rango custom. Quincenas: 1–15 = 1ª, 16–fin = 2ª.
                  </div>
                )}
                <KpiCards kpis={view.kpis} />
                <AutoCharts
                  charts={view.charts}
                  selectedCols={selectedCols}
                  onSelectCol={(col) =>
                    applyPeriod({ mode: "quincena", col })
                  }
                />
                <div className="grid gap-6 lg:grid-cols-2">
                  <section>
                    <h2 className="mb-3 text-lg font-semibold text-white">
                      Ingresos (filas 1–3)
                    </h2>
                    <SheetTable
                      headers={detailHeaders}
                      rows={incomeTableRows}
                    />
                  </section>
                  <section>
                    <h2 className="mb-3 text-lg font-semibold text-white">
                      Gastos del periodo
                    </h2>
                    <SheetTable
                      headers={detailHeaders}
                      rows={expenseTableRows}
                    />
                  </section>
                </div>
              </div>
            )}
          </>
        )
      )}
    </div>
  );
}
