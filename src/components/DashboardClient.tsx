"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import { AutoCharts } from "@/components/AutoCharts";
import { KpiCards } from "@/components/KpiCards";
import { SheetTable } from "@/components/SheetTable";
import { buildCharts, buildKpis } from "@/lib/dashboard";

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
    setActiveId(data.worksheets?.[0]?.id ?? null);
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
        const data = await res.json();

        if (res.status === 409 && data.candidates) {
          setCandidates(data.candidates);
          setError(data.message ?? "Elegí una planilla");
          setItem(null);
          setItemId(null);
          setDriveId(null);
          setWorksheets([]);
          return;
        }

        if (!res.ok) throw new Error(data.error ?? data.message ?? "No se pudo leer la Excel");
        applySummary(data);
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
      const data = await res.json();
      if (!res.ok || data.ok === false) {
        throw new Error(data.message ?? data.error ?? "No se pudo abrir el link");
      }
      applySummary(data);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error desconocido");
    } finally {
      setShareBusy(false);
    }
  }, [shareUrl, applySummary]);

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
        const data = await res.json();
        if (!res.ok) throw new Error(data.error ?? "No se pudo leer la hoja");
        setSheet(data);
      } catch (e) {
        setError(e instanceof Error ? e.message : "Error desconocido");
        setSheet(null);
      } finally {
        setSheetLoading(false);
      }
    },
    [],
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

  const kpis = useMemo(
    () => (sheet ? buildKpis(sheet.headers, sheet.rows) : []),
    [sheet],
  );
  const charts = useMemo(
    () => (sheet ? buildCharts(sheet.headers, sheet.rows) : []),
    [sheet],
  );

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
          <button
            type="button"
            onClick={() => {
              if (typeof window !== "undefined") {
                localStorage.removeItem(ITEM_KEY);
                localStorage.removeItem(DRIVE_KEY);
              }
              void loadSummary({ itemId: null, driveId: null });
            }}
            className="rounded-full bg-emerald-500 px-4 py-2 text-sm font-semibold text-zinc-950 hover:bg-emerald-400"
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
          {error}
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

            {sheet && !sheetLoading && (
              <div className="flex flex-col gap-6">
                <KpiCards kpis={kpis} />
                <AutoCharts charts={charts} />
                <section>
                  <h2 className="mb-3 text-lg font-semibold text-white">
                    {sheet.worksheet.name}
                  </h2>
                  <SheetTable headers={sheet.headers} rows={sheet.rows} />
                </section>
              </div>
            )}
          </>
        )
      )}
    </div>
  );
}
