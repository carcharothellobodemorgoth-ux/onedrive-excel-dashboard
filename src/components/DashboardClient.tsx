"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import { AutoCharts } from "@/components/AutoCharts";
import { KpiCards } from "@/components/KpiCards";
import { SheetTable } from "@/components/SheetTable";
import { buildCharts, buildKpis } from "@/lib/dashboard";

type Worksheet = { id: string; name: string; position: number };
type DriveItem = { id: string; name: string; webUrl?: string };
type SheetPayload = {
  worksheet: Worksheet;
  headers: string[];
  rows: (string | number | boolean | null)[][];
};

export function DashboardClient({
  userName,
}: {
  userName?: string | null;
}) {
  const [item, setItem] = useState<DriveItem | null>(null);
  const [itemId, setItemId] = useState<string | null>(null);
  const [worksheets, setWorksheets] = useState<Worksheet[]>([]);
  const [activeId, setActiveId] = useState<string | null>(null);
  const [sheet, setSheet] = useState<SheetPayload | null>(null);
  const [loading, setLoading] = useState(true);
  const [sheetLoading, setSheetLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const loadSummary = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetch("/api/excel");
      const data = await res.json();
      if (!res.ok) throw new Error(data.error ?? "No se pudo leer la Excel");
      setItem(data.item);
      setItemId(data.itemId ?? data.item?.id ?? null);
      setWorksheets(data.worksheets ?? []);
      const first = data.worksheets?.[0]?.id ?? null;
      setActiveId(first);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error desconocido");
    } finally {
      setLoading(false);
    }
  }, []);

  const loadSheet = useCallback(
    async (worksheetId: string, resolvedItemId: string) => {
      setSheetLoading(true);
      setError(null);
      try {
        const qs = new URLSearchParams({
          worksheetId,
          itemId: resolvedItemId,
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
    void loadSummary();
  }, [loadSummary]);

  useEffect(() => {
    if (activeId && itemId) void loadSheet(activeId, itemId);
  }, [activeId, itemId, loadSheet]);

  const kpis = useMemo(
    () =>
      sheet ? buildKpis(sheet.headers, sheet.rows) : [],
    [sheet],
  );
  const charts = useMemo(
    () =>
      sheet ? buildCharts(sheet.headers, sheet.rows) : [],
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
            onClick={() => void loadSummary()}
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

      {error && (
        <div className="rounded-2xl border border-red-500/30 bg-red-500/10 px-4 py-3 text-sm text-red-200">
          {error}
        </div>
      )}

      {loading ? (
        <p className="text-zinc-400">Cargando workbook…</p>
      ) : (
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
      )}
    </div>
  );
}
