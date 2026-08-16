"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import Link from "next/link";
import {
  detectCurrentQuincenaCol,
  listExpenseCategories,
  listQuincenaOptions,
  quincenaLabel,
} from "@/lib/dashboard";
import { apiFetch } from "@/lib/api-fetch";
import { isGastosVariosExpenseRow } from "@/lib/graph";

type Worksheet = { id: string; name: string; position: number };
type DriveItem = {
  id: string;
  name: string;
  parentReference?: { driveId?: string };
};

const ITEM_KEY = "excel-dashboard-item-id";
const DRIVE_KEY = "excel-dashboard-drive-id";

const inputClass =
  "w-full rounded-xl border border-white/15 bg-zinc-950/60 px-3 py-2.5 text-sm text-white placeholder:text-zinc-500";
const labelClass = "mb-1.5 block text-xs font-semibold uppercase tracking-wide text-zinc-400";

export function GastosClient({ userName }: { userName?: string | null }) {
  const [itemId, setItemId] = useState<string | null>(null);
  const [driveId, setDriveId] = useState<string | null>(null);
  const [itemName, setItemName] = useState<string | null>(null);
  const [categories, setCategories] = useState<string[]>([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [okMsg, setOkMsg] = useState<string | null>(null);

  const [description, setDescription] = useState("");
  const [category, setCategory] = useState("");
  const [amount, setAmount] = useState("");
  const [quincena, setQuincena] = useState<number>(1);
  const [recent, setRecent] = useState<
    { description: string; category: string; amount: number; quincena: number }[]
  >([]);

  const quincenaOptions = useMemo(() => listQuincenaOptions(24), []);

  const loadContext = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const storedItem =
        typeof window !== "undefined" ? localStorage.getItem(ITEM_KEY) : null;
      const storedDrive =
        typeof window !== "undefined" ? localStorage.getItem(DRIVE_KEY) : null;

      const qs = new URLSearchParams();
      if (storedItem && storedDrive) {
        qs.set("itemId", storedItem);
        qs.set("driveId", storedDrive);
      }
      const res = await apiFetch(`/api/excel${qs.toString() ? `?${qs}` : ""}`);
      const raw = await res.text();
      const data = raw ? (JSON.parse(raw) as Record<string, unknown>) : {};
      if (!res.ok) {
        throw new Error(
          typeof data.error === "string"
            ? data.error
            : typeof data.message === "string"
              ? data.message
              : `Error HTTP ${res.status}`,
        );
      }

      const worksheets = (data.worksheets as Worksheet[]) ?? [];
      const id = data.itemId as string;
      const dId = data.driveId as string;
      const item = data.item as DriveItem | undefined;

      setItemId(id);
      setDriveId(dId);
      setItemName(item?.name ?? null);
      if (typeof window !== "undefined") {
        localStorage.setItem(ITEM_KEY, id);
        localStorage.setItem(DRIVE_KEY, dId);
      }

      const main =
        worksheets.find((w) => /20262027/i.test(w.name)) ?? worksheets[0];
      if (main) {
        const sheetRes = await apiFetch(
          `/api/excel?${new URLSearchParams({
            worksheetId: main.id,
            itemId: id,
            driveId: dId,
          })}`,
        );
        const sheetRaw = await sheetRes.text();
        const sheetData = sheetRaw
          ? (JSON.parse(sheetRaw) as {
              rows?: (string | number | boolean | null)[][];
            })
          : {};
        if (sheetRes.ok && sheetData.rows) {
          const cats = listExpenseCategories(sheetData.rows);
          setCategories(cats);
          setCategory((prev) => prev || cats[0] || "");
          setQuincena(detectCurrentQuincenaCol(sheetData.rows));
        }
      }

      const gvRes = await apiFetch(
        `/api/excel/gastos?${new URLSearchParams({ itemId: id, driveId: dId })}`,
      );
      const gvRaw = await gvRes.text();
      const gvData = gvRaw ? (JSON.parse(gvRaw) as Record<string, unknown>) : {};
      if (gvRes.ok && Array.isArray(gvData.rows)) {
        const rows = gvData.rows as (string | number | boolean | null)[][];
        const parsed: {
          description: string;
          category: string;
          amount: number;
          quincena: number;
        }[] = [];
        for (let r = rows.length - 1; r >= 0 && parsed.length < 12; r--) {
          const row = rows[r];
          if (!isGastosVariosExpenseRow(row)) continue;
          const desc = String(row?.[0] ?? "").trim();
          const cat = String(row?.[1] ?? "").trim();
          let foundQ = 0;
          let foundAmt = 0;
          for (let q = 1; q <= 24; q++) {
            const cell = row?.[2 + q - 1];
            const n =
              typeof cell === "number"
                ? cell
                : typeof cell === "string"
                  ? Number(cell.replace(",", "."))
                  : 0;
            if (Number.isFinite(n) && n !== 0) {
              foundQ = q;
              foundAmt = Math.abs(n);
              break;
            }
          }
          if (foundQ) {
            parsed.push({
              description: desc,
              category: cat,
              amount: foundAmt,
              quincena: foundQ,
            });
          }
        }
        setRecent(parsed);
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error desconocido");
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    void loadContext();
  }, [loadContext]);

  const submit = useCallback(async () => {
    if (!itemId || !driveId) return;
    setSaving(true);
    setError(null);
    setOkMsg(null);
    try {
      const res = await apiFetch("/api/excel/gastos", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          description,
          category,
          amount,
          quincena,
          itemId,
          driveId,
        }),
      });
      const raw = await res.text();
      const data = raw ? (JSON.parse(raw) as Record<string, unknown>) : {};
      if (!res.ok) {
        throw new Error(
          typeof data.error === "string"
            ? data.error
            : `Error HTTP ${res.status}`,
        );
      }
      setOkMsg(
        `Guardado en fila ${data.row as number} · ${quincenaLabel(quincena)}`,
      );
      setDescription("");
      setAmount("");
      await loadContext();
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al guardar");
    } finally {
      setSaving(false);
    }
  }, [
    amount,
    category,
    description,
    driveId,
    itemId,
    loadContext,
    quincena,
  ]);

  const money = (n: number) =>
    new Intl.NumberFormat("es-AR", {
      style: "currency",
      currency: "ARS",
      maximumFractionDigits: 2,
    }).format(n);

  return (
    <div className="mx-auto flex w-full max-w-3xl flex-col gap-8 px-4 py-8 sm:px-6">
      <header className="flex flex-col gap-4 sm:flex-row sm:items-end sm:justify-between">
        <div>
          <p className="text-sm text-emerald-300/90">
            {userName ? `Hola, ${userName}` : "OneDrive Excel"}
          </p>
          <h1 className="mt-1 text-3xl font-semibold tracking-tight text-white">
            Cargar gasto
          </h1>
          <p className="mt-1 text-sm text-zinc-400">
            {itemName
              ? `Se escribe en la hoja «Gastos Varios» de ${itemName}`
              : "Se escribe en la hoja Gastos Varios del workbook"}
          </p>
        </div>
        <nav className="flex flex-wrap gap-2">
          <Link
            href="/dashboard"
            className="rounded-full border border-white/15 px-4 py-2 text-sm font-medium text-zinc-200 hover:bg-white/5"
          >
            Proyección
          </Link>
          <span className="rounded-full bg-white px-4 py-2 text-sm font-semibold text-zinc-950">
            Cargar gasto
          </span>
          <a
            href="/api/auth/signout"
            className="rounded-full border border-white/15 px-4 py-2 text-sm font-medium text-zinc-200 hover:bg-white/5"
          >
            Cerrar sesión
          </a>
        </nav>
      </header>

      {error && (
        <div className="rounded-2xl border border-red-500/30 bg-red-500/10 px-4 py-3 text-sm text-red-200">
          <p>{error}</p>
          {/sesión|permisos|Files\.ReadWrite|reauth|entrar/i.test(error) && (
            <p className="mt-2">
              <a
                href="/api/auth/signout"
                className="font-semibold text-white underline underline-offset-2"
              >
                Cerrar sesión y volver a entrar
              </a>{" "}
              (aceptá los nuevos permisos de escritura en Azure / Microsoft).
            </p>
          )}
        </div>
      )}

      {okMsg && (
        <div className="rounded-2xl border border-emerald-500/30 bg-emerald-500/10 px-4 py-3 text-sm text-emerald-100">
          {okMsg}
        </div>
      )}

      {loading ? (
        <p className="text-zinc-400">Cargando categorías y quincenas…</p>
      ) : (
        <form
          className="flex flex-col gap-5 rounded-2xl border border-white/10 bg-white/5 p-5"
          onSubmit={(e) => {
            e.preventDefault();
            void submit();
          }}
        >
          <div>
            <label className={labelClass} htmlFor="desc">
              Descripción <span className="text-red-400">*</span>
            </label>
            <input
              id="desc"
              className={inputClass}
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              placeholder="Ej. Supermercado, nafta, farmacia…"
              required
            />
          </div>

          <div>
            <label className={labelClass} htmlFor="amount">
              Monto <span className="text-red-400">*</span>
            </label>
            <input
              id="amount"
              className={inputClass}
              inputMode="decimal"
              value={amount}
              onChange={(e) => setAmount(e.target.value)}
              placeholder="15000"
              required
            />
          </div>

          <div>
            <label className={labelClass} htmlFor="cat">
              Categoría <span className="text-red-400">*</span>
            </label>
            {categories.length > 0 ? (
              <select
                id="cat"
                className={inputClass}
                value={category}
                onChange={(e) => setCategory(e.target.value)}
                required
              >
                {categories.map((c) => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </select>
            ) : (
              <input
                id="cat"
                className={inputClass}
                value={category}
                onChange={(e) => setCategory(e.target.value)}
                placeholder="Categoría"
                required
              />
            )}
            <p className="mt-1.5 text-xs text-zinc-500">
              Categorías tomadas de la columna B de la hoja 20262027.
            </p>
          </div>

          <div>
            <label className={labelClass} htmlFor="q">
              Quincena <span className="text-red-400">*</span>
            </label>
            <select
              id="q"
              className={inputClass}
              value={quincena}
              onChange={(e) => setQuincena(Number(e.target.value))}
              required
            >
              {quincenaOptions.map((q) => (
                <option key={q.value} value={q.value}>
                  {q.label}
                </option>
              ))}
            </select>
          </div>

          <div className="flex justify-end gap-2 pt-2">
            <button
              type="submit"
              disabled={saving || !description.trim() || !amount}
              className="rounded-full bg-emerald-500 px-5 py-2.5 text-sm font-semibold text-zinc-950 hover:bg-emerald-400 disabled:opacity-50"
            >
              {saving ? "Guardando…" : "Guardar gasto"}
            </button>
          </div>
        </form>
      )}

      {recent.length > 0 && (
        <section>
          <h2 className="mb-3 text-lg font-semibold text-white">
            Últimos gastos
          </h2>
          <ul className="divide-y divide-white/10 rounded-2xl border border-white/10 bg-white/5">
            {recent.map((r, i) => (
              <li
                key={`${r.description}-${i}`}
                className="flex flex-col gap-1 px-4 py-3 sm:flex-row sm:items-center sm:justify-between"
              >
                <div>
                  <p className="font-medium text-white">{r.description}</p>
                  <p className="text-xs text-zinc-400">
                    {r.category || "Sin categoría"} · {quincenaLabel(r.quincena)}
                  </p>
                </div>
                <p className="text-sm font-semibold text-emerald-300">
                  {money(r.amount)}
                </p>
              </li>
            ))}
          </ul>
        </section>
      )}
    </div>
  );
}
