"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import Link from "next/link";
import { flushSync } from "react-dom";
import {
  detectCurrentQuincenaCol,
  listExpenseCategories,
  listQuincenaOptions,
  quincenaLabel,
} from "@/lib/dashboard";
import { apiFetch } from "@/lib/api-fetch";
import { isGastosVariosExpenseRow } from "@/lib/graph";
import { pressable, tapFeedback } from "@/lib/tap";

type Worksheet = { id: string; name: string; position: number };
type DriveItem = {
  id: string;
  name: string;
  parentReference?: { driveId?: string };
};

type GastoItem = {
  excelRow: number;
  description: string;
  category: string;
  amount: number;
  quincena: number;
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
  /** Filter for the expense list: "all" or quincena 1–24 */
  const [listQuincena, setListQuincena] = useState<number | "all">("all");
  const [recent, setRecent] = useState<GastoItem[]>([]);
  const [editing, setEditing] = useState<{
    excelRow: number;
    quincena: number;
  } | null>(null);
  const [busyKey, setBusyKey] = useState<string | null>(null);
  const [confirmKey, setConfirmKey] = useState<string | null>(null);
  const [formFlash, setFormFlash] = useState(false);

  const quincenaOptions = useMemo(() => listQuincenaOptions(24), []);

  const filteredRecent = useMemo(() => {
    if (listQuincena === "all") return recent;
    return recent.filter((r) => r.quincena === listQuincena);
  }, [recent, listQuincena]);

  const filteredTotal = useMemo(
    () => filteredRecent.reduce((acc, r) => acc + r.amount, 0),
    [filteredRecent],
  );

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
          const currentQ = detectCurrentQuincenaCol(sheetData.rows);
          setQuincena(currentQ);
          setListQuincena((prev) => (prev === "all" ? currentQ : prev));
        }
      }

      const gvRes = await apiFetch(
        `/api/excel/gastos?${new URLSearchParams({ itemId: id, driveId: dId })}`,
      );
      const gvRaw = await gvRes.text();
      const gvData = gvRaw ? (JSON.parse(gvRaw) as Record<string, unknown>) : {};
      if (gvRes.ok && Array.isArray(gvData.rows)) {
        const rows = gvData.rows as (string | number | boolean | null)[][];
        const parsed: GastoItem[] = [];
        for (let r = rows.length - 1; r >= 0; r--) {
          const row = rows[r];
          if (!isGastosVariosExpenseRow(row)) continue;
          const desc =
            String(row?.[0] ?? "").trim() ||
            String(row?.[1] ?? "").trim() ||
            "Sin descripción";
          const cat = String(row?.[0] ?? "").trim()
            ? String(row?.[1] ?? "").trim()
            : "";
          for (let q = 1; q <= 24; q++) {
            const cell = row?.[2 + q - 1];
            const n =
              typeof cell === "number"
                ? cell
                : typeof cell === "string"
                  ? Number(
                      cell
                        .replace(/\$/g, "")
                        .replace(/\s/g, "")
                        .replace(/\./g, "")
                        .replace(",", "."),
                    )
                  : 0;
            if (Number.isFinite(n) && n !== 0) {
              parsed.push({
                excelRow: r + 1,
                description: desc,
                category: cat,
                amount: Math.abs(n),
                quincena: q,
              });
            }
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
    tapFeedback();
    flushSync(() => {
      setSaving(true);
      setError(null);
      setOkMsg(null);
    });
    try {
      const res = await apiFetch("/api/excel/gastos", {
        method: editing ? "PATCH" : "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          description,
          category,
          amount,
          quincena,
          itemId,
          driveId,
          ...(editing
            ? { row: editing.excelRow, previousQuincena: editing.quincena }
            : {}),
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
        editing
          ? `Actualizado · ${quincenaLabel(quincena)}`
          : `Guardado en fila ${data.row as number} · ${quincenaLabel(quincena)}`,
      );
      setDescription("");
      setAmount("");
      setEditing(null);
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
    editing,
    itemId,
    loadContext,
    quincena,
  ]);

  const startEdit = useCallback((item: GastoItem) => {
    tapFeedback();
    setConfirmKey(null);
    flushSync(() => {
      setEditing({ excelRow: item.excelRow, quincena: item.quincena });
      setDescription(item.description === "Sin descripción" ? "" : item.description);
      setCategory(item.category || categories[0] || "");
      setAmount(String(item.amount));
      setQuincena(item.quincena);
      setOkMsg(null);
      setError(null);
      setFormFlash(true);
    });
    document.getElementById("gasto-form")?.scrollIntoView({
      behavior: "auto",
      block: "start",
    });
    window.setTimeout(() => setFormFlash(false), 700);
  }, [categories]);

  const cancelEdit = useCallback(() => {
    tapFeedback();
    setEditing(null);
    setDescription("");
    setAmount("");
    setOkMsg(null);
    setFormFlash(false);
  }, []);

  const removeGasto = useCallback(
    async (item: GastoItem) => {
      if (!itemId || !driveId) return;
      const key = `${item.excelRow}-${item.quincena}`;
      tapFeedback();
      flushSync(() => {
        setBusyKey(key);
        setConfirmKey(null);
        setError(null);
        setOkMsg(null);
      });
      try {
        const qs = new URLSearchParams({
          itemId,
          driveId,
          row: String(item.excelRow),
          quincena: String(item.quincena),
        });
        const res = await apiFetch(`/api/excel/gastos?${qs}`, {
          method: "DELETE",
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
        if (
          editing &&
          editing.excelRow === item.excelRow &&
          editing.quincena === item.quincena
        ) {
          setEditing(null);
          setDescription("");
          setAmount("");
        }
        setOkMsg("Gasto eliminado");
        await loadContext();
      } catch (e) {
        setError(e instanceof Error ? e.message : "Error al eliminar");
      } finally {
        setBusyKey(null);
      }
    },
    [driveId, editing, itemId, loadContext],
  );

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
            {editing ? "Editar gasto" : "Cargar gasto"}
          </h1>
          <p className="mt-1 text-sm text-zinc-400">
            {itemName
              ? `Se escribe en la hoja «Gastos Varios» de ${itemName}`
              : "Se escribe en la hoja Gastos Varios del workbook"}
          </p>
        </div>
        <nav className="flex shrink-0 flex-nowrap items-center gap-1.5 sm:gap-2">
          <Link
            href="/dashboard"
            className={`${pressable} whitespace-nowrap rounded-full border border-white/15 px-2.5 py-1.5 text-xs font-medium text-zinc-200 hover:bg-white/5 sm:px-4 sm:py-2 sm:text-sm`}
          >
            Proyección
          </Link>
          <span className="whitespace-nowrap rounded-full bg-white px-2.5 py-1.5 text-xs font-semibold text-zinc-950 sm:px-4 sm:py-2 sm:text-sm">
            Cargar gasto
          </span>
          <a
            href="/api/auth/signout"
            onClick={() => tapFeedback()}
            className={`${pressable} whitespace-nowrap rounded-full border border-white/15 px-2.5 py-1.5 text-xs font-medium text-zinc-200 hover:bg-white/5 sm:px-4 sm:py-2 sm:text-sm`}
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
          id="gasto-form"
          className={`flex flex-col gap-5 rounded-2xl border bg-white/5 p-5 transition-shadow duration-300 ${
            formFlash
              ? "border-emerald-400/60 shadow-[0_0_0_3px_rgba(52,211,153,0.35)]"
              : editing
                ? "border-emerald-400/40"
                : "border-white/10"
          }`}
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
                {category && !categories.includes(category) && (
                  <option value={category}>{category}</option>
                )}
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
            {editing && (
              <button
                type="button"
                onClick={cancelEdit}
                disabled={saving}
                className={`${pressable} rounded-full border border-white/15 px-5 py-2.5 text-sm font-medium text-zinc-200 hover:bg-white/5 disabled:opacity-50`}
              >
                Cancelar
              </button>
            )}
            <button
              type="submit"
              disabled={saving || !description.trim() || !amount}
              className={`${pressable} rounded-full bg-emerald-500 px-5 py-2.5 text-sm font-semibold text-zinc-950 hover:bg-emerald-400 disabled:opacity-50`}
            >
              {saving
                ? "Guardando…"
                : editing
                  ? "Guardar cambios"
                  : "Guardar gasto"}
            </button>
          </div>
        </form>
      )}

      {recent.length > 0 && (
        <section>
          <div className="mb-3 flex flex-col gap-3 sm:flex-row sm:items-end sm:justify-between">
            <h2 className="text-lg font-semibold text-white">Gastos cargados</h2>
            <div className="sm:w-72">
              <label className={labelClass} htmlFor="list-q">
                Filtrar por quincena
              </label>
              <select
                id="list-q"
                className={inputClass}
                value={listQuincena === "all" ? "all" : String(listQuincena)}
                onChange={(e) => {
                  const v = e.target.value;
                  setListQuincena(v === "all" ? "all" : Number(v));
                }}
              >
                <option value="all">Todas las quincenas</option>
                {quincenaOptions.map((q) => (
                  <option key={q.value} value={q.value}>
                    {q.label}
                  </option>
                ))}
              </select>
            </div>
          </div>
          <p className="mb-2 text-sm text-zinc-400">
            {filteredRecent.length} gasto
            {filteredRecent.length === 1 ? "" : "s"}
            {" · "}
            Total:{" "}
            <strong className="text-emerald-300">{money(filteredTotal)}</strong>
          </p>
          {filteredRecent.length === 0 ? (
            <p className="rounded-2xl border border-white/10 bg-white/5 px-4 py-6 text-sm text-zinc-400">
              No hay gastos en esta quincena.
            </p>
          ) : (
            <ul className="divide-y divide-white/10 rounded-2xl border border-white/10 bg-white/5">
              {filteredRecent.map((r) => {
                const key = `${r.excelRow}-${r.quincena}`;
                const busy = busyKey === key;
                const confirming = confirmKey === key;
                const isEditing =
                  editing?.excelRow === r.excelRow &&
                  editing?.quincena === r.quincena;
                return (
                <li
                  key={key}
                  className={`flex flex-col gap-3 px-4 py-4 sm:flex-row sm:items-center sm:justify-between ${
                    isEditing ? "bg-emerald-500/10" : ""
                  } ${confirming ? "bg-red-500/10" : ""}`}
                >
                  <div className="min-w-0">
                    <p className="font-medium text-white">{r.description}</p>
                    <p className="text-xs text-zinc-400">
                      {r.category || "Sin categoría"} · {quincenaLabel(r.quincena)}
                    </p>
                    <p className="mt-1 text-base font-semibold text-emerald-300 sm:hidden">
                      {money(r.amount)}
                    </p>
                  </div>
                  {confirming ? (
                    <div className="grid grid-cols-2 gap-2 sm:flex sm:shrink-0 sm:items-center">
                      <p className="col-span-2 text-sm text-red-100 sm:hidden">
                        ¿Eliminar este gasto?
                      </p>
                      <button
                        type="button"
                        disabled={busy}
                        onClick={() => {
                          tapFeedback();
                          setConfirmKey(null);
                        }}
                        className={`${pressable} min-h-11 rounded-full border border-white/15 px-4 text-sm font-semibold text-zinc-100 hover:bg-white/10 disabled:opacity-50`}
                      >
                        No
                      </button>
                      <button
                        type="button"
                        disabled={busy}
                        onClick={() => void removeGasto(r)}
                        className={`${pressable} min-h-11 rounded-full bg-red-500 px-4 text-sm font-semibold text-white hover:bg-red-400 disabled:opacity-50`}
                      >
                        {busy ? "Eliminando…" : "Sí, borrar"}
                      </button>
                    </div>
                  ) : (
                  <div className="grid grid-cols-2 gap-2 sm:flex sm:shrink-0 sm:items-center sm:gap-2">
                    <p className="hidden text-sm font-semibold text-emerald-300 sm:block sm:pr-2">
                      {money(r.amount)}
                    </p>
                    <button
                      type="button"
                      disabled={saving || busy}
                      onClick={() => startEdit(r)}
                      className={`${pressable} min-h-11 rounded-full border border-white/15 px-4 text-sm font-semibold text-zinc-100 hover:bg-white/10 disabled:opacity-50`}
                    >
                      {isEditing ? "Editando" : "Editar"}
                    </button>
                    <button
                      type="button"
                      disabled={saving || busy}
                      onClick={() => {
                        tapFeedback();
                        setConfirmKey(key);
                      }}
                      className={`${pressable} min-h-11 rounded-full border border-red-400/30 px-4 text-sm font-semibold text-red-200 hover:bg-red-500/15 disabled:opacity-50`}
                    >
                      Eliminar
                    </button>
                  </div>
                  )}
                </li>
                );
              })}
            </ul>
          )}
        </section>
      )}
    </div>
  );
}
