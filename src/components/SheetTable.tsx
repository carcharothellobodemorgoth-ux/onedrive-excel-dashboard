"use client";

export function SheetTable({
  headers,
  rows,
}: {
  headers: string[];
  rows: (string | number | boolean | null)[][];
}) {
  if (headers.length === 0) {
    return (
      <p className="rounded-2xl border border-dashed border-white/15 p-8 text-center text-zinc-400">
        Esta hoja no tiene datos en el rango usado.
      </p>
    );
  }

  return (
    <div className="overflow-hidden rounded-2xl border border-white/10 bg-white/5">
      <div className="max-h-[28rem] overflow-auto">
        <table className="min-w-full border-collapse text-left text-sm">
          <thead className="sticky top-0 bg-zinc-900/95 backdrop-blur">
            <tr>
              {headers.map((h) => (
                <th
                  key={h}
                  className="whitespace-nowrap border-b border-white/10 px-4 py-3 font-semibold text-zinc-200"
                >
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, i) => (
              <tr
                key={i}
                className="border-b border-white/5 odd:bg-white/[0.02] hover:bg-emerald-500/5"
              >
                {headers.map((_, j) => {
                  const cell = row[j];
                  return (
                    <td
                      key={j}
                      className="whitespace-nowrap px-4 py-2.5 text-zinc-300"
                    >
                      {cell === null || cell === undefined ? "—" : String(cell)}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
