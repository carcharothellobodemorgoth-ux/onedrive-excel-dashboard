/** Client fetch that bypasses browser HTTP cache (critical on mobile). */
export function apiFetch(input: string, init?: RequestInit): Promise<Response> {
  const url = new URL(
    input,
    typeof window !== "undefined" ? window.location.origin : "http://localhost",
  );
  url.searchParams.set("_ts", String(Date.now()));
  return fetch(url.pathname + url.search, {
    ...init,
    cache: "no-store",
    headers: {
      "Cache-Control": "no-cache",
      Pragma: "no-cache",
      ...(init?.headers ?? {}),
    },
  });
}
