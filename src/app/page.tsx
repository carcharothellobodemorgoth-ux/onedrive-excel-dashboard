import { auth, signIn } from "@/auth";
import { redirect } from "next/navigation";

export default async function HomePage() {
  const session = await auth();
  if (session?.user) {
    redirect("/dashboard");
  }

  return (
    <main className="relative flex min-h-screen flex-1 flex-col overflow-hidden">
      <div
        aria-hidden
        className="pointer-events-none absolute inset-0 bg-[radial-gradient(ellipse_at_top,_#064e3b_0%,_#09090b_55%,_#000_100%)]"
      />
      <div
        aria-hidden
        className="pointer-events-none absolute inset-0 opacity-30 [background-image:linear-gradient(rgba(255,255,255,0.04)_1px,transparent_1px),linear-gradient(90deg,rgba(255,255,255,0.04)_1px,transparent_1px)] [background-size:48px_48px]"
      />

      <div className="relative z-10 mx-auto flex w-full max-w-3xl flex-1 flex-col justify-center px-6 py-16">
        <p className="text-sm font-medium uppercase tracking-[0.2em] text-emerald-300/90">
          OneDrive Excel
        </p>
        <h1 className="mt-4 text-5xl font-semibold tracking-tight text-white sm:text-6xl">
          Dashboard
        </h1>
        <p className="mt-4 max-w-xl text-lg text-zinc-300">
          Conectá tu cuenta Microsoft, leé la planilla de OneDrive y visualizá
          KPIs, tablas y gráficos en vivo.
        </p>

        <form
          className="mt-10"
          action={async () => {
            "use server";
            await signIn("microsoft-entra-id", { redirectTo: "/dashboard" });
          }}
        >
          <button
            type="submit"
            className="rounded-full bg-white px-6 py-3 text-sm font-semibold text-zinc-950 shadow-lg shadow-emerald-900/30 transition hover:bg-emerald-50"
          >
            Iniciar sesión con Microsoft / OneDrive
          </button>
        </form>

        <p className="mt-6 text-xs text-zinc-500">
          Lectura + escritura en OneDrive (hoja Gastos Varios). App Registration
          Azure con cuentas personales y permisos Files.Read / Files.ReadWrite.
        </p>
      </div>
    </main>
  );
}
