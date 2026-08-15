import { auth } from "@/auth";
import { DashboardClient } from "@/components/DashboardClient";
import { redirect } from "next/navigation";

export default async function DashboardPage() {
  const session = await auth();
  if (!session?.user) {
    redirect("/");
  }

  return (
    <main className="min-h-screen bg-zinc-950 text-white">
      <div
        aria-hidden
        className="pointer-events-none fixed inset-0 bg-[radial-gradient(ellipse_at_top_right,_#064e3b55_0%,_transparent_50%)]"
      />
      <div className="relative z-10">
        <DashboardClient userName={session.user.name ?? session.user.email} />
      </div>
    </main>
  );
}
