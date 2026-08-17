import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import { TapFix } from "@/components/TapFix";
import "./globals.css";

const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

export const metadata: Metadata = {
  title: "OneDrive Excel Dashboard",
  description: "Dashboard en vivo conectado a Excel en OneDrive personal",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html
      lang="es"
      className={`${geistSans.variable} ${geistMono.variable} h-full antialiased`}
    >
      <body className="min-h-full flex flex-col bg-zinc-950 text-zinc-100">
        <TapFix />
        {children}
      </body>
    </html>
  );
}
