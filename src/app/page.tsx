"use client";

import dynamic from "next/dynamic";

const CrcDashboard = dynamic(
  () => import("@/components/crc/crc-dashboard"),
  {
    ssr: false,
    loading: () => (
      <div className="min-h-[60vh] flex items-center justify-center text-sm text-slate-500 animate-pulse">
        Chargement du dashboard CRC…
      </div>
    ),
  }
);

export default function HomePage() {
  return <CrcDashboard />;
}