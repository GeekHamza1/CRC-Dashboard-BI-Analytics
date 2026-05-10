"use client";

import { useTheme } from "next-themes";
import { useEffect, useState } from "react";

export function ThemeToggle() {
  const { resolvedTheme, setTheme } = useTheme();
  const [mounted, setMounted] = useState(false);

  useEffect(() => setMounted(true), []);

  const next = mounted && resolvedTheme === "dark" ? "light" : "dark";

  return (
    <button
      type="button"
      onClick={() => setTheme(next)}
      className="rounded-full px-4 py-2 text-sm font-medium border border-slate-200/80 dark:border-slate-600 bg-white/60 dark:bg-slate-900/70 hover:bg-white dark:hover:bg-slate-900 transition shadow-sm backdrop-blur"
      aria-label="Basculer le thème"
    >
      {!mounted ? "…" : resolvedTheme === "dark" ? "Mode clair" : "Mode sombre"}
    </button>
  );
}
