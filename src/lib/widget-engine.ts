export type WidgetType = "chart" | "pivot" | "kpi" | "table" | "timeline";

export interface DashboardWidget {
  id: string;
  type: WidgetType;
  title: string;
  x: number;
  y: number;
  w: number;
  h: number;
  visible: boolean;
  config: Record<string, unknown>;
}

export interface DashboardLayoutConfig {
  widgets: DashboardWidget[];
  title: string;
  subtitle?: string;
  branding?: {
    primaryColor?: string;
    logoPath?: string;
  };
}

const KEY = "bi:dashboard-layout";

export function loadDashboardLayout(): DashboardLayoutConfig {
  if (typeof window === "undefined")
    return { widgets: [], title: "Dashboard Builder" };
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return { widgets: [], title: "Dashboard Builder" };
    return JSON.parse(raw) as DashboardLayoutConfig;
  } catch {
    return { widgets: [], title: "Dashboard Builder" };
  }
}

export function saveDashboardLayout(config: DashboardLayoutConfig) {
  if (typeof window === "undefined") return;
  localStorage.setItem(KEY, JSON.stringify(config));
}
