const LOGO_CANDIDATE_PATHS = [
  "/srm-logo.png",
  // "/srm-logo.jpg",
  // "/logo-srm.png",
  // "/logo.png",
] as const;

async function blobToDataUrl(blob: Blob): Promise<string> {
  return await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result ?? ""));
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

export async function loadBrandLogoDataUrl(): Promise<string | null> {
  for (const path of LOGO_CANDIDATE_PATHS) {
    try {
      const res = await fetch(path, { cache: "no-store" });
      if (!res.ok) continue;
      const blob = await res.blob();
      return await blobToDataUrl(blob);
    } catch {
      /** skip */
    }
  }
  return null;
}

/** Uploaded logo from Configuration du rapport, sinon logo public `/srm-logo.png` si présent */
export async function resolveReportLogo(base64Upload: string | null | undefined): Promise<string | null> {
  const u = typeof base64Upload === "string" ? base64Upload.trim() : "";
  if (u.startsWith("data:")) return u;
  if (u) return null;
  return await loadBrandLogoDataUrl();
}

export function reportTimestamp() {
  return new Date().toLocaleString("fr-FR");
}
