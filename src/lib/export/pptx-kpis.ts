import type PptxGenJS from "pptxgenjs";

import type { CrcKpiKey } from "../crc-report-config";
import type { CrcRow } from "../crc-types";

import { KPI_CARD, PPTX_CRC } from "./pptx-theme";

const KPI_LABEL: Record<CrcKpiKey, string> = {
  totalVolume: "Total interactions",
  abandons: "Appels abandonnés",
  informes: "Clients informés",
  tickets: "Tickets transmis",
  teleopsDistinct: "Téléopérateurs actifs",
  pctInformes: "% clients informés",
  pctTickets: "% tickets transmis",
  coverage: "Couverture filtre",
};

export function addKpiCardGrid(
  slide: PptxGenJS.Slide,
  items: { key: CrcKpiKey; value: string }[],
  origin: { x: number; y: number },
  cols = 4,
) {
  const { w, h, radius, titleSize, valueSize, pad } = KPI_CARD;
  let i = 0;
  for (const item of items) {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = origin.x + col * (w + 0.12);
    const y = origin.y + row * (h + 0.14);

    slide.addShape("roundRect", {
      x,
      y,
      w,
      h,
      fill: { color: PPTX_CRC.cardFill },
      line: { color: PPTX_CRC.cardBorder, pt: 0.75 },
      rectRadius: radius,
      shadow: {
        type: "outer",
        blur: 4,
        offset: 1.2,
        angle: 90,
        color: "000000",
        opacity: 0.08,
      },
    });

    slide.addText(KPI_LABEL[item.key], {
      x: x + pad,
      y: y + pad,
      w: w - pad * 2,
      h: 0.28,
      fontSize: titleSize,
      color: PPTX_CRC.slate600,
      bold: true,
    });

    slide.addText(item.value, {
      x: x + pad,
      y: y + 0.32,
      w: w - pad * 2,
      h: 0.62,
      fontSize: valueSize,
      bold: true,
      color: PPTX_CRC.navy,
      valign: "middle",
    });

    i += 1;
  }
}

export function buildKpiItems(
  rows: CrcRow[],
  kpisConfig: Record<CrcKpiKey, boolean>,
  k: {
    totalAppels: number;
    réclamations: number;
    appelsAbandonnés: number;
    clientsInformés: number;
    ticketsTransmis: number;
  },
): { key: CrcKpiKey; value: string }[] {
  const vol = rows.length;
  const distinctOps = new Set(rows.map((r) => r.téléopérateur)).size;
  const pct = (n: number) => (vol ? `${((n / vol) * 100).toFixed(1)} %` : "—");

  const map: Partial<Record<CrcKpiKey, string>> = {
    totalVolume: `${k.totalAppels.toLocaleString("fr-FR")}`,
    abandons: `${k.appelsAbandonnés.toLocaleString("fr-FR")}`,
    informes: `${k.clientsInformés.toLocaleString("fr-FR")}`,
    tickets: `${k.ticketsTransmis.toLocaleString("fr-FR")}`,
    teleopsDistinct: `${distinctOps}`,
    pctInformes: pct(k.clientsInformés),
    pctTickets: pct(k.ticketsTransmis),
    coverage: `${vol.toLocaleString("fr-FR")} lignes`,
  };

  return (Object.keys(map) as CrcKpiKey[])
    .filter((key) => kpisConfig[key])
    .map((key) => ({ key, value: map[key] ?? "—" }));
}
