import type PptxGenJS from "pptxgenjs";

import { reportTimestamp } from "../export-branding";

import { PPTX_CRC } from "./pptx-theme";

export function addSlideHeader(
  slide: PptxGenJS.Slide,
  title: string,
  subtitle: string,
  logoDataUrl: string | null,
) {
  slide.background = { color: PPTX_CRC.bg };

  slide.addShape("rect", {
    x: 0.2,
    y: 0.1,
    w: 13.0,
    h: 0.72,
    fill: { color: PPTX_CRC.header },
    line: { color: PPTX_CRC.header, pt: 0 },
    shadow: {
      type: "outer",
      blur: 3,
      offset: 1.5,
      angle: 90,
      color: "000000",
      opacity: 0.12,
    },
  });

  if (logoDataUrl) {
    try {
      slide.addImage({
        data: logoDataUrl,
        x: 0.3,
        y: 0.12,
        w: 0.78,
        h: 0.6,
      });
    } catch {
      /** logo optional */
    }
  }

  slide.addText(title, {
    x: 1.2,
    y: 0.22,
    w: 8.8,
    h: 0.2,
    fontSize: 16,
    bold: true,
    color: PPTX_CRC.white,
  });

  slide.addText(subtitle, {
    x: 1.2,
    y: 0.47,
    w: 10.4,
    h: 0.2,
    fontSize: 9,
    color: PPTX_CRC.cream,
  });

  slide.addText(`Généré le ${reportTimestamp()}`, {
    x: 9.8,
    y: 0.46,
    w: 3.2,
    h: 0.2,
    fontSize: 8,
    color: PPTX_CRC.bg,
    align: "right",
  });
}
