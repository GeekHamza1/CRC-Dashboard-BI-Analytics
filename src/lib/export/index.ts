/** Native CRC PowerPoint export (editable objects — no dashboard screenshots). */

export type { CrcPptxDashboardSnapshot, CrcPptxExportOptions, CrcPptxWidgetDescriptor } from "./crc-pptx-types";
export { buildAndSaveCrcPptx } from "./pptx-builder";
export { buildCrcPptxWidgetPlan } from "./pptx-widgets";
export {
  buildCrcPptxWidgetRegistry,
  buildRegionCardWidgetSpecs,
  type CrcPptxWidgetRegistry,
  type CrcRegionCardWidgetSpec,
} from "./widget-registry";
