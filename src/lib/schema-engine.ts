import type { DataRecord, DatasetSchema } from "./dataset-schema";
import { discoverDatasetSchema, loadSchemaLocal, saveSchemaLocal } from "./dataset-schema";

export function buildSchema(datasetId: string, rows: DataRecord[]) {
  const schema = discoverDatasetSchema(rows);
  saveSchemaLocal(datasetId, schema);
  return schema;
}

export function getSchema(datasetId: string): DatasetSchema | null {
  return loadSchemaLocal(datasetId);
}
