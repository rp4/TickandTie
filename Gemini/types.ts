export interface FieldDefinition {
  id: string;
  name: string;
  key: string; // Sanitized name for JSON key
  color: string; // Hex color for bounding box
}

// ymin, xmin, ymax, xmax (0-1000 scale)
export type BoundingBox = [number, number, number, number];

export interface ExtractedValue {
  value: string | number | null;
  box_2d?: BoundingBox | null;
}

export interface DocumentResult {
  id: string;
  file: File;
  fileName: string;
  fileType: string;
  previewUrl: string | null; // Data URL for image or rendered PDF page
  annotatedUrl?: string; // URL for the image with burnt-in annotations
  annotatedBlob?: Blob; // Blob for the image with burnt-in annotations
  status: 'idle' | 'processing' | 'success' | 'error';
  errorMsg?: string;
  data: Record<string, ExtractedValue>; // Keyed by FieldDefinition.key
}

export interface ReconcileResult {
  report: string;
  code: string;
  joinedData?: any[];
}

export const COLORS = [
  '#ef4444', // red-500
  '#3b82f6', // blue-500
  '#10b981', // emerald-500
  '#f59e0b', // amber-500
  '#8b5cf6', // violet-500
  '#ec4899', // pink-500
  '#06b6d4', // cyan-500
  '#84cc16', // lime-500
  '#f97316', // orange-500
  '#6366f1', // indigo-500
];