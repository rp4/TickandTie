import { COLORS, DocumentResult, ExtractedValue, FieldDefinition, ReconcileResult } from './types';

// Access globals loaded via CDN
declare global {
  interface Window {
    pdfjsLib: any;
    XLSX: any;
    JSZip: any;
    marked: any;
  }
}

export const sanitizeKey = (name: string): string => {
  return name.replace(/[^a-zA-Z0-9]/g, '_').toLowerCase();
};

export const getNextColor = (index: number): string => {
  return COLORS[index % COLORS.length];
};

export const fileToBase64 = (file: File): Promise<string> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => {
      const result = reader.result as string;
      // Remove data: type prefix for API calls if needed, but usually keep it for previews
      resolve(result);
    };
    reader.onerror = (error) => reject(error);
  });
};

/**
 * Shared helper to load a File (Image or PDF) into a Canvas.
 * PDF rendering uses scale. Images use natural size.
 */
const fileToCanvas = async (file: File, pdfScale: number = 1.5): Promise<HTMLCanvasElement> => {
  if (file.type.startsWith('image/')) {
    const base64 = await fileToBase64(file);
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        canvas.width = img.naturalWidth;
        canvas.height = img.naturalHeight;
        const ctx = canvas.getContext('2d');
        if (!ctx) {
          reject(new Error("Could not get 2d context"));
          return;
        }
        ctx.drawImage(img, 0, 0);
        resolve(canvas);
      };
      img.onerror = reject;
      img.src = base64;
    });
  }

  if (file.type === 'application/pdf') {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await window.pdfjsLib.getDocument(arrayBuffer).promise;
    const page = await pdf.getPage(1);
    
    const viewport = page.getViewport({ scale: pdfScale });
    
    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d');
    canvas.height = viewport.height;
    canvas.width = viewport.width;

    if (!context) throw new Error("Canvas context not available");

    await page.render({
      canvasContext: context,
      viewport: viewport
    }).promise;

    return canvas;
  }

  throw new Error('Unsupported file type for rendering');
};

/**
 * Renders the first page of a PDF or loads an image to a Base64 Image URL.
 */
export const renderPdfToImage = async (file: File): Promise<string> => {
  try {
    const canvas = await fileToCanvas(file, 1.5); // 1.5 scale is good for thumbnails
    return canvas.toDataURL('image/jpeg');
  } catch (e) {
    console.error("Preview generation failed", e);
    throw e;
  }
};

/**
 * Generates an annotated image (Blob and URL) with bounding boxes drawn on it.
 */
export const generateAnnotatedImage = async (
  file: File,
  data: Record<string, ExtractedValue>,
  fields: FieldDefinition[]
): Promise<{ blob: Blob; url: string }> => {
  // Use higher scale for annotated export (PDFs)
  const canvas = await fileToCanvas(file, 2.0);
  const ctx = canvas.getContext('2d');
  if (!ctx) throw new Error("No canvas context");

  fields.forEach(field => {
    const extracted = data[field.key];
    if (extracted && extracted.box_2d) {
      const [ymin, xmin, ymax, xmax] = extracted.box_2d;
      
      // Convert normalized (0-1000) coords to canvas coords
      const x = (xmin / 1000) * canvas.width;
      const y = (ymin / 1000) * canvas.height;
      const w = ((xmax - xmin) / 1000) * canvas.width;
      const h = ((ymax - ymin) / 1000) * canvas.height;

      // 1. Draw Box
      ctx.lineWidth = Math.max(2, canvas.width * 0.003); // Responsive line width
      ctx.strokeStyle = field.color;
      ctx.strokeRect(x, y, w, h);

      // 2. Draw Fill (Transparent)
      ctx.fillStyle = field.color + '20'; // Hex alpha ~12%
      ctx.fillRect(x, y, w, h);

      // 3. Draw Label Tag
      const fontSize = Math.max(12, canvas.width * 0.015);
      ctx.font = `bold ${fontSize}px sans-serif`;
      const text = field.name;
      const textMetrics = ctx.measureText(text);
      const padding = 4;
      const textW = textMetrics.width + padding * 2;
      const textH = fontSize + padding * 2;
      
      // Ensure label stays within canvas bounds (basic check for top edge)
      const labelY = y - textH > 0 ? y - textH : y;
      
      ctx.fillStyle = field.color;
      ctx.fillRect(x, labelY, textW, textH);
      
      ctx.fillStyle = 'white';
      ctx.textBaseline = 'top';
      ctx.fillText(text, x + padding, labelY + padding);
    }
  });

  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) {
        resolve({
          blob,
          url: URL.createObjectURL(blob)
        });
      } else {
        reject(new Error("Failed to create blob from canvas"));
      }
    }, 'image/jpeg', 0.85); // JPEG quality 0.85
  });
};

export const parseExcelFile = async (file: File): Promise<any[]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = window.XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = window.XLSX.utils.sheet_to_json(worksheet);
        resolve(json);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
};

// Helper to parse Markdown tables into array of arrays
const parseMarkdownTable = (markdown: string): any[][] | null => {
  const lines = markdown.split('\n');
  const data: any[][] = [];
  let tableStarted = false;

  for (let line of lines) {
    line = line.trim();
    if (line.startsWith('|') && line.endsWith('|')) {
      // Check if it's a separator line (e.g. |---|---|)
      if (line.replace(/\|/g, '').trim().match(/^-+$/)) {
        continue; 
      }
      
      const cells = line.split('|').slice(1, -1).map(c => c.trim());
      data.push(cells);
      tableStarted = true;
    } else if (tableStarted && line === '') {
      // Stop if we hit an empty line after finding a table
      break;
    }
  }
  
  return data.length > 0 ? data : null;
}

export const exportToZip = async (
  fields: FieldDefinition[], 
  documents: DocumentResult[],
  reconcileResult?: ReconcileResult | null
) => {
  if (!window.XLSX || !window.JSZip) {
    console.error("Required libraries (SheetJS or JSZip) not loaded");
    alert("Export libraries are still loading. Please try again in a moment.");
    return;
  }

  const zip = new window.JSZip();
  const folderName = "annotated_documents"; // Renamed to reflect content
  const docsFolder = zip.folder(folderName);

  // Handle duplicate filenames
  const filenameMap = new Map<string, string>();
  const usedFilenames = new Set<string>();

  documents.forEach(doc => {
    let name = doc.fileName;
    // If we have an annotated blob, we will save as .jpg regardless of input
    if (doc.annotatedBlob) {
      const lastDotIndex = name.lastIndexOf(".");
      const base = lastDotIndex !== -1 ? name.substring(0, lastDotIndex) : name;
      name = `${base}_annotated.jpg`;
    }

    const lastDotIndex = name.lastIndexOf(".");
    let base = lastDotIndex !== -1 ? name.substring(0, lastDotIndex) : name;
    let ext = lastDotIndex !== -1 ? name.substring(lastDotIndex) : "";
    
    // Ensure unique filename
    let counter = 1;
    let uniqueName = name;
    while (usedFilenames.has(uniqueName)) {
      uniqueName = `${base}_${counter}${ext}`;
      counter++;
    }
    usedFilenames.add(uniqueName);
    filenameMap.set(doc.id, uniqueName);

    // Add file to zip folder
    if (docsFolder) {
      if (doc.annotatedBlob) {
        docsFolder.file(uniqueName, doc.annotatedBlob);
      } else {
        // Fallback to original file if no annotation available (e.g. failed extraction)
        docsFolder.file(uniqueName, doc.file);
      }
    }
  });

  const workbook = window.XLSX.utils.book_new();

  // --- Sheet 1: Audit Data (Tick) ---
  const headers = [
    "File Name", 
    "Status", 
    ...fields.map(f => f.name), 
    ...fields.map(f => `${f.name} (Coords)`)
  ];

  const dataRows = documents.map(doc => {
    const row: Record<string, any> = {};
    // Use original filename for display row, but link will point to annotated file
    row["File Name"] = doc.fileName; 
    row["Status"] = doc.status;
    
    fields.forEach(field => {
      const extracted = doc.data[field.key];
      row[field.name] = extracted?.value || '';
      row[`${field.name} (Coords)`] = extracted?.box_2d ? JSON.stringify(extracted.box_2d) : '';
    });
    return row;
  });

  const worksheet = window.XLSX.utils.json_to_sheet(dataRows, { header: headers });

  // Add Hyperlinks
  const range = window.XLSX.utils.decode_range(worksheet['!ref']);
  for (let R = range.s.r + 1; R <= range.e.r; ++R) {
    const docIndex = R - 1; 
    if (docIndex < documents.length) {
      const doc = documents[docIndex];
      const uniqueName = filenameMap.get(doc.id);
      
      if (uniqueName) {
        const cellRef = window.XLSX.utils.encode_cell({ c: 0, r: R }); // Column 0
        if (!worksheet[cellRef]) {
            worksheet[cellRef] = { t: 's', v: doc.fileName };
        }
        worksheet[cellRef].l = { Target: `${folderName}/${uniqueName}` };
      }
    }
  }
  window.XLSX.utils.book_append_sheet(workbook, worksheet, "Source Data (Tick)");

  // --- Sheet 2+: Reconcile Data ---
  if (reconcileResult) {
    const parsedTable = parseMarkdownTable(reconcileResult.report);
    if (parsedTable) {
      const sheet = window.XLSX.utils.aoa_to_sheet(parsedTable);
      window.XLSX.utils.book_append_sheet(workbook, sheet, "Joined Data Details");
    }

    const reportSheet = window.XLSX.utils.aoa_to_sheet([
      ["Analysis Report"],
      [reconcileResult.report]
    ]);
    if (!reportSheet['!cols']) reportSheet['!cols'] = [];
    reportSheet['!cols'][0] = { wch: 100 }; 
    window.XLSX.utils.book_append_sheet(workbook, reportSheet, "Analysis Report");

    const codeSheet = window.XLSX.utils.aoa_to_sheet([
      ["Executed Python Code"],
      [reconcileResult.code]
    ]);
    if (!codeSheet['!cols']) codeSheet['!cols'] = [];
    codeSheet['!cols'][0] = { wch: 100 };
    window.XLSX.utils.book_append_sheet(workbook, codeSheet, "Python Code");
  }

  const excelBuffer = window.XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  zip.file("Audit_Report.xlsx", excelBuffer);

  try {
    const content = await zip.generateAsync({ type: "blob" });
    const url = URL.createObjectURL(content);
    const a = document.createElement("a");
    a.href = url;
    a.download = "TickAndTie_Audit_Annotated.zip";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } catch (error) {
    console.error("Failed to generate zip", error);
    alert("Failed to generate ZIP file.");
  }
};
