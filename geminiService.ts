import { GoogleGenAI, Type } from "@google/genai";
import { DocumentResult, FieldDefinition, ExtractedValue, ReconcileResult } from './types';
import { fileToBase64 } from './utils';

export const processDocument = async (
  doc: DocumentResult,
  fields: FieldDefinition[]
): Promise<Record<string, ExtractedValue>> => {
  
  if (!process.env.API_KEY) {
    throw new Error("API Key is missing");
  }

  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
  
  // 1. Construct Dynamic Schema
  const properties: Record<string, any> = {};
  
  fields.forEach(field => {
    properties[field.key] = {
      type: Type.OBJECT,
      properties: {
        value: { type: Type.STRING, description: `The extracted value for ${field.name}` },
        box_2d: { 
          type: Type.ARRAY, 
          items: { type: Type.INTEGER },
          description: "The bounding box of the value in [ymin, xmin, ymax, xmax] format (normalized 0-1000)."
        }
      },
      description: `Extraction result for ${field.name}`
    };
  });

  // 2. Prepare Prompt
  const base64DataUrl = await fileToBase64(doc.file);
  const base64Data = base64DataUrl.split(',')[1]; // Remove "data:image/xyz;base64," prefix

  const prompt = `
    You are an expert internal auditor. 
    Analyze the provided document.
    Extract the following fields: ${fields.map(f => f.name).join(', ')}.
    For each field, find the text value and the 2D bounding box coordinates.
    If a field is not found, return null for value.
    The bounding box should be normalized to a 0-1000 scale in [ymin, xmin, ymax, xmax] order.
  `;

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: {
        parts: [
          { inlineData: { mimeType: doc.file.type, data: base64Data } },
          { text: prompt }
        ]
      },
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: properties,
        }
      }
    });

    const text = response.text;
    if (!text) throw new Error("No response from Gemini");

    const json = JSON.parse(text);
    return json;

  } catch (error) {
    console.error("Gemini Error:", error);
    throw error;
  }
};

export const reconcileData = async (
  documents: DocumentResult[],
  fields: FieldDefinition[],
  referenceData: any[],
  userInstructions: string
): Promise<ReconcileResult> => {
  if (!process.env.API_KEY) {
    throw new Error("API Key is missing");
  }

  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });

  // Prepare data for the context
  const extractedData = documents
    .filter(d => d.status === 'success')
    .map(d => {
      const row: any = { fileName: d.fileName };
      fields.forEach(f => {
        row[f.name] = d.data[f.key]?.value || null;
      });
      return row;
    });

  const prompt = `
    You are an expert data analyst performing a 'Tick & Tie' audit procedure.
    
    You have two datasets available as JSON variables in your context.
    
    Dataset 1: Extracted Data (from source documents)
    ${JSON.stringify(extractedData, null, 2)}
    
    Dataset 2: Reference Data (from user upload)
    ${JSON.stringify(referenceData, null, 2)}

    User Instructions:
    "${userInstructions}"

    Goal:
    Write and execute Python code to perform the analysis requested by the user.
    Load the data into pandas DataFrames (create the dataframes directly from the provided JSON data).
    Perform the comparison/join/reconciliation.
    
    IMPORTANT OUTPUT INSTRUCTION:
    1. Print the results in a clear, readable format. 
    2. If you are generating a table of data (e.g. matched invoices, discrepancies), ALWAYS print it as a Markdown Table.
    3. Provide a summary text explanation of what was done and what was found.
  `;

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: { text: prompt },
      config: {
        tools: [{ codeExecution: {} }]
      }
    });

    const report = response.text || "Analysis completed, but no text explanation was returned.";
    
    // Extract executable code from candidates
    const parts = response.candidates?.[0]?.content?.parts || [];
    const codeBlocks = parts
      .filter((part: any) => part.executableCode)
      .map((part: any) => part.executableCode.code);
    
    const code = codeBlocks.join('\n\n# ---------------------------------------------------------\n# Next Code Block\n# ---------------------------------------------------------\n\n');

    return { report, code };

  } catch (error) {
    console.error("Reconciliation Error:", error);
    throw error;
  }
};