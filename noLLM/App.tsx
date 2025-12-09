import React, { useState, useRef } from 'react';
import { Upload, Plus, Trash2, FileText, Play, Download, Eye, CheckCircle, AlertCircle, Loader2, RefreshCw, FileSpreadsheet, Square } from 'lucide-react';
import { FieldDefinition, DocumentResult, ReconcileResult, COLORS } from './types';
import { sanitizeKey, getNextColor, renderPdfToImage, exportToZip, parseExcelFile, generateAnnotatedImage } from './utils';
import { processDocument, reconcileData } from './geminiService';
import DocumentViewer from './components/DocumentViewer';
import { v4 as uuidv4 } from 'uuid';

type Tab = 'tick' | 'tie';

const App: React.FC = () => {
  // --- State ---
  const [activeTab, setActiveTab] = useState<Tab>('tick');
  const [fields, setFields] = useState<FieldDefinition[]>([
    { id: '1', name: 'Invoice Number', key: 'invoice_number', color: COLORS[0] },
    { id: '2', name: 'Invoice Date', key: 'invoice_date', color: COLORS[1] },
    { id: '3', name: 'Total Amount', key: 'total_amount', color: COLORS[2] }
  ]);
  const [newFieldName, setNewFieldName] = useState('');
  const [documents, setDocuments] = useState<DocumentResult[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [selectedDocId, setSelectedDocId] = useState<string | null>(null);
  
  // Ref for aborting batch process
  const shouldStopRef = useRef(false);

  // Reconcile / Tie State
  const [referenceData, setReferenceData] = useState<any[]>([]);
  const [referenceFileName, setReferenceFileName] = useState<string | null>(null);
  const [reconcilePrompt, setReconcilePrompt] = useState("Compare the datasets and highlight any differences");
  const [isReconciling, setIsReconciling] = useState(false);
  const [reconcileResult, setReconcileResult] = useState<ReconcileResult | null>(null);

  // --- Handlers ---

  const addField = () => {
    if (!newFieldName.trim()) return;
    setFields(prev => [
      ...prev,
      {
        id: uuidv4(),
        name: newFieldName,
        key: sanitizeKey(newFieldName),
        color: getNextColor(prev.length)
      }
    ]);
    setNewFieldName('');
  };

  const removeField = (id: string) => {
    setFields(prev => prev.filter(f => f.id !== id));
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const newDocs: DocumentResult[] = [];
      
      // Process file previews immediately
      const files = Array.from(e.target.files) as File[];
      for (const file of files) {
        try {
          // Generate preview (Image or PDF page 1)
          const previewUrl = await renderPdfToImage(file);
          
          newDocs.push({
            id: uuidv4(),
            file,
            fileName: file.name,
            fileType: file.type,
            previewUrl,
            status: 'idle',
            data: {}
          });
        } catch (err) {
          console.error("Failed to load preview for", file.name, err);
          // Add anyway but maybe with error state or generic icon
           newDocs.push({
            id: uuidv4(),
            file,
            fileName: file.name,
            fileType: file.type,
            previewUrl: null, // Fallback
            status: 'error',
            errorMsg: 'Could not generate preview',
            data: {}
          });
        }
      }
      setDocuments(prev => [...prev, ...newDocs]);
    }
  };

  const handleReferenceUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      try {
        const data = await parseExcelFile(file);
        setReferenceData(data);
        setReferenceFileName(file.name);
      } catch (error) {
        console.error("Error parsing reference file", error);
        alert("Failed to parse the reference file. Please ensure it is a valid Excel or CSV.");
      }
    }
  };

  const handleStop = () => {
    shouldStopRef.current = true;
  };

  const runExtraction = async () => {
    if (documents.length === 0) return;
    setIsProcessing(true);
    shouldStopRef.current = false;

    // Clone docs to update state
    const docsToProcess = [...documents];

    // Process sequentially
    for (let i = 0; i < docsToProcess.length; i++) {
      // Check for stop signal at start of iteration
      if (shouldStopRef.current) {
        break;
      }

      const doc = docsToProcess[i];
      // Skip if success or already processing (to allow parallel single run + batch run mixing safely)
      if (doc.status === 'success' || doc.status === 'processing') continue;

      // Update status to processing
      setDocuments(prev => prev.map(d => d.id === doc.id ? { ...d, status: 'processing' } : d));

      try {
        const result = await processDocument(doc, fields);
        
        // Generate Annotated Image immediately for Viewer and Export
        let annotatedBlob: Blob | undefined;
        let annotatedUrl: string | undefined;

        try {
          const annotation = await generateAnnotatedImage(doc.file, result, fields);
          annotatedBlob = annotation.blob;
          annotatedUrl = annotation.url;
        } catch (annotErr) {
          console.error("Failed to generate annotation", annotErr);
        }
        
        setDocuments(prev => prev.map(d => 
          d.id === doc.id ? { 
            ...d, 
            status: 'success', 
            data: result,
            annotatedBlob,
            annotatedUrl
          } : d
        ));
      } catch (err) {
        setDocuments(prev => prev.map(d => 
          d.id === doc.id ? { ...d, status: 'error', errorMsg: 'Extraction failed' } : d
        ));
      }
    }

    setIsProcessing(false);
  };

  const runSingleExtraction = async (docId: string) => {
    const docIndex = documents.findIndex(d => d.id === docId);
    if (docIndex === -1) return;
    const doc = documents[docIndex];
    
    // Prevent double processing
    if (doc.status === 'processing') return;

    // Set status to processing
    setDocuments(prev => prev.map(d => d.id === docId ? { ...d, status: 'processing', errorMsg: undefined } : d));

    try {
        const result = await processDocument(doc, fields);
        
        let annotatedBlob: Blob | undefined;
        let annotatedUrl: string | undefined;

        try {
          const annotation = await generateAnnotatedImage(doc.file, result, fields);
          annotatedBlob = annotation.blob;
          annotatedUrl = annotation.url;
        } catch (annotErr) {
          console.error("Failed to generate annotation", annotErr);
        }
        
        setDocuments(prev => prev.map(d => 
          d.id === docId ? { 
            ...d, 
            status: 'success', 
            data: result,
            annotatedBlob,
            annotatedUrl
          } : d
        ));
    } catch (err) {
        setDocuments(prev => prev.map(d => 
          d.id === docId ? { ...d, status: 'error', errorMsg: 'Extraction failed' } : d
        ));
    }
  };

  const runReconciliation = async () => {
    const successDocs = documents.filter(d => d.status === 'success');
    if (successDocs.length === 0) {
      alert("Please extract data from documents first (Tick phase).");
      return;
    }
    if (referenceData.length === 0) {
      alert("Please upload a reference dataset.");
      return;
    }

    setIsReconciling(true);
    setReconcileResult(null);

    try {
      const result = await reconcileData(documents, fields, referenceData, reconcilePrompt);
      setReconcileResult(result);
    } catch (error) {
      console.error(error);
      setReconcileResult({ report: "An error occurred during analysis.", code: "" });
    } finally {
      setIsReconciling(false);
    }
  };

  const handleExport = () => {
    exportToZip(fields, documents, reconcileResult);
  };

  // --- Render ---

  const selectedDocument = documents.find(d => d.id === selectedDocId);
  const processedCount = documents.filter(d => d.status === 'success').length;

  return (
    <div className="min-h-screen flex flex-col bg-slate-50">
      {/* Navbar */}
      <header className="bg-white border-b px-8 py-4 flex items-center justify-between sticky top-0 z-20 shadow-sm">
        <div className="flex items-center gap-3">
          <img 
            src="https://storage.googleapis.com/toolbox-478717-storage/branding/worker.png" 
            alt="Logo" 
            className="w-10 h-10 rounded-lg"
          />
          <div>
            <h1 className="text-xl font-bold text-gray-900">Tick & Tie</h1>
            <p className="text-xs text-gray-500">by Auditor in the Loop</p>
          </div>
        </div>

        {/* Tab Switcher */}
        <div className="flex bg-gray-100 p-1 rounded-lg">
          <button
            onClick={() => setActiveTab('tick')}
            className={`px-4 py-1.5 rounded-md text-sm font-medium transition ${activeTab === 'tick' ? 'bg-white text-indigo-600 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}
          >
            1. Extract (Tick)
          </button>
          <button
            onClick={() => setActiveTab('tie')}
            className={`px-4 py-1.5 rounded-md text-sm font-medium transition ${activeTab === 'tie' ? 'bg-white text-indigo-600 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}
          >
            2. Reconcile (Tie)
          </button>
        </div>

        <div className="flex gap-3">
          {/* Actions for Tick Tab */}
          {activeTab === 'tick' && (
            <>
              {isProcessing ? (
                <button 
                  onClick={handleStop}
                  className="flex items-center gap-2 px-5 py-2 rounded-lg font-medium transition shadow-sm text-sm bg-red-600 text-white hover:bg-red-700 shadow-red-200"
                >
                  <Square size={16} fill="currentColor" />
                  Stop Extraction
                </button>
              ) : (
                <button 
                  onClick={runExtraction}
                  disabled={documents.length === 0}
                  className={`flex items-center gap-2 px-5 py-2 rounded-lg font-medium transition shadow-sm text-sm
                    ${documents.length === 0 
                      ? 'bg-gray-100 text-gray-400 cursor-not-allowed' 
                      : 'bg-indigo-600 text-white hover:bg-indigo-700 shadow-indigo-200'}`}
                >
                  <Play size={16} fill="currentColor" />
                  Run Extraction
                </button>
              )}
              
              <button 
                onClick={handleExport}
                disabled={documents.length === 0}
                className="flex items-center gap-2 px-4 py-2 border border-gray-300 text-gray-700 rounded-lg hover:bg-gray-50 font-medium transition disabled:opacity-50 text-sm"
              >
                <Download size={16} />
                Export Annotated ZIP
              </button>
            </>
          )}
          
          {/* Actions for Tie Tab - Added Export Button */}
          {activeTab === 'tie' && (
             <button 
                onClick={handleExport}
                disabled={documents.length === 0 && !reconcileResult}
                className="flex items-center gap-2 px-4 py-2 border border-gray-300 text-gray-700 rounded-lg hover:bg-gray-50 font-medium transition disabled:opacity-50 text-sm"
              >
                <Download size={16} />
                Export Annotated ZIP
              </button>
          )}
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 flex overflow-hidden">
        
        {/* Conditional Rendering based on Tab */}
        {activeTab === 'tick' ? (
          <>
            {/* Sidebar: Configuration */}
            <aside className="w-80 bg-white border-r flex flex-col overflow-hidden z-10">
              <div className="p-6 border-b bg-gray-50">
                <h2 className="text-sm font-semibold text-gray-900 uppercase tracking-wider mb-1">Fields to Extract</h2>
                <p className="text-xs text-gray-500">Define the data points for extraction.</p>
              </div>
              
              <div className="p-4 flex-1 overflow-y-auto">
                <div className="space-y-3">
                  {fields.map((field) => (
                    <div key={field.id} className="flex items-center justify-between p-3 bg-white border rounded-lg shadow-sm group hover:border-indigo-300 transition">
                      <div className="flex items-center gap-3">
                        <div className="w-4 h-4 rounded-full shadow-sm border border-black/10" style={{ backgroundColor: field.color }}></div>
                        <span className="font-medium text-gray-700">{field.name}</span>
                      </div>
                      <button 
                        onClick={() => removeField(field.id)}
                        className="text-gray-400 hover:text-red-500 opacity-0 group-hover:opacity-100 transition"
                      >
                        <Trash2 size={16} />
                      </button>
                    </div>
                  ))}
                </div>
              </div>

              <div className="p-4 border-t bg-gray-50">
                <div className="flex gap-2">
                  <input 
                    type="text" 
                    value={newFieldName}
                    onChange={(e) => setNewFieldName(e.target.value)}
                    onKeyDown={(e) => e.key === 'Enter' && addField()}
                    placeholder="e.g. Vendor Name"
                    className="flex-1 px-3 py-2 border rounded-md text-sm focus:ring-2 focus:ring-indigo-500 outline-none"
                  />
                  <button 
                    onClick={addField}
                    className="bg-gray-900 text-white p-2 rounded-md hover:bg-gray-800 transition"
                  >
                    <Plus size={20} />
                  </button>
                </div>
              </div>
            </aside>

            {/* Tick Content Area */}
            <div className="flex-1 flex flex-col overflow-hidden">
              
              {/* Upload Section */}
              <div className="p-8 pb-4">
                <label className="flex flex-col items-center justify-center w-full h-28 border-2 border-dashed border-gray-300 rounded-xl cursor-pointer bg-white hover:bg-gray-50 hover:border-indigo-400 transition group">
                  <div className="flex items-center justify-center gap-4">
                    <div className="bg-indigo-50 p-3 rounded-full group-hover:bg-indigo-100 transition">
                      <Upload className="text-indigo-600" size={24} />
                    </div>
                    <div className="text-left">
                        <p className="text-sm font-semibold text-gray-900">Click to upload source documents</p>
                        <p className="text-xs text-gray-500">PDF, PNG, JPG (Max 10MB per file)</p>
                    </div>
                  </div>
                  <input type="file" className="hidden" multiple onChange={handleFileUpload} accept=".pdf,.jpg,.jpeg,.png" />
                </label>
              </div>

              {/* Results Table */}
              <div className="flex-1 overflow-auto px-8 pb-8">
                <div className="bg-white border rounded-xl shadow-sm overflow-hidden min-h-[400px]">
                  <table className="w-full text-left border-collapse">
                    <thead className="bg-gray-50 text-xs uppercase text-gray-500 font-semibold sticky top-0 z-10 shadow-sm">
                      <tr>
                        <th className="px-6 py-4 border-b w-16">View</th>
                        <th className="px-6 py-4 border-b">Document</th>
                        <th className="px-6 py-4 border-b w-44">Status</th>
                        {fields.map(field => (
                          <th key={field.id} className="px-6 py-4 border-b" style={{ borderTop: `3px solid ${field.color}` }}>
                            {field.name}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100">
                      {documents.length === 0 ? (
                        <tr>
                          <td colSpan={fields.length + 3} className="px-6 py-20 text-center text-gray-400">
                            <div className="flex flex-col items-center">
                              <FileText size={48} className="mb-4 opacity-20" />
                              <p>No documents uploaded yet.</p>
                            </div>
                          </td>
                        </tr>
                      ) : (
                        documents.map((doc) => (
                          <tr key={doc.id} className="hover:bg-gray-50/80 transition group">
                            <td className="px-6 py-4">
                              <button 
                                onClick={() => setSelectedDocId(doc.id)}
                                className="text-gray-400 hover:text-indigo-600 transition"
                              >
                                <Eye size={20} />
                              </button>
                            </td>
                            <td className="px-6 py-4">
                              <div className="flex items-center gap-3">
                                {doc.previewUrl && (
                                  <img src={doc.previewUrl} alt="thumb" className="w-10 h-10 object-cover rounded border" />
                                )}
                                <div>
                                  <div className="font-medium text-gray-900 text-sm truncate max-w-[200px]" title={doc.fileName}>{doc.fileName}</div>
                                  <div className="text-xs text-gray-500">{doc.fileType.split('/')[1].toUpperCase()}</div>
                                </div>
                              </div>
                            </td>
                            <td className="px-6 py-4">
                              <div className="flex items-center justify-between gap-2">
                                <div className="flex-1">
                                  {doc.status === 'processing' && <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-blue-100 text-blue-800"><Loader2 size={12} className="mr-1 animate-spin" /> Processing</span>}
                                  {doc.status === 'success' && <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-800"><CheckCircle size={12} className="mr-1" /> Done</span>}
                                  {doc.status === 'error' && <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-red-100 text-red-800"><AlertCircle size={12} className="mr-1" /> Error</span>}
                                  {doc.status === 'idle' && <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-gray-100 text-gray-800">Ready</span>}
                                </div>
                                {doc.status !== 'processing' && (
                                  <button
                                    onClick={(e) => {
                                        e.stopPropagation();
                                        runSingleExtraction(doc.id);
                                    }}
                                    className="p-1.5 text-gray-400 hover:text-indigo-600 hover:bg-indigo-50 rounded-full transition"
                                    title={doc.status === 'idle' ? "Run Extraction" : "Re-run Extraction"}
                                  >
                                    {doc.status === 'idle' ? <Play size={16} /> : <RefreshCw size={16} />}
                                  </button>
                                )}
                              </div>
                            </td>
                            {fields.map(field => {
                              const val = doc.data[field.key];
                              return (
                                <td key={field.id} className="px-6 py-4 text-sm text-gray-700 font-mono">
                                  {val?.value || '-'}
                                </td>
                              );
                            })}
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </>
        ) : (
          // --- Reconcile (Tie) Tab ---
          <div className="flex-1 flex overflow-hidden">
            <div className="flex-1 flex flex-col p-8 overflow-auto max-w-6xl mx-auto w-full">
              
              <div className="bg-white border rounded-xl shadow-sm p-6 mb-6">
                <h2 className="text-lg font-semibold text-gray-900 mb-4 flex items-center gap-2">
                  <RefreshCw className="text-indigo-600" size={20}/>
                  Reconciliation Setup
                </h2>
                
                <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                   {/* Data Status */}
                   <div className="space-y-4">
                     <div className="bg-gray-50 rounded-lg p-4 border">
                        <h3 className="text-sm font-medium text-gray-700 mb-2">1. Source Data (Extracted)</h3>
                        <div className="flex items-center justify-between">
                          <span className="text-gray-600 text-sm">{processedCount} documents extracted successfully</span>
                          {processedCount > 0 ? <CheckCircle size={18} className="text-green-500" /> : <AlertCircle size={18} className="text-amber-500" />}
                        </div>
                     </div>

                     <div className="bg-gray-50 rounded-lg p-4 border">
                        <h3 className="text-sm font-medium text-gray-700 mb-2">2. Reference Data (Excel/CSV)</h3>
                        {referenceData.length > 0 ? (
                           <div className="flex items-center justify-between">
                             <div className="flex items-center gap-2">
                               <FileSpreadsheet className="text-green-600" size={20} />
                               <div>
                                 <div className="text-sm font-medium text-gray-900">{referenceFileName}</div>
                                 <div className="text-xs text-gray-500">{referenceData.length} rows loaded</div>
                               </div>
                             </div>
                             <button onClick={() => { setReferenceData([]); setReferenceFileName(null); }} className="text-gray-400 hover:text-red-500">
                               <Trash2 size={16} />
                             </button>
                           </div>
                        ) : (
                          <label className="flex flex-col items-center justify-center w-full h-20 border-2 border-dashed border-gray-300 rounded-lg cursor-pointer hover:bg-indigo-50 hover:border-indigo-400 transition">
                             <span className="text-sm text-gray-500 font-medium">Upload .xlsx or .csv</span>
                             <input type="file" className="hidden" accept=".csv, .xlsx, .xls" onChange={handleReferenceUpload} />
                          </label>
                        )}
                     </div>
                   </div>

                   {/* Prompt Section */}
                   <div className="flex flex-col">
                      <h3 className="text-sm font-medium text-gray-700 mb-2">3. Analysis Instructions</h3>
                      <textarea 
                        className="flex-1 border rounded-lg p-3 text-sm focus:ring-2 focus:ring-indigo-500 outline-none resize-none bg-gray-50"
                        placeholder="Describe how you want to compare the data..."
                        value={reconcilePrompt}
                        onChange={(e) => setReconcilePrompt(e.target.value)}
                      />
                   </div>
                </div>

                <div className="mt-6 flex justify-end">
                  <button 
                    onClick={runReconciliation}
                    disabled={isReconciling || processedCount === 0 || referenceData.length === 0}
                    className={`flex items-center gap-2 px-6 py-2.5 rounded-lg font-medium text-white transition shadow-sm
                      ${isReconciling || processedCount === 0 || referenceData.length === 0
                        ? 'bg-gray-300 cursor-not-allowed' 
                        : 'bg-indigo-600 hover:bg-indigo-700 shadow-indigo-200'}`}
                  >
                    {isReconciling ? <Loader2 className="animate-spin" size={18} /> : <Play size={18} fill="currentColor" />}
                    {isReconciling ? 'Running Analysis (Python)...' : 'Run Reconciliation'}
                  </button>
                </div>
              </div>

              {/* Results Area */}
              {(reconcileResult || isReconciling) && (
                <div className="bg-white border rounded-xl shadow-sm p-8 flex-1 overflow-auto animate-in fade-in slide-in-from-bottom-4 duration-500">
                   <h3 className="text-lg font-bold text-gray-900 mb-4 border-b pb-2">Analysis Results</h3>
                   {isReconciling ? (
                     <div className="flex flex-col items-center justify-center py-12 space-y-4">
                       <Loader2 size={40} className="animate-spin text-indigo-600" />
                       <p className="text-gray-500 text-sm">Gemini is writing and executing Python code...</p>
                     </div>
                   ) : (
                     <div 
                       className="prose prose-slate max-w-none prose-headings:text-indigo-900 prose-a:text-indigo-600"
                       dangerouslySetInnerHTML={{ __html: window.marked ? window.marked.parse(reconcileResult?.report || '') : reconcileResult?.report }}
                     />
                   )}
                </div>
              )}

            </div>
          </div>
        )}

      </main>

      {/* Document Viewer Modal */}
      {selectedDocument && (
        <DocumentViewer 
          document={selectedDocument} 
          fields={fields} 
          onClose={() => setSelectedDocId(null)} 
        />
      )}
    </div>
  );
};

export default App;