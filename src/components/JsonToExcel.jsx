import React, { useState, useEffect, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  FileSpreadsheet,
  UploadCloud,
  FileCode,
  CheckCircle2,
  AlertCircle,
  Copy,
  Trash2,
  FileText,
  Sparkles,
  RefreshCw,
  Check,
  FileUp,
  Table as TableIcon,
  ChevronRight
} from 'lucide-react';

// Sample JSON data for quick user onboarding
const SAMPLE_JSON = JSON.stringify(
  [
    {
      id: "USR-1001",
      name: "Sarah Jenkins",
      email: "sarah.j@example.com",
      role: "Engineering Lead",
      department: {
        name: "Cloud Infrastructure",
        location: "San Francisco, CA",
        floor: 4
      },
      skills: ["React", "TypeScript", "Node.js", "AWS"],
      metrics: {
        projectsCompleted: 14,
        rating: 4.9,
        active: true
      },
      salary: 165000,
      joinedDate: "2023-04-15"
    },
    {
      id: "USR-1002",
      name: "Marcus Vance",
      email: "marcus.v@example.com",
      role: "Senior UX Designer",
      department: {
        name: "Design Systems",
        location: "Austin, TX",
        floor: 2
      },
      skills: ["Figma", "Design Systems", "Prototyping", "User Research"],
      metrics: {
        projectsCompleted: 22,
        rating: 4.8,
        active: true
      },
      salary: 148000,
      joinedDate: "2022-09-01"
    },
    {
      id: "USR-1003",
      name: "Aaliyah Chen",
      email: "aaliyah.c@example.com",
      role: "Product Manager",
      department: {
        name: "Enterprise Growth",
        location: "New York, NY",
        floor: 12
      },
      skills: ["Strategy", "Analytics", "Agile", "SQL"],
      metrics: {
        projectsCompleted: 19,
        rating: 4.95,
        active: true
      },
      salary: 155000,
      joinedDate: "2023-01-10"
    },
    {
      id: "USR-1004",
      name: "Devon Bailey",
      email: "devon.b@example.com",
      role: "Security Analyst",
      department: {
        name: "SecOps",
        location: "Remote",
        floor: null
      },
      skills: ["Penetration Testing", "ISO 27001", "Cryptography"],
      metrics: {
        projectsCompleted: 9,
        rating: 4.7,
        active: false
      },
      salary: 135000,
      joinedDate: "2024-02-20"
    }
  ],
  null,
  2
);

/**
 * Deeply flatten nested objects & arrays for tabular presentation
 */
function flattenObject(obj, prefix = '', res = {}) {
  if (obj === null || typeof obj !== 'object') {
    res[prefix || 'value'] = obj;
    return res;
  }

  for (const key of Object.keys(obj)) {
    const val = obj[key];
    const newKey = prefix ? `${prefix}_${key}` : key;

    if (val === null || val === undefined) {
      res[newKey] = '';
    } else if (Array.isArray(val)) {
      if (val.length === 0) {
        res[newKey] = '[]';
      } else if (val.every((item) => typeof item !== 'object' || item === null)) {
        res[newKey] = val.join(', ');
      } else {
        val.forEach((item, index) => {
          if (typeof item === 'object' && item !== null) {
            flattenObject(item, `${newKey}_${index + 1}`, res);
          } else {
            res[`${newKey}_${index + 1}`] = item;
          }
        });
      }
    } else if (typeof val === 'object') {
      flattenObject(val, newKey, res);
    } else {
      res[newKey] = val;
    }
  }

  return res;
}

export default function JsonToExcel() {
  const [inputMode, setInputMode] = useState('text'); // 'text' | 'file'
  const [rawJson, setRawJson] = useState('');
  const [fileName, setFileName] = useState('data-export.xlsx');
  const [sheetName, setSheetName] = useState('Sheet1');
  const [flattenNested, setFlattenNested] = useState(true);
  
  // UI & Processing state
  const [parsedData, setParsedData] = useState(null);
  const [parseError, setParseError] = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  const [uploadedFile, setUploadedFile] = useState(null);
  const [isExporting, setIsExporting] = useState(false);
  const [exportSuccess, setExportSuccess] = useState(false);
  const [copiedNotification, setCopiedNotification] = useState(false);

  const fileInputRef = useRef(null);
  const textareaRef = useRef(null);

  // Automatically parse and validate JSON on change
  useEffect(() => {
    if (!rawJson.trim()) {
      setParsedData(null);
      setParseError(null);
      return;
    }

    try {
      const parsed = JSON.parse(rawJson);
      let normalized = [];

      if (Array.isArray(parsed)) {
        normalized = parsed.map((item) => {
          if (typeof item === 'object' && item !== null) {
            return flattenNested ? flattenObject(item) : item;
          }
          return { value: item };
        });
      } else if (typeof parsed === 'object' && parsed !== null) {
        // If single object, check if it contains a prominent array
        const keys = Object.keys(parsed);
        const arrayKey = keys.find((k) => Array.isArray(parsed[k]) && parsed[k].length > 0 && typeof parsed[k][0] === 'object');

        if (arrayKey && keys.length <= 3) {
          normalized = parsed[arrayKey].map((item) => (flattenNested ? flattenObject(item) : item));
        } else {
          normalized = [flattenNested ? flattenObject(parsed) : parsed];
        }
      } else {
        normalized = [{ value: parsed }];
      }

      setParsedData(normalized);
      setParseError(null);
    } catch (err) {
      setParsedData(null);
      setParseError(err.message || 'Invalid JSON format. Please verify syntax.');
    }
  }, [rawJson, flattenNested]);

  // Handle Drag and Drop
  const handleDragOver = (e) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      processFile(files[0]);
    }
  };

  const handleFileInputChange = (e) => {
    if (e.target.files && e.target.files.length > 0) {
      processFile(e.target.files[0]);
    }
  };

  const processFile = (file) => {
    if (!file) return;

    if (!file.name.endsWith('.json') && file.type !== 'application/json' && file.type !== 'text/plain') {
      setParseError('Please upload a valid .json file.');
      return;
    }

    setUploadedFile({
      name: file.name,
      size: (file.size / 1024).toFixed(1) + ' KB'
    });

    const baseName = file.name.replace(/\.json$/i, '');
    setFileName(`${baseName || 'data-export'}.xlsx`);

    const reader = new FileReader();
    reader.onload = (event) => {
      const content = event.target?.result;
      if (typeof content === 'string') {
        setRawJson(content);
      }
    };
    reader.onerror = () => {
      setParseError('Failed to read file. Please try again.');
    };
    reader.readAsText(file);
  };

  const handleLoadSample = () => {
    setRawJson(SAMPLE_JSON);
    setUploadedFile(null);
    setFileName('sample-employees.xlsx');
  };

  const handleClear = () => {
    setRawJson('');
    setParsedData(null);
    setParseError(null);
    setUploadedFile(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const handleFormatJson = () => {
    if (!rawJson.trim()) return;
    try {
      const parsed = JSON.parse(rawJson);
      setRawJson(JSON.stringify(parsed, null, 2));
    } catch {
      // Keep as-is if malformed
    }
  };

  const handleCopyJson = () => {
    if (!rawJson) return;
    navigator.clipboard.writeText(rawJson);
    setCopiedNotification(true);
    setTimeout(() => setCopiedNotification(false), 2000);
  };

  // Convert and export to XLSX
  const handleExport = () => {
    if (!parsedData || parsedData.length === 0) return;

    setIsExporting(true);

    try {
      // 1. Create a clean worksheet from JSON array
      const worksheet = XLSX.utils.json_to_sheet(parsedData);

      // 2. Auto-fit column widths based on maximum string length
      const colWidths = [];
      const headers = Object.keys(parsedData[0] || {});
      
      headers.forEach((header, colIndex) => {
        let maxLen = header.length;
        // Sample up to first 100 rows for performance
        const sampleRows = parsedData.slice(0, 100);
        sampleRows.forEach((row) => {
          const val = row[header];
          const len = val !== null && val !== undefined ? String(val).length : 0;
          if (len > maxLen) maxLen = len;
        });
        // Add padding, clamp width between 10 and 60 chars
        colWidths[colIndex] = { wch: Math.min(Math.max(maxLen + 4, 12), 65) };
      });

      worksheet['!cols'] = colWidths;

      // 3. Create a new workbook and append sheet
      const workbook = XLSX.utils.book_new();
      const safeSheetName = (sheetName || 'Sheet1').slice(0, 31); // Excel max 31 chars
      XLSX.utils.book_append_sheet(workbook, worksheet, safeSheetName);

      // 4. Save and trigger browser download
      const targetFileName = fileName.trim().endsWith('.xlsx') ? fileName.trim() : `${fileName.trim() || 'data-export'}.xlsx`;
      XLSX.writeFile(workbook, targetFileName);

      setExportSuccess(true);
      setTimeout(() => setExportSuccess(false), 4000);
    } catch (err) {
      console.error(err);
      setParseError('Failed to generate Excel file: ' + err.message);
    } finally {
      setIsExporting(false);
    }
  };

  // Columns for preview table
  const previewColumns = useMemo(() => {
    if (!parsedData || parsedData.length === 0) return [];
    return Object.keys(parsedData[0]);
  }, [parsedData]);

  // Preview slice (first 5 rows)
  const previewRows = useMemo(() => {
    if (!parsedData || parsedData.length === 0) return [];
    return parsedData.slice(0, 5);
  }, [parsedData]);

  return (
    <div class="space-y-6 sm:space-y-8">
      {/* Main Conversion Workspace Card */}
      <div class="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden transition-all duration-200">
        
        {/* Responsive Toolbar */}
        <div class="border-b border-slate-200 px-3.5 py-3 sm:px-5 sm:py-4 bg-slate-50/80 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          
          {/* Input Mode Selector & File Tag */}
          <div class="flex flex-wrap items-center gap-2">
            <div class="inline-flex bg-slate-200/80 p-1 rounded-xl w-full sm:w-auto justify-stretch">
              <button
                type="button"
                onClick={() => setInputMode('text')}
                class={`flex-1 sm:flex-none flex items-center justify-center gap-2 px-3 py-1.5 sm:px-3.5 sm:py-1.5 rounded-lg text-xs font-semibold transition-all ${
                  inputMode === 'text'
                    ? 'bg-white text-slate-900 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <FileCode class="w-3.5 h-3.5 text-emerald-600" />
                Paste JSON
              </button>
              <button
                type="button"
                onClick={() => setInputMode('file')}
                class={`flex-1 sm:flex-none flex items-center justify-center gap-2 px-3 py-1.5 sm:px-3.5 sm:py-1.5 rounded-lg text-xs font-semibold transition-all ${
                  inputMode === 'file'
                    ? 'bg-white text-slate-900 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <UploadCloud class="w-3.5 h-3.5 text-emerald-600" />
                Upload .json
              </button>
            </div>

            {uploadedFile && (
              <span class="inline-flex items-center gap-1.5 text-xs text-slate-600 bg-white border border-slate-200 px-2.5 py-1 rounded-lg max-w-full truncate">
                <FileText class="w-3 h-3 text-slate-400 shrink-0" />
                <span class="font-medium text-slate-800 truncate">{uploadedFile.name}</span>
                <span class="text-slate-400 shrink-0">({uploadedFile.size})</span>
              </span>
            )}
          </div>

          {/* Quick Action Buttons */}
          <div class="flex flex-wrap items-center gap-1.5 sm:gap-2 w-full sm:w-auto justify-start sm:justify-end pt-1 sm:pt-0 border-t sm:border-t-0 border-slate-200/60 sm:border-transparent">
            <button
              type="button"
              onClick={handleLoadSample}
              class="inline-flex items-center justify-center gap-1.5 px-2.5 py-1.5 sm:px-3 sm:py-1.5 rounded-lg text-xs font-medium bg-white text-slate-700 border border-slate-200 hover:bg-slate-50 hover:border-slate-300 transition-colors shadow-2xs active:scale-95"
            >
              <Sparkles class="w-3.5 h-3.5 text-amber-500" />
              <span>Sample</span>
            </button>

            {rawJson && (
              <>
                <button
                  type="button"
                  onClick={handleFormatJson}
                  title="Format JSON indentation"
                  class="inline-flex items-center justify-center gap-1.5 px-2.5 py-1.5 sm:px-3 sm:py-1.5 rounded-lg text-xs font-medium bg-white text-slate-700 border border-slate-200 hover:bg-slate-50 hover:border-slate-300 transition-colors shadow-2xs active:scale-95"
                >
                  <RefreshCw class="w-3.5 h-3.5 text-slate-500" />
                  <span>Prettify</span>
                </button>
                
                <button
                  type="button"
                  onClick={handleCopyJson}
                  title="Copy JSON to clipboard"
                  class="inline-flex items-center justify-center gap-1.5 px-2.5 py-1.5 sm:px-3 sm:py-1.5 rounded-lg text-xs font-medium bg-white text-slate-700 border border-slate-200 hover:bg-slate-50 hover:border-slate-300 transition-colors shadow-2xs active:scale-95"
                >
                  {copiedNotification ? (
                    <>
                      <Check class="w-3.5 h-3.5 text-emerald-600" />
                      <span class="text-emerald-700 font-semibold">Copied!</span>
                    </>
                  ) : (
                    <>
                      <Copy class="w-3.5 h-3.5 text-slate-500" />
                      <span>Copy</span>
                    </>
                  )}
                </button>

                <button
                  type="button"
                  onClick={handleClear}
                  title="Clear all input"
                  class="inline-flex items-center justify-center gap-1.5 px-2.5 py-1.5 sm:px-3 sm:py-1.5 rounded-lg text-xs font-semibold text-rose-700 bg-rose-50 hover:bg-rose-100 border border-rose-200 transition-colors shadow-2xs active:scale-95 ml-auto sm:ml-0"
                >
                  <Trash2 class="w-3.5 h-3.5 text-rose-600" />
                  <span>Clear</span>
                </button>
              </>
            )}
          </div>
        </div>

        {/* Input Area */}
        <div class="p-3.5 sm:p-6">
          {inputMode === 'text' ? (
            <div class="relative">
              <textarea
                ref={textareaRef}
                value={rawJson}
                onChange={(e) => setRawJson(e.target.value)}
                placeholder='Paste raw JSON here... e.g. [{"id": 1, "name": "Alice", "role": "Developer"}]'
                rows={9}
                class="w-full font-mono text-xs sm:text-sm bg-slate-50 text-slate-900 border border-slate-300 rounded-xl p-3 sm:p-4 focus:outline-hidden focus:ring-2 focus:ring-emerald-500/30 focus:border-emerald-600 transition duration-150 resize-y"
                spellCheck="false"
              />
              {rawJson && (
                <button
                  type="button"
                  onClick={handleClear}
                  title="Clear JSON input"
                  aria-label="Clear JSON"
                  class="absolute top-2.5 right-2.5 p-1.5 rounded-lg bg-white/95 text-slate-400 hover:text-rose-600 hover:bg-rose-50 border border-slate-200 shadow-2xs transition-colors"
                >
                  <Trash2 class="w-3.5 h-3.5" />
                </button>
              )}
              {!rawJson && (
                <div class="absolute inset-0 flex flex-col items-center justify-center pointer-events-none p-4 sm:p-6 text-center">
                  <FileCode class="w-8 h-8 sm:w-10 sm:h-10 text-slate-300 mb-2" />
                  <p class="text-xs sm:text-sm font-medium text-slate-500">Paste your JSON data here or click "Sample"</p>
                  <p class="text-[11px] sm:text-xs text-slate-400 mt-0.5">Supports arrays, single objects, and nested hierarchies</p>
                </div>
              )}
            </div>
          ) : (
            <div
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              onClick={() => fileInputRef.current?.click()}
              class={`border-2 border-dashed rounded-xl p-6 sm:p-12 text-center cursor-pointer transition-all duration-200 ${
                isDragging
                  ? 'border-emerald-500 bg-emerald-50/50 scale-[0.99]'
                  : 'border-slate-300 hover:border-emerald-500/70 hover:bg-slate-50/50'
              }`}
            >
              <input
                ref={fileInputRef}
                type="file"
                accept=".json,application/json"
                onChange={handleFileInputChange}
                class="hidden"
              />
              <div class="w-12 h-12 sm:w-14 sm:h-14 mx-auto rounded-2xl bg-emerald-50 text-emerald-600 flex items-center justify-center mb-3 sm:mb-4 border border-emerald-100 shadow-xs">
                <FileUp class="w-6 h-6 sm:w-7 sm:h-7" />
              </div>
              <h3 class="text-sm sm:text-base font-bold text-slate-800">
                {uploadedFile ? uploadedFile.name : 'Choose a JSON file or drag & drop here'}
              </h3>
              <p class="text-xs sm:text-sm text-slate-500 mt-1 max-w-sm mx-auto">
                Upload your local <code class="px-1.5 py-0.5 bg-slate-100 rounded text-slate-700 font-mono text-xs">.json</code> file. Processed 100% in your browser.
              </p>
              <button
                type="button"
                class="mt-3.5 sm:mt-4 px-4 py-2 rounded-xl text-xs font-semibold bg-white border border-slate-300 text-slate-700 shadow-2xs hover:bg-slate-50 transition-all inline-flex items-center gap-2"
              >
                <UploadCloud class="w-4 h-4 text-emerald-600" />
                Browse File
              </button>
            </div>
          )}

          {/* Inline Validation Error Banner */}
          {parseError && (
            <div class="mt-3 sm:mt-4 p-3.5 sm:p-4 rounded-xl bg-rose-50 border border-rose-200 flex items-start gap-2.5 sm:gap-3 text-rose-800 animate-in fade-in duration-200">
              <AlertCircle class="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />
              <div class="text-xs sm:text-sm">
                <p class="font-semibold">JSON Syntax Validation Error</p>
                <p class="mt-0.5 text-xs text-rose-700 font-mono break-all">{parseError}</p>
                <p class="mt-1.5 text-[11px] sm:text-xs text-rose-600">
                  Tip: Ensure keys and string values are enclosed in double quotes <code class="bg-rose-100 px-1 py-0.5 rounded text-rose-900 font-mono">"key": "value"</code> without trailing commas.
                </p>
              </div>
            </div>
          )}
        </div>

        {/* Responsive Configuration Options Bar */}
        <div class="border-t border-slate-200 px-3.5 sm:px-6 py-4 bg-slate-50/70">
          <div class="flex flex-col lg:flex-row lg:items-center justify-between gap-4 sm:gap-5">
            
            {/* Flatten Option Toggle */}
            <div class="flex items-center justify-between sm:justify-start gap-3 bg-white sm:bg-transparent p-3 sm:p-0 rounded-xl border sm:border-0 border-slate-200/80">
              <label class="relative flex items-center gap-3 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={flattenNested}
                  onChange={(e) => setFlattenNested(e.target.checked)}
                  class="sr-only peer"
                />
                <div class="w-11 h-6 bg-slate-300 peer-focus:outline-hidden rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-emerald-600"></div>
                <div>
                  <span class="text-xs sm:text-sm font-semibold text-slate-800 block">Flatten nested objects</span>
                  <span class="text-[11px] sm:text-xs text-slate-500 block -mt-0.5">e.g. user.name &rarr; user_name</span>
                </div>
              </label>
            </div>

            {/* Inputs & Download Action */}
            <div class="grid grid-cols-1 sm:grid-cols-2 lg:flex lg:items-end gap-3">
              {/* Filename Input */}
              <div class="w-full lg:w-48">
                <label class="block text-xs font-semibold text-slate-600 mb-1">Export File Name</label>
                <input
                  type="text"
                  value={fileName}
                  onChange={(e) => setFileName(e.target.value)}
                  placeholder="data-export.xlsx"
                  class="w-full text-xs font-mono bg-white border border-slate-300 rounded-lg px-3 py-2 text-slate-800 focus:outline-hidden focus:ring-1 focus:ring-emerald-500 focus:border-emerald-600 shadow-2xs"
                />
              </div>

              {/* Sheet Name Input */}
              <div class="w-full lg:w-28">
                <label class="block text-xs font-semibold text-slate-600 mb-1">Sheet Name</label>
                <input
                  type="text"
                  value={sheetName}
                  onChange={(e) => setSheetName(e.target.value)}
                  placeholder="Sheet1"
                  maxLength={31}
                  class="w-full text-xs font-mono bg-white border border-slate-300 rounded-lg px-3 py-2 text-slate-800 focus:outline-hidden focus:ring-1 focus:ring-emerald-500 focus:border-emerald-600 shadow-2xs"
                />
              </div>

              {/* Main Export Button */}
              <div class="w-full sm:col-span-2 lg:col-span-1 lg:w-auto pt-1 sm:pt-0">
                <button
                  type="button"
                  disabled={!parsedData || parsedData.length === 0 || isExporting}
                  onClick={handleExport}
                  class={`w-full lg:w-auto inline-flex items-center justify-center gap-2 px-5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold shadow-md transition-all duration-200 min-h-[42px] ${
                    !parsedData || parsedData.length === 0 || isExporting
                      ? 'bg-slate-200 text-slate-400 cursor-not-allowed shadow-none'
                      : 'bg-emerald-600 hover:bg-emerald-700 text-white shadow-emerald-600/20 active:scale-[0.98]'
                  }`}
                >
                  {isExporting ? (
                    <>
                      <RefreshCw class="w-4 h-4 animate-spin" />
                      Generating...
                    </>
                  ) : (
                    <>
                      <FileSpreadsheet class="w-4 h-4" />
                      Download Excel (.xlsx)
                    </>
                  )}
                </button>
              </div>
            </div>
          </div>
        </div>

        {/* Success Alert Banner */}
        {exportSuccess && (
          <div class="bg-emerald-50 border-t border-emerald-200 px-4 sm:px-6 py-3 flex flex-col sm:flex-row sm:items-center justify-between gap-1 text-emerald-800 text-xs sm:text-sm animate-in fade-in duration-200">
            <div class="flex items-center gap-2">
              <CheckCircle2 class="w-4 h-4 text-emerald-600 shrink-0" />
              <span>
                <strong>Success!</strong> Spreadhseet downloaded as <code class="font-mono bg-emerald-100/70 px-1 py-0.5 rounded text-emerald-900 text-xs">{fileName}</code>
              </span>
            </div>
            <span class="text-xs text-emerald-600 font-medium self-end sm:self-auto">100% Client-Side Export</span>
          </div>
        )}
      </div>

      {/* Live Data Preview Section */}
      {parsedData && parsedData.length > 0 && (
        <div class="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden animate-in fade-in duration-300">
          
          {/* Preview Header & Stats */}
          <div class="border-b border-slate-200 px-3.5 py-3 sm:px-5 sm:py-4 bg-slate-50 flex flex-col sm:flex-row sm:items-center justify-between gap-2.5">
            <div class="flex items-center gap-2.5">
              <div class="w-7 h-7 rounded-lg bg-emerald-100 text-emerald-700 flex items-center justify-center shrink-0">
                <TableIcon class="w-4 h-4" />
              </div>
              <div>
                <h3 class="text-xs sm:text-sm font-bold text-slate-900">Spreadsheet Live Preview</h3>
                <p class="text-[11px] sm:text-xs text-slate-500">First {previewRows.length} of {parsedData.length} records</p>
              </div>
            </div>

            {/* Metrics Badges */}
            <div class="flex items-center gap-2 self-start sm:self-auto">
              <span class="inline-flex items-center gap-1.5 px-2.5 py-0.5 sm:px-3 sm:py-1 rounded-full text-[11px] sm:text-xs font-semibold bg-slate-100 text-slate-700 border border-slate-200">
                <span class="w-2 h-2 rounded-full bg-emerald-500"></span>
                {parsedData.length} {parsedData.length === 1 ? 'Row' : 'Rows'}
              </span>
              <span class="inline-flex items-center gap-1.5 px-2.5 py-0.5 sm:px-3 sm:py-1 rounded-full text-[11px] sm:text-xs font-semibold bg-slate-100 text-slate-700 border border-slate-200">
                {previewColumns.length} Columns
              </span>
            </div>
          </div>

          {/* Table Container with Horizontal Scrolling */}
          <div class="overflow-x-auto max-w-full overscroll-x-contain">
            <table class="w-full text-left border-collapse text-xs sm:text-sm min-w-full">
              <thead>
                <tr class="bg-slate-100/90 border-b border-slate-200 text-slate-700 uppercase tracking-wider font-mono text-[10px] sm:text-[11px]">
                  <th class="px-3 sm:px-4 py-2.5 sm:py-3 text-slate-400 font-normal w-10 sm:w-12 text-center border-r border-slate-200 sticky left-0 bg-slate-100 z-10">
                    #
                  </th>
                  {previewColumns.map((col) => (
                    <th key={col} class="px-3 sm:px-4 py-2.5 sm:py-3 font-semibold border-r border-slate-200 last:border-r-0 whitespace-nowrap">
                      {col}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody class="divide-y divide-slate-200 font-sans">
                {previewRows.map((row, idx) => (
                  <tr key={idx} class="hover:bg-slate-50/80 transition-colors">
                    <td class="px-3 sm:px-4 py-2 sm:py-2.5 text-slate-400 font-mono text-[10px] sm:text-[11px] text-center bg-slate-50 border-r border-slate-200 sticky left-0 z-10">
                      {idx + 1}
                    </td>
                    {previewColumns.map((col) => {
                      const val = row[col];
                      const displayVal = val === null || val === undefined ? '' : typeof val === 'object' ? JSON.stringify(val) : String(val);
                      return (
                        <td key={col} class="px-3 sm:px-4 py-2 sm:py-2.5 text-slate-700 border-r border-slate-200 last:border-r-0 max-w-[180px] sm:max-w-xs truncate whitespace-nowrap" title={displayVal}>
                          {displayVal === '' ? (
                            <span class="text-slate-300 italic">null</span>
                          ) : (
                            displayVal
                          )}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Preview Footer note */}
          <div class="border-t border-slate-200 px-3.5 py-2.5 sm:px-5 sm:py-3 bg-slate-50/80 flex flex-col sm:flex-row sm:items-center justify-between gap-2 text-[11px] sm:text-xs text-slate-500">
            <span>
              Showing {previewRows.length} of {parsedData.length} records. All rows will be included in the Excel export.
            </span>
            <button
              type="button"
              onClick={handleExport}
              class="font-semibold text-emerald-700 hover:text-emerald-800 inline-flex items-center gap-1 hover:underline self-start sm:self-auto"
            >
              <span>Download full dataset</span>
              <ChevronRight class="w-3.5 h-3.5" />
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
