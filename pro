<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>DocuScan OCR Pro — Industrial Multi-Manufacturer PDF Collector</title>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet" />
  <script src="https://cdn.tailwindcss.com/3.4.17"></script>
  <script src="https://unpkg.com/lucide@0.460.0/dist/umd/lucide.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js"></script>
  <style>
    body {
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    }
    .drop-zone-active {
      border-color: #4f46e5 !important;
      background-color: #eef2ff !important;
    }
    .cell-input:focus {
      outline: 2px solid #6366f1;
      outline-offset: -2px;
      background-color: #ffffff;
    }
    @keyframes fadeSlideIn {
      from { opacity: 0; transform: translateY(8px); }
      to { opacity: 1; transform: translateY(0); }
    }
    .toast-enter {
      animation: fadeSlideIn 0.2s ease-out forwards;
    }
  </style>
</head>
<body class="bg-slate-50 text-slate-900 min-h-screen antialiased">
  <!-- Header -->
  <header class="bg-white border-b border-slate-200 sticky top-0 z-30">
    <div class="max-w-6xl mx-auto px-4 sm:px-6 h-16 flex items-center justify-between gap-3">
      <div class="flex items-center gap-2.5 min-w-0">
        <div class="w-9 h-9 rounded-xl bg-indigo-600 text-white flex items-center justify-center shrink-0 shadow-sm">
          <i data-lucide="scan-text" class="w-5 h-5"></i>
        </div>
        <div class="min-w-0">
          <h1 class="text-base sm:text-lg font-bold tracking-tight text-slate-900 truncate">DocuScan OCR Collector</h1>
          <p class="text-xs text-slate-500 hidden sm:block">Industrial multi-manufacturer PDF OCR &amp; spec collector (Excel &amp; PDF export)</p>
        </div>
      </div>

      <div class="flex items-center gap-2 shrink-0">
        <input id="projectFileInput" type="file" accept=".json,application/json" class="hidden" />
        <button id="loadProjectBtn" type="button" title="Import saved JSON project" class="hidden sm:inline-flex items-center gap-1.5 text-xs font-medium text-slate-700 hover:text-indigo-700 bg-white hover:bg-indigo-50 border border-slate-200 px-2.5 py-2 rounded-lg transition">
          <i data-lucide="folder-open" class="w-3.5 h-3.5"></i>
          <span>Import JSON</span>
        </button>
        <button id="saveProjectBtn" type="button" title="Save current collection as JSON" class="hidden sm:inline-flex items-center gap-1.5 text-xs font-medium text-slate-700 hover:text-indigo-700 bg-white hover:bg-indigo-50 border border-slate-200 px-2.5 py-2 rounded-lg transition">
          <i data-lucide="save" class="w-3.5 h-3.5"></i>
          <span>Save JSON</span>
        </button>
        <button id="openGuideBtn" type="button" class="inline-flex items-center gap-1.5 text-xs sm:text-sm font-medium text-slate-700 hover:text-indigo-700 bg-slate-100 hover:bg-indigo-50 border border-slate-200 hover:border-indigo-200 px-3 py-2 rounded-lg transition">
          <i data-lucide="file-code-2" class="w-4 h-4"></i>
          <span>Notepad Setup</span>
        </button>
      </div>
    </div>
  </header>

  <!-- Main Content -->
  <main class="max-w-6xl mx-auto px-4 sm:px-6 py-6 sm:py-8 space-y-6">
    <!-- Persistent Diagnostic Error Banner -->
    <div id="errorBanner" class="hidden bg-rose-50 border border-rose-200 text-rose-800 rounded-xl p-4 flex items-start justify-between gap-3">
      <div class="flex items-start gap-2.5 min-w-0">
        <i data-lucide="alert-triangle" class="w-5 h-5 text-rose-600 shrink-0 mt-0.5"></i>
        <div class="text-xs sm:text-sm min-w-0">
          <p id="errorBannerTitle" class="font-bold text-rose-900">Extraction Error</p>
          <p id="errorBannerDetail" class="mt-0.5 text-rose-700 break-words"></p>
        </div>
      </div>
      <button id="dismissErrorBtn" type="button" aria-label="Dismiss error" class="p-1 text-rose-500 hover:text-rose-800 rounded-md shrink-0">
        <i data-lucide="x" class="w-4 h-4"></i>
      </button>
    </div>

    <!-- Step 1: Tailor Fields & Extraction Rules -->
    <section class="bg-white rounded-2xl border border-slate-200 shadow-sm p-5 sm:p-6">
      <div class="flex flex-wrap items-center justify-between gap-3">
        <div class="flex items-center gap-2">
          <span class="w-6 h-6 rounded-full bg-indigo-100 text-indigo-700 text-xs font-bold flex items-center justify-center">1</span>
          <div>
            <h2 class="text-base font-semibold text-slate-900">Tailor Extraction Fields (Any Equipment)</h2>
            <p class="text-xs text-slate-500">Pick an industry preset or customize columns &amp; auto-match keywords.</p>
          </div>
        </div>
        <button id="toggleSchemaBtn" type="button" class="inline-flex items-center gap-1.5 text-xs font-semibold text-indigo-700 bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 px-3 py-1.5 rounded-lg transition">
          <i data-lucide="sliders-horizontal" class="w-3.5 h-3.5"></i>
          <span id="toggleSchemaLabel">Hide Field Builder</span>
        </button>
      </div>

      <div id="schemaPanel" class="mt-4 pt-4 border-t border-slate-100 space-y-4">
        <div class="flex flex-wrap items-center gap-2">
          <span class="text-xs font-semibold text-slate-600 mr-1">Quick Presets:</span>
          <button id="presetElectricalBtn" type="button" class="text-xs font-medium px-3 py-1.5 rounded-lg border border-indigo-200 bg-indigo-50 text-indigo-700 hover:bg-indigo-100 transition">
            Electrical Equipment Specs
          </button>
          <button id="presetCommercialBtn" type="button" class="text-xs font-medium px-3 py-1.5 rounded-lg border border-slate-200 bg-slate-50 text-slate-700 hover:bg-slate-100 transition">
            Quotation / Commercial
          </button>
          <button id="presetGeneralBtn" type="button" class="text-xs font-medium px-3 py-1.5 rounded-lg border border-slate-200 bg-slate-50 text-slate-700 hover:bg-slate-100 transition">
            Auto-Detect Table (General)
          </button>
        </div>

        <div id="schemaFieldsList" class="grid grid-cols-1 md:grid-cols-2 gap-2.5"></div>

        <div class="grid grid-cols-1 sm:grid-cols-12 gap-2.5 items-end bg-slate-50 p-3 rounded-xl border border-slate-200">
          <div class="sm:col-span-4">
            <label for="newFieldNameInput" class="block text-xs font-medium text-slate-600 mb-1">New Column Name</label>
            <input id="newFieldNameInput" type="text" placeholder="e.g., Insulation Class, Burden, Ratio" class="w-full text-xs sm:text-sm bg-white border border-slate-300 rounded-lg px-3 py-2 text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500" />
          </div>
          <div class="sm:col-span-5">
            <label for="newFieldKeywordsInput" class="block text-xs font-medium text-slate-600 mb-1">Auto-Match Keywords (comma-separated)</label>
            <input id="newFieldKeywordsInput" type="text" placeholder="e.g., insulation, class, thermal" class="w-full text-xs sm:text-sm bg-white border border-slate-300 rounded-lg px-3 py-2 text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500" />
          </div>
          <div class="sm:col-span-3 flex gap-2">
            <button id="addFieldBtn" type="button" class="flex-1 inline-flex items-center justify-center gap-1.5 bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold px-3 py-2 rounded-lg transition">
              <i data-lucide="plus" class="w-3.5 h-3.5"></i>
              Add Field
            </button>
          </div>
        </div>

        <div class="flex flex-wrap items-center justify-between gap-2 text-xs text-slate-500">
          <span>Every row is tagged with <strong>Manufacturer</strong>, <strong>Equipment</strong>, <strong>Source File</strong>, and <strong>Confidence</strong>.</span>
          <button id="reapplySchemaBtn" type="button" class="inline-flex items-center gap-1.5 text-xs font-semibold text-indigo-700 hover:text-indigo-900 underline">
            <i data-lucide="refresh-cw" class="w-3.5 h-3.5"></i>
            Re-apply Fields to Collected PDFs
          </button>
        </div>
      </div>
    </section>

    <!-- Step 2: Tag Manufacturer, Upload PDF(s) & Run Industrial OCR -->
    <section class="bg-white rounded-2xl border border-slate-200 shadow-sm p-5 sm:p-6">
      <div class="flex flex-wrap items-center justify-between gap-3 mb-4">
        <div class="flex items-center gap-2">
          <span class="w-6 h-6 rounded-full bg-indigo-100 text-indigo-700 text-xs font-bold flex items-center justify-center">2</span>
          <h2 class="text-base font-semibold text-slate-900">Tag Manufacturer, Upload PDF(s) &amp; Extract</h2>
        </div>
        <button id="loadSampleBtn" type="button" class="inline-flex items-center gap-1.5 text-xs font-medium text-indigo-600 hover:text-indigo-800 bg-indigo-50 hover:bg-indigo-100 px-3 py-1.5 rounded-lg transition">
          <i data-lucide="sparkles" class="w-3.5 h-3.5"></i>
          Load Multi-Manufacturer Sample
        </button>
      </div>

      <!-- Manufacturer & Equipment Tagging Bar -->
      <div class="grid grid-cols-1 md:grid-cols-12 gap-3 mb-4 bg-slate-50/80 p-3.5 rounded-xl border border-slate-200">
        <div class="md:col-span-4">
          <label for="manufacturerSelect" class="block text-xs font-medium text-slate-600 mb-1">Manufacturer</label>
          <select id="manufacturerSelect" class="w-full text-sm bg-white border border-slate-300 rounded-lg px-3 py-2 text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500"></select>
        </div>
        <div class="md:col-span-4">
          <label for="newManufacturerInput" class="block text-xs font-medium text-slate-600 mb-1">Add Custom Manufacturer</label>
          <div class="flex gap-1.5">
            <input id="newManufacturerInput" type="text" placeholder="e.g., Havells, Kirloskar" class="flex-1 min-w-0 text-sm bg-white border border-slate-300 rounded-lg px-3 py-2 text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500" />
            <button id="addManufacturerBtn" type="button" class="shrink-0 px-3 py-2 bg-slate-800 hover:bg-slate-900 text-white text-xs font-semibold rounded-lg transition">Add</button>
          </div>
        </div>
        <div class="md:col-span-4">
          <label for="equipmentInput" class="block text-xs font-medium text-slate-600 mb-1">Equipment Name / Category</label>
          <input id="equipmentInput" type="text" placeholder="e.g., Current Transformer, Switchgear, Motor" class="w-full text-sm bg-white border border-slate-300 rounded-lg px-3 py-2 text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500" />
        </div>
      </div>

      <!-- Drop Zone (Supports Multi-File Batch) -->
      <div id="dropZone" role="button" tabindex="0" aria-label="Upload one or more PDF or image files" class="border-2 border-dashed border-slate-300 hover:border-indigo-400 bg-slate-50/70 rounded-xl p-6 sm:p-7 text-center cursor-pointer transition">
        <input id="fileInput" type="file" multiple accept=".pdf,application/pdf,image/png,image/jpeg,image/webp" class="hidden" />
        <div class="mx-auto w-11 h-11 rounded-full bg-indigo-50 text-indigo-600 flex items-center justify-center mb-2.5">
          <i data-lucide="upload-cloud" class="w-6 h-6"></i>
        </div>
        <p class="text-sm sm:text-base font-semibold text-slate-800">Drop one or multiple PDFs here, or click to browse</p>
        <p class="text-xs sm:text-sm text-slate-500 mt-1">Batch-ready for scanned &amp; digital datasheets (PDF, PNG, JPG). Works offline-safe in local browser.</p>

        <div id="selectedFileBadge" class="hidden mt-4 inline-flex items-center gap-2 bg-white border border-indigo-200 text-indigo-800 text-xs sm:text-sm font-medium px-3.5 py-2 rounded-lg shadow-sm max-w-full">
          <i data-lucide="files" class="w-4 h-4 shrink-0 text-indigo-600"></i>
          <span id="selectedFileName" class="truncate max-w-[220px] sm:max-w-md"></span>
          <span id="selectedFileSize" class="text-xs text-slate-500 shrink-0"></span>
        </div>
      </div>

      <!-- OCR Precision & Page Range Controls -->
      <div class="mt-4 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-6 gap-3 items-end">
        <div>
          <label for="ocrModeSelect" class="block text-xs font-medium text-slate-600 mb-1.5">Extraction Mode</label>
          <select id="ocrModeSelect" class="w-full text-sm bg-white border border-slate-300 rounded-lg px-2.5 py-2.5 text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500">
            <option value="hybrid">Smart Auto (Text + OCR)</option>
            <option value="force_ocr">Force Full OCR</option>
          </select>
        </div>

        <div>
          <label for="ocrLangSelect" class="block text-xs font-medium text-slate-600 mb-1.5">OCR Language</label>
          <select id="ocrLangSelect" class="w-full text-sm bg-white border border-slate-300 rounded-lg px-2.5 py-2.5 text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500">
            <option value="eng">English (eng)</option>
            <option value="tam">Tamil (tam)</option>
            <option value="mar">Marathi (mar)</option>
            <option value="spa">Spanish (spa)</option>
            <option value="fra">French (fra)</option>
            <option value="deu">German (deu)</option>
          </select>
        </div>

        <div>
          <label for="dpiScaleSelect" class="block text-xs font-medium text-slate-600 mb-1.5">Scan Quality (DPI)</label>
          <select id="dpiScaleSelect" class="w-full text-sm bg-white border border-slate-300 rounded-lg px-2.5 py-2.5 text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500">
            <option value="2.0">Standard (2.0x — Fast)</option>
            <option value="2.8" selected>High Precision (2.8x)</option>
            <option value="3.5">Ultra Fine Print (3.5x)</option>
          </select>
        </div>

        <div>
          <label for="pageRangeInput" class="block text-xs font-medium text-slate-600 mb-1.5">Pages (e.g. 1-3, 5)</label>
          <input id="pageRangeInput" type="text" placeholder="All pages" class="w-full text-sm bg-white border border-slate-300 rounded-lg px-2.5 py-2.5 text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500" />
        </div>

        <div>
          <label for="collectionModeSelect" class="block text-xs font-medium text-slate-600 mb-1.5">Collection Action</label>
          <select id="collectionModeSelect" class="w-full text-sm bg-white border border-slate-300 rounded-lg px-2.5 py-2.5 text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500">
            <option value="append">Append to Master</option>
            <option value="replace">Replace Existing</option>
          </select>
        </div>

        <button id="extractBtn" type="button" disabled class="w-full inline-flex items-center justify-center gap-2 bg-indigo-600 hover:bg-indigo-700 disabled:bg-slate-300 disabled:cursor-not-allowed text-white font-semibold text-sm px-4 py-2.5 rounded-lg shadow-sm transition">
          <i data-lucide="scan-line" class="w-4 h-4"></i>
          <span>Extract &amp; Add</span>
        </button>
      </div>

      <!-- Preprocessing Options + Reset -->
      <div class="mt-3 flex flex-wrap items-center justify-between gap-3 pt-3 border-t border-slate-100">
        <label class="inline-flex items-center gap-2 text-xs text-slate-700 cursor-pointer">
          <input id="enhanceScanCheckbox" type="checkbox" checked class="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500" />
          <span><strong>Enhance Scanned PDFs</strong> (Grayscale + High-Contrast Sharpening for clearer tables &amp; fine print)</span>
        </label>
        <button id="clearBtn" type="button" class="inline-flex items-center gap-1.5 text-xs font-medium text-slate-600 hover:text-rose-700 bg-slate-100 hover:bg-rose-50 border border-slate-200 px-3 py-1.5 rounded-lg transition">
          <i data-lucide="rotate-ccw" class="w-3.5 h-3.5"></i>
          <span>Reset All</span>
        </button>
      </div>

      <!-- Progress Panel -->
      <div id="progressPanel" class="hidden mt-5 bg-indigo-50/70 border border-indigo-200 rounded-xl p-4">
        <div class="flex flex-col sm:flex-row sm:items-center gap-4">
          <div class="flex-1 min-w-0">
            <div class="flex items-center justify-between gap-2 mb-1.5">
              <span id="progressStatus" class="text-xs sm:text-sm font-semibold text-indigo-900 truncate">Initializing OCR engine...</span>
              <span id="progressPercent" class="text-xs font-bold text-indigo-700 shrink-0">0%</span>
            </div>
            <div class="w-full bg-indigo-200/70 rounded-full h-2.5 overflow-hidden">
              <div id="progressBar" class="bg-indigo-600 h-2.5 rounded-full transition-all duration-200" style="width: 0%"></div>
            </div>
            <div class="mt-2 flex items-center justify-between gap-2">
              <span id="progressDetail" class="text-xs text-indigo-700 truncate">Preparing PDF pages...</span>
              <button id="cancelOcrBtn" type="button" class="text-xs font-semibold text-rose-600 hover:text-rose-800 underline shrink-0">Cancel</button>
            </div>
          </div>
          <div id="previewCanvasWrap" class="hidden shrink-0 w-24 h-32 bg-white border border-indigo-200 rounded-lg overflow-hidden shadow-sm flex items-center justify-center">
            <canvas id="pagePreviewCanvas" class="max-w-full max-h-full object-contain"></canvas>
          </div>
        </div>
      </div>
    </section>

    <!-- Step 3: Two Export Options -->
    <section class="bg-white rounded-2xl border border-slate-200 shadow-sm p-5 sm:p-6">
      <div class="flex flex-wrap items-center justify-between gap-2 mb-4">
        <div class="flex items-center gap-2">
          <span class="w-6 h-6 rounded-full bg-indigo-100 text-indigo-700 text-xs font-bold flex items-center justify-center">3</span>
          <h2 class="text-base font-semibold text-slate-900">Export Collected Manufacturer Data</h2>
        </div>
        <span id="exportReadyBadge" class="text-xs font-medium text-slate-500 bg-slate-100 px-2.5 py-1 rounded-full">
          Upload a PDF or load sample data to export
        </span>
      </div>

      <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div class="border border-emerald-200 bg-emerald-50/40 rounded-xl p-4 flex flex-col justify-between gap-4">
          <div class="flex items-start gap-3">
            <div class="w-10 h-10 rounded-xl bg-emerald-600 text-white flex items-center justify-center shrink-0">
              <i data-lucide="file-spreadsheet" class="w-5 h-5"></i>
            </div>
            <div>
              <h3 class="text-sm font-bold text-slate-900">Option 1: Export as Excel (.xlsx)</h3>
              <p class="text-xs text-slate-600 mt-1">Downloads a multi-sheet Excel workbook with Master Collection, Manufacturer Summary, and Raw OCR Text.</p>
            </div>
          </div>
          <button id="exportExcelBtn" type="button" disabled class="w-full inline-flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-300 disabled:cursor-not-allowed text-white font-semibold text-sm px-4 py-2.5 rounded-lg shadow-sm transition">
            <i data-lucide="download" class="w-4 h-4"></i>
            <span>Download Excel (.xlsx)</span>
          </button>
        </div>

        <div class="border border-rose-200 bg-rose-50/40 rounded-xl p-4 flex flex-col justify-between gap-4">
          <div class="flex items-start gap-3">
            <div class="w-10 h-10 rounded-xl bg-rose-600 text-white flex items-center justify-center shrink-0">
              <i data-lucide="file-text" class="w-5 h-5"></i>
            </div>
            <div>
              <h3 class="text-sm font-bold text-slate-900">Option 2: Export as PDF (.pdf)</h3>
              <p class="text-xs text-slate-600 mt-1">Downloads a formatted multi-manufacturer comparison report PDF with styled tables and optional raw text appendix.</p>
              <label class="inline-flex items-center gap-1.5 mt-2 text-xs text-slate-600 cursor-pointer">
                <input id="includeRawInPdfCheckbox" type="checkbox" checked class="rounded border-slate-300 text-rose-600 focus:ring-rose-500" />
                <span>Include raw OCR text appendix in PDF</span>
              </label>
            </div>
          </div>
          <button id="exportPdfBtn" type="button" disabled class="w-full inline-flex items-center justify-center gap-2 bg-rose-600 hover:bg-rose-700 disabled:bg-slate-300 disabled:cursor-not-allowed text-white font-semibold text-sm px-4 py-2.5 rounded-lg shadow-sm transition">
            <i data-lucide="download" class="w-4 h-4"></i>
            <span>Download PDF (.pdf)</span>
          </button>
        </div>
      </div>
    </section>

    <!-- Step 4: Master Collection Table & Raw Text Editor -->
    <section class="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
      <div class="p-5 sm:p-6 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3">
        <div class="flex items-center gap-2">
          <span class="w-6 h-6 rounded-full bg-indigo-100 text-indigo-700 text-xs font-bold flex items-center justify-center">4</span>
          <h2 class="text-base font-semibold text-slate-900">Master Collection &amp; Data Editor</h2>
        </div>

        <div class="inline-flex bg-slate-100 p-1 rounded-lg border border-slate-200">
          <button id="tabTableBtn" type="button" class="px-3 py-1.5 text-xs sm:text-sm font-semibold rounded-md bg-white text-indigo-700 shadow-sm transition">
            Master Table
          </button>
          <button id="tabRawBtn" type="button" class="px-3 py-1.5 text-xs sm:text-sm font-medium rounded-md text-slate-600 hover:text-slate-900 transition">
            Raw OCR Text
          </button>
        </div>
      </div>

      <!-- Summary Stats Strip -->
      <div class="grid grid-cols-2 sm:grid-cols-4 divide-x divide-y sm:divide-y-0 divide-slate-200 bg-slate-50 border-b border-slate-200 text-center">
        <div class="p-3">
          <p class="text-xs text-slate-500">Pages Processed</p>
          <p id="statPages" class="text-base font-bold text-slate-900 mt-0.5">0</p>
        </div>
        <div class="p-3">
          <p class="text-xs text-slate-500">Collected Rows</p>
          <p id="statRows" class="text-base font-bold text-slate-900 mt-0.5">0</p>
        </div>
        <div class="p-3">
          <p class="text-xs text-slate-500">Words Extracted</p>
          <p id="statWords" class="text-base font-bold text-slate-900 mt-0.5">0</p>
        </div>
        <div class="p-3">
          <p class="text-xs text-slate-500">Manufacturers</p>
          <p id="statMethod" class="text-base font-bold text-indigo-700 mt-0.5">—</p>
        </div>
      </div>

      <!-- Master Table View -->
      <div id="tableView" class="p-5 sm:p-6">
        <div class="flex flex-wrap items-center justify-between gap-3 mb-4">
          <div class="flex flex-wrap items-center gap-2.5">
            <div class="flex items-center gap-1.5">
              <label for="parseModeSelect" class="text-xs font-medium text-slate-600">Parsing:</label>
              <select id="parseModeSelect" class="text-xs sm:text-sm bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500">
                <option value="custom">Tailored Custom Fields</option>
                <option value="auto">Auto-Detect Columns</option>
                <option value="kv">Key : Value Pairs</option>
                <option value="lines">One Line per Row</option>
              </select>
            </div>
            <div class="flex items-center gap-1.5">
              <label for="filterManufacturerSelect" class="text-xs font-medium text-slate-600">Filter:</label>
              <select id="filterManufacturerSelect" class="text-xs sm:text-sm bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500">
                <option value="ALL">All Manufacturers</option>
              </select>
            </div>
            <input id="tableSearchInput" type="search" placeholder="Search specs or models..." class="text-xs sm:text-sm bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500 w-44 sm:w-56" />
          </div>
          <button id="addRowBtn" type="button" class="inline-flex items-center gap-1.5 text-xs font-semibold text-indigo-700 bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 px-3 py-1.5 rounded-lg transition">
            <i data-lucide="plus" class="w-3.5 h-3.5"></i>
            Add Row
          </button>
        </div>

        <div id="tableEmptyState" class="py-12 text-center border border-dashed border-slate-200 rounded-xl bg-slate-50/50">
          <i data-lucide="table-2" class="w-8 h-8 text-slate-400 mx-auto mb-2"></i>
          <p class="text-sm font-medium text-slate-600">No manufacturer data collected yet</p>
          <p class="text-xs text-slate-500 mt-1">Select a manufacturer, upload one or more PDFs, and click <strong>Extract &amp; Add</strong> — or click <strong>Load Multi-Manufacturer Sample</strong>.</p>
        </div>

        <div id="tableContainer" class="hidden overflow-x-auto border border-slate-200 rounded-xl">
          <table class="w-full text-left border-collapse text-sm">
            <thead id="tableHead" class="bg-slate-100 text-slate-700 text-xs uppercase tracking-wider border-b border-slate-200"></thead>
            <tbody id="tableBody" class="divide-y divide-slate-200 bg-white"></tbody>
          </table>
        </div>
      </div>

      <!-- Raw Text View -->
      <div id="rawView" class="hidden p-5 sm:p-6">
        <div class="flex flex-wrap items-center justify-between gap-2 mb-3">
          <p class="text-xs text-slate-500">Review or edit raw OCR text below, then click <strong>Sync to Table</strong> to re-parse with your selected fields.</p>
          <div class="flex items-center gap-2">
            <button id="syncTextBtn" type="button" class="inline-flex items-center gap-1.5 text-xs font-semibold text-indigo-700 bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 px-3 py-1.5 rounded-lg transition">
              <i data-lucide="refresh-cw" class="w-3.5 h-3.5"></i>
              Sync to Table
            </button>
            <button id="copyTextBtn" type="button" class="inline-flex items-center gap-1.5 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 border border-slate-200 px-3 py-1.5 rounded-lg transition">
              <i data-lucide="copy" class="w-3.5 h-3.5"></i>
              Copy Text
            </button>
          </div>
        </div>
        <textarea id="rawTextArea" rows="12" placeholder="Extracted OCR text from uploaded manufacturer PDFs will appear here..." class="w-full font-mono text-xs sm:text-sm bg-slate-50 border border-slate-300 rounded-xl p-4 text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500"></textarea>
      </div>
    </section>
  </main>

  <!-- Notepad Local Setup Modal -->
  <div id="notepadModal" class="hidden fixed inset-0 z-50 bg-slate-900/50 flex items-center justify-center p-4">
    <div class="bg-white rounded-2xl border border-slate-200 shadow-xl max-w-lg w-full p-6 relative">
      <div class="flex items-center justify-between gap-3 mb-4">
        <div class="flex items-center gap-2">
          <i data-lucide="file-code-2" class="w-5 h-5 text-indigo-600"></i>
          <h3 class="text-base font-bold text-slate-900">How to Run Locally (Notepad Method)</h3>
        </div>
        <button id="closeGuideBtn" type="button" aria-label="Close modal" class="p-1.5 rounded-lg text-slate-500 hover:bg-slate-100 hover:text-slate-800 transition">
          <i data-lucide="x" class="w-5 h-5"></i>
        </button>
      </div>
      <ol class="space-y-2.5 text-sm text-slate-700 list-decimal list-inside">
        <li>Download this HTML file (or copy its entire source code into <strong>Notepad</strong>).</li>
        <li>Click <strong>File &rarr; Save As</strong>, set <em>Save as type</em> to <strong>All Files (*.*)</strong>, and name it <code class="bg-slate-100 px-1.5 py-0.5 rounded text-indigo-700 font-mono text-xs">ocr-collector.html</code>.</li>
        <li>Double-click <code class="bg-slate-100 px-1.5 py-0.5 rounded text-indigo-700 font-mono text-xs">ocr-collector.html</code> to open it in Chrome, Edge, or Firefox.</li>
        <li>Keep an active internet connection on first run so the browser loads PDF.js, Tesseract.js, SheetJS, and jsPDF from CDN.</li>
      </ol>
      <div class="mt-5 flex justify-end">
        <button id="closeGuideFooterBtn" type="button" class="bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-semibold px-4 py-2 rounded-lg transition">
          Got it
        </button>
      </div>
    </div>
  </div>

  <!-- Toast Container -->
  <div id="toastContainer" class="fixed bottom-4 right-4 z-50 flex flex-col gap-2 max-w-sm w-full pointer-events-none px-4 sm:px-0"></div>

  <script>
    (function () {
      'use strict';

      const PDF_WORKER_URL = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
      let pdfWorkerBlobUrl = null;

      // Bulletproof PDF.js worker loader (solves file:// & iframe CORS SecurityError)
      async function preparePdfWorker() {
        if (!window.pdfjsLib) return false;
        if (pdfWorkerBlobUrl) {
          window.pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerBlobUrl;
          return true;
        }
        try {
          const resp = await fetch(PDF_WORKER_URL);
          if (resp.ok) {
            const code = await resp.text();
            const blob = new Blob([code], { type: 'application/javascript' });
            pdfWorkerBlobUrl = URL.createObjectURL(blob);
            window.pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerBlobUrl;
            return true;
          }
        } catch (_err) {
          // Will fall back to main-thread fake worker below
        }
        window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER_URL;
        return false;
      }

      // Open PDF safely with cloned Uint8Array to prevent detached ArrayBuffer crash
      async function loadPdfDocumentSafely(arrayBuffer) {
        const blobReady = await preparePdfWorker();
        if (blobReady) {
          try {
            return await window.pdfjsLib.getDocument({
              data: new Uint8Array(arrayBuffer.slice(0)),
              isEvalSupported: false,
              useSystemFonts: true
            }).promise;
          } catch (firstErr) {
            if (firstErr && firstErr.name === 'PasswordException') throw firstErr;
          }
        }
        // Fallback: disableWorker on a fresh buffer clone
        return await window.pdfjsLib.getDocument({
          data: new Uint8Array(arrayBuffer.slice(0)),
          disableWorker: true,
          isEvalSupported: false,
          useSystemFonts: true
        }).promise;
      }

      // Bulletproof Tesseract.js v5 Worker Creator (SIMD -> LSTM fallback, Blob Worker safe)
      async function createOcrWorkerSafely(lang, onProgress) {
        const baseOpts = {
          workerPath: 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/worker.min.js',
          langPath: 'https://tessdata.projectnaptha.com/4.0.0',
          workerBlobURL: true,
          logger: function (m) {
            if (onProgress) onProgress(m);
          }
        };
        try {
          return await window.Tesseract.createWorker(lang, 1, Object.assign({}, baseOpts, {
            corePath: 'https://cdn.jsdelivr.net/npm/tesseract.js-core@5.1.0/tesseract-core-simd-lstm.wasm.js'
          }));
        } catch (_simdErr) {
          return await window.Tesseract.createWorker(lang, 1, Object.assign({}, baseOpts, {
            corePath: 'https://cdn.jsdelivr.net/npm/tesseract.js-core@5.1.0/tesseract-core-lstm.wasm.js'
          }));
        }
      }

      // Image Preprocessing for Scanned Datasheets (Grayscale + High-Contrast Stretch)
      function preprocessCanvasForOcr(sourceCanvas, enhance) {
        const off = document.createElement('canvas');
        off.width = sourceCanvas.width;
        off.height = sourceCanvas.height;
        const ctx = off.getContext('2d');
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, off.width, off.height);
        ctx.drawImage(sourceCanvas, 0, 0);

        if (enhance) {
          try {
            const imgData = ctx.getImageData(0, 0, off.width, off.height);
            const data = imgData.data;
            const contrast = 1.45;
            const intercept = 128 * (1 - contrast);
            for (let i = 0; i < data.length; i += 4) {
              const gray = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
              let val = gray * contrast + intercept;
              // Soft threshold to drop light watermarks while preserving fine print
              if (val > 215) val = 255;
              else if (val < 65) val = 0;
              else val = Math.max(0, Math.min(255, val));
              data[i] = val;
              data[i + 1] = val;
              data[i + 2] = val;
            }
            ctx.putImageData(imgData, 0, 0);
          } catch (_e) {
            // Skip pixel manipulation if tainted
          }
        }
        return off.toDataURL('image/png');
      }

      // Parse Page Range string (e.g. "1-3, 5") -> sorted unique 1-based page numbers
      function parsePageRange(inputStr, totalPages) {
        const cleaned = (inputStr || '').trim();
        if (!cleaned) {
          const all = [];
          for (let i = 1; i <= totalPages; i++) all.push(i);
          return all;
        }
        const set = new Set();
        cleaned.split(',').forEach(function (part) {
          const p = part.trim();
          if (!p) return;
          if (p.indexOf('-') !== -1) {
            const bounds = p.split('-').map(function (n) { return parseInt(n.trim(), 10); });
            if (!isNaN(bounds[0]) && !isNaN(bounds[1])) {
              const start = Math.max(1, Math.min(bounds[0], bounds[1]));
              const end = Math.min(totalPages, Math.max(bounds[0], bounds[1]));
              for (let i = start; i <= end; i++) set.add(i);
            }
          } else {
            const num = parseInt(p, 10);
            if (!isNaN(num) && num >= 1 && num <= totalPages) set.add(num);
          }
        });
        const list = Array.from(set).sort(function (a, b) { return a - b; });
        if (list.length === 0) {
          const fallback = [];
          for (let i = 1; i <= totalPages; i++) fallback.push(i);
          return fallback;
        }
        return list;
      }

      const DEFAULT_MANUFACTURERS = [
        'Siemens',
        'Schneider Electric',
        'Mehru Electricals',
        'ABB',
        'L&T Electrical',
        'Crompton Greaves',
        'BHEL',
        'Other'
      ];

      const PRESETS = {
        electrical: [
          { name: 'Model / Part No.', keywords: 'model, part no, type, catalog, code, order no, designation' },
          { name: 'Rated Voltage', keywords: 'voltage, rated voltage, kv, system voltage, highest voltage' },
          { name: 'Rated Current', keywords: 'current, rated current, amp, amps, primary current, ratio' },
          { name: 'Power / Rating', keywords: 'power, kva, kw, mva, burden, va, rating, breaking capacity' },
          { name: 'Frequency', keywords: 'frequency, hz, 50 hz, 60 hz' },
          { name: 'Accuracy / Standard', keywords: 'accuracy, class, standard, iec, is, ip rating, efficiency' },
          { name: 'Remarks', keywords: 'remarks, note, insulation, mounting, cooling, sealing' }
        ],
        commercial: [
          { name: 'Item / Model', keywords: 'item, model, part no, catalog, code' },
          { name: 'Description', keywords: 'description, particulars, specification, product' },
          { name: 'Quantity', keywords: 'qty, quantity, nos, units' },
          { name: 'Unit Price', keywords: 'unit price, rate, price, inr, usd' },
          { name: 'Total Amount', keywords: 'total, amount, net, grand total' },
          { name: 'Delivery / Terms', keywords: 'delivery, lead time, warranty, validity, payment' }
        ]
      };

      function loadSavedList(key, fallback) {
        try {
          const raw = window.localStorage.getItem(key);
          if (raw) {
            const parsed = JSON.parse(raw);
            if (Array.isArray(parsed) && parsed.length > 0) return parsed;
          }
        } catch (_e) {}
        return fallback.slice();
      }

      function saveList(key, val) {
        try {
          window.localStorage.setItem(key, JSON.stringify(val));
        } catch (_e) {}
      }

      const state = {
        manufacturers: loadSavedList('docuscan_manufacturers', DEFAULT_MANUFACTURERS),
        customFields: loadSavedList('docuscan_fields', PRESETS.electrical),
        files: [],
        documents: [],
        pagesCount: 0,
        rawText: '',
        headers: [],
        rows: [],
        filterManufacturer: 'ALL',
        searchQuery: '',
        isProcessing: false,
        cancelRequested: false,
        activeTab: 'table',
        schemaVisible: true
      };

      state.headers = ['Manufacturer', 'Equipment', 'Source File', 'Confidence'].concat(state.customFields.map(function (f) { return f.name; }));

      const els = {
        errorBanner: document.getElementById('errorBanner'),
        errorBannerTitle: document.getElementById('errorBannerTitle'),
        errorBannerDetail: document.getElementById('errorBannerDetail'),
        dismissErrorBtn: document.getElementById('dismissErrorBtn'),
        projectFileInput: document.getElementById('projectFileInput'),
        loadProjectBtn: document.getElementById('loadProjectBtn'),
        saveProjectBtn: document.getElementById('saveProjectBtn'),
        toggleSchemaBtn: document.getElementById('toggleSchemaBtn'),
        toggleSchemaLabel: document.getElementById('toggleSchemaLabel'),
        schemaPanel: document.getElementById('schemaPanel'),
        presetElectricalBtn: document.getElementById('presetElectricalBtn'),
        presetCommercialBtn: document.getElementById('presetCommercialBtn'),
        presetGeneralBtn: document.getElementById('presetGeneralBtn'),
        schemaFieldsList: document.getElementById('schemaFieldsList'),
        newFieldNameInput: document.getElementById('newFieldNameInput'),
        newFieldKeywordsInput: document.getElementById('newFieldKeywordsInput'),
        addFieldBtn: document.getElementById('addFieldBtn'),
        reapplySchemaBtn: document.getElementById('reapplySchemaBtn'),
        manufacturerSelect: document.getElementById('manufacturerSelect'),
        newManufacturerInput: document.getElementById('newManufacturerInput'),
        addManufacturerBtn: document.getElementById('addManufacturerBtn'),
        equipmentInput: document.getElementById('equipmentInput'),
        collectionModeSelect: document.getElementById('collectionModeSelect'),
        dpiScaleSelect: document.getElementById('dpiScaleSelect'),
        pageRangeInput: document.getElementById('pageRangeInput'),
        enhanceScanCheckbox: document.getElementById('enhanceScanCheckbox'),
        dropZone: document.getElementById('dropZone'),
        fileInput: document.getElementById('fileInput'),
        selectedFileBadge: document.getElementById('selectedFileBadge'),
        selectedFileName: document.getElementById('selectedFileName'),
        selectedFileSize: document.getElementById('selectedFileSize'),
        ocrModeSelect: document.getElementById('ocrModeSelect'),
        ocrLangSelect: document.getElementById('ocrLangSelect'),
        extractBtn: document.getElementById('extractBtn'),
        clearBtn: document.getElementById('clearBtn'),
        loadSampleBtn: document.getElementById('loadSampleBtn'),
        progressPanel: document.getElementById('progressPanel'),
        progressStatus: document.getElementById('progressStatus'),
        progressPercent: document.getElementById('progressPercent'),
        progressBar: document.getElementById('progressBar'),
        progressDetail: document.getElementById('progressDetail'),
        cancelOcrBtn: document.getElementById('cancelOcrBtn'),
        previewCanvasWrap: document.getElementById('previewCanvasWrap'),
        pagePreviewCanvas: document.getElementById('pagePreviewCanvas'),
        exportReadyBadge: document.getElementById('exportReadyBadge'),
        exportExcelBtn: document.getElementById('exportExcelBtn'),
        exportPdfBtn: document.getElementById('exportPdfBtn'),
        includeRawInPdfCheckbox: document.getElementById('includeRawInPdfCheckbox'),
        tabTableBtn: document.getElementById('tabTableBtn'),
        tabRawBtn: document.getElementById('tabRawBtn'),
        tableView: document.getElementById('tableView'),
        rawView: document.getElementById('rawView'),
        parseModeSelect: document.getElementById('parseModeSelect'),
        filterManufacturerSelect: document.getElementById('filterManufacturerSelect'),
        tableSearchInput: document.getElementById('tableSearchInput'),
        addRowBtn: document.getElementById('addRowBtn'),
        tableEmptyState: document.getElementById('tableEmptyState'),
        tableContainer: document.getElementById('tableContainer'),
        tableHead: document.getElementById('tableHead'),
        tableBody: document.getElementById('tableBody'),
        rawTextArea: document.getElementById('rawTextArea'),
        syncTextBtn: document.getElementById('syncTextBtn'),
        copyTextBtn: document.getElementById('copyTextBtn'),
        statPages: document.getElementById('statPages'),
        statRows: document.getElementById('statRows'),
        statWords: document.getElementById('statWords'),
        statMethod: document.getElementById('statMethod'),
        openGuideBtn: document.getElementById('openGuideBtn'),
        notepadModal: document.getElementById('notepadModal'),
        closeGuideBtn: document.getElementById('closeGuideBtn'),
        closeGuideFooterBtn: document.getElementById('closeGuideFooterBtn'),
        toastContainer: document.getElementById('toastContainer')
      };

      function showErrorBanner(title, detail) {
        els.errorBannerTitle.textContent = title || 'Extraction Error';
        els.errorBannerDetail.textContent = detail || 'An unexpected error occurred while processing the PDF.';
        els.errorBanner.classList.remove('hidden');
      }

      function hideErrorBanner() {
        els.errorBanner.classList.add('hidden');
      }

      function showToast(message, type) {
        const colors = {
          success: 'bg-emerald-600 text-white',
          error: 'bg-rose-600 text-white',
          info: 'bg-slate-800 text-white'
        };
        const toast = document.createElement('div');
        toast.className = 'toast-enter pointer-events-auto rounded-xl px-4 py-3 text-xs sm:text-sm font-medium shadow-lg flex items-center justify-between gap-3 ' + (colors[type] || colors.info);
        const span = document.createElement('span');
        span.textContent = message;
        toast.appendChild(span);
        els.toastContainer.appendChild(toast);
        setTimeout(function () {
          toast.style.opacity = '0';
          toast.style.transition = 'opacity 0.25s ease';
          setTimeout(function () {
            if (toast.parentNode) toast.parentNode.removeChild(toast);
          }, 260);
        }, 3800);
      }

      function formatBytes(bytes) {
        if (!bytes) return '0 B';
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), sizes.length - 1);
        return (bytes / Math.pow(1024, i)).toFixed(i === 0 ? 0 : 1) + ' ' + sizes[i];
      }

      function setProgress(percent, status, detail) {
        const clamped = Math.max(0, Math.min(100, Math.round(percent)));
        els.progressBar.style.width = clamped + '%';
        els.progressPercent.textContent = clamped + '%';
        if (status) els.progressStatus.textContent = status;
        if (detail) els.progressDetail.textContent = detail;
      }

      function renderManufacturers() {
        const current = els.manufacturerSelect.value || state.manufacturers[0];
        els.manufacturerSelect.innerHTML = '';
        state.manufacturers.forEach(function (m) {
          const opt = document.createElement('option');
          opt.value = m;
          opt.textContent = m;
          els.manufacturerSelect.appendChild(opt);
        });
        if (state.manufacturers.indexOf(current) !== -1) {
          els.manufacturerSelect.value = current;
        }
        renderManufacturerFilter();
      }

      function renderManufacturerFilter() {
        const prev = state.filterManufacturer;
        const present = Array.from(new Set(state.rows.map(function (r) { return r[0] || ''; }).filter(Boolean)));
        els.filterManufacturerSelect.innerHTML = '';
        const optAll = document.createElement('option');
        optAll.value = 'ALL';
        optAll.textContent = 'All Manufacturers (' + state.rows.length + ')';
        els.filterManufacturerSelect.appendChild(optAll);

        present.forEach(function (m) {
          const count = state.rows.filter(function (r) { return r[0] === m; }).length;
          const opt = document.createElement('option');
          opt.value = m;
          opt.textContent = m + ' (' + count + ')';
          els.filterManufacturerSelect.appendChild(opt);
        });

        if (prev === 'ALL' || present.indexOf(prev) !== -1) {
          els.filterManufacturerSelect.value = prev;
        } else {
          state.filterManufacturer = 'ALL';
          els.filterManufacturerSelect.value = 'ALL';
        }
      }

      function renderSchemaFields() {
        els.schemaFieldsList.innerHTML = '';
        state.customFields.forEach(function (field, idx) {
          const card = document.createElement('div');
          card.className = 'flex items-center gap-2 bg-slate-50 border border-slate-200 rounded-lg p-2.5';

          const inputsWrap = document.createElement('div');
          inputsWrap.className = 'flex-1 grid grid-cols-1 sm:grid-cols-2 gap-2 min-w-0';

          const nameIn = document.createElement('input');
          nameIn.type = 'text';
          nameIn.value = field.name;
          nameIn.placeholder = 'Column Name';
          nameIn.className = 'text-xs font-semibold bg-white border border-slate-200 rounded-md px-2.5 py-1.5 text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500';
          nameIn.addEventListener('input', function (e) {
            state.customFields[idx].name = e.target.value;
            saveList('docuscan_fields', state.customFields);
          });

          const kwIn = document.createElement('input');
          kwIn.type = 'text';
          kwIn.value = field.keywords;
          kwIn.placeholder = 'Keywords (comma-separated)';
          kwIn.className = 'text-xs bg-white border border-slate-200 rounded-md px-2.5 py-1.5 text-slate-600 focus:outline-none focus:ring-2 focus:ring-indigo-500';
          kwIn.addEventListener('input', function (e) {
            state.customFields[idx].keywords = e.target.value;
            saveList('docuscan_fields', state.customFields);
          });

          inputsWrap.appendChild(nameIn);
          inputsWrap.appendChild(kwIn);

          const delBtn = document.createElement('button');
          delBtn.type = 'button';
          delBtn.title = 'Remove field';
          delBtn.className = 'p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-md transition shrink-0';
          delBtn.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 6 6 18"/><path d="m6 6 12 12"/></svg>';
          delBtn.addEventListener('click', function () {
            if (state.customFields.length <= 1) {
              showToast('Keep at least one custom field.', 'info');
              return;
            }
            state.customFields.splice(idx, 1);
            saveList('docuscan_fields', state.customFields);
            renderSchemaFields();
          });

          card.appendChild(inputsWrap);
          card.appendChild(delBtn);
          els.schemaFieldsList.appendChild(card);
        });
      }

      // Improved Industry Keyword Matcher (supports Key: Value, multi-space columns, and next-line values)
      function extractCustomSchemaRows(docObj) {
        const lines = docObj.text
          .split(/\r?\n/)
          .map(function (l) { return l.trim(); })
          .filter(function (l) { return l.length > 0 && !/^---\s*Page\s+\d+.*---$/i.test(l); });

        const values = state.customFields.map(function (field) {
          const kws = field.keywords
            .split(',')
            .map(function (k) { return k.trim().toLowerCase(); })
            .filter(Boolean);
          if (kws.length === 0) kws.push(field.name.toLowerCase());

          for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            const lower = line.toLowerCase();
            for (let k = 0; k < kws.length; k++) {
              const kw = kws[k];
              if (lower.indexOf(kw) !== -1) {
                // Pattern 1: "Label : Value" or "Label = Value" or "Label | Value"
                const kvMatch = line.match(/^[^:=|]{1,50}\s*[:=|]\s*(.+)$/);
                if (kvMatch && kvMatch[1].trim()) {
                  return kvMatch[1].trim();
                }
                // Pattern 2: "Label    Value" (multi-space tabular separation)
                const cols = line.split(/\s{2,}/).map(function (s) { return s.trim(); }).filter(Boolean);
                if (cols.length >= 2) {
                  return cols.slice(1).join(' | ');
                }
                // Pattern 3: Strip keyword prefix; if nothing remains, grab next line if concise
                const escaped = kw.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                const cleaned = line.replace(new RegExp(escaped, 'i'), '').replace(/^[\s:\-–—.]+/, '').trim();
                if (cleaned) return cleaned;
                if (i + 1 < lines.length && lines[i + 1].length <= 80) {
                  return lines[i + 1];
                }
                return line;
              }
            }
          }
          return '';
        });

        const confStr = typeof docObj.confidence === 'number' ? (docObj.confidence + '% (' + docObj.method + ')') : docObj.method;
        return [[docObj.manufacturer, docObj.equipment || 'General Equipment', docObj.fileName, confStr].concat(values)];
      }

      function parseDocumentWithMode(docObj, mode) {
        const confStr = typeof docObj.confidence === 'number' ? (docObj.confidence + '% (' + docObj.method + ')') : docObj.method;
        if (mode === 'custom') {
          return {
            headers: ['Manufacturer', 'Equipment', 'Source File', 'Confidence'].concat(state.customFields.map(function (f) { return f.name; })),
            rows: extractCustomSchemaRows(docObj)
          };
        }

        const lines = docObj.text
          .split(/\r?\n/)
          .map(function (l) { return l.trim(); })
          .filter(function (l) { return l.length > 0 && !/^---\s*Page\s+\d+.*---$/i.test(l); });

        if (lines.length === 0) {
          return { headers: ['Manufacturer', 'Equipment', 'Source File', 'Confidence', 'Extracted Data'], rows: [] };
        }

        if (mode === 'lines') {
          return {
            headers: ['Manufacturer', 'Equipment', 'Source File', 'Confidence', 'Line #', 'Extracted Content'],
            rows: lines.map(function (line, idx) {
              return [docObj.manufacturer, docObj.equipment || 'General', docObj.fileName, confStr, String(idx + 1), line];
            })
          };
        }

        if (mode === 'kv') {
          const kvRows = lines.map(function (line) {
            const match = line.match(/^([^:=|]{2,45})\s*[:=]\s*(.+)$/);
            if (match) return [docObj.manufacturer, docObj.equipment || 'General', docObj.fileName, confStr, match[1].trim(), match[2].trim()];
            return [docObj.manufacturer, docObj.equipment || 'General', docObj.fileName, confStr, 'Note', line];
          });
          return { headers: ['Manufacturer', 'Equipment', 'Source File', 'Confidence', 'Field / Key', 'Extracted Value'], rows: kvRows };
        }

        const tokenized = lines.map(function (line) {
          if (line.indexOf('|') !== -1) return line.split('|').map(function (s) { return s.trim(); }).filter(Boolean);
          if (line.indexOf('\t') !== -1) return line.split('\t').map(function (s) { return s.trim(); });
          const multiSpace = line.split(/\s{2,}/).map(function (s) { return s.trim(); }).filter(Boolean);
          if (multiSpace.length > 1) return multiSpace;
          const kv = line.match(/^([^:]{2,40}):\s*(.+)$/);
          if (kv) return [kv[1].trim(), kv[2].trim()];
          return [line];
        });

        let maxCols = 1;
        tokenized.forEach(function (r) { if (r.length > maxCols) maxCols = r.length; });
        maxCols = Math.min(maxCols, 7);

        const firstRow = tokenized[0];
        const firstIsHeader = maxCols > 1 && firstRow.length >= 2 && firstRow.every(function (cell) {
          return isNaN(Number(cell.replace(/[$€£₹,]/g, '')));
        });

        const baseHeaders = ['Manufacturer', 'Equipment', 'Source File', 'Confidence'];
        let dataRows = tokenized;
        if (firstIsHeader && tokenized.length > 1) {
          for (let c = 0; c < maxCols; c++) baseHeaders.push(firstRow[c] || ('Column ' + (c + 1)));
          dataRows = tokenized.slice(1);
        } else {
          for (let c = 0; c < maxCols; c++) baseHeaders.push('Column ' + (c + 1));
        }

        const normalizedRows = dataRows.map(function (r) {
          const out = [docObj.manufacturer, docObj.equipment || 'General', docObj.fileName, confStr];
          for (let c = 0; c < maxCols; c++) out.push(r[c] !== undefined ? r[c] : '');
          return out;
        });

        return { headers: baseHeaders, rows: normalizedRows };
      }

      function rebuildMasterTableFromDocuments() {
        const mode = els.parseModeSelect.value;
        if (state.documents.length === 0) {
          state.rows = [];
          renderAll();
          return;
        }

        let masterHeaders = [];
        let masterRows = [];

        state.documents.forEach(function (docObj) {
          const parsed = parseDocumentWithMode(docObj, mode);
          if (parsed.headers.length > masterHeaders.length) {
            masterHeaders = parsed.headers.slice();
          }
          parsed.rows.forEach(function (r) { masterRows.push(r); });
        });

        masterRows = masterRows.map(function (r) {
          const copy = r.slice();
          while (copy.length < masterHeaders.length) copy.push('');
          return copy.slice(0, masterHeaders.length);
        });

        state.headers = masterHeaders;
        state.rows = masterRows;
        state.rawText = state.documents.map(function (d) {
          return '=== [' + d.manufacturer + ' | ' + (d.equipment || 'General') + ' | ' + d.fileName + '] ===\n' + d.text;
        }).join('\n\n');
        renderAll();
      }

      function handleSelectedFiles(fileList) {
        if (!fileList || fileList.length === 0) return;
        const valid = [];
        for (let i = 0; i < fileList.length; i++) {
          const f = fileList[i];
          const isPdf = f.type === 'application/pdf' || f.name.toLowerCase().endsWith('.pdf');
          const isImg = f.type.startsWith('image/');
          if (isPdf || isImg) valid.push(f);
        }
        if (valid.length === 0) {
          showToast('Please select PDF or image files (PNG, JPG, WebP).', 'error');
          return;
        }
        hideErrorBanner();
        state.files = valid;
        const totalSize = valid.reduce(function (acc, f) { return acc + f.size; }, 0);
        els.selectedFileName.textContent = valid.length === 1 ? valid[0].name : (valid.length + ' files queued (' + valid[0].name + '...)');
        els.selectedFileSize.textContent = '(' + formatBytes(totalSize) + ')';
        els.selectedFileBadge.classList.remove('hidden');
        els.extractBtn.disabled = false;
        showToast(valid.length + ' file(s) ready to extract.', 'info');
      }

      function getFilteredRowIndices() {
        const q = (state.searchQuery || '').trim().toLowerCase();
        const indices = [];
        state.rows.forEach(function (row, idx) {
          if (state.filterManufacturer !== 'ALL' && row[0] !== state.filterManufacturer) return;
          if (q) {
            const match = row.some(function (cell) {
              return String(cell || '').toLowerCase().indexOf(q) !== -1;
            });
            if (!match) return;
          }
          indices.push(idx);
        });
        return indices;
      }

      function renderAll() {
        const hasData = state.rows.length > 0 || state.rawText.trim().length > 0;
        els.exportExcelBtn.disabled = !hasData || state.isProcessing;
        els.exportPdfBtn.disabled = !hasData || state.isProcessing;

        renderManufacturerFilter();

        const uniqueMfrs = Array.from(new Set(state.rows.map(function (r) { return r[0]; }).filter(Boolean)));
        if (hasData) {
          els.exportReadyBadge.textContent = state.rows.length + ' rows across ' + (uniqueMfrs.length || 1) + ' manufacturer(s) ready';
          els.exportReadyBadge.className = 'text-xs font-semibold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2.5 py-1 rounded-full';
        } else {
          els.exportReadyBadge.textContent = 'Upload a PDF or load sample data to export';
          els.exportReadyBadge.className = 'text-xs font-medium text-slate-500 bg-slate-100 px-2.5 py-1 rounded-full';
        }

        els.statPages.textContent = String(state.pagesCount || (hasData ? state.documents.length || 1 : 0));
        els.statRows.textContent = String(state.rows.length);
        const words = state.rawText.trim() ? state.rawText.trim().split(/\s+/).length : 0;
        els.statWords.textContent = words.toLocaleString();
        els.statMethod.textContent = uniqueMfrs.length > 0 ? String(uniqueMfrs.length) : '—';

        if (els.rawTextArea.value !== state.rawText) {
          els.rawTextArea.value = state.rawText;
        }

        const visibleIndices = getFilteredRowIndices();
        if (visibleIndices.length === 0) {
          els.tableEmptyState.classList.remove('hidden');
          els.tableContainer.classList.add('hidden');
          return;
        }

        els.tableEmptyState.classList.add('hidden');
        els.tableContainer.classList.remove('hidden');

        els.tableHead.innerHTML = '';
        const trHead = document.createElement('tr');
        const thIdx = document.createElement('th');
        thIdx.className = 'px-3 py-2.5 w-10 text-center font-semibold';
        thIdx.textContent = '#';
        trHead.appendChild(thIdx);

        state.headers.forEach(function (h, colIdx) {
          const th = document.createElement('th');
          th.className = 'px-3 py-2.5 font-semibold min-w-[125px]';
          const input = document.createElement('input');
          input.type = 'text';
          input.value = h;
          input.setAttribute('aria-label', 'Column header ' + (colIdx + 1));
          input.className = 'w-full bg-transparent font-semibold uppercase tracking-wider text-xs text-slate-700 focus:outline-none focus:bg-white focus:px-1.5 rounded';
          input.addEventListener('input', function (e) {
            state.headers[colIdx] = e.target.value;
          });
          th.appendChild(input);
          trHead.appendChild(th);
        });

        const thAct = document.createElement('th');
        thAct.className = 'px-3 py-2.5 w-12 text-center font-semibold';
        trHead.appendChild(thAct);
        els.tableHead.appendChild(trHead);

        els.tableBody.innerHTML = '';
        visibleIndices.forEach(function (rowIdx, displayIdx) {
          const row = state.rows[rowIdx];
          const tr = document.createElement('tr');
          tr.className = 'hover:bg-slate-50/80 transition';

          const tdIdx = document.createElement('td');
          tdIdx.className = 'px-3 py-2 text-center text-xs text-slate-400 font-medium';
          tdIdx.textContent = String(displayIdx + 1);
          tr.appendChild(tdIdx);

          state.headers.forEach(function (_h, colIdx) {
            const td = document.createElement('td');
            td.className = 'p-0';
            const input = document.createElement('input');
            input.type = 'text';
            input.value = row[colIdx] !== undefined ? row[colIdx] : '';
            input.setAttribute('aria-label', 'Row ' + (displayIdx + 1) + ' Column ' + (colIdx + 1));
            let textStyle = 'text-slate-800';
            if (colIdx === 0) textStyle = 'font-semibold text-indigo-700';
            else if (colIdx === 3) textStyle = 'text-xs font-medium text-emerald-700';
            input.className = 'cell-input w-full px-3 py-2 text-xs sm:text-sm bg-transparent border-0 ' + textStyle;
            input.addEventListener('input', function (e) {
              state.rows[rowIdx][colIdx] = e.target.value;
            });
            td.appendChild(input);
            tr.appendChild(td);
          });

          const tdDel = document.createElement('td');
          tdDel.className = 'px-2 py-1.5 text-center';
          const delBtn = document.createElement('button');
          delBtn.type = 'button';
          delBtn.title = 'Delete row';
          delBtn.className = 'p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-md transition';
          delBtn.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h18"/><path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/><path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/></svg>';
          delBtn.addEventListener('click', function () {
            state.rows.splice(rowIdx, 1);
            renderAll();
          });
          tdDel.appendChild(delBtn);
          tr.appendChild(tdDel);

          els.tableBody.appendChild(tr);
        });
      }

      function extractStructuredPdfText(textContent) {
        const items = textContent.items || [];
        if (items.length === 0) return '';
        const linesMap = [];
        items.forEach(function (item) {
          const str = (item.str || '').trim();
          if (!str) return;
          const y = Math.round(item.transform[5] / 4) * 4;
          const x = item.transform[4];
          let lineObj = linesMap.find(function (l) { return Math.abs(l.y - y) <= 4; });
          if (!lineObj) {
            lineObj = { y: y, parts: [] };
            linesMap.push(lineObj);
          }
          lineObj.parts.push({ x: x, str: str });
        });

        linesMap.sort(function (a, b) { return b.y - a.y; });
        return linesMap.map(function (line) {
          line.parts.sort(function (a, b) { return a.x - b.x; });
          let out = '';
          for (let i = 0; i < line.parts.length; i++) {
            if (i > 0) {
              const gap = line.parts[i].x - line.parts[i - 1].x;
              out += gap > 40 ? '   ' : ' ';
            }
            out += line.parts[i].str;
          }
          return out.trim();
        }).filter(Boolean).join('\n');
      }

      // Main Industrial Batch Extraction Pipeline
      async function runExtraction() {
        if (state.files.length === 0 || state.isProcessing) return;

        hideErrorBanner();
        const hasPdf = state.files.some(function (f) { return f.type === 'application/pdf' || f.name.toLowerCase().endsWith('.pdf'); });
        if (hasPdf && !window.pdfjsLib) {
          showErrorBanner('PDF Engine Unavailable', 'PDF.js could not be loaded from CDN. Please check your internet connection and reload.');
          return;
        }
        if (!window.Tesseract) {
          showErrorBanner('OCR Engine Unavailable', 'Tesseract.js could not be loaded from CDN. Please check your internet connection and reload.');
          return;
        }

        state.isProcessing = true;
        state.cancelRequested = false;
        els.extractBtn.disabled = true;
        els.exportExcelBtn.disabled = true;
        els.exportPdfBtn.disabled = true;
        els.progressPanel.classList.remove('hidden');

        const mode = els.ocrModeSelect.value;
        const lang = els.ocrLangSelect.value;
        const dpiScale = parseFloat(els.dpiScaleSelect.value) || 2.8;
        const enhance = els.enhanceScanCheckbox.checked;
        const manufacturer = els.manufacturerSelect.value || 'Other';
        const equipment = els.equipmentInput.value.trim() || 'Equipment';
        const replaceMode = els.collectionModeSelect.value === 'replace';

        let worker = null;
        let currentProgressDetail = '';
        const newDocs = [];
        let newPagesTotal = 0;
        const failedFiles = [];

        try {
          if (replaceMode) {
            state.documents = [];
            state.pagesCount = 0;
          }

          for (let fIdx = 0; fIdx < state.files.length; fIdx++) {
            if (state.cancelRequested) break;
            const file = state.files[fIdx];
            const isPdf = file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf');
            const filePrefix = '[' + (fIdx + 1) + '/' + state.files.length + '] ' + file.name;
            setProgress((fIdx / state.files.length) * 100 + 2, 'Reading ' + filePrefix, 'Loading document...');

            try {
              const extractedPages = [];

              if (!isPdf) {
                const imgDataUrl = await new Promise(function (resolve, reject) {
                  const reader = new FileReader();
                  reader.onload = function () { resolve(reader.result); };
                  reader.onerror = function () { reject(new Error('Could not read image file')); };
                  reader.readAsDataURL(file);
                });

                if (!worker) {
                  setProgress(12, 'Loading OCR Engine (' + lang + ')...', 'Initializing Tesseract WASM...');
                  worker = await createOcrWorkerSafely(lang, function (m) {
                    if (m.status === 'recognizing text' && typeof m.progress === 'number') {
                      setProgress(15 + m.progress * 80, 'Running OCR on ' + filePrefix, 'Recognizing text: ' + Math.round(m.progress * 100) + '%');
                    }
                  });
                }

                // Draw image to canvas so preprocessing applies uniformly
                const imgEl = await new Promise(function (resolve, reject) {
                  const im = new Image();
                  im.onload = function () { resolve(im); };
                  im.onerror = function () { reject(new Error('Invalid image file')); };
                  im.src = imgDataUrl;
                });
                const canvas = els.pagePreviewCanvas;
                canvas.width = imgEl.naturalWidth || imgEl.width;
                canvas.height = imgEl.naturalHeight || imgEl.height;
                canvas.getContext('2d').drawImage(imgEl, 0, 0);
                els.previewCanvasWrap.classList.remove('hidden');

                const preparedPng = preprocessCanvasForOcr(canvas, enhance);
                const res = await worker.recognize(preparedPng);
                extractedPages.push({
                  pageNum: 1,
                  text: (res.data.text || '').trim(),
                  method: 'OCR',
                  confidence: Math.round(res.data.confidence || 0)
                });
              } else {
                const arrayBuffer = await file.arrayBuffer();
                const pdfDoc = await loadPdfDocumentSafely(arrayBuffer);
                const pageNumbers = parsePageRange(els.pageRangeInput.value, pdfDoc.numPages);
                els.previewCanvasWrap.classList.remove('hidden');

                for (let pIdx = 0; pIdx < pageNumbers.length; pIdx++) {
                  if (state.cancelRequested) break;
                  const pageNum = pageNumbers[pIdx];
                  const overallBase = ((fIdx + pIdx / pageNumbers.length) / state.files.length) * 100;
                  const slice = (1 / (state.files.length * pageNumbers.length)) * 100;
                  currentProgressDetail = filePrefix + ' — Page ' + pageNum + ' of ' + pdfDoc.numPages;
                  setProgress(overallBase + 1, 'Processing ' + currentProgressDetail, 'Inspecting PDF layer...');

                  const page = await pdfDoc.getPage(pageNum);
                  let pageText = '';
                  let usedMethod = 'Native Text';
                  let confidence = 100;

                  if (mode === 'hybrid') {
                    try {
                      const textContent = await page.getTextContent();
                      pageText = extractStructuredPdfText(textContent);
                    } catch (_tcErr) {
                      pageText = '';
                    }
                  }

                  const needsOcr = mode === 'force_ocr' || pageText.replace(/\s+/g, '').length < 25;
                  const renderScale = needsOcr ? dpiScale : 1.0;
                  const viewport = page.getViewport({ scale: renderScale });
                  const canvas = els.pagePreviewCanvas;
                  const ctx = canvas.getContext('2d');
                  canvas.width = Math.max(1, Math.floor(viewport.width));
                  canvas.height = Math.max(1, Math.floor(viewport.height));
                  await page.render({ canvasContext: ctx, viewport: viewport }).promise;

                  if (needsOcr) {
                    usedMethod = 'OCR';
                    if (!worker) {
                      setProgress(overallBase + 2, 'Loading OCR Engine (' + lang + ')...', 'Downloading language data (cached after first use)...');
                      worker = await createOcrWorkerSafely(lang, function (m) {
                        if (m.status === 'recognizing text' && typeof m.progress === 'number') {
                          setProgress(
                            overallBase + m.progress * slice,
                            'OCR Scanning: ' + currentProgressDetail,
                            'Recognizing characters: ' + Math.round(m.progress * 100) + '%'
                          );
                        }
                      });
                    }
                    const preparedPng = preprocessCanvasForOcr(canvas, enhance);
                    const ocrRes = await worker.recognize(preparedPng);
                    pageText = (ocrRes.data.text || '').trim();
                    confidence = Math.round(ocrRes.data.confidence || 0);
                  }

                  extractedPages.push({
                    pageNum: pageNum,
                    text: pageText,
                    method: usedMethod,
                    confidence: confidence
                  });
                }
              }

              if (!state.cancelRequested && extractedPages.length > 0) {
                const avgConf = Math.round(
                  extractedPages.reduce(function (s, p) { return s + p.confidence; }, 0) / extractedPages.length
                );
                const combinedText = extractedPages.map(function (p) {
                  return '--- Page ' + p.pageNum + ' (' + p.method + ', ' + p.confidence + '%) ---\n' + (p.text || '(No text detected)');
                }).join('\n\n');

                newDocs.push({
                  id: Date.now() + fIdx,
                  manufacturer: manufacturer,
                  equipment: equipment,
                  fileName: file.name,
                  method: Array.from(new Set(extractedPages.map(function (p) { return p.method; }))).join('+'),
                  confidence: avgConf,
                  pages: extractedPages.length,
                  text: combinedText
                });
                newPagesTotal += extractedPages.length;
              }
            } catch (fileErr) {
              console.error('Error processing ' + file.name, fileErr);
              let reason = fileErr && fileErr.message ? fileErr.message : 'Unknown error';
              if (fileErr && fileErr.name === 'PasswordException') {
                reason = 'PDF is password-protected. Please unlock it first.';
              }
              failedFiles.push(file.name + ': ' + reason);
            }
          }

          if (worker) {
            await worker.terminate();
            worker = null;
          }

          if (newDocs.length > 0) {
            state.documents = state.documents.concat(newDocs);
            state.pagesCount += newPagesTotal;
            rebuildMasterTableFromDocuments();
          }

          if (failedFiles.length > 0) {
            showErrorBanner(
              'Some files could not be extracted (' + failedFiles.length + ' failed)',
              failedFiles.join(' | ')
            );
          }

          if (state.cancelRequested) {
            showToast('Extraction cancelled.', 'info');
          } else if (newDocs.length > 0) {
            setProgress(100, 'Extraction complete!', 'Collected ' + newDocs.length + ' file(s) (' + newPagesTotal + ' pages).');
            showToast('Extracted ' + newDocs.length + ' file(s) for ' + manufacturer + '!', 'success');
          }
        } catch (err) {
          console.error(err);
          showErrorBanner('Extraction Failed', err && err.message ? err.message : 'Unexpected OCR error. Check internet connection for CDN models.');
        } finally {
          if (worker) {
            try { await worker.terminate(); } catch (_e) {}
          }
          state.isProcessing = false;
          els.extractBtn.disabled = state.files.length === 0;
          setTimeout(function () {
            if (!state.isProcessing) els.progressPanel.classList.add('hidden');
          }, 1200);
          renderAll();
        }
      }

      // Option 1: Export to Excel (.xlsx)
      function exportToExcel() {
        if (!window.XLSX) {
          showToast('SheetJS library not loaded. Check your internet connection.', 'error');
          return;
        }
        if (state.rows.length === 0 && !state.rawText.trim()) {
          showToast('No extracted data to export.', 'error');
          return;
        }

        try {
          const wb = window.XLSX.utils.book_new();
          const exportRows = getFilteredRowIndices().map(function (i) { return state.rows[i]; });

          const tableData = [state.headers].concat(exportRows);
          const wsTable = window.XLSX.utils.aoa_to_sheet(tableData);
          wsTable['!cols'] = state.headers.map(function (h, idx) {
            let maxLen = String(h || '').length;
            exportRows.forEach(function (r) {
              const len = String(r[idx] || '').length;
              if (len > maxLen) maxLen = len;
            });
            return { wch: Math.min(Math.max(maxLen + 4, 14), 48) };
          });
          window.XLSX.utils.book_append_sheet(wb, wsTable, 'Master Collection');

          const summaryMap = {};
          exportRows.forEach(function (r) {
            const mfr = r[0] || 'Unspecified';
            const eq = r[1] || 'General';
            const key = mfr + '||' + eq;
            summaryMap[key] = (summaryMap[key] || 0) + 1;
          });
          const summaryData = [['Manufacturer', 'Equipment Category', 'Collected Rows']];
          Object.keys(summaryMap).forEach(function (k) {
            const parts = k.split('||');
            summaryData.push([parts[0], parts[1], summaryMap[k]]);
          });
          const wsSummary = window.XLSX.utils.aoa_to_sheet(summaryData);
          wsSummary['!cols'] = [{ wch: 24 }, { wch: 28 }, { wch: 16 }];
          window.XLSX.utils.book_append_sheet(wb, wsSummary, 'Manufacturer Summary');

          const rawRows = [['Manufacturer', 'Equipment', 'Source File', 'Confidence', 'Extracted Text']];
          if (state.documents.length > 0) {
            state.documents.forEach(function (d) {
              rawRows.push([d.manufacturer, d.equipment, d.fileName, (d.confidence || 0) + '%', d.text]);
            });
          } else {
            rawRows.push(['All', 'Edited', 'Manual', '-', state.rawText]);
          }
          const wsRaw = window.XLSX.utils.aoa_to_sheet(rawRows);
          wsRaw['!cols'] = [{ wch: 20 }, { wch: 22 }, { wch: 26 }, { wch: 12 }, { wch: 90 }];
          window.XLSX.utils.book_append_sheet(wb, wsRaw, 'Raw OCR Text');

          const filename = 'Manufacturer_Equipment_Data_' + new Date().toISOString().slice(0, 10) + '.xlsx';
          window.XLSX.writeFile(wb, filename);
          showToast('Excel workbook downloaded: ' + filename, 'success');
        } catch (err) {
          console.error(err);
          showToast('Excel export failed: ' + err.message, 'error');
        }
      }

      // Option 2: Export to PDF (.pdf)
      function exportToPdf() {
        if (!window.jspdf || !window.jspdf.jsPDF) {
          showToast('jsPDF library not loaded. Check your internet connection.', 'error');
          return;
        }
        if (state.rows.length === 0 && !state.rawText.trim()) {
          showToast('No extracted data to export.', 'error');
          return;
        }

        try {
          const jsPDF = window.jspdf.jsPDF;
          const exportRows = getFilteredRowIndices().map(function (i) { return state.rows[i]; });
          const doc = new jsPDF({ orientation: state.headers.length > 5 ? 'landscape' : 'portrait', unit: 'pt', format: 'a4' });
          const margin = 36;

          doc.setFillColor(79, 70, 229);
          doc.rect(0, 0, doc.internal.pageSize.getWidth(), 64, 'F');
          doc.setTextColor(255, 255, 255);
          doc.setFont('helvetica', 'bold');
          doc.setFontSize(15);
          doc.text('Multi-Manufacturer Equipment Data Report', margin, 30);
          doc.setFont('helvetica', 'normal');
          doc.setFontSize(9);
          const uniqueMfrs = Array.from(new Set(exportRows.map(function (r) { return r[0]; }).filter(Boolean)));
          const metaLine = 'Manufacturers: ' + (uniqueMfrs.join(', ') || 'All') + '   |   Generated: ' + new Date().toLocaleString();
          doc.text(metaLine, margin, 48);

          let cursorY = 84;

          if (exportRows.length > 0 && typeof doc.autoTable === 'function') {
            doc.setTextColor(15, 23, 42);
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(11);
            doc.text('Collected Equipment Comparison (' + exportRows.length + ' rows)', margin, cursorY);

            doc.autoTable({
              startY: cursorY + 10,
              head: [state.headers],
              body: exportRows,
              margin: { left: margin, right: margin },
              styles: { fontSize: 7.5, cellPadding: 4, textColor: [30, 41, 59], overflow: 'linebreak' },
              headStyles: { fillColor: [79, 70, 229], textColor: [255, 255, 255], fontStyle: 'bold' },
              alternateRowStyles: { fillColor: [248, 250, 252] }
            });
            cursorY = (doc.lastAutoTable ? doc.lastAutoTable.finalY : cursorY) + 24;
          }

          if (els.includeRawInPdfCheckbox.checked && state.rawText.trim()) {
            const pageHeight = doc.internal.pageSize.getHeight();
            const pageWidth = doc.internal.pageSize.getWidth();
            if (cursorY > pageHeight - 100) {
              doc.addPage();
              cursorY = 44;
            }
            doc.setTextColor(15, 23, 42);
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(11);
            doc.text('Raw Extracted OCR Text Appendix', margin, cursorY);
            cursorY += 15;

            doc.setFont('courier', 'normal');
            doc.setFontSize(8);
            doc.setTextColor(51, 65, 85);
            const wrapped = doc.splitTextToSize(state.rawText.trim(), pageWidth - margin * 2);
            wrapped.forEach(function (line) {
              if (cursorY > pageHeight - 36) {
                doc.addPage();
                cursorY = 40;
              }
              doc.text(line, margin, cursorY);
              cursorY += 11;
            });
          }

          const filename = 'Manufacturer_Equipment_Report_' + new Date().toISOString().slice(0, 10) + '.pdf';
          doc.save(filename);
          showToast('PDF report downloaded: ' + filename, 'success');
        } catch (err) {
          console.error(err);
          showToast('PDF export failed: ' + err.message, 'error');
        }
      }

      // Save / Load JSON Project Session
      function saveProjectJson() {
        const payload = {
          version: 2,
          exportedAt: new Date().toISOString(),
          manufacturers: state.manufacturers,
          customFields: state.customFields,
          documents: state.documents,
          pagesCount: state.pagesCount,
          headers: state.headers,
          rows: state.rows,
          rawText: state.rawText
        };
        const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'DocuScan_Project_' + new Date().toISOString().slice(0, 10) + '.json';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        showToast('Project JSON saved.', 'success');
      }

      function loadProjectJson(file) {
        if (!file) return;
        const reader = new FileReader();
        reader.onload = function () {
          try {
            const data = JSON.parse(reader.result);
            if (Array.isArray(data.manufacturers)) {
              state.manufacturers = data.manufacturers;
              saveList('docuscan_manufacturers', state.manufacturers);
            }
            if (Array.isArray(data.customFields)) {
              state.customFields = data.customFields;
              saveList('docuscan_fields', state.customFields);
            }
            state.documents = Array.isArray(data.documents) ? data.documents : [];
            state.pagesCount = Number(data.pagesCount) || state.documents.length;
            state.headers = Array.isArray(data.headers) ? data.headers : state.headers;
            state.rows = Array.isArray(data.rows) ? data.rows : [];
            state.rawText = typeof data.rawText === 'string' ? data.rawText : '';
            renderManufacturers();
            renderSchemaFields();
            renderAll();
            showToast('Project loaded: ' + state.rows.length + ' rows restored.', 'success');
          } catch (_e) {
            showToast('Invalid project JSON file.', 'error');
          }
        };
        reader.readAsText(file);
      }

      function loadSampleData() {
        hideErrorBanner();
        state.customFields = PRESETS.electrical.map(function (f) { return { name: f.name, keywords: f.keywords }; });
        saveList('docuscan_fields', state.customFields);
        renderSchemaFields();
        els.parseModeSelect.value = 'custom';

        state.documents = [
          {
            id: 1,
            manufacturer: 'Siemens',
            equipment: '3-Phase Induction Motor',
            fileName: 'Siemens_1LE1_Datasheet.pdf',
            method: 'Sample OCR',
            confidence: 98,
            pages: 1,
            text: [
              '--- Page 1 (Sample OCR, 98%) ---',
              'Model: SIMOTICS GP 1LE1001-1CB2',
              'Rated Voltage: 415 V AC (3-Phase)',
              'Rated Current: 14.2 A',
              'Power: 7.5 kW (10 HP)',
              'Frequency: 50 Hz',
              'Standard: IE3 / IEC 60034-30-1 / IP55',
              'Insulation: Class F, Foot Mounted B3'
            ].join('\n')
          },
          {
            id: 2,
            manufacturer: 'Schneider Electric',
            equipment: 'Air Circuit Breaker (ACB)',
            fileName: 'Schneider_MasterPact_MTZ1.pdf',
            method: 'Sample OCR',
            confidence: 97,
            pages: 1,
            text: [
              '--- Page 1 (Sample OCR, 97%) ---',
              'Model: MasterPact MTZ1 10 H1',
              'Rated Voltage: 690 V AC',
              'Rated Current: 1000 A',
              'Rating: 42 kA Breaking Capacity',
              'Frequency: 50/60 Hz',
              'Standard: IEC 60947-2 / MicroLogic 5.0 X',
              'Remarks: Drawout 3P with MicroLogic trip unit'
            ].join('\n')
          },
          {
            id: 3,
            manufacturer: 'Mehru Electricals',
            equipment: 'Current Transformer (CT)',
            fileName: 'Mehru_33kV_Outdoor_CT.pdf',
            method: 'Sample OCR',
            confidence: 96,
            pages: 1,
            text: [
              '--- Page 1 (Sample OCR, 96%) ---',
              'Model: ME-CT-33-400/1A',
              'System Voltage: 33 kV (Highest 36 kV)',
              'Primary Current: 400-200 / 1-1 A',
              'Burden: 15 VA (Core 1 & Core 2)',
              'Frequency: 50 Hz',
              'Accuracy: Class 0.2S / 5P20 (IS 2705 / IEC 61869)',
              'Insulation: Oil-Filled Hermetically Sealed'
            ].join('\n')
          }
        ];

        state.pagesCount = 3;
        state.files = [];
        els.selectedFileBadge.classList.add('hidden');
        els.extractBtn.disabled = true;
        rebuildMasterTableFromDocuments();
        showToast('Loaded sample data for Siemens, Schneider Electric & Mehru Electricals!', 'success');
      }

      function resetAll() {
        if (state.isProcessing) state.cancelRequested = true;
        hideErrorBanner();
        state.files = [];
        state.documents = [];
        state.pagesCount = 0;
        state.rawText = '';
        state.rows = [];
        state.filterManufacturer = 'ALL';
        state.searchQuery = '';
        els.fileInput.value = '';
        els.tableSearchInput.value = '';
        els.selectedFileBadge.classList.add('hidden');
        els.extractBtn.disabled = true;
        els.progressPanel.classList.add('hidden');
        els.previewCanvasWrap.classList.add('hidden');
        els.rawTextArea.value = '';
        renderAll();
        showToast('All collected data cleared.', 'info');
      }

      function switchTab(tab) {
        state.activeTab = tab;
        if (tab === 'table') {
          els.tableView.classList.remove('hidden');
          els.rawView.classList.add('hidden');
          els.tabTableBtn.className = 'px-3 py-1.5 text-xs sm:text-sm font-semibold rounded-md bg-white text-indigo-700 shadow-sm transition';
          els.tabRawBtn.className = 'px-3 py-1.5 text-xs sm:text-sm font-medium rounded-md text-slate-600 hover:text-slate-900 transition';
        } else {
          els.tableView.classList.add('hidden');
          els.rawView.classList.remove('hidden');
          els.tabRawBtn.className = 'px-3 py-1.5 text-xs sm:text-sm font-semibold rounded-md bg-white text-indigo-700 shadow-sm transition';
          els.tabTableBtn.className = 'px-3 py-1.5 text-xs sm:text-sm font-medium rounded-md text-slate-600 hover:text-slate-900 transition';
        }
      }

      // Event Listeners
      els.dismissErrorBtn.addEventListener('click', hideErrorBanner);
      els.saveProjectBtn.addEventListener('click', saveProjectJson);
      els.loadProjectBtn.addEventListener('click', function () { els.projectFileInput.click(); });
      els.projectFileInput.addEventListener('change', function (e) {
        if (e.target.files && e.target.files[0]) {
          loadProjectJson(e.target.files[0]);
          e.target.value = '';
        }
      });

      els.toggleSchemaBtn.addEventListener('click', function () {
        state.schemaVisible = !state.schemaVisible;
        if (state.schemaVisible) {
          els.schemaPanel.classList.remove('hidden');
          els.toggleSchemaLabel.textContent = 'Hide Field Builder';
        } else {
          els.schemaPanel.classList.add('hidden');
          els.toggleSchemaLabel.textContent = 'Show Field Builder';
        }
      });

      els.presetElectricalBtn.addEventListener('click', function () {
        state.customFields = PRESETS.electrical.map(function (f) { return { name: f.name, keywords: f.keywords }; });
        saveList('docuscan_fields', state.customFields);
        els.parseModeSelect.value = 'custom';
        renderSchemaFields();
        if (state.documents.length > 0) rebuildMasterTableFromDocuments();
        showToast('Applied Electrical Equipment Specs preset.', 'info');
      });

      els.presetCommercialBtn.addEventListener('click', function () {
        state.customFields = PRESETS.commercial.map(function (f) { return { name: f.name, keywords: f.keywords }; });
        saveList('docuscan_fields', state.customFields);
        els.parseModeSelect.value = 'custom';
        renderSchemaFields();
        if (state.documents.length > 0) rebuildMasterTableFromDocuments();
        showToast('Applied Quotation / Commercial preset.', 'info');
      });

      els.presetGeneralBtn.addEventListener('click', function () {
        els.parseModeSelect.value = 'auto';
        if (state.documents.length > 0) rebuildMasterTableFromDocuments();
        showToast('Switched to Auto-Detect Table mode.', 'info');
      });

      els.addFieldBtn.addEventListener('click', function () {
        const name = els.newFieldNameInput.value.trim();
        const kws = els.newFieldKeywordsInput.value.trim();
        if (!name) {
          showToast('Enter a column name first.', 'error');
          return;
        }
        state.customFields.push({ name: name, keywords: kws || name.toLowerCase() });
        saveList('docuscan_fields', state.customFields);
        els.newFieldNameInput.value = '';
        els.newFieldKeywordsInput.value = '';
        renderSchemaFields();
        showToast('Added field "' + name + '". Click Re-apply to update existing rows.', 'success');
      });

      els.reapplySchemaBtn.addEventListener('click', function () {
        if (state.documents.length === 0) {
          showToast('No extracted PDFs yet. Upload a PDF or load sample data first.', 'info');
          return;
        }
        rebuildMasterTableFromDocuments();
        showToast('Master table updated with your tailored fields.', 'success');
      });

      els.addManufacturerBtn.addEventListener('click', function () {
        const name = els.newManufacturerInput.value.trim();
        if (!name) {
          showToast('Enter a manufacturer name to add.', 'error');
          return;
        }
        if (state.manufacturers.indexOf(name) === -1) {
          state.manufacturers.unshift(name);
          saveList('docuscan_manufacturers', state.manufacturers);
        }
        renderManufacturers();
        els.manufacturerSelect.value = name;
        els.newManufacturerInput.value = '';
        showToast('Manufacturer added & selected: ' + name, 'success');
      });

      els.dropZone.addEventListener('click', function () { els.fileInput.click(); });
      els.dropZone.addEventListener('keydown', function (e) {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          els.fileInput.click();
        }
      });
      els.fileInput.addEventListener('change', function (e) {
        if (e.target.files && e.target.files.length > 0) handleSelectedFiles(e.target.files);
      });

      ['dragenter', 'dragover'].forEach(function (evt) {
        els.dropZone.addEventListener(evt, function (e) {
          e.preventDefault();
          e.stopPropagation();
          els.dropZone.classList.add('drop-zone-active');
        });
      });
      ['dragleave', 'drop'].forEach(function (evt) {
        els.dropZone.addEventListener(evt, function (e) {
          e.preventDefault();
          e.stopPropagation();
          els.dropZone.classList.remove('drop-zone-active');
        });
      });
      els.dropZone.addEventListener('drop', function (e) {
        if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
          handleSelectedFiles(e.dataTransfer.files);
        }
      });

      els.extractBtn.addEventListener('click', runExtraction);
      els.clearBtn.addEventListener('click', resetAll);
      els.loadSampleBtn.addEventListener('click', loadSampleData);
      els.cancelOcrBtn.addEventListener('click', function () {
        state.cancelRequested = true;
        els.progressStatus.textContent = 'Cancelling...';
      });

      els.exportExcelBtn.addEventListener('click', exportToExcel);
      els.exportPdfBtn.addEventListener('click', exportToPdf);

      els.tabTableBtn.addEventListener('click', function () { switchTab('table'); });
      els.tabRawBtn.addEventListener('click', function () { switchTab('raw'); });

      els.parseModeSelect.addEventListener('change', function () {
        if (state.documents.length > 0) rebuildMasterTableFromDocuments();
      });

      els.filterManufacturerSelect.addEventListener('change', function () {
        state.filterManufacturer = els.filterManufacturerSelect.value;
        renderAll();
      });

      els.tableSearchInput.addEventListener('input', function (e) {
        state.searchQuery = e.target.value;
        renderAll();
      });

      els.addRowBtn.addEventListener('click', function () {
        if (state.headers.length === 0) {
          state.headers = ['Manufacturer', 'Equipment', 'Source File', 'Confidence'].concat(state.customFields.map(function (f) { return f.name; }));
        }
        const newRow = state.headers.map(function (_h, idx) {
          if (idx === 0) return els.manufacturerSelect.value || 'Other';
          if (idx === 1) return els.equipmentInput.value.trim() || 'Equipment';
          if (idx === 2) return 'Manual Entry';
          if (idx === 3) return 'Manual';
          return '';
        });
        state.rows.push(newRow);
        renderAll();
      });

      els.syncTextBtn.addEventListener('click', function () {
        const text = els.rawTextArea.value;
        state.rawText = text;
        if (state.documents.length === 1) {
          state.documents[0].text = text;
          rebuildMasterTableFromDocuments();
        } else {
          const docObj = {
            id: Date.now(),
            manufacturer: els.manufacturerSelect.value || 'Edited',
            equipment: els.equipmentInput.value.trim() || 'General',
            fileName: 'Edited_Text',
            method: 'Manual',
            confidence: 100,
            pages: 1,
            text: text
          };
          const parsed = parseDocumentWithMode(docObj, els.parseModeSelect.value);
          state.headers = parsed.headers;
          state.rows = parsed.rows;
          renderAll();
        }
        switchTab('table');
        showToast('Master table updated from raw text.', 'success');
      });

      els.copyTextBtn.addEventListener('click', function () {
        const text = els.rawTextArea.value;
        if (!text.trim()) {
          showToast('Nothing to copy yet.', 'info');
          return;
        }
        if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(text).then(function () {
            showToast('Extracted text copied to clipboard.', 'success');
          }).catch(function () {
            els.rawTextArea.select();
            document.execCommand('copy');
            showToast('Extracted text copied to clipboard.', 'success');
          });
        } else {
          els.rawTextArea.select();
          document.execCommand('copy');
          showToast('Extracted text copied to clipboard.', 'success');
        }
      });

      function toggleModal(show) {
        if (show) els.notepadModal.classList.remove('hidden');
        else els.notepadModal.classList.add('hidden');
      }
      els.openGuideBtn.addEventListener('click', function () { toggleModal(true); });
      els.closeGuideBtn.addEventListener('click', function () { toggleModal(false); });
      els.closeGuideFooterBtn.addEventListener('click', function () { toggleModal(false); });
      els.notepadModal.addEventListener('click', function (e) {
        if (e.target === els.notepadModal) toggleModal(false);
      });

      if (window.lucide) {
        window.lucide.createIcons();
      }
      renderManufacturers();
      renderSchemaFields();
      renderAll();
    })();
  </script>
</body>
</html>
