import { useState, useRef } from "react";

type EntryType = "inward" | "outward";
type CapacityMode = "manual" | "carryforward";
type AppState = "idle" | "processing" | "done" | "error";

export default function App() {
  const [files, setFiles] = useState<File[]>([]);
  const [entryType, setEntryType] = useState<EntryType>("inward");
  const [capacityMode, setCapacityMode] = useState<CapacityMode>("manual");
  const [openingCapacity, setOpeningCapacity] = useState("");
  const [appState, setAppState] = useState<AppState>("idle");
  const [error, setError] = useState("");
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFiles = (selected: FileList | null) => {
    if (!selected) return;
    setFiles(Array.from(selected));
    setDownloadUrl(null);
    setError("");
    setAppState("idle");
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    handleFiles(e.dataTransfer.files);
  };

  const handleSubmit = async () => {
    if (files.length === 0) {
      setError("Upload at least one image.");
      return;
    }
    setAppState("processing");
    setError("");
    setDownloadUrl(null);

    const formData = new FormData();
    for (const f of files) formData.append("images", f);
    formData.append("entryType", entryType);
    formData.append("openingCapacityMode", capacityMode);
    if (capacityMode === "manual" && openingCapacity) {
      formData.append("openingCapacityValue", openingCapacity);
    }

    try {
      const res = await fetch("/api/warehouse/process", {
        method: "POST",
        body: formData,
      });

      if (!res.ok) {
        const json = await res.json().catch(() => ({ error: "Unknown error" }));
        throw new Error(json.error || "Request failed");
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      setDownloadUrl(url);
      setAppState("done");
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : "Something went wrong";
      setError(msg);
      setAppState("error");
    }
  };

  const reset = () => {
    setFiles([]);
    setDownloadUrl(null);
    setError("");
    setAppState("idle");
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col items-center py-10 px-4">
      <div className="w-full max-w-lg">
        {/* Header */}
        <div className="mb-8 text-center">
          <div className="inline-flex items-center justify-center w-14 h-14 rounded-2xl bg-blue-600 mb-4">
            <svg className="w-7 h-7 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
          </div>
          <h1 className="text-2xl font-bold text-slate-800">Warehouse Register</h1>
          <p className="text-slate-500 text-sm mt-1">Upload images → Get Excel report</p>
        </div>

        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 space-y-6">
          {/* Upload */}
          <div>
            <label className="block text-sm font-semibold text-slate-700 mb-2">Upload Images</label>
            <div
              className="border-2 border-dashed border-slate-300 rounded-xl p-6 text-center cursor-pointer hover:border-blue-400 hover:bg-blue-50 transition-colors"
              onClick={() => fileInputRef.current?.click()}
              onDrop={handleDrop}
              onDragOver={(e) => e.preventDefault()}
            >
              <input
                ref={fileInputRef}
                type="file"
                multiple
                accept="image/*"
                className="hidden"
                onChange={(e) => handleFiles(e.target.files)}
              />
              {files.length === 0 ? (
                <div>
                  <svg className="w-10 h-10 text-slate-400 mx-auto mb-2" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z" />
                  </svg>
                  <p className="text-slate-500 text-sm">Tap to choose or drag images here</p>
                </div>
              ) : (
                <div>
                  <svg className="w-8 h-8 text-green-500 mx-auto mb-1" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                  </svg>
                  <p className="text-slate-700 font-medium text-sm">{files.length} image{files.length > 1 ? "s" : ""} selected</p>
                  <p className="text-slate-400 text-xs mt-0.5">{files.map(f => f.name).join(", ")}</p>
                  <button
                    className="text-blue-600 text-xs mt-2 underline"
                    onClick={(e) => { e.stopPropagation(); reset(); }}
                  >
                    Clear
                  </button>
                </div>
              )}
            </div>
          </div>

          {/* Entry Type */}
          <div>
            <label className="block text-sm font-semibold text-slate-700 mb-2">Entry Type</label>
            <div className="flex gap-3">
              {(["inward", "outward"] as EntryType[]).map((type) => (
                <label
                  key={type}
                  className={`flex-1 flex items-center justify-center gap-2 py-3 rounded-xl border-2 cursor-pointer font-medium text-sm transition-colors ${
                    entryType === type
                      ? "border-blue-600 bg-blue-50 text-blue-700"
                      : "border-slate-200 text-slate-600 hover:border-slate-300"
                  }`}
                >
                  <input
                    type="radio"
                    name="entryType"
                    value={type}
                    checked={entryType === type}
                    onChange={() => setEntryType(type)}
                    className="hidden"
                  />
                  <span className={`w-4 h-4 rounded-full border-2 flex items-center justify-center ${entryType === type ? "border-blue-600" : "border-slate-400"}`}>
                    {entryType === type && <span className="w-2 h-2 rounded-full bg-blue-600" />}
                  </span>
                  {type.charAt(0).toUpperCase() + type.slice(1)}
                </label>
              ))}
            </div>
          </div>

          {/* Opening Capacity Mode */}
          <div>
            <label className="block text-sm font-semibold text-slate-700 mb-2">Opening Capacity</label>
            <div className="flex gap-3">
              {([["manual", "Manual Entry"], ["carryforward", "Carry Forward"]] as [CapacityMode, string][]).map(([mode, label]) => (
                <label
                  key={mode}
                  className={`flex-1 flex items-center justify-center gap-2 py-3 rounded-xl border-2 cursor-pointer font-medium text-sm transition-colors ${
                    capacityMode === mode
                      ? "border-blue-600 bg-blue-50 text-blue-700"
                      : "border-slate-200 text-slate-600 hover:border-slate-300"
                  }`}
                >
                  <input
                    type="radio"
                    name="capacityMode"
                    value={mode}
                    checked={capacityMode === mode}
                    onChange={() => setCapacityMode(mode)}
                    className="hidden"
                  />
                  <span className={`w-4 h-4 rounded-full border-2 flex items-center justify-center ${capacityMode === mode ? "border-blue-600" : "border-slate-400"}`}>
                    {capacityMode === mode && <span className="w-2 h-2 rounded-full bg-blue-600" />}
                  </span>
                  {label}
                </label>
              ))}
            </div>

            {capacityMode === "manual" && (
              <div className="mt-3">
                <input
                  type="number"
                  step="0.001"
                  placeholder="e.g. 148.520"
                  value={openingCapacity}
                  onChange={(e) => setOpeningCapacity(e.target.value)}
                  className="w-full border border-slate-300 rounded-xl px-4 py-3 text-slate-800 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                />
              </div>
            )}
          </div>

          {/* Error */}
          {error && (
            <div className="bg-red-50 border border-red-200 rounded-xl px-4 py-3 text-red-700 text-sm">
              {error}
            </div>
          )}

          {/* Submit */}
          {appState !== "done" && (
            <button
              onClick={handleSubmit}
              disabled={appState === "processing"}
              className="w-full bg-blue-600 hover:bg-blue-700 disabled:bg-blue-400 text-white font-semibold py-3.5 rounded-xl transition-colors text-sm flex items-center justify-center gap-2"
            >
              {appState === "processing" ? (
                <>
                  <svg className="animate-spin w-4 h-4" fill="none" viewBox="0 0 24 24">
                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
                  </svg>
                  Extracting & Building Excel...
                </>
              ) : (
                <>
                  <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                  </svg>
                  Generate Report
                </>
              )}
            </button>
          )}

          {/* Download */}
          {appState === "done" && downloadUrl && (
            <div className="space-y-3">
              <div className="bg-green-50 border border-green-200 rounded-xl px-4 py-3 text-green-700 text-sm flex items-center gap-2">
                <svg className="w-4 h-4 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                </svg>
                Report ready! Download below.
              </div>
              <a
                href={downloadUrl}
                download="warehouse-register.xlsx"
                className="w-full bg-green-600 hover:bg-green-700 text-white font-semibold py-3.5 rounded-xl transition-colors text-sm flex items-center justify-center gap-2"
              >
                <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                </svg>
                Download Excel Report
              </a>
              <button
                onClick={reset}
                className="w-full text-slate-500 hover:text-slate-700 text-sm py-2"
              >
                Process another batch
              </button>
            </div>
          )}
        </div>

        <p className="text-center text-xs text-slate-400 mt-6">
          AI extracts data • Formulas preserved • One sheet per financial month
        </p>
      </div>
    </div>
  );
}
