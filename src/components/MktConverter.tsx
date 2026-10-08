import { useState } from "react";
import { AlertTriangle, CheckCircle2, Download, FileSpreadsheet, Loader2, Upload, Wrench } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { buildWorkbook, parseResponses, type Correction, type ParsedResponses } from "@/lib/mkt-converter";

const iso = (date: Date) => date.toISOString().slice(0, 10);

function FileDrop({ label, hint, file, onFile }: { label: string; hint: string; file: File | null; onFile: (file: File) => void }) {
  return <label className="flex cursor-pointer items-center gap-4 rounded-xl border-2 border-dashed border-slate-300 bg-white p-5 transition hover:border-teal-600"><div className="rounded-lg bg-teal-50 p-3 text-teal-700">{file ? <FileSpreadsheet /> : <Upload />}</div><div className="min-w-0"><p className="font-medium">{label}</p><p className="truncate text-sm text-slate-500">{file?.name ?? hint}</p></div><input className="hidden" type="file" accept=".xlsx" onChange={(event) => { const selected = event.target.files?.[0]; if (selected) onFile(selected); }} /></label>;
}

export function MktConverter() {
  const [responseFile, setResponseFile] = useState<File | null>(null);
  const [workbookFile, setWorkbookFile] = useState<File | null>(null);
  const [parsed, setParsed] = useState<ParsedResponses | null>(null);
  const [start, setStart] = useState("");
  const [result, setResult] = useState<{ url: string; name: string; filled: number; missing: string[]; corrections: Correction[]; notes: string[] } | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  async function readResponses(file: File) {
    setResponseFile(file); setResult(null); setError(""); setBusy(true);
    try {
      const data = await parseResponses(await file.arrayBuffer());
      setParsed(data);
      if (data.maxDate) {
        const date = new Date(data.maxDate);
        date.setUTCDate(date.getUTCDate() - 6);
        date.setUTCDate(date.getUTCDate() - ((date.getUTCDay() + 6) % 7));
        setStart(iso(date));
      }
    } catch { setError("Could not read the response workbook."); }
    setBusy(false);
  }

  async function convert() {
    if (!parsed || !start || !workbookFile) return;
    setBusy(true); setError("");
    try {
      const data = await buildWorkbook(parsed, new Date(`${start}T00:00:00Z`), await workbookFile.arrayBuffer());
      if (result) URL.revokeObjectURL(result.url);
      setResult({ ...data, name: data.fileName, url: URL.createObjectURL(data.blob) });
    } catch { setError("Conversion failed. Check that the MKT workbook contains the Summary sheet and weekly template."); }
    setBusy(false);
  }

  const end = start ? iso(new Date(Date.parse(start) + 6 * 864e5)) : "";
  return <div className="mx-auto max-w-3xl"><div className="mb-7"><p className="text-sm text-slate-500">Temperature quality workspace</p><h2 className="mt-1 text-3xl font-semibold">Room temperature to MKT</h2><p className="mt-2 text-slate-600">Upload both workbooks to create the weekly monitoring sheet with corrected readings, formulas, and a Summary column.</p></div><div className="space-y-4"><FileDrop label="1. ULC CCT room temperature responses" hint="Required Google Form export (.xlsx)" file={responseFile} onFile={(file) => void readResponses(file)} /><FileDrop label="2. Existing MKT monitoring workbook" hint="Required workbook with Summary and weekly sheets" file={workbookFile} onFile={(file) => { setWorkbookFile(file); setResult(null); }} />{parsed && <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm"><p className="text-sm text-slate-500">{parsed.rows.toLocaleString()} readings found · {parsed.minDate && iso(parsed.minDate)} to {parsed.maxDate && iso(parsed.maxDate)}</p>{parsed.corrections.length > 0 && <p className="mt-2 text-sm font-medium text-amber-700">{parsed.corrections.length} values need automatic correction and will be listed after conversion.</p>}<div className="mt-5 flex flex-wrap items-end gap-4"><div><Label htmlFor="mkt-start">Week starts</Label><Input className="mt-2 w-44" id="mkt-start" type="date" value={start} onChange={(event) => { setStart(event.target.value); setResult(null); }} /></div><span className="pb-2 font-mono text-sm text-slate-500">→ {end} (7 days)</span></div><Button className="mt-5 w-full" size="lg" disabled={busy || !start || !workbookFile} onClick={() => void convert()}>{busy ? <Loader2 className="animate-spin" /> : <FileSpreadsheet />} Convert workbook</Button></div>}{busy && !parsed && <p className="flex items-center gap-2 text-sm text-slate-500"><Loader2 className="h-4 w-4 animate-spin" /> Reading workbook…</p>}{error && <p className="flex items-center gap-2 rounded-lg bg-red-50 p-3 text-sm text-red-800"><AlertTriangle className="h-4 w-4" />{error}</p>}{result && <div className="rounded-2xl border border-teal-200 bg-teal-50 p-6"><p className="flex items-center gap-2 font-medium text-teal-950"><CheckCircle2 className="h-5 w-5" /> Ready: {result.filled} of 168 hours filled</p>{result.notes.map((note) => <p className="mt-2 text-sm text-teal-900" key={note}>{note}</p>)}{result.corrections.length > 0 && <details className="mt-4 text-sm text-teal-950"><summary className="flex cursor-pointer items-center gap-2 font-medium"><Wrench className="h-4 w-4" /> Review {result.corrections.length} corrections</summary><div className="mt-3 max-h-56 overflow-auto rounded-lg bg-white p-3 font-mono text-xs"><table className="w-full"><thead><tr className="border-b text-left"><th className="py-2">Date / time</th><th className="py-2">Sensor</th><th className="py-2 text-right">Original</th><th className="py-2 text-right">Result</th></tr></thead><tbody>{result.corrections.filter((correction) => correction.date >= start && correction.date <= end).map((correction) => <tr className="border-b border-slate-100" key={`${correction.date}-${correction.hour}-${correction.target}`}><td className="py-2">{correction.date} {correction.hour}</td><td className="py-2">{correction.label}</td><td className="py-2 text-right text-red-700">{correction.raw}</td><td className="py-2 text-right">{correction.cleaned === null ? "blank" : correction.cleaned}</td></tr>)}</tbody></table></div></details>}<Button asChild className="mt-5 w-full" size="lg"><a download={result.name} href={result.url}><Download /> Download {result.name}</a></Button></div>}</div></div>;
}
