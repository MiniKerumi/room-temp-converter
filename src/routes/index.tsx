import { createFileRoute } from "@tanstack/react-router";
import { useState } from "react";
import { FileSpreadsheet, Download, Upload, Loader as Loader2, CircleCheck as CheckCircle2, TriangleAlert as AlertTriangle, Wrench } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "ULC Temp → MKT Converter" },
      { name: "description", content: "Convert ULC CCT room temperature form responses into the weekly MKT monitoring workbook." },
      { property: "og:title", content: "ULC Temp → MKT Converter" },
      { property: "og:description", content: "Turn guard temperature logs into weekly MKT sheets with formulas." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary" },
    ],
  }),
  component: Index,
});

type Parsed = Awaited<ReturnType<typeof import("@/lib/mkt-converter").parseResponses>>;
type Correction = Parsed["corrections"][number];

const iso = (d: Date) => d.toISOString().slice(0, 10);

function FileDrop({ label, hint, file, onFile }: { label: string; hint: string; file: File | null; onFile: (f: File) => void }) {
  return (
    <label className="flex cursor-pointer items-center gap-4 rounded-lg border-2 border-dashed border-border bg-card p-5 transition-colors hover:border-primary">
      <div className="rounded-md bg-secondary p-3 text-primary">
        {file ? <FileSpreadsheet className="h-6 w-6" /> : <Upload className="h-6 w-6" />}
      </div>
      <div className="min-w-0">
        <div className="font-medium">{label}</div>
        <div className="truncate text-sm text-muted-foreground">{file ? file.name : hint}</div>
      </div>
      <input type="file" accept=".xlsx" className="hidden" onChange={(e) => e.target.files?.[0] && onFile(e.target.files[0])} />
    </label>
  );
}

function Index() {
  const [respFile, setRespFile] = useState<File | null>(null);
  const [mktFile, setMktFile] = useState<File | null>(null);
  const [parsed, setParsed] = useState<Parsed | null>(null);
  const [start, setStart] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState<{ url: string; name: string; filled: number; missing: string[]; corrections: Correction[]; notes: string[] } | null>(null);

  async function onResponses(f: File) {
    setRespFile(f); setError(""); setResult(null); setBusy(true);
    try {
      const { parseResponses } = await import("@/lib/mkt-converter");
      const p = await parseResponses(await f.arrayBuffer());
      setParsed(p);
      if (p.maxDate) {
        // default: most recent Monday-starting full week in the data
        const d = new Date(p.maxDate);
        d.setUTCDate(d.getUTCDate() - 6);
        d.setUTCDate(d.getUTCDate() - ((d.getUTCDay() + 6) % 7));
        setStart(iso(d));
      }
    } catch (e) {
      setError("Couldn't read that file. Make sure it's the ULC CCT Form responses sheet.");
      console.error(e);
    } finally { setBusy(false); }
  }

  async function convert() {
    if (!parsed || !start) return;
    if (!mktFile) { setError("Please upload the existing MKT Monitoring workbook — its formulas are needed for the conversion."); return; }
    setBusy(true); setError("");
    try {
      const { buildWorkbook } = await import("@/lib/mkt-converter");
      const res = await buildWorkbook(parsed, new Date(`${start}T00:00:00Z`), await mktFile.arrayBuffer());
      if (result) URL.revokeObjectURL(result.url);
      setResult({ url: URL.createObjectURL(res.blob), name: res.fileName, filled: res.filled, missing: res.missing, corrections: res.corrections, notes: res.notes });
    } catch (e) {
      setError(`Conversion failed: ${e instanceof Error ? e.message : "check that the MKT workbook is the right file."}`);
      console.error(e);
    } finally { setBusy(false); }
  }

  const endLabel = start ? iso(new Date(Date.parse(start) + 6 * 864e5)) : "";

  return (
    <main className="min-h-screen bg-background px-4 py-12">
      <div className="mx-auto max-w-2xl">
        <p className="font-mono text-xs uppercase tracking-widest text-primary">ULC · Merck</p>
        <h1 className="mt-2 text-3xl font-semibold tracking-tight">Room Temp → MKT Converter</h1>
        <p className="mt-2 text-muted-foreground">
          Upload the guard's form responses, pick the week, and download the MKT workbook with MIN / MAX / AVERAGE / MKT formulas filled in.
        </p>

        <div className="mt-8 space-y-4">
          <FileDrop label="1. ULC CCT Room Temperature Responses" hint="Required — the Google Form export (.xlsx)" file={respFile} onFile={onResponses} />
          <FileDrop label="2. Existing MKT Monitoring workbook" hint="Required — must contain the Summary sheet and weekly MKT formulas" file={mktFile} onFile={(f) => { setMktFile(f); setResult(null); }} />

          {parsed && (
            <div className="rounded-lg border border-border bg-card p-5">
              <div className="text-sm text-muted-foreground">
                {parsed.rows.toLocaleString()} readings found · {parsed.minDate && iso(parsed.minDate)} to {parsed.maxDate && iso(parsed.maxDate)}
              </div>
              <div className="mt-4 flex flex-wrap items-end gap-4">
                <div className="space-y-1.5">
                  <Label htmlFor="start">3. Week starts</Label>
                  <Input id="start" type="date" value={start} onChange={(e) => { setStart(e.target.value); setResult(null); }} className="w-44" />
                </div>
                <div className="pb-2 font-mono text-sm text-muted-foreground">→ {endLabel} (7 days)</div>
              </div>
              <Button className="mt-5 w-full" size="lg" onClick={convert} disabled={busy || !start || !mktFile}>
                {busy ? <Loader2 className="animate-spin" /> : <FileSpreadsheet />} Convert
              </Button>
            </div>
          )}

          {busy && !parsed && <div className="flex items-center gap-2 text-sm text-muted-foreground"><Loader2 className="h-4 w-4 animate-spin" /> Reading file…</div>}
          {error && <div className="flex items-center gap-2 rounded-md bg-destructive/10 p-3 text-sm text-destructive"><AlertTriangle className="h-4 w-4" />{error}</div>}

          {result && (
            <div className="rounded-lg border border-primary/30 bg-secondary p-5">
              <div className="flex items-center gap-2 font-medium"><CheckCircle2 className="h-5 w-5 text-primary" /> Ready: {result.filled} of 168 hours filled</div>
              {result.notes.map((n) => <div key={n} className="mt-2 text-sm text-muted-foreground">{n}</div>)}
              {result.corrections.length > 0 && (
                <details className="mt-3 text-sm text-muted-foreground">
                  <summary className="cursor-pointer flex items-center gap-1.5"><Wrench className="h-4 w-4 text-accent" /> {result.corrections.length} values were auto-corrected this week (highlighted yellow in the sheet)</summary>
                  <div className="mt-2 max-h-48 overflow-auto font-mono text-xs">
                    <table className="w-full border-collapse">
                      <thead>
                        <tr className="border-b text-left">
                          <th className="py-1 pr-3 font-medium">Date</th>
                          <th className="py-1 pr-3 font-medium">Time</th>
                          <th className="py-1 pr-3 font-medium">Sensor</th>
                          <th className="py-1 pr-3 font-medium text-right">Original</th>
                          <th className="py-1 font-medium text-right">Corrected</th>
                        </tr>
                      </thead>
                      <tbody>
                        {result.corrections.map((c, idx) => (
                          <tr key={idx} className="border-b border-border/50">
                            <td className="py-1 pr-3">{c.date}</td>
                            <td className="py-1 pr-3">{c.hour}</td>
                            <td className="py-1 pr-3 max-w-[12rem] truncate">{c.label}</td>
                            <td className="py-1 pr-3 text-right text-destructive">{c.raw}</td>
                            <td className="py-1 text-right text-primary">{c.cleaned === null ? "(blank)" : c.cleaned}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </details>
              )}
              {result.missing.length > 0 && (
                <details className="mt-2 text-sm text-muted-foreground">
                  <summary className="cursor-pointer">{result.missing.length} hours have no reading (left blank)</summary>
                  <div className="mt-2 max-h-40 overflow-auto font-mono text-xs">{result.missing.join(" · ")}</div>
                </details>
              )}
              <Button asChild className="mt-4 w-full" size="lg">
                <a href={result.url} download={result.name}><Download /> Download {result.name}</a>
              </Button>
            </div>
          )}
        </div>

        <p className="mt-10 text-xs text-muted-foreground">
          Files are processed in your browser and never uploaded. If several entries exist for the same hour, the one submitted closest to that hour is used. Chiller uses the NEW Chiller WH1 column; Bio Ref uses BioRef 1. Values outside the expected range are auto-corrected: a missing decimal point is reinserted (e.g. 300 → 30.0), and impossible values like dates or 3°C readings are left blank. Every change is listed in the results so you can verify.
        </p>
      </div>
    </main>
  );
}
