export type CheckValue = "yes" | "none" | "";

export type InspectionForm = {
  date: string;
  direction: string;
  arrivalTime: string;
  dispatchTime: string;
  truckingCompany: string;
  plateNumber: string;
  truckType: string;
  driverName: string;
  licenseNumber: string;
  helpers: string;
  ppe: Record<string, CheckValue>;
  ppeRemarks: Record<string, string>;
  chemical: Record<string, CheckValue>;
  chemicalRemarks: Record<string, string>;
  vehicle: Record<string, CheckValue>;
  vehicleRemarks: Record<string, string>;
  preparedBy: string;
  notedBy: string;
};

export const ppeItems = [
  ["uniform", "Complete uniform and ID"],
  ["vest", "With safety vest"],
  ["shoes", "With safety shoes"],
  ["hardHat", "With hard hat"],
] as const;

export const chemicalItems = [
  ["gloves", "With gloves"],
  ["spillKit", "With spill kit"],
  ["firstAid", "With first aid kit"],
  ["fireExtinguisher", "With fire extinguisher"],
] as const;

export const vehicleItems = [
  ["dripPan", "Drip pan"],
  ["wheelChocks", "Wheel chocks"],
  ["spareTire", "Spare tire"],
  ["reflectiveTriangle", "Reflective triangle"],
  ["maintenanceTools", "Maintenance tools"],
  ["noLeak", "No leak"],
  ["noDarkSmoke", "No dark smoke"],
  ["cleanInterior", "Clean interior and exterior"],
  ["noObviousDamages", "No obvious damages"],
] as const;

const blankChecks = (items: readonly (readonly [string, string])[]) =>
  Object.fromEntries(items.map(([key]) => [key, ""])) as Record<string, CheckValue>;
const blankRemarks = (items: readonly (readonly [string, string])[]) =>
  Object.fromEntries(items.map(([key]) => [key, ""])) as Record<string, string>;

export function createEmptyInspection(): InspectionForm {
  return {
    date: new Date().toISOString().slice(0, 10), direction: "", arrivalTime: "", dispatchTime: "",
    truckingCompany: "", plateNumber: "", truckType: "", driverName: "", licenseNumber: "", helpers: "",
    ppe: blankChecks(ppeItems), ppeRemarks: blankRemarks(ppeItems),
    chemical: blankChecks(chemicalItems), chemicalRemarks: blankRemarks(chemicalItems),
    vehicle: blankChecks(vehicleItems), vehicleRemarks: blankRemarks(vehicleItems),
    preparedBy: "", notedBy: "",
  };
}

const esc = (value: unknown) => String(value ?? "").replace(/[&<>\"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;" })[char] ?? char);
const check = (value: CheckValue, expected: CheckValue) => value === expected ? "☒" : "☐";

export function inspectionHtml(form: InspectionForm): string {
  const itemRows = (items: readonly (readonly [string, string])[], values: Record<string, CheckValue>, remarks: Record<string, string>, left = true) => items.map(([key, label]) => `<tr><td>${esc(label)}</td><td>${check(values[key]!, left ? "none" : "yes")} ${left ? "None" : "Yes"}</td><td>${check(values[key]!, left ? "yes" : "none")} ${left ? "Yes" : "No"}</td><td>${esc(remarks[key])}</td></tr>`).join("");
  const info = (label: string, value: string) => `<div><b>${label}</b><span>${esc(value)}</span></div>`;
  return `<!doctype html><html><head><meta charset="utf-8"><title>Delivery Vehicle Inspection Form</title><style>body{font:12px Arial;color:#111;margin:24px}h1{text-align:center;font-size:22px;border:1px solid #111;margin:0;padding:7px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:4px 34px;border:1px solid #111;border-top:0;padding:12px}.grid div{display:grid;grid-template-columns:140px 1fr;border-bottom:1px solid #777;min-height:18px}.section{border:1px solid #111;border-top:0;padding:8px}.columns{display:grid;grid-template-columns:1fr 1fr;gap:24px}.columns h3{text-align:center;font-size:12px;margin:0 0 4px}.columns table{width:100%;border-collapse:collapse}.columns td{border-bottom:1px dotted #555;padding:3px}.columns td:nth-child(2),.columns td:nth-child(3){white-space:nowrap;width:45px}.note{font-weight:bold;border:1px solid #111;border-top:0;padding:10px}.signatures{display:grid;grid-template-columns:1fr 1fr;border:1px solid #111;border-top:0;text-align:center}.signature{min-height:72px;border-right:1px solid #111;padding-top:12px}.signature:last-child{border:0}@media print{body{margin:0}}</style></head><body><h1>Delivery Vehicle Inspection Form</h1><div class="grid">${info("Date", form.date)}${info("Truck Type", form.truckType)}${info("Incoming or Outgoing", form.direction)}${info("Name of Driver", form.driverName)}${info("Arrival Time", form.arrivalTime)}${info("Driver's License No.", form.licenseNumber)}${info("Dispatch Time", form.dispatchTime)}${info("Name of helper(s)", form.helpers)}${info("Trucking Company", form.truckingCompany)}${info("Plate Number", form.plateNumber)}</div><div class="section columns"><div><h3>Driver / Helper PPE Requirements</h3><table><tr><th>Requirement</th><th>None</th><th>Yes</th><th>Remarks</th></tr>${itemRows(ppeItems, form.ppe, form.ppeRemarks)}</table></div><div><h3>For Chemical Handling</h3><table><tr><th>Requirement</th><th>Yes</th><th>No</th><th>Remarks</th></tr>${itemRows(chemicalItems, form.chemical, form.chemicalRemarks, false)}</table></div></div><div class="section columns"><div><h3>Vehicle Requirements</h3><table><tr><th>Requirement</th><th>None</th><th>Yes</th><th>Remarks</th></tr>${itemRows(vehicleItems.slice(0, 5), form.vehicle, form.vehicleRemarks)}</table></div><div><h3>Vehicle Requirements</h3><table><tr><th>Requirement</th><th>None</th><th>Yes</th><th>Remarks</th></tr>${itemRows(vehicleItems.slice(5), form.vehicle, form.vehicleRemarks)}</table></div></div><div class="note">Note: If there are issues with trucker or forwarder, immediately inform Warehouse in-charge</div><div class="signatures"><div class="signature"><b>Prepared by</b><br><br><u>${esc(form.preparedBy)}</u><br>Signature over Printed Name<br><b>Guard</b></div><div class="signature"><b>Noted by</b><br><br><u>${esc(form.notedBy)}</u><br>Signature over Printed Name<br><b>Assistant Manager</b></div></div><p style="text-align:right;margin-top:28px">Document No. OPS-FORMS-06<br>Revision 01</p></body></html>`;
}

export function downloadInspection(form: InspectionForm, format: "doc" | "pdf") {
  const html = inspectionHtml(form);
  if (format === "pdf") {
    const printWindow = window.open("", "_blank", "noopener,noreferrer");
    if (!printWindow) throw new Error("Pop-up blocked");
    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();
    return;
  }
  const blob = new Blob([html], { type: "application/msword" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `Delivery-Vehicle-Inspection-${form.date || "form"}.doc`;
  anchor.click();
  URL.revokeObjectURL(url);
}
