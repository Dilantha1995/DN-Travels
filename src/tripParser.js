import * as pdfjsLib from "pdfjs-dist";
import workerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";

pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl;

// Read all text out of a PDF File object (browser)
export async function readPdfText(file) {
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  let text = "";
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    // Reconstruct lines using the y-position of each item
    let lastY = null;
    let line = "";
    const lines = [];
    for (const item of content.items) {
      const y = Math.round(item.transform[5]);
      if (lastY !== null && Math.abs(y - lastY) > 3) {
        lines.push(line.trim());
        line = "";
      }
      line += item.str + " ";
      lastY = y;
    }
    if (line.trim()) lines.push(line.trim());
    text += lines.join("\n") + "\n";
  }
  return text;
}

// Normalize the special unicode colon Trip.com uses (∶ U+2236) and spacing
function norm(t) {
  return t.replace(/\u2236/g, ":").replace(/\r/g, "").replace(/[ \t]+/g, " ");
}

function shortRoute(airport) {
  const map = [
    [/vel(a|e)na|mal[eé]/i, "Malé"],
    [/colombo|bandaranaike/i, "Colombo"],
  ];
  for (const [re, name] of map) if (re.test(airport)) return name;
  return airport.split(" ")[0];
}

export function parseReceipt(raw) {
  const t = norm(raw);
  const out = {};
  let m;
  if ((m = t.match(/Booking No\.?\s*([0-9]{6,})/i))) out.tripBookingNo = m[1];
  if ((m = t.match(/Contact Name\s+(.+)/i))) out.customerName = m[1].trim();
  if ((m = t.match(/Email\s+([^\s]+@[^\s]+)/i))) out.customerEmail = m[1].trim();
  if ((m = t.match(/Passenger & E-ticket No\.\s*\n?\s*(.+?)\s+(\d{3}-\d{6,})/i))) {
    out.paxName = m[1].trim();
    out.eticket = m[2].trim();
  }
  if ((m = t.match(/Fare\s*\$?\s*([\d,]+\.\d{2})/i))) out.fare = parseFloat(m[1].replace(/,/g, ""));
  if ((m = t.match(/Taxes? & fees\s*\$?\s*([\d,]+\.\d{2})/i))) out.taxes = parseFloat(m[1].replace(/,/g, ""));
  if ((m = t.match(/Total\b[^$]*\$?\s*([\d,]+\.\d{2})/i))) out.total = parseFloat(m[1].replace(/,/g, ""));
  return out;
}

export function parseItinerary(raw) {
  const t = norm(raw);
  const out = { segments: [] };
  let m;
  if ((m = t.match(/Booking No\.?\s*([0-9]{6,})/i))) out.tripBookingNo = m[1];
  if ((m = t.match(/E-ticket No\.\s*Airline[\s\S]*?(\d{3}-\d{6,})/i))) out.eticket = m[1].trim();
  if ((m = t.match(/Economy\s+\d{3}-\d{6,}\s+([A-Z0-9]{5,7})/i))) out.airlineRef = m[1].trim();
  if ((m = t.match(/([A-Z](?:\s[A-Z])*)\s*\(First\s*\n?\s*name\)\s*([A-Z]+)\s*\(Last/i))) {
    out.paxName = m[1].replace(/\s+/g, " ").trim() + " " + m[2].trim();
  }
  const flightSection = t.split(/Flight Information/i)[1] || t;
  const blockRe = /Departure\s+(\d{1,2}:\d{2}),\s*([^,]+),\s*([^,]+),\s*(.+?)\s*\n\s*Arrival\s+(\d{1,2}:\d{2}),\s*([^,]+),\s*([^,]+),\s*(.+?)\s*\n\s*Airline\s+(.+?)\s+([A-Z0-9]{2}\d{2,4})/gi;
  let b;
  while ((b = blockRe.exec(flightSection)) !== null) {
    const depAirport = b[4].trim();
    const arrAirport = b[8].trim();
    out.segments.push({
      depTime: b[1], depDate: `${b[2].trim()}, ${b[3].trim()}`, depAirport,
      arrTime: b[5], arrDate: `${b[6].trim()}, ${b[7].trim()}`, arrAirport,
      airline: b[9].trim(), flightNo: b[10].trim(), cls: "Economy",
      route: shortRoute(depAirport) + " - " + shortRoute(arrAirport),
    });
  }
  return out;
}

// Merge receipt + itinerary into one form-ready object
export function mergeParsed(receipt, itinerary) {
  const r = receipt || {};
  const it = itinerary || {};
  const segments = (it.segments && it.segments.length) ? it.segments : [];
  return {
    customerName: r.customerName || "",
    customerEmail: (r.customerEmail && !r.customerEmail.includes("*")) ? r.customerEmail : "",
    paxName: it.paxName || r.paxName || "",
    eticket: r.eticket || it.eticket || "",
    airlineRef: it.airlineRef || "",
    tripComPrice: r.total != null ? String(r.total) : "",
    tripBookingNo: r.tripBookingNo || it.tripBookingNo || "",
    segments,
  };
}
