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
    text += content.items.map((it) => it.str).join(" ") + " ";
  }
  return text;
}

// Send extracted text to our serverless function, which calls Gemini and
// returns clean structured data. Works for receipt, itinerary, or both.
export async function aiParse({ receiptText = "", itineraryText = "" }) {
  const res = await fetch("/api/parse", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ receiptText, itineraryText }),
  });
  if (!res.ok) {
    let detail = "";
    try { detail = (await res.json()).detail || ""; } catch {}
    throw new Error(`AI parse failed (${res.status}). ${detail}`);
  }
  return res.json();
}

// Collapse ALL whitespace (incl. non-breaking spaces \xa0) to single spaces so
// parsing never depends on newlines. Also normalize Trip.com's unicode colon (∶),
// including when the PDF reader puts spaces around it ("07 ∶ 30" -> "07:30").
function norm(t) {
  return t
    .replace(/(\d)\s*[\u2236:]\s*(\d)/g, "$1:$2")  // "07 ∶ 30" / "07 : 30" -> "07:30"
    .replace(/\u2236/g, ":")     // any remaining ∶ -> :
    .replace(/\u00a0/g, " ")     // non-breaking space -> space
    .replace(/\r/g, "")
    .replace(/\s+/g, " ")
    .trim();
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
  if ((m = t.match(/Contact Name\s+(.+?)\s+Email/i))) out.customerName = m[1].trim();
  if ((m = t.match(/Email\s+(\S+@\S+)/i))) out.customerEmail = m[1].trim();
  if ((m = t.match(/Passenger & E-ticket No\.\s*(.+?)\s+(\d{3}-\d{6,})/i))) {
    out.paxName = m[1].trim();
    out.eticket = m[2];
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
  if ((m = t.match(/E-ticket No\.\s*Airline.*?(\d{3}-\d{6,})/i))) out.eticket = m[1];
  if ((m = t.match(/Ticket Number:\s*(\d{3}-\d{6,})/i))) out.eticket = out.eticket || m[1];
  if ((m = t.match(/Economy\s+\d{3}-\d{6,}\s+([A-Z0-9]{5,7})/i))) out.airlineRef = m[1].trim();
  if ((m = t.match(/PNR\)\s*:?\s*([A-Z0-9]{5,7})/i))) out.airlineRef = out.airlineRef || m[1].trim();
  if ((m = t.match(/Dear\s+(.+?),/i))) out.customerName = m[1].trim();
  if ((m = t.match(/([\w.+-]+@[\w.-]+\.\w+)/))) {
    if (!m[1].includes("noreply") && !m[1].includes("trip.com")) out.customerEmail = m[1];
  }
  if ((m = t.match(/(?:name|^|\s)\s*((?:[A-Z]\s){1,6}[A-Z]?)\(First\s*name\)\s*([A-Z]+)\s*\(Last/i))) {
    out.paxName = m[1].replace(/\s+/g, " ").trim() + " " + m[2].trim();
  }

  // Format A: attached itinerary — "Departure HH:MM, Month D, YYYY, Airport ... Airline FLT"
  const reA = /Departure\s+(\d{1,2}[:.]\d{2}),\s*([A-Za-z]+\s+\d{1,2}),\s*(\d{4}),\s*(.+?)\s+Arrival\s+(\d{1,2}[:.]\d{2}),\s*([A-Za-z]+\s+\d{1,2}),\s*(\d{4}),\s*(.+?)\s+Airline\s+(.+?)\s+([A-Z]{1,3}\s?\d{2,4})\b/gi;
  let b;
  while ((b = reA.exec(t)) !== null) {
    const dep = b[4].trim();
    const arr = b[8].trim();
    out.segments.push({
      depTime: b[1], depDate: `${b[2]}, ${b[3]}`, depAirport: dep,
      arrTime: b[5], arrDate: `${b[6]}, ${b[7]}`, arrAirport: arr,
      airline: b[9].trim(), flightNo: b[10].replace(/\s/g, ""), cls: "Economy",
      route: shortRoute(dep) + " - " + shortRoute(arr),
    });
  }

  // Format B (fallback): Gmail confirmation email —
  // "Departure: Route | Date ... HH:MM CODE Airport  Airline FLT | Class ... HH:MM CODE Airport"
  if (out.segments.length === 0) {
    const reHeader = /(?:Departure|Return):\s*(.+?)\s*\|\s*([A-Za-z]+\s+\d{1,2},\s*\d{4})/gi;
    const headers = [];
    let h;
    while ((h = reHeader.exec(t)) !== null) headers.push({ route: h[1].trim(), date: h[2].trim(), index: h.index });
    const reB = /(\d{1,2}:\d{2})\s+([A-Z]{3})\s+(.+?)\s+([A-Z][A-Za-z]+(?:\s[A-Za-z]+)?)\s+([A-Z]{1,3}\d{2,4})\s*\|\s*([A-Za-z]+).*?\s+(\d{1,2}:\d{2})\s+([A-Z]{3})\s+(.+?)(?=\s+(?:Return|View|D L U N|Modify|$)|\s+\d{1,2}:\d{2}\s+[A-Z]{3})/gi;
    let s, i = 0;
    while ((s = reB.exec(t)) !== null) {
      const dep = s[3].trim();
      const arr = s[9].trim();
      const hdr = headers[i] || {};
      out.segments.push({
        depTime: s[1], depDate: hdr.date || "", depAirport: dep,
        arrTime: s[7], arrDate: hdr.date || "", arrAirport: arr,
        airline: s[4].trim(), flightNo: s[5].trim(), cls: s[6].trim() || "Economy",
        route: hdr.route || (shortRoute(dep) + " - " + shortRoute(arr)),
      });
      i++;
    }
  }

  return out;
}

// Merge receipt + itinerary into one form-ready object
export function mergeParsed(receipt, itinerary) {
  const r = receipt || {};
  const it = itinerary || {};
  const segments = (it.segments && it.segments.length) ? it.segments : [];
  const email = r.customerEmail || it.customerEmail || "";
  return {
    customerName: r.customerName || it.customerName || "",
    customerEmail: (email && !email.includes("*")) ? email : "",
    paxName: it.paxName || r.paxName || "",
    eticket: r.eticket || it.eticket || "",
    airlineRef: it.airlineRef || r.airlineRef || "",
    tripComPrice: r.total != null ? String(r.total) : "",
    tripBookingNo: r.tripBookingNo || it.tripBookingNo || "",
    segments,
  };
}
