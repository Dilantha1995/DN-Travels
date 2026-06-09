import React, { useState, useEffect, useRef, useCallback } from "react";
import {
  Plane, Plus, History, Download, Printer, Trash2, FileText,
  X, Check, Compass, Search, ChevronRight, RefreshCw, Edit3,
  Upload, FileUp, Loader2, Mail, Settings, BarChart3, FileSignature, Image as ImageIcon,
  ExternalLink, Globe
} from "lucide-react";
import * as XLSX from "xlsx";
import html2canvas from "html2canvas";
import { readPdfText, parseReceipt, parseItinerary, mergeParsed, aiParseImage } from "./tripParser";

// ---------- Branding ----------
const NAVY = "#0d3b66";
const OCEAN = "#1d6fb8";
const SKY = "#3f9fe0";
const PAPER = "#f5f8fb";

const CURRENCIES = {
  USD: { symbol: "$", label: "USD" },
  MVR: { symbol: "MVR", label: "MVR" },
  LKR: { symbol: "LKR", label: "LKR" },
};

// Live rates (USD -> currency). Updated at runtime from the API or manual override.
const RATES = { USD: 1, MVR: 15.42, LKR: 300 };

const MARKUP = 0.05; // 5%

// ---------- Helpers ----------
const pad = (n, len = 2) => String(n).padStart(len, "0");

function genBookingRef(seq) {
  // DN + YYMMDD + 3-digit sequence  ->  DN260528001
  const d = new Date();
  const stamp = `${pad(d.getFullYear() % 100)}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;
  return `DN${stamp}${pad(seq, 3)}`;
}

function genQuoteRef(seq) {
  // QT + YYMMDD + 3-digit sequence
  const d = new Date();
  const stamp = `${pad(d.getFullYear() % 100)}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;
  return `QT${stamp}${pad(seq, 3)}`;
}

function money(amount, cur) {
  const rate = RATES[cur] || 1;
  const val = amount * rate;
  const fixed = val.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return cur === "USD" ? `${CURRENCIES[cur].symbol}${fixed}` : `${CURRENCIES[cur].symbol} ${fixed}`;
}

const blankSegment = () => ({
  route: "", airline: "", flightNo: "", cls: "Economy",
  depTime: "", depDate: "", depAirport: "",
  arrTime: "", arrDate: "", arrAirport: "",
});

const blankForm = () => ({
  customerName: "",
  customerEmail: "",
  customerPhone: "",
  paxName: "",
  eticket: "",
  airlineRef: "",
  tripComPrice: "",       // base fare cost from Trip.com (USD)
  currency: "USD",
  tripType: "round",       // "one" | "round"
  segments: [blankSegment(), blankSegment()],
  baggage: "Checked 25kg • Carry-on 6kg • 1 Personal item",
  notes: "",
});

const blankQuote = () => ({
  customerName: "",
  customerEmail: "",
  customerPhone: "",
  description: "",          // free-text description (optional)
  costPrice: "",           // your cost (USD); +5% applied
  validUntil: "",          // optional validity date
  includeFlights: false,   // show flight segments?
  tripType: "round",       // "one" | "round"
  segments: [blankSegment(), blankSegment()],
  notes: "",
});

const defaultSettings = () => ({
  companyName: "DN TRAVELS",
  tagline: "Your Maldivian Connection",
  email: "info@dntravels.mv",
  phone: "",
  address: "Malé, Republic of Maldives",
  website: "",
  ratesMVR: "15.42",
  ratesLKR: "300",
  autoRates: true,
});

// ============================================================
export default function App() {
  const [tab, setTab] = useState("new");        // new | history
  const [records, setRecords] = useState([]);
  const [seq, setSeq] = useState(1);
  const [form, setForm] = useState(blankForm());
  const [loading, setLoading] = useState(true);
  const [saved, setSaved] = useState(null);     // the just-created record (preview)
  const [viewing, setViewing] = useState(null); // record opened from history
  const [docType, setDocType] = useState("invoice"); // invoice | itinerary
  const [search, setSearch] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [importing, setImporting] = useState(false);
  const [importMsg, setImportMsg] = useState(null);
  const [settings, setSettings] = useState(defaultSettings());
  const [ratesInfo, setRatesInfo] = useState({ MVR: 15.42, LKR: 300, fetchedAt: null, loading: false });
  const [reportFrom, setReportFrom] = useState("");
  const [reportTo, setReportTo] = useState("");
  const [rateNonce, setRateNonce] = useState(0); // forces re-render when inline rate changes
  const [shotBusy, setShotBusy] = useState(null); // "quote" | "ticket" | null
  const [shotMsg, setShotMsg] = useState(null);
  const [quotes, setQuotes] = useState([]);
  const [quoteSeq, setQuoteSeq] = useState(1);
  const [quoteForm, setQuoteForm] = useState(blankQuote());
  const [savedQuote, setSavedQuote] = useState(null);
  const [viewingQuote, setViewingQuote] = useState(null);
  const [editingQuoteId, setEditingQuoteId] = useState(null);

  // ---- Load persisted data (localStorage) ----
  useEffect(() => {
    let loadedSettings = defaultSettings();
    try {
      const r = localStorage.getItem("dn_records");
      if (r) setRecords(JSON.parse(r));
      const s = localStorage.getItem("dn_seq");
      if (s) setSeq(parseInt(s, 10));
      const q = localStorage.getItem("dn_quotes");
      if (q) setQuotes(JSON.parse(q));
      const qs = localStorage.getItem("dn_quoteSeq");
      if (qs) setQuoteSeq(parseInt(qs, 10));
      const st = localStorage.getItem("dn_settings");
      if (st) { loadedSettings = { ...loadedSettings, ...JSON.parse(st) }; setSettings(loadedSettings); }
    } catch (e) {
      console.error("Load error", e);
    }
    DOC_SETTINGS = loadedSettings;
    // apply saved manual rates immediately
    RATES.MVR = parseFloat(loadedSettings.ratesMVR) || RATES.MVR;
    RATES.LKR = parseFloat(loadedSettings.ratesLKR) || RATES.LKR;
    setRatesInfo((ri) => ({ ...ri, MVR: RATES.MVR, LKR: RATES.LKR }));
    // fetch live rates if auto is on
    if (loadedSettings.autoRates) fetchRates();
    setLoading(false);
  }, []);

  // ---- Fetch live exchange rates (free, no key) ----
  const fetchRates = useCallback(async () => {
    setRatesInfo((ri) => ({ ...ri, loading: true }));
    try {
      const res = await fetch("https://open.er-api.com/v6/latest/USD");
      const data = await res.json();
      if (data && data.rates && data.rates.MVR && data.rates.LKR) {
        RATES.MVR = data.rates.MVR;
        RATES.LKR = data.rates.LKR;
        setRatesInfo({ MVR: data.rates.MVR, LKR: data.rates.LKR, fetchedAt: new Date(), loading: false });
        setSettings((s) => ({ ...s, ratesMVR: data.rates.MVR.toFixed(2), ratesLKR: data.rates.LKR.toFixed(2) }));
      } else {
        setRatesInfo((ri) => ({ ...ri, loading: false }));
      }
    } catch (e) {
      console.error("Rate fetch failed", e);
      setRatesInfo((ri) => ({ ...ri, loading: false }));
    }
  }, []);

  const saveSettings = useCallback((newSettings) => {
    setSettings(newSettings);
    DOC_SETTINGS = newSettings;
    RATES.MVR = parseFloat(newSettings.ratesMVR) || RATES.MVR;
    RATES.LKR = parseFloat(newSettings.ratesLKR) || RATES.LKR;
    setRatesInfo((ri) => ({ ...ri, MVR: RATES.MVR, LKR: RATES.LKR }));
    try { localStorage.setItem("dn_settings", JSON.stringify(newSettings)); } catch (e) { console.error(e); }
  }, []);

  const persist = useCallback((newRecords, newSeq) => {
    try {
      localStorage.setItem("dn_records", JSON.stringify(newRecords));
      if (newSeq != null) localStorage.setItem("dn_seq", String(newSeq));
    } catch (e) {
      console.error("Storage error", e);
    }
  }, []);

  // ---- Derived pricing ----
  const basePrice = parseFloat(form.tripComPrice) || 0;
  const sellPrice = basePrice * (1 + MARKUP);
  const margin = sellPrice - basePrice;

  // ---- Form field updates ----
  // ---- Import Trip.com PDFs via AI (Gemini) ----
  const applyParsed = (p) => {
    setForm((f) => ({
      ...f,
      customerName: p.customerName || f.customerName,
      customerEmail: p.customerEmail || f.customerEmail,
      paxName: p.paxName || f.paxName,
      eticket: p.eticket || f.eticket,
      airlineRef: p.airlineRef || f.airlineRef,
      tripComPrice: (p.tripComPrice != null && p.tripComPrice !== 0 && p.tripComPrice !== "")
        ? String(p.tripComPrice) : f.tripComPrice,
      segments: (p.segments && p.segments.length)
        ? p.segments.map((s) => ({ ...blankSegment(), ...s }))
        : f.segments,
    }));
  };

  const runImport = async (file, which) => {
    if (!file) return;
    setImporting(which);
    setImportMsg(null);
    try {
      const text = await readPdfText(file);
      let p;
      if (which === "receipt") {
        const r = parseReceipt(text);
        p = mergeParsed(r, null);
      } else {
        const it = parseItinerary(text);
        p = mergeParsed(null, it);
      }
      applyParsed(p);
      const found = [];
      if (p.customerName) found.push("customer");
      if (p.paxName) found.push("passenger");
      if (p.segments && p.segments.length) found.push(`${p.segments.length} flight(s)`);
      if (p.tripComPrice) found.push("price");
      if (p.airlineRef) found.push("airline ref");
      setImportMsg({
        ok: found.length > 0,
        text: found.length
          ? `Imported \u2014 got: ${[...new Set(found)].join(", ")}. Please review.`
          : `Couldn't find details. Is this the Trip.com ${which === "receipt" ? "Receipt" : "Itinerary"} PDF?`,
      });
    } catch (e) {
      console.error(e);
      setImportMsg({ ok: false, text: `Couldn't read this PDF. ${e.message || ""}` });
    } finally {
      setImporting(false);
    }
  };

  const importReceipt = (fileList) => runImport((fileList || [])[0], "receipt");
  const importItinerary = (fileList) => runImport((fileList || [])[0], "itinerary");

  // ---- AI screenshot upload (image -> form fields) ----
  const handleScreenshot = async (file, target /* "ticket" | "quote" */) => {
    if (!file) return;
    if (!file.type || !file.type.startsWith("image/")) {
      setShotMsg({ ok: false, target, text: "Please upload an image file (PNG, JPG)." });
      return;
    }
    setShotBusy(target);
    setShotMsg(null);
    try {
      const p = await aiParseImage(file);
      const found = [];
      if (target === "ticket") {
        setForm((f) => ({
          ...f,
          customerName: p.customerName || f.customerName,
          customerEmail: p.customerEmail || f.customerEmail,
          paxName: p.paxName || f.paxName,
          eticket: p.eticket || f.eticket,
          airlineRef: p.airlineRef || f.airlineRef,
          tripComPrice: (p.tripComPrice != null && p.tripComPrice !== 0) ? String(p.tripComPrice) : f.tripComPrice,
          segments: (p.segments && p.segments.length)
            ? p.segments.map((s) => ({ ...blankSegment(), ...s }))
            : f.segments,
        }));
      } else {
        setQuoteForm((f) => {
          const segs = (p.segments && p.segments.length)
            ? p.segments.map((s) => ({ ...blankSegment(), ...s }))
            : f.segments;
          return {
            ...f,
            customerName: p.customerName || f.customerName,
            customerEmail: p.customerEmail || f.customerEmail,
            costPrice: (p.tripComPrice != null && p.tripComPrice !== 0) ? String(p.tripComPrice) : f.costPrice,
            includeFlights: (p.segments && p.segments.length) ? true : f.includeFlights,
            tripType: (p.segments && p.segments.length >= 2) ? "round" : (p.segments && p.segments.length === 1 ? "one" : f.tripType),
            segments: segs,
          };
        });
      }
      if (p.segments && p.segments.length) found.push(`${p.segments.length} flight(s)`);
      if (p.tripComPrice) found.push("price");
      if (p.customerName) found.push("customer");
      if (p.airlineRef) found.push("airline ref");
      setShotMsg({
        ok: found.length > 0,
        target,
        text: found.length
          ? `AI read screenshot \u2014 got: ${[...new Set(found)].join(", ")}. Please review.`
          : "AI couldn't extract clear details. Try a clearer screenshot or enter manually.",
      });
    } catch (e) {
      console.error(e);
      const msg = e.message || "Unknown error";
      const friendly = msg.includes("GEMINI_API_KEY")
        ? "AI is not configured (server missing API key). Add GEMINI_API_KEY in Vercel."
        : `Couldn't read this image. ${msg}`;
      setShotMsg({ ok: false, target, text: friendly });
    } finally {
      setShotBusy(null);
    }
  };

  // ---- Form field updates ----
  const setField = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  const setSeg = (i, k, v) =>
    setForm((f) => {
      const segs = [...f.segments];
      segs[i] = { ...segs[i], [k]: v };
      return { ...f, segments: segs };
    });
  const addSeg = () => setForm((f) => ({ ...f, segments: [...f.segments, blankSegment()] }));
  const removeSeg = (i) =>
    setForm((f) => ({ ...f, segments: f.segments.filter((_, idx) => idx !== i) }));

  // ---- Save / issue ticket ----
  const issueTicket = async () => {
    if (!form.customerName.trim() || !form.paxName.trim() || !basePrice) {
      alert("Please fill in Customer Name, Passenger Name, and Trip.com Price.");
      return;
    }

    if (editingId) {
      // update existing
      const updated = records.map((r) =>
        r.id === editingId
          ? { ...r, ...form, basePrice, sellPrice, margin }
          : r
      );
      setRecords(updated);
      persist(updated);
      const rec = updated.find((r) => r.id === editingId);
      setSaved(rec);
      setEditingId(null);
      setForm(blankForm());
      return;
    }

    const ref = genBookingRef(seq);
    const rec = {
      id: Date.now(),
      ref,
      issuedAt: new Date().toISOString(),
      ...form,
      basePrice,
      sellPrice,
      margin,
    };
    const newRecords = [rec, ...records];
    const newSeq = seq + 1;
    setRecords(newRecords);
    setSeq(newSeq);
    persist(newRecords, newSeq);
    setSaved(rec);
    setForm(blankForm());
  };

  const editRecord = (rec) => {
    setForm({ ...blankForm(), ...rec, tripComPrice: String(rec.basePrice) });
    setEditingId(rec.id);
    setTab("new");
    setSaved(null);
    setViewing(null);
  };

  const deleteRecord = async (id) => {
    if (!confirm("Delete this booking permanently?")) return;
    const newRecords = records.filter((r) => r.id !== id);
    setRecords(newRecords);
    persist(newRecords);
    if (viewing && viewing.id === id) setViewing(null);
  };

  // ---- Excel export ----
  const exportExcel = () => {
    if (records.length === 0) { alert("No tickets to export yet."); return; }
    const rows = records.map((r) => ({
      "Booking Ref": r.ref,
      "Issued Date": new Date(r.issuedAt).toLocaleString(),
      "Customer": r.customerName,
      "Customer Email": r.customerEmail,
      "Customer Phone": r.customerPhone,
      "Passenger": r.paxName,
      "E-ticket No.": r.eticket,
      "Airline Ref": r.airlineRef,
      "Route(s)": r.segments.map((s) => s.route || `${s.depAirport}-${s.arrAirport}`).join(" | "),
      "Flight(s)": r.segments.map((s) => `${s.airline} ${s.flightNo}`).join(" | "),
      "Trip.com Cost (USD)": Number(r.basePrice.toFixed(2)),
      "Sell Price (USD)": Number(r.sellPrice.toFixed(2)),
      "Margin (USD)": Number(r.margin.toFixed(2)),
      "Currency Quoted": r.currency,
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [
      { wch: 14 }, { wch: 20 }, { wch: 22 }, { wch: 24 }, { wch: 16 },
      { wch: 22 }, { wch: 16 }, { wch: 12 }, { wch: 28 }, { wch: 24 },
      { wch: 18 }, { wch: 16 }, { wch: 14 }, { wch: 14 },
    ];
    // totals row
    const totalCost = records.reduce((a, r) => a + r.basePrice, 0);
    const totalSell = records.reduce((a, r) => a + r.sellPrice, 0);
    const totalMargin = records.reduce((a, r) => a + r.margin, 0);
    XLSX.utils.sheet_add_aoa(ws, [[
      "TOTALS", "", "", "", "", "", "", "", "", "",
      Number(totalCost.toFixed(2)), Number(totalSell.toFixed(2)), Number(totalMargin.toFixed(2)), "",
    ]], { origin: -1 });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Issued Tickets");
    XLSX.writeFile(wb, `DN_Travels_Tickets_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // ---- Print ----
  const printRef = useRef(null);
  // ---- Download a document node as PNG image ----
  const downloadAsImage = async (node, filename) => {
    if (!node) return;
    try {
      const canvas = await html2canvas(node, { scale: 2, backgroundColor: "#ffffff", useCORS: true });
      const link = document.createElement("a");
      link.download = `${filename}.png`;
      link.href = canvas.toDataURL("image/png");
      link.click();
    } catch (e) {
      console.error("Image export failed", e);
      alert("Couldn't generate image. Try the PDF option instead.");
    }
  };

  const printRefQuote = useRef(null);

  const printNode = (node, fileTitle) => {
    if (!node) return;
    const w = window.open("", "_blank");
    w.document.write(`<!DOCTYPE html><html><head><title>${fileTitle}</title>
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700;800&family=Manrope:wght@400;500;600;700&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        body{font-family:'Manrope',sans-serif;color:#1a2733;padding:32px;}
        @page{margin:14mm;}
      </style></head><body>${node.innerHTML}</body></html>`);
    w.document.close();
    setTimeout(() => { w.focus(); w.print(); }, 400);
  };

  const handlePrint = () => {
    const rec = viewing || saved;
    const docLabel = docType === "invoice" ? "E-receipt" : "Itinerary";
    printNode(printRef.current, `${rec.ref} ${docLabel}`);
  };

  const activeRecord = viewing || saved;

  // ---- Quotation handlers ----
  const quoteBase = parseFloat(quoteForm.costPrice) || 0;
  const quoteSell = quoteBase * (1 + MARKUP);

  const setQField = (k, v) => setQuoteForm((f) => ({ ...f, [k]: v }));
  const setQSeg = (i, k, v) => setQuoteForm((f) => {
    const segs = [...f.segments]; segs[i] = { ...segs[i], [k]: v }; return { ...f, segments: segs };
  });
  const addQSeg = () => setQuoteForm((f) => ({ ...f, segments: [...f.segments, blankSegment()] }));
  const removeQSeg = (i) => setQuoteForm((f) => ({ ...f, segments: f.segments.filter((_, idx) => idx !== i) }));

  const persistQuotes = (list, sq) => {
    try {
      localStorage.setItem("dn_quotes", JSON.stringify(list));
      if (sq != null) localStorage.setItem("dn_quoteSeq", String(sq));
    } catch (e) { console.error(e); }
  };

  const issueQuote = () => {
    if (!quoteForm.customerName.trim() || !quoteBase) {
      alert("Please enter Customer Name and a Price.");
      return;
    }
    if (editingQuoteId) {
      const updated = quotes.map((q) => q.id === editingQuoteId
        ? { ...q, ...quoteForm, basePrice: quoteBase, sellPrice: quoteSell } : q);
      setQuotes(updated); persistQuotes(updated);
      setSavedQuote(updated.find((q) => q.id === editingQuoteId));
      setEditingQuoteId(null); setQuoteForm(blankQuote());
      return;
    }
    const ref = genQuoteRef(quoteSeq);
    const rec = {
      id: Date.now(), ref, issuedAt: new Date().toISOString(),
      ...quoteForm, basePrice: quoteBase, sellPrice: quoteSell,
    };
    const list = [rec, ...quotes];
    const nseq = quoteSeq + 1;
    setQuotes(list); setQuoteSeq(nseq); persistQuotes(list, nseq);
    setSavedQuote(rec); setQuoteForm(blankQuote());
  };

  const editQuote = (q) => {
    setQuoteForm({ ...blankQuote(), ...q, costPrice: String(q.basePrice) });
    setEditingQuoteId(q.id); setTab("quote"); setSavedQuote(null); setViewingQuote(null);
  };
  const deleteQuote = (id) => {
    if (!confirm("Delete this quotation?")) return;
    const list = quotes.filter((q) => q.id !== id);
    setQuotes(list); persistQuotes(list);
    if (viewingQuote && viewingQuote.id === id) setViewingQuote(null);
  };

  const activeQuote = viewingQuote || savedQuote;

  // ---- Compose customer email (opens Gmail with body prefilled) ----
  const emailToCustomer = (rec) => {
    const co = settings.companyName || "DN Travels";
    const segLines = rec.segments.map((s) => {
      const route = s.route || `${s.depAirport} → ${s.arrAirport}`;
      return `• ${route} | ${s.depDate}\n   ${s.depTime}  ${s.depAirport}\n   ${s.airline} ${s.flightNo} | ${s.cls}\n   ${s.arrTime}  ${s.arrAirport}`;
    }).join("\n\n");

    const subject = `Flight Booking Confirmed: ${rec.segments.map((s) => s.route || "").join(", ")}`;
    const body =
`Dear ${rec.customerName || "Customer"},

Thank you for choosing ${co}. Your flight(s) have been booked successfully! Please see your itinerary and e-receipt attached to this email.

Booking Reference: ${rec.ref}
Airline Booking Reference (PNR): ${rec.airlineRef || "-"}
Ticket Number: ${rec.eticket || "-"}

Passenger: ${rec.paxName}

FLIGHT DETAILS
${segLines}

Baggage Allowance: ${rec.baggage}

Total: ${money(rec.sellPrice, rec.currency)}

IMPORTANT
• Please carry a valid ID matching the passenger name above.
• Arrive at the airport at least 3 hours prior to departure.
• Tickets must be used in the sequence shown above.

We wish you a pleasant journey!

Best regards,
${co}
${settings.tagline || ""}
${settings.phone ? "Tel: " + settings.phone : ""}
${settings.email ? "Email: " + settings.email : ""}
${settings.website || ""}

------------------------------------------
NOTE: Please remember to attach the Itinerary.pdf and E-receipt.pdf before sending.`;

    const to = encodeURIComponent(rec.customerEmail || "");
    const su = encodeURIComponent(subject);
    const bo = encodeURIComponent(body);
    // Gmail compose URL — opens a new email ready to send
    const gmailUrl = `https://mail.google.com/mail/?view=cm&fs=1&to=${to}&su=${su}&body=${bo}`;
    window.open(gmailUrl, "_blank");
  };

  if (loading) {
    return (
      <div style={{ ...styles.shell, display: "flex", alignItems: "center", justifyContent: "center" }}>
        <div style={{ color: NAVY, fontFamily: "Sora, sans-serif", display: "flex", alignItems: "center", gap: 10 }}>
          <Compass className="spin" size={22} /> Loading DN Travels…
        </div>
      </div>
    );
  }

  return (
    <div style={styles.shell}>
      <FontInjector />
      {/* ===== Header ===== */}
      <header className="dn-header" style={styles.header}>
        <div style={styles.brand}>
          <Logo />
          <div>
            <div style={styles.brandName}>{settings.companyName || "DN TRAVELS"}</div>
            <div style={styles.brandTag}>{settings.tagline || "Your Maldivian Connection"}</div>
          </div>
        </div>
        <nav className="dn-nav" style={styles.nav}>
          <button
            style={{ ...styles.navBtn, ...(tab === "new" ? styles.navBtnActive : {}) }}
            onClick={() => { setTab("new"); setViewing(null); }}
          >
            <Plus size={16} /> New Ticket
          </button>
          <button
            style={{ ...styles.navBtn, ...(tab === "history" ? styles.navBtnActive : {}) }}
            onClick={() => { setTab("history"); setSaved(null); }}
          >
            <History size={16} /> History
            <span style={styles.badge}>{records.length}</span>
          </button>
          <button
            style={{ ...styles.navBtn, ...(tab === "quote" ? styles.navBtnActive : {}) }}
            onClick={() => { setTab("quote"); setSaved(null); setViewing(null); setViewingQuote(null); }}
          >
            <FileSignature size={16} /> Quotation
          </button>
          <button
            style={{ ...styles.navBtn, ...(tab === "reports" ? styles.navBtnActive : {}) }}
            onClick={() => { setTab("reports"); setSaved(null); setViewing(null); }}
          >
            <BarChart3 size={16} /> Reports
          </button>
          <button
            style={{ ...styles.navBtn, ...(tab === "settings" ? styles.navBtnActive : {}) }}
            onClick={() => { setTab("settings"); setSaved(null); setViewing(null); }}
          >
            <Settings size={16} /> Settings
          </button>
        </nav>
      </header>

      <main style={styles.main}>
        {/* ============ NEW TICKET ============ */}
        {tab === "new" && !saved && (
          <div className="dn-grid" style={styles.grid}>
            {/* --- Form --- */}
            <div style={styles.card}>
              <SectionTitle icon={<Edit3 size={16} />}>
                {editingId ? "Edit Booking" : "Issue New Ticket"}
              </SectionTitle>

              <BookingSites />
              <ScreenshotUploader target="ticket" busy={shotBusy === "ticket"} msg={shotMsg}
                onPick={(f) => handleScreenshot(f, "ticket")} />
              <div style={styles.importBox}>
                <div style={styles.importHead}>
                  <FileUp size={18} color={OCEAN} />
                  <div>
                    <div style={styles.importTitle}>Auto-fill from Trip.com</div>
                    <div style={styles.importSub}>Upload each PDF in its own slot</div>
                  </div>
                </div>
                <div style={styles.importBtnRow}>
                  <label style={{ ...styles.importBtn, ...(importing === "receipt" ? styles.importBtnBusy : {}) }}>
                    {importing === "receipt" ? <Loader2 size={16} className="spin" /> : <Upload size={16} />}
                    {importing === "receipt" ? "Reading\u2026" : "Receipt PDF"}
                    <input type="file" accept="application/pdf" style={{ display: "none" }}
                      disabled={!!importing}
                      onChange={(e) => { importReceipt(e.target.files); e.target.value = ""; }} />
                  </label>
                  <label style={{ ...styles.importBtn, ...(importing === "itinerary" ? styles.importBtnBusy : {}) }}>
                    {importing === "itinerary" ? <Loader2 size={16} className="spin" /> : <Upload size={16} />}
                    {importing === "itinerary" ? "Reading\u2026" : "Itinerary PDF"}
                    <input type="file" accept="application/pdf" style={{ display: "none" }}
                      disabled={!!importing}
                      onChange={(e) => { importItinerary(e.target.files); e.target.value = ""; }} />
                  </label>
                </div>
                {importMsg && (
                  <div style={{ ...styles.importMsg, color: importMsg.ok ? "#1a7a4c" : "#c0392b" }}>
                    {importMsg.ok ? <Check size={14} /> : <X size={14} />} {importMsg.text}
                  </div>
                )}
              </div>

              <Group label="Customer Details">
                <Field label="Customer Name *" value={form.customerName} onChange={(v) => setField("customerName", v)} placeholder="Full name" />
                <Row>
                  <Field label="Email" value={form.customerEmail} onChange={(v) => setField("customerEmail", v)} placeholder="email@example.com" />
                  <Field label="Phone" value={form.customerPhone} onChange={(v) => setField("customerPhone", v)} placeholder="+960 …" />
                </Row>
              </Group>

              <Group label="Passenger & Ticket">
                <Field label="Passenger Name (as on ticket) *" value={form.paxName} onChange={(v) => setField("paxName", v)} placeholder="D L U N PREMATHILAKA" />
                <Row>
                  <Field label="E-ticket No." value={form.eticket} onChange={(v) => setField("eticket", v)} placeholder="072-9411581236" />
                  <Field label="Airline Booking Ref" value={form.airlineRef} onChange={(v) => setField("airlineRef", v)} placeholder="NMYDEB" />
                </Row>
              </Group>

              <Group label="Flight Segments">
                <div style={styles.tripTypeRow}>
                  <button
                    style={{ ...styles.tripBtn, ...(form.tripType === "one" ? styles.tripBtnActive : {}) }}
                    onClick={() => setForm((f) => ({ ...f, tripType: "one", segments: [f.segments[0] || blankSegment()] }))}
                  >One-way</button>
                  <button
                    style={{ ...styles.tripBtn, ...(form.tripType === "round" ? styles.tripBtnActive : {}) }}
                    onClick={() => setForm((f) => ({ ...f, tripType: "round", segments: f.segments.length >= 2 ? f.segments : [f.segments[0] || blankSegment(), blankSegment()] }))}
                  >Round-trip</button>
                </div>
                {form.segments.map((s, i) => (
                  <div key={i} style={styles.segBox}>
                    <div style={styles.segHead}>
                      <span style={{ fontWeight: 700, color: NAVY, fontSize: 13 }}>
                        {form.tripType === "round" ? (i === 0 ? "Outbound" : i === 1 ? "Return" : `Segment ${i + 1}`) : `Segment ${i + 1}`}
                      </span>
                      {form.segments.length > 1 && form.tripType !== "round" && (
                        <button style={styles.iconBtnSm} onClick={() => removeSeg(i)} title="Remove">
                          <X size={14} />
                        </button>
                      )}
                    </div>
                    <Field label="Route" value={s.route} onChange={(v) => setSeg(i, "route", v)} placeholder="Malé - Colombo" />
                    <Row>
                      <Field label="Airline" value={s.airline} onChange={(v) => setSeg(i, "airline", v)} placeholder="Gulf Air" />
                      <Field label="Flight No." value={s.flightNo} onChange={(v) => setSeg(i, "flightNo", v)} placeholder="GF144" />
                    </Row>
                    <Row>
                      <DateField label="Dep. Date" value={s.depDate} onChange={(v) => setSeg(i, "depDate", v)} />
                      <Field label="Dep. Time" value={s.depTime} onChange={(v) => setSeg(i, "depTime", v)} placeholder="07:30" />
                    </Row>
                    <Field label="Dep. Airport" value={s.depAirport} onChange={(v) => setSeg(i, "depAirport", v)} placeholder="Velana International Airport T1" />
                    <Row>
                      <DateField label="Arr. Date" value={s.arrDate} onChange={(v) => setSeg(i, "arrDate", v)} />
                      <Field label="Arr. Time" value={s.arrTime} onChange={(v) => setSeg(i, "arrTime", v)} placeholder="09:40" />
                    </Row>
                    <Field label="Arr. Airport" value={s.arrAirport} onChange={(v) => setSeg(i, "arrAirport", v)} placeholder="Colombo Bandaranaike International Airport" />
                  </div>
                ))}
                {form.tripType !== "round" && (
                  <button style={styles.addSegBtn} onClick={addSeg}>
                    <Plus size={15} /> Add another segment
                  </button>
                )}
              </Group>

              <Group label="Baggage & Notes">
                <Field label="Baggage Allowance" value={form.baggage} onChange={(v) => setField("baggage", v)} />
                <Field label="Notes (optional)" value={form.notes} onChange={(v) => setField("notes", v)} placeholder="Any remarks for the customer" />
              </Group>
            </div>

            {/* --- Pricing & action --- */}
            <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
              <div style={styles.card}>
                <SectionTitle icon={<FileText size={16} />}>Pricing (auto +5%)</SectionTitle>
                <Field
                  label="Trip.com Price — your cost (USD) *"
                  value={form.tripComPrice}
                  onChange={(v) => setField("tripComPrice", v.replace(/[^0-9.]/g, ""))}
                  placeholder="268.10"
                />
                <div style={styles.currencyRow}>
                  {Object.keys(CURRENCIES).map((c) => (
                    <button
                      key={c}
                      onClick={() => setField("currency", c)}
                      style={{ ...styles.curBtn, ...(form.currency === c ? styles.curBtnActive : {}) }}
                    >
                      {c}
                    </button>
                  ))}
                </div>

                <div style={styles.priceBreak}>
                  <PriceLine label="Trip.com cost" value={money(basePrice, form.currency)} />
                  <PriceLine label="Markup (5%)" value={money(margin, form.currency)} accent />
                  <div style={styles.priceDivider} />
                  <PriceLine label="Customer pays" value={money(sellPrice, form.currency)} big />
                </div>
                {form.currency !== "USD" && (
                  <div style={styles.fxEdit}>
                    <label style={styles.fxLabel}>Rate: 1 USD =</label>
                    <input
                      style={styles.fxInput}
                      value={RATES[form.currency]}
                      onChange={(e) => {
                        const v = e.target.value.replace(/[^0-9.]/g, "");
                        RATES[form.currency] = parseFloat(v) || 0;
                        setRateNonce((n) => n + 1);
                      }}
                    />
                    <span style={styles.fxCur}>{form.currency}</span>
                    <button
                      style={styles.fxRefresh}
                      title="Fetch live rate"
                      onClick={async () => { await fetchRates(); setRateNonce((n) => n + 1); }}
                    >
                      <RefreshCw size={13} /> Live
                    </button>
                  </div>
                )}
                {form.currency !== "USD" && (
                  <div style={styles.fxHint}>
                    Edit the rate above for this booking, or set a default in Settings.
                  </div>
                )}
              </div>

              <button style={styles.primaryBtn} onClick={issueTicket}>
                <Check size={18} />
                {editingId ? "Update Booking" : "Issue Ticket & Generate Docs"}
              </button>
              {editingId && (
                <button style={styles.ghostBtn} onClick={() => { setEditingId(null); setForm(blankForm()); }}>
                  Cancel edit
                </button>
              )}
              <p style={styles.hint}>
                A booking reference like <b>{genBookingRef(seq)}</b> will be generated automatically. Everything is saved to history.
              </p>
            </div>
          </div>
        )}

        {/* ============ DOCUMENT PREVIEW (after issuing or from history) ============ */}
        {activeRecord && (tab === "new" ? saved : true) && (
          <div>
            <div style={styles.docToolbar}>
              <div style={styles.docTabs}>
                <button
                  style={{ ...styles.docTab, ...(docType === "invoice" ? styles.docTabActive : {}) }}
                  onClick={() => setDocType("invoice")}
                >Invoice</button>
                <button
                  style={{ ...styles.docTab, ...(docType === "itinerary" ? styles.docTabActive : {}) }}
                  onClick={() => setDocType("itinerary")}
                >Itinerary</button>
              </div>
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                <button style={styles.toolBtn} onClick={handlePrint}><Printer size={15} /> PDF</button>
                <button style={styles.toolBtn} onClick={() => downloadAsImage(printRef.current, `${activeRecord.ref} ${docType === "invoice" ? "E-receipt" : "Itinerary"}`)}><ImageIcon size={15} /> Image</button>
                <button style={styles.emailBtn} onClick={() => emailToCustomer(activeRecord)}><Mail size={15} /> Email</button>
                <button style={styles.toolBtn} onClick={() => editRecord(activeRecord)}><Edit3 size={15} /> Edit</button>
                <button
                  style={styles.toolBtnGhost}
                  onClick={() => { setSaved(null); setViewing(null); setTab(viewing ? "history" : "new"); }}
                ><X size={15} /> Close</button>
              </div>
            </div>

            <div style={styles.docWrap}>
              <div ref={printRef}>
                {docType === "invoice"
                  ? <Invoice rec={activeRecord} />
                  : <Itinerary rec={activeRecord} />}
              </div>
            </div>
          </div>
        )}

        {/* ============ HISTORY ============ */}
        {tab === "history" && !viewing && (
          <div>
            <div style={styles.histBar}>
              <div style={styles.searchBox}>
                <Search size={16} color="#7d93a8" />
                <input
                  style={styles.searchInput}
                  placeholder="Search by name, ref, e-ticket…"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                />
              </div>
              <button style={styles.excelBtn} onClick={exportExcel}>
                <Download size={16} /> Export Excel
              </button>
            </div>

            {records.length === 0 ? (
              <div style={styles.empty}>
                <Plane size={40} color="#b8cad9" />
                <p>No tickets issued yet. Create your first one under <b>New Ticket</b>.</p>
              </div>
            ) : (
              <div style={styles.histList}>
                {records
                  .filter((r) => {
                    const q = search.toLowerCase();
                    return !q ||
                      r.ref.toLowerCase().includes(q) ||
                      r.customerName.toLowerCase().includes(q) ||
                      r.paxName.toLowerCase().includes(q) ||
                      (r.eticket || "").toLowerCase().includes(q);
                  })
                  .map((r) => (
                    <div key={r.id} style={styles.histRow} onClick={() => { setViewing(r); setDocType("invoice"); }}>
                      <div style={styles.refTag}>{r.ref}</div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={styles.histName}>{r.customerName}</div>
                        <div style={styles.histSub}>
                          {r.segments.map((s) => s.route || `${s.depAirport}→${s.arrAirport}`).join("  •  ")}
                        </div>
                        <div style={styles.histMeta}>
                          {new Date(r.issuedAt).toLocaleDateString()} · {r.paxName}
                        </div>
                      </div>
                      <div style={{ textAlign: "right" }}>
                        <div style={styles.histPrice}>{money(r.sellPrice, r.currency)}</div>
                        <div style={styles.histMargin}>+{money(r.margin, r.currency)} margin</div>
                      </div>
                      <button
                        style={styles.iconBtnSm}
                        onClick={(e) => { e.stopPropagation(); deleteRecord(r.id); }}
                        title="Delete"
                      ><Trash2 size={15} /></button>
                      <ChevronRight size={18} color="#b8cad9" />
                    </div>
                  ))}
              </div>
            )}
          </div>
        )}

        {/* ============ QUOTATION ============ */}
        {tab === "quote" && !savedQuote && !viewingQuote && (
          <div className="dn-grid" style={styles.grid}>
            <div style={styles.card}>
              <SectionTitle icon={<FileSignature size={16} />}>
                {editingQuoteId ? "Edit Quotation" : "New Quotation"}
              </SectionTitle>
              <BookingSites />
              <ScreenshotUploader target="quote" busy={shotBusy === "quote"} msg={shotMsg}
                onPick={(f) => handleScreenshot(f, "quote")} />
              <Group label="Customer Details">
                <Field label="Customer Name *" value={quoteForm.customerName} onChange={(v) => setQField("customerName", v)} placeholder="Full name" />
                <Row>
                  <Field label="Email" value={quoteForm.customerEmail} onChange={(v) => setQField("customerEmail", v)} placeholder="email@example.com" />
                  <Field label="Phone" value={quoteForm.customerPhone} onChange={(v) => setQField("customerPhone", v)} placeholder="+960 …" />
                </Row>
              </Group>
              <Group label="Quotation Details">
                <Field label="Description" value={quoteForm.description} onChange={(v) => setQField("description", v)} placeholder="e.g. Return ticket Malé–Colombo, Gulf Air, Economy" />
                <label style={{ ...styles.label, display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
                  <input type="checkbox" checked={quoteForm.includeFlights} onChange={(e) => setQField("includeFlights", e.target.checked)} />
                  Include detailed flight segments
                </label>
                {quoteForm.includeFlights && (
                  <>
                    <div style={styles.tripTypeRow}>
                      <button
                        style={{ ...styles.tripBtn, ...(quoteForm.tripType === "one" ? styles.tripBtnActive : {}) }}
                        onClick={() => setQuoteForm((f) => ({ ...f, tripType: "one", segments: [f.segments[0] || blankSegment()] }))}
                      >One-way</button>
                      <button
                        style={{ ...styles.tripBtn, ...(quoteForm.tripType === "round" ? styles.tripBtnActive : {}) }}
                        onClick={() => setQuoteForm((f) => ({ ...f, tripType: "round", segments: f.segments.length >= 2 ? f.segments : [f.segments[0] || blankSegment(), blankSegment()] }))}
                      >Round-trip</button>
                    </div>
                    {quoteForm.segments.map((s, i) => (
                      <div key={i} style={styles.segBox}>
                        <div style={styles.segHead}>
                          <span style={{ fontWeight: 700, color: NAVY, fontSize: 13 }}>
                            {quoteForm.tripType === "round" ? (i === 0 ? "Outbound" : i === 1 ? "Return" : `Segment ${i + 1}`) : `Segment ${i + 1}`}
                          </span>
                          {quoteForm.segments.length > 1 && quoteForm.tripType !== "round" && (
                            <button style={styles.iconBtnSm} onClick={() => removeQSeg(i)}><X size={14} /></button>
                          )}
                        </div>
                        <Field label="Route" value={s.route} onChange={(v) => setQSeg(i, "route", v)} placeholder="Malé - Colombo" />
                        <Row>
                          <Field label="Airline" value={s.airline} onChange={(v) => setQSeg(i, "airline", v)} placeholder="Gulf Air" />
                          <Field label="Flight No." value={s.flightNo} onChange={(v) => setQSeg(i, "flightNo", v)} placeholder="GF144" />
                        </Row>
                        <Row>
                          <DateField label="Dep. Date" value={s.depDate} onChange={(v) => setQSeg(i, "depDate", v)} />
                          <Field label="Dep. Time" value={s.depTime} onChange={(v) => setQSeg(i, "depTime", v)} placeholder="07:30" />
                        </Row>
                      </div>
                    ))}
                    {quoteForm.tripType !== "round" && (
                      <button style={styles.addSegBtn} onClick={addQSeg}><Plus size={15} /> Add segment</button>
                    )}
                  </>
                )}
                <DateField label="Valid Until (optional)" value={quoteForm.validUntil} onChange={(v) => setQField("validUntil", v)} />
                <Field label="Notes (optional)" value={quoteForm.notes} onChange={(v) => setQField("notes", v)} />
              </Group>
            </div>

            <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
              <div style={styles.card}>
                <SectionTitle icon={<FileText size={16} />}>Price (auto +5%)</SectionTitle>
                <Field label="Your cost (USD) *" value={quoteForm.costPrice} onChange={(v) => setQField("costPrice", v.replace(/[^0-9.]/g, ""))} placeholder="268.10" />
                <div style={styles.priceBreak}>
                  <PriceLine label="Cost" value={money(quoteBase, "USD")} />
                  <PriceLine label="Markup (5%)" value={money(quoteSell - quoteBase, "USD")} accent />
                  <div style={styles.priceDivider} />
                  <PriceLine label="Quote (USD)" value={money(quoteSell, "USD")} big />
                  <PriceLine label="Quote (MVR)" value={money(quoteSell, "MVR")} />
                  <PriceLine label="Quote (LKR)" value={money(quoteSell, "LKR")} />
                </div>
                <div style={styles.fxHint}>The quotation shows the price in all three currencies.</div>
              </div>
              <button style={styles.primaryBtn} onClick={issueQuote}>
                <Check size={18} /> {editingQuoteId ? "Update Quotation" : "Generate Quotation"}
              </button>
              {editingQuoteId && (
                <button style={styles.ghostBtn} onClick={() => { setEditingQuoteId(null); setQuoteForm(blankQuote()); }}>Cancel edit</button>
              )}
              <p style={styles.hint}>A quote number like <b>{genQuoteRef(quoteSeq)}</b> is generated automatically. Saved to Quotation history below.</p>
            </div>
          </div>
        )}

        {/* Quotation document preview */}
        {tab === "quote" && activeQuote && (
          <div>
            <div style={styles.docToolbar}>
              <div style={{ fontFamily: "Sora, sans-serif", fontWeight: 700, color: NAVY }}>Quotation {activeQuote.ref}</div>
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                <button style={styles.toolBtn} onClick={() => printNode(printRefQuote.current, `${activeQuote.ref} Quotation`)}><Printer size={15} /> PDF</button>
                <button style={styles.toolBtn} onClick={() => downloadAsImage(printRefQuote.current, `${activeQuote.ref} Quotation`)}><ImageIcon size={15} /> Image</button>
                <button style={styles.toolBtn} onClick={() => editQuote(activeQuote)}><Edit3 size={15} /> Edit</button>
                <button style={styles.toolBtnGhost} onClick={() => { setSavedQuote(null); setViewingQuote(null); }}><X size={15} /> Close</button>
              </div>
            </div>
            <div style={styles.docWrap}>
              <div ref={printRefQuote}><Quotation rec={activeQuote} /></div>
            </div>
          </div>
        )}

        {/* Quotation history list */}
        {tab === "quote" && !savedQuote && !viewingQuote && quotes.length > 0 && (
          <div style={{ marginTop: 26 }}>
            <div style={{ fontFamily: "Sora, sans-serif", fontWeight: 700, color: NAVY, marginBottom: 12 }}>Saved Quotations</div>
            <div style={styles.histList}>
              {quotes.map((q) => (
                <div key={q.id} style={styles.histRow} onClick={() => setViewingQuote(q)}>
                  <div style={{ ...styles.refTag, background: OCEAN }}>{q.ref}</div>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={styles.histName}>{q.customerName}</div>
                    <div style={styles.histSub}>{q.description || (q.segments && q.segments.map((s) => s.route).filter(Boolean).join(", ")) || "Quotation"}</div>
                    <div style={styles.histMeta}>{new Date(q.issuedAt).toLocaleDateString()}{q.validUntil ? ` · valid until ${q.validUntil}` : ""}</div>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <div style={styles.histPrice}>{money(q.sellPrice, "USD")}</div>
                    <div style={styles.histMargin}>{money(q.sellPrice, "MVR")}</div>
                  </div>
                  <button style={styles.iconBtnSm} onClick={(e) => { e.stopPropagation(); deleteQuote(q.id); }}><Trash2 size={15} /></button>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ============ REPORTS ============ */}
        {tab === "reports" && (
          <ReportsView
            records={records}
            reportFrom={reportFrom} setReportFrom={setReportFrom}
            reportTo={reportTo} setReportTo={setReportTo}
          />
        )}

        {/* ============ SETTINGS ============ */}
        {tab === "settings" && (
          <SettingsView
            settings={settings}
            saveSettings={saveSettings}
            ratesInfo={ratesInfo}
            fetchRates={fetchRates}
          />
        )}
      </main>
    </div>
  );
}

// ============================================================
// Document components
// ============================================================
// Holds current settings so document components can read contact details.
let DOC_SETTINGS = defaultSettings();

// ============================================================
// Reports
// ============================================================
function ReportsView({ records, reportFrom, setReportFrom, reportTo, setReportTo }) {
  const from = reportFrom ? new Date(reportFrom + "T00:00:00") : null;
  const to = reportTo ? new Date(reportTo + "T23:59:59") : null;
  const filtered = records.filter((r) => {
    const d = new Date(r.issuedAt);
    if (from && d < from) return false;
    if (to && d > to) return false;
    return true;
  });
  const revenue = filtered.reduce((a, r) => a + r.sellPrice, 0);
  const cost = filtered.reduce((a, r) => a + r.basePrice, 0);
  const profit = filtered.reduce((a, r) => a + r.margin, 0);

  const exportReport = () => {
    if (filtered.length === 0) { alert("No tickets in this period."); return; }
    const rows = filtered.map((r) => ({
      "Booking Ref": r.ref,
      "Issued": new Date(r.issuedAt).toLocaleString(),
      "Customer": r.customerName,
      "Passenger": r.paxName,
      "Route(s)": r.segments.map((s) => s.route || `${s.depAirport}-${s.arrAirport}`).join(" | "),
      "Cost (USD)": Number(r.basePrice.toFixed(2)),
      "Revenue (USD)": Number(r.sellPrice.toFixed(2)),
      "Profit (USD)": Number(r.margin.toFixed(2)),
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.sheet_add_aoa(ws, [[
      "TOTALS", "", "", "", "",
      Number(cost.toFixed(2)), Number(revenue.toFixed(2)), Number(profit.toFixed(2)),
    ]], { origin: -1 });
    ws["!cols"] = [{ wch: 14 }, { wch: 20 }, { wch: 22 }, { wch: 22 }, { wch: 28 }, { wch: 12 }, { wch: 13 }, { wch: 12 }];
    const wb = XLSX.utils.book_new();
    const label = `${reportFrom || "all"}_to_${reportTo || "now"}`;
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    XLSX.writeFile(wb, `DN_Travels_Report_${label}.xlsx`);
  };

  const quick = (days) => {
    const t = new Date();
    const f = new Date(); f.setDate(f.getDate() - days);
    setReportFrom(f.toISOString().slice(0, 10));
    setReportTo(t.toISOString().slice(0, 10));
  };

  return (
    <div style={{ maxWidth: 880, margin: "0 auto" }}>
      <div style={styles.card}>
        <SectionTitle icon={<BarChart3 size={16} />}>Revenue & Profit Report</SectionTitle>
        <div style={styles.reportRange}>
          <div style={styles.field}>
            <label style={styles.label}>From</label>
            <input type="date" style={styles.input} value={reportFrom} onChange={(e) => setReportFrom(e.target.value)} />
          </div>
          <div style={styles.field}>
            <label style={styles.label}>To</label>
            <input type="date" style={styles.input} value={reportTo} onChange={(e) => setReportTo(e.target.value)} />
          </div>
        </div>
        <div style={styles.quickRow}>
          <button style={styles.quickBtn} onClick={() => quick(7)}>Last 7 days</button>
          <button style={styles.quickBtn} onClick={() => quick(30)}>Last 30 days</button>
          <button style={styles.quickBtn} onClick={() => quick(90)}>Last 90 days</button>
          <button style={styles.quickBtn} onClick={() => { setReportFrom(""); setReportTo(""); }}>All time</button>
        </div>

        <div className="dn-report-stats" style={styles.statGrid}>
          <div style={styles.statBox}>
            <div style={styles.statLabel}>Tickets</div>
            <div style={styles.statVal}>{filtered.length}</div>
          </div>
          <div style={styles.statBox}>
            <div style={styles.statLabel}>Revenue (customer paid)</div>
            <div style={styles.statVal}>${revenue.toFixed(2)}</div>
          </div>
          <div style={{ ...styles.statBox, background: "#eafaf1", borderColor: "#bfe8d0" }}>
            <div style={styles.statLabel}>Profit (your 5%)</div>
            <div style={{ ...styles.statVal, color: "#1a7a4c" }}>${profit.toFixed(2)}</div>
          </div>
        </div>
        <div style={styles.costNote}>Cost paid to Trip.com in this period: ${cost.toFixed(2)}</div>

        <button style={{ ...styles.excelBtn, marginTop: 18 }} onClick={exportReport}>
          <Download size={16} /> Export this report to Excel
        </button>
      </div>
    </div>
  );
}

// ============================================================
// Settings
// ============================================================
function SettingsView({ settings, saveSettings, ratesInfo, fetchRates }) {
  const [local, setLocal] = useState(settings);
  const [savedFlash, setSavedFlash] = useState(false);
  useEffect(() => { setLocal(settings); }, [settings]);
  const set = (k, v) => setLocal((s) => ({ ...s, [k]: v }));
  const doSave = () => {
    saveSettings(local);
    setSavedFlash(true);
    setTimeout(() => setSavedFlash(false), 1800);
  };

  return (
    <div style={{ maxWidth: 720, margin: "0 auto", display: "flex", flexDirection: "column", gap: 18 }}>
      <div style={styles.card}>
        <SectionTitle icon={<Settings size={16} />}>Company Details</SectionTitle>
        <p style={styles.hint}>These appear on your invoices, itineraries, and customer emails.</p>
        <Field label="Company Name" value={local.companyName} onChange={(v) => set("companyName", v)} />
        <Field label="Tagline" value={local.tagline} onChange={(v) => set("tagline", v)} />
        <div style={styles.row}>
          <Field label="Email" value={local.email} onChange={(v) => set("email", v)} placeholder="info@dntravels.mv" />
          <Field label="Phone" value={local.phone} onChange={(v) => set("phone", v)} placeholder="+960 …" />
        </div>
        <Field label="Address" value={local.address} onChange={(v) => set("address", v)} />
        <Field label="Website (optional)" value={local.website} onChange={(v) => set("website", v)} placeholder="www.dntravels.mv" />
      </div>

      <div style={styles.card}>
        <SectionTitle icon={<RefreshCw size={16} />}>Exchange Rates (per 1 USD)</SectionTitle>
        <p style={styles.hint}>
          {ratesInfo.fetchedAt
            ? `Live rates fetched ${ratesInfo.fetchedAt.toLocaleString()}.`
            : "Using saved/manual rates."} You can override them below.
        </p>
        <label style={{ ...styles.label, display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
          <input type="checkbox" checked={!!local.autoRates} onChange={(e) => set("autoRates", e.target.checked)} />
          Auto-fetch live rates on startup
        </label>
        <div style={styles.row}>
          <Field label="1 USD = ? MVR" value={local.ratesMVR} onChange={(v) => set("ratesMVR", v.replace(/[^0-9.]/g, ""))} />
          <Field label="1 USD = ? LKR" value={local.ratesLKR} onChange={(v) => set("ratesLKR", v.replace(/[^0-9.]/g, ""))} />
        </div>
        <button
          style={{ ...styles.quickBtn, marginTop: 4 }}
          onClick={async () => { await fetchRates(); }}
          disabled={ratesInfo.loading}
        >
          {ratesInfo.loading ? "Fetching…" : "Fetch live rates now"}
        </button>
        {ratesInfo.fetchedAt && (
          <div style={styles.costNote}>
            Latest live: 1 USD = {ratesInfo.MVR?.toFixed(2)} MVR · {ratesInfo.LKR?.toFixed(2)} LKR.
            Click "Fetch live rates now" then Save to use them.
          </div>
        )}
      </div>

      <button style={styles.primaryBtn} onClick={doSave}>
        <Check size={18} /> {savedFlash ? "Saved!" : "Save Settings"}
      </button>
    </div>
  );
}

function DocHeader({ subtitle }) {
  const s = DOC_SETTINGS;
  return (
    <div className="dn-doc-header" style={doc.header}>
      <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
        <Logo size={54} />
        <div>
          <div style={doc.brandName}>{s.companyName || "DN TRAVELS"}</div>
          <div style={doc.brandTag}>{s.tagline || "Your Maldivian Connection"}</div>
        </div>
      </div>
      <div style={{ textAlign: "right" }}>
        <div style={doc.docType}>{subtitle}</div>
        {s.address && <div style={doc.contact}>{s.address}</div>}
        {s.phone && <div style={doc.contact}>{s.phone}</div>}
        {s.email && <div style={doc.contact}>{s.email}</div>}
        {s.website && <div style={doc.contact}>{s.website}</div>}
      </div>
    </div>
  );
}

function Invoice({ rec }) {
  const c = rec.currency;
  return (
    <div className="dn-doc-page" style={doc.page}>
      <DocHeader subtitle="INVOICE" />
      <div className="dn-meta-grid" style={doc.metaGrid}>
        <div>
          <div style={doc.metaLabel}>Billed To</div>
          <div style={doc.metaStrong}>{rec.customerName}</div>
          {rec.customerEmail && <div style={doc.metaLine}>{rec.customerEmail}</div>}
          {rec.customerPhone && <div style={doc.metaLine}>{rec.customerPhone}</div>}
        </div>
        <div style={{ textAlign: "right" }}>
          <div style={doc.metaLabel}>Booking Reference</div>
          <div style={doc.refBig}>{rec.ref}</div>
          <div style={doc.metaLine}>Issued: {new Date(rec.issuedAt).toLocaleDateString()}</div>
          {rec.eticket && <div style={doc.metaLine}>E-ticket: {rec.eticket}</div>}
        </div>
      </div>

      <table style={doc.table}>
        <thead>
          <tr>
            <th style={{ ...doc.th, textAlign: "left" }}>Description</th>
            <th style={doc.th}>Passenger</th>
            <th style={{ ...doc.th, textAlign: "right" }}>Amount</th>
          </tr>
        </thead>
        <tbody>
          {rec.segments.map((s, i) => (
            <tr key={i}>
              <td style={doc.td}>
                <b>{s.route || `${s.depAirport} → ${s.arrAirport}`}</b><br />
                <span style={{ color: "#5b7185", fontSize: 12 }}>
                  {s.airline} {s.flightNo} · {s.cls} · {s.depDate}
                </span>
              </td>
              <td style={{ ...doc.td, textAlign: "center" }}>{rec.paxName}</td>
              <td style={{ ...doc.td, textAlign: "right" }}>
                {i === 0 ? money(rec.sellPrice, c) : "—"}
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      <div style={doc.totalsWrap}>
        <div style={doc.totalsBox}>
          <div style={doc.totalRow}><span>Subtotal</span><span>{money(rec.sellPrice, c)}</span></div>
          <div style={{ ...doc.totalRow, ...doc.grandTotal }}>
            <span>Total Due</span><span>{money(rec.sellPrice, c)}</span>
          </div>
        </div>
      </div>

      {rec.notes && <div style={doc.notes}><b>Notes: </b>{rec.notes}</div>}
      <div style={doc.footer}>
        Thank you for booking with DN Travels — Your Maldivian Connection.<br />
        This is a computer-generated invoice.
      </div>
    </div>
  );
}

function Itinerary({ rec }) {
  return (
    <div className="dn-doc-page" style={doc.page}>
      <DocHeader subtitle="ITINERARY" />
      <div style={doc.itinTop}>
        <div>
          <div style={doc.metaLabel}>Passenger</div>
          <div style={doc.metaStrong}>{rec.paxName}</div>
          <div style={doc.metaLine}>Booking Ref: <b>{rec.ref}</b></div>
        </div>
        <div style={{ textAlign: "right" }}>
          {rec.eticket && <div style={doc.metaLine}>E-ticket: {rec.eticket}</div>}
          {rec.airlineRef && <div style={doc.metaLine}>Airline Ref: {rec.airlineRef}</div>}
          <div style={doc.metaLine}>Issued: {new Date(rec.issuedAt).toLocaleDateString()}</div>
        </div>
      </div>

      {rec.segments.map((s, i) => (
        <div key={i} style={doc.flightCard}>
          <div style={doc.flightHead}>
            <span style={{ fontWeight: 800, color: NAVY }}>{s.route || `${s.depAirport} → ${s.arrAirport}`}</span>
            <span style={doc.flightAirline}>{s.airline} {s.flightNo} · {s.cls}</span>
          </div>
          <div style={doc.flightBody}>
            <div style={doc.fSegment}>
              <div style={doc.fTime}>{s.depTime || "--:--"}</div>
              <div style={doc.fLabel}>Departure</div>
              <div style={doc.fAirport}>{s.depAirport}</div>
              <div style={doc.fDate}>{s.depDate}</div>
            </div>
            <div style={doc.fArrow}><Plane size={18} color={OCEAN} /></div>
            <div style={{ ...doc.fSegment, textAlign: "right" }}>
              <div style={doc.fTime}>{s.arrTime || "--:--"}</div>
              <div style={doc.fLabel}>Arrival</div>
              <div style={doc.fAirport}>{s.arrAirport}</div>
              <div style={doc.fDate}>{s.arrDate}</div>
            </div>
          </div>
        </div>
      ))}

      <div style={doc.baggage}>
        <div style={doc.metaLabel}>Baggage Allowance</div>
        <div style={{ fontSize: 13, color: "#33485c" }}>{rec.baggage}</div>
      </div>

      <div style={doc.important}>
        <b>Important:</b> Please carry a valid ID and arrive at the airport at least 3 hours prior to departure.
        Tickets must be used in the sequence shown above.
      </div>

      <div style={doc.footer}>DN Travels — Your Maldivian Connection · info@dntravels.mv</div>
    </div>
  );
}

// ============================================================
// Small UI pieces
// ============================================================
function Quotation({ rec }) {
  return (
    <div className="dn-doc-page" style={doc.page}>
      <DocHeader subtitle="QUOTATION" />
      <div className="dn-meta-grid" style={doc.metaGrid}>
        <div>
          <div style={doc.metaLabel}>Prepared For</div>
          <div style={doc.metaStrong}>{rec.customerName}</div>
          {rec.customerEmail && <div style={doc.metaLine}>{rec.customerEmail}</div>}
          {rec.customerPhone && <div style={doc.metaLine}>{rec.customerPhone}</div>}
        </div>
        <div style={{ textAlign: "right" }}>
          <div style={doc.metaLabel}>Quotation No.</div>
          <div style={doc.refBig}>{rec.ref}</div>
          <div style={doc.metaLine}>Date: {new Date(rec.issuedAt).toLocaleDateString()}</div>
          {rec.validUntil && <div style={doc.metaLine}>Valid until: {rec.validUntil}</div>}
        </div>
      </div>

      <table style={doc.table}>
        <thead>
          <tr><th style={{ ...doc.th, textAlign: "left" }}>Description</th></tr>
        </thead>
        <tbody>
          <tr>
            <td style={doc.td}>
              {rec.description && <div style={{ marginBottom: rec.includeFlights ? 10 : 0 }}>{rec.description}</div>}
              {rec.includeFlights && rec.segments.filter((s) => s.route || s.flightNo).map((s, i) => (
                <div key={i} style={{ fontSize: 13, color: "#33485c", marginBottom: 6 }}>
                  <b>{s.route || `${s.depAirport} → ${s.arrAirport}`}</b>
                  {(s.airline || s.flightNo) && <span style={{ color: "#5b7185" }}> — {s.airline} {s.flightNo}</span>}
                  {(s.depDate || s.depTime) && <span style={{ color: "#5b7185" }}> · {s.depDate} {s.depTime}</span>}
                </div>
              ))}
              {!rec.description && !rec.includeFlights && <span style={{ color: "#9aaabb" }}>Air ticket quotation</span>}
            </td>
          </tr>
        </tbody>
      </table>

      <div style={doc.totalsWrap}>
        <div style={{ width: 320 }}>
          <div style={doc.totalRow}><span>Price (USD)</span><span style={{ fontWeight: 700, color: NAVY }}>{money(rec.sellPrice, "USD")}</span></div>
          <div style={doc.totalRow}><span>Price (MVR)</span><span style={{ fontWeight: 700, color: NAVY }}>{money(rec.sellPrice, "MVR")}</span></div>
          <div style={{ ...doc.totalRow, ...doc.grandTotal }}><span>Price (LKR)</span><span>{money(rec.sellPrice, "LKR")}</span></div>
        </div>
      </div>
      <div style={doc.fxLine}>Currency conversions are indicative and based on current exchange rates.</div>

      {rec.notes && <div style={doc.notes}><b>Notes: </b>{rec.notes}</div>}
      <div style={doc.footer}>
        This is a quotation only and does not constitute a confirmed booking.<br />
        Please contact us to confirm and issue your ticket. Thank you for considering DN Travels.
      </div>
    </div>
  );
}

function Logo({ size = 44 }) {
  return (
    <div style={{ width: size, height: size, flexShrink: 0 }}>
      <svg viewBox="0 0 100 100" width={size} height={size}>
        <defs>
          <linearGradient id="dnGrad" x1="0" y1="0" x2="1" y2="1">
            <stop offset="0%" stopColor={SKY} />
            <stop offset="100%" stopColor={NAVY} />
          </linearGradient>
        </defs>
        <circle cx="50" cy="50" r="46" fill="none" stroke="url(#dnGrad)" strokeWidth="3" />
        <polygon points="50,6 56,46 50,40 44,46" fill={NAVY} />
        <polygon points="50,94 56,54 50,60 44,54" fill={OCEAN} />
        <polygon points="6,50 46,44 40,50 46,56" fill={OCEAN} />
        <polygon points="94,50 54,44 60,50 54,56" fill={NAVY} />
        <text x="50" y="58" textAnchor="middle" fontFamily="Sora, sans-serif"
          fontWeight="800" fontSize="30" fill="url(#dnGrad)">DN</text>
      </svg>
    </div>
  );
}

const SectionTitle = ({ icon, children }) => (
  <div style={styles.sectionTitle}>{icon} {children}</div>
);
const Group = ({ label, children }) => (
  <div style={styles.group}>
    <div style={styles.groupLabel}>{label}</div>
    {children}
  </div>
);
const Row = ({ children }) => <div style={styles.row}>{children}</div>;

// Quick links to common airline / OTA booking sites — opens each in a new tab.
const BOOKING_SITES = [
  { name: "Trip.com", url: "https://www.trip.com/flights/", color: "#287dfa" },
  { name: "SriLankan", url: "https://www.srilankan.com/", color: "#1f4e79" },
  { name: "Gulf Air", url: "https://www.gulfair.com/", color: "#9b2335" },
  { name: "flydubai", url: "https://www.flydubai.com/", color: "#fb5d2c" },
  { name: "FITS Air", url: "https://www.fitsair.com/", color: "#0f3d6e" },
  { name: "Emirates", url: "https://www.emirates.com/", color: "#d71920" },
];
const BookingSites = () => (
  <div style={styles.sitesBox}>
    <div style={styles.importHead}>
      <Globe size={18} color={OCEAN} />
      <div>
        <div style={styles.importTitle}>Search & Book on</div>
        <div style={styles.importSub}>Open an airline / OTA, complete the booking, then upload the PDF</div>
      </div>
    </div>
    <div style={styles.sitesGrid}>
      {BOOKING_SITES.map((s) => (
        <a key={s.name} href={s.url} target="_blank" rel="noopener noreferrer"
          style={{ ...styles.siteBtn, borderColor: s.color, color: s.color }}>
          <ExternalLink size={13} /> {s.name}
        </a>
      ))}
    </div>
  </div>
);

// Screenshot uploader — pick or paste an image; AI extracts the booking details.
const ScreenshotUploader = ({ target, busy, msg, onPick }) => {
  // Allow paste-from-clipboard while the panel is mounted
  React.useEffect(() => {
    const handler = (e) => {
      if (busy) return;
      const items = e.clipboardData?.items || [];
      for (const it of items) {
        if (it.type && it.type.startsWith("image/")) {
          const file = it.getAsFile();
          if (file) { onPick(file); e.preventDefault(); return; }
        }
      }
    };
    window.addEventListener("paste", handler);
    return () => window.removeEventListener("paste", handler);
  }, [busy, onPick]);

  return (
    <div style={{ ...styles.importBox, background: "linear-gradient(135deg,#f4f0ff,#f8f5ff)", borderColor: "#c8b6ff" }}>
      <div style={styles.importHead}>
        <ImageIcon size={18} color="#6e3bd1" />
        <div>
          <div style={styles.importTitle}>Auto-fill from screenshot (AI)</div>
          <div style={styles.importSub}>Upload or paste (Ctrl+V) any booking page or search result</div>
        </div>
      </div>
      <label style={{ ...styles.importBtn, background: "#6e3bd1", width: "100%" }}>
        {busy ? <Loader2 size={16} className="spin" /> : <Upload size={16} />}
        {busy ? "Reading screenshot…" : "Choose screenshot"}
        <input type="file" accept="image/*" style={{ display: "none" }}
          disabled={!!busy}
          onChange={(e) => { onPick(e.target.files?.[0]); e.target.value = ""; }} />
      </label>
      {msg && msg.target === target && (
        <div style={{ ...styles.importMsg, color: msg.ok ? "#1a7a4c" : "#c0392b" }}>
          {msg.ok ? <Check size={14} /> : <X size={14} />} {msg.text}
        </div>
      )}
    </div>
  );
};
const Field = ({ label, value, onChange, placeholder }) => (
  <div style={styles.field}>
    <label style={styles.label}>{label}</label>
    <input style={styles.input} value={value} placeholder={placeholder}
      onChange={(e) => onChange(e.target.value)} />
  </div>
);

// Date picker — stores readable text ("May 29, 2026") but lets user pick from a calendar.
const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
function toISO(dateStr) {
  if (!dateStr) return "";
  // Already ISO?
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return dateStr;
  // Try "Month D, YYYY"
  const m = dateStr.match(/([A-Za-z]+)\s+(\d{1,2}),?\s*(\d{4})/);
  if (m) {
    const mi = MONTHS.findIndex((x) => x.toLowerCase().startsWith(m[1].toLowerCase().slice(0,3)));
    if (mi >= 0) return `${m[3]}-${String(mi+1).padStart(2,"0")}-${String(m[2]).padStart(2,"0")}`;
  }
  return "";
}
function fromISO(iso) {
  if (!iso || !/^\d{4}-\d{2}-\d{2}$/.test(iso)) return iso || "";
  const [y, m, d] = iso.split("-").map(Number);
  return `${MONTHS[m-1]} ${d}, ${y}`;
}
const DateField = ({ label, value, onChange }) => (
  <div style={styles.field}>
    <label style={styles.label}>{label}</label>
    <input
      type="date"
      style={styles.input}
      value={toISO(value)}
      onChange={(e) => onChange(fromISO(e.target.value))}
    />
  </div>
);
const PriceLine = ({ label, value, accent, big }) => (
  <div style={styles.priceLine}>
    <span style={{ color: big ? NAVY : "#5b7185", fontWeight: big ? 700 : 500, fontSize: big ? 15 : 13 }}>{label}</span>
    <span style={{
      color: accent ? OCEAN : NAVY,
      fontWeight: big ? 800 : 600,
      fontSize: big ? 22 : 14,
      fontFamily: "Sora, sans-serif",
    }}>{value}</span>
  </div>
);

function FontInjector() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700;800&family=Manrope:wght@400;500;600;700&display=swap');
      * { box-sizing: border-box; }
      input:focus { outline: none; border-color: ${OCEAN} !important; box-shadow: 0 0 0 3px rgba(29,111,184,.12); }
      .spin { animation: sp 1.2s linear infinite; }
      @keyframes sp { to { transform: rotate(360deg); } }
      ::-webkit-scrollbar { width: 9px; height: 9px; }
      ::-webkit-scrollbar-thumb { background: #c5d6e4; border-radius: 6px; }
      @media (max-width: 760px) {
        .dn-grid { grid-template-columns: 1fr !important; }
        .dn-doc-page { width: 100% !important; padding: 22px 16px !important; }
        .dn-header { flex-direction: column; align-items: flex-start !important; gap: 12px; }
        .dn-nav { width: 100%; overflow-x: auto; flex-wrap: nowrap !important; }
        .dn-nav button { white-space: nowrap; }
        .dn-doc-header { flex-direction: column; gap: 12px; }
        .dn-meta-grid { flex-direction: column; gap: 14px; }
        .dn-report-stats { grid-template-columns: 1fr !important; }
      }
    `}</style>
  );
}

// ============================================================
// Styles
// ============================================================
const styles = {
  shell: { fontFamily: "Manrope, sans-serif", background: PAPER, minHeight: "100vh", color: "#1a2733" },
  header: {
    display: "flex", justifyContent: "space-between", alignItems: "center",
    padding: "16px 26px", background: "#fff", borderBottom: `1px solid #e3edf5`,
    position: "sticky", top: 0, zIndex: 20, flexWrap: "wrap", gap: 12,
  },
  brand: { display: "flex", alignItems: "center", gap: 12 },
  brandName: { fontFamily: "Sora, sans-serif", fontWeight: 800, fontSize: 20, color: NAVY, letterSpacing: 1 },
  brandTag: { fontSize: 11.5, color: OCEAN, fontWeight: 600, letterSpacing: .3 },
  nav: { display: "flex", gap: 8 },
  navBtn: {
    display: "flex", alignItems: "center", gap: 7, padding: "9px 16px",
    border: `1px solid #dbe7f0`, background: "#fff", borderRadius: 10,
    fontFamily: "Manrope, sans-serif", fontWeight: 600, fontSize: 14, color: "#5b7185",
    cursor: "pointer", transition: "all .15s",
  },
  navBtnActive: { background: NAVY, color: "#fff", borderColor: NAVY },
  badge: { background: "rgba(255,255,255,.25)", padding: "1px 8px", borderRadius: 20, fontSize: 12 },
  main: { maxWidth: 1180, margin: "0 auto", padding: "26px 22px 60px" },
  grid: { display: "grid", gridTemplateColumns: "1.35fr 1fr", gap: 22, alignItems: "start" },
  card: { background: "#fff", border: "1px solid #e3edf5", borderRadius: 16, padding: 22, boxShadow: "0 1px 3px rgba(13,59,102,.04)" },
  sectionTitle: { display: "flex", alignItems: "center", gap: 8, fontFamily: "Sora, sans-serif", fontWeight: 700, fontSize: 16, color: NAVY, marginBottom: 18 },
  group: { marginBottom: 20 },
  importBox: { background: "linear-gradient(135deg,#eef6fc,#f5f9fc)", border: `1.5px dashed ${SKY}`, borderRadius: 14, padding: 16, marginBottom: 22 },
  importHead: { display: "flex", alignItems: "center", gap: 11, marginBottom: 12 },
  importTitle: { fontFamily: "Sora, sans-serif", fontWeight: 700, fontSize: 14.5, color: NAVY },
  importSub: { fontSize: 12, color: "#5b7185", marginTop: 1 },
  importBtn: { display: "inline-flex", alignItems: "center", justifyContent: "center", gap: 8, padding: "10px 18px", background: OCEAN, color: "#fff", borderRadius: 10, fontWeight: 600, fontSize: 13.5, cursor: "pointer", border: "none", flex: 1 },
  importBtnRow: { display: "flex", gap: 10 },
  importBtnBusy: { background: "#7fb4dd", cursor: "wait" },
  importMsg: { display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, fontWeight: 600, marginTop: 11, lineHeight: 1.4 },
  sitesBox: { background: "#fff", border: "1px solid #e3edf5", borderRadius: 14, padding: 16, marginBottom: 22 },
  sitesGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(120px, 1fr))", gap: 8 },
  siteBtn: { display: "inline-flex", alignItems: "center", justifyContent: "center", gap: 6, padding: "9px 12px", background: "#fff", border: "1.5px solid", borderRadius: 9, fontWeight: 700, fontSize: 12.5, textDecoration: "none", transition: "all .15s" },
  groupLabel: { fontSize: 11.5, fontWeight: 700, color: OCEAN, textTransform: "uppercase", letterSpacing: .6, marginBottom: 10 },
  row: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 },
  field: { marginBottom: 12 },
  label: { display: "block", fontSize: 12.5, fontWeight: 600, color: "#5b7185", marginBottom: 5 },
  input: {
    width: "100%", padding: "10px 12px", border: "1px solid #d9e6f0", borderRadius: 9,
    fontFamily: "Manrope, sans-serif", fontSize: 14, color: "#1a2733", transition: "all .15s",
  },
  segBox: { border: "1px solid #e3edf5", borderRadius: 12, padding: 14, marginBottom: 12, background: "#fbfdff" },
  tripTypeRow: { display: "flex", gap: 8, marginBottom: 14 },
  tripBtn: { flex: 1, padding: "10px", border: "1px solid #d9e6f0", background: "#fff", borderRadius: 9, fontWeight: 700, fontSize: 13.5, color: "#5b7185", cursor: "pointer" },
  tripBtnActive: { background: OCEAN, color: "#fff", borderColor: OCEAN },
  segHead: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 },
  iconBtnSm: { display: "grid", placeItems: "center", width: 30, height: 30, border: "1px solid #e3edf5", background: "#fff", borderRadius: 8, color: "#c0392b", cursor: "pointer" },
  addSegBtn: { display: "flex", alignItems: "center", justifyContent: "center", gap: 6, width: "100%", padding: "10px", border: `1.5px dashed ${SKY}`, background: "rgba(63,159,224,.06)", borderRadius: 10, color: OCEAN, fontWeight: 600, fontSize: 13.5, cursor: "pointer" },
  currencyRow: { display: "flex", gap: 8, margin: "14px 0" },
  curBtn: { flex: 1, padding: "8px", border: "1px solid #d9e6f0", background: "#fff", borderRadius: 9, fontWeight: 700, fontSize: 13, color: "#5b7185", cursor: "pointer" },
  curBtnActive: { background: OCEAN, color: "#fff", borderColor: OCEAN },
  priceBreak: { background: "#f5f9fc", borderRadius: 12, padding: "14px 16px", marginTop: 6 },
  priceLine: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "5px 0" },
  priceDivider: { height: 1, background: "#dde9f2", margin: "8px 0" },
  fxEdit: { display: "flex", alignItems: "center", gap: 8, marginTop: 12, flexWrap: "wrap" },
  fxLabel: { fontSize: 12.5, fontWeight: 600, color: "#5b7185" },
  fxInput: { width: 90, padding: "7px 10px", border: "1px solid #d9e6f0", borderRadius: 8, fontFamily: "Manrope, sans-serif", fontSize: 14, color: NAVY, fontWeight: 700 },
  fxCur: { fontSize: 13, fontWeight: 700, color: NAVY },
  fxRefresh: { display: "inline-flex", alignItems: "center", gap: 5, padding: "7px 11px", background: "#eef6fc", border: "1px solid #cfe3f3", borderRadius: 8, color: OCEAN, fontWeight: 600, fontSize: 12, cursor: "pointer" },
  fxHint: { fontSize: 11.5, color: "#8499aa", marginTop: 8, lineHeight: 1.4 },
  primaryBtn: { display: "flex", alignItems: "center", justifyContent: "center", gap: 9, padding: "15px", background: NAVY, color: "#fff", border: "none", borderRadius: 12, fontFamily: "Sora, sans-serif", fontWeight: 700, fontSize: 15.5, cursor: "pointer", boxShadow: "0 4px 14px rgba(13,59,102,.22)" },
  ghostBtn: { padding: "11px", background: "transparent", color: "#8499aa", border: "1px solid #d9e6f0", borderRadius: 10, fontWeight: 600, cursor: "pointer" },
  hint: { fontSize: 12, color: "#8499aa", lineHeight: 1.5, textAlign: "center" },
  // doc toolbar
  docToolbar: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18, flexWrap: "wrap", gap: 12 },
  docTabs: { display: "flex", gap: 6, background: "#fff", padding: 5, borderRadius: 12, border: "1px solid #e3edf5" },
  docTab: { padding: "8px 22px", border: "none", background: "transparent", borderRadius: 8, fontWeight: 700, fontSize: 14, color: "#5b7185", cursor: "pointer" },
  docTabActive: { background: NAVY, color: "#fff" },
  toolBtn: { display: "flex", alignItems: "center", gap: 7, padding: "10px 16px", background: OCEAN, color: "#fff", border: "none", borderRadius: 10, fontWeight: 600, fontSize: 13.5, cursor: "pointer" },
  toolBtnGhost: { display: "flex", alignItems: "center", gap: 7, padding: "10px 16px", background: "#fff", color: "#5b7185", border: "1px solid #d9e6f0", borderRadius: 10, fontWeight: 600, fontSize: 13.5, cursor: "pointer" },
  docWrap: { display: "flex", justifyContent: "center" },
  // history
  histBar: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18, gap: 12, flexWrap: "wrap" },
  searchBox: { display: "flex", alignItems: "center", gap: 8, background: "#fff", border: "1px solid #d9e6f0", borderRadius: 11, padding: "9px 14px", flex: 1, maxWidth: 380 },
  searchInput: { border: "none", outline: "none", fontFamily: "Manrope, sans-serif", fontSize: 14, flex: 1, background: "transparent" },
  excelBtn: { display: "flex", alignItems: "center", gap: 8, padding: "11px 20px", background: "#1a7a4c", color: "#fff", border: "none", borderRadius: 11, fontWeight: 700, fontSize: 14, cursor: "pointer", boxShadow: "0 3px 10px rgba(26,122,76,.2)" },
  emailBtn: { display: "flex", alignItems: "center", gap: 7, padding: "10px 16px", background: "#1a7a4c", color: "#fff", border: "none", borderRadius: 10, fontWeight: 600, fontSize: 13.5, cursor: "pointer" },
  reportRange: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 },
  quickRow: { display: "flex", gap: 8, flexWrap: "wrap", margin: "4px 0 18px" },
  quickBtn: { padding: "8px 14px", background: "#fff", border: "1px solid #d9e6f0", borderRadius: 9, fontWeight: 600, fontSize: 13, color: OCEAN, cursor: "pointer" },
  statGrid: { display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12, marginTop: 6 },
  statBox: { background: "#f5f9fc", border: "1px solid #e3edf5", borderRadius: 12, padding: "16px 18px" },
  statLabel: { fontSize: 11.5, fontWeight: 700, color: "#5b7185", textTransform: "uppercase", letterSpacing: .4 },
  statVal: { fontFamily: "Sora, sans-serif", fontWeight: 800, fontSize: 24, color: NAVY, marginTop: 6 },
  costNote: { fontSize: 12.5, color: "#8499aa", marginTop: 12 },
  empty: { textAlign: "center", padding: "70px 20px", color: "#8499aa", background: "#fff", borderRadius: 16, border: "1px dashed #d9e6f0" },
  histList: { display: "flex", flexDirection: "column", gap: 10 },
  histRow: { display: "flex", alignItems: "center", gap: 16, background: "#fff", border: "1px solid #e3edf5", borderRadius: 13, padding: "15px 18px", cursor: "pointer", transition: "all .15s" },
  refTag: { fontFamily: "Sora, sans-serif", fontWeight: 800, fontSize: 13, color: "#fff", background: NAVY, padding: "6px 11px", borderRadius: 8, letterSpacing: .5 },
  histName: { fontWeight: 700, fontSize: 15, color: NAVY },
  histSub: { fontSize: 12.5, color: "#5b7185", marginTop: 2, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" },
  histMeta: { fontSize: 11.5, color: "#9aaabb", marginTop: 2 },
  histPrice: { fontFamily: "Sora, sans-serif", fontWeight: 800, fontSize: 16, color: NAVY },
  histMargin: { fontSize: 11.5, color: "#1a7a4c", fontWeight: 600 },
};

const doc = {
  page: { width: 720, background: "#fff", padding: "40px 44px", boxShadow: "0 6px 30px rgba(13,59,102,.1)", borderRadius: 6, fontFamily: "Manrope, sans-serif", color: "#1a2733" },
  header: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", borderBottom: `3px solid ${NAVY}`, paddingBottom: 18, marginBottom: 22 },
  brandName: { fontFamily: "Sora, sans-serif", fontWeight: 800, fontSize: 22, color: NAVY, letterSpacing: 1 },
  brandTag: { fontSize: 12, color: OCEAN, fontWeight: 600 },
  docType: { fontFamily: "Sora, sans-serif", fontWeight: 800, fontSize: 26, color: NAVY, letterSpacing: 2 },
  contact: { fontSize: 11.5, color: "#7d93a8", marginTop: 2 },
  metaGrid: { display: "flex", justifyContent: "space-between", marginBottom: 24 },
  metaLabel: { fontSize: 10.5, fontWeight: 700, color: OCEAN, textTransform: "uppercase", letterSpacing: .8, marginBottom: 5 },
  metaStrong: { fontWeight: 700, fontSize: 16, color: NAVY },
  metaLine: { fontSize: 12.5, color: "#5b7185", marginTop: 2 },
  refBig: { fontFamily: "Sora, sans-serif", fontWeight: 800, fontSize: 19, color: NAVY },
  table: { width: "100%", borderCollapse: "collapse", marginBottom: 4 },
  th: { background: NAVY, color: "#fff", padding: "11px 14px", fontSize: 12, fontWeight: 600, textAlign: "center" },
  td: { padding: "13px 14px", borderBottom: "1px solid #eef3f8", fontSize: 13.5, verticalAlign: "top" },
  totalsWrap: { display: "flex", justifyContent: "flex-end", marginTop: 14 },
  totalsBox: { width: 280 },
  totalRow: { display: "flex", justifyContent: "space-between", padding: "8px 4px", fontSize: 14, color: "#5b7185" },
  grandTotal: { borderTop: `2px solid ${NAVY}`, marginTop: 4, paddingTop: 12, fontFamily: "Sora, sans-serif", fontWeight: 800, fontSize: 18, color: NAVY },
  notes: { marginTop: 22, padding: "12px 16px", background: "#f5f9fc", borderRadius: 8, fontSize: 12.5, color: "#5b7185" },
  fxLine: { fontSize: 11, color: "#9aaabb", marginTop: 8, textAlign: "right" },
  footer: { marginTop: 30, paddingTop: 16, borderTop: "1px solid #eef3f8", textAlign: "center", fontSize: 11.5, color: "#9aaabb", lineHeight: 1.6 },
  itinTop: { display: "flex", justifyContent: "space-between", marginBottom: 22 },
  flightCard: { border: "1px solid #e3edf5", borderRadius: 12, marginBottom: 14, overflow: "hidden" },
  flightHead: { display: "flex", justifyContent: "space-between", alignItems: "center", background: "#f5f9fc", padding: "11px 18px", fontSize: 14 },
  flightAirline: { fontSize: 12.5, color: OCEAN, fontWeight: 600 },
  flightBody: { display: "flex", alignItems: "center", padding: "18px" },
  fSegment: { flex: 1 },
  fTime: { fontFamily: "Sora, sans-serif", fontWeight: 800, fontSize: 24, color: NAVY },
  fLabel: { fontSize: 10.5, fontWeight: 700, color: OCEAN, textTransform: "uppercase", letterSpacing: .6, marginTop: 2 },
  fAirport: { fontSize: 12.5, color: "#33485c", marginTop: 5, lineHeight: 1.3 },
  fDate: { fontSize: 11.5, color: "#9aaabb", marginTop: 3 },
  fArrow: { padding: "0 22px" },
  baggage: { marginTop: 8, padding: "14px 18px", background: "#f5f9fc", borderRadius: 10, marginBottom: 14 },
  important: { fontSize: 11.5, color: "#7d93a8", lineHeight: 1.6, padding: "0 4px" },
};
