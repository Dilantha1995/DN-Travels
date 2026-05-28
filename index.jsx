import React, { useState, useEffect, useRef, useCallback } from "react";
import {
  Plane, Plus, History, Download, Printer, Trash2, FileText,
  X, Check, Compass, Search, ChevronRight, RefreshCw, Edit3
} from "lucide-react";
import * as XLSX from "xlsx";

// ---------- Branding ----------
const NAVY = "#0d3b66";
const OCEAN = "#1d6fb8";
const SKY = "#3f9fe0";
const PAPER = "#f5f8fb";

const CURRENCIES = {
  USD: { symbol: "$", label: "USD", rate: 1 },
  MVR: { symbol: "MVR", label: "MVR", rate: 15.42 },     // editable below
  LKR: { symbol: "LKR", label: "LKR", rate: 300 },        // editable below
};

const MARKUP = 0.05; // 5%

// ---------- Helpers ----------
const pad = (n, len = 2) => String(n).padStart(len, "0");

function genBookingRef(seq) {
  // DN + YYMMDD + 3-digit sequence  ->  DN260528001
  const d = new Date();
  const stamp = `${pad(d.getFullYear() % 100)}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;
  return `DN${stamp}${pad(seq, 3)}`;
}

function money(amount, cur) {
  const c = CURRENCIES[cur];
  const val = (amount * c.rate);
  const fixed = val.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return cur === "USD" ? `${c.symbol}${fixed}` : `${c.symbol} ${fixed}`;
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
  segments: [blankSegment()],
  baggage: "Checked 25kg • Carry-on 6kg • 1 Personal item",
  notes: "",
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

  // ---- Load persisted data ----
  useEffect(() => {
    (async () => {
      try {
        const r = await window.storage.get("records");
        if (r && r.value) setRecords(JSON.parse(r.value));
      } catch (e) { /* none yet */ }
      try {
        const s = await window.storage.get("seq");
        if (s && s.value) setSeq(parseInt(s.value, 10));
      } catch (e) { /* none yet */ }
      setLoading(false);
    })();
  }, []);

  const persist = useCallback(async (newRecords, newSeq) => {
    try {
      await window.storage.set("records", JSON.stringify(newRecords));
      if (newSeq != null) await window.storage.set("seq", String(newSeq));
    } catch (e) {
      console.error("Storage error", e);
    }
  }, []);

  // ---- Derived pricing ----
  const basePrice = parseFloat(form.tripComPrice) || 0;
  const sellPrice = basePrice * (1 + MARKUP);
  const margin = sellPrice - basePrice;

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
      await persist(updated);
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
    await persist(newRecords, newSeq);
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
    await persist(newRecords);
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
  const handlePrint = () => {
    const node = printRef.current;
    if (!node) return;
    const w = window.open("", "_blank");
    w.document.write(`<!DOCTYPE html><html><head><title>${(viewing||saved).ref}</title>
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700;800&family=Manrope:wght@400;500;600;700&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        body{font-family:'Manrope',sans-serif;color:#1a2733;padding:32px;}
        @page{margin:14mm;}
      </style></head><body>${node.innerHTML}</body></html>`);
    w.document.close();
    setTimeout(() => { w.focus(); w.print(); }, 400);
  };

  const activeRecord = viewing || saved;

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
      <header style={styles.header}>
        <div style={styles.brand}>
          <Logo />
          <div>
            <div style={styles.brandName}>DN TRAVELS</div>
            <div style={styles.brandTag}>Your Maldivian Connection</div>
          </div>
        </div>
        <nav style={styles.nav}>
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
        </nav>
      </header>

      <main style={styles.main}>
        {/* ============ NEW TICKET ============ */}
        {tab === "new" && !saved && (
          <div style={styles.grid}>
            {/* --- Form --- */}
            <div style={styles.card}>
              <SectionTitle icon={<Edit3 size={16} />}>
                {editingId ? "Edit Booking" : "Issue New Ticket"}
              </SectionTitle>

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
                {form.segments.map((s, i) => (
                  <div key={i} style={styles.segBox}>
                    <div style={styles.segHead}>
                      <span style={{ fontWeight: 700, color: NAVY, fontSize: 13 }}>Segment {i + 1}</span>
                      {form.segments.length > 1 && (
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
                      <Field label="Dep. Date" value={s.depDate} onChange={(v) => setSeg(i, "depDate", v)} placeholder="May 29, 2026" />
                      <Field label="Dep. Time" value={s.depTime} onChange={(v) => setSeg(i, "depTime", v)} placeholder="07:30" />
                    </Row>
                    <Field label="Dep. Airport" value={s.depAirport} onChange={(v) => setSeg(i, "depAirport", v)} placeholder="Velana International Airport T1" />
                    <Row>
                      <Field label="Arr. Date" value={s.arrDate} onChange={(v) => setSeg(i, "arrDate", v)} placeholder="May 29, 2026" />
                      <Field label="Arr. Time" value={s.arrTime} onChange={(v) => setSeg(i, "arrTime", v)} placeholder="09:40" />
                    </Row>
                    <Field label="Arr. Airport" value={s.arrAirport} onChange={(v) => setSeg(i, "arrAirport", v)} placeholder="Colombo Bandaranaike International Airport" />
                  </div>
                ))}
                <button style={styles.addSegBtn} onClick={addSeg}>
                  <Plus size={15} /> Add another segment
                </button>
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
                  <div style={styles.fxNote}>
                    Converted at {CURRENCIES[form.currency].rate} {form.currency}/USD — adjust the rate in code if needed.
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
                <button style={styles.toolBtn} onClick={handlePrint}><Printer size={15} /> Print / Save PDF</button>
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
      </main>
    </div>
  );
}

// ============================================================
// Document components
// ============================================================
function DocHeader({ subtitle }) {
  return (
    <div style={doc.header}>
      <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
        <Logo size={54} />
        <div>
          <div style={doc.brandName}>DN TRAVELS</div>
          <div style={doc.brandTag}>Your Maldivian Connection</div>
        </div>
      </div>
      <div style={{ textAlign: "right" }}>
        <div style={doc.docType}>{subtitle}</div>
        <div style={doc.contact}>Malé, Republic of Maldives</div>
        <div style={doc.contact}>info@dntravels.mv</div>
      </div>
    </div>
  );
}

function Invoice({ rec }) {
  const c = rec.currency;
  return (
    <div style={doc.page}>
      <DocHeader subtitle="INVOICE" />
      <div style={doc.metaGrid}>
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
    <div style={doc.page}>
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
const Field = ({ label, value, onChange, placeholder }) => (
  <div style={styles.field}>
    <label style={styles.label}>{label}</label>
    <input style={styles.input} value={value} placeholder={placeholder}
      onChange={(e) => onChange(e.target.value)} />
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
  groupLabel: { fontSize: 11.5, fontWeight: 700, color: OCEAN, textTransform: "uppercase", letterSpacing: .6, marginBottom: 10 },
  row: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 },
  field: { marginBottom: 12 },
  label: { display: "block", fontSize: 12.5, fontWeight: 600, color: "#5b7185", marginBottom: 5 },
  input: {
    width: "100%", padding: "10px 12px", border: "1px solid #d9e6f0", borderRadius: 9,
    fontFamily: "Manrope, sans-serif", fontSize: 14, color: "#1a2733", transition: "all .15s",
  },
  segBox: { border: "1px solid #e3edf5", borderRadius: 12, padding: 14, marginBottom: 12, background: "#fbfdff" },
  segHead: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 },
  iconBtnSm: { display: "grid", placeItems: "center", width: 30, height: 30, border: "1px solid #e3edf5", background: "#fff", borderRadius: 8, color: "#c0392b", cursor: "pointer" },
  addSegBtn: { display: "flex", alignItems: "center", justifyContent: "center", gap: 6, width: "100%", padding: "10px", border: `1.5px dashed ${SKY}`, background: "rgba(63,159,224,.06)", borderRadius: 10, color: OCEAN, fontWeight: 600, fontSize: 13.5, cursor: "pointer" },
  currencyRow: { display: "flex", gap: 8, margin: "14px 0" },
  curBtn: { flex: 1, padding: "8px", border: "1px solid #d9e6f0", background: "#fff", borderRadius: 9, fontWeight: 700, fontSize: 13, color: "#5b7185", cursor: "pointer" },
  curBtnActive: { background: OCEAN, color: "#fff", borderColor: OCEAN },
  priceBreak: { background: "#f5f9fc", borderRadius: 12, padding: "14px 16px", marginTop: 6 },
  priceLine: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "5px 0" },
  priceDivider: { height: 1, background: "#dde9f2", margin: "8px 0" },
  fxNote: { fontSize: 11.5, color: "#8499aa", marginTop: 10, lineHeight: 1.4 },
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
