// Vercel Serverless Function — keeps the Gemini API key safe on the server.
// The browser sends extracted PDF text here; this calls Gemini and returns clean JSON.

export default async function handler(req, res) {
  // Allow the browser to call this
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const key = process.env.GEMINI_API_KEY;
  if (!key) return res.status(500).json({ error: "Server missing GEMINI_API_KEY" });

  try {
    const { receiptText = "", itineraryText = "", imageBase64 = "", imageMime = "image/png" } = req.body || {};
    if (!receiptText && !itineraryText && !imageBase64) {
      return res.status(400).json({ error: "No document text or image provided" });
    }

    const prompt = `You are extracting flight booking details from airline or travel-agency documents OR from a SCREENSHOT of a search result, booking page, itinerary, or confirmation.
The source may come from Trip.com, Emirates, SriLankan Airlines, Gulf Air, flydubai, FITS Air,
or any other airline or booking site. The input may be raw text (extracted from a PDF) and/or
an image (screenshot). Read the meaning, not a fixed format.
Return ONLY a JSON object (no markdown, no backticks, no explanation) with this exact shape:

{
  "customerName": string,        // the contact/billing name from the receipt, or "" if unknown
  "customerEmail": string,       // email if a real one is present (ignore masked ones with *), else ""
  "paxName": string,             // the passenger name as printed on the ticket
  "eticket": string,             // e-ticket number, e.g. "072-9411581236"
  "airlineRef": string,          // airline booking reference / PNR, e.g. "NMYDEB", else ""
  "tripComPrice": number,        // the grand TOTAL price (number, no currency symbol)
  "segments": [                  // one object per flight segment, in order
    {
      "route": string,           // e.g. "Malé - Colombo"
      "airline": string,         // e.g. "Gulf Air"
      "flightNo": string,        // e.g. "GF144"
      "cls": string,             // travel class, e.g. "Economy"
      "depTime": string,         // 24h time, e.g. "07:30"
      "depDate": string,         // e.g. "May 29, 2026"
      "depAirport": string,      // full departure airport name
      "arrTime": string,         // e.g. "09:40"
      "arrDate": string,         // e.g. "May 29, 2026"
      "arrAirport": string       // full arrival airport name
    }
  ]
}

Rules:
- If a field is unknown, use "" for strings, 0 for tripComPrice, or [] for segments.
- tripComPrice must be the grand TOTAL the buyer paid (including taxes & fees), not just the base fare.
  Look for labels like "Total", "Total Amount", "Grand Total", "Amount Paid", "Total Fare".
- The price may be in any currency (USD, MVR, LKR, AED, etc). Return only the numeric amount.
- "airlineRef" is the booking reference / PNR / confirmation code (e.g. "NMYDEB"), whatever the airline calls it.
- For multi-leg trips, return every flight segment in travel order.
- Different airlines label things differently (e.g. "Passenger", "Traveller", "Guest"); map them to the right field.
- For screenshots showing only search results (no passenger info yet), return whatever IS visible (flights, price)
  and leave unseen fields as "".
- Return strictly valid JSON and nothing else.

=== RECEIPT TEXT ===
${receiptText || "(none provided)"}

=== ITINERARY TEXT ===
${itineraryText || "(none provided)"}`;

    // Build the Gemini parts. Image goes alongside the text prompt.
    const parts = [{ text: prompt }];
    if (imageBase64) {
      parts.push({ inline_data: { mime_type: imageMime, data: imageBase64 } });
    }

    const geminiRes = await fetch(
      "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent",
      {
        method: "POST",
        headers: { "Content-Type": "application/json", "x-goog-api-key": key },
        body: JSON.stringify({
          contents: [{ parts }],
          generationConfig: { temperature: 0, responseMimeType: "application/json" },
        }),
      }
    );

    if (!geminiRes.ok) {
      const errText = await geminiRes.text();
      return res.status(502).json({ error: "Gemini error", detail: errText.slice(0, 500) });
    }

    const data = await geminiRes.json();
    const raw = data?.candidates?.[0]?.content?.parts?.[0]?.text || "";
    let parsed;
    try {
      parsed = JSON.parse(raw);
    } catch {
      // strip any stray code fences just in case
      const cleaned = raw.replace(/```json|```/g, "").trim();
      parsed = JSON.parse(cleaned);
    }
    return res.status(200).json(parsed);
  } catch (e) {
    return res.status(500).json({ error: "Parse failed", detail: String(e).slice(0, 300) });
  }
}
