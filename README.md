# DN Travels — Ticket Booking Tool

Internal tool for issuing customer tickets (Trip.com cost + 5% markup),
generating branded invoices & itineraries, with auto booking references,
saved history, and Excel export.

## Run locally
```
npm install
npm run dev
```

## Deploy
Push to GitHub and import the repo in Vercel.
Framework preset: **Vite**. Build command: `npm run build`. Output dir: `dist`.

## Notes
- History is saved in the browser via localStorage (per device/browser).
- Edit exchange rates in `src/App.jsx` (CURRENCIES object).
