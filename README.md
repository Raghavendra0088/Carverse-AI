# CarVerse — DealerVerse AI

A gamified workflow platform for automotive dealerships. It turns a dealership's real
operational data into a role-based gamification system across all **7 departments**
(Sales, Finance, Accounts, EDP, Customer Care, PDI, RTO/Registration), while modeling
the complete used-car (C2B2C) buy/sell loop and the new-car sales lifecycle.

Built for the **"Dealership Gamification"** hackathon.

## Highlights

- **React 19 + Vite** SPA with 20+ screens and a custom design system, built on
  **Zustand**, **Tailwind CSS**, **Framer Motion**, **React Router**, and **Recharts**.
- **Node.js data pipeline** (PapaParse) that ingests a real dealership event log of
  **170,000+ rows** and precomputes a compact JSON the client loads instantly.
- **Anti-gaming scoring engine** — caps meaningful actions at 20 milestones and credits
  the *actual* actor's real department, rewarding genuine work over vanity metrics.
- **Full C2B2C used-car loop**: seller listing → PDI inspection → sales offer → payout →
  auto-priced inventory → buyer reservation → department-gated fulfillment → delivery.
- **11-stage new-car booking flow** with strict sequential, role-gated state transitions.
- **PIN + OTP authentication**, real-time event-driven missions, per-step document
  upload/verification, an email + support-ticket lifecycle, and an intent-based AI coach.
- All screens and E2E flows verified with **Playwright**.

## Tech Stack

React 19 · Vite · Zustand · Tailwind CSS · Framer Motion · React Router · Recharts ·
Node.js · PapaParse · Playwright

## Getting Started

```bash
cd dealerverse-ai
npm install
npm run data    # build src/data/dealerverse.json from the source CSVs
npm run dev     # start the Vite dev server
```

## Project Structure

```
CARVERSE/
├── dealerverse-ai/          # the main React + Vite application
│   ├── src/                 # components, pages, stores (Zustand), lib, data
│   └── scripts/build-data.mjs   # Node.js ETL: CSV event log → precomputed JSON
├── carverse files/          # source dataset (event log, employees, locations)
└── app.py, scoring_engine.py    # early Streamlit prototype (superseded)
```
