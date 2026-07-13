# DICT Abra — Project Monitoring System
### Consolidated Target vs. Accomplishment · FY 2025

A web application for monitoring the five DICT Abra provincial projects:
**ILCDB · DTC · SPARK · Cybersecurity · Free Wi-Fi**

---

## Features
- **Dashboard** — Overview of all 5 programs with completion rates per quarter
- **Encode Data** — Enter targets and accomplishments per indicator, per quarter
- **Full Report** — Consolidated table view of all programs and all quarters
- **Add/Remove Indicators** — Customize performance indicators per program
- **Notes & Remarks** — Add remarks per program per quarter
- **Export CSV** — Download the full report as a spreadsheet
- **Auto-save** — All data is saved automatically in the browser (localStorage)

---

## How to Run (Development)

### Requirements
- Node.js v16 or higher
- npm v7 or higher

### Steps

```bash
# 1. Install dependencies
npm install

# 2. Start the development server
npm start
```

The app will open at **http://localhost:3000**

---

## How to Build for Production

```bash
npm run build
```

This creates a `build/` folder you can deploy to any static hosting service
(e.g., Vercel, Netlify, GitHub Pages, or your agency's web server).

---

## Data Storage
All data is stored in **browser localStorage** — no backend or internet connection required.
Data persists across page refreshes but is device/browser-specific.

To back up data, use the **Export CSV** button in the app.

---

## Program Indicators (Default)

| Program | Default Indicators |
|---|---|
| ILCDB | Barangays trained, Officials reached, Modules completed, Assessments, Helpdesks |
| DTC | Walk-in clients, Services rendered, Training sessions, Partner agencies, Equipment |
| SPARK | Scholars enrolled, Courses completed, NC certifications, Employed graduates, Partners |
| Cybersecurity | Seminars, Participants, LGUs covered, Incidents responded, Trainors capacitated |
| Free Wi-Fi | Sites deployed, Active connections, Barangays covered, Uptime rate, Tickets resolved |

Custom indicators can be added per program via the **+ Add Indicator** button.

---

## Status Legend
| Label | Condition |
|---|---|
| ✅ Met | ≥ 90% accomplishment |
| 🟡 On track | 65% – 89% |
| 🔴 At risk | Below 65% |

---

*Developed for DICT Abra Provincial Office · FY 2025*
