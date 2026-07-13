# Supply Traffic Optimizer

Static React + TypeScript web app for ad-ops supply optimization.

## What it does

- Upload an Excel report (`.xlsx` / `.xls`)
- Group traffic by:
  - Supply Name
  - Bundle (Platform)
- Detect bad traffic rows per group using configurable rules:
  - zero sRPM
  - low sRPM
  - high requests share + low cost share
  - very low fill rate
- Show before/after optimization KPIs:
  - total demand requests
  - total impressions
  - total cost
  - fill rate (`impressions / requests`)
- Display flagged rows with reasons.

## Run locally

```bash
npm install
npm run dev
```

## Build static bundle

```bash
npm run build
```

Static files are generated in `dist/` and can be deployed to any static host.
