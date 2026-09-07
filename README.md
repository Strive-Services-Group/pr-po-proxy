# PR / PO Proxy (Azure Function)

Small Azure Function that reads purchase requisition (PR) and purchase order (PO) data from D365 Finance & Operations over OData and returns clean JSON for the dashboard. It exists so the public dashboard never holds D365 credentials and so the data can be assembled/cached server-side.

**Companion repo:** [PR-PO-Pipeline-Dashboard](https://github.com/Strive-Services-Group/PR-PO-Pipeline-Dashboard) — the dashboard website. Same project, kept separate because the repositories deploy to different places.

## Deploys to

The authorised company target for PR/PO work is **`ssg-prpo-proxy`**. This repository does not currently contain a deployment workflow for that app. Its checked-in workflow targets the out-of-scope **`pr-po-dashboard-proxy`** app; do not run or edit that workflow for company deployment. Deploy `ssg-prpo-proxy` only through a separately verified, explicitly authorised path.

## Workbook retirement status

The 7 September 2026 reconciliation concluded **cannot retire**. Current live sources cannot deterministically separate dashboard `PR in review` from `Sourcing`, and the stage thresholds did not pass. The existing workbook-based dashboard, email and snapshot paths remain protected. No live-dataset cutover or deployment was made.

## Endpoints
- `GET /api/pr` — assembled purchase requisitions
- `GET /api/po` — assembled purchase orders

Legacy out-of-scope URL (reference only; do not deploy): `https://pr-po-dashboard-proxy-b4budzexh7eveved.uaenorth-01.azurewebsites.net`

## What it does
- Authenticates to D365 F&O (client-credentials) using app settings.
- Reads `PurchaseRequisitionHeaders` + lines + `WorkflowWorkItems` (PR), and `PurchaseOrderHeadersV2` + lines + `VendorsV2` (PO).
- Returns one JSON row per PR/PO with the fields the dashboard needs; caches for 3 minutes.

## Key files
- `src/functions/pr.js` — the function (token + OData queries + assembly + CORS).
- `package.json`, `host.json` — Functions config.
- `stepMap.json` / `poStepMap.json` — workflow element GUID → step-name maps (legacy; the dashboard now overlays steps from its own export, so these are secondary).
- `README-DEPLOY.md` — deployment + troubleshooting notes.

## App settings (in the Azure Function App, not in code)
`TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`, `FO_RESOURCE`, `ALLOWED_ORIGIN` (and `DASHBOARD_CLIENT_ID` if token-auth is enabled).

## Source-of-truth note

F&O virtual entities are the source for current headers and lines. The development `ssg_` capture is the source for current approval work items and assignment observations. The workbook remains the production source for workflow detail that those sources cannot reproduce. See the companion repository's `evidence/workbook-retirement-report.md` before changing that boundary.
