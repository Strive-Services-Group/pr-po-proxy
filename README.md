# PR / PO Proxy (Azure Function)

Small Azure Function that reads purchase requisition (PR) and purchase order (PO) data from D365 Finance & Operations over OData and returns clean JSON for the dashboard. It exists so the public dashboard never holds D365 credentials and so the data can be assembled/cached server-side.

**Companion repo:** [PR-PO-Pipeline-Dashboard](https://github.com/Chandansah605/PR-PO-Pipeline-Dashboard) — the dashboard website that calls this proxy. Same project, kept separate because they deploy to different places.

## Deploys to
Azure Function App **`pr-po-dashboard-proxy`** (Flex Consumption, Node), via Azure **Deployment Center → GitHub Actions**. Pushing to `main` redeploys.

## Endpoints
- `GET /api/pr` — assembled purchase requisitions
- `GET /api/po` — assembled purchase orders

Base URL: `https://pr-po-dashboard-proxy-b4budzexh7eveved.uaenorth-01.azurewebsites.net`

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

## Note
The granular workflow "Step name" is a stored `IFAHR*` field in D365 not yet exposed on OData; until the F&O developer exposes it, the dashboard supplies the step from its own export overlay. See the companion repo's README.
