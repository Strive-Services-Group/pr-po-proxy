# PR / PO Proxy (Azure Function)

Small Azure Function that reads purchase requisition (PR) and purchase order (PO) data from D365 Finance & Operations over OData and returns clean JSON for the dashboard. It exists so the public dashboard never holds D365 credentials and so the data can be assembled/cached server-side.

**Companion repo:** [PR-PO-Pipeline-Dashboard](https://github.com/Strive-Services-Group/PR-PO-Pipeline-Dashboard) — the dashboard website. Same project, kept separate because the repositories deploy to different places.

## Deploys to

The authorised company target for PR/PO work is **`ssg-prpo-proxy`**. `.github/workflows/deploy-ssg-prpo-proxy.yml` is manual-only, uses Azure OIDC, and requires the exact tested commit SHA. The legacy workflow is disabled and has no push trigger or deployment step.

## Live dataset contract

Correction 04 replaces the final workbook path. F&O is authoritative for PR/PO state and line amounts; the development `ssg_` capture is authoritative for approval assignments and PO stage observations. The final workbook is used once only to seed otherwise unavailable PO clocks, explicitly labelled `SEEDED_FROM_FINAL_WORKBOOK`.

## Endpoints
- `GET /api/version` — public full commit SHA baked into the deployed package
- `GET /api/dataset` — shared revision used by dashboard and email
- `GET /api/pr` — PR slice of that revision
- `GET /api/po` — PO slice of that revision

The authorised deployment fails unless `/api/version` reports the exact source SHA after OneDeploy. Its response contains no setting or secret:

```json
{"commit":"0123456789abcdef0123456789abcdef01234567"}
```

Legacy out-of-scope URL (reference only; do not deploy): `https://pr-po-dashboard-proxy-b4budzexh7eveved.uaenorth-01.azurewebsites.net`

## What it does
- Authenticates to D365 F&O (client-credentials) using app settings.
- Reads all companies from F&O with `cross-company=true`.
- Derives PR stages from approval capture and active-line pricing.
- Derives PO stages from F&O status, confirmations, receipts and invoices.
- Preserves the first observation of PO stage changes in development Dataverse.
- Returns one revision shared by the dashboard and email; caches for 3 minutes.

## Key files
- `src/shared/prpoDataset.js` — the source-of-truth assembly and clock model.
- `src/functions/pr.js` — live dataset endpoints and CORS.
- `package.json`, `host.json` — Functions config.
- `stepMap.json` / `poStepMap.json` — workflow element GUID → step-name maps (legacy; the dashboard now overlays steps from its own export, so these are secondary).
- `README-DEPLOY.md` — deployment + troubleshooting notes.

## App settings (in the Azure Function App, not in code)
`TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`, `FO_RESOURCE`, `ALLOWED_ORIGIN`, `DATAVERSE_API_URL`, `DATAVERSE_RESOURCE`, `DATAVERSE_MI_CLIENT_ID`, `PRPO_STAGE_OBSERVATION_WRITE` (and `DASHBOARD_CLIENT_ID` if token-auth is enabled).

## Source-of-truth note

F&O virtual entities are the source for current headers, lines and PO lifecycle state. The development `ssg_` capture is the source for current approval work items and first-observed PO stage clocks. No workbook is a runtime source.
