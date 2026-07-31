# PR / PO Pipeline — Daily 10 AM Email (setup)

Sends the **PR / PO Pipeline — Open & Pending** report every day at **10:00 AM Dubai**,
fully headless from the existing **pr-po-dashboard-proxy** Azure Function App (no laptop,
no Power Automate). Numbers reconcile 1:1 with the dashboard (same live-pipeline logic).

New function: `pr-po-proxy/src/functions/prpoEmail.js`
- Timer `0 0 6 * * *` (06:00 UTC = 10:00 AM Dubai).
- Fetches `pr.xlsx` + `po.xlsx` from the published dashboard (GitHub Pages).
- Reuses the telemetry email's Graph `Mail.Send` + app registration — **no new admin consent needed.**

Verified locally against today's export: PR buckets = **555** (115/403/9/22/6/0),
PO buckets = **901** (15/14/0/1/637/234), department total **1,456**, 7-day inflow **96** — all matching the dashboard.

---

## Part 1 — Recipients (you, 1 min)

Function App **pr-po-dashboard-proxy** → **Settings → Environment variables** → add:

| Name | Value |
|---|---|
| `PRPO_MAIL_TO` | semicolon-separated recipients, e.g. `you@striveservicesgroup.com; boss@…` |

If you skip this, it falls back to the telemetry list (`MAIL_TO`). **Save** (app restarts).
Everything else (`TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`, `MAIL_FROM`) is already set from the telemetry email.

## Part 2 — Deploy (you)

GitHub Desktop → **pr-po-proxy** repo → commit `src/functions/prpoEmail.js` → **Push**
(auto-redeploys via `main_pr-po-dashboard-proxy.yml`). The `xlsx` dependency is already in `package.json`.

## Part 3 — Test (in this order)

Function App → Functions → **prpo-email** → **Get function URL**, then in a browser:

1. `…/api/prpo-email?code=…&debug=1` → JSON. Expect `prSum: 555`, `poSum: 901` — compare to the dashboard header.
2. `…/api/prpo-email?code=…&format=html` → the full email in the browser.
3. `…/api/prpo-email?code=…&send=1` → sends it to `PRPO_MAIL_TO` now.
4. Done — the timer repeats the send every day at **10:00 AM Dubai**. Watch the function's **Monitor** tab.

---

## Settings reference

| Name | Purpose | Status |
|---|---|---|
| `PRPO_MAIL_TO` | recipient list | **add this** |
| `PRPO_MAIL_FROM` | sender mailbox | optional (defaults to `MAIL_FROM`) |
| `PRPO_PR_URL` / `PRPO_PO_URL` | data sources | optional (default = GitHub Pages `pr.xlsx` / `po.xlsx`) |
| `TENANT_ID` / `CLIENT_ID` / `CLIENT_SECRET` / `MAIL_FROM` | auth + sender | ✅ already set (telemetry) |

## Troubleshooting

| Symptom | Fix |
|---|---|
| `debug=1` errors on `fetch 404` | `pr.xlsx`/`po.xlsx` not published yet — run `refresh_data.py`, or set `PRPO_PR_URL`/`PRPO_PO_URL`. |
| Counts differ from dashboard | The dashboard build changed its step→bucket map — re-sync `PR_MAP`/`PO_MAP` in `prpoEmail.js` from `index.html`. |
| `sendMail 403` | Graph `Mail.Send` consent (same as telemetry) — should already be granted. |
| Email but no data rows | Data fetch fell back/empty — check the `debug=1` totals first. |

Design/preview of the email: `PR_PO_Pipeline_Email_v2_last7d-by-dept.html` (same layout the function produces).
