# PR / PO Pipeline — Daily 10 AM Email (setup)

Builds the PR/PO personal digests and team lists every weekday at **10:00 AM Dubai**
from the authorised **ssg-prpo-proxy** Function App.

New function: `pr-po-proxy/src/functions/prpoEmail.js`
- Timer `0 0 6 * * 1-5` (06:00 UTC = 10:00 AM Dubai, weekdays).
- Reads the live `/api/dataset` response. Workbooks are not a runtime source.
- Every Graph message is hard-coded to `w.amjad@striveservicesgroup.com`, with no Cc or Bcc.
- Each subject starts with `[FOR <F&O person or team>]` so the intended audience is visible.
- Environment recipient settings and `PRPO_PERSONAL_TEST` cannot override this guard.

---

## Recipient safety

Do not add or change a recipient setting. The code-level Waqas-only guard is the production test channel.
Removing that guard requires a separate authorised change with new tests.

## Part 2 — Deploy (you)

Use the manual authorised deployment workflow for `ssg-prpo-proxy` and supply the exact tested full commit SHA.
The workflow verifies `/api/version` after deployment.

## Part 3 — Test (in this order)

Function App → Functions → **prpo-email** → **Get function URL**, then in a browser:

1. `…/api/prpo-email?code=…` → JSON summary from the live dataset revision.
2. `…/api/prpo-email?code=…&format=html` → the full email in the browser.
3. Do not use `send=1` for verification. Rendering the preview is sufficient.
4. The timer evaluates the same Waqas-only message builder on weekdays.

---

## Settings reference

| Name | Purpose | Status |
|---|---|---|
| `PRPO_MAIL_FROM` | sender mailbox | optional (defaults to `MAIL_FROM`) |
| `TENANT_ID` / `CLIENT_ID` / `CLIENT_SECRET` / `MAIL_FROM` | auth + sender | ✅ already set (telemetry) |

## Troubleshooting

| Symptom | Fix |
|---|---|
| Dataset request fails | Check `/api/dataset`; no workbook fallback is used. |
| Counts differ from dashboard | The dashboard build changed its step→bucket map — re-sync `PR_MAP`/`PO_MAP` in `prpoEmail.js` from `index.html`. |
| `sendMail 403` | Graph `Mail.Send` consent (same as telemetry) — should already be granted. |
| Email but no data rows | Data fetch fell back/empty — check the `debug=1` totals first. |

Design/preview of the email: `PR_PO_Pipeline_Email_v2_last7d-by-dept.html` (same layout the function produces).
