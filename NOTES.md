# 10 September 2026 production release investigation

## What I found

- Canonical repository: `C:\Users\w.amjad\Documents\GitHub\pr-po-proxy`, clean `main` at `737793e9f6ef2ce2166ad898ed784602cfae5c57` before this change.
- Azure app: `ssg-prpo-proxy` in resource group `ssgprpoproxy`; Linux Flex Consumption, Node 20, `Recreate` update strategy.
- Flex run-from-package source: private container `app-package-ssg-prpo-proxy-04aac11` in storage account `ssgprpoproxy`, authenticated by the existing user-assigned identity. Flex correctly has no `WEBSITE_RUN_FROM_PACKAGE` setting.
- Azure deployment `762873a8-6690-4003-ac39-5a67ee374703` was accepted from GitHub at `2026-09-09T10:24:03Z` and marked successful/active at `2026-09-09T10:25:05Z`. Deployments `417005a3-0d58-42ce-ab5f-14be5ca47031` and `067b29c4-82cf-421b-93bc-2ca00d468082` were made by Core Tools, explaining why Azure recorded successful deployments when the GitHub workflows failed.
- The 7 and 8 September GitHub deployments failed before Azure login because their required publish-profile value was absent. The 9 September OIDC workflow checked out and tested exact commit `737793e`, used OneDeploy, and completed successfully.
- Before this change, `/api/version` returned HTTP 404.
- Read-only live previews on 10 September returned the post-8 September team title `Suppliers, Open Orders & Unowned PRs`, and Zaheer preview JSON resolved `Zaheer.Ahmed@domus-housing.com` with 14 items. This proves current HTTP workers were already running the new source.
- Application Insights recorded `prpo-email-daily` starting at `2026-09-10T06:00:00Z` with the new weekday cron. It failed the two division and eleven personal send attempts at `06:01:40Z` because `MAIL_FROM / PRPO_MAIL_FROM` is not present. No protected app setting was changed.
- Outlook transport headers for the old team email show internal MAPI submission from Chandan Kumar's Exchange mailbox at `2026-09-10T06:00:04Z`. The Azure function's failed attempts occurred about 96 seconds later. Therefore the twelve old-format emails did not come from `ssg-prpo-proxy`; they came from Chandan's separately protected sender/flow.
- No second PR/PO Function App, Logic App, container app, or local Windows PR/PO scheduled task exists in the one accessible Azure subscription or this machine. Per repository instructions, Chandan's app, flow, OneDrive, and tokens were inspected only through the received message headers and were not modified.

## Problems and risks

- Azure's green deployment status did not previously prove the application output or identify a Git commit.
- Two separate send mechanisms were treated as one. Redeploying `ssg-prpo-proxy` cannot update Chandan's separate sender.
- The GitHub workbook workflow is active, but scheduled run history contains no run on 9 or 10 September. Its single unattended run on 8 September began at `08:41:05Z`, outside both the former `*/15 4-5` UTC expression and the current `7,19,31,43 4-5` UTC expression. GitHub's scheduler is not a dependable pre-06:00 trigger for this job.
- Direct deployment-package blob inspection was denied by storage data-plane RBAC, and key authentication is disabled. No permission or storage setting was changed. The new public build endpoint removes the need to grant package-container access for routine live-build verification.

## Exact changes made

- Added anonymous `GET /api/version`. It returns only `{ "commit": "<full SHA>" }`, disables caching, and fails closed with a generic HTTP 503 if build identity is missing or invalid.
- The authorised deploy workflow now creates `build-info.json` from the exact tested `source_sha` immediately before packaging.
- The workflow now polls the public endpoint after OneDeploy and fails unless the running response equals that full SHA.
- `build-info.json` is ignored locally so a developer cannot accidentally commit stale build identity.
- Added deterministic tests for valid, missing, short, and malformed build identity and for the non-leaking failure response.

## What I did not change

- No email wording, address, recipient, manager copy, BCC, `PRPO_*_MAIL_TO`, `PRPO_PERSONAL_TEST`, token, Graph permission, secret, app setting, mail sender, or 06:00 UTC timer changed.
- No `/api/dataset` logic, Dataverse data, VAT logic, dashboard code, or Chandan-controlled app/flow/OneDrive/token changed.
- No email was sent. All email route requests omitted `send=1`.

## Commands and evidence

- `git fetch --prune origin`, `git status --short --branch`, `git rev-parse HEAD origin/main`, and `git log` established clean current branches and exact remote heads.
- `az functionapp list`, `az resource show ... --api-version 2024-04-01`, `az functionapp config appsettings list` with name-only/allow-listed output, and `az webapp log deployment list/show` established the app plan, deployment storage contract, protected setting names, and deployment actors/times without printing a secret.
- `gh run list/view` established the exact 7-9 September workflow results and OneDeploy output.
- `az monitor app-insights query` established the 06:00 invocation, current registered routes, and all thirteen failed mail attempts.
- Outlook web message details established MAPI submission, sender mailbox, and exact transport time. No Outlook data was changed.
- `node --check src/functions/version.js`, `npm test`, `npm run build --if-present`, and `git diff --check` passed before commit. Test result: 24/24.
- First deployment run `34446690877` stopped before Azure login because its fail-closed unit test assumed the deploy-time `build-info.json` would be absent. The test now injects an explicit nonexistent path, so it covers the same failure response whether or not the deploy workflow has baked build metadata.
- Corrected deployment run `34446874692` passed all 24 tests, Azure OIDC login, OneDeploy, and the public build-identity check. The deployed response was HTTP 200, `Cache-Control: no-store`, and `{ "commit": "d9860329a7917b84971d34954888b8793a227497" }`, exactly matching the deployed source SHA.
- The running `format=html` preview was fetched without `send=1`. After the dashboard workbook refresh, Adnan's visible header was `183 items pending your action · 183 still being priced · not yet priced.` and the next line was `Oldest item was with you since 2025-05-26 (471 days).` No money total was presented as queue coverage.
- The running team preview title was `PR / PO Pipeline — Suppliers, Open Orders & Unowned PRs`. Zaheer's preview resolved `Zaheer.Ahmed@domus-housing.com`, 14 items, and dataset revision `897143ea9ac3965a190ea7e11960dae46c597cc315c22e0a0a04f76746959fc7`.
- A same-input regression ran the current sender code against the exact pre-refresh morning workbooks from dashboard commit `15fb646b85c3e854d1e4ce99fb8ae5caa4093174`. All expected personal counts matched: dinesh 425, Adnan 189, shijil 94, roderick 88, Gokul 74, Aparna 56, Shakir 22, pramod 3, arman 2, Abdul 1. This separates code behavior from later live-data movement.
- The HTML preview files were saved outside both repositories. The in-app browser rejected their local `file:` URL under its navigation policy, so evidence is the running service's HTML response plus extracted rendered-visible text, not a browser screenshot. No alternate-browser or debugging-protocol bypass was attempted.
- Assumption correction: the brief's proposed cause (successful Azure deployment serving old code) was disproved by the live route output, App Insights timing, and Outlook transport headers. The safe release work therefore adds deterministic live-build proof while leaving the separately owned sender unchanged.
- During the final notes-only publication, manual run `34447833177` was accidentally dispatched with an incorrectly expanded full SHA. It was cancelled before it could become a valid deployment. Run `34447870083` was then dispatched with the actual commit `da1499c945a87439e3a4bb354a874ca367abee8b`; the exact-SHA checkout guard is the control that makes this mistake fail closed.

## Remaining risk and recommended next step

- Chandan's separate production sender will continue using its own old template until its owner updates or retires it. This repository explicitly forbids changing that system.
- Keep the deployment workflow as the only authorised `ssg-prpo-proxy` publisher. Treat `/api/version`, not an Azure green tick, as release proof.

# 10 September 2026 F&O owner-of-record and Waqas-only sender guard

## Source trace and assumptions

- Canonical proxy checkout was clean `main` at `b8119a71bc32dd97e121ac733769e23decb140b7`, equal to `origin/main`, before this change.
- The production `/api/dataset` route is unchanged. One HTTP 200 live snapshot was saved outside both repositories at `C:\Windows\Temp\prpo-fno-owner-20260910-0705.json`; every reconciliation and preview in this change uses that file only.
- Snapshot revision `a80bb2da0eb478efa17f19fd9c3d2a343af476a3ab0e182f9206a9d000109bce`, generated `2026-09-10T07:23:53.999Z`, contains 4,445 PR rows and 3,216 PO rows.
- "F&O owner" means the exact nonblank token in `Pending Approver/User`. Comma-joined values are split, trimmed and case-insensitively de-duplicated within each document. No department reassignment, preparer, accepted/assigned field, email-address map or name alias supplies ownership.
- Purely numeric or blank values do not name a person and remain in the explicit no-named-owner path. The saved revision contains 109 actionable PR documents with no named owner and no numeric pending-owner token.
- Legacy `pr.xlsx`, `po.xlsx`, their generator and Chandan's sender are a separate protected delivery path. They were not changed or regenerated.

## Exact implementation

- `loadItems()` now reads the live `/api/dataset` response directly and refuses a non-LIVE or malformed response. It no longer reads GitHub Pages workbooks or their state file.
- All Draft, In review and Approved requisitions are included. Each named F&O pending owner receives one attribution; a shared F&O field produces one attribution per de-duplicated name.
- Priced items keep their existing class, ageing, price and queue wording, but ownership is no longer reassigned to a department operations person. Draft and Approved items no longer substitute Preparer or Accepted By/Assign To.
- Inactive or unmapped-address names remain F&O owners because this Function App no longer delivers to their address. Blank F&O ownership alone enters the no-named-owner team list.
- Personal and division Graph message builders hard-code `w.amjad@striveservicesgroup.com` as the only To address, never create Cc/Bcc, and prefix each subject with `[FOR <F&O person or team>]`. No environment recipient setting or `PRPO_PERSONAL_TEST` branch can bypass the guard.
- The send-from lookup and Graph authentication path were not changed. No send request was made.

## Live F&O reconciliation

| person | F&O count | our PR count | difference |
|---|---:|---:|---:|
| Adnan.Ullah | 428 | 428 | 0 |
| roderick.red | 307 | 307 | 0 |
| Layusha.cleatus | 140 | 140 | 0 |
| Aparna.Pauly | 133 | 133 | 0 |
| arman.b | 10 | 10 | 0 |
| Judhin.prabhakar | 3 | 3 | 0 |
| Mahmud.hasan | 2 | 2 | 0 |
| Mohammad.w | 2 | 2 | 0 |
| Dan.roberts | 1 | 1 | 0 |
| Ernie.Lavalle | 1 | 1 | 0 |
| Firas.altamimi | 1 | 1 | 0 |
| Muhammad.faisal | 1 | 1 | 0 |
| Nathan.Buys | 1 | 1 | 0 |
| ruben.senesan | 1 | 1 | 0 |

The comparison covers 954 actionable PR documents. Every named-owner difference is zero, and our output has no person absent from F&O.

## Commands, tests and previews

- `git status --short --branch`, `git rev-parse HEAD` and `git ls-remote origin refs/heads/main` proved the starting branch and remote equality.
- One cache-busted `Invoke-WebRequest` fetched `/api/dataset`; no second live dataset fetch was used for verification.
- `node --check src/functions/prpoEmail.js`, `npm test` and `git diff --check` passed. Final proxy result before commit: 26/26 tests.
- The first test run had one fixture failure because an analytical division renderer was given an empty list. The recipient test was corrected to exercise all four pure team-message builders directly; no production logic was weakened.
- `node tests/reconcile_fno_owner_counts.js <saved revision> <proxy repo>` proved F&O, Function sender and dashboard counts from independent paths.
- No-send HTML previews were rendered outside the repositories: `C:\Windows\Temp\prpo-adnan-fno-waqas-only.html` (990,987 bytes) and `C:\Windows\Temp\prpo-procurement-fno-waqas-only.html` (93,222 bytes).
- Adnan preview subject: `[FOR Adnan.Ullah] Action needed — 428 PR/PO items pending with you (10 Sept 2026)`. The team subject begins `[FOR procurement team]`. Both message objects contain only Waqas in To and have empty Cc/Bcc.

## Protected state

- No app setting, secret, token, Graph permission, `PRPO_PERSONAL_TEST`, send-from setting or 06:00 UTC timer changed or was printed.
- A read-only Azure check confirmed the authorised target name `ssg-prpo-proxy` and listed setting names only: 17 names, with `PRPO_PERSONAL_TEST`, `MAIL_FROM` and `PRPO_MAIL_FROM` absent. No value was requested or displayed.
- No Dataverse write and no `/api/dataset` code change occurred.
- Chandan's sender, flow, OneDrive, tokens and recipients were not touched.
- No email was sent.

# 11 September 2026 PR/PO sender activation and reconciliation

- Task authority: turn on the PR/PO Function App sender for Waqas-only production proving, add one daily reconciliation email, add an off-by-default direct-recipient switch, change only the send-from mailbox app setting, deploy, and send the real Waqas-only run.
- Source inspected: `src/functions/prpoEmail.js`, `test/emailPopulation.test.js`, `user-email-addresses.json`, `PRPO-EMAIL-SETUP.md`, Azure app setting names for `ssg-prpo-proxy`, and the dashboard `CLAUDE.md`.
- App setting changed: added `PRPO_MAIL_FROM=w.amjad@striveservicesgroup.com` on `ssg-prpo-proxy`. The command output listed setting names only. No secret, token, tenant id, client id, client secret, Graph permission, recipient list, or Chandan setting was read or changed.
- Delivery switch added in code as `PRPO_PERSONAL_DELIVERY_MODE`. It is set to `waqas_only`. To flip the sender later, change it to `direct_to_owner` and redeploy after Waqas approves.
- Switch-off rule: every personal digest, every active team list, and the new reconciliation email build one `toRecipients` entry for `w.amjad@striveservicesgroup.com`; no `ccRecipients` or `bccRecipients` fields are created.
- Switch-on test rule: personal digests resolve the F&O owner through `user-email-addresses.json`, remove the `[FOR <person>]` prefix, and still create no Cc or Bcc. If a named owner has no address, that person is listed in the reconciliation email and their items are kept in the team list instead of silently dropping.
- Reconciliation email added: after the daily digests it sends one Waqas-only message with columns `person`, `F&O export says`, `our email said`, and `difference`; it also names the export files/date, total matched/different people, missing-address people, and the verdict sentence.
- The reconciliation counts are built from the IT-supplied F&O export values already carried by `/api/dataset`, not from yesterday's workbook or from the old rebuilt holder layer.
- Tests run in proxy: `npm test` at 2026-09-11T06:34Z. Result: 33 passed, 0 failed. New coverage includes switch off, switch on, no-address handling, and reconciliation arithmetic.
- Tests run in dashboard after fast-forwarding main: `node --test tests/*.test.js`. Result: 24 passed, 0 failed.
- Python dashboard tests run: `python -m unittest discover -s tests -p "test_*.py"`. Result: 23 passed, 0 failed.
- Build check note: `node -c index.html` is not a valid production-build check for this HTML dashboard and failed with Node's unknown `.html` extension error. It is not counted as a dashboard failure; the repository has no `package.json` build command.
- Local no-send reconciliation render attempted with `node -e` against `email.loadItems()`. It was stopped after the live dataset fetch took too long locally. A direct read of the production `/api/dataset` returned HTTP 200, so final preview/send verification will use the deployed Function endpoint.
- Chandan Kumar's sender, flow, OneDrive, tokens, template, recipients, and schedule were not opened or changed.

## Final production proof for sender activation

- First deploy run for code commit `76deaeb4c3a592a24e638b0d99bdc2366c9e596f`: GitHub Actions run `34571327086` passed tests, Azure deployment, and `/api/version` proof.
- Final code deploy run for commit `4a4baa54c30fabcf3b600fa650aadddc3a3e5d07`: GitHub Actions run `34571529162` passed tests, Azure deployment, and `/api/version` proof.
- Independent `/api/version` read returned `4a4baa54c30fabcf3b600fa650aadddc3a3e5d07`.
- Live no-send reconciliation preview through `/api/prpo-email?reconcile=1` returned 21 matched people, 0 differences, and these no-address names: `admin.hk`, `D365CRMADMIN`, `Layusha.cleatus`, `Nathan.Buys`, `Patrick.Smith`, `Qasim.Jahangir`, `ruben.senesan`, `Shaik.baba`.
- Real production send through `/api/prpo-email?send=1` completed at about `2026-09-11T06:53:22Z`. It sent 24 messages total: 2 active team lists, 21 personal digests, and 1 reconciliation email. Every returned `to` value was `w.amjad@striveservicesgroup.com`; no send result contained a Cc or Bcc field.
- Sent subjects:
  - `[FOR procurement team] PR / PO Pipeline — Suppliers, Open Orders & Unowned PRs (11 Sept 2026)`
  - `[FOR invoicing team] PR / PO Pipeline — Pending Invoicing (11 Sept 2026)`
  - `[FOR Adnan.Ullah] Action needed — 246 PR/PO items pending with you (11 Sept 2026)`
  - `[FOR roderick.red] Action needed — 183 PR/PO items pending with you (11 Sept 2026)`
  - `[FOR Aparna.Pauly] Action needed — 96 PR/PO items pending with you (11 Sept 2026)`
  - `[FOR Layusha.cleatus] Action needed — 31 PR/PO items pending with you (11 Sept 2026)`
  - `[FOR arman.b] Action needed — 31 PR/PO items pending with you (11 Sept 2026)`
  - `[FOR Riyaz.n] Action needed — 29 PR/PO items pending with you (11 Sept 2026)`
  - `[FOR Mohamed.Ashraf] Action needed — 6 PR/PO items pending with you (11 Sept 2026)`
  - `[FOR dinesh.laxman] Action needed — 5 PR/PO items pending with you (11 Sept 2026)`
  - `[FOR D365CRMADMIN] Action needed — 4 PR/PO items pending with you (11 Sept 2026)`
  - `[FOR Gokul.Krishna] Action needed — 4 PR/PO items pending with you (11 Sept 2026)`
  - `[FOR Ayman.ismail] Action needed — 3 PR/PO items pending with you (11 Sept 2026)`
  - `[FOR Mohammad.w] Action needed — 2 PR/PO items pending with you (11 Sept 2026)`
  - `[FOR it.solutions] Action needed — 1 PR/PO item pending with you (11 Sept 2026)`
  - `[FOR Shaik.baba] Action needed — 1 PR/PO item pending with you (11 Sept 2026)`
  - `[FOR admin.hk] Action needed — 1 PR/PO item pending with you (11 Sept 2026)`
  - `[FOR Nathan.Buys] Action needed — 1 PR/PO item pending with you (11 Sept 2026)`
  - `[FOR Abdul.Muqeet] Action needed — 1 PR/PO item pending with you (11 Sept 2026)`
  - `[FOR Judhin.prabhakar] Action needed — 1 PR/PO item pending with you (11 Sept 2026)`
  - `[FOR Qasim.Jahangir] Action needed — 1 PR/PO item pending with you (11 Sept 2026)`
  - `[FOR ruben.senesan] Action needed — 1 PR/PO item pending with you (11 Sept 2026)`
  - `[FOR Patrick.Smith] Action needed — 1 PR/PO item pending with you (11 Sept 2026)`
  - `[FOR Waqas] PR / PO sender reconciliation (11 Sept 2026)`
- Rendered reconciliation table:
  - `Adnan.Ullah | 246 | 246 | 0`
  - `roderick.red | 183 | 183 | 0`
  - `Aparna.Pauly | 96 | 96 | 0`
  - `arman.b | 31 | 31 | 0`
  - `Layusha.cleatus | 31 | 31 | 0`
  - `Riyaz.n | 29 | 29 | 0`
  - `Mohamed.Ashraf | 6 | 6 | 0`
  - `dinesh.laxman | 5 | 5 | 0`
  - `D365CRMADMIN | 4 | 4 | 0`
  - `Gokul.Krishna | 4 | 4 | 0`
  - `Ayman.ismail | 3 | 3 | 0`
  - `Mohammad.w | 2 | 2 | 0`
  - `Abdul.Muqeet | 1 | 1 | 0`
  - `admin.hk | 1 | 1 | 0`
  - `it.solutions | 1 | 1 | 0`
  - `Judhin.prabhakar | 1 | 1 | 0`
  - `Nathan.Buys | 1 | 1 | 0`
  - `Patrick.Smith | 1 | 1 | 0`
  - `Qasim.Jahangir | 1 | 1 | 0`
  - `ruben.senesan | 1 | 1 | 0`
  - `Shaik.baba | 1 | 1 | 0`
- The reconciliation verdict was `Every person matches`. The export source/date line was rendered from the current export authority carried by the live dataset.
- Chandan Kumar's sender, flow, OneDrive, tokens, template, recipients, and schedule were untouched. The only production app setting changed was `PRPO_MAIL_FROM`.

# 11 September 2026 reconciliation correction 01

- Correction accepted: the first reconciliation compared digest-built `items` against digest-built `personalPool` counts, so it could not prove the email counts independently.
- Code changed so the reconciliation export column is now built by `countOwnersFromExportWorkbooks()` in `src/functions/prpoEmail.js`. That function counts directly from workbook rows supplied by `readExportReconciliationSource()`, which opens the current `Purchase Reques*.xlsx` and `Purchase order*.xlsx` files through `loadExportAuthority()` in `src/shared/prpoDataset.js`.
- The reconciliation no longer uses `items`, `personalPool`, `groupByOwner`, or `personalAttributionPool` to build the `F&O export says` column. The `our email said` column still comes from the digest count that was actually built for each person.
- Added direct-cutover guard: if `PRPO_PERSONAL_DELIVERY_MODE` is later changed to `direct_to_owner`, stale export data or any non-zero reconciliation difference forces that run back to Waqas-only delivery and puts the hold reason at the top of the reconciliation email. The mode remains `waqas_only`.
- Live folder checked with Graph against the exact configured source `w.amjad@striveservicesgroup.com` / `Claude/PR PO Pipeline Dashboard/Email-Drops`. The folder contains only `Purchase order.xlsx` modified `2026-09-07T11:34:19Z` and `Purchase Reques.xlsx` modified `2026-09-07T11:34:18Z`.
- Plain cause of stale export: the Function is correctly reading the newest pair available in the OneDrive folder it can access, but Abdul's newer 10 and 11 September exports are not in that folder. No mailbox, Chandan flow, OneDrive configuration, token, permission, app setting, or timer was changed to work around that.
- Tests run after the fix: `npm test`. Result: 36 passed, 0 failed. New tests prove the reconciliation can go red when a digest count is deliberately wrong, direct delivery is held on stale export, and direct delivery is held on a non-zero reconciliation difference.
- Protected items: Chandan Kumar's sender, flow, OneDrive, tokens, template, recipients, and schedule were untouched. No secret, token, permission, timer, recipient setting, or app setting was changed.
- Deployed commit `6a4dfe02e2c2eaa4c0ad6c9b0cd36ec50ab2426c` with GitHub Actions run `34573316187`; the run passed its package deploy step and public `/api/version` proof.
- Live `/api/version` then returned `6a4dfe02e2c2eaa4c0ad6c9b0cd36ec50ab2426c`, matching both local `HEAD` and `origin/main`.
- Live no-send reconciliation from `/api/prpo-email?reconcile=1` used `Purchase Reques.xlsx` and `Purchase order.xlsx`, export date `2026-09-07T11:34:19Z`. Result: 21 matched people, 0 people with a difference, 8 export-named people without an address on file. The rendered stale warning said the latest export supplied by IT is dated 7 September 2026 and is older than this morning's send.
- Real production run from `/api/prpo-email?send=1` was executed after deployment. Result: `sentAll=true`, `deliveryMode=waqas_only`, `waqasOnly=true`, stale export warning present, 2 team messages sent to `w.amjad@striveservicesgroup.com`, 21 personal messages sent to `w.amjad@striveservicesgroup.com`, and 1 reconciliation message sent to `w.amjad@striveservicesgroup.com`. The reconciliation result in that run was 21 matched and 0 differences.
