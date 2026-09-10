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

## Remaining risk and recommended next step

- Chandan's separate production sender will continue using its own old template until its owner updates or retires it. This repository explicitly forbids changing that system.
- Keep the deployment workflow as the only authorised `ssg-prpo-proxy` publisher. Treat `/api/version`, not an Azure green tick, as release proof.
