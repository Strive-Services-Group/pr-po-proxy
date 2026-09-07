# PR/PO proxy source of truth

## Deployment safety

- The only authorised company deployment target is `ssg-prpo-proxy`.
- The checked-in GitHub workflow targets `pr-po-dashboard-proxy`, which is out of scope.
- Do not run or modify that workflow for company deployment.
- Verify the exact Azure target and tested artifact hash before any future deployment.

## Current data boundary

- F&O virtual entities are read-only sources for PR/PO headers, lines and packing-slip journals.
- Development Dataverse `ssg_` PR/PO tables are read sources for approval capture.
- The workbook path remains in production because the live sources did not pass retirement gates on 7 September 2026.
- Do not infer a workflow step from an approver identity or choose a majority label for a shared element.
- `UNRESOLVED-*` work items are not documents and never enter the distinct-document headline.

## Protected behaviour

- Keep email senders, recipients and quiet-mode settings unchanged unless explicitly approved.
- Keep the existing workbook email and weekly-snapshot paths until a full cutover passes.
- Never touch Chandan's app, flow, OneDrive or tokens.
- Never write to `operations-ifahr-live`.

## Retirement gate

The current capture cannot deterministically separate `PR in review` from `Sourcing`. Waqas must approve a new business rule before a live-only model is implemented or deployed. A fresh reconciliation must then pass PR and PO stage, clock, approver and count gates before workbook dependencies are removed.
