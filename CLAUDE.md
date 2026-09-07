# PR/PO proxy source of truth

## Deployment safety

- The only authorised company deployment target is `ssg-prpo-proxy`.
- The legacy `pr-po-dashboard-proxy` workflow is disabled and cannot deploy.
- The `ssg-prpo-proxy` workflow is manual-only and requires an exact tested SHA.
- Its only secret is `AZURE_FUNCTIONAPP_PUBLISH_PROFILE_SSG_PRPO_PROXY` for that authorised app.
- Verify the exact Azure target and tested artifact hash before any future deployment.

## Current data boundary

- F&O virtual entities are read-only sources for PR/PO headers, lines and packing-slip journals.
- Development Dataverse `ssg_` PR/PO tables are read sources for approval capture.
- The workbook path remains in production because correction 01 still fails PR stage, PO stage and PR amount gates.
- Do not infer a workflow step from an approver identity or choose a majority label for a shared element.
- `UNRESOLVED-*` work items are not documents and never enter the distinct-document headline.

## Protected behaviour

- Keep email senders, recipients and quiet-mode settings unchanged unless explicitly approved.
- Keep the existing workbook email and weekly-snapshot paths until a full cutover passes.
- Never touch Chandan's app, flow, OneDrive or tokens.
- Never write to `operations-ifahr-live`.

## Retirement gate

`PR in review` now merges into `Sourcing`; live amounts are excluding VAT; PO stages follow live events. No business decision is pending. A fresh reconciliation must pass PR and PO stage, clock, amount, approver and count gates before a live-only model is deployed or workbook dependencies are removed.
