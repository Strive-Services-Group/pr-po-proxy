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
- No workbook is a production source after correction 04.
- PO stage and clock are separate: state comes from F&O; time comes from a live event, an explicitly flagged final-workbook seed, or our first observation.
- Do not infer a workflow step from an approver identity or choose a majority label for a shared element.
- `UNRESOLVED-*` work items are not documents and never enter the distinct-document headline.

## Protected behaviour

- Keep email senders, recipients and quiet-mode settings unchanged unless explicitly approved.
- Keep the morning and parallel email chains running; never send during a dry run.
- Never touch Chandan's app, flow, OneDrive or tokens.
- Never write to `operations-ifahr-live`.

## Retirement gate

The accepted correction-04 gate is P1a/P2 at 100%, with P1b reported separately. Never turn a seeded or missing clock into a live event date, and never overwrite a first-observed stage record with a later observation.
