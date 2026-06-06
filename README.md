# Ask Your Business - Owner Command Center v1.6

This demo keeps the accountant-side data intake/review/certification workflow, but changes the owner experience into a finished product view.

## v1.6 focus

- Owner sees a clean command center only.
- Owner does not perform setup, review, mapping, or certification.
- Accountant side handles company setup, Serbian account mapping, review, certification, and report preparation.
- Owner gets dedicated pages for:
  - Owner Home
  - Cash & Obligations
  - Financial Reports
  - Tax Reports
  - Company Health
  - AI Models & Analytics
  - AI Assistant
  - Documents & Published Reports

## Railway

The app has no external dependencies and runs with:

```bash
npm start
```

Health check:

```text
/health
```

Expected version: `1.6.0`.

## Important demo limitation

This is still browser-local. Review decisions, company setup, mappings, and certification state are saved in localStorage only until PostgreSQL is added.
