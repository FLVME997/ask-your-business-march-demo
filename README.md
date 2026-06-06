# Ask Your Business - Certified Data View Demo v1.0

This is a zero-dependency Railway demo for the owner/accountant workflow.

## What v1.0 adds

- Certified Data tab
- Management-ready data view
- Accountant-certified data view
- Before vs after totals
- Remaining limitations panel
- Management-ready category breakdown
- Rejected/excluded rows view
- Manual adjustments view
- Management analysis export

## Local run

```bash
npm start
```

Open http://localhost:3000.

## Railway deploy

Upload the contents of this folder to the existing GitHub repo connected to Railway. Keep `Dockerfile`, `server.js`, `package.json`, and `public/` at the repository root.

Test deployment with `/health`. Expected version: `1.0.0`.

## Important limitation

Review decisions are still saved in browser localStorage only. Export certified packages before clearing browser data. The next major technical step is PostgreSQL persistence.
