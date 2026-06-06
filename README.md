# Ask Your Business - Company Setup & Serbian Accounting Profile Demo v1.1

This is a zero-dependency Railway demo for the owner/accountant workflow with a new Company Setup foundation for Serbian accounting mapping.

## What v1.1 adds

- Company Setup tab
- Serbian legal form / entity type selector
- Serbian accounting framework selector
- Serbian chart-of-accounts template field
- PDV / SEF / fiscalization / payroll / bank-cash setup
- Setup completeness score and certification gate
- Serbian konto dropdown inside the Review Center popup
- Posting treatment and tax/PDV/SEF treatment fields
- Blocking of accountant certification until the company setup profile is complete
- Company setup JSON export

## Local run

```bash
npm start
```

Open http://localhost:3000.

## Railway deploy

Upload the contents of this folder to the existing GitHub repo connected to Railway. Keep `Dockerfile`, `server.js`, `package.json`, and `public/` at the repository root.

Test deployment with `/health`. Expected version: `1.1.0`.

## Important limitation

Review decisions and the company setup profile are still saved in browser localStorage only. Export the company setup / certified packages before clearing browser data. The next major technical step is PostgreSQL persistence.
