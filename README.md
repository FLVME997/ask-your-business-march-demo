# Ask Your Business - Serbian COA Mapping Dictionary v1.2

This zero-dependency Railway demo expands the Company Setup and Review Center into a Serbian accounting mapping layer.

## What v1.2 adds

- Searchable Serbian Chart of Accounts starter dictionary.
- Classes 5 and 6 cost/income starter mapping plus common cash, bank, receivable, supplier, PDV, payroll, liability, and equity accounts.
- Serbian konto suggestions in the review popup.
- Posting treatment selector.
- Tax / PDV / SEF treatment selector.
- Reusable mapping rules that can store: raw category + mapped category + Serbian konto + posting treatment + tax treatment.
- Serbian COA tab for accountant-side review and filtering.
- Export Serbian dictionary JSON for later PostgreSQL import.

## Local run

```bash
npm start
```

Open http://localhost:3000.

## Railway deploy

Upload the contents of this folder to the existing GitHub repo connected to Railway. Keep `Dockerfile`, `server.js`, `package.json`, and `public/` at the repository root.

Test deployment with `/health`. Expected version: `1.2.0`.

## Important limitation

This is still browser-local. Company profile, review decisions, Serbian konto selections, and certifications are stored in localStorage. Export packages before clearing browser data. PostgreSQL comes later.
