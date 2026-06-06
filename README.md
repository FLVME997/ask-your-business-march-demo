# Ask Your Business - Simplified Owner + Accountant Operations v0.7

This Railway demo upgrades the current one-company workflow before adding PostgreSQL.

## What changed in v0.7

- Simplified Owner Portal: the owner sees only business-use status, accounting status, cash summary, and questions requiring owner input.
- Accountant Operations Center: accountant team sees exception queue, month operations board, template reuse, owner questions, and priority items.
- Better management-ready vs accounting-ready split.
- Bulk handling in the Review Center after filtering a safe group of items.
- Exception grouping by broad review area instead of many technical issue types.
- Still uses browser localStorage for decisions and exports JSON packages.

## What is still intentionally not included

- PostgreSQL persistence
- Real login/permissions
- Real Excel upload/parsing inside Railway
- OpenAI/API assistant
- Bank feeds
- OCR/document extraction

## Deploy

Upload the contents of this folder to the root of the existing GitHub repository connected to Railway.

The root of the repository should contain:

```text
Dockerfile
server.js
package.json
README.md
.dockerignore
public/
```

After Railway redeploys, check:

```text
/health
```

It should show version `0.7.0`.

## Test flow

1. Open Owner Portal and confirm it is simple: cash, business-use status, accounting status, owner questions.
2. Switch to Accountant Operations.
3. Open exception groups or priority queue.
4. In Review Center, filter a safe group.
5. Test one bulk action, such as sending filtered items to accountant review.
6. Export accountant review pack.
7. Reset local progress if needed.

## Important warning

This is still Option 1: browser-only local review state. Export your package before clearing browser data or changing computers.
