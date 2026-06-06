# Ask Your Business - Smart Category Suggestions v0.9

This Railway demo adds reusable mapped-category suggestions inside the Review Center popup.

## New in v0.9

- The **Mapped / management category** field now has a dropdown sourced from previous mappings, transactions, mapping rules, and manually added categories.
- You can still type a brand-new category manually.
- New typed categories are remembered locally after you save/certify the review item.
- The Mapping tab now shows the current mapped-category suggestion dictionary.
- Validation/control issues and manual adjustment categories use the same suggestion library.
- The v0.8 source evidence viewer remains available in the popup.

## Important demo limitation

Review decisions and new categories are saved in browser localStorage for now. Export the certified package or accountant review pack before clearing browser data. Later, this will move into PostgreSQL.

## Deploy

Upload the contents of this folder to the root of the existing GitHub repository connected to Railway.

Check `/health` after deployment. It should show version `0.9.0`.


## v0.9 mapped-category suggestions

The Review Center mapped-category fields now use dropdown/autocomplete suggestions from previous imports, seed categories, approved mapping rules, local transaction edits, manual adjustments, and manually typed categories. Reviewers can still type a new category; once they save/certify the item, the category is stored in browser local storage and appears in future dropdown suggestions and in the Mapping tab.
