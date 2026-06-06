# Ask Your Business - Source Evidence Review v0.8

This Railway demo adds a source evidence viewer inside the Review Center popup.

## New in v0.8

- When you click a Needs Review line, the popup now shows a Source Evidence Viewer.
- Shows the original workbook name, month, sheet, row/cell reference, and source reference.
- Shows the normalized/imported transaction or validation/control record created by the app.
- Shows source rows near the referenced sheet/row, sorted by sheet and row.
- Shows daily balance controls and monthly status rows from the generated review documents.
- Includes buttons to open the PDF report and download the generated review workbook.

## Important demo limitation

The raw original Excel workbooks are not bundled in this public Railway demo because they can contain sensitive company/personal data. The app shows workbook locator information and extracted source evidence. In production, this same panel will open the secured original file from private storage.

## Deploy

Upload the contents of this folder to the root of the existing GitHub repository connected to Railway.

Check `/health` after deployment. It should show version `0.8.0`.
