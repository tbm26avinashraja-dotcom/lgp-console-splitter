# LGP Console Splitter

A **standalone** React/Vite finance utility. Upload one consolidated XLSX, group by the existing `Party Name` field, and download one ZIP containing an XLSX per party. Does not recalculate Aggregator rules or change CH&W.

## Deploy

1. Create a new GitHub repository and upload this project's files.
2. In Vercel, choose **Add New → Project**, import the repository, and deploy (Vite preset, `npm run build`, output `dist`).
3. Open the deployment in current Chrome or Edge over HTTPS (OPFS required).

## Development

`npm install` then `npm run dev`.

## Important limitations / validation

- This initial version targets the supplied 100 MB sample's **inline-string XLSX** format and first worksheet. It rejects large shared-string workbooks rather than producing corrupt output.
- Original cell formatting is not preserved. Header names, order, numeric cell types, and stored cell values are retained. Formula elements are stripped; their cached `<v>` values are used where present. The original workbook must contain cached values.
- Empty Party Name rows are retained in `LGP_Console_UNASSIGNED.xlsx` for Finance review.
- Case/whitespace-normalized party names are grouped; original `Party Name` cell contents are not rewritten.
- Processing requires local disk space for intermediate XML, per-party XLSX and final ZIP. OPFS may be constrained by browser quota. For production use, test browser storage quotas, performance and output reconciliation on Finance hardware.
- Verify that all output transaction row counts sum to the input count, every party appears in one file, and numeric `Final Amount` values remain numeric. A successful build alone is not a production acceptance test.
- No files are uploaded to a server. Browser local storage is used for temporary processing; clearing site data removes generated artifacts.

## Large-file architecture

The worker streams worksheet XML from ZIP and partitions rows into per-party OPFS files, then creates each XLSX and adds it to an outer ZIP sequentially. It avoids loading the complete ~1 GB decompressed XML into JS heap. ZIP processing uses `@zip.js/zip.js`.
