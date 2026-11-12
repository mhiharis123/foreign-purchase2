# Foreign Purchase Data Processor

Web tool that filters outstanding foreign purchase Excel files and downloads a cleaned `.xlsx` with a fixed column layout.

## Requirements

- Node.js (for the local server), **or** open `index.html` directly in a browser
- Input files must be `.xlsx`

## Quick start

```bash
npm install
npm start
```

Open [http://localhost:3000](http://localhost:3000), upload one or more outstanding-purchases `.xlsx` files, then click **Process Files**.

Alternatively:

```bash
python server.py
```

Opens [http://127.0.0.1:5000](http://127.0.0.1:5000).

## How filtering works

Each row is kept only if **both** are true:

1. **Contract No.** (`Cnt No`) starts with `CNPG`
2. **DR code** is in the allowlist in `script.js` (e.g. `CC`, `CC1`, `CE12`, `DMJ`, `DMS`, …)

Matching rows are sorted by Contract No. A dealer-code count summary is shown after processing.

## Column mapping

| Source column | Output column |
|---|---|
| DR code | DR Code |
| Cnt No | Contract No. |
| Name | Client Name |
| Code | Client Code |
| Counter | Counter |
| Qty | Quantity |
| Buy Price | Price |
| CCY | Traded Currency |
| CCY (Sett) | Settlement Currency |
| Amount (traded) | Contract Value |
| Action | Action |

Output column order matches the table above.

## Output

- Single file: `processed_data.xlsx`
- Multiple files: `processed_data_1.xlsx`, `processed_data_2.xlsx`, …

Sheet name: `Filtered Data`.

## Project structure

| File | Role |
|---|---|
| `index.html` | UI |
| `script.js` | Upload, filter, map, export |
| `styles.css` | Styling |
| `server.js` | Express static server (port 3000) |
| `server.py` | Flask static server (port 5000) |

## Updating the dealer allowlist

Edit `allowedDealerCodes` in `script.js`.
