# Heat Register

Furnace production ledger with batch costing per heat. The app is one static page (`index.html`); your data lives in a Google Sheet, written through a small Google Apps Script backend.

```
index.html              the app (host it on GitHub Pages)
apps-script/Code.gs     backend that reads/writes your Google Sheet
apps-script/appsscript.json   web-app settings (only needed if you deploy with clasp)
```

## One-time setup (about 10 minutes)

**1. Create the Sheet and backend**
1. Create a new Google Sheet (e.g. "Heat Register Data").
2. Extensions → Apps Script. Delete the default code and paste in `apps-script/Code.gs`. Save.
3. Select the `setup` function and click Run. Approve the permissions when asked.
4. Go back to the Sheet. The tabs `Heats`, `HeatItems`, `StockReceipts`, `Assets`, `Meta` and `Setup` now exist. Copy the **access token** from cell B2 of the `Setup` tab.

**2. Publish the backend**
1. In Apps Script: Deploy → New deployment → type **Web app**.
2. Execute as: **Me**. Who has access: **Anyone**. Deploy.
3. Copy the Web app URL (ends in `/exec`).

**3. Host the app**
1. Push this folder to your repository.
2. Repository → Settings → Pages → deploy from branch `main`, folder `/ (root)`.
3. Open the Pages URL.

**4. Connect**
Open the app → *Sheet & Sync* → paste the Web app URL and the access token → Save & connect.

Do the same on each device you use. Data loads from the Sheet automatically.

## How it behaves
- Every edit is saved on the device first, then written to the Sheet about a second later. Offline edits are kept and sent when the connection returns.
- A device that has never loaded the Sheet merges its entries into it rather than overwriting.
- Edit data in the app, not in the Sheet: each save rewrites the data tabs. Use the Sheet for filters, pivots and charts. File → Version history in Google Sheets is your rollback.
- Built for one user. Two people editing at the same moment can overwrite each other's last change.

## Security
The access token is the only lock on the backend, and the Web app URL is public by design. Do not commit the token or URL to the repository. To rotate the token, delete the `TOKEN` script property (Project Settings) and run `setup` again.
