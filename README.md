# Leads Calculator

Tracks leads received from `leads@soumissionsdemenageurs.ca`, detects the ones
you rejected (RETOUR replies), and rolls up totals per city each time the
renewal email arrives. Runs entirely as a Google Apps Script bound to a
Google Sheet — no server, no local computer, $0 cost (covered by your
existing Google Workspace plan).

## How it works

- **Ingest**: every 15 minutes, scans for new emails from
  `leads@soumissionsdemenageurs.ca` to `tom@tommoving.ca` and logs each one
  as a row in the **Leads Log** tab. City defaults to Gatineau.
- **RETOUR detection**: scans for threads where you replied containing the
  word "RETOUR" and marks that lead **Rejected**.
- **Renewal rollup**: watches for an email from
  `admin@soumissionsdemenageurs.ca` titled
  *"Renouvellement automatique bloc de soumissions"*. When one arrives, it
  closes out the current period (per city), counts total leads / rejected /
  net, computes cost using the **Config** tab rates, and writes a row to
  **Renewal Periods** — including a check against the known invoice total
  (574.88 CAD default).
- **Ottawa**: not distinguishable in the email yet. When you're ready, apply
  the Gmail label `LeadsCalc/City-Ottawa` to a lead thread and it will be
  logged as Ottawa instead of Gatineau. (See "Open items" below.)

## One-time setup

1. **Create the Sheet**: in Google Drive, create a new Google Sheet named
   `Leads Calculator`.
2. **Create the bound script**: in the Sheet, go to
   `Extensions > Apps Script`. This opens an empty script project already
   attached to the Sheet.
3. **Copy the Script ID**: in the Apps Script editor, go to
   `Project Settings` (gear icon) and copy the **Script ID**.
4. **Point this local project at it**: copy `.clasp.json.template` to
   `.clasp.json` and paste the Script ID in:
   ```
   cp .clasp.json.template .clasp.json
   ```
   then edit `.clasp.json` and replace `PASTE_YOUR_SCRIPT_ID_HERE`.
5. **Install clasp** (Google's official CLI, dev dependency only — nothing
   runs through it later):
   ```
   npm install
   ```
6. **Log in** (opens a browser to authorize your Google account — one time):
   ```
   npm run login
   ```
7. **Push the code**:
   ```
   npm run push
   ```
   This uploads everything in `src/` to the bound script. The empty default
   `Code.gs` that Google creates in step 2 is harmless — you can delete it
   from the Apps Script editor for tidiness, or leave it.
8. **Run setup**: open the Sheet, reload the page (so the custom menu
   appears), then click `Leads Calculator > Setup (first time...)`.
   Google will prompt you to authorize Gmail + Sheets access the first time
   — approve it (it's your own script). This creates the three tabs, the
   Gmail labels, and the 15-minute trigger.

That's it — from here it runs on its own. You can also click
`Leads Calculator > Run now` any time to trigger a check immediately instead
of waiting for the next scheduled run.

## Filling in costs

Open the **Config** tab and enter the before-tax cost per lead in column B
for each city. The with-tax cost (column D) is computed automatically using
the tax rate in column C (defaults: 14.975% for Gatineau/QC, 13% for
Ottawa/ON — adjust if wrong). Once filled in, new rollups in
**Renewal Periods** will compute Total Cost and flag whether it matches the
574.88 CAD invoice.

## Open items to confirm later

- **Ottawa city marking**: currently manual (apply the
  `LeadsCalc/City-Ottawa` Gmail label to a thread). If the vendor agrees to
  tag the city in the subject/body instead, the `ingestLeads` logic in
  `src/Ingest.gs` can be switched to detect that automatically.
- **Multiple renewal emails per cycle**: once Ottawa is live, confirm
  whether you'll get one combined renewal email or one per city bloc. The
  rollup logic currently applies to whichever cities have counted leads
  since their own last renewal, so either pattern should work, but worth
  confirming once real Ottawa emails start arriving.
