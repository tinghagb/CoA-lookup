# CoA Lookup Tool — User Manual

A local web tool that turns barcode scans (or an inventory spreadsheet) into
fully populated Certificate of Analysis records from **BioLegend**,
**BD Biosciences**, and **Miltenyi Biotec**.

---

## 1. What it does

For each antibody / reagent you give it (either by scanning a label or by
uploading a spreadsheet), the tool:

1. **Parses** the catalog and lot number from the scan.
2. **Fetches** the matching Certificate of Analysis (CoA) directly from the
   vendor's site.
3. **Extracts** product name, clone, fluorochrome, isotype, host species,
   reactivity, concentration, expiration date, storage conditions, applications,
   and per-test results.
4. **Displays** the result in a sortable table you can export to Excel or CSV.

Everything runs locally on your computer. No data is sent anywhere except to
the vendor websites you would normally visit yourself.

---

## 2. System requirements

| Component | Requirement |
|---|---|
| OS | macOS 12+ or Windows 10/11 |
| Python | 3.10 or newer (the Windows launcher installs it for you if missing) |
| Browser | Any modern browser (Chrome, Safari, Firefox, Edge) |
| Internet | Required — the tool fetches CoAs from vendor sites |
| Scanner | Any USB or Bluetooth 2D barcode scanner that emulates a keyboard (recommended; manual entry also works) |

---

## 3. Installation & first launch

### macOS

1. Double-click **`setup_and_run.command`**.
2. If macOS warns "cannot be opened because it is from an unidentified
   developer," right-click the file → **Open** → **Open** again. (You only
   need to do this once.)
3. A Terminal window opens, installs dependencies the first time, then
   launches the server.
4. Your browser opens automatically to **http://localhost:5050**.

### Windows

1. Double-click **`setup_and_run.bat`**.
2. The script auto-detects Python; if it's not installed, it offers to
   install Python 3.12 (via winget or the official python.org installer).
3. A Command Prompt window opens, installs dependencies, then launches the
   server.
4. Your browser opens automatically to **http://localhost:5050**.

### Manual / advanced

```bash
pip install -r requirements.txt
python app.py
```

Then open `http://localhost:5050`.

**Keep the Terminal / Command Prompt window open** — closing it shuts down
the server.

---

## 4. Using the tool

### 4.1 Single Scan tab

The fastest workflow if you have a barcode scanner:

1. **Pick a vendor** from the dropdown (BioLegend / BD / Miltenyi Biotec).
   The placeholder text shows the expected barcode format for that vendor.
2. **Click into the "Scan barcode" box** so it has keyboard focus.
3. **Scan the label** (or type/paste the barcode and press Enter).
4. The tool shows a "Detected" chip with the catalog and lot it parsed,
   then fetches the CoA and renders the result card below.

If you don't have a scanner, type the catalog and lot into the **Cat #** /
**Lot #** boxes and press **Look Up CoA**.

#### Scan history

Every successful scan is added to the **Scan History** table at the bottom
of the page. You can:

- **Select rows** with the checkboxes (or click "Select All").
- **Export to CSV** or **Export to Excel** — selected rows only.
- **Clear** the history without exporting.

### 4.2 Batch Process Spreadsheet tab

For populating an existing inventory file:

1. Pick a **default vendor** (used for rows that don't specify their own).
2. Drag-and-drop an `.xlsx` file onto the drop zone (or click to browse).
3. **Required columns**: the tool auto-detects columns named like `Cat #`,
   `Catalog Number`, `Lot #`, `Lot Number`. Optional: a `Vendor` column
   with the value `biolegend`, `bd`, or `miltenyi` per row.
4. Click **Start Processing**. A progress bar shows live status.
5. When done, click **Download Filled Spreadsheet**. New cells the tool
   populated are highlighted **light purple** so you can spot them at a
   glance; existing values are preserved.

---

## 5. Supported barcode formats

| Vendor | Example | Decoding |
|---|---|---|
| BioLegend | `344742 B402098 2026/03/05 04:16:50` | Whitespace-separated: catalog, lot, scan date/time |
| BD | `0100382905669699172707311050508742400026` | GS1 DataMatrix: `01`+GTIN, `17`+expiry, `10`+7-char lot |
| Miltenyi | `916561301237105250101832000017…` | AI 91 + 3-digit subcode + 9-digit catalog (e.g. `130-123-710`) + 10-digit lot |

Auto-detection runs across all three formats when no vendor is selected, but
selecting the correct vendor up front is faster and more reliable.

---

## 6. Field reference

The result card and export columns include the following fields. Fields the
vendor's CoA doesn't include are left blank.

| Field | Description |
|---|---|
| Catalog # | Vendor catalog / product number |
| Lot # | Manufacturing lot |
| Product Name | Full product name from the CoA |
| Marker / Target | Antigen target (e.g. CD3) |
| Fluorochrome | Conjugate (e.g. PE, APC, Brilliant Violet 421) |
| Clone | Antibody clone name |
| Isotype | e.g. IgG1, κ |
| Host Species | e.g. Mouse, Rat |
| Reactivity | Species the antibody recognizes |
| Concentration | µg/test, mg/mL, etc. |
| Optimal Dilution | Recommended dilution |
| Expiry Date | Lot expiration |
| Storage | Storage conditions |
| Formulation | Buffer composition |

---

## 7. Troubleshooting

### "CoA not found" or empty fields

Click the **"View on <vendor>"** link in the result card to verify the
catalog/lot on the vendor's own site. Some lots are only published a few
days after manufacture; very new lots may not yet have a public CoA.

### Lookup works on macOS but not Windows (or vice-versa)

Check the Command Prompt / Terminal window — it logs every HTTP request.
A typical fix:

- Ensure Python 3.10+ is installed (`python --version`).
- Reinstall dependencies: `pip install --upgrade -r requirements.txt`.
- Make sure your firewall allows the Python interpreter to reach the
  vendor domains (`biolegend.com`, `regdocs.bd.com`,
  `assets.miltenyibiotec.com`).

### Port 5050 already in use

Another instance of the tool (or another app) is using that port. Close it
and re-launch, or edit `app.py` and change `PORT = 5050` to a free port.

### Browser didn't open automatically

Open it yourself and go to **http://localhost:5050**.

### The barcode parses but no CoA comes back

The vendor sometimes throttles or temporarily rate-limits anonymous
requests. Wait a minute and try again, or use the manual link to fetch
the CoA directly from the vendor's site.

---

## 8. Shutting it down

To stop the server, switch to the Terminal / Command Prompt window that's
running it and press **Ctrl + C** (or just close the window).

---

## 9. Privacy & network

- The tool runs entirely on your computer — there is no Anthropic, Google,
  or other cloud component.
- The only outbound network requests are to the three vendor sites
  (`biolegend.com`, `regdocs.bd.com`, `assets.miltenyibiotec.com`) for the
  CoAs themselves.
- Scan history is held in memory only and disappears when you close the
  server. Export to Excel/CSV if you need to keep it.

---

## 10. Getting help

- Read the **README.md** for a one-page summary.
- Check the Terminal / Command Prompt window for HTTP logs — they almost
  always pinpoint where a lookup fails.
- File an issue on the project's GitHub repository (see README) with the
  failing barcode and the relevant log lines.
