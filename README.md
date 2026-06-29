# CoA Lookup Tool

**Multi-vendor Certificate of Analysis lookup for antibody inventories.**
Scan a 2D barcode or upload a spreadsheet — the tool fetches the matching
CoA from **BioLegend**, **BD Biosciences**, or **Miltenyi Biotec** and
fills in product name, clone, fluorochrome, isotype, expiry, and more.

Runs locally on your computer. No cloud, no account, no data leaves your
machine except the request to the vendor itself.

---

## Quick start

### macOS
Double-click **`setup_and_run.command`**.

### Windows
Double-click **`setup_and_run.bat`**.
(The Windows launcher installs Python 3.12 for you if it isn't already.)

### Manual
```bash
pip install -r requirements.txt
python app.py
```

The tool opens in your browser at **http://localhost:5050**.

---

## Features

- **Single Scan tab** — point a USB/Bluetooth barcode scanner at a label,
  the tool parses it and pulls the CoA in one shot.
- **Batch tab** — drop an `.xlsx` inventory file in and download a
  filled-in spreadsheet, with newly populated cells highlighted in light
  purple.
- **Scan history** with selective export to CSV or Excel.
- **Vendor-aware** parsing of three barcode standards:
  - BioLegend space-separated (`344742 B402098 2026/03/05`)
  - BD GS1 DataMatrix (`010038290…172707311050508742400026`)
  - Miltenyi AI 91 (`9165613012371052501018320…`)

---

## Files

```
coa_lookup/
├── app.py                 ← Flask server + HTML UI (single file)
├── requirements.txt       ← Python dependencies
├── setup_and_run.command  ← macOS one-click launcher
├── setup_and_run.bat      ← Windows one-click launcher
├── README.md              ← This file
├── USER_MANUAL.md         ← Full user guide
└── LICENSE                ← MIT
```

See **[USER_MANUAL.md](USER_MANUAL.md)** for the full walkthrough,
field reference, and troubleshooting.

---

## Requirements

- macOS 12+ or Windows 10/11
- Python 3.10+ (auto-installed on Windows if missing)
- Internet connection (to reach vendor sites)

---

## License

MIT — see [LICENSE](LICENSE).

## Acknowledgements

CoAs are fetched directly from each vendor's public documents portal. This
tool is not affiliated with or endorsed by BioLegend, BD Biosciences, or
Miltenyi Biotec.
