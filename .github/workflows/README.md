# Philippine CCET WordPress Dashboard — code-only starter

This repository is intentionally small. It does **not** contain the CCET Excel workbook, generated JSON, `node_modules`, or build artifacts.

## Architecture

1. GitHub Actions downloads the source CCET `.xlsx` URL stored in the repository variable `CCET_XLSX_URL`.
2. `scripts/build_data.py` normalizes the FY2017–FY2026 workbook sheets and generates compact JSON summaries.
3. `site/` renders the data with Chart.js and Leaflet.
4. GitHub Pages serves the dashboard.
5. WordPress.com embeds the GitHub Pages URL in a Custom HTML block using an `<iframe>`.

## First-time GitHub setup

1. Create a new empty GitHub repository.
2. Upload only the files/folders in this starter package. Do not upload the Excel workbook.
3. Go to **Settings → Secrets and variables → Actions → Variables → New repository variable**.
4. Name: `CCET_XLSX_URL`.
5. Value: a **direct public URL** that downloads the `.xlsx` file (not a web preview page).
6. Go to **Settings → Pages → Build and deployment → Source → GitHub Actions**.
7. Go to **Actions → Build and Deploy CCET Dashboard → Run workflow**.
8. After the green check mark, open **Settings → Pages**. Your live URL is shown there, normally `https://YOUR-USERNAME.github.io/YOUR-REPOSITORY/`.

The workflow also runs daily at 08:15 Asia/Manila time. Change the cron expression if desired.

## Local test (optional)

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
# source .venv/bin/activate
pip install -r requirements.txt
python scripts/build_data.py --input path/to/ccet.xlsx --out site/data
python -m http.server 8000 --directory site
```

Open `http://localhost:8000`.

## WordPress embed

On a WordPress.com paid plan with hosting features activated, add a **Custom HTML** block and paste:

```html
<div style="width:100%;max-width:1600px;margin:0 auto;">
  <iframe
    src="https://YOUR-USERNAME.github.io/YOUR-REPOSITORY/"
    title="Philippine Climate Budget Explorer"
    loading="lazy"
    style="width:100%;height:2200px;border:0;border-radius:12px;overflow:hidden;"
    allowfullscreen>
  </iframe>
</div>
```

For best results, use a full-width WordPress page template and remove the page sidebar.

## If the source file URL changes

You do **not** edit code. Update **Settings → Secrets and variables → Actions → Variables → `CCET_XLSX_URL`**, then run the workflow manually.

## Why this avoids GitHub's file-size problem

The repository stores only code. The Excel workbook is downloaded temporarily on the GitHub Actions runner and discarded. The generated JSON is uploaded directly as the Pages deployment artifact and is not committed to Git history.

## Map caveat

The Excel workbook does not provide a clean, consistent geographic field for every PAP. The starter infers a province/city from words in the PAP/agency text and matches those names against geoBoundaries ADM2 reference names. Therefore the map is exploratory, not an authoritative geocoded inventory. For an official map, add a structured `province`, `city_municipality`, and ideally project coordinates to the CCET source data.
