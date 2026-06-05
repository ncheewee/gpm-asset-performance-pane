# GPM Asset Performance Pane

Static UAT prototype for portfolio performance review across assets, countries, and asset types.

## Run Locally

```sh
python3 -m http.server 5187
```

Then open:

```text
http://localhost:5187/index.html
```

## Data

The editable UAT data sheet lives in Google Drive:

```text
https://docs.google.com/spreadsheets/d/10m7sdj9aaPqG82xpeBWbx4es4zOh_Aadmmxbxl1RB_M/edit?gid=1349902348#gid=1349902348
```

Local workbook and CSV copies are also kept in GitHub as reference backups:

```text
data/portfolio-data.xlsx
data/portfolio-data.csv
```

The app loads the Google Sheet at runtime through Google Visualization JSONP, then falls back to embedded sample data if the Sheet is not reachable. For live UAT updates, keep the Sheet set to "Anyone with the link can view".

## GitHub Pages

Publish from the `main` branch root.
