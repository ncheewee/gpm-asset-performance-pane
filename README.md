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
https://docs.google.com/spreadsheets/d/1Mwb7E5CXmoQtvB7yq4R7DmyH-O1qnxTb54NtVe9k34M/edit?gid=0#gid=0
```

Local workbook and CSV copies are also kept in GitHub as reference backups:

```text
data/portfolio-data.xlsx
data/portfolio-data.csv
```

For now, the app uses embedded sample data for speed and reliability. The Google Sheet, workbook, and CSV mirror the same structure so the Sheet can become the source of truth in the next pass.

## GitHub Pages

Publish from the `main` branch root.
