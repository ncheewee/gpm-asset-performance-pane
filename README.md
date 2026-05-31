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

The app tries to load the Google Sheet as CSV at runtime, then falls back to embedded sample data if the Sheet is not published/reachable. For live UAT updates, publish the Google Sheet to the web or otherwise make the CSV endpoint readable by the dashboard.

## GitHub Pages

Publish from the `main` branch root.
