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

The editable UAT data workbook lives in:

```text
data/portfolio-data.xlsx
```

A CSV fallback is also kept in:

```text
data/portfolio-data.csv
```

For now, the app uses embedded sample data for speed and reliability. The workbook and CSV mirror the same structure so either can become the source of truth in the next pass.

## GitHub Pages

Publish from the `main` branch root.
