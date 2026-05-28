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

The editable UAT data lives in:

```text
data/portfolio-data.csv
```

For now, the app uses embedded sample data for speed and reliability. The CSV mirrors the same structure so it can become the source of truth in the next pass.

## GitHub Pages

Publish from the `main` branch root.
