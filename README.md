# RAV Regulation Framework and Regulation Analysis

A static GitHub Pages template for exploring automated-vehicle regulation across all 50 states.

Live site: https://xinyangdu0210.github.io/RAV-regulation/

Current release: Version 1.0 (September 2, 2026)

## Publish

1. Push this folder to a GitHub repository on the `main` branch.
2. Before the first workflow run, open **Settings → Pages** and select **GitHub Actions** under **Build and deployment → Source**.
3. Run the workflow again. It publishes `index.html`, `styles.css`, `script.js`, `assets/`, and `data/`.

## Refresh the data

Place the updated workbook at `RAV-Task1.2-50states.xlsx`, then run:

```powershell
.\tools\extract-regulations.ps1
```

The original Office files are not included in the public Pages artifact.
