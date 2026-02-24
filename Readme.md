# Superstore Sales Data Analysis


## Objective

A Python script that takes the Superstore dataset and turns it into a proper P&L breakdown with summaries, charts, and an excel export in one run.

## Summary

- load the raw CSV 'Sample_Superstore.csv'
- cleans the data and add 'Profit Margin' and 'COGS' columns to the dataframe
- generate two summaries with describe() and groupby()
- create two visuals: bar chart for top sales and sankey chart for P&L
- produce 4 exports: 'Superstore Analysis Output.xlsx', 'bar_chart.png', 'sankey-chart.png', 'sankey_chart.html'

## Requirements
- **pandas** — data loading, cleaning, groupby aggregations, and Excel export
- **numpy** — used for `np.where()` to handle conditional column calculations
- **matplotlib** — builds the horizontal bar chart for top 10 products by sales
- **plotly** — builds the interactive Sankey chart
- **kaleido** — Enables exporting static PNG images via `write_image()`.
- **xlsxwriter** — Enables `pd.ExcelWriter` to build and write the Excel workbook
- **os** — Used to check the current working directory
- **io** — Used to handle in-memory image buffers for embedding charts in Excel

## Notes
- Sankey uses fixed positions on the visual. If labels start to overlap, please use 'arrangement=snap'