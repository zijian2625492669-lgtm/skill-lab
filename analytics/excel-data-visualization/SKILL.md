---
name: excel-data-visualization
description: Read, profile, analyze, and visualize Excel workbooks (.xlsx/.xls) by returning available sheets and fields, helping users choose columns, generating charts such as pie charts, line charts, bar charts, scatter plots, correlation heatmaps, and quadrant charts, and running SPSS-style basic statistics including descriptive statistics, correlation analysis, and linear regression. Use when the user provides Excel data and asks for field inspection, chart recommendations, dashboard-style analysis, exploratory data analysis, statistical analysis, correlation, regression, or visual reports.
---

# Excel Data Visualization

## Workflow

1. Inspect the workbook before choosing chart types:
   ```bash
   python scripts/xlsx_visualize.py profile --input "workbook.xlsx" --format markdown
   ```
2. Report the available sheets and fields to the user. Include inferred field types, missingness, sample values, and which columns are likely suitable for dimensions, measures, dates, and labels.
3. Ask the user to choose fields when intent is underspecified. If intent is clear, choose defensible defaults and state them.
4. Generate charts with `scripts/xlsx_visualize.py chart` or run statistics with `scripts/xlsx_visualize.py analyze`.
5. Summarize what each chart or statistical table shows and call out data quality issues that affect interpretation.

## Field Selection Rules

- Use numeric fields for measures, correlation heatmaps, scatter axes, and quadrant axes.
- Use low-cardinality categorical fields for pie slices, grouped bars, legends, and segment comparisons.
- Use date/time fields for line chart x-axes when time trend analysis is requested.
- Avoid pie charts when there are too many categories; aggregate to top N and combine the rest into `Other`.
- For correlation heatmaps, include numeric columns only and exclude IDs or near-unique numeric keys when they are obvious.
- For quadrant charts, use two numeric metrics as `x` and `y`; use a label column only when labels are readable at the chart size.
- For descriptive statistics, use numeric fields and report N, missing values, mean, standard deviation, quantiles, skewness, and kurtosis.
- For correlation analysis, use Pearson by default for approximately linear numeric relationships; use Spearman when rank/order or non-normal monotonic relationships are more appropriate.
- For linear regression, require one numeric dependent variable and one or more numeric predictors. Mention that significance does not prove causality.

## Script Usage

Profile a workbook:
```bash
python scripts/xlsx_visualize.py profile --input "sales.xlsx" --sheet "Sheet1" --format markdown
```

Generate common charts:
```bash
python scripts/xlsx_visualize.py chart --input "sales.xlsx" --sheet "Sheet1" --kind pie --category "region" --value "revenue" --output "charts/revenue_by_region.png"
python scripts/xlsx_visualize.py chart --input "sales.xlsx" --sheet "Sheet1" --kind line --x "date" --y "revenue" --group "channel" --output "charts/revenue_trend.png"
python scripts/xlsx_visualize.py chart --input "sales.xlsx" --sheet "Sheet1" --kind corr-heatmap --columns "revenue,profit,orders" --output "charts/correlation.png"
python scripts/xlsx_visualize.py chart --input "sales.xlsx" --sheet "Sheet1" --kind quadrant --x "growth_rate" --y "profit_margin" --label "product" --output "charts/quadrant.png"
```

Supported `--kind` values: `pie`, `bar`, `line`, `scatter`, `corr-heatmap`, `quadrant`.

Run SPSS-style basic statistics:
```bash
python scripts/xlsx_visualize.py analyze --input "sales.xlsx" --method descriptive --columns "revenue,profit,growth_rate" --output "stats/descriptive.md"
python scripts/xlsx_visualize.py analyze --input "sales.xlsx" --method correlation --columns "revenue,profit,growth_rate" --corr-method pearson --output "stats/correlation.md"
python scripts/xlsx_visualize.py analyze --input "sales.xlsx" --method linear-regression --y "profit" --x "revenue,growth_rate" --output "stats/regression.md"
```

Supported `--method` values: `descriptive`, `correlation`, `linear-regression`.

Useful options:
- `--agg sum|mean|median|count` controls category aggregation.
- `--top-n 12` limits crowded categorical charts.
- `--split mean|median|zero` controls quadrant divider lines.
- `--columns "a,b,c"` selects numeric columns for correlation heatmaps.
- `--title "..."` overrides the generated chart title.
- `--format markdown|json` controls statistical output format.
- `--output "stats/result.md"` writes statistical output to a file.

## Output Standards

- Prefer clear file names under a `charts/` folder near the source workbook.
- Prefer statistical tables under a `stats/` folder near the source workbook.
- Use UTF-8-safe column names exactly as they appear in the workbook.
- Do not silently drop rows; mention rows removed for missing chart fields.
- Do not overclaim causality from correlations or quadrant positions.
- For regression, report dependent variable, predictors, N, R, R Square, adjusted R Square, F test, coefficient B, standardized Beta, t, p-value, and confidence interval.
