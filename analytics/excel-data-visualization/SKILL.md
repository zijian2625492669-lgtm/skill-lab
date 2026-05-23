---
name: excel-data-visualization
description: Read, profile, analyze, and visualize Excel workbooks (.xlsx/.xls) by returning available sheets and fields, helping users choose columns, and generating charts such as pie charts, line charts, bar charts, scatter plots, correlation heatmaps, and quadrant charts. Use when the user provides Excel data and asks for field inspection, chart recommendations, dashboard-style analysis, exploratory data analysis, or visual reports.
---

# Excel Data Visualization

## Workflow

1. Inspect the workbook before choosing chart types:
   ```bash
   python scripts/xlsx_visualize.py profile --input "workbook.xlsx" --format markdown
   ```
2. Report the available sheets and fields to the user. Include inferred field types, missingness, sample values, and which columns are likely suitable for dimensions, measures, dates, and labels.
3. Ask the user to choose fields when intent is underspecified. If intent is clear, choose defensible defaults and state them.
4. Generate charts with `scripts/xlsx_visualize.py chart`. Save outputs as PNG files unless the user requests another format.
5. Summarize what each chart shows and call out data quality issues that affect interpretation.

## Field Selection Rules

- Use numeric fields for measures, correlation heatmaps, scatter axes, and quadrant axes.
- Use low-cardinality categorical fields for pie slices, grouped bars, legends, and segment comparisons.
- Use date/time fields for line chart x-axes when time trend analysis is requested.
- Avoid pie charts when there are too many categories; aggregate to top N and combine the rest into `Other`.
- For correlation heatmaps, include numeric columns only and exclude IDs or near-unique numeric keys when they are obvious.
- For quadrant charts, use two numeric metrics as `x` and `y`; use a label column only when labels are readable at the chart size.

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

Useful options:
- `--agg sum|mean|median|count` controls category aggregation.
- `--top-n 12` limits crowded categorical charts.
- `--split mean|median|zero` controls quadrant divider lines.
- `--columns "a,b,c"` selects numeric columns for correlation heatmaps.
- `--title "..."` overrides the generated chart title.

## Output Standards

- Prefer clear file names under a `charts/` folder near the source workbook.
- Use UTF-8-safe column names exactly as they appear in the workbook.
- Do not silently drop rows; mention rows removed for missing chart fields.
- Do not overclaim causality from correlations or quadrant positions.
