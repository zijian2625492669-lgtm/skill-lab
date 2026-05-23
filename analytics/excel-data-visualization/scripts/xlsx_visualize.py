#!/usr/bin/env python3
"""Profile Excel workbooks, generate charts, and run basic SPSS-style statistics."""

from __future__ import annotations

import argparse
import json
import math
import warnings
from pathlib import Path
from typing import Iterable

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
import statsmodels.api as sm
from scipy import stats


def read_sheet(path: Path, sheet: str | None) -> pd.DataFrame:
    try:
        return pd.read_excel(path, sheet_name=sheet or 0)
    except ImportError as exc:
        raise SystemExit("Missing Excel dependency. Install pandas and openpyxl.") from exc


def workbook_sheet_names(path: Path) -> list[str]:
    with pd.ExcelFile(path) as xls:
        return list(xls.sheet_names)


def infer_type(series: pd.Series) -> str:
    non_null = series.dropna()
    if non_null.empty:
        return "empty"
    if pd.api.types.is_datetime64_any_dtype(series):
        return "datetime"
    if pd.api.types.is_numeric_dtype(series):
        unique_ratio = non_null.nunique(dropna=True) / max(len(non_null), 1)
        if unique_ratio > 0.9 and len(non_null) > 20:
            return "numeric-id"
        return "numeric"

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        parsed_dates = pd.to_datetime(non_null, errors="coerce")
    if parsed_dates.notna().mean() >= 0.8:
        return "datetime"

    unique_count = non_null.astype(str).nunique(dropna=True)
    unique_ratio = unique_count / max(len(non_null), 1)
    if unique_count <= 30 and unique_ratio <= 0.5:
        return "categorical"
    return "text"


def profile_dataframe(df: pd.DataFrame) -> list[dict[str, object]]:
    rows = len(df)
    fields: list[dict[str, object]] = []
    for column in df.columns:
        series = df[column]
        non_null = int(series.notna().sum())
        sample = [str(value) for value in series.dropna().head(5).tolist()]
        fields.append(
            {
                "name": str(column),
                "dtype": str(series.dtype),
                "inferred_type": infer_type(series),
                "non_null": non_null,
                "missing": rows - non_null,
                "missing_pct": round((rows - non_null) / rows * 100, 2) if rows else 0,
                "unique": int(series.nunique(dropna=True)),
                "sample": sample,
            }
        )
    return fields


def profile_workbook(path: Path, sheet: str | None) -> dict[str, object]:
    names = workbook_sheet_names(path)
    selected = [sheet] if sheet else names
    result = {"workbook": str(path), "sheets": []}
    for sheet_name in selected:
        df = read_sheet(path, sheet_name)
        result["sheets"].append(
            {
                "name": sheet_name,
                "rows": int(len(df)),
                "columns": int(len(df.columns)),
                "fields": profile_dataframe(df),
            }
        )
    return result


def print_markdown(profile: dict[str, object]) -> None:
    print(f"# Workbook: {profile['workbook']}")
    for sheet in profile["sheets"]:
        print(f"\n## Sheet: {sheet['name']}")
        print(f"Rows: {sheet['rows']}  Columns: {sheet['columns']}\n")
        print("| Field | Type | Missing | Unique | Sample |")
        print("| --- | --- | ---: | ---: | --- |")
        for field in sheet["fields"]:
            sample = ", ".join(field["sample"])
            print(
                f"| `{field['name']}` | {field['inferred_type']} ({field['dtype']}) | "
                f"{field['missing_pct']}% | {field['unique']} | {sample} |"
            )


def parse_columns(raw: str | None) -> list[str] | None:
    if not raw:
        return None
    return [item.strip() for item in raw.split(",") if item.strip()]


def format_number(value: object, digits: int = 4) -> str:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, (int, np.integer)):
        return str(value)
    if isinstance(value, (float, np.floating)):
        if 0 < abs(value) < 10**-digits:
            return f"<{10**-digits:.{digits}f}"
        return f"{value:.{digits}f}".rstrip("0").rstrip(".")
    return str(value)


def markdown_table(rows: list[dict[str, object]], columns: list[str]) -> str:
    lines = [
        "| " + " | ".join(columns) + " |",
        "| " + " | ".join("---" for _ in columns) + " |",
    ]
    for row in rows:
        lines.append("| " + " | ".join(format_number(row.get(column)) for column in columns) + " |")
    return "\n".join(lines)


def write_analysis_result(result: dict[str, object], args: argparse.Namespace) -> None:
    if args.output:
        output = Path(args.output)
        output.parent.mkdir(parents=True, exist_ok=True)
        if args.format == "json":
            output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        else:
            output.write_text(result["markdown"], encoding="utf-8")
        print(f"saved: {output}")
        return

    if args.format == "json":
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(result["markdown"])


def ensure_columns(df: pd.DataFrame, columns: Iterable[str]) -> None:
    missing = [column for column in columns if column and column not in df.columns]
    if missing:
        raise SystemExit(f"Missing columns: {', '.join(missing)}")


def clean_for_columns(df: pd.DataFrame, columns: list[str]) -> tuple[pd.DataFrame, int]:
    before = len(df)
    cleaned = df.dropna(subset=columns)
    return cleaned, before - len(cleaned)


def aggregate_category(
    df: pd.DataFrame, category: str, value: str | None, agg: str, top_n: int
) -> pd.DataFrame:
    if value:
        if agg == "count":
            grouped = df.groupby(category, dropna=False)[value].count()
        else:
            grouped = df.groupby(category, dropna=False)[value].agg(agg)
    else:
        grouped = df.groupby(category, dropna=False).size()
        value = "count"

    data = grouped.sort_values(ascending=False).reset_index(name=value)
    if top_n and len(data) > top_n:
        top = data.head(top_n)
        other_value = data.iloc[top_n:][value].sum()
        other = pd.DataFrame([{category: "Other", value: other_value}])
        data = pd.concat([top, other], ignore_index=True)
    return data


def save_current_figure(output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    plt.tight_layout()
    plt.savefig(output, dpi=180, bbox_inches="tight")
    plt.close()
    print(f"saved: {output}")


def chart_pie(df: pd.DataFrame, args: argparse.Namespace) -> None:
    ensure_columns(df, [args.category] + ([args.value] if args.value else []))
    plot_df, dropped = clean_for_columns(df, [args.category] + ([args.value] if args.value else []))
    data = aggregate_category(plot_df, args.category, args.value, args.agg, args.top_n)
    value_col = args.value if args.value else "count"
    plt.figure(figsize=(9, 7))
    plt.pie(data[value_col], labels=data[args.category].astype(str), autopct="%1.1f%%", startangle=90)
    plt.title(args.title or f"{value_col} by {args.category}")
    save_current_figure(Path(args.output))
    if dropped:
        print(f"rows_dropped_missing_fields: {dropped}")


def chart_bar(df: pd.DataFrame, args: argparse.Namespace) -> None:
    ensure_columns(df, [args.category] + ([args.value] if args.value else []))
    plot_df, dropped = clean_for_columns(df, [args.category] + ([args.value] if args.value else []))
    data = aggregate_category(plot_df, args.category, args.value, args.agg, args.top_n)
    value_col = args.value if args.value else "count"
    plt.figure(figsize=(10, 6))
    sns.barplot(data=data, x=args.category, y=value_col)
    plt.xticks(rotation=35, ha="right")
    plt.title(args.title or f"{value_col} by {args.category}")
    save_current_figure(Path(args.output))
    if dropped:
        print(f"rows_dropped_missing_fields: {dropped}")


def chart_line(df: pd.DataFrame, args: argparse.Namespace) -> None:
    ensure_columns(df, [args.x, args.y] + ([args.group] if args.group else []))
    plot_df, dropped = clean_for_columns(df, [args.x, args.y])
    plot_df = plot_df.sort_values(args.x)
    plt.figure(figsize=(11, 6))
    if args.group:
        sns.lineplot(data=plot_df, x=args.x, y=args.y, hue=args.group, marker="o")
    else:
        sns.lineplot(data=plot_df, x=args.x, y=args.y, marker="o")
    plt.xticks(rotation=30, ha="right")
    plt.title(args.title or f"{args.y} over {args.x}")
    save_current_figure(Path(args.output))
    if dropped:
        print(f"rows_dropped_missing_fields: {dropped}")


def chart_scatter(df: pd.DataFrame, args: argparse.Namespace) -> None:
    ensure_columns(df, [args.x, args.y] + ([args.group] if args.group else []))
    plot_df, dropped = clean_for_columns(df, [args.x, args.y])
    plt.figure(figsize=(9, 7))
    sns.scatterplot(data=plot_df, x=args.x, y=args.y, hue=args.group if args.group else None)
    plt.title(args.title or f"{args.y} vs {args.x}")
    save_current_figure(Path(args.output))
    if dropped:
        print(f"rows_dropped_missing_fields: {dropped}")


def chart_corr_heatmap(df: pd.DataFrame, args: argparse.Namespace) -> None:
    columns = parse_columns(args.columns)
    numeric_df = df.select_dtypes(include="number")
    if columns:
        ensure_columns(df, columns)
        numeric_df = df[columns].apply(pd.to_numeric, errors="coerce")
    if numeric_df.shape[1] < 2:
        raise SystemExit("Correlation heatmap requires at least two numeric columns.")
    corr = numeric_df.corr(numeric_only=True)
    size = max(7, min(14, 0.6 * len(corr.columns) + 4))
    plt.figure(figsize=(size, size))
    sns.heatmap(corr, annot=True, fmt=".2f", cmap="RdBu_r", center=0, square=True, linewidths=0.5)
    plt.title(args.title or "Correlation heatmap")
    save_current_figure(Path(args.output))


def split_value(series: pd.Series, method: str) -> float:
    if method == "median":
        return float(series.median())
    if method == "zero":
        return 0.0
    return float(series.mean())


def chart_quadrant(df: pd.DataFrame, args: argparse.Namespace) -> None:
    ensure_columns(df, [args.x, args.y] + ([args.label] if args.label else []))
    plot_df, dropped = clean_for_columns(df, [args.x, args.y])
    x_split = split_value(plot_df[args.x], args.split)
    y_split = split_value(plot_df[args.y], args.split)
    plt.figure(figsize=(10, 8))
    sns.scatterplot(data=plot_df, x=args.x, y=args.y)
    plt.axvline(x_split, color="#555555", linestyle="--", linewidth=1)
    plt.axhline(y_split, color="#555555", linestyle="--", linewidth=1)
    if args.label:
        for _, row in plot_df.iterrows():
            if not (math.isfinite(row[args.x]) and math.isfinite(row[args.y])):
                continue
            plt.text(row[args.x], row[args.y], str(row[args.label]), fontsize=8, alpha=0.75)
    plt.title(args.title or f"Quadrant: {args.y} vs {args.x}")
    save_current_figure(Path(args.output))
    print(f"x_split_{args.split}: {x_split}")
    print(f"y_split_{args.split}: {y_split}")
    if dropped:
        print(f"rows_dropped_missing_fields: {dropped}")


def numeric_analysis_frame(df: pd.DataFrame, columns: list[str] | None) -> pd.DataFrame:
    if columns:
        ensure_columns(df, columns)
        frame = df[columns].apply(pd.to_numeric, errors="coerce")
    else:
        frame = df.select_dtypes(include="number")
    if frame.empty:
        raise SystemExit("No numeric columns available for this analysis.")
    return frame


def analysis_descriptive(df: pd.DataFrame, args: argparse.Namespace) -> dict[str, object]:
    columns = parse_columns(args.columns)
    frame = numeric_analysis_frame(df, columns)
    rows: list[dict[str, object]] = []
    for column in frame.columns:
        series = frame[column].dropna()
        rows.append(
            {
                "Variable": column,
                "N": int(series.count()),
                "Missing": int(frame[column].isna().sum()),
                "Mean": series.mean(),
                "Std. Deviation": series.std(ddof=1),
                "Minimum": series.min(),
                "Q1": series.quantile(0.25),
                "Median": series.median(),
                "Q3": series.quantile(0.75),
                "Maximum": series.max(),
                "Skewness": series.skew(),
                "Kurtosis": series.kurt(),
            }
        )
    columns_out = [
        "Variable",
        "N",
        "Missing",
        "Mean",
        "Std. Deviation",
        "Minimum",
        "Q1",
        "Median",
        "Q3",
        "Maximum",
        "Skewness",
        "Kurtosis",
    ]
    markdown = "# Descriptive Statistics\n\n" + markdown_table(rows, columns_out)
    return {"analysis": "descriptive", "rows": rows, "markdown": markdown}


def pairwise_correlation(x: pd.Series, y: pd.Series, method: str) -> tuple[float, float, int]:
    valid = pd.concat([x, y], axis=1).dropna()
    n = len(valid)
    if n < 3:
        return np.nan, np.nan, n
    if method == "spearman":
        coef, p_value = stats.spearmanr(valid.iloc[:, 0], valid.iloc[:, 1])
    else:
        coef, p_value = stats.pearsonr(valid.iloc[:, 0], valid.iloc[:, 1])
    return float(coef), float(p_value), n


def analysis_correlation(df: pd.DataFrame, args: argparse.Namespace) -> dict[str, object]:
    columns = parse_columns(args.columns)
    frame = numeric_analysis_frame(df, columns)
    if frame.shape[1] < 2:
        raise SystemExit("Correlation analysis requires at least two numeric columns.")

    rows: list[dict[str, object]] = []
    names = list(frame.columns)
    for i, left in enumerate(names):
        for right in names[i + 1 :]:
            coef, p_value, n = pairwise_correlation(frame[left], frame[right], args.corr_method)
            rows.append(
                {
                    "Variable 1": left,
                    "Variable 2": right,
                    "Method": args.corr_method,
                    "N": n,
                    "Correlation": coef,
                    "Sig. (2-tailed)": p_value,
                }
            )

    markdown = (
        f"# Correlation Analysis ({args.corr_method})\n\n"
        + markdown_table(rows, ["Variable 1", "Variable 2", "Method", "N", "Correlation", "Sig. (2-tailed)"])
    )
    return {"analysis": "correlation", "method": args.corr_method, "rows": rows, "markdown": markdown}


def analysis_regression(df: pd.DataFrame, args: argparse.Namespace) -> dict[str, object]:
    predictors = parse_columns(args.x)
    if not args.y or not predictors:
        raise SystemExit("Linear regression requires --y and --x predictor columns.")
    ensure_columns(df, [args.y] + predictors)

    model_df = df[[args.y] + predictors].apply(pd.to_numeric, errors="coerce").dropna()
    if len(model_df) <= len(predictors) + 1:
        raise SystemExit("Not enough complete rows for linear regression.")

    y = model_df[args.y]
    x = sm.add_constant(model_df[predictors], has_constant="add")
    model = sm.OLS(y, x).fit()

    model_summary = [
        {"Metric": "Dependent Variable", "Value": args.y},
        {"Metric": "N", "Value": int(model.nobs)},
        {"Metric": "R", "Value": math.sqrt(max(model.rsquared, 0))},
        {"Metric": "R Square", "Value": model.rsquared},
        {"Metric": "Adjusted R Square", "Value": model.rsquared_adj},
        {"Metric": "Std. Error of the Estimate", "Value": math.sqrt(model.mse_resid)},
        {"Metric": "F", "Value": model.fvalue},
        {"Metric": "Sig. F", "Value": model.f_pvalue},
        {"Metric": "Df Regression", "Value": int(model.df_model)},
        {"Metric": "Df Residual", "Value": int(model.df_resid)},
    ]
    coefficients: list[dict[str, object]] = []
    conf_int = model.conf_int()
    for name in model.params.index:
        coefficients.append(
            {
                "Variable": "Constant" if name == "const" else name,
                "B": model.params[name],
                "Std. Error": model.bse[name],
                "Beta": "" if name == "const" else standardized_beta(model_df, args.y, name),
                "t": model.tvalues[name],
                "Sig.": model.pvalues[name],
                "95% CI Lower": conf_int.loc[name, 0],
                "95% CI Upper": conf_int.loc[name, 1],
            }
        )

    markdown = (
        "# Linear Regression Analysis\n\n"
        "## Model Summary\n\n"
        + markdown_table(model_summary, ["Metric", "Value"])
        + "\n\n## Coefficients\n\n"
        + markdown_table(
            coefficients,
            ["Variable", "B", "Std. Error", "Beta", "t", "Sig.", "95% CI Lower", "95% CI Upper"],
        )
    )
    return {
        "analysis": "linear-regression",
        "dependent": args.y,
        "predictors": predictors,
        "model_summary": model_summary,
        "coefficients": coefficients,
        "markdown": markdown,
    }


def standardized_beta(model_df: pd.DataFrame, y_column: str, x_column: str) -> float:
    y_std = model_df[y_column].std(ddof=1)
    x_std = model_df[x_column].std(ddof=1)
    if y_std == 0 or x_std == 0:
        return np.nan
    y_z = (model_df[y_column] - model_df[y_column].mean()) / y_std
    x_columns = [column for column in model_df.columns if column != y_column]
    x_z = (model_df[x_columns] - model_df[x_columns].mean()) / model_df[x_columns].std(ddof=1)
    fitted = sm.OLS(y_z, sm.add_constant(x_z, has_constant="add")).fit()
    return float(fitted.params[x_column])


def run_profile(args: argparse.Namespace) -> None:
    profile = profile_workbook(Path(args.input), args.sheet)
    if args.format == "json":
        print(json.dumps(profile, ensure_ascii=False, indent=2))
    else:
        print_markdown(profile)


def run_chart(args: argparse.Namespace) -> None:
    df = read_sheet(Path(args.input), args.sheet)
    chart_map = {
        "pie": chart_pie,
        "bar": chart_bar,
        "line": chart_line,
        "scatter": chart_scatter,
        "corr-heatmap": chart_corr_heatmap,
        "quadrant": chart_quadrant,
    }
    chart_map[args.kind](df, args)


def run_analyze(args: argparse.Namespace) -> None:
    df = read_sheet(Path(args.input), args.sheet)
    analysis_map = {
        "descriptive": analysis_descriptive,
        "correlation": analysis_correlation,
        "linear-regression": analysis_regression,
    }
    result = analysis_map[args.method](df, args)
    write_analysis_result(result, args)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    profile = subparsers.add_parser("profile", help="List sheets and profile fields.")
    profile.add_argument("--input", required=True, help="Path to .xlsx/.xls workbook.")
    profile.add_argument("--sheet", help="Sheet name. Defaults to all sheets.")
    profile.add_argument("--format", choices=["markdown", "json"], default="markdown")
    profile.set_defaults(func=run_profile)

    chart = subparsers.add_parser("chart", help="Generate a chart PNG from a workbook sheet.")
    chart.add_argument("--input", required=True)
    chart.add_argument("--sheet", help="Sheet name. Defaults to first sheet.")
    chart.add_argument("--kind", required=True, choices=["pie", "bar", "line", "scatter", "corr-heatmap", "quadrant"])
    chart.add_argument("--output", required=True)
    chart.add_argument("--x")
    chart.add_argument("--y")
    chart.add_argument("--category")
    chart.add_argument("--value")
    chart.add_argument("--group")
    chart.add_argument("--label")
    chart.add_argument("--columns", help="Comma-separated numeric columns for correlation heatmap.")
    chart.add_argument("--agg", choices=["sum", "mean", "median", "count"], default="sum")
    chart.add_argument("--top-n", type=int, default=12)
    chart.add_argument("--split", choices=["mean", "median", "zero"], default="mean")
    chart.add_argument("--title")
    chart.set_defaults(func=run_chart)

    analyze = subparsers.add_parser("analyze", help="Run SPSS-style basic statistical analyses.")
    analyze.add_argument("--input", required=True)
    analyze.add_argument("--sheet", help="Sheet name. Defaults to first sheet.")
    analyze.add_argument(
        "--method",
        required=True,
        choices=["descriptive", "correlation", "linear-regression"],
    )
    analyze.add_argument("--columns", help="Comma-separated numeric columns for descriptive/correlation.")
    analyze.add_argument("--corr-method", choices=["pearson", "spearman"], default="pearson")
    analyze.add_argument("--y", help="Dependent variable for linear regression.")
    analyze.add_argument("--x", help="Comma-separated predictors for linear regression.")
    analyze.add_argument("--format", choices=["markdown", "json"], default="markdown")
    analyze.add_argument("--output", help="Optional output file for markdown or JSON results.")
    analyze.set_defaults(func=run_analyze)
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    required_by_kind = {
        "pie": ["category"],
        "bar": ["category"],
        "line": ["x", "y"],
        "scatter": ["x", "y"],
        "quadrant": ["x", "y"],
    }
    if getattr(args, "command", None) == "chart":
        missing = [name for name in required_by_kind.get(args.kind, []) if not getattr(args, name)]
        if missing:
            parser.error(f"--kind {args.kind} requires: {', '.join('--' + name for name in missing)}")
    args.func(args)


if __name__ == "__main__":
    main()
