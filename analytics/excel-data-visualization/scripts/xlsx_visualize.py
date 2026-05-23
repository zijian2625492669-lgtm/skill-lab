#!/usr/bin/env python3
"""Profile Excel workbooks and generate common exploratory charts."""

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
import pandas as pd
import seaborn as sns


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
