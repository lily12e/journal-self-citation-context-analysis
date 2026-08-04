"""Export manuscript- and supplement-facing CSV files from analysis results."""

from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path


DIMENSION_ORDER = [
    "Citation position",
    "Citation distance",
    "Citation strength",
    "Citation function",
    "Citation depth",
    "Semantic similarity",
]
FIELD_NAMES = {
    "1": "Information Science & Library Science",
    "2": "Gerontology",
    "3": "Engineering, Multidisciplinary",
    "4": "Veterinary Sciences",
    "5": "Geography",
}
FIELD_CODES = {
    "1": ("G1-1", "G2-1"),
    "2": ("G1-2", "G2-2"),
    "3": ("G1-3", "G2-3"),
    "4": ("G1-4", "G2-4"),
    "5": ("G1-5", "G2-5"),
}
HYPOTHESES = {
    "Citation position": "H1a",
    "Citation distance": "H1b",
    "Citation strength": "H1c",
    "Citation function": "H2a",
    "Citation depth": "H2b",
    "Semantic similarity": "H2c",
}

# Preserve the original machine-readable annotation strings in the released
# data and analysis JSON, while using the manuscript's formal display labels
# in report-facing tables and figure source files.
DISPLAY_LABELS = {
    "Simple mention": "Simple Mention",
    "Historical background": "Historical Background",
    "Related work": "Related Work",
    "Shallow citation": "Shallow Citation",
}


def display_label(value: object) -> object:
    return DISPLAY_LABELS.get(value, value)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--analysis-results",
        type=Path,
        default=Path("analysis/results/analysis_results.json"),
    )
    parser.add_argument(
        "--annotation-results",
        type=Path,
        default=Path("analysis/results/annotation_evaluation.json"),
    )
    parser.add_argument(
        "--sample-selection",
        type=Path,
        default=Path("data/sample_selection_counts.csv"),
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("analysis/report_tables"),
    )
    return parser.parse_args()


def write_csv(path: Path, headers: list[str], rows: list[list[object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        writer.writerows(rows)


def significance(value: float) -> str:
    return "Yes" if value < 0.05 else "No"


def confusion_rows(result: dict[str, object]) -> tuple[list[str], list[list[object]]]:
    labels = result["labels"]
    matrix = result["confusionMatrix"]
    rows = [
        [display_label(actual)]
        + [matrix[actual][predicted] for predicted in labels]
        for actual in labels
    ]
    return ["True label"] + [display_label(label) for label in labels], rows


def discipline_source_rows(field_result: dict[str, object]) -> list[list[object]]:
    rows: list[list[object]] = []
    for dimension in DIMENSION_ORDER:
        result = field_result["dimensions"][dimension]
        if dimension == "Semantic similarity":
            rows.append(
                [
                    dimension,
                    "Mean score",
                    "",
                    result["mean1"],
                    result["sd1"],
                    "",
                    result["mean2"],
                    result["sd2"],
                ]
            )
            continue
        for index, category in enumerate(result["categories"]):
            rows.append(
                [
                    dimension,
                    display_label(category),
                    result["counts"][0][index],
                    result["percentages"][0][index],
                    "",
                    result["counts"][1][index],
                    result["percentages"][1][index],
                    "",
                ]
            )
    return rows


def main() -> int:
    args = parse_args()
    results = json.loads(args.analysis_results.read_text(encoding="utf-8"))
    annotation = json.loads(args.annotation_results.read_text(encoding="utf-8"))
    output_dir = args.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    with args.sample_selection.open(encoding="utf-8-sig", newline="") as handle:
        sample_rows = list(csv.DictReader(handle))
    sample_headers = list(sample_rows[0])
    write_csv(
        output_dir / "table_s1_sample_selection.csv",
        sample_headers,
        [[row[header] for header in sample_headers] for row in sample_rows],
    )

    field_by_code = {
        code: FIELD_NAMES[field]
        for field, codes in FIELD_CODES.items()
        for code in codes
    }
    table2_rows = [
        [
            field_by_code[row["Code"]],
            row["Code"],
            row["Journal"],
            row["Published articles"],
            row["Self-citing articles"],
            row["Self-cited article occurrences"],
            row["Self-citation mentions"],
        ]
        for row in sample_rows
        if row["Code"] != "Total"
    ]
    write_csv(
        output_dir / "main_table2_sample_statistics.csv",
        [
            "JCR category",
            "Code",
            "Journal",
            "Published articles",
            "Self-citing articles",
            "Self-cited article occurrences",
            "Self-citation mentions",
        ],
        table2_rows,
    )

    table4_rows: list[list[object]] = []
    for dimension in DIMENSION_ORDER:
        result = results["aggregate"][dimension]
        if dimension == "Semantic similarity":
            table4_rows.append(
                [
                    dimension,
                    "Mean score",
                    "",
                    result["mean1"],
                    result["sd1"],
                    "",
                    result["mean2"],
                    result["sd2"],
                ]
            )
            continue
        for index, category in enumerate(result["categories"]):
            table4_rows.append(
                [
                    dimension,
                    display_label(category),
                    result["counts"][0][index],
                    result["percentages"][0][index],
                    "",
                    result["counts"][1][index],
                    result["percentages"][1][index],
                    "",
                ]
            )
    table4_headers = [
        "Dimension",
        "Category",
        "G1 count",
        "G1 proportion",
        "G1 SD",
        "G2 count",
        "G2 proportion",
        "G2 SD",
    ]
    write_csv(output_dir / "main_table4_aggregate_counts.csv", table4_headers, table4_rows)
    write_csv(output_dir / "figure6_source_data.csv", table4_headers, table4_rows)

    table5_rows: list[list[object]] = []
    for field in FIELD_NAMES:
        dimensions = results["discipline"][field]["dimensions"]
        flags = []
        for dimension in DIMENSION_ORDER:
            item = dimensions[dimension]
            p_value = item["p"] if dimension == "Semantic similarity" else item["test"]["p"]
            flags.append("Yes" if p_value < 0.05 else "No")
        table5_rows.append(
            [
                FIELD_NAMES[field],
                *flags,
                f"{sum(flag == 'Yes' for flag in flags)}/6",
            ]
        )
    write_csv(
        output_dir / "main_table5_significance_summary.csv",
        ["JCR category", *DIMENSION_ORDER, "Total"],
        table5_rows,
    )

    agreement = annotation["interAnnotatorAgreement"]
    model = annotation["heldOutModelEvaluation"]
    write_csv(
        output_dir / "table_s2_validation_summary.csv",
        ["Validation component", "Task", "Metric", "Value", "n"],
        [
            [
                "Human inter-annotator agreement",
                "Citation function",
                "Cohen's kappa",
                agreement["citationFunction"]["cohenKappa"],
                agreement["citationFunction"]["n"],
            ],
            [
                "Human inter-annotator agreement",
                "Citation depth",
                "Cohen's kappa",
                agreement["citationDepth"]["cohenKappa"],
                agreement["citationDepth"]["n"],
            ],
            [
                "GPT-4o held-out evaluation",
                "Citation function",
                "Overall accuracy",
                model["citationFunction"]["accuracy"],
                model["citationFunction"]["n"],
            ],
            [
                "GPT-4o held-out evaluation",
                "Citation function",
                "Weighted F1",
                model["citationFunction"]["weightedAverage"]["f1"],
                model["citationFunction"]["n"],
            ],
            [
                "GPT-4o held-out evaluation",
                "Citation depth",
                "Overall accuracy",
                model["citationDepth"]["accuracy"],
                model["citationDepth"]["n"],
            ],
            [
                "GPT-4o held-out evaluation",
                "Citation depth",
                "Weighted F1",
                model["citationDepth"]["weightedAverage"]["f1"],
                model["citationDepth"]["n"],
            ],
        ],
    )

    for key, filename in (
        ("citationFunction", "table_s3_function_performance.csv"),
        ("citationDepth", "table_s4_depth_performance.csv"),
    ):
        item = model[key]
        metric_rows = [
            [
                display_label(row["category"]),
                row["precision"],
                row["recall"],
                row["f1"],
                row["support"],
            ]
            for row in item["categoryMetrics"]
        ]
        metric_rows.extend(
            [
                [
                    "Macro average",
                    item["macroAverage"]["precision"],
                    item["macroAverage"]["recall"],
                    item["macroAverage"]["f1"],
                    item["n"],
                ],
                [
                    "Weighted average",
                    item["weightedAverage"]["precision"],
                    item["weightedAverage"]["recall"],
                    item["weightedAverage"]["f1"],
                    item["n"],
                ],
            ]
        )
        write_csv(
            output_dir / filename,
            ["Category", "Precision", "Recall", "F1-score", "Support"],
            metric_rows,
        )

    function_headers, function_rows = confusion_rows(model["citationFunction"])
    depth_headers, depth_rows = confusion_rows(model["citationDepth"])
    write_csv(output_dir / "table_s5_function_confusion.csv", function_headers, function_rows)
    write_csv(output_dir / "table_s6_depth_confusion.csv", depth_headers, depth_rows)

    fdr = {row["dimension"]: row["adjustedP"] for row in results["fdr"]}
    aggregate_rows: list[list[object]] = []
    for dimension in DIMENSION_ORDER:
        item = results["aggregate"][dimension]
        if dimension == "Semantic similarity":
            aggregate_rows.append(
                [
                    dimension,
                    "Mann-Whitney U",
                    item["u1"],
                    "",
                    item["p"],
                    "",
                    item["rankBiserial"],
                    "Rank-biserial correlation",
                    item["n"],
                    fdr[dimension],
                ]
            )
        else:
            test = item["test"]
            aggregate_rows.append(
                [
                    dimension,
                    "Pearson chi-square",
                    test["statistic"],
                    test["df"],
                    test["p"],
                    test["monteCarlo"]["p"],
                    test["cramersV"],
                    "Cramer's V",
                    test["n"],
                    fdr[dimension],
                ]
            )
    write_csv(
        output_dir / "table_s7_aggregate_tests.csv",
        [
            "Dimension",
            "Test",
            "Statistic",
            "df",
            "Asymptotic p",
            "Monte Carlo p",
            "Effect size",
            "Effect-size measure",
            "n",
            "BH-adjusted p",
        ],
        aggregate_rows,
    )

    discipline_test_rows: list[list[object]] = []
    for field in FIELD_NAMES:
        field_result = results["discipline"][field]
        for dimension in DIMENSION_ORDER:
            item = field_result["dimensions"][dimension]
            if dimension == "Semantic similarity":
                discipline_test_rows.append(
                    [
                        field_result["fieldName"],
                        dimension,
                        "Mann-Whitney U",
                        item["u1"],
                        "",
                        item["p"],
                        "",
                        item["rankBiserial"],
                        "Rank-biserial correlation",
                        item["n"],
                    ]
                )
            else:
                test = item["test"]
                discipline_test_rows.append(
                    [
                        field_result["fieldName"],
                        dimension,
                        "Pearson chi-square",
                        test["statistic"],
                        test["df"],
                        test["p"],
                        test["monteCarlo"]["p"],
                        test["cramersV"],
                        "Cramer's V",
                        test["n"],
                    ]
                )
    write_csv(
        output_dir / "table_s8_discipline_tests.csv",
        [
            "JCR category",
            "Dimension",
            "Test",
            "Statistic",
            "df",
            "Asymptotic p",
            "Monte Carlo p",
            "Effect size",
            "Effect-size measure",
            "n",
        ],
        discipline_test_rows,
    )

    discipline_headers = [
        "Dimension",
        "Category",
        "G1 count",
        "G1 proportion",
        "G1 SD",
        "G2 count",
        "G2 proportion",
        "G2 SD",
    ]
    for field in FIELD_NAMES:
        rows = discipline_source_rows(results["discipline"][field])
        table_number = int(field) + 8
        write_csv(
            output_dir / f"table_s{table_number}_discipline_counts.csv",
            discipline_headers,
            rows,
        )
        write_csv(
            output_dir / f"figure_s{field}_source_data.csv",
            discipline_headers,
            rows,
        )

    write_csv(
        output_dir / "table_s14_bootstrap.csv",
        [
            "Dimension",
            "Category",
            "Difference G1-G2",
            "Article-level CI lower",
            "Article-level CI upper",
            "Article-level significant",
            "Journal-level CI lower",
            "Journal-level CI upper",
            "Journal-level significant",
        ],
        [
            [
                row["dimension"],
                display_label(row["category"]),
                row["differenceG1MinusG2"],
                row["articleCI"][0],
                row["articleCI"][1],
                row["articleSignificant"],
                row["journalCI"][0],
                row["journalCI"][1],
                row["journalSignificant"],
            ]
            for row in results["bootstrap"]
        ],
    )
    write_csv(
        output_dir / "table_s15_leave_one_journal_out.csv",
        [
            "Dimension",
            "Excluded code",
            "Excluded journal",
            "n",
            "Chi-square",
            "df",
            "Asymptotic p",
            "Monte Carlo p",
            "Cramer's V",
        ],
        [
            [
                row["dimension"],
                row["excludedLabel"],
                row["excludedJournal"],
                row["n"],
                row["statistic"],
                row["df"],
                row["p"],
                row["monteCarlo"]["p"],
                row["cramersV"],
            ]
            for row in results["leaveOneJournalOut"]
        ],
    )
    write_csv(
        output_dir / "table_s16_bh_correction.csv",
        ["Dimension", "Hypothesis", "Raw p", "BH-adjusted p", "Significant"],
        [
            [
                row["dimension"],
                HYPOTHESES[row["dimension"]],
                row["p"],
                row["adjustedP"],
                row["significant"],
            ]
            for row in results["fdr"]
        ],
    )
    print(f"Exported report-facing CSV files to {output_dir.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
