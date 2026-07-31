# Statistical analysis and report outputs

This directory contains the deterministic analysis workflow used for the
main-text and supplementary statistical tables and data figures.

## Run from the repository root

```text
python analysis/prepare_analysis_data.py --force
python analysis/evaluate_annotation.py --force
python analysis/run_full_analysis.py
python analysis/export_report_tables.py
```

The workflow does not call an external API.

## Files

- `prepare_analysis_data.py` reads `data/full_annotation_results.xlsx`,
  worksheet `Table S2`, verifies 2,435 citation mentions and 1,883 Citation
  Strength occurrences, and writes `results/analysis_records.json`.
- `evaluate_annotation.py` recalculates Cohen's kappa from
  `data/inter_annotator_agreement.xlsx` and held-out model metrics from
  `semantic_annotation/training_data/test_set_84.xlsx`.
- `run_full_analysis.py` performs aggregate and discipline-level tests,
  fixed-margin Monte Carlo tests, hierarchical bootstrap analyses,
  leave-one-journal-out analyses, and Benjamini-Hochberg correction.
- `export_report_tables.py` converts the machine-readable JSON results into
  one CSV file per reported statistical table or data figure.

## Deterministic settings

- Base seed: `20260726`
- Aggregate Monte Carlo iterations: `100000`
- Discipline-level Monte Carlo iterations: `50000`
- Hierarchical bootstrap iterations: `2000`

These values are the command defaults and are recorded in
`results/analysis_results.json`.

## Citation Function category accounting

The complete conceptual Citation Function taxonomy contains 15 categories.
The operational model label set and aggregate dataset contain 14 categories;
`Irrelevant citation` is not a GPT-4o output option and has zero observations in
both aggregate comparison groups. It is therefore excluded before the
chi-square test, which uses 14 nonzero categories and 13 degrees of freedom.
Separately, the 84-instance held-out model-evaluation set represents 10 of the
14 model-output categories. The evaluation JSON records the conceptual,
model-output, and represented-label counts separately.

## Figure source files

The CSV files named `figure*_source_data.csv` contain the numerical values
used in the data figures. The editable GraphPad Prism project is released as
`figure_sources/data_figures.prism`. Conceptual and workflow diagrams were
prepared in Microsoft Visio; their editable source is
`figure_sources/manuscript_diagrams.vsdx`.

The statistical values can be reproduced without Prism or Visio. Those
applications are needed only to reproduce the exact presentation layout of
the submitted figures.
