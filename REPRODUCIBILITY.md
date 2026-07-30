# Reproducibility map for manuscript and supplementary outputs

This map follows common replication-package practice by identifying, for each
reported output, the exact source data, analysis or generation file, and
report-facing output. Conceptual and illustrative items that are not generated
from statistical data are explicitly marked as such.

All paths are relative to the repository root. Statistical commands are
documented in `analysis/README.md`.

| Reported output | Exact source data | Analysis, generation file, or procedure | Report-facing output / editable source |
|---|---|---|---|
| Main Table 1. Six-dimensional annotation framework | `semantic_annotation/annotation_guidelines.docx` | Not data-derived; definitions were consolidated from the released annotation guidelines and Methods | Manuscript table |
| Main Table 2. Basic statistics of sample data | `data/sample_selection_counts.csv` | `analysis/export_report_tables.py` | `analysis/report_tables/main_table2_sample_statistics.csv` |
| Main Table 3. Example of the training corpus | `data/inter_annotator_agreement.xlsx`, `'2022 Annotation Comparison'!A2:N4` | Illustrative excerpt; no statistical transformation | Manuscript table |
| Main Table 4. Annotated self-citation characteristics | `data/full_annotation_results.xlsx`, `Table S2!A1:M2436` | `analysis/prepare_analysis_data.py`; `analysis/run_full_analysis.py`; `analysis/export_report_tables.py` | `analysis/report_tables/main_table4_aggregate_counts.csv` |
| Main Table 5. Significant differences by JCR category | `data/full_annotation_results.xlsx`, `Table S2!A1:M2436` | `analysis/prepare_analysis_data.py`; `analysis/run_full_analysis.py`; `analysis/export_report_tables.py` | `analysis/report_tables/main_table5_significance_summary.csv` |
| Main Figure 1. Six-dimensional conceptual framework | `semantic_annotation/annotation_guidelines.docx` | Not data-derived; editable diagram prepared in Microsoft Visio | `figure_sources/manuscript_diagrams.vsdx` |
| Main Figure 2. Dataset construction and citation extraction workflow | `data/sample_selection_counts.csv` | Flow diagram assembled in Microsoft Visio from the released selection counts | `figure_sources/manuscript_diagrams.vsdx` |
| Main Figure 3. Semi-automated annotation workflow | Released scripts under `citation_extraction_and_structural_annotation/`, `semantic_annotation/`, and `semantic_similarity/` | Not statistically generated; workflow diagram prepared in Microsoft Visio | `figure_sources/manuscript_diagrams.vsdx` |
| Main Figure 4. Prompt structure | `semantic_annotation/prompts/promptsprompt_long.txt`; `semantic_annotation/prompts/promptsprompt_short.txt` | Prompt components arranged in Microsoft Visio | `figure_sources/manuscript_diagrams.vsdx` |
| Main Figure 5. Citation-sentence filtering process | `semantic_similarity/citation_content_filtering.py` and its embedded prompt | Process diagram prepared in Microsoft Visio | `figure_sources/manuscript_diagrams.vsdx` |
| Main Figure 6. Citation characteristics and semantic similarity | `data/full_annotation_results.xlsx`, `Table S2!A1:M2436` | `analysis/prepare_analysis_data.py`; `analysis/run_full_analysis.py`; `analysis/export_report_tables.py`; GraphPad Prism for layout | `analysis/report_tables/figure6_source_data.csv`; `figure_sources/data_figures.prism` |
| Table S1. Data collection and exclusion statistics | `data/sample_selection_counts.csv` | `analysis/export_report_tables.py` | `analysis/report_tables/table_s1_sample_selection.csv` |
| Table S2. Annotation-pipeline validation | `data/inter_annotator_agreement.xlsx`, `'2022 Annotation Comparison'!K2:N40`; `semantic_annotation/training_data/test_set_84.xlsx`, `Sheet1!E2:H85` | `analysis/evaluate_annotation.py`; `analysis/export_report_tables.py` | `analysis/report_tables/table_s2_validation_summary.csv` |
| Table S3. Citation Function classification performance | `semantic_annotation/training_data/test_set_84.xlsx`, `Sheet1!E2:F85` | `analysis/evaluate_annotation.py`; `analysis/export_report_tables.py` | `analysis/report_tables/table_s3_function_performance.csv` |
| Table S4. Citation Depth classification performance | `semantic_annotation/training_data/test_set_84.xlsx`, `Sheet1!G2:H85` | `analysis/evaluate_annotation.py`; `analysis/export_report_tables.py` | `analysis/report_tables/table_s4_depth_performance.csv` |
| Table S5. Citation Function confusion matrix | `semantic_annotation/training_data/test_set_84.xlsx`, `Sheet1!E2:F85` | `analysis/evaluate_annotation.py`; `analysis/export_report_tables.py` | `analysis/report_tables/table_s5_function_confusion.csv` |
| Table S6. Citation Depth confusion matrix | `semantic_annotation/training_data/test_set_84.xlsx`, `Sheet1!G2:H85` | `analysis/evaluate_annotation.py`; `analysis/export_report_tables.py` | `analysis/report_tables/table_s6_depth_confusion.csv` |
| Table S7. Aggregate-level statistical comparisons | `data/full_annotation_results.xlsx`, `Table S2!A1:M2436` | `analysis/prepare_analysis_data.py`; `analysis/run_full_analysis.py`; `analysis/export_report_tables.py` | `analysis/report_tables/table_s7_aggregate_tests.csv` |
| Table S8. Discipline-level statistical comparisons | `data/full_annotation_results.xlsx`, `Table S2!A1:M2436` | `analysis/prepare_analysis_data.py`; `analysis/run_full_analysis.py`; `analysis/export_report_tables.py` | `analysis/report_tables/table_s8_discipline_tests.csv` |
| Table S9. Information Science & Library Science | `data/full_annotation_results.xlsx`, `Table S2!A1:M2436`, labels `G1-1` and `G2-1` | `analysis/prepare_analysis_data.py`; `analysis/run_full_analysis.py`; `analysis/export_report_tables.py` | `analysis/report_tables/table_s9_discipline_counts.csv` |
| Table S10. Gerontology | `data/full_annotation_results.xlsx`, `Table S2!A1:M2436`, labels `G1-2` and `G2-2` | `analysis/prepare_analysis_data.py`; `analysis/run_full_analysis.py`; `analysis/export_report_tables.py` | `analysis/report_tables/table_s10_discipline_counts.csv` |
| Table S11. Engineering, Multidisciplinary | `data/full_annotation_results.xlsx`, `Table S2!A1:M2436`, labels `G1-3` and `G2-3` | `analysis/prepare_analysis_data.py`; `analysis/run_full_analysis.py`; `analysis/export_report_tables.py` | `analysis/report_tables/table_s11_discipline_counts.csv` |
| Table S12. Veterinary Sciences | `data/full_annotation_results.xlsx`, `Table S2!A1:M2436`, labels `G1-4` and `G2-4` | `analysis/prepare_analysis_data.py`; `analysis/run_full_analysis.py`; `analysis/export_report_tables.py` | `analysis/report_tables/table_s12_discipline_counts.csv` |
| Table S13. Geography | `data/full_annotation_results.xlsx`, `Table S2!A1:M2436`, labels `G1-5` and `G2-5` | `analysis/prepare_analysis_data.py`; `analysis/run_full_analysis.py`; `analysis/export_report_tables.py` | `analysis/report_tables/table_s13_discipline_counts.csv` |
| Figure S1. Information Science & Library Science | `data/full_annotation_results.xlsx`, `Table S2!A1:M2436`, labels `G1-1` and `G2-1` | `analysis/prepare_analysis_data.py`; `analysis/run_full_analysis.py`; `analysis/export_report_tables.py`; GraphPad Prism for layout | `analysis/report_tables/figure_s1_source_data.csv`; `figure_sources/data_figures.prism` |
| Figure S2. Gerontology | `data/full_annotation_results.xlsx`, `Table S2!A1:M2436`, labels `G1-2` and `G2-2` | `analysis/prepare_analysis_data.py`; `analysis/run_full_analysis.py`; `analysis/export_report_tables.py`; GraphPad Prism for layout | `analysis/report_tables/figure_s2_source_data.csv`; `figure_sources/data_figures.prism` |
| Figure S3. Engineering, Multidisciplinary | `data/full_annotation_results.xlsx`, `Table S2!A1:M2436`, labels `G1-3` and `G2-3` | `analysis/prepare_analysis_data.py`; `analysis/run_full_analysis.py`; `analysis/export_report_tables.py`; GraphPad Prism for layout | `analysis/report_tables/figure_s3_source_data.csv`; `figure_sources/data_figures.prism` |
| Figure S4. Veterinary Sciences | `data/full_annotation_results.xlsx`, `Table S2!A1:M2436`, labels `G1-4` and `G2-4` | `analysis/prepare_analysis_data.py`; `analysis/run_full_analysis.py`; `analysis/export_report_tables.py`; GraphPad Prism for layout | `analysis/report_tables/figure_s4_source_data.csv`; `figure_sources/data_figures.prism` |
| Figure S5. Geography | `data/full_annotation_results.xlsx`, `Table S2!A1:M2436`, labels `G1-5` and `G2-5` | `analysis/prepare_analysis_data.py`; `analysis/run_full_analysis.py`; `analysis/export_report_tables.py`; GraphPad Prism for layout | `analysis/report_tables/figure_s5_source_data.csv`; `figure_sources/data_figures.prism` |
| Table S14. Hierarchical bootstrap | `data/full_annotation_results.xlsx`, `Table S2!A1:M2436` | `analysis/prepare_analysis_data.py`; `analysis/run_full_analysis.py`; `analysis/export_report_tables.py` | `analysis/report_tables/table_s14_bootstrap.csv` |
| Table S15. Leave-one-journal-out analyses | `data/full_annotation_results.xlsx`, `Table S2!A1:M2436` | `analysis/prepare_analysis_data.py`; `analysis/run_full_analysis.py`; `analysis/export_report_tables.py` | `analysis/report_tables/table_s15_leave_one_journal_out.csv` |
| Table S16. Benjamini-Hochberg correction | `data/full_annotation_results.xlsx`, `Table S2!A1:M2436` | `analysis/prepare_analysis_data.py`; `analysis/run_full_analysis.py`; `analysis/export_report_tables.py` | `analysis/report_tables/table_s16_bh_correction.csv` |

## Notes

- `analysis/results/analysis_records.json` is the deterministic record-oriented
  representation of the released workbook.
- `analysis/results/analysis_results.json` records the exact statistical
  results, random seed, Monte Carlo iterations, and bootstrap iterations.
- `analysis/results/annotation_evaluation.json` records recalculated agreement
  and held-out model-evaluation metrics.
- Prism and Visio are required only for the exact presentation layout. The
  underlying statistical values are available in open CSV and JSON formats.
