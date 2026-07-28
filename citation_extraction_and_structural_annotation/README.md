# Citation extraction and structural annotation

This directory contains the scripts used to extract and check citation records.
Re-running extraction from source articles requires legally obtained `.docx`
files supplied by the user; the source full text is not redistributed.

## Structural annotation

`structural_annotation.py` extracts target journal self-citations from a
directory of user-supplied `.docx` articles and exports citation position,
distance, strength, sentence, and context fields. Paths and study-specific
target settings are supplied explicitly at run time; the script does not
depend on a local `input` directory or a timestamped output name.

Check the paths and parameters without processing:

```text
python citation_extraction_and_structural_annotation/structural_annotation.py \
  --input-dir path/to/source_articles \
  --output outputs/structural_annotations.xlsx \
  --target-year 2022 \
  --target-journal "Int J Syst Assur Eng Manag" \
  --target-journal "International Journal of System Assurance Engineering and Management" \
  --validate-only
```

Remove `--validate-only` to run the extraction. Repeat `--target-journal` for
abbreviated and full-name variants of the same target journal. Run the command
separately for other target journal/year combinations. Existing output is not
replaced unless `--overwrite` is supplied. The source `.docx` articles are
non-public inputs and must be obtained legally by the reproducing researcher.

## Error checking

`error_checking.py` cleans citation-context whitespace, writes a cleaned-text
column, and highlights records that need manual inspection:

- red: the cleaned citation context is shorter than the selected threshold;
- blue in the citation-content column: more than 10 dash characters;
- blue in the self-cited-article column: more than one distinct publication
  year.

The script accepts a workbook path instead of relying on machine-specific or
placeholder directories. By default, it automatically selects the first
worksheet containing compatible citation-content and self-cited-article
columns.

Validate the released data without writing a file:

```text
python citation_extraction_and_structural_annotation/error_checking.py \
  data/full_annotation_results.xlsx \
  --validate-only
```

Create a checked copy:

```text
python citation_extraction_and_structural_annotation/error_checking.py \
  data/full_annotation_results.xlsx \
  --output outputs/full_annotation_results_checked.xlsx
```

Use `--help` to list optional worksheet, column-name, length-threshold, and
overwrite arguments. Existing output is not replaced unless `--overwrite` is
provided, and the input workbook is never overwritten.
