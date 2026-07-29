# Semantic similarity analysis

This script reproduces the semantic-similarity values reported in the study
using the specified BERT model and the supplied whitening parameters.

The default `reported` mode uses first/last-layer CLS pooling and the same
tensor layout and normalization procedure used to generate the values in the
released workbook.

## Files

```text
similarity_computation.py
all_utils.py
requirements.txt
resources/bert-base-uncased-first_last_avg-whiten(NLI).pkl
```

## Installation

Create a Python environment and install:

```text
python -m pip install -r requirements.txt
```

The BERT model can be specified as the Hugging Face model ID
`bert-base-uncased` or as a local model directory.

## Optional GPT-based citation-context filtering

`citation_content_filtering.py` is an upstream API-dependent step that applies
the released linguistic prompt to citation contexts. It now reads paths,
worksheet names, column names, and the model ID from command-line arguments.
It recognizes the released `Table S2` columns directly.

Validate the released workbook without an API key or model call:

```text
python semantic_similarity/citation_content_filtering.py \
  --input data/full_annotation_results.xlsx \
  --sheet "Table S2" \
  --validate-only
```

An API run requires the user's own `OPENAI_API_KEY`, model access, billing,
an explicit output path, and the `--yes-run` confirmation:

```text
python semantic_similarity/citation_content_filtering.py \
  --input data/full_annotation_results.xlsx \
  --sheet "Table S2" \
  --output outputs/full_annotation_results_filtered.xlsx \
  --model gpt-4o \
  --temperature 0.3 \
  --max-tokens 400 \
  --yes-run
```

The historical script specified the `gpt-4o` model alias, temperature `0.3`,
and a maximum of 400 output tokens. No client-side seed was fixed, so API
outputs may vary across runs, and model aliases or account access may change.
The released annotation workbook preserves the exact data used for downstream
analysis and can be used without rerunning this API step.

The script reports token usage but does not estimate monetary cost because API
pricing changes over time. Existing outputs are not overwritten unless
`--overwrite` is supplied. An optional error log can contain article text and
should be reviewed before sharing.

## Validate the input without loading BERT

```text
python similarity_computation.py \
  --input ../data/full_annotation_results.xlsx \
  --sheet "Table S2" \
  --validate-only
```

Validation does not load or download the BERT model and does not modify the
workbook.

## Reproduce the reported values

```text
python similarity_computation.py \
  --input ../data/full_annotation_results.xlsx \
  --output ../data/full_annotation_results_reproduced.xlsx \
  --sheet "Table S2" \
  --model bert-base-uncased \
  --mode reported
```

For an offline local model, replace `bert-base-uncased` with the local model
directory and add `--local-files-only`.

Input workbooks are not overwritten unless `--in-place` is explicitly passed.
Existing output files are not replaced unless `--force` is explicitly passed.
