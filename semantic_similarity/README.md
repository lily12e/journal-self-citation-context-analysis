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
