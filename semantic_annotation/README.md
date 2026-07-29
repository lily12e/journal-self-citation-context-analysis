# Semantic annotation and fine-tuning data

## Build JSONL from an annotated Excel workbook

`build_training_data.py` reads the released column names directly and also
accepts `Self-cited Article Index` as an alias for
`Self-citing Article Index`.

Required columns:

```text
Self-citing Article Index (or Self-cited Article Index)
Self-cited Article Title
Citation Content
citation location
citation function
citation depth
```

The script loads the two system prompts from the repository's `prompts`
directory. It does not contain local absolute paths and does not call an API.

To reconstruct the released 84-record test JSONL while preserving row order
and using the short prompt:

```text
python build_training_data.py --input training_data/test_set_84.xlsx --train-output training_data/test_set_84_rebuilt.jsonl --long-prompt-rate 0 --no-shuffle
```

To create a deterministic training/validation split:

```text
python build_training_data.py --input path/to/annotated_data.xlsx --train-output training_data/train.jsonl --validation-output training_data/validation.jsonl --validation-ratio 0.2 --seed 42
```

Existing outputs are not replaced unless `--force` is explicitly supplied.

## Fine-tuning

`train_finetune.py` validates chat-format JSONL files locally before any API
operation. Paths, the base model, the output file, and hyperparameters are
command-line arguments rather than machine-specific constants.

Validate the historical third-stage files without uploading anything:

```text
python semantic_annotation/train_finetune.py \
  --train semantic_annotation/training_data/batch3_train_53.jsonl \
  --validation semantic_annotation/training_data/test_set_84.jsonl \
  --validate-only
```

The historical script passed `test_set_84.jsonl` to the API as its
`validation_file`; that file name and role are retained here to document the
released workflow. The validation command reports 53 training records and 84
validation records.

Submitting a new job requires the user's own `OPENAI_API_KEY`, API billing,
and access to the selected base model:

```text
python semantic_annotation/train_finetune.py \
  --train semantic_annotation/training_data/batch3_train_53.jsonl \
  --validation semantic_annotation/training_data/test_set_84.jsonl \
  --base-model YOUR_ACCESSIBLE_MODEL_ID \
  --n-epochs 1 \
  --learning-rate-multiplier 0.5 \
  --batch-size auto \
  --yes-submit
```

`--yes-submit` is deliberately required because the command uploads files and
may incur charges. Fine-tuned model identifiers may be account- or
project-specific; the released `fine_tuned_model_v1/v2/v3.txt` files document
the identifiers used in the study but do not grant access.

The managed fine-tuning job does not receive a client-side random seed from
this script. The released JSONL files and recorded model identifiers preserve
the exact study inputs even when another user cannot rerun the account-specific
fine-tuning job.
