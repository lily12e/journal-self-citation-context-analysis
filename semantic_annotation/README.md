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

`train_finetune.py` reads the released JSONL files from `training_data`.
Running it requires the user's own OpenAI API key and access to the model
specified in that script. Fine-tuned model identifiers may be account- or
project-specific.
