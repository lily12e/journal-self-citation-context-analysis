# Journal Self-Citation Context Analysis

Code and data for the paper:

**"Contextual Patterns of Journal Self-Citation in JCR-Suppressed and
Comparator Journals: A Multi-Dimensional Citation Content Analysis"**

## Overview

This study applies a six-dimensional citation-content framework—citation
position, distance, strength, function, depth, and semantic similarity—to
compare JCR-suppressed and comparator journals across five disciplinary
categories.

## Repository structure

```text
├── citation_extraction_and_structural_annotation/
│   ├── structural_annotation.py
│   └── error_checking.py
├── semantic_annotation/
│   ├── annotation_guidelines.docx
│   ├── build_training_data.py
│   ├── train_finetune.py
│   ├── prompts/
│   │   ├── promptsprompt_long.txt
│   │   └── promptsprompt_short.txt
│   └── training_data/
│       ├── batch1_train_130.jsonl
│       ├── batch1_val_32.jsonl
│       ├── batch2_train_55.jsonl
│       ├── batch2_val_20.jsonl
│       ├── batch3_train_53.jsonl
│       ├── test_set_84.xlsx
│       ├── test_set_84.jsonl
│       └── fine_tuned_model_v1/v2/v3.txt
├── semantic_similarity/
│   ├── citation_content_filtering.py
│   ├── similarity_computation.py
│   ├── all_utils.py
│   ├── requirements.txt
│   └── resources/
│       └── bert-base-uncased-first_last_avg-whiten(NLI).pkl
├── data/
│   ├── full_annotation_results.xlsx
│   ├── citation_extraction_validation.xlsx
│   └── inter_annotator_agreement.xlsx
├── requirements.txt
└── README.md
```

## Environment

Python 3.11 is recommended. The scripts use Python 3.10+ syntax.

Create an isolated environment and install the complete dependency list:

```text
python -m venv .venv
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

PyTorch installation can be platform-specific. The requirement file installs
the standard build; users who need a particular CUDA build should follow the
PyTorch installation instructions for their operating system and hardware.

## Models and access requirements

### BERT-whitening semantic similarity

`semantic_similarity/similarity_computation.py` uses
`bert-base-uncased`. The model can be downloaded by `transformers` or supplied
as a local model directory. Internet access is needed only when the model is
not already available locally. The whitening parameter file is included under
`semantic_similarity/resources/`.

### OpenAI API-dependent steps

The following scripts require the user's own `OPENAI_API_KEY`, an OpenAI API
account with access to the specified model, and available API billing:

```text
semantic_similarity/citation_content_filtering.py
semantic_annotation/train_finetune.py
```

Set the key as an environment variable. Never place an API key in a script,
README, committed `.env` file, or released output.

```text
OPENAI_API_KEY=your_own_key
```

Fine-tuned model identifiers are account- or project-specific and may not be
accessible to other users. The `fine_tuned_model_v1/v2/v3.txt` files document
the model identifiers used in the study; they do not grant access to those
models. Re-running a fine-tuning job requires the user's own API credentials,
model access, and billing.

The released annotation workbook can be used for downstream analyses without
calling the OpenAI API or accessing the fine-tuned models.

## Randomness and deterministic settings

| Step | Randomness setting |
|---|---|
| Structural extraction and error checking | No random sampling in the released scripts |
| Excel-to-JSONL conversion | `--seed 42` by default; the seed controls prompt selection and train/validation shuffling |
| GPT-based citation-content filtering | No client-side seed is fixed; API outputs may vary across reruns |
| Managed fine-tuning | No client-side training seed is specified by the released script |
| BERT-whitening similarity | Model is run in evaluation mode; use the same model files, whitening parameters, package versions, and device for closest numerical reproduction |

The released JSONL files and final annotation workbook preserve the exact data
used for the reported analyses even when an upstream API-dependent step is not
rerun.

## Non-public and user-supplied inputs

The source full-text journal articles are not redistributed in this repository.
Re-running citation extraction from raw documents requires legally obtained
`.docx` article files supplied by the user. The released validation and final
annotation workbooks allow inspection of the extracted and annotated data
without those source documents.

## Pipeline

1. `structural_annotation.py` extracts in-text citations and annotates
   citation position, distance, and strength.
2. `build_training_data.py` converts manually annotated Excel rows to
   API-ready JSONL. For the released test set, `test_set_84.xlsx` is the
   annotated source and `test_set_84.jsonl` is the exact derived file supplied
   to the fine-tuning workflow.
3. `train_finetune.py` submits an optional API-based fine-tuning job for
   citation function and depth.
4. `citation_content_filtering.py` performs optional API-based filtering of
   citation contexts.
5. `similarity_computation.py` computes BERT-whitening semantic similarity.

Subdirectory READMEs provide commands and file-specific notes.

## Citation

To be added upon acceptance.

## License

This project is released for academic and research purposes.
