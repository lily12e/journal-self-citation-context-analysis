# Journal Self-Citation Context Analysis

Code and data for the paper:

**"Contextual Patterns of Journal Self-Citation in JCR-Suppressed and Comparator Journals: A Multi-Dimensional Citation Content Analysis"**

## Overview

This study proposes a six-dimensional citation content annotation framework—covering citation position, distance, strength, function, depth, and semantic similarity—and applies it to compare self-citation patterns between JCR-suppressed journals and comparator journals across five disciplinary categories.

## Repository structure

```
├── citation_extraction_and_structural_annotation/
│   ├── structural_annotation.py    # Citation extraction, position, distance, and strength annotation
│   └── error_checking.py           # Automated detection of extraction anomalies
│
├── semantic_annotation/
│   ├── build_training_data.py      # Convert annotated data to JSONL for GPT-4o fine-tuning
│   ├── train_finetune.py           # Fine-tuning pipeline for citation function and depth classification
│   ├── annotation_guidelines.md    # Annotation criteria for citation function and depth
│   ├── prompts/
│   │   ├── prompt_long.txt         # Full prompt with category definitions and examples
│   │   └── prompt_short            # Abbreviated prompt for inference
│   └── training_data/
│       ├── batch1_train_130.jsonl   # Stage 1 training set (130 instances)
│       ├── batch1_val_32.jsonl      # Stage 1 validation set (32 instances)
│       ├── batch2_train_55.jsonl    # Stage 2 training set (55 instances)
│       ├── batch2_val_20.jsonl      # Stage 2 validation set (20 instances)
│       ├── batch3_train_53.jsonl    # Stage 3 training set (53 instances)
│       ├── fine_tuned_model_v1.txt  # Model ID after stage 1
│       ├── fine_tuned_model_v2.txt  # Model ID after stage 2
│       ├── fine_tuned_model_v3.txt  # Model ID after stage 3 (final)
│       └── test_set_84.xlsx         # Independent test set (84 instances)
│
├── semantic_similarity/
│   ├── citation_content_filtering.py   # GPT-4o-based citation context filtering
│   └── similarity_computation.py       # BERT-based semantic similarity with whitening
│
└── data/
    └── full_annotation_results.xlsx    # Complete annotation results across six dimensions
```

## Pipeline

1. **Citation extraction and structural annotation** — Extract in-text citations from .docx files, match them to reference lists, and annotate citation position, distance, and strength.

2. **Semantic annotation** — Classify citation function (14 categories) and citation depth (3 levels) using a fine-tuned GPT-4o model.

3. **Semantic similarity** — Filter citation contexts for relevance, then compute cosine similarity between citation content embeddings and cited article abstract embeddings using BERT with whitening transformation.

## Requirements

- Python 3.8+
- python-docx
- pandas
- openpyxl
- openai
- transformers
- torch

Install dependencies:

```bash
pip install python-docx pandas openpyxl openai transformers torch
```

## Usage

Each script is configured with input/output paths at the top of the file. Modify these paths to point to your data before running:

```bash
# Step 1: Extract citations and annotate structural features
python citation_extraction_and_structural_annotation/structural_annotation.py

# Step 2: Fine-tune GPT-4o for function and depth classification
python semantic_annotation/train_finetune.py

# Step 3: Filter citation content
python semantic_similarity/citation_content_filtering.py

# Step 4: Compute semantic similarity
python semantic_similarity/similarity_computation.py
```

## Citation

(To be added upon acceptance)

## License

This project is released for academic and research purposes.
