"""Filter citation contexts with an explicitly configured OpenAI API call."""

from __future__ import annotations

import argparse
import os
import re
import time
from datetime import datetime
from pathlib import Path
from typing import Sequence

import pandas as pd

try:
    from openai import OpenAI
except ModuleNotFoundError:
    OpenAI = None


DEFAULT_CITATION_COLUMNS = (
    "Citation content",
    "Citation Content",
    "Citation Content(±3)",
    "Citation Content (±3)",
)
DEFAULT_SELF_CITED_COLUMNS = (
    "Self-cited article",
    "Self-cited Article",
    "Self-cited article title",
    "Self-cited Article Title",
)

SYSTEM_PROMPT = """
【Role】
You are an expert in citation analysis and academic text processing with extensive experience in bibliometrics research. Your specialty is extracting precise citation contexts while maintaining strict adherence to linguistic inclusion rules.

【Task Overview】
Extract citation content/context for a given reference from academic text, following specific linguistic rules for including surrounding sentences.

【Key Concept】
Citation content = the citation sentence itself + any qualifying surrounding sentences based on linguistic markers.

【Extraction Rules - Follow These Steps】
STEP 1: Locate and always include the citation sentence containing the target reference mark.

STEP 2: Analyze preceding sentences for inclusion:
- Check if the citation sentence contains ANY of these linguistic elements:
  • Conjunctive adverbs: Use your complete internal knowledge of discourse connectives
    (including but not limited to: however, therefore, moreover, furthermore, nevertheless,
    consequently, similarly, in contrast, in addition, additionally, likewise, thus, hence, etc.)
  • Demonstrative pronouns: this, that, these, those, such
  • Third-person pronouns: it, they, them, their, its, theirs, he, she, him, her, his, hers
  • First-person plural pronouns when referring to research: we, our, us
- IF the citation sentence contains any of the above → include the immediately preceding sentence
- IF that preceding sentence ALSO contains any of the above elements → include ITS preceding sentence too

STEP 3: Analyze following sentences:
- Check the immediately following sentence for:
  • Demonstrative pronouns that refer back to the citation content
  • Third-person pronouns that refer back to the citation content
  • Conjunctive adverbs that continue the discourse thread
  • Any pronoun or discourse marker that creates clear cohesive links back to the citation

- IF present → include the following sentence

【Strict Constraints】
Never rewrite, paraphrase, summarize, or alter the original wording.
Output must be verbatim text from the input.
Each block of output = one citation content (citation sentence + qualifying surrounding sentences).
""".strip()


def normalize_header(value: object) -> str:
    return re.sub(r"\s+", " ", str(value)).strip().casefold()


def detect_column_name(
    df: pd.DataFrame,
    candidates: Sequence[str],
    label: str,
) -> str:
    for candidate in candidates:
        if candidate in df.columns:
            return candidate

    normalized_columns = {
        normalize_header(column): column
        for column in df.columns
    }
    for candidate in candidates:
        normalized = normalize_header(candidate)
        if normalized in normalized_columns:
            return normalized_columns[normalized]

    raise KeyError(
        f"Could not find the {label} column. Expected one of {list(candidates)}; "
        f"found {list(df.columns)}."
    )


def resolve_sheet_name(
    input_xlsx: Path,
    requested_sheet: str,
    citation_candidates: Sequence[str],
    self_cited_candidates: Sequence[str],
) -> str:
    with pd.ExcelFile(input_xlsx, engine="openpyxl") as workbook:
        sheet_names = workbook.sheet_names

        if requested_sheet.casefold() == "auto":
            for sheet_name in sheet_names:
                headers = pd.read_excel(workbook, sheet_name=sheet_name, nrows=0)
                try:
                    detect_column_name(
                        headers,
                        citation_candidates,
                        "citation-content",
                    )
                    detect_column_name(
                        headers,
                        self_cited_candidates,
                        "self-cited-article",
                    )
                    return sheet_name
                except KeyError:
                    continue
            raise KeyError(
                "No worksheet contains compatible citation-content and "
                f"self-cited-article columns. Available sheets: {sheet_names}."
            )

        if requested_sheet.isdigit():
            index = int(requested_sheet)
            if index >= len(sheet_names):
                raise IndexError(
                    f"Worksheet index {index} is out of range. "
                    f"Available sheets: {sheet_names}."
                )
            return sheet_names[index]

        if requested_sheet not in sheet_names:
            raise KeyError(
                f"Worksheet '{requested_sheet}' was not found. "
                f"Available sheets: {sheet_names}."
            )
        return requested_sheet


def validate_workbook(
    input_xlsx: Path,
    requested_sheet: str,
    citation_candidates: Sequence[str],
    self_cited_candidates: Sequence[str],
) -> tuple[pd.DataFrame, str, str, str]:
    if not input_xlsx.is_file():
        raise FileNotFoundError(f"Input workbook does not exist: {input_xlsx}")
    if input_xlsx.suffix.casefold() != ".xlsx":
        raise ValueError(f"Input must be an .xlsx file: {input_xlsx}")

    sheet_name = resolve_sheet_name(
        input_xlsx,
        requested_sheet,
        citation_candidates,
        self_cited_candidates,
    )
    df = pd.read_excel(input_xlsx, sheet_name=sheet_name, engine="openpyxl")
    citation_column = detect_column_name(
        df,
        citation_candidates,
        "citation-content",
    )
    self_cited_column = detect_column_name(
        df,
        self_cited_candidates,
        "self-cited-article",
    )
    return df, sheet_name, citation_column, self_cited_column


def build_user_prompt(self_cited_article: object, citation_context: object) -> str:
    article_text = "" if pd.isna(self_cited_article) else str(self_cited_article)
    context_text = "" if pd.isna(citation_context) else str(citation_context)
    return (
        f"Self-cited article title and authors: {article_text}\n"
        f"Citation context (given text segment): {context_text}\n"
    )


def append_error_log(path: Path, user_prompt: str, error: Exception) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as log:
        log.write("=" * 60 + "\n")
        log.write(f"API failure time: {datetime.now().isoformat()}\n")
        log.write(f"Prompt:\n{user_prompt}\n")
        log.write(f"Error: {error}\n")
        log.write("=" * 60 + "\n\n")


def call_model(
    client,
    model: str,
    user_prompt: str,
    fallback_text: str,
    temperature: float,
    max_tokens: int,
    retries: int,
    error_log: Path | None,
) -> tuple[str, int, int, bool]:
    """Return text, input tokens, output tokens, and whether fallback was used."""
    for attempt in range(1, retries + 1):
        try:
            response = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": user_prompt},
                ],
                temperature=temperature,
                max_tokens=max_tokens,
            )
            content = (response.choices[0].message.content or "").strip()
            usage = getattr(response, "usage", None)
            input_tokens = getattr(usage, "prompt_tokens", 0) or 0
            output_tokens = getattr(usage, "completion_tokens", 0) or 0
            if content:
                return content, input_tokens, output_tokens, False
            return fallback_text, input_tokens, output_tokens, True
        except Exception as exc:
            if attempt == retries:
                if error_log is not None:
                    append_error_log(error_log, user_prompt, exc)
                return fallback_text, 0, 0, True
            time.sleep(min(2 ** (attempt - 1), 10))

    return fallback_text, 0, 0, True


def run_filtering(
    df: pd.DataFrame,
    citation_column: str,
    self_cited_column: str,
    client,
    model: str,
    temperature: float,
    max_tokens: int,
    retries: int,
    error_log: Path | None,
) -> tuple[pd.DataFrame, dict[str, int]]:
    filtered_values = []
    input_tokens = 0
    output_tokens = 0
    fallback_count = 0

    for processed, (_, row) in enumerate(df.iterrows(), start=1):
        citation_text = (
            ""
            if pd.isna(row[citation_column])
            else str(row[citation_column])
        )
        prompt = build_user_prompt(
            row[self_cited_column],
            citation_text,
        )
        filtered, prompt_tokens, completion_tokens, used_fallback = call_model(
            client=client,
            model=model,
            user_prompt=prompt,
            fallback_text=citation_text,
            temperature=temperature,
            max_tokens=max_tokens,
            retries=retries,
            error_log=error_log,
        )
        filtered_values.append(filtered)
        input_tokens += prompt_tokens
        output_tokens += completion_tokens
        fallback_count += int(used_fallback)

        if processed % 25 == 0 or processed == len(df):
            print(f"Processed {processed}/{len(df)} records.")

    result = df.copy()
    output_column = "citation content (filtered)"
    if output_column in result.columns:
        result = result.drop(columns=[output_column])
    citation_index = result.columns.get_loc(citation_column)
    result.insert(citation_index + 1, output_column, filtered_values)
    statistics = {
        "records": len(result),
        "input_tokens": input_tokens,
        "output_tokens": output_tokens,
        "fallback_records": fallback_count,
    }
    return result, statistics


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Validate an annotation workbook and optionally filter citation "
            "contexts through the OpenAI Chat Completions API."
        )
    )
    parser.add_argument("--input", type=Path, required=True, help="Input .xlsx file.")
    parser.add_argument(
        "--output",
        type=Path,
        help="Output .xlsx file. Required unless --validate-only is used.",
    )
    parser.add_argument(
        "--sheet",
        default="auto",
        help=(
            "Worksheet name or zero-based index. The default, 'auto', selects "
            "the first compatible worksheet."
        ),
    )
    parser.add_argument(
        "--model",
        help="Accessible OpenAI model ID. Required when running API filtering.",
    )
    parser.add_argument(
        "--citation-column",
        help="Exact citation-content column name, if automatic detection is unsuitable.",
    )
    parser.add_argument(
        "--self-cited-column",
        help="Exact self-cited-article column name, if automatic detection is unsuitable.",
    )
    parser.add_argument(
        "--temperature",
        type=float,
        default=0.3,
        help="Sampling temperature used in the original script (default: 0.3).",
    )
    parser.add_argument(
        "--max-tokens",
        type=int,
        default=400,
        help="Maximum output tokens per record (default: 400).",
    )
    parser.add_argument(
        "--retries",
        type=int,
        default=2,
        help="API attempts per record before using the input as fallback (default: 2).",
    )
    parser.add_argument(
        "--error-log",
        type=Path,
        help=(
            "Optional error-log path. Prompts may contain article text, so "
            "review the file before sharing it."
        ),
    )
    parser.add_argument(
        "--validate-only",
        action="store_true",
        help="Validate the workbook and columns without calling the API.",
    )
    parser.add_argument(
        "--yes-run",
        action="store_true",
        help="Confirm API processing, which may incur charges.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Allow replacement of an existing output file.",
    )
    return parser


def cli() -> None:
    parser = build_parser()
    args = parser.parse_args()

    citation_candidates = (
        (args.citation_column,)
        if args.citation_column
        else DEFAULT_CITATION_COLUMNS
    )
    self_cited_candidates = (
        (args.self_cited_column,)
        if args.self_cited_column
        else DEFAULT_SELF_CITED_COLUMNS
    )

    input_xlsx = args.input.resolve()
    try:
        df, sheet_name, citation_column, self_cited_column = validate_workbook(
            input_xlsx,
            requested_sheet=args.sheet,
            citation_candidates=citation_candidates,
            self_cited_candidates=self_cited_candidates,
        )
    except (FileNotFoundError, ValueError, KeyError, IndexError) as exc:
        parser.error(str(exc))
    if not 0 <= args.temperature <= 2:
        parser.error("--temperature must be between 0 and 2")
    if args.max_tokens < 1:
        parser.error("--max-tokens must be at least 1")
    if args.retries < 1:
        parser.error("--retries must be at least 1")

    print(f"Input workbook: {input_xlsx}")
    print(f"Worksheet: {sheet_name}")
    print(f"Rows: {len(df)}")
    print(f"Citation column: {citation_column}")
    print(f"Self-cited article column: {self_cited_column}")
    print(f"Model: {args.model or 'not supplied'}")
    print(f"Temperature: {args.temperature}")
    print(
        "Randomness: no client-side seed is fixed; API output may vary "
        "between runs."
    )

    if args.validate_only:
        print("Validation passed; the API was not called and no output was written.")
        return

    if args.output is None:
        parser.error("--output is required unless --validate-only is used")
    output_xlsx = args.output.resolve()
    if output_xlsx.suffix.casefold() != ".xlsx":
        parser.error("--output must use the .xlsx extension")
    if output_xlsx == input_xlsx:
        parser.error("the output path must differ from the input path")
    if output_xlsx.exists() and not args.overwrite:
        parser.error(
            f"output already exists: {output_xlsx}. "
            "Use --overwrite to replace it."
        )
    if not args.yes_run:
        parser.error(
            "API processing is disabled by default. Add --yes-run to confirm "
            "a potentially billable run."
        )
    if not args.model:
        parser.error("--model is required when running API filtering")
    if OpenAI is None:
        parser.error(
            "the 'openai' package is not installed. Install requirements.txt "
            "before running API filtering."
        )
    if not os.getenv("OPENAI_API_KEY"):
        parser.error("OPENAI_API_KEY is not set")

    error_log = args.error_log.resolve() if args.error_log else None
    client = OpenAI()
    result, statistics = run_filtering(
        df=df,
        citation_column=citation_column,
        self_cited_column=self_cited_column,
        client=client,
        model=args.model,
        temperature=args.temperature,
        max_tokens=args.max_tokens,
        retries=args.retries,
        error_log=error_log,
    )

    output_xlsx.parent.mkdir(parents=True, exist_ok=True)
    result.to_excel(
        output_xlsx,
        index=False,
        sheet_name="Filtered Results",
        engine="openpyxl",
    )
    print(f"Output written to: {output_xlsx}")
    print(f"Input tokens: {statistics['input_tokens']}")
    print(f"Output tokens: {statistics['output_tokens']}")
    print(f"Fallback records: {statistics['fallback_records']}")
    print(
        "Cost is not estimated because API pricing is time-dependent; "
        "check the provider's current billing records."
    )


if __name__ == "__main__":
    cli()
