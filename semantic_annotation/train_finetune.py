"""Validate fine-tuning JSONL files and optionally submit an OpenAI job."""

from __future__ import annotations

import argparse
import json
import os
import time
from pathlib import Path
from typing import Optional

try:
    from openai import OpenAI
except ModuleNotFoundError:
    OpenAI = None


def validate_jsonl(path: Path, label: str) -> int:
    """Validate a chat-format fine-tuning JSONL file and return its record count."""
    if not path.is_file():
        raise FileNotFoundError(f"{label} file does not exist: {path}")
    if path.suffix.casefold() != ".jsonl":
        raise ValueError(f"{label} file must use the .jsonl extension: {path}")

    count = 0
    with path.open("r", encoding="utf-8") as handle:
        for line_number, raw_line in enumerate(handle, start=1):
            if not raw_line.strip():
                raise ValueError(
                    f"{label} file contains a blank line at line {line_number}: {path}"
                )
            try:
                record = json.loads(raw_line)
            except json.JSONDecodeError as exc:
                raise ValueError(
                    f"{label} file contains invalid JSON at line {line_number}: {exc}"
                ) from exc

            messages = record.get("messages")
            if not isinstance(messages, list) or not messages:
                raise ValueError(
                    f"{label} line {line_number} must contain a non-empty "
                    "'messages' list."
                )
            for message_index, message in enumerate(messages, start=1):
                if not isinstance(message, dict):
                    raise ValueError(
                        f"{label} line {line_number}, message {message_index} "
                        "must be an object."
                    )
                if message.get("role") not in {"system", "user", "assistant"}:
                    raise ValueError(
                        f"{label} line {line_number}, message {message_index} "
                        "has an unsupported role."
                    )
                if not isinstance(message.get("content"), str):
                    raise ValueError(
                        f"{label} line {line_number}, message {message_index} "
                        "must contain string content."
                    )
            count += 1

    if count == 0:
        raise ValueError(f"{label} file contains no records: {path}")
    return count


def parse_batch_size(value: str):
    """Return 'auto' or a positive integer accepted by the fine-tuning API."""
    if value.casefold() == "auto":
        return "auto"
    try:
        parsed = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            "--batch-size must be 'auto' or a positive integer"
        ) from exc
    if parsed < 1:
        raise argparse.ArgumentTypeError(
            "--batch-size must be 'auto' or a positive integer"
        )
    return parsed


def upload_file(client, path: Path, purpose: str = "fine-tune") -> str:
    print(f"Uploading: {path}")
    with path.open("rb") as handle:
        file_object = client.files.create(file=handle, purpose=purpose)
    print(f"Uploaded file ID: {file_object.id}")
    return file_object.id


def create_job(
    client,
    train_file_id: str,
    base_model: str,
    suffix: str,
    n_epochs: int,
    learning_rate_multiplier: float,
    batch_size,
    validation_file_id: Optional[str] = None,
):
    parameters = {
        "model": base_model,
        "training_file": train_file_id,
        "suffix": suffix,
        "hyperparameters": {
            "batch_size": batch_size,
            "learning_rate_multiplier": learning_rate_multiplier,
            "n_epochs": n_epochs,
        },
    }
    if validation_file_id is not None:
        parameters["validation_file"] = validation_file_id

    print("Submitting the fine-tuning job.")
    job = client.fine_tuning.jobs.create(**parameters)
    print(f"Fine-tuning job ID: {job.id}")
    return job


def poll_job(client, job_id: str, interval_seconds: int):
    """Poll until the fine-tuning job succeeds, fails, or is cancelled."""
    printed_event_ids = set()
    while True:
        information = client.fine_tuning.jobs.retrieve(job_id)
        status = information.status
        event_limit = 50 if status in {"succeeded", "failed", "cancelled"} else 20
        events = client.fine_tuning.jobs.list_events(job_id, limit=event_limit)
        for event in reversed(events.data):
            if event.id not in printed_event_ids:
                print(f"[{event.created_at}] {event.type} | {event.message}")
                printed_event_ids.add(event.id)

        if status == "succeeded":
            print(f"Fine-tuning succeeded: {information.fine_tuned_model}")
            return information.fine_tuned_model
        if status in {"failed", "cancelled"}:
            raise RuntimeError(f"Fine-tuning ended with status: {status}")

        time.sleep(interval_seconds)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Validate chat-format JSONL data and optionally upload it to create "
            "an OpenAI fine-tuning job."
        )
    )
    parser.add_argument("--train", type=Path, required=True, help="Training JSONL.")
    parser.add_argument(
        "--validation",
        type=Path,
        help="Optional validation JSONL. Do not use a held-out test set here.",
    )
    parser.add_argument(
        "--base-model",
        help=(
            "Base model or accessible fine-tuned model ID. Required when "
            "submitting a job."
        ),
    )
    parser.add_argument(
        "--suffix",
        default="citation-func-depth",
        help="Fine-tuning job suffix (default: citation-func-depth).",
    )
    parser.add_argument(
        "--n-epochs",
        type=int,
        default=1,
        help="Number of training epochs (default: 1).",
    )
    parser.add_argument(
        "--learning-rate-multiplier",
        type=float,
        default=0.5,
        help="Learning-rate multiplier (default: 0.5).",
    )
    parser.add_argument(
        "--batch-size",
        type=parse_batch_size,
        default="auto",
        help="Batch size: 'auto' or a positive integer (default: auto).",
    )
    parser.add_argument(
        "--poll-seconds",
        type=int,
        default=15,
        help="Status polling interval in seconds (default: 15).",
    )
    parser.add_argument(
        "--model-id-output",
        type=Path,
        help=(
            "Text file for the resulting model ID. Defaults to "
            "<train_directory>/fine_tuned_model.txt."
        ),
    )
    parser.add_argument(
        "--validate-only",
        action="store_true",
        help="Validate local files and parameters without calling the API.",
    )
    parser.add_argument(
        "--yes-submit",
        action="store_true",
        help="Confirm upload and fine-tuning submission, which may incur charges.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Allow replacement of an existing model-ID output file.",
    )
    return parser


def cli() -> None:
    parser = build_parser()
    args = parser.parse_args()

    train_path = args.train.resolve()
    validation_path = args.validation.resolve() if args.validation else None
    model_id_output = (
        args.model_id_output.resolve()
        if args.model_id_output
        else train_path.parent / "fine_tuned_model.txt"
    )

    try:
        training_records = validate_jsonl(train_path, "training")
        validation_records = (
            validate_jsonl(validation_path, "validation")
            if validation_path
            else None
        )
    except (FileNotFoundError, ValueError) as exc:
        parser.error(str(exc))

    if args.n_epochs < 1:
        parser.error("--n-epochs must be at least 1")
    if args.learning_rate_multiplier <= 0:
        parser.error("--learning-rate-multiplier must be greater than 0")
    if args.poll_seconds < 1:
        parser.error("--poll-seconds must be at least 1")
    if model_id_output.suffix.casefold() != ".txt":
        parser.error("--model-id-output must use the .txt extension")

    print(f"Training file: {train_path} ({training_records} records)")
    if validation_path:
        print(
            f"Validation file: {validation_path} "
            f"({validation_records} records)"
        )
    else:
        print("Validation file: not supplied")
    print(f"Base model: {args.base_model or 'not supplied'}")
    print(
        "Hyperparameters: "
        f"n_epochs={args.n_epochs}, "
        f"learning_rate_multiplier={args.learning_rate_multiplier}, "
        f"batch_size={args.batch_size}"
    )
    print(
        "Randomness: the managed fine-tuning job does not receive a "
        "client-side seed from this script."
    )

    if args.validate_only:
        print("Validation passed; the API was not called and no files were uploaded.")
        return

    if not args.yes_submit:
        parser.error(
            "submission is disabled by default. Add --yes-submit to upload "
            "files and create a potentially billable fine-tuning job."
        )
    if not args.base_model:
        parser.error("--base-model is required when submitting a job")
    if OpenAI is None:
        parser.error(
            "the 'openai' package is not installed. Install requirements.txt "
            "before submitting."
        )
    if not os.getenv("OPENAI_API_KEY"):
        parser.error("OPENAI_API_KEY is not set")
    if model_id_output.exists() and not args.overwrite:
        parser.error(
            f"model-ID output already exists: {model_id_output}. "
            "Use --overwrite to replace it."
        )

    client = OpenAI()
    training_file_id = upload_file(client, train_path)
    validation_file_id = (
        upload_file(client, validation_path)
        if validation_path
        else None
    )
    job = create_job(
        client=client,
        train_file_id=training_file_id,
        validation_file_id=validation_file_id,
        base_model=args.base_model,
        suffix=args.suffix,
        n_epochs=args.n_epochs,
        learning_rate_multiplier=args.learning_rate_multiplier,
        batch_size=args.batch_size,
    )
    model_name = poll_job(client, job.id, interval_seconds=args.poll_seconds)

    model_id_output.parent.mkdir(parents=True, exist_ok=True)
    model_id_output.write_text(model_name or "", encoding="utf-8")
    print(f"Model ID written to: {model_id_output}")


if __name__ == "__main__":
    cli()
