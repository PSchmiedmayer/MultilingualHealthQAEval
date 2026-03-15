<!--
This source file is part of the multilingual medical QA evaluation project

SPDX-FileCopyrightText: 2026 Stanford University and the project authors (see CONTRIBUTORS.md)

SPDX-License-Identifier: MIT
-->

# Multilingual Medical QA Evaluation

This project is a notebook-based workflow for evaluating multilingual medical questions across LLMs. It uses direct OpenRouter API calls, Excel review files for human raters, and Parquet checkpoints so interrupted runs can resume.

## Files

- `notebooks/medical_multilingual_eval.ipynb`: main workflow
- `src/medical_eval.py`: helper functions for model calls, checkpoints, Excel export, and summaries
- `requirements.txt`: local Python dependencies
- `.env.example`: secrets template
- `LICENSE`, `LICENSES/MIT.txt`, `REUSE.toml`: license and REUSE metadata
- `.github/workflows/lint.yml`: CI checks

## Setup

1. `python3 -m venv .venv`
2. `source .venv/bin/activate`
3. `python -m pip install --upgrade pip`
4. `python -m pip install -r requirements.txt`
5. `cp .env.example .env`
6. Add `OPENROUTER_API_KEY=...` to `.env`
7. `python -m jupyter lab`

## Notebook Flow

1. Define questions, languages, models, and review criteria in the notebook.
2. Generate translation suggestions.
3. Send one translation workbook per language from `review/translations/generated/`.
4. Place completed translation files in `review/translations/completed/`.
5. Generate model answers and review workbooks.
6. Send clinician files from `review/clinician/generated/` and lay files from `review/lay/generated/`.
7. Place completed review files in the matching `completed/` folders.
8. Import the completed reviews and generate summary tables and charts.

## Current Behavior

- Questions default to English source text.
- Translation and answer requests go directly to OpenRouter.
- The default notebook model is `openai/gpt-5-mini` for both translation and answer generation.
- Concurrency is configurable in the notebook with `TRANSLATION_CONCURRENCY` and `EVALUATION_CONCURRENCY`.
- Long-running model cells save progress and resume from saved checkpoints unless you force a restart.
- Progress output includes completed counts and ETA during model execution.

## Outputs

- `inputs/`: saved copies of questions, languages, schemas, and model settings
- `outputs/translation_suggestions.parquet`: saved translation suggestions
- `outputs/raw_answers.parquet`: saved model answers
- `outputs/approved_translations.parquet`: reviewed translations after import
- `outputs/clinician_review_summary.parquet`: clinician summary output
- `outputs/lay_review_summary.parquet`: lay summary output

## Review Files

- Translation review files are split per language and include question context, the suggested translation, and editable review fields.
- Clinician and lay review files include stable `response_key` values, `reviewer_name`, dropdown fields for rubric scoring, and protected sheets with highlighted editable cells.
- Multiple reviewers can review the same response as long as they return separate completed files.

## Repo Metadata

- Licensed under MIT.
- Contributors are listed in `CONTRIBUTORS.md`.
- REUSE metadata is configured in `REUSE.toml`.
- CI runs Ruff, Python compilation, notebook JSON validation, and `reuse lint`.
