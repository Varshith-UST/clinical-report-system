# AI Clinical Report Summarization Assistant

A local, synthetic-data capstone demonstrating safe PDF ingestion, source-linked
report analysis, deterministic summary assembly, guideline retrieval, draft
review suggestions, and exact-draft human review.

This is a software-training demo, not medical advice or a medical device. Use
synthetic reports and synthetic guidance only.

## What is implemented

The project has exactly three AI agents:

1. Report Analysis Agent extracts measurements and bounded narrative facts,
   validates page evidence, and performs supported numeric comparisons.
2. Summary Agent selects validated finding IDs; deterministic code renders facts,
   repeat groups, uncertainties, and evidence.
3. Recommendation Agent uses bounded local guideline passages to draft
   source-linked review suggestions or abstain.

CrewAI Flow resolves analysis first, runs Summary Agent work and local guideline
retrieval concurrently, joins them, verifies the guideline fingerprint, and then
runs or reuses Recommendation Agent output. Parallelism is one AI branch plus
one local retrieval branch, not multiple AI agents running in parallel.

## Prerequisites

Windows 11 is the primary tested setup:

- Git
- uv
- internet for the first dependency sync, first embedding-model download, and
  any deliberately requested Gemini call

The project supports Python 3.12 only. You do not need to install Python
separately when uv can download its managed Python 3.12 interpreter.

### Install Git and uv on Windows

Open PowerShell and install Git with Windows Package Manager:

    winget install --id Git.Git -e --source winget

If `winget` is unavailable, use the official Git for Windows installer:
https://git-scm.com/install/windows.

Install uv with Astral's official Windows installer:

    powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

Alternatively, install uv with Windows Package Manager:

    winget install --id=astral-sh.uv -e

Close and reopen PowerShell after installation so its refreshed `PATH` includes
the new commands. The official uv installation page also lists other supported
methods: https://docs.astral.sh/uv/getting-started/installation/.

Verify the tools:

    git --version
    uv --version

## Clone and install

    git clone https://github.com/Varshith-UST/clinical-report-system.git
    cd clinical-report-system
    uv sync --locked
    uv run python --version

pyproject.toml and uv.lock are the canonical dependency sources. The repository
does not maintain requirements.txt because a second hand-written dependency
list can drift from the lockfile.

## Configure Gemini

Create an ignored local file from the placeholder template:

    Copy-Item .env.example .env

Edit .env and set both values explicitly:

    GEMINI_API_KEY=your_key_here
    GEMINI_MODEL=your_model_id_here

Do not put a real key in .env.example. Configuration precedence is process
environment, Streamlit secrets, then local .env. The app does not guess model
availability or free quota, enable billing, or switch models automatically.
Gemini calls consume the quota associated with your account and selected model.

## Generate synthetic demo PDFs

    uv run python scripts/generate_synthetic_report.py
    uv run python scripts/generate_synthetic_guidelines.py

The generated directory is ignored by Git. Generated files are explicitly
labeled synthetic and contain no real patient data.

## Run the app

    uv run streamlit run src/clinical_report_assistant/app.py

Open the local URL printed by Streamlit, normally http://localhost:8501.

### Index the synthetic guidance

In Guideline Knowledge Base:

1. Upload generated/synthetic_guidelines/synthetic_capstone_review_guidance.pdf.
2. Select Index guideline.
3. Wait for extraction, local embedding, and Chroma indexing.
4. Optionally enter a retrieval query to inspect page-linked passages.

The first real embedding run downloads
sentence-transformers/all-MiniLM-L6-v2. After it is cached, embedding and
retrieval run locally and use no Gemini quota. Guideline data persists under
data/chroma_guidelines for this clone only.

### Run the report workflow

1. Upload generated/synthetic_reports/synthetic_lab_report.pdf.
2. Select Run remaining workflow to resolve the pipeline.
3. Review the validated Analysis and Summary sections.
4. Review the Recommendation draft or abstention.
5. Enter review notes and explicitly approve or reject the exact draft.

Nothing calls Gemini automatically on upload, rerun, or restart. The original
Analyze Report, Generate Summary, Generate Recommendations, and regeneration
buttons remain available for stage-by-stage testing. Editing a draft creates a
new pending version and never inherits approval.

Compatible validated results are reused from
data/local_results/results.sqlite3 without Gemini calls. The database contains
findings and source excerpts, which are report data. It stores neither original
PDF bytes nor API keys. Local storage is not a shared-user isolation boundary.

## Offline verification

Run the complete offline suite and static checks:

    uv run pytest
    uv run ruff check .
    uv lock --check
    uv sync --check

Run the mocked end-to-end demo:

    uv run python scripts/run_offline_workflow_demo.py

The demo clearly labels its agent boundaries mocked. It uses synthetic PDFs,
deterministic fake embeddings, temporary Chroma, and temporary SQLite. It needs
no API key, makes no network calls, and does not prove live Gemini behavior.

The optional real-embedding integration test can trigger the first model
download:

    uv run pytest tests/test_guideline_integration.py -v

The explicit one-request synthetic Gemini smoke test is:

    uv run gemini-smoke --live

Run it only when you intend to consume API quota. It never sends uploaded PDFs.

## Troubleshooting

- uv or Git not found: reopen the terminal after installation or invoke the
  existing executable by its full path. This project does not alter permanent
  PATH automatically.
- Dependency drift: run uv sync --locked. If it reports a stale lock, do not
  bypass it; update pyproject.toml and uv.lock together in a deliberate change.
- Windows cloud/hardlink error 396: retry from a clean environment with copy
  mode. If the first attempt created a partial .venv, move it aside first:

      Move-Item .venv .venv.failed
      uv sync --locked --link-mode copy

- Streamlit watcher errors on Windows:

      uv run streamlit run src/clinical_report_assistant/app.py --server.fileWatcherType none

- No PDF text: the file may be scanned. OCR is not implemented.
- Empty guideline index: generate and index the synthetic guidance before
  expecting source-linked suggestions.
- First embedding run is slow: model download is separate from Gemini and does
  not consume Gemini quota.
- Gemini authentication, quota, unavailable-model, timeout, blocked, and
  incomplete-output errors are sanitized. Confirm the local key, explicit model,
  account access, and quota without posting secret values.

## macOS and Linux

These commands are expected but are not verified in this Windows-focused
handoff:

    git clone https://github.com/Varshith-UST/clinical-report-system.git
    cd clinical-report-system
    uv sync --locked
    cp .env.example .env
    uv run streamlit run src/clinical_report_assistant/app.py

## More documentation

- [Architecture and editable workflow diagram](docs/architecture.md)
- [Development, offline demo, and fresh-clone verification](docs/development.md)
- [Data safety, compatibility, and limitations](docs/data-safety-and-limitations.md)
- [Project module rules](AGENTS.md)
