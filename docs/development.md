# Development and verification

## Canonical environment

The supported project environment is Python 3.12 and uv. The Python constraint
is declared in pyproject.toml, the preferred interpreter is in .python-version,
and exact resolution is in uv.lock. Do not maintain a second hand-written
requirements file: it can drift from the lockfile. If an external consumer
requires requirements format, generate it from the lock with uv export rather
than treating it as a second dependency source.

Locked setup:

    uv sync --locked
    uv run python --version

uv can download a compatible managed Python interpreter when Python 3.12 is not
already installed. Git must still be installed separately.

On the verified Windows machine, a fresh environment on cloud-backed storage
could not use uv hardlinks and failed with OS error 396. Copy mode succeeded:

    uv sync --locked --link-mode copy

If the failed attempt already created a partial .venv, move that directory aside
before retrying so uv does not trust incomplete package metadata:

    Move-Item .venv .venv.failed
    uv sync --locked --link-mode copy

## Offline checks

    uv run pytest
    uv run ruff check .
    uv lock --check
    uv sync --check

The mocked end-to-end demo needs no key, no model download, and no network:

    uv run python scripts/run_offline_workflow_demo.py

It creates synthetic PDFs in memory and temporary Chroma and SQLite storage. Its
three agent boundaries are mocked, which is printed prominently in the output.
It exercises production validation and persistence but is not evidence of live
Gemini behavior or clinical accuracy.

## Synthetic files

    uv run python scripts/generate_synthetic_report.py
    uv run python scripts/generate_synthetic_guidelines.py

Outputs go under generated and are ignored by Git. They contain synthetic
software-test content only.

The optional real-embedding integration test may download
sentence-transformers/all-MiniLM-L6-v2 on its first run:

    uv run pytest tests/test_guideline_integration.py -v

After the model is cached, embedding and retrieval are local and consume no
Gemini quota. The normal unit suite mocks embeddings and needs no download.

## Live Gemini smoke test

Only run this deliberately:

    uv run gemini-smoke --live

It sends one short built-in synthetic prompt and can consume API quota. It never
sends an uploaded PDF. There is no automatic model switching or billing setup.

## Streamlit health check

Start the app:

    uv run streamlit run src/clinical_report_assistant/app.py

Then open the local URL and verify:

    http://localhost:8501/_stcore/health

Stop the server with Ctrl+C. On Windows file-watcher errors, use:

    uv run streamlit run src/clinical_report_assistant/app.py --server.fileWatcherType none

## Fresh-clone verification

A meaningful fresh-clone check copies only intended tracked files into a new
temporary directory. It must not copy .venv, .env, generated files, local
SQLite, or Chroma data. Run uv sync --locked there, then imports, generators,
the offline suite, the mocked demo, and the Streamlit health check.

uv normally reuses its machine-wide download cache. A successful cached run is
not proof that a completely cold machine has internet access to every registry
or model host. Record the operating system, cache reuse, and any network failure.
