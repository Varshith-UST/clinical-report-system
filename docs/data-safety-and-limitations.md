# Data safety and limitations

## Local data

Each clone is a separate local installation. The following repository-relative
locations are ignored by Git:

- generated synthetic PDFs under generated
- guideline Chroma data under data/chroma_guidelines
- validated result records under data/local_results/results.sqlite3
- SQLite sidecars, tool caches, uploads, and output directories

Uploaded PDF bytes remain in memory and are not stored. The local SQLite
database does store validated findings, summaries, draft suggestions, and source
excerpts. Those records contain report data. Local storage is not a shared-user
isolation or access-control boundary.

Secrets are loaded server-side with this precedence:

1. process environment
2. Streamlit secrets
3. local .env

.env.example contains placeholders only. Never commit .env, Streamlit secrets,
API keys, real reports, local databases, or vector stores.

## Validation and reuse

- Reports are identified by SHA-256 of uploaded bytes, not filename.
- Analysis reuse binds report hash, model, prompt and schema versions,
  generation settings, schema hash, and dependency versions.
- Summary reuse additionally binds the exact immutable analysis.
- Recommendation reuse additionally binds the exact summary, retrieval settings,
  guideline document hashes, and full index fingerprint.
- Every saved load is revalidated with Pydantic and upstream evidence checks.
- A changed guideline index invalidates current recommendation reuse.
- Human approval or rejection binds exactly one immutable draft. Editing creates
  a new pending version.

## Limitations

This is a synthetic-data educational capstone, not a medical device. Source
validation proves that cited text exists in the supplied source; it does not
prove clinical accuracy, completeness, guideline authority, or correctness.

PDF extraction supports selectable text only. Scans may require OCR, which is
not implemented. Range parsing is deliberately narrow. Unsupported compound
values, conditional ranges, missing units, unit mismatches, and ambiguous source
association are left for review rather than guessed.

Vector similarity is a retrieval mechanism, not a confidence score. The
Recommendation Agent must abstain when bounded guidance is absent, unrelated,
contradictory, or insufficient. It cannot diagnose, prescribe, change
medications, supply doses, or assign urgency.

Gemini availability, account quota, and configured model access are external.
The project does not guess free quota, enable paid billing, retry automatically
at multiple layers, or switch models automatically.
