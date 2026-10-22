# Architecture

This capstone has exactly three AI agents:

1. Report Analysis Agent extracts and source-validates structured report facts.
2. Summary Agent selects bounded finding IDs; deterministic code rebuilds facts.
3. Recommendation Agent drafts source-linked review suggestions or abstains.

PDF ingestion, Gemini configuration, guideline retrieval, SQLite persistence,
session state, and human review are supporting components, not agents.

## Workflow

The diagram is Mermaid text and can be edited directly.

~~~mermaid
flowchart TD
    U[Explicit UI action] --> A{Compatible analysis?}
    A -->|yes| AR[Reuse validated analysis]
    A -->|no| AG[Generate with Report Analysis Agent]
    AG --> AV{Validation passed?}
    AV -->|no| F[failed]
    AV -->|yes| AJ[Persist immutable analysis]
    AR --> FORK
    AJ --> FORK

    FORK{{Bounded two-worker fork}} --> S{Compatible summary?}
    FORK --> R[Local guideline retrieval from analysis]
    S -->|yes| SR[Reuse validated summary]
    S -->|no| SG[Generate with Summary Agent]
    SG --> SV{Validation passed?}
    SV -->|no| F
    SV -->|yes| SJ[Persist immutable summary]
    R --> RF[Capture exact guideline fingerprint and passages]

    SR --> JOIN{{Join}}
    SJ --> JOIN
    RF --> JOIN
    JOIN --> C{Index still identical?}
    C -->|no| F
    C -->|yes| D{Compatible draft?}
    D -->|yes| DR[Reuse exact validated draft]
    D -->|no| DG[Generate with Recommendation Agent]
    DG --> DV{Validation passed?}
    DV -->|no| F
    DV -->|yes| DP[Persist immutable draft]

    DR --> O{Outcome}
    DP --> O
    O -->|insufficient guidance| AB[abstain]
    O -->|new or unreviewed draft| HR[await_review]
    O -->|exact reviewed draft| RV[approved or rejected]
    HR --> H[Explicit human Approve or Reject]
    H --> HP[Persist decision for exact draft version]
~~~

## Execution semantics

- The UI never starts the Flow on upload, rerun, or restart. The user must select
  Run remaining workflow. Existing Analyze, Generate Summary, Generate
  Recommendations, and regeneration controls remain available.
- Analysis resolves first. Summary resolution or generation and local retrieval
  then run concurrently with a maximum of two workers. This is one AI branch
  and one local retrieval branch; it is not parallel execution of multiple AI
  agents.
- Retrieval queries are derived from validated analysis findings and uncertainty
  reasons. Retrieval does not need a generated summary.
- The join requires both branches. Recommendation resolution cannot begin before
  the join and an immediate guideline fingerprint recheck.
- Routes are recorded as reuse, generate, retrieve, abstain, await_review,
  approved, rejected, or failed.
- Each worker creates its own SQLite store wrapper. The retrieval worker creates
  and closes its own Chroma client. Streamlit is not imported by worker code.
- Agent timeouts and output limits remain owned by the three agent modules.
  Retrieval is bounded by query count, top-k, and total passages.
- A failed branch does not erase a successfully saved analysis or summary. A
  failed recommendation does not erase existing valid upstream results.
- Flow execution never approves a draft. Approval and rejection remain explicit,
  immutable, exact-draft UI operations.

## Trust boundaries

Uploaded reports and retrieved passages are untrusted data. Pydantic models,
source-page validation, exact evidence checks, deterministic numeric comparison,
compatibility keys, and guideline fingerprints enforce the boundaries before a
result is displayed or reused.

Guideline retrieval is application-level RAG. It is not an agent-invoked CrewAI
tool, and the project does not claim that it is. Adding agent tools, additional
retries, and monitoring infrastructure is outside this change.

## Main modules

| Module | Responsibility |
| --- | --- |
| src/clinical_report_assistant/pdf_processing.py | In-memory PDF validation and page extraction |
| src/clinical_report_assistant/report_analysis_agent.py | First agent and source validation |
| src/clinical_report_assistant/summary_agent.py | Second agent and deterministic summary assembly |
| src/clinical_report_assistant/guideline_kb.py | Local embeddings, persistent Chroma, retrieval |
| src/clinical_report_assistant/recommendation_agent.py | Third agent, citations, abstention |
| src/clinical_report_assistant/capstone_flow.py | Sequencing, two-branch join, routers, reuse |
| src/clinical_report_assistant/persistence.py | Immutable validated SQLite results and review |
| src/clinical_report_assistant/session_state.py | Session-local stale-result clearing |
| src/clinical_report_assistant/app.py | Explicit controls and presentation |

## Reference documentation

- [CrewAI Flows](https://docs.crewai.com/en/concepts/flows)
- [uv project synchronization](https://docs.astral.sh/uv/concepts/projects/sync/)
- [LangChain Chroma integration](https://docs.langchain.com/oss/python/integrations/vectorstores/chroma)
- [Chroma collections](https://docs.trychroma.com/docs/collections/manage-collections)

