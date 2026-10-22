"""CrewAI Flow orchestration for the three-agent capstone workflow.

The Flow coordinates existing validated modules. It does not define another AI
agent and it never imports Streamlit. Summary generation and local guideline
retrieval are the only concurrent branches.
"""

from __future__ import annotations

from collections.abc import Callable
from concurrent.futures import Future, ThreadPoolExecutor
from enum import StrEnum
from pathlib import Path

from crewai.flow import Flow, listen, router, start
from pydantic import BaseModel, ConfigDict, Field, PrivateAttr

from clinical_report_assistant.guideline_kb import (
    EmptyGuidelineIndexError,
    GuidelineIndexFingerprint,
    GuidelineKnowledgeBaseError,
    RetrievedPassage,
)
from clinical_report_assistant.persistence import (
    AnalysisGenerationSettings,
    AnalysisResolution,
    PersistedResultService,
    PersistenceError,
    RecommendationGenerationSettings,
    RecommendationResolution,
    RuntimeVersions,
    SavedAnalysisRecord,
    SavedReviewDecision,
    SQLiteResultStore,
    SummaryGenerationSettings,
    SummaryResolution,
)
from clinical_report_assistant.recommendation_agent import (
    MAX_RETRIEVED_PASSAGES,
    RecommendationDraftResult,
    build_analysis_retrieval_queries,
    safe_recommendation_error_message,
)
from clinical_report_assistant.report_analysis_agent import (
    ReportAnalysisResult,
    safe_report_analysis_error_message,
)
from clinical_report_assistant.summary_agent import (
    ReportSummaryResult,
    safe_summary_error_message,
)


class WorkflowRoute(StrEnum):
    """Auditable stage decisions and terminal outcomes."""

    REUSE = "reuse"
    GENERATE = "generate"
    RETRIEVE = "retrieve"
    ABSTAIN = "abstain"
    AWAIT_REVIEW = "await_review"
    APPROVED = "approved"
    REJECTED = "rejected"
    FAILED = "failed"


class WorkflowTransition(BaseModel):
    model_config = ConfigDict(frozen=True, extra="forbid")

    stage: str = Field(min_length=1)
    route: WorkflowRoute


class WorkflowFailure(BaseModel):
    """Sanitized failure safe for logs and the UI."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    stage: str
    category: str
    message: str


class GuidelineRetrievalSnapshot(BaseModel):
    """Bounded passages tied to the exact index identity used for retrieval."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    index: GuidelineIndexFingerprint
    queries: list[str]
    passages: list[RetrievedPassage]


class CapstoneFlowState(BaseModel):
    """Structured state retained across sequential and joined Flow stages."""

    model_config = ConfigDict(extra="forbid")

    analysis: AnalysisResolution | None = None
    summary: SummaryResolution | None = None
    retrieval: GuidelineRetrievalSnapshot | None = None
    recommendation: RecommendationResolution | None = None
    review: SavedReviewDecision | None = None
    transitions: list[WorkflowTransition] = Field(default_factory=list)
    failure: WorkflowFailure | None = None
    outcome: WorkflowRoute | None = None


class CapstoneWorkflowResult(BaseModel):
    """Immutable result returned to framework-specific presentation code."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    analysis: AnalysisResolution | None = None
    summary: SummaryResolution | None = None
    retrieval: GuidelineRetrievalSnapshot | None = None
    recommendation: RecommendationResolution | None = None
    review: SavedReviewDecision | None = None
    transitions: list[WorkflowTransition]
    failure: WorkflowFailure | None = None
    outcome: WorkflowRoute


class GuidelineSnapshotError(RuntimeError):
    """Raised when retrieval cannot use one immutable guideline index view."""


AnalysisGenerator = Callable[[], ReportAnalysisResult]
SummaryGenerator = Callable[[ReportAnalysisResult], ReportSummaryResult]
GuidelineRetriever = Callable[[ReportAnalysisResult], GuidelineRetrievalSnapshot]
RecommendationGenerator = Callable[
    [ReportAnalysisResult, ReportSummaryResult, list[RetrievedPassage]],
    RecommendationDraftResult,
]
IndexIdentityProvider = Callable[[], GuidelineIndexFingerprint]


class ClinicalReportFlow(Flow[CapstoneFlowState]):
    """Resolve the existing three-agent pipeline with one parallel join."""

    _database_path: Path = PrivateAttr()
    _report_hash: str = PrivateAttr()
    _analysis_settings: AnalysisGenerationSettings = PrivateAttr()
    _summary_settings: SummaryGenerationSettings = PrivateAttr()
    _recommendation_settings: RecommendationGenerationSettings = PrivateAttr()
    _runtime_versions: RuntimeVersions = PrivateAttr()
    _generate_analysis: AnalysisGenerator = PrivateAttr()
    _generate_summary: SummaryGenerator = PrivateAttr()
    _retrieve_guidelines: GuidelineRetriever = PrivateAttr()
    _generate_recommendation: RecommendationGenerator = PrivateAttr()
    _current_index_identity: IndexIdentityProvider = PrivateAttr()

    def __init__(
        self,
        *,
        database_path: str | Path,
        report_hash: str,
        analysis_settings: AnalysisGenerationSettings,
        summary_settings: SummaryGenerationSettings,
        recommendation_settings: RecommendationGenerationSettings,
        runtime_versions: RuntimeVersions,
        generate_analysis: AnalysisGenerator,
        generate_summary: SummaryGenerator,
        retrieve_guidelines: GuidelineRetriever,
        generate_recommendation: RecommendationGenerator,
        current_index_identity: IndexIdentityProvider,
    ) -> None:
        super().__init__(
            initial_state=CapstoneFlowState(),
            max_method_calls=8,
            tracing=False,
            suppress_flow_events=True,
        )
        self._database_path = Path(database_path)
        self._report_hash = report_hash
        self._analysis_settings = analysis_settings
        self._summary_settings = summary_settings
        self._recommendation_settings = recommendation_settings
        self._runtime_versions = runtime_versions
        self._generate_analysis = generate_analysis
        self._generate_summary = generate_summary
        self._retrieve_guidelines = retrieve_guidelines
        self._generate_recommendation = generate_recommendation
        self._current_index_identity = current_index_identity

    def _service(self) -> PersistedResultService:
        # Workers share no SQLite connection. The store opens one per operation.
        return PersistedResultService(
            SQLiteResultStore(self._database_path),
            runtime_versions=self._runtime_versions,
        )

    def _transition(self, stage: str, route: WorkflowRoute) -> None:
        self.state.transitions.append(WorkflowTransition(stage=stage, route=route))

    def _fail(self, stage: str, category: str, message: str) -> None:
        self.state.failure = WorkflowFailure(
            stage=stage,
            category=category,
            message=message,
        )
        self.state.outcome = WorkflowRoute.FAILED
        self._transition(stage, WorkflowRoute.FAILED)

    @start()
    def resolve_analysis(self) -> None:
        try:
            resolution = self._service().resolve_analysis(
                self._report_hash,
                self._analysis_settings,
                self._generate_analysis,
                force=False,
            )
        except Exception as exc:  # noqa: BLE001 - sanitized workflow boundary
            self._fail(
                "analysis",
                "analysis_resolution",
                safe_report_analysis_error_message(exc),
            )
            return
        self.state.analysis = resolution
        self._transition(
            "analysis",
            WorkflowRoute.REUSE if resolution.reused else WorkflowRoute.GENERATE,
        )

    @router(resolve_analysis, emit=["branches_ready", "failed"])
    def route_after_analysis(self) -> str:
        return "failed" if self.state.failure is not None else "branches_ready"

    def _summary_branch(
        self,
        analysis_record: SavedAnalysisRecord,
    ) -> SummaryResolution:
        return self._service().resolve_summary(
            analysis_record,
            self._summary_settings,
            lambda: self._generate_summary(analysis_record.result),
            force=False,
        )

    @listen("branches_ready")
    def resolve_parallel_branches(self) -> None:
        """Overlap one Summary Agent call with independent local retrieval."""

        analysis_resolution = self.state.analysis
        if analysis_resolution is None:
            self._fail(
                "parallel_join",
                "missing_analysis",
                "No valid analysis is available.",
            )
            return
        analysis_record = analysis_resolution.record
        executor = ThreadPoolExecutor(max_workers=2, thread_name_prefix="capstone-flow")
        summary_future: Future[SummaryResolution] = executor.submit(
            self._summary_branch,
            analysis_record,
        )
        retrieval_future: Future[GuidelineRetrievalSnapshot] = executor.submit(
            self._retrieve_guidelines,
            analysis_record.result,
        )
        summary_error: Exception | None = None
        retrieval_error: Exception | None = None
        try:
            try:
                self.state.summary = summary_future.result()
            except Exception as exc:  # noqa: BLE001 - sanitized below
                summary_error = exc
            try:
                self.state.retrieval = retrieval_future.result()
            except Exception as exc:  # noqa: BLE001 - sanitized below
                retrieval_error = exc
        finally:
            executor.shutdown(wait=True, cancel_futures=True)

        if self.state.summary is not None:
            self._transition(
                "summary",
                WorkflowRoute.REUSE
                if self.state.summary.reused
                else WorkflowRoute.GENERATE,
            )
        if self.state.retrieval is not None:
            self._transition("guideline_retrieval", WorkflowRoute.RETRIEVE)
        if summary_error is not None:
            self._fail(
                "summary",
                "summary_resolution",
                safe_summary_error_message(summary_error),
            )
        elif retrieval_error is not None:
            self._fail(
                "guideline_retrieval",
                "retrieval",
                _safe_retrieval_error(retrieval_error),
            )

    @router(resolve_parallel_branches, emit=["recommendation_ready", "failed"])
    def route_after_parallel_branches(self) -> str:
        if self.state.failure is not None:
            return "failed"
        if self.state.summary is None or self.state.retrieval is None:
            self._fail(
                "parallel_join",
                "incomplete_join",
                "The summary and retrieval branches did not both complete.",
            )
            return "failed"
        return "recommendation_ready"

    @listen("recommendation_ready")
    def resolve_recommendation(self) -> None:
        analysis = self.state.analysis
        summary = self.state.summary
        retrieval = self.state.retrieval
        if analysis is None or summary is None or retrieval is None:
            self._fail(
                "recommendation",
                "missing_dependency",
                "A validated workflow dependency is unavailable.",
            )
            return
        try:
            _validate_recommendation_index(
                retrieval.index,
                self._recommendation_settings,
            )
            if self._current_index_identity() != retrieval.index:
                raise GuidelineSnapshotError(
                    "The guideline index changed after retrieval; "
                    "no draft was generated."
                )
            resolution = self._service().resolve_recommendation(
                analysis.record,
                summary.record,
                self._recommendation_settings,
                lambda: self._generate_recommendation(
                    analysis.record.result,
                    summary.record.result,
                    retrieval.passages,
                ),
                force=False,
            )
            review = SQLiteResultStore(self._database_path).load_latest_review(
                resolution.record
            )
        except GuidelineSnapshotError as exc:
            self._fail("recommendation", "stale_guidelines", str(exc))
            return
        except Exception as exc:  # noqa: BLE001 - sanitized workflow boundary
            self._fail(
                "recommendation",
                "recommendation_resolution",
                safe_recommendation_error_message(exc),
            )
            return
        self.state.recommendation = resolution
        self.state.review = review
        self._transition(
            "recommendation",
            WorkflowRoute.REUSE if resolution.reused else WorkflowRoute.GENERATE,
        )

    @router(
        resolve_recommendation,
        emit=["abstain", "await_review", "approved", "rejected", "failed"],
    )
    def route_after_recommendation(self) -> str:
        if self.state.failure is not None or self.state.recommendation is None:
            self.state.outcome = WorkflowRoute.FAILED
            return WorkflowRoute.FAILED.value
        if self.state.recommendation.record.result.abstention is not None:
            outcome = WorkflowRoute.ABSTAIN
        elif self.state.review is None:
            outcome = WorkflowRoute.AWAIT_REVIEW
        elif self.state.review.decision == "approved":
            outcome = WorkflowRoute.APPROVED
        else:
            outcome = WorkflowRoute.REJECTED
        self.state.outcome = outcome
        self._transition("workflow", outcome)
        return outcome.value

    def run(self) -> CapstoneWorkflowResult:
        """Run synchronously and return only validated structured state."""

        self.kickoff()
        outcome = self.state.outcome or WorkflowRoute.FAILED
        return CapstoneWorkflowResult(
            analysis=self.state.analysis,
            summary=self.state.summary,
            retrieval=self.state.retrieval,
            recommendation=self.state.recommendation,
            review=self.state.review,
            transitions=list(self.state.transitions),
            failure=self.state.failure,
            outcome=outcome,
        )


def retrieve_guideline_snapshot(
    analysis: ReportAnalysisResult,
    retriever: object,
    *,
    top_k: int,
    max_queries: int,
) -> GuidelineRetrievalSnapshot:
    """Retrieve bounded passages from one stable local index instance."""

    if not 1 <= top_k <= 10:
        raise ValueError("top_k must be between 1 and 10.")
    if not 1 <= max_queries <= 5:
        raise ValueError("max_queries must be between 1 and 5.")
    identity_method = getattr(retriever, "index_fingerprint", None)
    retrieve_method = getattr(retriever, "retrieve", None)
    if not callable(identity_method) or not callable(retrieve_method):
        raise TypeError("Retriever must expose index_fingerprint() and retrieve().")
    before = identity_method()
    queries = build_analysis_retrieval_queries(analysis, max_queries=max_queries)
    passages: list[RetrievedPassage] = []
    try:
        for query in queries:
            passages.extend(retrieve_method(query, top_k=top_k))
    except EmptyGuidelineIndexError:
        passages = []
    after = identity_method()
    if before != after:
        raise GuidelineSnapshotError(
            "The guideline index changed during retrieval; no draft was generated."
        )
    unique: dict[tuple[str, str], RetrievedPassage] = {}
    for passage in passages:
        unique.setdefault((passage.document_hash, passage.chunk_id), passage)
    return GuidelineRetrievalSnapshot(
        index=before,
        queries=queries,
        passages=list(unique.values())[:MAX_RETRIEVED_PASSAGES],
    )


def _validate_recommendation_index(
    snapshot: GuidelineIndexFingerprint,
    settings: RecommendationGenerationSettings,
) -> None:
    if (
        snapshot.fingerprint != settings.guideline_index_fingerprint
        or snapshot.document_hashes != settings.guideline_document_hashes
    ):
        raise GuidelineSnapshotError(
            "The guideline index is incompatible with recommendation settings; "
            "no draft was generated."
        )


def _safe_retrieval_error(error: Exception) -> str:
    if isinstance(error, (GuidelineKnowledgeBaseError, GuidelineSnapshotError)):
        return str(error)
    if isinstance(error, PersistenceError):
        return str(error)
    return "Local guideline retrieval failed. No draft was generated."
