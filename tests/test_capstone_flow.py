"""Offline orchestration tests for the three-agent CrewAI Flow."""

from __future__ import annotations

import threading
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest

from clinical_report_assistant.capstone_flow import (
    CapstoneWorkflowResult,
    ClinicalReportFlow,
    GuidelineRetrievalSnapshot,
    WorkflowRoute,
    retrieve_guideline_snapshot,
)
from clinical_report_assistant.guideline_kb import (
    GuidelineIndexFingerprint,
    RetrievedPassage,
)
from clinical_report_assistant.persistence import (
    AnalysisGenerationSettings,
    RecommendationGenerationSettings,
    RuntimeVersions,
    SQLiteResultStore,
    SummaryGenerationSettings,
    report_sha256,
)
from clinical_report_assistant.recommendation_agent import (
    AbstentionCategory,
    ExtractedFindingCitation,
    ExtractedPassageCitation,
    ExtractedRecommendationDraft,
    ExtractedSuggestion,
    RecommendationAbstention,
    RecommendationAgent,
    RecommendationDraftResult,
)
from clinical_report_assistant.report_analysis_agent import (
    ApplicabilityStatus,
    FindingAssessment,
    FindingKind,
    ReportAnalysisResult,
    ReportFinding,
)
from clinical_report_assistant.summary_agent import (
    ExtractedReportSummary,
    ReportSummaryResult,
    SummaryAgent,
)

_DOCUMENT_HASH = "a" * 64
_INDEX_HASH = "b" * 64


class _SummaryRunner:
    def summarize(self, analysis: ReportAnalysisResult) -> ExtractedReportSummary:
        finding_id = analysis.findings[0].finding_id
        return ExtractedReportSummary(
            overview_finding_ids=[finding_id],
            important_finding_ids=[finding_id],
        )


class _RecommendationRunner:
    def recommend(
        self,
        analysis: ReportAnalysisResult,
        summary: ReportSummaryResult,
        passages: list[RetrievedPassage],
    ) -> ExtractedRecommendationDraft:
        del summary
        finding = analysis.findings[0]
        passage = passages[0]
        return ExtractedRecommendationDraft(
            suggestions=[
                ExtractedSuggestion(
                    text="Ask the reviewer to verify and document fasting status.",
                    patient_evidence=[
                        ExtractedFindingCitation(
                            finding_id=finding.finding_id,
                            quoted_excerpt="Glucose result: 118 mg/dL",
                        )
                    ],
                    guideline_evidence=[
                        ExtractedPassageCitation(
                            passage_id=passage.chunk_id,
                            quoted_excerpt="verify and document fasting status",
                        )
                    ],
                )
            ]
        )


class _Counter:
    def __init__(self, function: Callable[..., Any]) -> None:
        self.function = function
        self.calls = 0

    def __call__(self, *args: Any) -> Any:
        self.calls += 1
        return self.function(*args)


def _analysis() -> ReportAnalysisResult:
    finding = ReportFinding(
        name="Glucose",
        reported_value="118",
        reported_date="2026-01-01",
        reported_unit="mg/dL",
        reference_text="70-99 mg/dL",
        reference_range="70-99",
        reference_unit="mg/dL",
        specimen_event="initial",
        applicability_status=ApplicabilityStatus.REVIEW_REQUIRED,
        finding_kind=FindingKind.PATIENT_MEASUREMENT,
        finding_id="glucose-initial",
        source_page=1,
        supporting_excerpt="Glucose result: 118 mg/dL; fasting status unknown.",
        assessment=FindingAssessment.ABOVE_RANGE,
        review_required=True,
        review_reasons=["Fasting status is unknown."],
    )
    return ReportAnalysisResult(filename="synthetic.pdf", findings=[finding])


def _summary(analysis: ReportAnalysisResult) -> ReportSummaryResult:
    return SummaryAgent(runner=_SummaryRunner()).summarize(analysis)


def _passage() -> RetrievedPassage:
    return RetrievedPassage(
        text=(
            "SYNTHETIC CAPSTONE REVIEW GUIDANCE - NOT MEDICAL ADVICE. "
            "For glucose with unknown fasting status, ask the reviewer to verify "
            "and document fasting status before interpreting the supplied range."
        ),
        document_name="synthetic_capstone_review_guidance.pdf",
        document_hash=_DOCUMENT_HASH,
        page_number=1,
        chunk_id="glucose-guidance",
    )


def _identity(*, fingerprint: str = _INDEX_HASH) -> GuidelineIndexFingerprint:
    return GuidelineIndexFingerprint(
        fingerprint=fingerprint,
        document_hashes=[_DOCUMENT_HASH],
        configuration={"embedding_model": "offline-fake"},
    )


def _snapshot() -> GuidelineRetrievalSnapshot:
    return GuidelineRetrievalSnapshot(
        index=_identity(),
        queries=["review follow-up guidance for Glucose"],
        passages=[_passage()],
    )


def _draft(
    analysis: ReportAnalysisResult,
    summary: ReportSummaryResult,
    passages: list[RetrievedPassage],
) -> RecommendationDraftResult:
    return RecommendationAgent(runner=_RecommendationRunner()).recommend(
        analysis,
        summary,
        passages,
    )


def _versions() -> RuntimeVersions:
    return RuntimeVersions(
        application="test-app",
        python="3.12",
        pydantic="test",
        crewai="test",
        google_genai="test",
    )


def _settings() -> tuple[
    AnalysisGenerationSettings,
    SummaryGenerationSettings,
    RecommendationGenerationSettings,
]:
    return (
        AnalysisGenerationSettings(
            model="synthetic-model",
            max_output_tokens=8192,
            timeout_seconds=60,
            schema_hash="1" * 64,
        ),
        SummaryGenerationSettings(
            model="synthetic-model",
            max_output_tokens=4096,
            timeout_seconds=60,
            schema_hash="2" * 64,
        ),
        RecommendationGenerationSettings(
            model="synthetic-model",
            max_output_tokens=4096,
            timeout_seconds=60,
            schema_hash="3" * 64,
            retrieval_top_k=4,
            max_retrieval_queries=3,
            guideline_index_fingerprint=_INDEX_HASH,
            guideline_document_hashes=[_DOCUMENT_HASH],
        ),
    )


def _flow(
    database_path: Path,
    *,
    generate_analysis: Callable[[], ReportAnalysisResult] = _analysis,
    generate_summary: Callable[[ReportAnalysisResult], ReportSummaryResult] = _summary,
    retrieve_guidelines: Callable[
        [ReportAnalysisResult], GuidelineRetrievalSnapshot
    ] = lambda analysis: _snapshot(),
    generate_recommendation: Callable[
        [ReportAnalysisResult, ReportSummaryResult, list[RetrievedPassage]],
        RecommendationDraftResult,
    ] = _draft,
    current_index_identity: Callable[
        [], GuidelineIndexFingerprint
    ] = _identity,
) -> ClinicalReportFlow:
    analysis_settings, summary_settings, recommendation_settings = _settings()
    return ClinicalReportFlow(
        database_path=database_path,
        report_hash=report_sha256(b"synthetic report"),
        analysis_settings=analysis_settings,
        summary_settings=summary_settings,
        recommendation_settings=recommendation_settings,
        runtime_versions=_versions(),
        generate_analysis=generate_analysis,
        generate_summary=generate_summary,
        retrieve_guidelines=retrieve_guidelines,
        generate_recommendation=generate_recommendation,
        current_index_identity=current_index_identity,
    )


def _routes(result: CapstoneWorkflowResult) -> list[tuple[str, WorkflowRoute]]:
    return [(item.stage, item.route) for item in result.transitions]


def test_retrieval_queries_need_validated_analysis_but_no_summary() -> None:
    class Retriever:
        def __init__(self) -> None:
            self.queries: list[str] = []

        def index_fingerprint(self) -> GuidelineIndexFingerprint:
            return _identity()

        def retrieve(self, query: str, *, top_k: int) -> list[RetrievedPassage]:
            assert top_k == 2
            self.queries.append(query)
            return [_passage()]

    retriever = Retriever()
    snapshot = retrieve_guideline_snapshot(
        _analysis(),
        retriever,
        top_k=2,
        max_queries=3,
    )

    assert retriever.queries
    assert all("Glucose" in query for query in retriever.queries)
    assert snapshot.passages == [_passage()]


def test_summary_and_retrieval_branches_overlap_and_join(tmp_path: Path) -> None:
    barrier = threading.Barrier(2)
    branch_threads: set[int] = set()

    def summary_branch(analysis: ReportAnalysisResult) -> ReportSummaryResult:
        branch_threads.add(threading.get_ident())
        barrier.wait(timeout=2)
        return _summary(analysis)

    def retrieval_branch(
        analysis: ReportAnalysisResult,
    ) -> GuidelineRetrievalSnapshot:
        assert analysis.findings[0].finding_id == "glucose-initial"
        branch_threads.add(threading.get_ident())
        barrier.wait(timeout=2)
        return _snapshot()

    result = _flow(
        tmp_path / "results.sqlite3",
        generate_summary=summary_branch,
        retrieve_guidelines=retrieval_branch,
    ).run()

    assert len(branch_threads) == 2
    assert result.outcome is WorkflowRoute.AWAIT_REVIEW
    assert ("summary", WorkflowRoute.GENERATE) in _routes(result)
    assert ("guideline_retrieval", WorkflowRoute.RETRIEVE) in _routes(result)
    assert result.recommendation is not None


@pytest.mark.parametrize("failed_branch", ["summary", "retrieval"])
def test_branch_failure_preserves_successful_stages(
    tmp_path: Path,
    failed_branch: str,
) -> None:
    recommendation = _Counter(_draft)

    def fail_summary(analysis: ReportAnalysisResult) -> ReportSummaryResult:
        del analysis
        raise RuntimeError("SENSITIVE REPORT CONTENT")

    def fail_retrieval(
        analysis: ReportAnalysisResult,
    ) -> GuidelineRetrievalSnapshot:
        del analysis
        raise RuntimeError("SENSITIVE GUIDELINE CONTENT")

    result = _flow(
        tmp_path / "results.sqlite3",
        generate_summary=fail_summary if failed_branch == "summary" else _summary,
        retrieve_guidelines=(
            fail_retrieval if failed_branch == "retrieval" else lambda analysis: _snapshot()
        ),
        generate_recommendation=recommendation,
    ).run()

    assert result.outcome is WorkflowRoute.FAILED
    assert result.analysis is not None
    assert recommendation.calls == 0
    assert "SENSITIVE" not in result.failure.message
    if failed_branch == "summary":
        assert result.summary is None
        assert result.retrieval is not None
    else:
        assert result.summary is not None
        assert result.retrieval is None
        assert SQLiteResultStore(tmp_path / "results.sqlite3").load_summary_by_id(
            result.summary.record.record_id
        )


def test_changed_guideline_index_fails_before_recommendation(tmp_path: Path) -> None:
    recommendation = _Counter(_draft)
    changed = _identity(fingerprint="c" * 64)

    result = _flow(
        tmp_path / "results.sqlite3",
        generate_recommendation=recommendation,
        current_index_identity=lambda: changed,
    ).run()

    assert result.outcome is WorkflowRoute.FAILED
    assert result.failure.category == "stale_guidelines"
    assert result.analysis is not None
    assert result.summary is not None
    assert recommendation.calls == 0


def test_recommendation_failure_preserves_valid_upstream_results(
    tmp_path: Path,
) -> None:
    def fail_recommendation(
        analysis: ReportAnalysisResult,
        summary: ReportSummaryResult,
        passages: list[RetrievedPassage],
    ) -> RecommendationDraftResult:
        del analysis, summary, passages
        raise RuntimeError("SENSITIVE GENERATED CONTENT")

    result = _flow(
        tmp_path / "results.sqlite3",
        generate_recommendation=fail_recommendation,
    ).run()

    assert result.outcome is WorkflowRoute.FAILED
    assert result.analysis is not None
    assert result.summary is not None
    assert result.recommendation is None
    assert "SENSITIVE" not in result.failure.message


def test_abstention_is_an_explicit_terminal_route(tmp_path: Path) -> None:
    def abstain(
        analysis: ReportAnalysisResult,
        summary: ReportSummaryResult,
        passages: list[RetrievedPassage],
    ) -> RecommendationDraftResult:
        del analysis, summary, passages
        return RecommendationDraftResult(
            abstention=RecommendationAbstention(
                category=AbstentionCategory.INSUFFICIENT_GUIDANCE,
                reason="Retrieved synthetic guidance was insufficient.",
            )
        )

    result = _flow(
        tmp_path / "results.sqlite3",
        generate_recommendation=abstain,
    ).run()

    assert result.outcome is WorkflowRoute.ABSTAIN
    assert ("recommendation", WorkflowRoute.GENERATE) in _routes(result)


def test_restart_reuses_agents_and_preserves_exact_review(tmp_path: Path) -> None:
    database_path = tmp_path / "results.sqlite3"
    first = _flow(database_path).run()
    assert first.recommendation is not None
    saved_review = SQLiteResultStore(database_path).save_review_decision(
        first.recommendation.record,
        "approved",
        "Reviewed synthetic draft.",
    )
    analysis_calls = _Counter(_analysis)
    summary_calls = _Counter(_summary)
    recommendation_calls = _Counter(_draft)

    second = _flow(
        database_path,
        generate_analysis=analysis_calls,
        generate_summary=summary_calls,
        generate_recommendation=recommendation_calls,
    ).run()

    assert analysis_calls.calls == 0
    assert summary_calls.calls == 0
    assert recommendation_calls.calls == 0
    assert second.outcome is WorkflowRoute.APPROVED
    assert second.review == saved_review
    assert second.recommendation.record.record_id == first.recommendation.record.record_id
    assert ("analysis", WorkflowRoute.REUSE) in _routes(second)
    assert ("summary", WorkflowRoute.REUSE) in _routes(second)
    assert ("recommendation", WorkflowRoute.REUSE) in _routes(second)
