"""Run a fully offline, mocked-boundary capstone workflow demonstration."""

from __future__ import annotations

import gc
import hashlib
import io
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any

from langchain_core.embeddings import Embeddings
from reportlab.pdfgen import canvas

from clinical_report_assistant.capstone_flow import (
    ClinicalReportFlow,
    retrieve_guideline_snapshot,
)
from clinical_report_assistant.guideline_kb import GuidelineKnowledgeBase
from clinical_report_assistant.pdf_processing import PdfDocument, process_pdf
from clinical_report_assistant.persistence import (
    AnalysisGenerationSettings,
    RecommendationGenerationSettings,
    RuntimeVersions,
    SummaryGenerationSettings,
    report_sha256,
)
from clinical_report_assistant.recommendation_agent import (
    ExtractedFindingCitation,
    ExtractedPassageCitation,
    ExtractedRecommendationDraft,
    ExtractedSuggestion,
    RecommendationAgent,
)
from clinical_report_assistant.report_analysis_agent import (
    ExtractedApplicabilityContext,
    ExtractedFinding,
    ExtractedReportAnalysis,
    ReportAnalysisAgent,
)
from clinical_report_assistant.summary_agent import (
    ExtractedReportSummary,
    SummaryAgent,
)

_REPORT_PAGE = (
    "SYNTHETIC CAPSTONE REPORT - NOT MEDICAL ADVICE\n"
    "Glucose result: 118 mg/dL.\n"
    "Glucose reference range: 70-99 mg/dL.\n"
    "Glucose fasting status is unknown."
)
_GUIDANCE_PAGE = (
    "SYNTHETIC CAPSTONE REVIEW GUIDANCE - NOT MEDICAL ADVICE.\n"
    "For a glucose result with unknown fasting status, ask the reviewer to "
    "verify and document fasting status before interpreting the supplied range."
)


class _OfflineEmbeddings(Embeddings):
    """Small deterministic vectors used only by this no-download demo."""

    def embed_documents(self, texts: list[str]) -> list[list[float]]:
        return [self._vector(text) for text in texts]

    def embed_query(self, text: str) -> list[float]:
        return self._vector(text)

    @staticmethod
    def _vector(text: str) -> list[float]:
        digest = hashlib.sha256(text.casefold().encode("utf-8")).digest()
        return [float(value) / 255.0 for value in digest[:16]]


class _OfflineAnalysisRunner:
    def analyze(self, document: PdfDocument) -> ExtractedReportAnalysis:
        del document
        return ExtractedReportAnalysis(
            findings=[
                ExtractedFinding(
                    name="Glucose",
                    reported_value="118",
                    reported_unit="mg/dL",
                    reference_text="70-99 mg/dL",
                    reference_range="70-99",
                    reference_unit="mg/dL",
                    reference_supporting_excerpt=(
                        "Glucose reference range: 70-99 mg/dL."
                    ),
                    specimen_event="synthetic initial specimen",
                    applicability_contexts=[
                        ExtractedApplicabilityContext(
                            text="Glucose fasting status is unknown.",
                            source_page=1,
                            supporting_excerpt=(
                                "Glucose fasting status is unknown."
                            ),
                            specimen_event="synthetic initial specimen",
                        )
                    ],
                    source_page=1,
                    supporting_excerpt="Glucose result: 118 mg/dL.",
                )
            ]
        )


class _OfflineSummaryRunner:
    def summarize(self, analysis: Any) -> ExtractedReportSummary:
        finding_id = analysis.findings[0].finding_id
        return ExtractedReportSummary(
            overview_finding_ids=[finding_id],
            important_finding_ids=[finding_id],
            uncertainty_finding_ids=[finding_id],
        )


class _OfflineRecommendationRunner:
    def recommend(
        self,
        analysis: Any,
        summary: Any,
        passages: list[Any],
    ) -> ExtractedRecommendationDraft:
        del summary
        finding = analysis.findings[0]
        passage = passages[0]
        return ExtractedRecommendationDraft(
            suggestions=[
                ExtractedSuggestion(
                    text=(
                        "Ask the reviewer to verify and document fasting status "
                        "before interpreting the supplied range."
                    ),
                    patient_evidence=[
                        ExtractedFindingCitation(
                            finding_id=finding.finding_id,
                            quoted_excerpt="Glucose result: 118 mg/dL.",
                        )
                    ],
                    guideline_evidence=[
                        ExtractedPassageCitation(
                            passage_id=passage.chunk_id,
                            quoted_excerpt=(
                                "verify and document fasting status before "
                                "interpreting the supplied range"
                            ),
                        )
                    ],
                )
            ]
        )


def _pdf_bytes(text: str) -> bytes:
    output = io.BytesIO()
    pdf = canvas.Canvas(output)
    text_object = pdf.beginText(72, 760)
    for line in text.splitlines():
        text_object.textLine(line)
    pdf.drawText(text_object)
    pdf.showPage()
    pdf.save()
    return output.getvalue()


def _versions() -> RuntimeVersions:
    return RuntimeVersions(
        application="offline-mocked-demo",
        python="3.12",
        pydantic="offline-demo",
        crewai="offline-demo",
        google_genai="not-called",
    )


def main() -> int:
    """Exercise PDF, Chroma, Flow, validation, SQLite, and routing offline."""

    print("MOCKED OFFLINE WORKFLOW DEMO - NO LIVE AI CALLS")
    with TemporaryDirectory(prefix="clinical-capstone-demo-") as temporary:
        root = Path(temporary)
        chroma_path = root / "chroma"
        database_path = root / "results.sqlite3"
        report_bytes = _pdf_bytes(_REPORT_PAGE)
        guidance_bytes = _pdf_bytes(_GUIDANCE_PAGE)
        report = process_pdf(report_bytes, "synthetic_report.pdf")

        indexer = GuidelineKnowledgeBase(
            storage_path=chroma_path,
            embeddings=_OfflineEmbeddings(),
        )
        indexed = indexer.index_pdf(
            guidance_bytes,
            "synthetic_capstone_review_guidance.pdf",
        )
        identity = indexer.index_fingerprint()

        analysis_settings = AnalysisGenerationSettings(
            model="mocked-offline-boundary",
            max_output_tokens=8192,
            timeout_seconds=60,
            schema_hash="1" * 64,
        )
        summary_settings = SummaryGenerationSettings(
            model="mocked-offline-boundary",
            max_output_tokens=4096,
            timeout_seconds=60,
            schema_hash="2" * 64,
        )
        recommendation_settings = RecommendationGenerationSettings(
            model="mocked-offline-boundary",
            max_output_tokens=4096,
            timeout_seconds=60,
            schema_hash="3" * 64,
            retrieval_top_k=4,
            max_retrieval_queries=3,
            guideline_index_fingerprint=identity.fingerprint,
            guideline_document_hashes=identity.document_hashes,
        )

        def retrieve(analysis: Any):
            worker_kb = GuidelineKnowledgeBase(
                storage_path=chroma_path,
                embeddings=_OfflineEmbeddings(),
            )
            try:
                return retrieve_guideline_snapshot(
                    analysis,
                    worker_kb,
                    top_k=recommendation_settings.retrieval_top_k,
                    max_queries=recommendation_settings.max_retrieval_queries,
                )
            finally:
                worker_kb.close()

        def current_identity():
            return GuidelineKnowledgeBase(
                storage_path=chroma_path,
                embeddings=_OfflineEmbeddings(),
            ).index_fingerprint()

        analysis_agent = ReportAnalysisAgent(runner=_OfflineAnalysisRunner())
        summary_agent = SummaryAgent(runner=_OfflineSummaryRunner())
        recommendation_agent = RecommendationAgent(
            runner=_OfflineRecommendationRunner()
        )
        flow = ClinicalReportFlow(
            database_path=database_path,
            report_hash=report_sha256(report_bytes),
            analysis_settings=analysis_settings,
            summary_settings=summary_settings,
            recommendation_settings=recommendation_settings,
            runtime_versions=_versions(),
            generate_analysis=lambda: analysis_agent.analyze(report),
            generate_summary=summary_agent.summarize,
            retrieve_guidelines=retrieve,
            generate_recommendation=recommendation_agent.recommend,
            current_index_identity=current_identity,
        )
        result = flow.run()

        print(f"Indexed synthetic guideline chunks: {indexed.chunks_added}")
        print(
            "Route: "
            + " -> ".join(
                f"{transition.stage}:{transition.route.value}"
                for transition in result.transitions
            )
        )
        print(f"Final outcome: {result.outcome.value}")
        print(
            "Validated suggestions: "
            f"{len(result.recommendation.record.result.suggestions)}"
        )
        indexer.close()
        print("Temporary Chroma and SQLite resources were isolated for this run.")
        del flow, result, indexer
        gc.collect()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
