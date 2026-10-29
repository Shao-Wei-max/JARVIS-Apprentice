# JARVIS Apprentice

**Correction-driven procedural memory for personalized document processing**

JARVIS Apprentice is a research prototype that investigates whether an AI agent can reuse human-approved document-formatting corrections across future tasks, without modifying model parameters. The current **v0.5** implementation focuses on converting medical lecture-note PDFs into editable Word documents using a supplied DOCX template.

> **Project status:** Functional research prototype. Rule proposals are heuristic and require human approval. This release does **not** implement LLM-based rule induction or establish end-to-end medical content accuracy.

## Research question

Can persistent, human-approved procedural rules reduce repeated formatting corrections across related document-processing tasks?

The prototype implements an **execute → correct → learn → reuse** loop. It stores procedural rules externally in SQLite and supports controlled memory-on versus memory-off comparisons.

## Implemented features

| Component | Current implementation |
| --- | --- |
| Workflow initialization | Uses a user-supplied DOCX template and fixed document-processing defaults. |
| Workflow executor | Extracts PDF text blocks, image regions, and URL annotations; generates a DOCX using template styles and optional approved rules. |
| Correction analyzer | Compares original and human-corrected DOCX files for paragraph style, content/structure, margin, and image-wrapping changes. |
| Procedural memory | Persists candidate rules in SQLite; rules must be explicitly approved before use. |
| Evaluation | Compares memory-off and memory-on structural outputs, and optionally evaluates outputs against a reference DOCX. |

The Streamlit interface has four tabs: **Execute & A/B**, **Correct & Learn**, **Memory**, and **Evaluate**.

## Repository files

- `app.py` — Streamlit interface and workflow orchestration.
- `jarvis.py` — PDF extraction, DOCX generation, correction analysis, rule proposals, SQLite memory, and DOCX integrity checks.
- `evaluation.py` — Structural evaluation and A/B comparison utilities.
- `requirements.txt` — Python package dependencies.

## Installation

**Recommended:** Python 3.11.

```bash
python -m venv .venv
```

On Windows (PowerShell):

```powershell
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

On macOS/Linux:

```bash
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Streamlit will display a local URL (usually `http://localhost:8501`). No hosted demo is provided by this repository.

## How to use

1. **Execute & A/B:** Upload a source PDF and DOCX template, select the page range, and generate documents with memory disabled and/or enabled. Download the resulting DOCX and JSON reports.
2. **Correct & Learn:** Edit a generated DOCX manually in Word. Upload the original and corrected versions, analyze differences, and propose candidate rules.
3. **Memory:** Inspect proposed rules and explicitly approve the rules you want to reuse. Approved rules are stored in `workspace/memory.db` locally.
4. **Evaluate:** Upload a reference DOCX and compare structural proxy metrics for available A/B outputs.

Generated outputs and the local database are stored in `workspace/`. Do not commit real medical records, sensitive documents, or local memory databases to a public repository.

## Rule examples

The current rule-proposal logic can suggest:

- **Image wrapping:** Apply Microsoft Word's **Top and Bottom** wrapping after a corresponding human correction.
- **Margins:** Reuse a corrected uniform page margin when within a supported range.
- **Paragraph continuation:** Suggest context-sensitive style inheritance after repeated style changes; the heuristic requires review.
- **Legacy header removal:** Suggest filtering a recognized prior-year `Conote for KMU Medicine ...` header in an expected page-top region.

These are candidate or conditional behaviors, **not guaranteed automatic learning outcomes**. A rule affects later generations only when it has been approved and its implementation conditions are met.

## Evaluation and interpretation

`evaluation.py` reports structural measurements such as image count, Top-and-Bottom image-wrapping rate, page margins, exact-text-matched paragraph style accuracy, and a text-similarity proxy. The A/B comparison also reports approved rule-application events.

**Important limitations:** These measurements do not establish clinical accuracy, semantic content preservation, or overall document correctness. Text extraction order, paragraph segmentation, tables, and image-contained text may require manual review. Results from a small set of pages should not be generalized to all documents.

## Planned research extensions

The research proposal additionally envisions LLM-based correction interpretation, confidence-aware rule refinement, broader document formatting actions, and comparisons with LLM-only and fixed-workflow baselines. **These are future directions, not features claimed as complete in v0.5.**

## Safety and privacy

This software is a research prototype, not a clinical document-processing system. Use synthetic or authorized test documents, review generated outputs manually, and keep personal data, API credentials, and `workspace/` files out of public commits.

