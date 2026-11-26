# JARVIS Apprentice v0.9.1

Continual Personalization of AI Agents through Human Corrections

## Overview

JARVIS Apprentice is a personalized AI agent that learns reusable procedural rules from human corrections without model fine-tuning.

The current application focuses on medical lecture note preparation and cross-year educational content integration.

## Key Features

- Automated PDF-to-Word document generation
- Correction-driven procedural memory
- Human-approved rule learning
- Memory OFF/ON A/B comparison
- Cross-year lecture slide comparison
- Human-in-the-loop content integration
- Quantitative evaluation

## System Workflow

Execute → Correct → Learn → Reuse

## Experimental Results

- 3 medical lecture topics
- 25 chapter-level tasks
- 50 successful executions
- 3 approved procedural rules
- Test-set image wrapping accuracy: 0% → 100% (53 images)
- Matched paragraph style accuracy: 94.35% → 95.25%
- Text retention proxy: 100% → 93.74%

The results demonstrate cross-topic reuse of specific procedural rules, while overall content preservation requires further improvement.

## Installation

Python 3.11 is recommended.

```bash
pip install -r requirements.txt
python -m streamlit run app.py
```

Open http://localhost:8501 in your browser.

## Project Structure

- app.py: Streamlit user interface
- jarvis.py: Document generation and rule execution
- learn.py: Correction-driven rule learning
- content_comparator.py: Lecture content comparison
- evaluation.py: Quantitative evaluation
- dataset_loader.py: Dataset management
- experiment_runner.py: A/B experiment execution

## Data Availability

Original medical lecture materials, generated documents, and private correction records are not publicly distributed.

Users must provide their own authorized documents and templates.

## Limitations

- Procedural learning currently relies on domain-specific heuristics.
- Potential new content requires human verification.
- Content comparison does not guarantee medical correctness.
- Reduction in human correction time requires further evaluation.

## Future Work

Extend procedural memory to diverse personal workflows and develop an adaptive AI assistant capable of learning continuously through human collaboration.

## Academic Project

Introduction to Artificial Intelligence
Final Project — HW5
Group 191
Version 0.9.1
November 2026
