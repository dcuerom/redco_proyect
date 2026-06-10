# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A data-science / business-intelligence workspace for **REDCO**, a strategic-transformation consulting program. The work is bilingual: documentation and domain concepts are in **Spanish** (with English technical terms preserved intentionally — `Capability`, `Personal Identity`, `Mindset`, `Stratex`, `BSC`), while code is in English. The repository is in its early stages — the conceptual context is fully documented but the analytical code (`pilar_a/notebooks/pruebas.ipynb`) is essentially empty.

The intended consumer of `context/*.md` is an AI agent specialized in data science and BI; treat those documents as the authoritative source of truth for domain concepts, terminology, and what metrics/artifacts the analysis must produce.

## Layout

The git repo root (`proyect_I/`) is self-contained — it holds the `.venv`, `CLAUDE.md`, and `.claude/`. The repo currently has no commits; everything is untracked.

- `context/` — the conceptual specification (see below).
- `pilar_a/` — strategic-dimension analysis, with `notebooks/`, `scripts/`, `data/` subdirs.
- `pilar_b/` — organizational-dimension analysis (directory exists, empty).
- `.venv/` — the Python environment (untracked; should be gitignored once a `.gitignore` is added).

## Environment & commands

The virtualenv lives at the repo root (`.venv/`). Python 3.12; the data/ML stack installed for the agent skills (see "Configured skills" below) is `pandas` 3.0, `numpy` 2.4, `scipy` 1.17, `scikit-learn` 1.9, `statsmodels` 0.14, `umap-learn` 0.5 + `hdbscan`, `shap` 0.52, `networkx` 3.6, `matplotlib` 3.10, `seaborn` 0.13, `polars` 1.41, plus document I/O (`markitdown`, `python-pptx`, `python-docx`, `openpyxl`/`python-calamine`, `pypdf`/`pdfplumber`) and full Jupyter/JupyterLab. No `requirements.txt`, lint, or test setup exists yet — there is nothing to build or test.

```bash
# Activate the env (run from the proyect_I repo root)
source .venv/bin/activate

# Run JupyterLab
jupyter lab

# Execute a notebook headless
jupyter nbconvert --to notebook --execute pilar_a/notebooks/pruebas.ipynb
```

## Configured skills

A curated subset of [K-Dense-AI/scientific-agent-skills](https://github.com/K-Dense-AI/scientific-agent-skills) is installed at **project level** in `.claude/skills/` (18 of the repo's 144 — only those that serve a BI / strategy / data-science consulting workflow; the bio/chem/physics skills were deliberately left out). Each skill is a folder with a `SKILL.md` plus reference docs and runnable scripts. Their Python dependencies are already installed in `.venv` (versions above), so the scripts run as-is. Invoke a skill via the Skill tool by its name.

Mapping skills to where they apply in this project:

**Modeling & analytics — primarily `pilar_b` (human-capital baseline) and `pilar_a` KPIs**
- `scikit-learn` — clustering of organizational profiles, capability-gap models, the Capability × Personal-Identity 2×2 segmentation, pipelines.
- `umap-learn` — dimensionality reduction / 2D visualization of the profile space (pairs with hdbscan for density clustering).
- `statistical-analysis` — descriptive stats, hypothesis tests, correlations (e.g. validating value-creation hypotheses from `pilar_a`).
- `statsmodels` — regression and inferential models for the business case / KPI drivers.
- `shap` — explainability of any predictive model (which features drive a profile classification or KPI).
- `exploratory-data-analysis` — first-pass profiling of any new dataset (`pilar_a/data/`, `pilar_b` evaluations).

**Visualization — Balanced Scorecard, dashboards, strategy maps**
- `matplotlib`, `seaborn` — core plotting.
- `scientific-visualization` — publication-quality, grayscale-safe figures for the deliverables.
- `networkx` — strategy maps, decision-governance graphs, org/relationship networks.

**Document I/O — ingest sources (.pptx) and produce consulting deliverables**
- `markitdown` — convert `.pptx` / `.docx` / `.pdf` / `.xlsx` to Markdown. This is the exact workflow that produced `context/*.md` from `260526_Gestión_Estratégica_REDCO.pptx`; use it to ingest new source decks.
- `pptx` — read/build PowerPoint (source format of the strategy decks; also for output presentations).
- `docx` — Word reports.
- `xlsx` — Excel models / BSC spreadsheets (read with `python-calamine`, write with `openpyxl`).
- `pdf` — read/split/merge PDFs and OCR scans (OCR also needs system `poppler-utils` + `tesseract-ocr`, not yet installed).

**Writing & data wrangling**
- `scientific-writing` — structured long-form reports.
- `markdown-mermaid-writing` — Markdown docs with Mermaid diagrams (the documentation default for this repo; matches the style of `context/*.md`).
- `polars` — fast dataframe alternative to pandas for larger evaluation datasets.

> The full upstream catalog (144 skills) is also available globally via the Skill tool; the list above is the project-relevant subset whose dependencies are guaranteed present in this `.venv`.

When adding dependencies, install into this `.venv`. If you introduce tests or linting, there is no existing convention to follow — pick one and document it here.

## Domain architecture (the "two pillars")

The whole project is organized around two interdependent analytical pillars defined in `context/current_iteration.md`. Code directories (`pilar_a`, `pilar_b`) map directly onto these — keep that correspondence when adding analysis.

- **`pilar_a` — Strategic dimension** ("hacia dónde"). Produces the strategic KPIs. Its outputs (value-creation hypotheses, prioritized initiatives) are the **direct inputs to the Balanced Scorecard and operational dashboards**. This is the source of the strategic-control metrics.

- **`pilar_b` — Organizational dimension** ("con quiénes"). Builds an organizational baseline by evaluating each person on two cross-cutting sub-dimensions: **Capability** ("how you do it") and **Personal Identity** ("who you are"). The combined 2×2 matrix is the analytical frame for human-capital indicators — expect clustering of organizational profiles, capability-gap analysis, and cultural-compatibility metrics. This baseline enables longitudinal comparison across future iterations.

A third, non-analytical dimension (integrated Finance & Strategy capability) institutionalizes the pillars' outputs but produces no data artifacts.

## Conceptual frameworks the analysis is built on

These are referenced throughout the context docs and define the BI artifacts the project must deliver:

- **Kaplan–Norton "Execution Premium"** — a 6-stage closed-loop strategy model. The central BI artifacts are the **Balanced Scorecard (BSC)**, **strategy map**, **Stratex** (strategic budget), and **dashboards**. `pilar_a` feeds stages 1–2 (Develop/Translate Strategy); `pilar_b` feeds stage 3 (Align).
- **Two-Agenda system** — *Agenda de Cumplimiento* (compliance KPIs, deliver commitments) and *Agenda de Aspiraciones* (track strategic hypotheses, build future capabilities). Note: the 2024 "Agenda de Desarrollo" was renamed "Agenda de Aspiraciones" — content is equivalent.
- **Three-level decision governance** — strategic / management / operational.

## Context files

- `context/first_iteration.md` — the 2024 strategic-definition iteration (diagnosis, core business, propósito/misión/aspiración, initial KPIs).
- `context/current_iteration.md` — the **current** phase (REDCO 2028 horizon, the pillar definitions above, frameworks, team). This supersedes and builds on the first; when the two conflict, prefer `current_iteration.md`.

The overarching meta-objective stated in the docs — relevant when prioritizing what to analyze — is **reducing the organization's operational dependence on the founder-owner**; treat that as the primary risk variable the BI work should illuminate.
