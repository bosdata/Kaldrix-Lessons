# Kaldrix Data Analysis Mastery Programme

A Kaldrix Limited practical learning system that develops learners from **absolute beginner** to **expert-track data analyst** through 50 trainer-reviewed lessons.

## Progression

- **Beginner** — Google Sheets foundations, descriptive statistics, formulas, cleaning and charts.
- **Intermediate** — multi-condition formulas, lookups, pivots, reconciliation, QA and KPI design.
- **Advanced** — SQL, Power BI, DAX, statistics and dashboard design.
- **Pro** — Python/pandas, reproducible analysis, automation, experimentation, forecasting and governance.
- **Expert** — problem framing, data architecture, metric governance, model audit, causal reasoning, AI verification and stakeholder defense.

## Main applications

- `Kaldrix_Data_Analysis_Mastery_Learner.html` — learner application.
- `Kaldrix_Data_Analysis_Mastery_Trainer.html` — trainer review/approval application.

The generated HTML files are self-contained: lesson PDFs and day-specific exercise datasets are embedded so the files can be shared directly.

## Submission and approval workflow

1. Learner studies the lesson.
2. Learner downloads the day-specific exercise dataset.
3. For Google Sheets lessons, the dataset is imported into Google Sheets and completed there.
4. Learner shares the Google Sheet and submits its URL in the app.
5. Trainer imports the learner course record and reviews the submission.
6. Trainer approves or returns the lesson with feedback.
7. Learner imports the reviewed record and marks an approved lesson **Done**.
8. Progress counts completed lessons.
9. A level certificate unlocks only after all 10 lessons in that level are approved and marked Done.

## AI Learning Lab

The AI Lab is deliberately **API-key free** in the standalone HTML. It builds a lesson-aware coaching prompt, reliably copies it even when the app is opened from `file://`, and opens ChatGPT in a new tab. This avoids placing a secret API key in a distributed learner file.

Learners can use the coach for concept explanation, one-hint-at-a-time guidance, formula/SQL/DAX/Python debugging, challenging conclusions, mock trainer questioning, extra practice, dashboard review, and AI-output verification.

The programme teaches: **attempt first → ask AI intelligently → verify independently → explain the result yourself**.

## Build

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# Linux/macOS: source .venv/bin/activate
pip install -r requirements.txt
python build_project.py
```

The build creates all 50 lesson PDFs, all 50 exercise datasets, both self-contained HTML applications, the roadmap/assessor PDFs and a distributable ZIP.

GitHub Actions rebuilds the programme automatically whenever the source curriculum, template or build logic changes.

## Kaldrix Limited

Designed and built by **Kaldrix Limited**.

Programme certificates contain the authorization block:

**Brian Sifuna Obware**  
CEO & Founder  
Kaldrix Limited
