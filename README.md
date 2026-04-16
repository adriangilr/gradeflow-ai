# GradeFlow AI (Hybrid Pipeline)
A semi-automated grading system that combines heuristic rules, AI-assisted evaluation, and human validation to process classroom submissions at scale.

## Why I built this
This project solves a practical workflow problem: collecting assignment deliverables and attached files from Google Classroom in a structured way.

It also reflects the kind of work I enjoy most: combining automation, APIs, file handling, and traceable outputs to reduce manual effort.

## Main features
- Rule-based grading (deterministic scoring)
- AI-assisted feedback generation
- Human-in-the-loop review for edge cases
- Structured CSV output for auditability
- Support for multiple file types (PDF, DOCX, etc.)
- Late submission penalty handling


## Tech stack
- Python
- Google Classroom API
- Google Drive API
- Pandas
- OAuth 2.0

### Output Fields
- submission_status
- is_readable
- word_count
- keyword_hits
- auto_score
- auto_feedback
- requires_manual_review
- final_grade


## Project structure
```text
classroom-downloader/
├── README.md
├── requirements.txt
├── .gitignore
├── .env.example
├── credentials/
├── data/
├── src/
└── tests/



┌──────────────────────────────────────────────┐
│ HYBRID CLASSROOM GRADING PIPELINE           │
│ Heuristic + AI + Human-in-the-loop          │
└──────────────────────────────────────────────┘


[INPUT]                [PROCESS]                         [OUTPUT]

Submissions ───▶  File Processing ───▶  Heuristic Score ───▶ CSV Export
(files, API)       (readability, type)     (+ rules)

                         │
                         ▼
                   AI Feedback (optional)
                         │
                         ▼
                  Manual Review Flag
                         │
                         ▼
                   Final Grade / Feedback




### System Flow

           ┌──────────────────────┐
           │       INPUT          │
           │──────────────────────│
           │ Google Classroom API │
           │ Student submissions  │
           │ Files (PDF, DOCX...) │
           └──────────┬───────────┘
                      │
                      ▼
           ┌──────────────────────┐
           │  FILE PROCESSING     │
           │──────────────────────│
           │ - Detect file type   │
           │ - Extract text       │
           │ - Check readability  │
           └──────────┬───────────┘
                      │
                      ▼
           ┌──────────────────────┐
           │  HEURISTIC ENGINE    │
           │──────────────────────│
           │ - Word count         │
           │ - Keyword hits       │
           │ - Late penalty       │
           │ - Base scoring       │
           └──────────┬───────────┘
                      │
                      ▼
           ┌──────────────────────┐
           │   AI LAYER (optional)│
           │──────────────────────│
           │ - Feedback generation│
           │ - Content analysis   │
           └──────────┬───────────┘
                      │
                      ▼
           ┌──────────────────────┐
           │     DECISION         │
           │──────────────────────│
           │ Is readable?         │
           │                      │
           │  Yes ──▶ continue    │
           │  No  ──▶ manual      │
           └──────────┬───────────┘
                      │
                      ▼
           ┌──────────────────────┐
           │       OUTPUT         │
           │──────────────────────│
           │ CSV Export           │
           │ - auto_score         │
           │ - auto_feedback      │
           │ - confidence_score   │
           │ - requires_manual    │
           │ - final_grade        │
           └──────────────────────┘

---
