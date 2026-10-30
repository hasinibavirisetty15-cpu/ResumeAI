# ResumeAI --- AI Resume Analyzer & ATS Optimizer

A Streamlit portfolio project that compares a resume with a job
description and provides ATS-style heuristic scores, skill and keyword
matches, improvement suggestions, an optimized resume draft, and
interview preparation materials.

> **Important:** This project uses rule-based text matching and
> heuristics. Its scores are estimates, not scores from an employer's
> ATS, not a hiring prediction, and not a guarantee of interview
> selection.

## Features

-   Upload resumes in PDF or DOCX format.
-   Extract and display resume text.
-   Compare detected skills and keywords against a pasted job
    description.
-   Calculate ATS-style, job-match, skill-match, keyword-match,
    structure, and content scores.
-   Show matched and potentially missing skills/keywords.
-   Provide strengths, improvement suggestions, and a suggested
    professional summary.
-   Generate an optimized resume draft for review.
-   Suggest projects and interview questions, with a preparation plan.
-   Download a PDF analysis report.
-   Reset the current analysis.
-   Handle empty job descriptions and unsupported file types with
    user-facing messages.

## Requirements

-   Python 3.10 or later recommended
-   The packages listed in `requirements.txt`
-   Visual Studio Code or another terminal

## Setup on Windows

1.  Open the project folder in VS Code.

2.  Open **Terminal → New Terminal**.

3.  (Recommended) Create a virtual environment:

    ``` powershell
    python -m venv .venv
    ```

4.  Activate it:

    ``` powershell
    .\.venv\Scripts\Activate.ps1
    ```

    If PowerShell blocks activation, you can skip virtual-environment
    activation and run the next commands with your normal Python
    installation.

5.  Install dependencies:

    ``` powershell
    python -m pip install -r requirements.txt
    ```

6.  Start the app from the folder containing `resume_analyzer.py`:

    ``` powershell
    python -m streamlit run resume_analyzer.py
    ```

7.  Open the local address printed in the terminal (usually
    `http://localhost:8501`).

## How to use

1.  Upload a text-based PDF or DOCX resume.
2.  Paste the target job description.
3.  Select **Analyze Resume**.
4.  Review the overview, job match, skills, improvement suggestions,
    project suggestions, optimized resume, interview preparation,
    structure, and extracted resume text.
5.  Download the PDF report if needed.

Scanned/image-only PDFs may not provide extractable text. If extraction
fails, try a text-based PDF or DOCX.

## Test checklist

-   [ ] PDF resume upload and analysis
-   [ ] DOCX resume upload and analysis
-   [ ] Empty job description shows a helpful message
-   [ ] Reset clears the previous analysis
-   [ ] Different job descriptions produce updated results
-   [ ] PDF report downloads and opens
-   [ ] Unsupported file types are rejected safely
-   [ ] Optimized resume and interview preparation sections display

## Privacy and responsible use

Use resumes only with the person's permission. The app is designed not
to intentionally save uploaded resumes to a database, but users should
still avoid uploading sensitive documents to environments they do not
control. Verify all extracted information and generated suggestions.
Never add achievements, skills, or metrics that the candidate cannot
substantiate.

## Portfolio description

**ResumeAI --- AI Resume Analyzer & ATS Optimizer:** Built a Streamlit
application that extracts resume content from PDF/DOCX files, compares
skills and keywords against job descriptions, calculates heuristic
ATS-style match scores, suggests resume improvements, generates an
optimized draft, and provides interview preparation resources with
downloadable PDF reports.

## Suggested next improvements

-   Add automated unit tests for score calculations and file extraction.
-   Improve keyword matching for synonyms and related phrases.
-   Add more robust handling for scanned PDFs using OCR if needed.
-   Add deployment instructions if publishing the app online.
-   Consider accessibility and security review before public deployment.
