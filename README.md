# AI Job Search Agent

A Python-based job discovery and matching tool that collects public job listings, evaluates them against configurable role and skill keywords, and records relevant opportunities in an Excel workbook.

## Features

- Retrieves listings from Himalayas, Remote OK, and Arbeitnow.
- Scores listings using configurable role and skill keywords.
- Filters out duplicates and lower-relevance results.
- Stores results in an Excel workbook.
- Preserves existing application statuses and personal notes when refreshing listings.
- Uses public job feeds without requiring a paid API key.

## Requirements

- Python 3.11 or later
- `requests`
- `openpyxl`

## Setup

1. Clone or download this repository.
2. Open a terminal in the project directory.
3. Create and activate a virtual environment:

   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

4. Install dependencies:

   ```powershell
   python -m pip install -r requirements.txt
   ```

5. Review `config.json` and adjust the role keywords, skill keywords, matching threshold, and output filename to suit your needs.
6. Run the agent:

   ```powershell
   python .\job_agent.py
   ```

The results are written to the configured Excel workbook, normally `job_tracker.xlsx`.

## Important limitations

- The available sources do not cover every job opportunity.
- Remote listings may restrict applicants by country or location.
- Match scores and fresher-eligibility assessments are keyword-based estimates, not guarantees.
- Verify the experience requirements, location restrictions, application deadline, and job status on the employer's original listing.
- Public feeds and APIs may change or become unavailable.
- Follow the usage and attribution requirements of each source.

## Privacy

Generated job-tracking workbooks and personal application notes are intended to remain local. They are excluded from version control by the repository's `.gitignore`.

## License

No license has been specified yet. Add a license before permitting reuse or redistribution.
