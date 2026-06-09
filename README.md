# Interview Prep & Guide Generator

An automated tool to scrape, synthesize, and generate tailored interview preparation guides (in premium PDF format) for target companies and roles, fully customized using candidate resumes and background details.

## Features

- **Google Search Grounded Sourcing**: Queries live Google Search results via Gemini to find real-world, recent interview questions specific to targeted companies and roles.
- **Tailored Answers**: Synthesizes custom-tailored STAR narrative and technical responses using the candidate's resume (`.docx`) and background details.
- **Premium PDF Generation**: Automatically formats and builds highly structured PDFs featuring styled sections, category badges, key terms, custom typography, cover page, and page numberings.
- **Local LLM Backend Fallback**: Supports local LLM endpoints (like Ollama or PewDiePie's Odysseus) as a fallback or primary provider when Gemini APIs are rate-limited or unavailable.

## Setup & Configuration

1. Clone this repository to your local machine.
2. Install the python dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Copy `.env.example` to `.env` and fill in your API keys:
   ```bash
   cp .env.example .env
   ```
   *Note: Add your resume file as `Hrudhay_Kumar_Updated.docx` and RUC details file as `Hrudhay_RUC.docx` in the project root directory. These files are ignored by git to protect personal data.*

## Usage

Run the script directly from your terminal:

```bash
# Interactively enter target companies & roles
python interview_prep_generator.py

# Command-line shortcut
python interview_prep_generator.py --companies "PNC, Google" --role "Cybersecurity Analyst"

# Run with local LLM provider (Ollama / Odysseus)
python interview_prep_generator.py --provider local --companies "PNC" --role "Cybersecurity Analyst"
```

## Security & Privacy

This project is configured with a strict `.gitignore` to prevent any personal documents (such as `.docx`, `.pdf`, `.md` guides) or credentials (`.env`) from being uploaded to public version control hosting (e.g. GitHub).
