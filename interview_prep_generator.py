#!/usr/bin/env python3
"""
Interview Prep Generator
========================
Scrapes interview questions for target companies using Google Search via Gemini,
then generates tailored PDF prep guides based on the candidate's resume and RUC.

Usage:
    python interview_prep_generator.py
    python interview_prep_generator.py --companies "PNC, Google" --role "SOC Analyst"
"""

import argparse
import os
import sys
import re
import textwrap
import time
from datetime import datetime
import json
import urllib.request
import urllib.error

from dotenv import load_dotenv
import docx
from google import genai
from google.genai import types

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, KeepTogether,
)

# ──────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ──────────────────────────────────────────────────────────────────────────────

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
RESUME_FILE = os.path.join(SCRIPT_DIR, "Hrudhay_Kumar_Updated.docx")
RUC_FILE = os.path.join(SCRIPT_DIR, "Hrudhay_RUC.docx")
CANDIDATE_NAME = "Hrudhay Kumar"

# Color palette – deep professional tones
CLR_PRIMARY      = colors.HexColor("#0D1B2A")   # Dark navy
CLR_SECONDARY    = colors.HexColor("#1B2838")   # Slate
CLR_ACCENT       = colors.HexColor("#415A77")   # Steel blue
CLR_HIGHLIGHT    = colors.HexColor("#E63946")   # Coral-red accent
CLR_GOLD         = colors.HexColor("#D4A843")   # Gold accent
CLR_LIGHT_BG     = colors.HexColor("#F1F3F5")   # Very light gray
CLR_QUESTION_BG  = colors.HexColor("#E8ECF1")   # Question card bg
CLR_ANSWER_BG    = colors.HexColor("#FFFFFF")   # White
CLR_TEXT_DARK     = colors.HexColor("#1A1A2E")   # Nearly black
CLR_TEXT_BODY     = colors.HexColor("#2C3E50")   # Body text
CLR_TEXT_MUTED    = colors.HexColor("#6C757D")   # Muted text
CLR_DIVIDER       = colors.HexColor("#CED4DA")   # Light divider

CATEGORY_COLORS = {
    "Technical":        colors.HexColor("#0F3460"),
    "Behavioral":       colors.HexColor("#E63946"),
    "Situational":      colors.HexColor("#6A0572"),
    "Company-Specific": colors.HexColor("#2B7A78"),
    "General":          colors.HexColor("#415A77"),
}


# ──────────────────────────────────────────────────────────────────────────────
# PARAGRAPH STYLES
# ──────────────────────────────────────────────────────────────────────────────

def build_styles():
    """Return a dict of custom ParagraphStyles."""
    return {
        "cover_title": ParagraphStyle(
            "cover_title",
            fontName="Helvetica-Bold",
            fontSize=28,
            leading=34,
            textColor=CLR_PRIMARY,
            alignment=TA_CENTER,
            spaceAfter=6,
        ),
        "cover_subtitle": ParagraphStyle(
            "cover_subtitle",
            fontName="Helvetica",
            fontSize=16,
            leading=22,
            textColor=CLR_ACCENT,
            alignment=TA_CENTER,
            spaceAfter=4,
        ),
        "cover_meta": ParagraphStyle(
            "cover_meta",
            fontName="Helvetica",
            fontSize=11,
            leading=16,
            textColor=CLR_TEXT_MUTED,
            alignment=TA_CENTER,
            spaceAfter=2,
        ),
        "section_header": ParagraphStyle(
            "section_header",
            fontName="Helvetica-Bold",
            fontSize=18,
            leading=24,
            textColor=CLR_PRIMARY,
            spaceBefore=20,
            spaceAfter=10,
        ),
        "category_badge": ParagraphStyle(
            "category_badge",
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=10,
            textColor=colors.white,
            alignment=TA_CENTER,
        ),
        "question_num": ParagraphStyle(
            "question_num",
            fontName="Helvetica-Bold",
            fontSize=10,
            leading=14,
            textColor=CLR_HIGHLIGHT,
        ),
        "question_text": ParagraphStyle(
            "question_text",
            fontName="Helvetica-Bold",
            fontSize=11,
            leading=16,
            textColor=CLR_TEXT_DARK,
            spaceAfter=4,
        ),
        "answer_label": ParagraphStyle(
            "answer_label",
            fontName="Helvetica-Bold",
            fontSize=9,
            leading=13,
            textColor=CLR_ACCENT,
            spaceBefore=4,
        ),
        "answer_text": ParagraphStyle(
            "answer_text",
            fontName="Helvetica",
            fontSize=10,
            leading=15,
            textColor=CLR_TEXT_BODY,
            alignment=TA_JUSTIFY,
        ),
        "key_terms_label": ParagraphStyle(
            "key_terms_label",
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=12,
            textColor=CLR_GOLD,
            spaceBefore=6,
        ),
        "key_terms_text": ParagraphStyle(
            "key_terms_text",
            fontName="Helvetica-Oblique",
            fontSize=8,
            leading=12,
            textColor=CLR_TEXT_MUTED,
        ),
        "footer_text": ParagraphStyle(
            "footer_text",
            fontName="Helvetica",
            fontSize=7,
            leading=10,
            textColor=CLR_TEXT_MUTED,
            alignment=TA_CENTER,
        ),
        "summary_text": ParagraphStyle(
            "summary_text",
            fontName="Helvetica",
            fontSize=10,
            leading=15,
            textColor=CLR_TEXT_BODY,
            alignment=TA_JUSTIFY,
            spaceBefore=4,
            spaceAfter=4,
        ),
        "toc_item": ParagraphStyle(
            "toc_item",
            fontName="Helvetica",
            fontSize=10,
            leading=18,
            textColor=CLR_TEXT_BODY,
            leftIndent=12,
        ),
    }


# ──────────────────────────────────────────────────────────────────────────────
# DOCUMENT HELPERS
# ──────────────────────────────────────────────────────────────────────────────

def load_docx(filepath):
    """Extract all non-empty paragraph text from a .docx file."""
    if not os.path.exists(filepath):
        print(f"  ⚠  File not found: {filepath}")
        return ""
    doc = docx.Document(filepath)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())


def get_api_key():
    """Resolve the Gemini API key from environment, .env, or interactive prompt."""
    load_dotenv(os.path.join(SCRIPT_DIR, ".env"))
    key = os.getenv("GEMINI_API_KEY")
    if key:
        return key
    print("\n┌─────────────────────────────────────────────┐")
    print("│  No GEMINI_API_KEY found in environment.    │")
    print("│  You can set it in a .env file or export    │")
    print("│  it as an environment variable.             │")
    print("└─────────────────────────────────────────────┘")
    key = input("  🔑  Enter your Gemini API key: ").strip()
    if not key:
        print("  ❌  No API key provided. Exiting.")
        sys.exit(1)
    return key


def sanitize_filename(name):
    """Turn an arbitrary string into a safe filename fragment."""
    return re.sub(r'[^\w\-]+', '_', name).strip('_')


# ──────────────────────────────────────────────────────────────────────────────
# GEMINI  –  RETRY HELPER
# ──────────────────────────────────────────────────────────────────────────────

MAX_RETRIES = 5
BASE_BACKOFF = 15  # seconds

def _gemini_call_with_retry(client, model, contents, config, label="API call"):
    """Wrap a Gemini generate_content call with retry + exponential backoff."""
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = client.models.generate_content(
                model=model,
                contents=contents,
                config=config,
            )
            return response.text
        except Exception as e:
            err_str = str(e)
            is_retryable = "429" in err_str or "503" in err_str or "RESOURCE_EXHAUSTED" in err_str or "UNAVAILABLE" in err_str

            if not is_retryable or attempt == MAX_RETRIES:
                raise

            # Try to parse retry delay from error message
            delay_match = re.search(r'retry\s+in\s+([\d.]+)s', err_str, re.IGNORECASE)
            if delay_match:
                wait = float(delay_match.group(1)) + 2  # add small buffer
            else:
                wait = BASE_BACKOFF * (2 ** (attempt - 1))  # exponential backoff

            print(f"  ⏳  {label}: Rate limited (attempt {attempt}/{MAX_RETRIES}). "
                  f"Waiting {wait:.0f}s …")
            time.sleep(wait)

    # Should never reach here, but just in case
    raise RuntimeError(f"Failed after {MAX_RETRIES} retries")


def _local_call(endpoint, model, prompt):
    """Call a local OpenAI-compatible API endpoint (like Ollama or Odysseus)."""
    endpoint = endpoint.rstrip("/")
    url = f"{endpoint}/chat/completions"
    
    headers = {
        "Content-Type": "application/json",
    }
    
    data = {
        "model": model,
        "messages": [
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.4,
    }
    
    try:
        req = urllib.request.Request(
            url,
            data=json.dumps(data).encode("utf-8"),
            headers=headers,
            method="POST"
        )
        print(f"  🤖  Sending request to local LLM ({model}) at {url} …")
        with urllib.request.urlopen(req, timeout=300) as response:
            res_data = json.loads(response.read().decode("utf-8"))
            return res_data["choices"][0]["message"]["content"]
    except Exception as e:
        print(f"  ❌  Local LLM error: {e}")
        raise


# ──────────────────────────────────────────────────────────────────────────────
# GEMINI  –  SEARCH & GENERATE
# ──────────────────────────────────────────────────────────────────────────────

def fetch_interview_questions(client, company, role):
    """Use Gemini with Google Search grounding to find real interview Qs."""
    prompt = textwrap.dedent(f"""\
        Search for the most common and recent interview questions asked at
        **{company}** for the role of **{role}**.

        Find 15–20 specific interview questions including:
        • Technical / domain-specific questions
        • Behavioral / situational questions (STAR format)
        • Company-specific questions unique to {company}
        • Common screening questions for this role

        For each question, prefix it with its category in square brackets.
        Example:
        [Technical] What is the difference between IDS and IPS?
        [Behavioral] Describe a time you handled a critical incident under pressure.
    """)

    print(f"  🔍  Searching for {company} interview questions …")
    return _gemini_call_with_retry(
        client,
        model="gemini-2.5-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            tools=[types.Tool(google_search=types.GoogleSearch())],
            temperature=0.7,
        ),
        label=f"Search ({company})",
    )


def generate_tailored_answers(client, company, role, questions_text, resume_text, ruc_text, provider="gemini", local_endpoint=None, local_model=None):
    """Ask Gemini or a local LLM to produce tailored answers grounded in the candidate's background."""
    # Truncate RUC to stay within context limits
    ruc_excerpt = ruc_text[:12000]

    prompt = textwrap.dedent(f"""\
        You are an elite career coach preparing a cybersecurity professional for
        an interview at **{company}** for the role of **{role}**.

        ─── CANDIDATE RESUME ───
        {resume_text}

        ─── CANDIDATE DETAILED BACKGROUND (RUC) ───
        {ruc_excerpt}

        ─── INTERVIEW QUESTIONS FOUND FOR {company.upper()} ───
        {questions_text}

        ═══════════════════════════════════════════════════════════════
        INSTRUCTIONS
        ═══════════════════════════════════════════════════════════════
        For EVERY question above, produce a detailed, tailored answer that:
        1. Directly references the candidate's ACTUAL experience, tools,
           metrics, and achievements from the resume and RUC.
        2. Uses STAR format (Situation → Task → Action → Result) for
           behavioral and situational questions.
        3. Includes specific technical details, tool names, and metrics.
        4. Is confident but honest — only mention things the candidate
           has actually done.
        5. For technical questions: 3–6 sentences with depth.
           For behavioral questions: full STAR narrative.

        FORMAT — output EXACTLY this structure for each Q/A pair:

        ###QUESTION_START###
        CATEGORY: <Technical | Behavioral | Situational | Company-Specific>
        QUESTION: <The question text>
        ANSWER: <Your tailored answer>
        KEY_TERMS: <Comma-separated list of 3–5 key technical terms>
        ###QUESTION_END###

        Generate answers for ALL questions.
    """)

    print(f"  ✍️   Generating tailored answers for {company} using {provider} …")
    
    if provider == "local":
        if not local_endpoint:
            local_endpoint = "http://localhost:11434/v1"
        if not local_model:
            local_model = "llama3"
        return _local_call(local_endpoint, local_model, prompt)

    try:
        return _gemini_call_with_retry(
            client,
            model="gemini-2.5-flash",
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.4,
                max_output_tokens=30000,
            ),
            label=f"Answers ({company})",
        )
    except Exception as e:
        print(f"  ⚠  Gemini answer generation failed: {e}")
        print(f"  🔄  Attempting automatic fallback to local LLM ({local_model}) …")
        try:
            if not local_endpoint:
                local_endpoint = "http://localhost:11434/v1"
            if not local_model:
                local_model = "llama3"
            return _local_call(local_endpoint, local_model, prompt)
        except Exception as local_err:
            print(f"  ❌  Fallback to local LLM failed: {local_err}")
            raise e


def parse_qa_response(raw_text):
    """Parse the ###QUESTION_START### / ###QUESTION_END### blocks."""
    qa_pairs = []
    blocks = raw_text.split("###QUESTION_START###")

    for block in blocks:
        if "###QUESTION_END###" not in block:
            continue
        content = block.split("###QUESTION_END###")[0].strip()

        qa = {}
        cat = re.search(r"CATEGORY:\s*(.+?)(?:\n|$)", content)
        qa["category"] = cat.group(1).strip() if cat else "General"

        q = re.search(r"QUESTION:\s*(.+?)(?=\nANSWER:)", content, re.DOTALL)
        qa["question"] = q.group(1).strip() if q else ""

        a = re.search(r"ANSWER:\s*(.+?)(?=\nKEY_TERMS:|$)", content, re.DOTALL)
        qa["answer"] = a.group(1).strip() if a else ""

        kt = re.search(r"KEY_TERMS:\s*(.+?)$", content, re.DOTALL)
        qa["key_terms"] = kt.group(1).strip() if kt else ""

        if qa["question"] and qa["answer"]:
            qa_pairs.append(qa)

    return qa_pairs


# ──────────────────────────────────────────────────────────────────────────────
# PDF GENERATION
# ──────────────────────────────────────────────────────────────────────────────

def _header_footer(canvas, doc, company, role):
    """Draw header and footer on every page (except page 1 cover)."""
    page_num = doc.page
    canvas.saveState()

    if page_num > 1:
        # ── Header line ──
        canvas.setStrokeColor(CLR_DIVIDER)
        canvas.setLineWidth(0.5)
        canvas.line(
            doc.leftMargin,
            letter[1] - 0.55 * inch,
            letter[0] - doc.rightMargin,
            letter[1] - 0.55 * inch,
        )
        canvas.setFont("Helvetica-Bold", 8)
        canvas.setFillColor(CLR_ACCENT)
        canvas.drawString(
            doc.leftMargin,
            letter[1] - 0.48 * inch,
            f"{company}  ·  {role}  ·  Interview Preparation Guide",
        )

    # ── Footer ──
    canvas.setStrokeColor(CLR_DIVIDER)
    canvas.setLineWidth(0.5)
    canvas.line(
        doc.leftMargin,
        0.55 * inch,
        letter[0] - doc.rightMargin,
        0.55 * inch,
    )
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(CLR_TEXT_MUTED)
    canvas.drawString(
        doc.leftMargin,
        0.40 * inch,
        f"Prepared for {CANDIDATE_NAME}  ·  {datetime.now().strftime('%B %d, %Y')}",
    )
    canvas.drawRightString(
        letter[0] - doc.rightMargin,
        0.40 * inch,
        f"Page {page_num}",
    )

    canvas.restoreState()


def _build_cover_page(story, sty, company, role):
    """Append cover-page flowables to *story*."""
    story.append(Spacer(1, 1.6 * inch))

    # Decorative top rule
    story.append(HRFlowable(
        width="60%", thickness=3, color=CLR_HIGHLIGHT,
        spaceAfter=14, spaceBefore=0, hAlign="CENTER",
    ))

    story.append(Paragraph("INTERVIEW PREPARATION GUIDE", sty["cover_meta"]))
    story.append(Spacer(1, 10))
    story.append(Paragraph(company.upper(), sty["cover_title"]))
    story.append(Spacer(1, 6))
    story.append(Paragraph(role, sty["cover_subtitle"]))
    story.append(Spacer(1, 20))

    # Decorative bottom rule
    story.append(HRFlowable(
        width="40%", thickness=1, color=CLR_ACCENT,
        spaceAfter=20, spaceBefore=0, hAlign="CENTER",
    ))

    story.append(Paragraph(
        f"Prepared for <b>{CANDIDATE_NAME}</b>", sty["cover_meta"],
    ))
    story.append(Paragraph(
        datetime.now().strftime("%B %d, %Y"), sty["cover_meta"],
    ))

    story.append(Spacer(1, 1.2 * inch))

    # Disclaimer
    disclaimer = (
        "This document was auto-generated using AI-powered search and your "
        "personal resume &amp; background details. Review each answer to ensure "
        "it accurately reflects your experience before your interview."
    )
    story.append(Paragraph(disclaimer, ParagraphStyle(
        "disclaimer",
        fontName="Helvetica-Oblique",
        fontSize=8,
        leading=12,
        textColor=CLR_TEXT_MUTED,
        alignment=TA_CENTER,
    )))

    story.append(PageBreak())


def _build_summary_page(story, sty, company, role, qa_pairs):
    """Append a summary / table-of-contents page."""
    story.append(Paragraph("📋 &nbsp;Overview", sty["section_header"]))
    story.append(HRFlowable(
        width="100%", thickness=1, color=CLR_DIVIDER,
        spaceAfter=12, spaceBefore=0,
    ))

    # Category counts
    counts = {}
    for qa in qa_pairs:
        cat = qa["category"]
        counts[cat] = counts.get(cat, 0) + 1

    summary_lines = [
        f"This guide contains <b>{len(qa_pairs)}</b> interview questions "
        f"sourced for <b>{company}</b> targeting the <b>{role}</b> role.",
        "Each answer is tailored to your resume, projects, and professional background.",
    ]
    for line in summary_lines:
        story.append(Paragraph(line, sty["summary_text"]))

    story.append(Spacer(1, 12))

    # Category breakdown table
    table_data = [
        [Paragraph("<b>Category</b>", sty["answer_label"]),
         Paragraph("<b>Count</b>", sty["answer_label"])],
    ]
    for cat, cnt in sorted(counts.items(), key=lambda x: -x[1]):
        clr = CATEGORY_COLORS.get(cat, CLR_ACCENT)
        badge = f'<font color="{clr.hexval()}">{cat}</font>'
        table_data.append([
            Paragraph(badge, sty["answer_text"]),
            Paragraph(str(cnt), sty["answer_text"]),
        ])

    cat_table = Table(table_data, colWidths=[3.5 * inch, 1.2 * inch])
    cat_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), CLR_LIGHT_BG),
        ("GRID", (0, 0), (-1, -1), 0.4, CLR_DIVIDER),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
    ]))
    story.append(cat_table)
    story.append(Spacer(1, 16))

    # Quick-reference list
    story.append(Paragraph("📝 &nbsp;Questions at a Glance", sty["section_header"]))
    story.append(HRFlowable(
        width="100%", thickness=1, color=CLR_DIVIDER,
        spaceAfter=10, spaceBefore=0,
    ))
    for i, qa in enumerate(qa_pairs, 1):
        story.append(Paragraph(
            f"<b>Q{i}.</b>  {qa['question'][:120]}{'…' if len(qa['question']) > 120 else ''}",
            sty["toc_item"],
        ))

    story.append(PageBreak())


def _build_qa_section(story, sty, qa_pairs):
    """Append the main Q&A section with styled question/answer cards."""
    story.append(Paragraph("💡 &nbsp;Questions &amp; Tailored Answers", sty["section_header"]))
    story.append(HRFlowable(
        width="100%", thickness=1, color=CLR_DIVIDER,
        spaceAfter=16, spaceBefore=0,
    ))

    for i, qa in enumerate(qa_pairs, 1):
        card = _build_qa_card(i, qa, sty)
        story.append(KeepTogether(card))
        story.append(Spacer(1, 14))


def _build_qa_card(index, qa, sty):
    """Return a list of flowables representing one Q/A card."""
    elements = []
    cat = qa["category"]
    cat_color = CATEGORY_COLORS.get(cat, CLR_ACCENT)

    # ── Question card (with category badge + number) ──
    badge_text = f'<font color="white">&nbsp;{cat.upper()}&nbsp;</font>'
    badge_para = Paragraph(badge_text, sty["category_badge"])

    badge_cell = Table(
        [[badge_para]],
        colWidths=[1.4 * inch],
        rowHeights=[16],
    )
    badge_cell.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), cat_color),
        ("ALIGN", (0, 0), (0, 0), "CENTER"),
        ("VALIGN", (0, 0), (0, 0), "MIDDLE"),
        ("TOPPADDING", (0, 0), (0, 0), 2),
        ("BOTTOMPADDING", (0, 0), (0, 0), 2),
        ("ROUNDEDCORNERS", [4, 4, 4, 4]),
    ]))

    elements.append(badge_cell)
    elements.append(Spacer(1, 6))

    # Question text
    q_para = Paragraph(
        f'<font color="{CLR_HIGHLIGHT.hexval()}">Q{index}.</font>&nbsp;&nbsp;{_escape(qa["question"])}',
        sty["question_text"],
    )

    # Wrap question in a shaded box
    q_table = Table(
        [[q_para]],
        colWidths=[6.5 * inch],
    )
    q_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), CLR_QUESTION_BG),
        ("TOPPADDING", (0, 0), (0, 0), 10),
        ("BOTTOMPADDING", (0, 0), (0, 0), 10),
        ("LEFTPADDING", (0, 0), (0, 0), 14),
        ("RIGHTPADDING", (0, 0), (0, 0), 14),
        ("ROUNDEDCORNERS", [6, 6, 6, 6]),
    ]))
    elements.append(q_table)
    elements.append(Spacer(1, 8))

    # ── Answer section ──
    elements.append(Paragraph("TAILORED ANSWER", sty["answer_label"]))
    elements.append(Spacer(1, 2))

    # Format answer text — handle STAR labels
    answer_html = _format_answer_html(qa["answer"])
    a_para = Paragraph(answer_html, sty["answer_text"])

    a_table = Table(
        [[a_para]],
        colWidths=[6.5 * inch],
    )
    a_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), CLR_ANSWER_BG),
        ("BOX", (0, 0), (0, 0), 0.5, CLR_DIVIDER),
        ("TOPPADDING", (0, 0), (0, 0), 10),
        ("BOTTOMPADDING", (0, 0), (0, 0), 10),
        ("LEFTPADDING", (0, 0), (0, 0), 14),
        ("RIGHTPADDING", (0, 0), (0, 0), 14),
        ("ROUNDEDCORNERS", [4, 4, 4, 4]),
    ]))
    elements.append(a_table)

    # ── Key Terms ──
    if qa.get("key_terms"):
        elements.append(Spacer(1, 4))
        elements.append(Paragraph("🔑 KEY TERMS", sty["key_terms_label"]))
        elements.append(Paragraph(_escape(qa["key_terms"]), sty["key_terms_text"]))

    # Divider between cards
    elements.append(Spacer(1, 6))
    elements.append(HRFlowable(
        width="100%", thickness=0.3, color=CLR_DIVIDER,
        spaceAfter=4, spaceBefore=0,
    ))

    return elements


def _escape(text):
    """Escape HTML special chars for ReportLab Paragraph markup."""
    text = text.replace("&", "&amp;")
    text = text.replace("<", "&lt;")
    text = text.replace(">", "&gt;")
    return text


def _format_answer_html(answer_text):
    """Highlight STAR labels and add visual structure to answer text."""
    text = _escape(answer_text)

    # Bold STAR labels
    for label in ("Situation:", "Task:", "Action:", "Result:"):
        escaped = label
        text = text.replace(
            escaped,
            f'<br/><b><font color="{CLR_ACCENT.hexval()}">{escaped}</font></b> ',
        )

    # Also handle **bold** markdown
    text = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)

    # Convert newlines
    text = text.replace("\n", "<br/>")

    return text


def generate_pdf(qa_pairs, company, role, output_path):
    """Build and save the styled PDF document."""
    sty = build_styles()

    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        rightMargin=0.75 * inch,
        leftMargin=0.75 * inch,
        topMargin=0.7 * inch,
        bottomMargin=0.7 * inch,
        title=f"{company} – {role} Interview Preparation Guide",
        author=CANDIDATE_NAME,
    )

    story = []

    # 1) Cover page
    _build_cover_page(story, sty, company, role)

    # 2) Summary / TOC
    _build_summary_page(story, sty, company, role, qa_pairs)

    # 3) Q&A cards
    _build_qa_section(story, sty, qa_pairs)

    # Build with header/footer callback
    doc.build(
        story,
        onFirstPage=lambda c, d: _header_footer(c, d, company, role),
        onLaterPages=lambda c, d: _header_footer(c, d, company, role),
    )


# ──────────────────────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Generate tailored interview prep PDFs for target companies.",
    )
    parser.add_argument(
        "--companies",
        type=str,
        default=None,
        help='Comma-separated company names, e.g. "PNC, Google, Deloitte"',
    )
    parser.add_argument(
        "--role",
        type=str,
        default=None,
        help='Target job role (default: Cybersecurity Analyst)',
    )
    parser.add_argument(
        "--output-dir",
        type=str,
        default=SCRIPT_DIR,
        help="Directory to save generated PDFs (default: script directory)",
    )
    parser.add_argument(
        "--provider",
        type=str,
        choices=["gemini", "local"],
        default="gemini",
        help="LLM provider: 'gemini' (default) or 'local' (Odysseus/Ollama)",
    )
    parser.add_argument(
        "--local-endpoint",
        type=str,
        default="http://localhost:11434/v1",
        help="Local OpenAI-compatible API endpoint (default: http://localhost:11434/v1)",
    )
    parser.add_argument(
        "--local-model",
        type=str,
        default="llama3",
        help="Model name for local provider (default: llama3)",
    )
    args = parser.parse_args()

    # ── Banner ──
    print()
    print("╔══════════════════════════════════════════════════════════╗")
    print("║       🛡️  INTERVIEW PREP GENERATOR                     ║")
    print("║       Tailored Q&A guides from your resume & RUC       ║")
    print("╚══════════════════════════════════════════════════════════╝")
    print()

    # ── Resolve companies ──
    if args.companies:
        companies = [c.strip() for c in args.companies.split(",") if c.strip()]
    else:
        raw = input("  📋  Enter company names (comma-separated): ").strip()
        if not raw:
            print("  ❌  No companies entered. Exiting.")
            sys.exit(1)
        companies = [c.strip() for c in raw.split(",") if c.strip()]

    # ── Resolve role ──
    if args.role:
        role = args.role
    else:
        role_input = input("  💼  Enter target role (press Enter for 'Cybersecurity Analyst'): ").strip()
        role = role_input if role_input else "Cybersecurity Analyst"

    print()
    print(f"  ✅  Companies : {', '.join(companies)}")
    print(f"  ✅  Role      : {role}")
    print()

    # ── Load candidate documents ──
    print("  📄  Loading resume and RUC …")
    resume_text = load_docx(RESUME_FILE)
    ruc_text = load_docx(RUC_FILE)

    if not resume_text:
        print(f"  ❌  Could not load resume from {RESUME_FILE}")
        sys.exit(1)
    if not ruc_text:
        print(f"  ⚠   RUC file not found or empty — continuing with resume only.")

    print(f"       Resume : {len(resume_text):,} chars loaded")
    print(f"       RUC    : {len(ruc_text):,} chars loaded")
    print()

    # ── Initialize Gemini client ──
    client = None
    if args.provider == "gemini":
        try:
            api_key = get_api_key()
            client = genai.Client(api_key=api_key)
            print("  🤖  Gemini client initialized.\n")
        except Exception as e:
            print(f"  ⚠  Could not initialize Gemini: {e}")
            print("  🔄  Switching provider to local (Odysseus/Ollama) as a fallback.\n")
            args.provider = "local"

    # ── Process each company ──
    generated_files = []
    os.makedirs(args.output_dir, exist_ok=True)

    for idx, company in enumerate(companies, 1):
        print(f"  {'─' * 54}")
        print(f"  [{idx}/{len(companies)}]  Processing: {company}")
        print(f"  {'─' * 54}")

        try:
            # Step 1: Search for questions
            if client:
                try:
                    questions_text = fetch_interview_questions(client, company, role)
                except Exception as e:
                    print(f"  ⚠  Gemini question search failed: {e}")
                    print("  🔄  Using default common cybersecurity/interview questions fallback …")
                    questions_text = textwrap.dedent(f"""\
                        [Technical] What is the difference between IDS and IPS?
                        [Technical] How do you secure a web application?
                        [Technical] Can you describe the difference between symmetric and asymmetric encryption?
                        [Technical] What is DNS spoofing and how do you prevent it?
                        [Behavioral] Describe a time you handled a critical incident under pressure.
                        [Behavioral] Tell me about a time you had to explain a complex technical issue to a non-technical stakeholder.
                        [Situational] What would you do if you detected an active ransomware attack on a critical server?
                        [Company-Specific] Why do you want to work at {company} as a {role}?
                    """)
            else:
                print("  🔄  No Gemini client active. Using default common interview questions fallback …")
                questions_text = textwrap.dedent(f"""\
                    [Technical] What is the difference between IDS and IPS?
                    [Technical] How do you secure a web application?
                    [Technical] Can you describe the difference between symmetric and asymmetric encryption?
                    [Technical] What is DNS spoofing and how do you prevent it?
                    [Behavioral] Describe a time you handled a critical incident under pressure.
                    [Behavioral] Tell me about a time you had to explain a complex technical issue to a non-technical stakeholder.
                    [Situational] What would you do if you detected an active ransomware attack on a critical server?
                    [Company-Specific] Why do you want to work at {company} as a {role}?
                """)

            # Step 2: Generate tailored answers
            raw_answers = generate_tailored_answers(
                client, company, role, questions_text, resume_text, ruc_text,
                provider=args.provider,
                local_endpoint=args.local_endpoint,
                local_model=args.local_model,
            )

            # Step 3: Parse response
            qa_pairs = parse_qa_response(raw_answers)

            if not qa_pairs:
                print(f"  ⚠   No Q&A pairs parsed for {company}. Skipping PDF.")
                continue

            print(f"  ✅  Parsed {len(qa_pairs)} Q&A pairs.")

            # Step 4: Generate PDF
            safe_company = sanitize_filename(company)
            safe_role = sanitize_filename(role)
            filename = f"{safe_company}_{safe_role}_Interview_Prep.pdf"
            output_path = os.path.join(args.output_dir, filename)

            generate_pdf(qa_pairs, company, role, output_path)

            generated_files.append(output_path)
            print(f"  📕  PDF saved: {filename}\n")

        except Exception as e:
            print(f"  ❌  Error processing {company}: {e}\n")

    # ── Summary ──
    print()
    print("╔══════════════════════════════════════════════════════════╗")
    print("║  ✅  GENERATION COMPLETE                                ║")
    print("╚══════════════════════════════════════════════════════════╝")
    print()
    if generated_files:
        print(f"  Generated {len(generated_files)} PDF(s):")
        for f in generated_files:
            print(f"    📄  {os.path.basename(f)}")
        print(f"\n  📂  Output directory: {args.output_dir}")
    else:
        print("  ⚠   No PDFs were generated.")
    print()


if __name__ == "__main__":
    main()
