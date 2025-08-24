#!/usr/bin/env python3
"""Command-line tool to tailor a base resume to a job description using an LLM."""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from openai import OpenAI


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def read_docx_text(path: str) -> str:
    """Extract text from a .docx file."""
    document = Document(path)
    return "\n".join(p.text for p in document.paragraphs)


def read_job_description(path: Optional[str], text: Optional[str]) -> str:
    """Return job description text from a path or raw string."""
    if path:
        file_path = Path(path)
        if not file_path.exists():
            raise FileNotFoundError(f"Job description file not found: {path}")
        if file_path.suffix.lower() == ".docx":
            return read_docx_text(str(file_path))
        return file_path.read_text(encoding="utf-8")
    if text:
        return text
    raise ValueError("Either job description path or text must be provided.")


# ---------------------------------------------------------------------------
# LLM Backends
# ---------------------------------------------------------------------------


class LLMBackend(ABC):
    """Abstract LLM backend interface."""

    @abstractmethod
    def generate(self, system_prompt: str, user_prompt: str, model: str) -> str:
        """Return the LLM's text output."""


class OpenAIBackend(LLMBackend):
    """OpenAI backend using the Responses API."""

    def generate(self, system_prompt: str, user_prompt: str, model: str) -> str:
        client = OpenAI()
        response = client.responses.create(
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
        )
        return response.output_text  # type: ignore[attr-defined]


class OllamaBackend(LLMBackend):
    """Ollama backend calling the local HTTP API."""

    def generate(self, system_prompt: str, user_prompt: str, model: str) -> str:
        url = "http://localhost:11434/api/chat"
        payload = {
            "model": model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "stream": False,
        }
        response = requests.post(url, json=payload, timeout=300)
        response.raise_for_status()
        data = response.json()
        return data.get("message", {}).get("content", "")


# ---------------------------------------------------------------------------
# Resume Rendering
# ---------------------------------------------------------------------------


def render_resume_to_docx(data: Dict[str, Any], out_path: str) -> None:
    """Render resume data to a .docx file."""
    doc = Document()

    # Set default font
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    style.paragraph_format.line_spacing = 1
    style.paragraph_format.space_after = Pt(6)

    # Header
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(data["header"])
    run.bold = True
    run.font.size = Pt(18)

    def add_section_heading(text: str) -> None:
        para = doc.add_paragraph()
        run = para.add_run(text.upper())
        run.bold = True
        run.font.size = Pt(12)

    # Executive Summary
    add_section_heading("Executive Summary")
    doc.add_paragraph(data["summary"])

    # Education
    add_section_heading("Education")
    doc.add_paragraph(data["education"])

    # Experience
    add_section_heading("Experience")
    for item in data["experience"]:
        header_parts = [item.get("title", ""), item.get("employer", "")]
        header = " - ".join(part for part in header_parts if part)
        loc = item.get("location")
        dates = item.get("dates")
        if loc:
            header += f", {loc}"
        if dates:
            header += f" | {dates}"
        doc.add_paragraph(header)
        for bullet in item.get("bullets", []):
            doc.add_paragraph(bullet, style="List Bullet")

    # Skills
    add_section_heading("Skills")
    skills: Dict[str, List[str]] = data.get("skills", {})
    for category, items in skills.items():
        doc.add_paragraph(f"{category}: {', '.join(items)}")

    # Interests (optional)
    interests = data.get("interests")
    if interests:
        add_section_heading("Interests")
        doc.add_paragraph(", ".join(interests))

    doc.save(out_path)


# ---------------------------------------------------------------------------
# JSON handling
# ---------------------------------------------------------------------------


def parse_llm_json(text: str) -> Dict[str, Any]:
    """Parse JSON from LLM output, attempting recovery if necessary."""
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start == -1 or end == -1 or start >= end:
            raise
        candidate = text[start : end + 1]
        return json.loads(candidate)


def validate_resume_data(data: Dict[str, Any]) -> None:
    required = ["header", "summary", "education", "experience", "skills"]
    for key in required:
        if key not in data:
            raise ValueError(f"Missing key: {key}")
    if not isinstance(data["experience"], list):
        raise TypeError("'experience' must be a list")
    if not isinstance(data["skills"], dict):
        raise TypeError("'skills' must be an object")


# ---------------------------------------------------------------------------
# Main CLI
# ---------------------------------------------------------------------------


def build_user_prompt(resume_text: str, company: str, role: str, location: Optional[str], jd_text: str) -> str:
    contract = (
        "Return JSON only with fields:\n"
        '{"header":"Name + contact exactly from base resume (no new info)",'
        '"summary":"3–4 line Executive Summary with JD-aligned keywords",'
        '"education":"1–3 lines max",'
        '"experience":[{"title":"...","employer":"...","location":"...","dates":"...","bullets":["...","..."]}],'
        '"skills":{"Technical":[...],"Tools":[...],"Other":[...]},'
        '"interests":["optional","list"]}'
    )
    parts = [
        f"Base resume:\n{resume_text}",
        f"Company: {company}",
        f"Role: {role}",
    ]
    if location:
        parts.append(f"Location: {location}")
    parts.append(f"Job description:\n{jd_text}")
    parts.append(contract)
    return "\n\n".join(parts)


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--base-resume", required=True, help="Path to base resume .docx")
    jd_group = parser.add_mutually_exclusive_group(required=True)
    jd_group.add_argument("--job-desc", help="Path to job description")
    jd_group.add_argument("--job-desc-text", help="Raw job description text")
    parser.add_argument("--company", required=True, help="Company name")
    parser.add_argument("--role", required=True, help="Role title")
    parser.add_argument("--location", help="Job location")
    parser.add_argument("--out", required=True, help="Output .docx path")
    parser.add_argument("--backend", choices=["openai", "ollama"], default="openai")
    parser.add_argument("--model", help="Model identifier", default=None)

    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(message)s")
    logging.info("Backend: %s", args.backend)

    resume_text = read_docx_text(args.base_resume)
    jd_text = read_job_description(args.job_desc, args.job_desc_text)

    system_prompt = (
        "You are a resume tailoring assistant for a recent mechanical engineering graduate."
    )

    model = args.model or ("gpt-5" if args.backend == "openai" else "llama3")

    user_prompt = build_user_prompt(
        resume_text, args.company, args.role, args.location, jd_text
    )

    backend: LLMBackend = OpenAIBackend() if args.backend == "openai" else OllamaBackend()
    try:
        output_text = backend.generate(system_prompt, user_prompt, model)
    except Exception as exc:
        logging.error("LLM call failed: %s", exc)
        return 1

    try:
        data = parse_llm_json(output_text)
        validate_resume_data(data)
    except Exception as exc:
        raw_path = os.path.splitext(args.out)[0] + "_RAW.txt"
        Path(raw_path).write_text(output_text, encoding="utf-8")
        logging.error(
            "Model output was not valid JSON. Raw output saved to %s. %s",
            raw_path,
            exc,
        )
        return 1

    try:
        render_resume_to_docx(data, args.out)
    except Exception as exc:
        logging.error("Failed to render document: %s", exc)
        return 1

    logging.info("Resume written to %s", args.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())

