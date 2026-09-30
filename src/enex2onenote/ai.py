"""Use Claude to propose a trustworthy title, tags and summary for each note."""

from __future__ import annotations

import base64
import json
import os
import shutil
import subprocess
import tempfile
from dataclasses import asdict, dataclass
from html import escape
from pathlib import Path

from .content import extract_for_analysis, pdf_first_pages
from .enex import Note

DEFAULT_API_MODEL = "claude-opus-5-5"

SYSTEM_PROMPT = """\
You catalog personal notes that are being migrated from Evernote to OneNote. \
The person does not trust the original titles: many are scanner filenames \
("Scan_0012.pdf", "IMG_2231"), placeholders ("Untitled", "Note"), or simply wrong. \
Read the note content and attachments and describe what the note actually is.

Produce:
- title: specific and descriptive, at most 80 characters, so the person can find \
the note by scanning a list of titles. Lead with what the document is and who or \
what it concerns, e.g. "Comcast bill - March 2021", "Knee MRI results - Dr. Patel", \
"2019 federal tax return (Form 1040)". Include the date the document is about when \
there is one ("March 2021" or "2021-03-14"). No file names, surrounding quotes, or \
trailing period.
- tags: 2 to 6 short lowercase kebab-case tags covering the topic, the document \
type, and key organizations or people (e.g. "taxes", "receipt", "comcast"). Reuse \
tags from the existing vocabulary whenever one fits; invent a new tag only when \
none does.
- summary: one or two plain sentences describing the content.
- original_title_ok: true only if the original title is already accurate and \
specific enough that you would not improve it.
- confidence: "high", "medium", or "low". Use "low" when the content is empty, \
unreadable, or ambiguous, and base the title on whatever metadata exists.\
"""

ANALYSIS_SCHEMA = {
    "type": "object",
    "properties": {
        "title": {"type": "string"},
        "tags": {"type": "array", "items": {"type": "string"}},
        "summary": {"type": "string"},
        "original_title_ok": {"type": "boolean"},
        "confidence": {"type": "string", "enum": ["high", "medium", "low"]},
    },
    "required": ["title", "tags", "summary", "original_title_ok", "confidence"],
    "additionalProperties": False,
}


class AnalysisError(Exception):
    pass


@dataclass
class Analysis:
    title: str
    tags: list[str]
    summary: str
    original_title_ok: bool
    confidence: str
    model: str = ""

    @classmethod
    def from_dict(cls, d: dict, model: str = "") -> "Analysis":
        tags = []
        for t in d.get("tags", []):
            t = "-".join(str(t).strip().lower().lstrip("#").split())
            if t and t not in tags:
                tags.append(t)
        title = " ".join(str(d.get("title", "")).split()).strip().strip('"').rstrip(".")
        if not title:
            raise AnalysisError("Claude returned an empty title")
        return cls(
            title=title[:120],
            tags=tags[:8],
            summary=str(d.get("summary", "")).strip(),
            original_title_ok=bool(d.get("original_title_ok", False)),
            confidence=str(d.get("confidence", "low")),
            model=d.get("model", model),
        )

    def to_dict(self) -> dict:
        return asdict(self)


def build_prompt(note: Note, text: str, truncated: bool, vocabulary: list[str],
                 extra_instructions: str | None, attached_note: str = "") -> str:
    attachments = "\n".join(
        f"  <attachment name=\"{escape(r.filename)}\" type=\"{r.mime}\" bytes=\"{r.size}\"/>"
        for r in note.resources
    )
    parts = []
    if vocabulary:
        parts.append("<existing_tag_vocabulary>\n" + ", ".join(vocabulary)
                     + "\n</existing_tag_vocabulary>")
    if extra_instructions:
        parts.append(f"<additional_instructions>\n{extra_instructions}\n</additional_instructions>")
    trunc = (' truncated="true" note="Only the beginning of the content is shown."'
             if truncated else "")
    parts.append(
        "<note>\n"
        f"<original_title>{escape(note.title)}</original_title>\n"
        f"<created>{note.created.date().isoformat() if note.created else 'unknown'}</created>\n"
        f"<evernote_tags>{escape(', '.join(note.tags))}</evernote_tags>\n"
        + (f"<source_url>{escape(note.source_url)}</source_url>\n" if note.source_url else "")
        + f"<attachments>\n{attachments}\n</attachments>\n"
        f"<content{trunc}>\n{escape(text) if text else '(no extractable text)'}\n</content>\n"
        "</note>"
    )
    if attached_note:
        parts.append(attached_note)
    parts.append("Describe this note using the requested JSON format.")
    return "\n\n".join(parts)


class Analyzer:
    name = "base"

    def __init__(self, model: str | None, max_chars: int = 20000,
                 extra_instructions: str | None = None):
        self.model = model
        self.max_chars = max_chars
        self.extra_instructions = extra_instructions

    def analyze(self, note: Note, vocabulary: list[str]) -> Analysis:
        raise NotImplementedError


class NoAIAnalyzer(Analyzer):
    """Keeps Evernote titles and tags. Useful for dry runs and testing."""

    name = "none"

    def analyze(self, note: Note, vocabulary: list[str]) -> Analysis:
        return Analysis(title=note.title, tags=[t.lower() for t in note.tags], summary="",
                        original_title_ok=True, confidence="n/a", model="none")


class AnthropicAnalyzer(Analyzer):
    """Calls the Claude API directly (needs an API key from platform.claude.com)."""

    name = "api"
    max_images = 3
    max_image_bytes = 5 * 1024 * 1024
    scanned_pdf_pages = 5

    def __init__(self, model: str | None, effort: str = "low", **kw):
        super().__init__(model or DEFAULT_API_MODEL, **kw)
        import anthropic

        self.anthropic = anthropic
        self.client = anthropic.Anthropic(max_retries=5)
        self.effort = effort

    def _visual_blocks(self, extracted, text_len: int) -> list[dict]:
        """Scanned PDFs and photos carry no text, so let Claude look at them."""
        blocks: list[dict] = []
        for r in extracted.scanned_pdfs[:2]:
            data = pdf_first_pages(r.data, self.scanned_pdf_pages)
            if data and len(data) < 20 * 1024 * 1024:
                blocks.append({"type": "document", "title": r.filename, "source": {
                    "type": "base64", "media_type": "application/pdf",
                    "data": base64.standard_b64encode(data).decode()}})
        if text_len < 500:
            ok = {"image/jpeg", "image/png", "image/gif", "image/webp"}
            imgs = [r for r in extracted.images if r.mime in ok and r.size <= self.max_image_bytes]
            for r in imgs[: self.max_images]:
                blocks.append({"type": "image", "source": {
                    "type": "base64", "media_type": r.mime,
                    "data": base64.standard_b64encode(r.data).decode()}})
        return blocks

    def analyze(self, note: Note, vocabulary: list[str]) -> Analysis:
        extracted = extract_for_analysis(note, self.max_chars)
        visual = self._visual_blocks(extracted, len(extracted.text))
        prompt = build_prompt(note, extracted.text, extracted.truncated, vocabulary,
                              self.extra_instructions,
                              "The scanned attachments/images are included above." if visual else "")
        a = self.anthropic
        try:
            response = self.client.beta.messages.create(
                model=self.model,
                max_tokens=8000,
                system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": [*visual, {"type": "text", "text": prompt}]}],
                output_config={
                    "effort": self.effort,
                    "format": {"type": "json_schema", "schema": ANALYSIS_SCHEMA},
                },
                # If a safety classifier declines, re-run on Anthropic's
                # recommended fallback model instead of failing the note.
                betas=["server-side-fallback-2026-07-01"],
                fallbacks="default",
            )
        except a.BadRequestError as e:
            raise AnalysisError(f"Bad request: {e.message}") from e
        except (a.AuthenticationError, a.PermissionDeniedError) as e:
            raise SystemExit(f"Claude API rejected the credentials: {e.message}") from e
        except a.RateLimitError as e:
            raise AnalysisError("Rate limited after retries; re-run later") from e
        except a.APIStatusError as e:
            raise AnalysisError(f"API error {e.status_code}: {e.message}") from e
        except a.APIConnectionError as e:
            raise AnalysisError(f"Network error: {e}") from e

        if response.stop_reason == "refusal":
            raise AnalysisError("Claude declined to analyze this note")
        if response.stop_reason == "max_tokens":
            raise AnalysisError("Response was cut off (max_tokens)")
        text = next((b.text for b in response.content if b.type == "text"), None)
        if not text:
            raise AnalysisError("No text in Claude's response")
        try:
            return Analysis.from_dict(json.loads(text), model=response.model)
        except json.JSONDecodeError as e:
            raise AnalysisError(f"Could not parse Claude's JSON: {e}") from e


class ClaudeCodeAnalyzer(Analyzer):
    """Runs `claude -p` so analysis uses your Claude Pro/Max subscription login."""

    name = "claude-code"

    def __init__(self, model: str | None, effort: str | None = None, timeout: int = 600, **kw):
        super().__init__(model, **kw)
        self.effort = effort
        self.timeout = timeout
        self.exe = shutil.which("claude")
        if not self.exe:
            raise SystemExit("The `claude` command was not found. Install Claude Code "
                             "(https://claude.com/claude-code) and run `claude` once to log in.")

    def analyze(self, note: Note, vocabulary: list[str]) -> Analysis:
        extracted = extract_for_analysis(note, self.max_chars)
        with tempfile.TemporaryDirectory(prefix="enex2onenote-") as tmp:
            visuals = []
            for r in extracted.scanned_pdfs[:2] + (
                    extracted.images[:3] if len(extracted.text) < 500 else []):
                safe = "".join(c if c.isalnum() or c in "._-" else "_" for c in r.filename)
                path = Path(tmp) / f"{len(visuals)}-{safe}"
                path.write_bytes(r.data)
                visuals.append(path)
            hint = ""
            if visuals:
                hint = ("Some attachments have no extractable text. Use the Read tool to look "
                        "at these files (for PDFs, the first few pages are enough):\n"
                        + "\n".join(str(p) for p in visuals))
            prompt = build_prompt(note, extracted.text, extracted.truncated, vocabulary,
                                  self.extra_instructions, hint)
            cmd = [self.exe, "-p", "--output-format", "json", "--no-session-persistence",
                   "--json-schema", json.dumps(ANALYSIS_SCHEMA),
                   "--system-prompt", SYSTEM_PROMPT]
            if visuals:
                cmd += ["--tools", "Read", "--allowedTools", "Read", "--add-dir", tmp]
            else:
                cmd += ["--tools", ""]
            if self.model:
                cmd += ["--model", self.model]
            if self.effort:
                cmd += ["--effort", self.effort]
            try:
                # claude reads and writes UTF-8. Without an explicit encoding, Windows
                # would use its ANSI code page (cp1252) and fail on characters like
                # U+2011, or garble Claude's reply.
                proc = subprocess.run(cmd, input=prompt, capture_output=True, text=True,
                                      encoding="utf-8", errors="replace",
                                      cwd=tmp, timeout=self.timeout, env=os.environ.copy())
            except subprocess.TimeoutExpired as e:
                raise AnalysisError("claude timed out") from e
        try:
            out = json.loads(proc.stdout)
        except json.JSONDecodeError:
            raise AnalysisError(f"claude failed (exit {proc.returncode}): "
                                f"{(proc.stderr or proc.stdout).strip()[:500]}")
        if out.get("is_error") or out.get("subtype") != "success":
            raise AnalysisError(f"claude reported an error: {str(out.get('result'))[:500]}")
        data = out.get("structured_output")
        if data is None:
            try:
                data = json.loads(out.get("result", ""))
            except json.JSONDecodeError as e:
                raise AnalysisError("claude returned no structured output") from e
        model = next(iter(out.get("modelUsage", {}) or {}), self.model or "claude-code")
        return Analysis.from_dict(data, model=model)


def make_analyzer(backend: str, model: str | None, effort: str | None, max_chars: int,
                  extra_instructions: str | None) -> Analyzer:
    if backend == "auto":
        if os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN"):
            backend = "api"
        elif shutil.which("claude"):
            backend = "claude-code"
        else:
            raise SystemExit("No Claude backend available. Set ANTHROPIC_API_KEY, or install "
                             "Claude Code and log in with your Claude subscription.")
    kw = {"max_chars": max_chars, "extra_instructions": extra_instructions}
    if backend == "api":
        return AnthropicAnalyzer(model, effort=effort or "low", **kw)
    if backend == "claude-code":
        return ClaudeCodeAnalyzer(model, effort=effort, **kw)
    if backend == "none":
        return NoAIAnalyzer(model, **kw)
    raise SystemExit(f"Unknown backend: {backend}")
