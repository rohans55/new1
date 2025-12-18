#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Document comparator (DOCX/PDF).

Goals:
- Detect missing sections (based on headings)
- Detect missing/changed sentences within matching sections
- Ignore differences due to: whitespace, case, punctuation quirks, smart quotes/dashes,
  ligatures, soft hyphens, PDF hyphenation across line breaks

Output:
- Excel file with sheets: Sections, SentenceDiffs

Dependencies:
- pandas, openpyxl
- python-docx (for .docx)
- PyPDF2 (for .pdf)
"""

import os
import re
import argparse
import unicodedata
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import pandas as pd

try:
    from docx import Document as DocxDocument
except Exception:
    DocxDocument = None

try:
    from PyPDF2 import PdfReader
except Exception:
    PdfReader = None


# ---------------- Normalization ----------------
_LIGATURES = {
    "ﬁ": "fi",
    "ﬂ": "fl",
    "ﬀ": "ff",
    "ﬃ": "ffi",
    "ﬄ": "ffl",
}
_DASHES = "\u2010\u2011\u2012\u2013\u2014\u2212"  # hyphen..minus
_QUOTES = {
    "\u2018": "'",
    "\u2019": "'",
    "\u201c": '"',
    "\u201d": '"',
}
_SOFT_HYPHEN = "\u00ad"


def norm_ws(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()


def normalize_unicode(s: str) -> str:
    s = (s or "")
    s = unicodedata.normalize("NFKC", s)
    s = s.replace(_SOFT_HYPHEN, "")
    for k, v in _LIGATURES.items():
        s = s.replace(k, v)
    for k, v in _QUOTES.items():
        s = s.replace(k, v)
    s = re.sub(f"[{_DASHES}]", "-", s)
    return s


def dehyphenate_linebreaks(s: str) -> str:
    """Turn PDF artifacts like 'inter-\n national' into 'international'."""
    s = (s or "")
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"(\w)-\s*\n\s*(\w)", r"\1\2", s)
    s = re.sub(r"\n+", " ", s)
    return s


def canonical_text(s: str) -> str:
    s = normalize_unicode(s)
    s = dehyphenate_linebreaks(s)
    s = norm_ws(s).casefold()
    return s


def strip_leading_numbering(s: str) -> str:
    t = (s or "").strip()
    # a) / (a).
    t = re.sub(r"^\(?\s*[a-zA-Z]\s*\)?[)\.:\-]\s+", "", t)
    # 1. / 1.2.3)
    t = re.sub(r"^\s*\d+(?:[\.\-:\)]\d+)*[)\.:\-]\s+", "", t)
    # bullets
    t = re.sub(r"^\s*[-•*]\s+", "", t)
    return t


def sentence_tokens(sentence: str) -> List[str]:
    """Tokenize sentence for comparison (ignores punctuation/case/whitespace)."""
    s = canonical_text(sentence)
    s = strip_leading_numbering(s)
    return re.findall(r"[a-z0-9]+", s)


def token_jaccard(a_tokens: List[str], b_tokens: List[str]) -> float:
    A = set(a_tokens)
    B = set(b_tokens)
    if not A and not B:
        return 1.0
    if not A or not B:
        return 0.0
    return len(A & B) / len(A | B)


def split_sentences(text: str) -> List[str]:
    """Conservative splitter suitable for legal/contract text."""
    t = canonical_text(text)
    if not t:
        return []
    # split at sentence end punctuation OR double-space fragments common in PDFs
    parts = re.split(r"(?<=[.!?;])\s+|\s{2,}", t)
    parts = [p.strip() for p in parts if p and p.strip()]
    return parts


# ---------------- Section keying ----------------

def heading_key(heading: str) -> str:
    """Flexible key so '1', '1.', '1)' match; ANNEX/APPENDIX normalize."""
    h = norm_ws(normalize_unicode(heading))
    low = h.casefold()

    if low.startswith(("annex", "appendix")):
        return re.sub(r"\s+", " ", re.sub(r"[^\w\s]", " ", low)).strip()

    m = re.match(r"^([\d]+(?:[.\-:\)\s]+[\d]+)*)\s*(.*)$", h)
    num_key = ""
    rest = h
    if m:
        nums = re.findall(r"\d+", m.group(1))
        rest = m.group(2)
        if nums:
            num_key = "n:" + "-".join(nums)

    rest_key = re.sub(r"\s+", " ", re.sub(r"[^\w\s]", " ", rest.casefold())).strip()
    return f"{num_key}|{rest_key}" if num_key else rest_key


# ---------------- Data structures ----------------
@dataclass
class Section:
    key: str
    heading: str
    content: str


# ---------------- Extractors ----------------

def extract_docx_sections(path: str) -> List[Section]:
    if DocxDocument is None:
        raise RuntimeError("Missing dependency: python-docx (pip install python-docx)")

    doc = DocxDocument(path)
    sections: List[Section] = []

    cur_heading: Optional[str] = None
    cur_key: Optional[str] = None
    cur_lines: List[str] = []

    def push() -> None:
        nonlocal cur_heading, cur_key, cur_lines
        if cur_key is None:
            return
        sections.append(Section(key=cur_key, heading=cur_heading or "", content="\n".join(cur_lines).strip()))

    for p in doc.paragraphs:
        txt = (p.text or "").strip()
        style_name = getattr(p.style, "name", "") or ""
        is_heading = bool(txt) and ("Heading" in style_name)

        if is_heading:
            push()
            cur_heading = norm_ws(txt)
            cur_key = heading_key(cur_heading)
            cur_lines = []
        else:
            if cur_key is None:
                cur_heading = "WHOLE_DOCUMENT"
                cur_key = heading_key(cur_heading)
                cur_lines = []
            if txt:
                cur_lines.append(txt)

    push()
    return sections


STOPWORDS_CAPS = {"AND", "BETWEEN", "AGREEMENT"}
HEADER_SKIP_RX = [
    r"^\s*Page\s+\d+(\s+of\s+\d+)?\s*$",
    r"^SAT\s*-\s*Projects",
]


def is_pdf_heading(line: str) -> bool:
    t = line.strip()
    if re.match(r"^\s*\d+(?:[.\-:\)\s]+\d+)*[)\.\-:\s]+\S+", t):
        return True
    if re.match(r"^[A-Z0-9][A-Z0-9\s\-_\/]{2,80}$", t):
        if (
            len(t) >= 8
            and len(t.split()) >= 2
            and t not in STOPWORDS_CAPS
            and "confidential" not in t.casefold()
        ):
            return True
    if re.match(r"^\s*(Clause|Section|Article|Chapter)\s+\d+", t, re.I):
        return True
    if t.casefold().startswith(("annex", "appendix")):
        return True
    return False


def extract_pdf_sections(path: str) -> List[Section]:
    if PdfReader is None:
        raise RuntimeError("Missing dependency: PyPDF2 (pip install PyPDF2)")

    reader = PdfReader(path)

    lines: List[str] = []
    for page in reader.pages:
        txt = page.extract_text() or ""
        txt = normalize_unicode(txt)
        for raw in txt.splitlines():
            s = (raw or "").strip()
            if not s:
                continue
            if any(re.match(rx, s) for rx in HEADER_SKIP_RX):
                continue
            if "confidential" in s.casefold():
                continue
            lines.append(s)

    sections: List[Section] = []
    cur_heading: Optional[str] = None
    cur_key: Optional[str] = None
    cur_lines: List[str] = []

    def push() -> None:
        nonlocal cur_heading, cur_key, cur_lines
        if cur_key is None:
            return
        sections.append(Section(key=cur_key, heading=cur_heading or "", content="\n".join(cur_lines).strip()))

    for line in lines:
        if is_pdf_heading(line):
            push()
            cur_heading = norm_ws(line)
            cur_key = heading_key(cur_heading)
            cur_lines = []
        else:
            if cur_key is None:
                cur_heading = "WHOLE_DOCUMENT"
                cur_key = heading_key(cur_heading)
                cur_lines = []
            cur_lines.append(line)

    push()
    return sections


def extract_sections(path: str) -> List[Section]:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return extract_docx_sections(path)
    if ext == ".pdf":
        return extract_pdf_sections(path)
    raise ValueError(f"Unsupported file type: {ext} (use .docx or .pdf)")


# ---------------- Sentence matching ----------------

def greedy_sentence_match(
    a_sents: List[str],
    b_sents: List[str],
    same_threshold: float,
    match_threshold: float,
) -> Tuple[bool, List[Dict[str, str]]]:
    """Match sentences using token-Jaccard similarity."""

    a_tokens = [sentence_tokens(s) for s in a_sents]
    b_tokens = [sentence_tokens(s) for s in b_sents]

    used_b = set()
    diffs: List[Dict[str, str]] = []
    any_changed = False

    for a_raw, a_tok in zip(a_sents, a_tokens):
        best_j = None
        best_sim = -1.0
        for j, b_tok in enumerate(b_tokens):
            if j in used_b:
                continue
            sim = token_jaccard(a_tok, b_tok)
            if sim > best_sim:
                best_sim = sim
                best_j = j

        if best_j is None or best_sim < match_threshold:
            any_changed = True
            diffs.append(
                {
                    "ChangeType": "Missing in Doc2",
                    "Doc1 Sentence": a_raw,
                    "Doc2 Sentence": "",
                    "Similarity": f"{best_sim:.3f}" if best_sim >= 0 else "",
                }
            )
            continue

        used_b.add(best_j)
        b_raw = b_sents[best_j]

        if best_sim < same_threshold:
            any_changed = True
            diffs.append(
                {
                    "ChangeType": "Changed",
                    "Doc1 Sentence": a_raw,
                    "Doc2 Sentence": b_raw,
                    "Similarity": f"{best_sim:.3f}",
                }
            )

    for j, b_raw in enumerate(b_sents):
        if j not in used_b:
            any_changed = True
            diffs.append(
                {
                    "ChangeType": "Missing in Doc1",
                    "Doc1 Sentence": "",
                    "Doc2 Sentence": b_raw,
                    "Similarity": "",
                }
            )

    return any_changed, diffs


# ---------------- Compare + export ----------------

def sections_to_map(sections: List[Section]) -> Dict[str, Section]:
    m: Dict[str, Section] = {}
    for s in sections:
        if s.key not in m:
            m[s.key] = s
    return m


def compare_documents(doc1_path: str, doc2_path: str, out_xlsx: str, same_threshold: float, match_threshold: float) -> None:
    s1 = extract_sections(doc1_path)
    s2 = extract_sections(doc2_path)

    m1 = sections_to_map(s1)
    m2 = sections_to_map(s2)

    keys = sorted(set(m1.keys()) | set(m2.keys()))
    section_rows: List[Dict[str, str]] = []
    sentence_rows: List[Dict[str, str]] = []

    for k in keys:
        a = m1.get(k)
        b = m2.get(k)

        if a is None:
            section_rows.append(
                {
                    "SectionKey": k,
                    "Doc1 Heading": "",
                    "Doc2 Heading": b.heading if b else "",
                    "Status": "Section missing in Doc1",
                }
            )
            continue

        if b is None:
            section_rows.append(
                {
                    "SectionKey": k,
                    "Doc1 Heading": a.heading,
                    "Doc2 Heading": "",
                    "Status": "Section missing in Doc2",
                }
            )
            continue

        a_sents = split_sentences(a.content)
        b_sents = split_sentences(b.content)

        changed, diffs = greedy_sentence_match(
            a_sents,
            b_sents,
            same_threshold=same_threshold,
            match_threshold=match_threshold,
        )

        section_rows.append(
            {
                "SectionKey": k,
                "Doc1 Heading": a.heading,
                "Doc2 Heading": b.heading,
                "Status": "Changed" if changed else "Same",
            }
        )

        for d in diffs:
            sentence_rows.append(
                {
                    "SectionKey": k,
                    "Doc1 Heading": a.heading,
                    "Doc2 Heading": b.heading,
                    **d,
                }
            )

    os.makedirs(os.path.dirname(out_xlsx) or ".", exist_ok=True)
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as w:
        pd.DataFrame(section_rows).to_excel(w, index=False, sheet_name="Sections")
        pd.DataFrame(sentence_rows).to_excel(w, index=False, sheet_name="SentenceDiffs")

    print(f"Done. Report written to: {out_xlsx}")


# ---------------- CLI ----------------

def main() -> None:
    ap = argparse.ArgumentParser(
        description=(
            "Compare two documents (.docx/.pdf): missing sections + missing/changed sentences. "
            "Ignores case/whitespace/formatting and common PDF extraction artifacts."
        )
    )
    ap.add_argument("doc1")
    ap.add_argument("doc2")
    ap.add_argument("-o", "--output", default="comparison_report.xlsx")
    ap.add_argument(
        "--same-threshold",
        type=float,
        default=0.98,
        help="Sentence considered SAME if token-similarity >= this (default: 0.98)",
    )
    ap.add_argument(
        "--match-threshold",
        type=float,
        default=0.85,
        help="Sentence considered a MATCH (else missing) if token-similarity >= this (default: 0.85)",
    )

    args = ap.parse_args()

    compare_documents(
        args.doc1,
        args.doc2,
        args.output,
        same_threshold=args.same_threshold,
        match_threshold=args.match_threshold,
    )


if __name__ == "__main__":
    main()
