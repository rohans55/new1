#!/usr/bin/env python3
"""Clause-by-clause document comparison utility.

This script compares two documents that contain clause-like sections, typically
introduced by headings such as "Clause 1", "Section 2.1", "1.0 Introduction",
or UPPERCASE headings. It highlights which clauses are missing from either
document and prints a diff for clauses that exist in both files but differ in
content.

The script works with plain-text files by default and optionally supports
`.docx` files if the `python-docx` package is installed (`pip install
python-docx`). See the README for usage examples.
"""

from __future__ import annotations

import argparse
import difflib
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Sequence

try:
    from docx import Document  # type: ignore
except ImportError:  # pragma: no cover - optional dependency
    Document = None  # type: ignore


@dataclass
class Clause:
    """A single clause/section extracted from a document."""

    heading: str
    body: str
    index: int  # the order in which the clause appears in the document

    def normalized_heading(self) -> str:
        """Return a case/format-insensitive token for matching clauses."""
        heading = self.heading.lower()
        heading = re.sub(r"[^a-z0-9]+", " ", heading)
        return heading.strip()

    def normalized_body(self) -> str:
        """Normalize clause body for equality checks."""
        body = self.body.strip()
        body = re.sub(r"\s+", " ", body)
        return body.lower()


ClauseMap = Dict[str, List[Clause]]


HEADING_PATTERNS: Sequence[re.Pattern[str]] = (
    re.compile(r"^(?:clause|section)\s+\d+[\w .-]*:?$", re.IGNORECASE),
    re.compile(r"^\d+(?:\.\d+)*\s+.+$"),
    re.compile(r"^[A-Z0-9][A-Z0-9 .,:;/-]{3,}$"),
)


def is_probable_heading(line: str) -> bool:
    """Return True if the line likely represents a clause heading."""
    stripped = line.strip()
    if not stripped:
        return False
    return any(pattern.match(stripped) for pattern in HEADING_PATTERNS)


def parse_clauses(text: str) -> ClauseMap:
    """Split raw text into clauses keyed by normalized heading."""
    clauses: ClauseMap = {}
    current_heading: str | None = None
    current_lines: List[str] = []
    clause_index = 0

    lines = text.splitlines()
    for i, raw_line in enumerate(lines):
        if is_probable_heading(raw_line):
            if current_heading is not None:
                _store_clause(
                    clauses,
                    Clause(
                        heading=current_heading,
                        body="\n".join(current_lines).strip(),
                        index=clause_index,
                    ),
                )
                clause_index += 1
                current_lines = []
            current_heading = raw_line.strip()
        else:
            current_lines.append(raw_line.rstrip())

    if current_heading is not None:
        _store_clause(
            clauses,
            Clause(
                heading=current_heading,
                body="\n".join(current_lines).strip(),
                index=clause_index,
            ),
        )

    return clauses


def _store_clause(clauses: ClauseMap, clause: Clause) -> None:
    """Append the clause to the correct heading bucket."""
    key = clause.normalized_heading()
    clauses.setdefault(key, []).append(clause)


def read_document(path: Path) -> str:
    """Return the textual contents of the provided document."""
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    if path.suffix.lower() == ".docx":
        if Document is None:
            raise RuntimeError(
                "python-docx is required to read .docx files. "
                "Install it with `pip install python-docx`."
            )
        doc = Document(str(path))
        return "\n".join(paragraph.text for paragraph in doc.paragraphs)

    return path.read_text(encoding="utf-8")


@dataclass
class ComparisonResult:
    missing_from_doc_two: List[Clause]
    additional_in_doc_two: List[Clause]
    changed: List[Dict[str, str]]


def compare_clauses(first: ClauseMap, second: ClauseMap) -> ComparisonResult:
    """Compare two clause maps with document one treated as the reference."""
    missing_from_doc_two: List[Clause] = []
    changed: List[Dict[str, str]] = []

    for heading_key, first_clauses in first.items():
        second_clauses = second.get(heading_key)
        if not second_clauses:
            missing_from_doc_two.extend(first_clauses)
            continue

        for idx, first_clause in enumerate(first_clauses):
            if idx >= len(second_clauses):
                missing_from_doc_two.append(first_clause)
                continue

            second_clause = second_clauses[idx]
            if first_clause.normalized_body() != second_clause.normalized_body():
                diff = "\n".join(
                    difflib.unified_diff(
                        first_clause.body.splitlines(),
                        second_clause.body.splitlines(),
                        fromfile="doc_one",
                        tofile="doc_two",
                        lineterm="",
                    )
                )
                changed.append(
                    {
                        "heading": first_clause.heading,
                        "doc_one": first_clause.body,
                        "doc_two": second_clause.body,
                        "diff": diff,
                    }
                )

    additional_in_doc_two: List[Clause] = []
    for heading_key, second_clauses in second.items():
        first_count = len(first.get(heading_key, []))
        if len(second_clauses) > first_count:
            additional_in_doc_two.extend(second_clauses[first_count:])

    return ComparisonResult(
        missing_from_doc_two=missing_from_doc_two,
        additional_in_doc_two=additional_in_doc_two,
        changed=changed,
    )


def format_result(result: ComparisonResult) -> str:
    """Create a readable, beginner-friendly summary."""
    lines: List[str] = []

    lines.append("CLAUSE COMPARISON REPORT")
    lines.append("=" * 27)
    lines.append("Document 2 vs Document 1 overview:")
    lines.append(
        f"  - Missing clauses (only in doc 1): {len(result.missing_from_doc_two)}"
    )
    lines.append(
        f"  - Extra clauses (only in doc 2): {len(result.additional_in_doc_two)}"
    )
    lines.append(
        f"  - Content changes: {len(result.changed)}"
    )

    def _format_clause_list(label: str, clauses: Iterable[Clause], doc_label: str) -> None:
        clauses = list(clauses)
        if not clauses:
            lines.append(f"{label}: none")
            return
        lines.append(f"{label}:")
        for clause in clauses:
            snippet = clause.body[:80].replace("\n", " ")
            lines.append(
                f"  - {clause.heading} ({doc_label} clause #{clause.index + 1}): {snippet}..."
            )

    _format_clause_list(
        "Missing in document 2 (present in document 1)",
        result.missing_from_doc_two,
        "doc 1",
    )
    _format_clause_list(
        "Additional clauses only in document 2",
        result.additional_in_doc_two,
        "doc 2",
    )

    if result.changed:
        lines.append("Same heading but content changed:")
        for idx, change in enumerate(result.changed, start=1):
            lines.append(f"  {idx}. {change['heading']}")
            lines.append("     Diff:")
            diff_lines = change["diff"].splitlines() or ["(whitespace-only differences)"]
            for diff_line in diff_lines:
                lines.append(f"       {diff_line}")
    else:
        lines.append("Same heading but content changed: none")

    return "\n".join(lines)


def parse_args(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Compare clause headings between two documents."
    )
    parser.add_argument("doc_one", type=Path, help="Path to the first (reference) document")
    parser.add_argument("doc_two", type=Path, help="Path to the second document to compare")
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print the comparison result as JSON instead of plain text",
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])

    doc_one_text = read_document(args.doc_one)
    doc_two_text = read_document(args.doc_two)

    clauses_one = parse_clauses(doc_one_text)
    clauses_two = parse_clauses(doc_two_text)

    result = compare_clauses(clauses_one, clauses_two)

    if args.json:
        payload = {
            "missing_from_doc_two": [
                clause.heading for clause in result.missing_from_doc_two
            ],
            "additional_in_doc_two": [
                clause.heading for clause in result.additional_in_doc_two
            ],
            "changed": result.changed,
        }
        print(json.dumps(payload, indent=2))
    else:
        print(format_result(result))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
