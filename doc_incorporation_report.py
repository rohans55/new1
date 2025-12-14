#!/usr/bin/env python3

"""Generate a Doc2->Doc1 incorporation report.

Use-case:
- Doc1_template: the standard template before enrichment
- Doc2: supplementary document with requirements/details
- Doc1_updated: the system output (template after incorporating Doc2)

This script reports which Doc2 items are present in Doc1_updated but not in
Doc1_template using similarity matching.

Backends:
- tfidf (default): lightweight, no PyTorch; good baseline but less semantic
- semantic: SentenceTransformers embeddings (more \"meaning-based\")

Outputs:
- Markdown (default) to stdout
- Optional JSON or Markdown file via --out

Notes:
- Matching is approximate. Tune --threshold and chunk sizes.
- For .pdf/.docx, basic text extraction is used.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from dataclasses import dataclass
from typing import Iterable, List, Optional, Tuple


def _read_text_file(path: str) -> str:
    with open(path, "rb") as f:
        raw = f.read()
    # Try utf-8, then latin-1 as a last resort.
    for enc in ("utf-8", "utf-8-sig", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="ignore")


def _read_docx(path: str) -> str:
    try:
        import docx  # python-docx
    except Exception as e:  # pragma: no cover
        raise RuntimeError(
            "Reading .docx requires python-docx. Install requirements.txt"
        ) from e

    d = docx.Document(path)
    parts: List[str] = []
    for p in d.paragraphs:
        t = (p.text or "").strip()
        if t:
            parts.append(t)
    # Tables too
    for table in d.tables:
        for row in table.rows:
            for cell in row.cells:
                t = (cell.text or "").strip()
                if t:
                    parts.append(t)
    return "\n\n".join(parts)


def _read_pdf(path: str) -> str:
    try:
        from pypdf import PdfReader
    except Exception as e:  # pragma: no cover
        raise RuntimeError("Reading .pdf requires pypdf. Install requirements.txt") from e

    reader = PdfReader(path)
    parts: List[str] = []
    for page in reader.pages:
        try:
            t = page.extract_text() or ""
        except Exception:
            t = ""
        t = t.strip()
        if t:
            parts.append(t)
    return "\n\n".join(parts)


def load_document_text(path: str) -> str:
    if not os.path.exists(path):
        raise FileNotFoundError(path)

    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        text = _read_docx(path)
    elif ext == ".pdf":
        text = _read_pdf(path)
    else:
        text = _read_text_file(path)

    # Normalize whitespace
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[\t\f\v]+", " ", text)
    text = re.sub(r"[ ]{2,}", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


_SENT_SPLIT = re.compile(r"(?<=[.!?])\s+(?=[A-Z0-9(\[])")


def split_into_claims(doc2_text: str, min_len: int, max_len: int) -> List[str]:
    """Split Doc2 into "claims" (items to check for incorporation)."""
    raw_paras = [p.strip() for p in re.split(r"\n\s*\n", doc2_text) if p.strip()]

    claims: List[str] = []
    for para in raw_paras:
        # Skip likely headings/titles (common in Doc2)
        wc = len(re.findall(r"\b\w+\b", para))
        if (
            len(para) < 80
            and wc <= 10
            and not re.search(r"[.!?:]", para)
            and (para == para.title() or para.isupper())
        ):
            continue

        # Further split if it looks like a list/bullets
        bullet_lines = [
            ln.strip()
            for ln in para.split("\n")
            if ln.strip().startswith(("- ", "* ", "•", "1.", "2.", "3.", "4.", "5."))
        ]
        if len(bullet_lines) >= 2:
            for ln in bullet_lines:
                c = re.sub(r"^([\-*•]|\d+\.)\s+", "", ln).strip()
                if len(c) >= min_len:
                    claims.append(c)
            continue

        if len(para) <= max_len:
            if len(para) >= min_len:
                claims.append(para)
            continue

        # Split long paragraphs into sentence groups
        sentences = [s.strip() for s in _SENT_SPLIT.split(para) if s.strip()]
        buf: List[str] = []
        cur = 0
        for s in sentences:
            if cur + len(s) + 1 > max_len and buf:
                c = " ".join(buf).strip()
                if len(c) >= min_len:
                    claims.append(c)
                buf = [s]
                cur = len(s)
            else:
                buf.append(s)
                cur += len(s) + 1
        if buf:
            c = " ".join(buf).strip()
            if len(c) >= min_len:
                claims.append(c)

    # De-duplicate while preserving order
    seen = set()
    uniq: List[str] = []
    for c in claims:
        k = re.sub(r"\s+", " ", c).strip().lower()
        if k in seen:
            continue
        seen.add(k)
        uniq.append(c)
    return uniq


def chunk_text(text: str, chunk_size: int, overlap: int) -> List[str]:
    """Chunk text into overlapping windows (by characters, but sentence-aware)."""
    text = text.strip()
    if not text:
        return []

    # First split into pseudo-sentences, then pack into chunks.
    parts = [p.strip() for p in re.split(r"\n\n+", text) if p.strip()]
    sentences: List[str] = []
    for p in parts:
        # Keep headers / short lines as-is
        lines = [ln.strip() for ln in p.split("\n") if ln.strip()]
        if len(lines) > 1 and all(len(ln) < 120 for ln in lines):
            sentences.extend(lines)
        else:
            sentences.extend([s.strip() for s in _SENT_SPLIT.split(p) if s.strip()])

    chunks: List[str] = []
    cur: List[str] = []
    cur_len = 0

    def flush():
        nonlocal cur, cur_len
        if cur:
            chunks.append(" ".join(cur).strip())
        cur = []
        cur_len = 0

    for s in sentences:
        if not s:
            continue
        add_len = len(s) + (1 if cur else 0)
        if cur_len + add_len <= chunk_size:
            cur.append(s)
            cur_len += add_len
            continue

        # Flush and start new chunk; apply overlap
        flush()
        if overlap > 0 and chunks:
            prev = chunks[-1]
            cur = [prev[-overlap:]] if len(prev) > overlap else [prev]
            cur_len = sum(len(x) for x in cur) + (len(cur) - 1)

        if len(s) > chunk_size:
            # Hard split very long sentence
            for i in range(0, len(s), chunk_size):
                chunks.append(s[i : i + chunk_size])
            cur = []
            cur_len = 0
        else:
            cur = [s]
            cur_len = len(s)

    flush()

    # Clean up overlap artifacts
    cleaned = [re.sub(r"\s+", " ", c).strip() for c in chunks if c.strip()]
    return cleaned


@dataclass
class Match:
    claim: str
    status: str  # incorporated | already_in_template | not_found
    sim_updated: float
    sim_template: float
    best_updated_chunk: str
    best_template_chunk: str
    best_updated_idx: int
    best_template_idx: int
    lexical_updated: float
    lexical_template: float


def _lexical_score(claim: str, chunk: str) -> float:
    try:
        from rapidfuzz.fuzz import partial_ratio

        return float(partial_ratio(claim, chunk)) / 100.0
    except Exception:
        return 0.0


def build_report(
    doc1_template: str,
    doc2: str,
    doc1_updated: str,
    backend: str,
    model_name: str,
    threshold: float,
    chunk_size: int,
    chunk_overlap: int,
    min_claim_len: int,
    max_claim_len: int,
    topk: int,
) -> Tuple[List[Match], dict]:
    # Prepare text units
    claims = split_into_claims(doc2, min_len=min_claim_len, max_len=max_claim_len)
    t_chunks = chunk_text(doc1_template, chunk_size=chunk_size, overlap=chunk_overlap)
    u_chunks = chunk_text(doc1_updated, chunk_size=chunk_size, overlap=chunk_overlap)

    if not claims:
        raise ValueError("Doc2 produced zero claims; try lowering --min-claim-len")
    if not u_chunks:
        raise ValueError("Doc1_updated produced zero chunks; input seems empty")
    if not t_chunks:
        # Allow empty template
        t_chunks = [""]

    backend = (backend or "").strip().lower()
    if backend not in {"tfidf", "semantic"}:
        raise ValueError(f"Unsupported backend: {backend}")

    import numpy as np

    if backend == "semantic":
        try:
            from sentence_transformers import SentenceTransformer
        except Exception as e:  # pragma: no cover
            raise RuntimeError(
                "Semantic backend requires sentence-transformers (and torch). "
                "Install: pip install -r requirements-semantic.txt"
            ) from e

        model = SentenceTransformer(model_name)
        claim_emb = model.encode(claims, normalize_embeddings=True)
        t_emb = model.encode(t_chunks, normalize_embeddings=True)
        u_emb = model.encode(u_chunks, normalize_embeddings=True)

        sim_u = claim_emb @ u_emb.T
        sim_t = claim_emb @ t_emb.T

        def best(sim_mat):
            idx = np.argmax(sim_mat, axis=1)
            val = sim_mat[np.arange(sim_mat.shape[0]), idx]
            return idx.tolist(), val.tolist()

        u_idx, u_val = best(sim_u)
        t_idx, t_val = best(sim_t)

    else:
        # TF-IDF cosine similarity (fast + light)
        try:
            from sklearn.feature_extraction.text import TfidfVectorizer
        except Exception as e:  # pragma: no cover
            raise RuntimeError(
                "TF-IDF backend requires scikit-learn. Install: pip install -r requirements.txt"
            ) from e

        vec = TfidfVectorizer(ngram_range=(1, 2), min_df=1, max_df=0.95)
        all_texts = claims + u_chunks + t_chunks
        X = vec.fit_transform(all_texts)

        cX = X[: len(claims)]
        uX = X[len(claims) : len(claims) + len(u_chunks)]
        tX = X[len(claims) + len(u_chunks) :]

        sim_u = (cX @ uX.T).toarray()
        sim_t = (cX @ tX.T).toarray()

        u_idx = np.argmax(sim_u, axis=1).tolist()
        u_val = sim_u[np.arange(sim_u.shape[0]), u_idx].tolist()

        t_idx = np.argmax(sim_t, axis=1).tolist()
        t_val = sim_t[np.arange(sim_t.shape[0]), t_idx].tolist()

    matches: List[Match] = []
    for i, claim in enumerate(claims):
        bi_u = int(u_idx[i])
        bi_t = int(t_idx[i])
        su = float(u_val[i])
        st = float(t_val[i])
        bu = u_chunks[bi_u] if 0 <= bi_u < len(u_chunks) else ""
        bt = t_chunks[bi_t] if 0 <= bi_t < len(t_chunks) else ""

        # Decide status
        in_updated = su >= threshold
        in_template = st >= threshold
        if in_updated and not in_template:
            status = "incorporated"
        elif in_updated and in_template:
            status = "already_in_template"
        else:
            status = "not_found"

        matches.append(
            Match(
                claim=claim,
                status=status,
                sim_updated=su,
                sim_template=st,
                best_updated_chunk=bu,
                best_template_chunk=bt,
                best_updated_idx=bi_u,
                best_template_idx=bi_t,
                lexical_updated=_lexical_score(claim, bu),
                lexical_template=_lexical_score(claim, bt),
            )
        )

    summary = {
        "claims": len(matches),
        "incorporated": sum(1 for m in matches if m.status == "incorporated"),
        "already_in_template": sum(1 for m in matches if m.status == "already_in_template"),
        "not_found": sum(1 for m in matches if m.status == "not_found"),
        "threshold": threshold,
        "backend": backend,
        "model": model_name if backend == "semantic" else "",
        "chunk_size": chunk_size,
        "chunk_overlap": chunk_overlap,
    }

    # Sort: incorporated first, then already, then not_found; within group by sim_updated desc
    order = {"incorporated": 0, "already_in_template": 1, "not_found": 2}
    matches.sort(key=lambda m: (order.get(m.status, 9), -m.sim_updated, -m.sim_template))

    return matches, summary


def render_markdown(matches: List[Match], summary: dict, show_chunks: int) -> str:
    lines: List[str] = []
    lines.append("## Doc2 incorporation report")
    lines.append("")
    lines.append(
        f"- **Claims checked**: {summary['claims']}\n"
        f"- **Incorporated (present in updated, not in template)**: {summary['incorporated']}\n"
        f"- **Already in template**: {summary['already_in_template']}\n"
        f"- **Not found in updated**: {summary['not_found']}\n"
        f"- **Threshold**: {summary['threshold']}\n"
        f"- **Backend**: {summary['backend']}\n"
        + (f"- **Model**: {summary['model']}\n" if summary.get("model") else "")
    )

    def add_section(title: str, statuses: Iterable[str]):
        items = [m for m in matches if m.status in set(statuses)]
        if not items:
            return
        lines.append(f"## {title}")
        lines.append("")
        for idx, m in enumerate(items, 1):
            lines.append(f"### {idx}. {m.status}")
            lines.append("")
            lines.append(f"- **Doc2 claim**: {m.claim}")
            lines.append(
                f"- **Similarity**: updated={m.sim_updated:.3f}, template={m.sim_template:.3f}"
            )
            lines.append(
                f"- **Lexical (sanity)**: updated={m.lexical_updated:.3f}, template={m.lexical_template:.3f}"
            )
            if show_chunks > 0:
                lines.append("")
                lines.append("Best match in updated Doc1:")
                lines.append("")
                excerpt_u = m.best_updated_chunk
                if len(excerpt_u) > show_chunks:
                    excerpt_u = excerpt_u[:show_chunks].rstrip() + " …"
                lines.append("```")
                lines.append(excerpt_u)
                lines.append("```")

                lines.append("Best match in template Doc1:")
                lines.append("")
                excerpt_t = m.best_template_chunk
                if len(excerpt_t) > show_chunks:
                    excerpt_t = excerpt_t[:show_chunks].rstrip() + " …"
                lines.append("```")
                lines.append(excerpt_t)
                lines.append("```")
            lines.append("")

    add_section("Incorporated items", ["incorporated"])
    add_section("Already present in template", ["already_in_template"])
    add_section("Not found in updated output", ["not_found"])

    return "\n".join(lines).rstrip() + "\n"


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(
        description="Report which Doc2 content appears incorporated into updated Doc1."
    )
    ap.add_argument("--doc1-template", required=True, help="Path to original Doc1 template")
    ap.add_argument("--doc2", required=True, help="Path to supplementary Doc2")
    ap.add_argument("--doc1-updated", required=True, help="Path to system-updated Doc1")

    ap.add_argument(
        "--backend",
        choices=["tfidf", "semantic"],
        default="tfidf",
        help="Matching backend (tfidf is lightweight; semantic is more meaning-based)",
    )
    ap.add_argument(
        "--model",
        default="sentence-transformers/all-MiniLM-L6-v2",
        help="SentenceTransformer model name (semantic backend only)",
    )
    ap.add_argument(
        "--threshold",
        type=float,
        default=None,
        help="Similarity threshold for 'present' (default depends on backend)",
    )
    ap.add_argument("--chunk-size", type=int, default=900, help="Chunk size (chars)")
    ap.add_argument("--chunk-overlap", type=int, default=120, help="Overlap (chars)")
    ap.add_argument(
        "--min-claim-len", type=int, default=25, help="Minimum length of Doc2 claim"
    )
    ap.add_argument(
        "--max-claim-len",
        type=int,
        default=600,
        help="Maximum length of a claim before splitting",
    )
    ap.add_argument(
        "--show-chunks",
        type=int,
        default=550,
        help="How many chars of matched chunks to show (0 disables)",
    )
    ap.add_argument(
        "--out",
        default="",
        help="Optional output file path (.md or .json). If omitted, prints Markdown to stdout.",
    )
    ap.add_argument(
        "--format",
        choices=["md", "json"],
        default="md",
        help="Output format (default: md)",
    )

    args = ap.parse_args(argv)

    if args.threshold is None:
        # Reasonable defaults: TF-IDF scores run lower than semantic cosine.
        args.threshold = 0.28 if args.backend == "tfidf" else 0.72

    doc1_template = load_document_text(args.doc1_template)
    doc2 = load_document_text(args.doc2)
    doc1_updated = load_document_text(args.doc1_updated)

    matches, summary = build_report(
        doc1_template=doc1_template,
        doc2=doc2,
        doc1_updated=doc1_updated,
        backend=args.backend,
        model_name=args.model,
        threshold=args.threshold,
        chunk_size=args.chunk_size,
        chunk_overlap=args.chunk_overlap,
        min_claim_len=args.min_claim_len,
        max_claim_len=args.max_claim_len,
        topk=1,
    )

    if args.format == "json" or (args.out.lower().endswith(".json") if args.out else False):
        payload = {
            "summary": summary,
            "matches": [m.__dict__ for m in matches],
        }
        out_s = json.dumps(payload, indent=2, ensure_ascii=False)
    else:
        out_s = render_markdown(matches, summary, show_chunks=args.show_chunks)

    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(out_s)
    else:
        sys.stdout.write(out_s)

    # Exit code: non-zero if any 'not_found' claims
    return 2 if summary["not_found"] > 0 else 0


if __name__ == "__main__":
    raise SystemExit(main())
