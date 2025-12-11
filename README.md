# Clause Comparison Helper

This repository contains a single Python script, `compare_clauses.py`, that lets
you compare two documents clause-by-clause. It is written for beginners, so the
code is heavily commented and the README walks you through every step.

## What the script does

1. Reads two documents (plain text or `.docx` if you install `python-docx`).
2. Detects clause headings such as:
   - `Clause 1: Introduction`
   - `2.0 Background`
   - `SECTION 5 – TERMINATION`
3. Groups everything after each heading as the body of that clause.
4. Treats the first document as the reference copy.
5. Prints a clear report that shows:
   - Clauses that exist in document one but are missing in document two.
   - Clauses that appear only in document two (extra clauses).
   - Clauses that exist in both files but have different wording (with a diff).

## Requirements

- Python 3.9 or newer already installed on your machine.
- Optional: `python-docx` if you plan to read `.docx` files.

Install the optional dependency with:

```bash
pip install python-docx
```

## File structure

```
/workspace
├── compare_clauses.py   # The script you run
└── README.md            # This guide
```

## How to run the script

1. Put your two documents in the same folder as the script. For example:
   - `doc_one.txt`
   - `doc_two.txt`
2. Open a terminal in that folder.
3. Run:

```bash
python compare_clauses.py doc_one.txt doc_two.txt
```

To get machine-readable output:

```bash
python compare_clauses.py doc_one.txt doc_two.txt --json
```

### Sample documents

`doc_one.txt`

```
Clause 1: Introduction
This agreement covers how we will work together.

Clause 2: Payment
Invoices are due within 30 days.

Clause 3: Termination
Either party may terminate with 30 days written notice.
```

`doc_two.txt`

```
Clause 1: Introduction
This agreement covers how we will work together and collaborate.

Clause 3: Termination
Either party may terminate with 15 days notice.

Clause 4: Governing Law
This contract follows the laws of California.
```

Command:

```bash
python compare_clauses.py doc_one.txt doc_two.txt
```

Output:

```
CLAUSE COMPARISON REPORT
===========================
Document 2 vs Document 1 overview:
  - Missing clauses (only in doc 1): 1
  - Extra clauses (only in doc 2): 1
  - Content changes: 2
Missing in document 2 (present in document 1):
  - Clause 2: Payment (doc 1 clause #2): Invoices are due within 30 days...
Additional clauses only in document 2:
  - Clause 4: Governing Law (doc 2 clause #3): This contract follows the laws of California....
Same heading but content changed:
  - Clause 1: Introduction
  - Clause 3: Termination
What changed (details):
  1. Clause 1: Introduction
     Diff:
       --- doc_one
       +++ doc_two
       @@
       -This agreement covers how we will work together.
       +This agreement covers how we will work together and collaborate.
  2. Clause 3: Termination
     Diff:
       --- doc_one
       +++ doc_two
       @@
       -Either party may terminate with 30 days written notice.
       +Either party may terminate with 15 days notice.
```

## How the detection works

- Each line is checked to see if it looks like a heading (starts with `Clause`,
  `Section`, a numeric outline like `1.2`, or is in all caps).
- The script groups the lines that follow the heading as the clause body until
  the next heading is encountered.
- Headings are normalized (lowercase, punctuation removed) to match similar
  headings even when formatted differently (`Clause 1:` vs `clause 1`), and
  pure numbering differences (e.g., `16.` vs `17.`) are ignored when matching.
- Clause bodies plus the descriptive portion of each heading are compared (after
  trimming whitespace) so that inline sentences such as `16.1 ...` still count
  as clause content, while minor spacing differences do not matter.

## Troubleshooting

- **`FileNotFoundError`:** Check the file names and paths you passed in.
- **`python-docx` error:** Install the dependency or convert your `.docx`
  documents to plain text.
- **Headings not detected:** Ensure each clause uses a clear heading as shown
  in the samples.
