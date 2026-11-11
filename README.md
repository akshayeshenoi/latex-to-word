# Portable LaTeX-to-Word compiler

Move this entire folder to any directory. It contains the compiler, tests,
dependency list, Word formatting template, and matching JMIR citation style.
The manuscript stays in its own project: pass its main LaTeX file with `--source`.
Generated files go into this folder's `output/` directory by default.

```text
latex-to-word/
  compile_word.py
  requirements.txt
  README.md
  templates/JMIR.docx
  styles/journal-of-medical-internet-research.csl
  tests/
  .venv/          # Local Python environment, created during setup
  output/        # Generated documents, reports, and verification artifacts
```

## Setup and use

Use Python 3.10 or newer and Pandoc 3.8 or newer. From this folder, create an
isolated environment:

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

Python environments contain absolute installation paths. After moving to a new
location or machine, recreate `.venv` with the setup commands (remove the old
`.venv` first). No source-code changes are needed. Python, Pandoc, and Zotero are
system dependencies, not bundled applications.

Put Pandoc on `PATH`, beside the chosen Python executable, or set `PANDOC` to its executable path. Zotero must be open,
with its local API and the existing Better BibTeX citation-key integration
available. The compiler never enables these or changes Zotero settings.

```sh
.venv/bin/python compile_word.py --source /path/to/manuscript/00_main.tex --check
.venv/bin/python compile_word.py --source /path/to/manuscript/00_main.tex
```

The first command checks sources, images, the Word template, the CSL
style, and every citation key without writing output. The second generates and
validates the document. A sibling `.build.json` report records counts, citation
links, hashes, and validation results. Generated output and the virtual
environment are ignored by Git.

Options:

| Option | Meaning | Default |
| --- | --- | --- |
| `--source` | Main LaTeX file containing the manuscript | Required |
| `--template` | Word formatting and Zotero settings reference | Bundled `templates/JMIR.docx` |
| `--output` | Generated Word file | `output/<source-stem>.docx` inside this folder |
| `--pandoc` | Pandoc executable | `PANDOC`, then `PATH`, then beside Python |
| `--csl` | Existing CSL file matching the template's Zotero style ID | Bundled `styles/<style-id>.csl`, then `~/Zotero/styles/<style-id>.csl` |
| `--check` | Read-only preflight; no document generated | Off |

The defaults work from any working directory and follow the compiler when it
moves. Relative option values resolve from the directory in which the command
is invoked. Output cannot replace the template. For example, from elsewhere:

```sh
/path/to/latex-to-word/.venv/bin/python /path/to/latex-to-word/compile_word.py \
  --source /path/to/manuscript/00_main.tex
```

The bundled template is a blank, generic reference file containing only styles,
page settings, font/theme/list definitions, and Zotero preferences. It contains
no manuscript text, figures, citations, reference entries, or author metadata.
Table and caption formatting live in reusable styles rather than example content.
Replace it or pass `--template` to use
another Word document containing the desired styles and Zotero settings. Pass
`--csl` for a different matching citation style when needed.

## Content and formatting

LaTeX is the content source. Author names, keywords, declarations, placeholders,
and section order come from LaTeX. Word-only additions are not copied. The full
appendix follows the references in the same document.

The compiler retains the template's styles, theme, page size and margins. It
maps title/sections to the template's heading hierarchy, keeps source tables
editable, applies the reference table style to source-derived column widths, and fits figures to the text
width. PDF figures are rendered at 600 dpi in temporary build files; original
assets are untouched. Mathematics becomes editable Word equations. Table and
figure references remain internal hyperlinks.

Layout-only LaTeX commands are ignored. In particular, boxes such as `\makebox`
and `\resizebox` are unwrapped while preserving their contents. This is a
compiler for this manuscript's LaTeX conventions, not a complete TeX engine.
Unsupported content-bearing commands fail visibly rather than disappearing.
Page breaks can differ from the reference document as text changes.

Each build replaces the generated Word file. Manual edits in generated Word
files are not merged back into LaTeX. Sources, `references.bib`, and the template
are not modified. A failed conversion or validation retains the last successful
Word file.

## Zotero behavior and acceptance check

All library requests are read-only GETs to `127.0.0.1:23119`. Exact citation keys
are resolved in group `aimi-llm` (`5808658`) before the personal library. Ambiguous
keys, missing records, and failed requests stop the build. There is no web lookup,
library write, API setup, or offline fallback.

`\cite{key}` and `\cite{key1,key2}` become native Zotero Word fields with current
item metadata and library links. References are contained in one native Zotero
bibliography field. Each generated document receives its own Zotero session ID;
the template's citation style and other preferences are retained. Optional
locators, citation affixes, and narrative citation commands are currently rejected.

Automated validation checks XML, package links, field boundaries, citation
coverage, unique citation IDs, reference counts, preserved styles/page settings,
figures, tables, equations, and internal cross-references. It does not prove that
the installed Word plugin works. To check that final integration:

1. Open a copy of the generated document in Microsoft Word with Zotero running.
2. Use Zotero **Refresh**. Check numbering, bibliography, and the JMIR style.
3. Click inside a citation and use **Add/Edit Citation**. Confirm the cited
   library items are recognized, then cancel without changing them.
4. Review the abstract, wide table, figures, references, and appendix visually.

Word/Zotero may normalize journal abbreviations or other rendered citation
details on Refresh; Pandoc's CSL rendering provides the initial visible result.
Never use **Unlink Citations** if you want to retain Zotero editing.

## Tests

```sh
.venv/bin/python -m unittest discover -s tests -v
```

Tests use a fake read-only Zotero client and do not access the library. The Word
integration tests require Pandoc and the bundled template and JMIR CSL style;
they skip explicitly if these are missing. They cover relocation to a directory
with spaces and an external manuscript, citation
reordering, grouping, repetition and additions, appendix preservation, failures
without replacing output, read-only preflight, exact-key resolution, group-first
lookup, pagination, missing assets, and broken references.
