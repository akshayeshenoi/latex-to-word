#!/usr/bin/env python3
"""Compile this manuscript to DOCX with native, editable Zotero fields.

Zotero access is exclusively read-only HTTP GET. Sources and templates are inputs.
"""
from __future__ import annotations

import argparse
from collections import Counter
from copy import deepcopy
import hashlib
import json
import os
from pathlib import Path
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.error
import urllib.parse
import urllib.request
import uuid
from zipfile import ZipFile, ZIP_DEFLATED

try:
    from lxml import etree as ET
    import pymupdf
except ImportError as exc:
    requirements = Path(__file__).resolve().with_name("requirements.txt")
    raise SystemExit(f'Install dependencies with: python -m pip install -r "{requirements}"') from exc

HERE = Path(__file__).resolve().parent
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W, "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "m": "http://schemas.openxmlformats.org/officeDocument/2006/math"}
CP = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT = "http://schemas.openxmlformats.org/package/2006/content-types"
BIB_MARK = "WORDCOMPILERBIBLIOGRAPHY"
APP_MARK = "WORDCOMPILERAPPENDIX"
XML_PARSER = ET.XMLParser(resolve_entities=False, no_network=True)


class BuildError(Exception):
    pass


def require(condition, message):
    if not condition:
        raise BuildError(message)


def parse(data):
    return ET.fromstring(data, XML_PARSER)


def xml(node):
    return ET.tostring(node, encoding="UTF-8", xml_declaration=True, standalone=True)


def q(name):
    return f"{{{W}}}{name}"


def child(parent, name, **attrs):
    node = parent.find(q(name))
    if node is None:
        node = ET.Element(q(name))
        if name.endswith("Pr"):
            parent.insert(0, node)
        else:
            parent.append(node)
    for key, value in attrs.items():
        node.set(q(key), str(value))
    return node


def nodes(obj, kind=None):
    if isinstance(obj, dict):
        if "t" in obj and (kind is None or obj["t"] == kind):
            yield obj
        for value in obj.values():
            yield from nodes(value, kind)
    elif isinstance(obj, list):
        for value in obj:
            yield from nodes(value, kind)


def text(obj):
    if isinstance(obj, dict):
        if obj.get("t") == "Str":
            return obj["c"]
        if obj.get("t") in ("Space", "SoftBreak", "LineBreak"):
            return " "
        if obj.get("t") in ("Code", "Math"):
            return obj["c"][1]
        return text(obj.get("c", []))
    if isinstance(obj, list):
        return "".join(text(x) for x in obj)
    return ""


def string(value):
    return {"t": "Str", "c": value}


def header(title, level=2, identifier=""):
    return {"t": "Header", "c": [level, [identifier, ["unnumbered"], []], [string(title)]]}


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def run_pandoc(binary, arguments, *, data=None, cwd=None):
    result = subprocess.run([binary, *arguments], input=data, text=True, capture_output=True, cwd=cwd)
    require(result.returncode == 0, f"Pandoc failed:\n{result.stderr.strip()}")
    require(not result.stderr.strip(), f"Pandoc reported a conversion problem:\n{result.stderr.strip()}")
    return result.stdout


def balanced(s, start):
    """Return the contents/end of a braced TeX argument, respecting escaped braces."""
    require(start < len(s) and s[start] == "{", "Expected a braced LaTeX argument")
    depth, i = 1, start + 1
    while i < len(s):
        if s[i] == "\\":
            i += 2
            continue
        if s[i] == "{":
            depth += 1
        elif s[i] == "}":
            depth -= 1
            if depth == 0:
                return s[start + 1:i], i + 1
        i += 1
    raise BuildError("Unbalanced braces in LaTeX source")


def outside_verbatim(s, transform):
    pattern = r"(\\begin\{(Verbatim|verbatim|lstlisting)\}(?:.|\n)*?\\end\{\2\})"
    out, pos = [], 0
    for match in re.finditer(pattern, s):
        out.extend([transform(s[pos:match.start()]), match[0]])
        pos = match.end()
    return "".join(out) + transform(s[pos:])


def strip_comments(s):
    return re.sub(r"(?<!\\)(?:\\\\)*%[^\n]*", lambda m: m[0][:m[0].index('%')], s)


def expand_source(path, tracked, stack=(), root=None):
    path = path.resolve()
    root = root or path.parent
    require(path not in stack, f"Circular LaTeX input: {path}")
    require(path.is_file(), f"Missing LaTeX input: {path}")
    tracked[path] = digest(path)

    def expand(s):
        s = strip_comments(s)
        def include(match):
            value = Path(match[1])
            if not value.suffix:
                value = value.with_suffix(".tex")
            candidates = [root / value, path.parent / value]
            target = next((p for p in candidates if p.is_file()), candidates[0])
            return expand_source(target, tracked, (*stack, path), root)
        return re.sub(r"\\(?:input|include)\s*\{([^}]+)\}", include, s)

    return outside_verbatim(path.read_text(), expand)


def preprocess(s):
    """Ignore layout wrappers, retaining their complete content."""
    def convert(part):
        # These commands affect TeX layout only; don't discard their arguments.
        pattern = r"\\(?:makebox|mbox|fbox|framebox|resizebox|scalebox)(?![A-Za-z])"
        while match := re.search(pattern, part):
            i = match.end()
            name = match[0][1:]
            for _ in range({"resizebox": 2, "scalebox": 1}.get(name, 0)):
                while i < len(part) and part[i].isspace():
                    i += 1
                _, i = balanced(part, i)
            while True:
                while i < len(part) and part[i].isspace():
                    i += 1
                if i < len(part) and part[i] == "[":
                    end = part.find("]", i)
                    require(end >= 0, "Unclosed layout option")
                    i = end + 1
                else:
                    break
            content, end = balanced(part, i)
            part = part[:match.start()] + content + part[end:]
        part = part.replace(r"\ackname", "Acknowledgments").replace(r"\discintname", "Disclosure of Interests")
        # Pandoc otherwise emits cline's column range as literal cell content.
        part = re.sub(r"\\(?:cline|cmidrule)(?:\([^)]*\))?\s*\{[^}]*\}", "", part)
        part = re.sub(r"\\keywords\{([^}]+)\}", lambda m: r"\par\textbf{Keywords:} " + m[1].replace(r"\and", "; "), part)
        part = re.sub(r"\\bibliography\{[^}]+\}", lambda m: r"\section*{" + BIB_MARK + "}", part)
        part = re.sub(r"\\appendix\b", lambda m: r"\section*{" + APP_MARK + "}", part)
        return part
    return outside_verbatim(s, convert)


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        raise BuildError("Zotero redirected a read request; refusing to follow it")


class Zotero:
    def __init__(self, base="http://127.0.0.1:23119"):
        self.base = base
        self.opener = urllib.request.build_opener(urllib.request.ProxyHandler({}), NoRedirect())

    def get(self, path, **params):
        url = self.base + path
        if params:
            url += "?" + urllib.parse.urlencode(params)
        try:
            req = urllib.request.Request(url, method="GET", headers={"Zotero-API-Version": "3"})
            with self.opener.open(req, timeout=30) as response:
                return json.load(response)
        except (urllib.error.URLError, OSError, ValueError) as exc:
            hint = ("Local network access is blocked by the execution sandbox." if "Operation not permitted" in str(exc)
                    else "Check that Zotero is open and its local API is enabled.")
            raise BuildError(f"Read-only Zotero request failed ({path}): {exc}. {hint} No settings were changed.") from exc

    def inventory(self, path):
        result, start = [], 0
        while True:
            page = self.get(path + "/items/top", limit=100, start=start)
            require(isinstance(page, list), "Unexpected Zotero inventory response")
            result.extend(page)
            if len(page) < 100:
                return result
            start += len(page)

    def resolve(self, keys):
        unresolved, resolved = set(keys), {}
        for library in ("/api/groups/5808658", "/api/users/0"):
            if not unresolved:
                break
            records = self.inventory(library)
            for key in sorted(unresolved):
                matches = [item for item in records if item.get("data", {}).get("citationKey") == key]
                require(len(matches) <= 1, f"Ambiguous Zotero citation key {key!r} in {library}")
                if not matches:
                    continue
                item = matches[0]
                csl = self.get(f"{library}/items/{item['key']}", format="csljson")
                require(isinstance(csl, list) and len(csl) == 1, f"Expected one CSL item for {key}")
                require(csl[0].get("citation-key") == key, f"CSL citation key mismatch for {key}")
                uri = item["links"]["alternate"]["href"]
                # Zotero's canonical integration URIs use http://zotero.org.
                uri = uri.replace("https://www.zotero.org/", "http://zotero.org/").replace("https://zotero.org/", "http://zotero.org/")
                data = deepcopy(csl[0])
                data["id"] = key
                resolved[key] = {"key": key, "uri": uri, "item_key": item["key"], "library": library, "itemData": data}
            unresolved.difference_update(resolved)
        require(not unresolved, "No exact Zotero match for: " + ", ".join(sorted(unresolved)))
        return resolved


def read_template(path):
    require(path.is_file(), f"Missing Word template: {path}")
    with ZipFile(path) as z:
        parts = {name: z.read(name) for name in z.namelist()}
    require("docProps/custom.xml" in parts, "Template has no Zotero document settings")
    props = parse(parts["docProps/custom.xml"])
    prefs = sorted((p for p in props if p.get("name", "").startswith("ZOTERO_PREF_")), key=lambda p: int(p.get("name").rsplit("_", 1)[1]))
    require(prefs, "Template has no ZOTERO_PREF properties")
    settings = parse("".join("".join(p.itertext()) for p in prefs).encode())
    style = settings.find("style")
    require(style is not None, "Template has no Zotero citation style")
    field = settings.find("prefs/pref[@name='fieldType']")
    require(field is not None and field.get("value") == "Field", "Template must use Zotero Word fields")
    doc = parse(parts["word/document.xml"])
    sect = doc.find(".//w:sectPr", NS)
    require(sect is not None, "Template has no page settings")
    size, margins = sect.find("w:pgSz", NS), sect.find("w:pgMar", NS)
    width = int(size.get(q("w"))) - int(margins.get(q("left"))) - int(margins.get(q("right")))
    return {"parts": parts, "props": prefs, "settings": settings, "style": style.get("id"), "locale": style.get("locale", "en-US"), "doc": doc, "width": width}


def transform_ast(ast):
    meta, blocks = ast["meta"], ast["blocks"]
    front = []
    if "title" in meta:
        front.append({"t": "Header", "c": [1, ["manuscript-title", ["unnumbered"], []], meta["title"]["c"]]})
    for author in meta.get("author", {"c": []})["c"]:
        front.append({"t": "Para", "c": author["c"]})
    if "abstract" in meta:
        front.extend([header("Abstract", identifier="abstract"), *meta["abstract"]["c"]])
    # Flatten LLNCS credits without changing the declaration text.
    flat = []
    for b in blocks:
        if b["t"] == "Div" and "credits" in b["c"][0][1]:
            for sub in b["c"][1]:
                if sub["t"] == "Header":
                    sub["c"][0] = 2
            flat.extend(b["c"][1])
        else:
            flat.append(b)
    labels, result, appendix, appnum, bib_count = {}, [], False, 0, 0
    for b in flat:
        if b["t"] == "Header":
            c = b["c"]
            title = text(c[2])
            if title == APP_MARK:
                appendix = True
                continue
            if title == BIB_MARK:
                bib_count += 1
                result.extend([header("References", identifier="references"), {"t": "Div", "c": [["refs", [], []], []]}])
                continue
            if appendix and c[0] == 1:
                require(appnum < 26, "More than 26 appendices are unsupported")
                letter = chr(ord("A") + appnum)
                appnum += 1
                labels[c[1][0]] = letter
                c[2].insert(0, string(f"Appendix {letter}. "))
            if c[0] == 4:
                # Carry a marker so the next body paragraph can become run-in.
                b = {"t": "Div", "c": [[c[1][0], ["word-run-in"], []], [{"t": "Para", "c": [{"t": "Strong", "c": c[2]}]}]]}
            else:
                c[0] += 1
                if "unnumbered" not in c[1][1]:
                    c[1][1].append("unnumbered")
        result.append(b)
    require(bib_count == 1, "Expected exactly one \\bibliography command")
    ast["meta"] = {}
    ast["blocks"] = front + result
    # Figure/table labels can reside on the object itself or its wrapping Div.
    counts = Counter()
    def number(obj, enclosing=""):
        if isinstance(obj, dict):
            typ, c = obj.get("t"), obj.get("c")
            if typ == "Div":
                enclosing = c[0][0] or enclosing
            if typ in ("Figure", "Table"):
                counts[typ] += 1
                key = c[0][0] or enclosing
                if key:
                    require(key not in labels, f"Duplicate label: {key}")
                    labels[key] = str(counts[typ])
                if c[1][1]:
                    c[1][1][0]["c"].insert(0, string(f"{typ} {counts[typ]}. "))
            for value in obj.values():
                number(value, enclosing)
        elif isinstance(obj, list):
            for value in obj:
                number(value, enclosing)
    number(ast["blocks"])
    for link in nodes(ast, "Link"):
        attrs, inlines, target = link["c"]
        if target[0].startswith("#") and dict(attrs[2]).get("reference-type"):
            label = target[0][1:]
            require(label in labels, f"Unresolved cross-reference: {label}")
            link["c"][1] = [string(labels[label])]
    # Merge the heading into the following paragraph where possible.
    bs, merged = ast["blocks"], []
    i = 0
    while i < len(bs):
        b = bs[i]
        if b["t"] == "Div" and "word-run-in" in b["c"][0][1] and i + 1 < len(bs) and bs[i + 1]["t"] == "Para":
            b["c"][1][0]["c"].extend([{"t": "Space"}, *bs[i + 1]["c"]])
            i += 1
        merged.append(b)
        i += 1
    ast["blocks"] = merged
    return counts


def check_ast(ast):
    """Reject semantic content that Pandoc would drop; disregard layout remnants."""
    ignored = []
    layout = re.compile(r"^\\(?:centering|raggedright|raggedleft|small|footnotesize|scriptsize|tiny|normalsize|large|Large|LARGE|noindent|clearpage|newpage|pagebreak|hfill|vfill|hspace\*?|vspace\*?|setlength|addtolength|renewcommand|fontsize|selectfont)\b")
    def visit(x):
        if isinstance(x, dict):
            if x.get("t") == "Cite":
                for cite in x["c"][0]:
                    require(cite["citationMode"]["t"] == "NormalCitation" and not cite["citationPrefix"] and not cite["citationSuffix"], "This compiler supports plain \\cite{key} and grouped citations; citation locators/affixes are not yet supported")
                return
            if x.get("t") in ("RawBlock", "RawInline") and x["c"][0] == "latex":
                raw = x["c"][1].strip()
                require(bool(layout.match(raw)), f"Unsupported content-bearing LaTeX: {raw[:180]}")
                ignored.append(raw)
                x.update({"t": "RawBlock" if x["t"] == "RawBlock" else "RawInline", "c": ["openxml", ""]})
            for value in x.values():
                visit(value)
        elif isinstance(x, list):
            for value in x:
                visit(value)
    visit(ast)
    return ignored


def collect_assets(ast, root, tracked):
    assets = {}
    for image in nodes(ast, "Image"):
        name = image["c"][2][0]
        path = (root / name).resolve()
        require(path.is_file(), f"Missing figure: {path}")
        require(path.suffix.lower() in (".pdf", ".png", ".jpg", ".jpeg"), f"Unsupported figure format: {path}")
        tracked[path] = digest(path)
        assets[name] = path
    return assets


def render_assets(ast, assets, directory, width_twips):
    rendered = {}
    for name, source in assets.items():
        if source.suffix.lower() == ".pdf":
            dest = directory / (hashlib.sha256(name.encode()).hexdigest()[:12] + ".png")
            with pymupdf.open(source) as pdf:
                require(len(pdf) == 1, f"Figure PDF must have one page: {source}")
                pdf[0].get_pixmap(dpi=600, alpha=False).save(dest)
            rendered[name] = str(dest)
        else:
            rendered[name] = str(source)
    for image in nodes(ast, "Image"):
        image["c"][2][0] = rendered[image["c"][2][0]]
        attrs = [(k, v) for k, v in image["c"][0][2] if k not in ("width", "height")]
        attrs.append(("width", f"{width_twips / 1440:.6f}in"))
        image["c"][0][2] = attrs


def field_runs(instruction=None, end=False):
    def fld(kind):
        r = ET.Element(q("r"), nsmap={"w": W})
        child(r, "fldChar", fldCharType=kind)
        return r
    if end:
        return [fld("end")]
    result = [fld("begin")]
    for start in range(0, len(instruction), 8000):
        run = ET.Element(q("r"), nsmap={"w": W})
        node = child(run, "instrText")
        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        node.text = instruction[start:start + 8000]
        result.append(run)
    return [*result, fld("separate")]


def raw_runs(runs):
    return {"t": "RawInline", "c": ["openxml", "".join(ET.tostring(r, encoding="unicode") for r in runs)]}


def embed_citations(ast, resolved):
    citations = []
    for cite in list(nodes(ast, "Cite")):
        items, visible = cite["c"]
        citation_items = []
        for item in items:
            ref = resolved[item["citationId"]]
            metadata = deepcopy(ref["itemData"])
            metadata["id"] = ref["uri"]
            citation_items.append({"id": ref["uri"], "uris": [ref["uri"]], "itemData": metadata})
        display = text(visible)
        payload = {"citationID": uuid.uuid4().hex[:12], "properties": {"formattedCitation": display, "plainCitation": display, "noteIndex": 0}, "citationItems": citation_items, "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"}
        instruction = " ADDIN ZOTERO_ITEM CSL_CITATION " + json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + " "
        cite.clear()
        cite.update({"t": "Span", "c": [["", [], []], [raw_runs(field_runs(instruction)), *visible, raw_runs(field_runs(end=True))]]})
        citations.append({"keys": [i["citationId"] for i in items], "text": display})
    return citations


def replace_child(parent, tag, source):
    existing = parent.find(q(tag))
    if existing is not None:
        parent.remove(existing)
    if source is not None:
        parent.insert(0, deepcopy(source))


def finish_docx(path, template):
    with ZipFile(path) as z:
        parts = {name: z.read(name) for name in z.namelist()}
    doc = parse(parts["word/document.xml"])
    # Restore exact template style definitions, while retaining writer-generated styles.
    styles = parse(parts["word/styles.xml"])
    original_styles = parse(template["parts"]["word/styles.xml"])
    for source in original_styles:
        if source.tag == q("style"):
            target = next((s for s in styles if s.get(q("styleId")) == source.get(q("styleId"))), None)
        else:
            target = styles.find(source.tag)
        if target is not None:
            styles.replace(target, deepcopy(source))
        else:
            styles.append(deepcopy(source))
    parts["word/styles.xml"] = xml(styles)
    for name, data in template["parts"].items():
        if name.startswith("word/theme/"):
            parts[name] = data
    # Body paragraphs in the target use Normal directly, not Pandoc's extra spacing styles.
    for style in doc.findall(".//w:pPr/w:pStyle", NS):
        if style.get(q("val")) in ("BodyText", "FirstParagraph", "Compact", "Author"):
            style.set(q("val"), "Normal")
        if style.get(q("val")) == "SourceCode":
            # The template inherits justified Normal; literal prompt lines must not stretch.
            child(style.getparent(), "jc", val="left")
            child(style.getparent(), "wordWrap", val="on")
    sample_tables = template["doc"].findall(".//w:tbl", NS)
    generic_table = next((s for s in styles if s.get(q("styleId")) == "JMIRTable"), None)
    for tbl in doc.findall(".//w:tbl", NS):
        columns = tbl.findall("w:tblGrid/w:gridCol", NS)
        sample = next((t for t in sample_tables if len(t.findall("w:tblGrid/w:gridCol", NS)) == len(columns)), None)
        if sample is None and generic_table is None:
            continue
        # Reuse layout, never the original table text/caption or cell merges.
        props = deepcopy((sample if sample is not None else generic_table).find("w:tblPr", NS))
        if sample is None:
            props.insert(0, ET.Element(q("tblStyle"), {q("val"): "JMIRTable"}))
        for stale in list(props):
            if stale.tag in (q("tblCaption"), q("tblDescription")):
                props.remove(stale)
        new_caption = tbl.find("w:tblPr/w:tblCaption", NS)
        if new_caption is not None:
            props.append(deepcopy(new_caption))
        replace_child(tbl, "tblPr", props)
        widths = [int(c.get(q("w"))) for c in (sample.findall("w:tblGrid/w:gridCol", NS) if sample is not None else columns)]
        # Generic styles use source-derived proportions and fit the available text width.
        require(widths and sum(widths) > 0, "Table has no usable column widths")
        factor = template["width"] / sum(widths)
        if sample is not None:
            factor = min(1, factor)
        widths = [int(w * factor) for w in widths]
        child(tbl.find("w:tblPr", NS), "tblW", w=sum(widths), type="dxa")
        for col, width in zip(columns, widths):
            col.set(q("w"), str(width))
        sample_run = sample.find(".//w:r/w:rPr", NS) if sample is not None else generic_table.find("w:rPr", NS)
        rows = tbl.findall("w:tr", NS)
        for row_index, row in enumerate(rows):
            child(child(row, "trPr"), "cantSplit")
            pos = 0
            for cell in row.findall("w:tc", NS):
                cp = cell.find("w:tcPr", NS)
                if cp is None:
                    cp = ET.Element(q("tcPr")); cell.insert(0, cp)
                span = cp.find("w:gridSpan", NS)
                count = int(span.get(q("val"))) if span is not None else 1
                child(cp, "tcW", w=sum(widths[pos:pos + count]), type="dxa")
                pos += count
                if len(rows) <= 12 and row_index < len(rows) - 1:
                    for p in cell.findall("w:p", NS):
                        child(child(p, "pPr"), "keepNext")
                for run in cell.findall(".//w:r", NS):
                    rp = run.find("w:rPr", NS)
                    if rp is None:
                        rp = ET.Element(q("rPr")); run.insert(0, rp)
                    if sample_run is not None:
                        for key in ("rFonts", "sz", "szCs"):
                            src = sample_run.find(q(key))
                            if src is not None:
                                replace_child(rp, key, src)
    # Match caption formatting from the target paragraph itself when no named style exists.
    sample_caption = next((p for p in template["doc"].findall("w:body/w:p", NS) if re.match(r"(?:Figure|Table) \d+\.", "".join(p.itertext()))), None)
    for p in doc.findall(".//w:p", NS):
        value = "".join(p.xpath(".//w:t/text()", namespaces=NS))
        if re.match(r"(?:Figure|Table) \d+\.", value):
            if sample_caption is not None:
                replace_child(p, "pPr", sample_caption.find("w:pPr", NS))
            if value.startswith("Table "):
                child(child(p, "pPr"), "keepNext")
    # Citeproc bibliography paragraphs are wrapped in one native Zotero field.
    bibliography = [p for p in doc.findall("w:body/w:p", NS) if p.find("w:pPr/w:pStyle[@w:val='Bibliography']", NS) is not None]
    require(bibliography, "Generated bibliography is empty")
    start = bibliography[0]
    offset = 1 if start.find("w:pPr", NS) is not None else 0
    for run in field_runs(' ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY '):
        start.insert(offset, run); offset += 1
    bibliography[-1].extend(field_runs(end=True))
    parts["word/document.xml"] = xml(doc)
    # Preserve Zotero preferences, with a fresh session identity for this new document.
    props = parse(parts["docProps/custom.xml"]) if "docProps/custom.xml" in parts else ET.Element(f"{{{CP}}}Properties", nsmap={None: CP})
    for p in list(props):
        if p.get("name", "").startswith("ZOTERO_PREF_") or p.get("name") in ("bibliography", "csl"):
            props.remove(p)
    pid = max([int(p.get("pid", 1)) for p in props] + [1]) + 1
    settings = deepcopy(template["settings"])
    session = settings.find("session")
    if session is not None:
        session.set("id", uuid.uuid4().hex[:12])
    serialized = ET.tostring(settings, encoding="unicode")
    for index, start in enumerate(range(0, len(serialized), 255), 1):
        prop = deepcopy(template["props"][0])
        prop.set("pid", str(pid)); pid += 1
        prop.set("name", f"ZOTERO_PREF_{index}")
        prop[0].text = serialized[start:start + 255]
        props.append(prop)
    parts["docProps/custom.xml"] = xml(props)
    rels = parse(parts["_rels/.rels"])
    if not any(r.get("Type", "").endswith("/custom-properties") for r in rels):
        ET.SubElement(rels, f"{{{REL}}}Relationship", Id="rIdWordCompilerCustom", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties", Target="docProps/custom.xml")
    parts["_rels/.rels"] = xml(rels)
    types = parse(parts["[Content_Types].xml"])
    if not any(t.get("PartName") == "/docProps/custom.xml" for t in types):
        ET.SubElement(types, f"{{{CONTENT}}}Override", PartName="/docProps/custom.xml", ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml")
    parts["[Content_Types].xml"] = xml(types)
    with ZipFile(path, "w", ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


def validate_docx(path, expected, resolved, template):
    with ZipFile(path) as z:
        require(z.testzip() is None, "Corrupt DOCX archive")
        names = set(z.namelist())
        for name in names:
            if name.endswith((".xml", ".rels")):
                root = parse(z.read(name))
                if name.endswith(".rels"):
                    base = "" if name == "_rels/.rels" else name.rsplit("/_rels/", 1)[0]
                    for rel in root:
                        if rel.get("TargetMode") == "External":
                            continue
                        target = urllib.parse.unquote(rel.get("Target", ""))
                        target = posixpath.normpath(posixpath.join(base, target)).lstrip("/")
                        require(target in names, f"Broken DOCX relationship: {name} -> {target}")
        doc = parse(z.read("word/document.xml"))
        styles = parse(z.read("word/styles.xml"))
        for original in parse(template["parts"]["word/styles.xml"]):
            if original.tag == q("style"):
                match = next((s for s in styles if s.get(q("styleId")) == original.get(q("styleId"))), None)
                require(match is not None and ET.tostring(original, method="c14n", exclusive=True) == ET.tostring(match, method="c14n", exclusive=True), f"Template style changed: {original.get(q('styleId'))}")
        stack, instructions = [], []
        for elem in doc.iter():
            if elem.tag == q("fldChar"):
                kind = elem.get(q("fldCharType"))
                if kind == "begin":
                    stack.append({"instruction": "", "separate": False})
                elif kind == "separate":
                    require(bool(stack) and not stack[-1]["separate"], "Invalid field separator")
                    stack[-1]["separate"] = True
                elif kind == "end":
                    require(bool(stack) and stack[-1]["separate"], "Invalid field end")
                    instructions.append(stack.pop()["instruction"])
            elif elem.tag == q("instrText"):
                require(bool(stack) and not stack[-1]["separate"], "Instruction outside field")
                stack[-1]["instruction"] += elem.text or ""
        require(not stack, "Unclosed Word field")
        citations, bibs, ids, used = [], 0, set(), set()
        for instruction in instructions:
            if "ZOTERO_ITEM" in instruction:
                data = json.loads(instruction[instruction.index("{"):])
                require(data["citationID"] not in ids, "Duplicate Zotero citation ID")
                ids.add(data["citationID"]); citations.append(data)
                for item in data["citationItems"]:
                    key = item["itemData"]["citation-key"]
                    require(key in resolved and item["uris"] == [resolved[key]["uri"]], f"Invalid Zotero link for {key}")
                    used.add(key)
            elif "ZOTERO_BIBL" in instruction:
                bibs += 1
        require(len(citations) == expected["citations"] and bibs == 1, "Citation/bibliography field count mismatch")
        require(used == set(resolved), "Zotero reference coverage mismatch")
        for typ, xpath in (("Table", ".//w:tbl"), ("Image", ".//w:drawing"), ("Math", ".//m:oMath")):
            require(len(doc.findall(xpath, NS)) >= expected[typ], f"Lost {typ} content")
        bookmarks = {b.get(q("name")) for b in doc.findall(".//w:bookmarkStart", NS)}
        anchors = {a.get(q("anchor")) for a in doc.findall(".//w:hyperlink", NS) if a.get(q("anchor"))}
        require(not anchors - bookmarks, "Broken internal links: " + ", ".join(anchors - bookmarks))
        for extent in doc.findall(".//wp:extent", NS):
            require(int(extent.get("cx")) <= template["width"] * 635 + 10, "Figure exceeds page width")
        for tag in ("pgSz", "pgMar"):
            require(doc.find(f".//w:sectPr/w:{tag}", NS).attrib == template["doc"].find(f".//w:sectPr/w:{tag}", NS).attrib, f"Template {tag} changed")
        count_bib = len(doc.findall("w:body/w:p/w:pPr/w:pStyle[@w:val='Bibliography']", NS))
        require(count_bib == len(resolved), "Bibliography entry count mismatch")
        plain = "".join(doc.xpath(".//w:t/text()", namespaces=NS))
        require(BIB_MARK not in plain and APP_MARK not in plain, "Internal conversion marker remains")
        return {"citation_fields": len(citations), "references": count_bib, "tables": len(doc.findall('.//w:tbl', NS)), "figures": len(doc.findall('.//w:drawing', NS)), "word_plugin_check": "not performed; verify Refresh and Add/Edit Citation in Microsoft Word"}


def compile_document(args):
    source, template_path = args.source.resolve(), args.template.resolve()
    output = (args.output or HERE / "output" / f"{source.stem}.docx").resolve()
    require(output.suffix.lower() == ".docx", "Output must end in .docx")
    require(output != template_path, "Output must differ from the template")
    python_pandoc = Path(sys.executable).with_name("pandoc")
    binary = args.pandoc or os.environ.get("PANDOC") or shutil.which("pandoc") or (str(python_pandoc) if python_pandoc.is_file() else None)
    require(bool(binary) and (shutil.which(str(binary)) or Path(binary).is_file()), "Pandoc 3.8+ is required. Put it on PATH, set PANDOC, or pass --pandoc.")
    version = run_pandoc(str(binary), ["--version"]).splitlines()[0]
    match = re.search(r"pandoc (\d+)\.(\d+)", version)
    require(match and tuple(map(int, match.groups())) >= (3, 8), "Pandoc 3.8+ is required")
    tracked = {template_path: digest(template_path)}
    template = read_template(template_path)
    style_name = template["style"].rstrip("/").split("/")[-1]
    bundled_style = HERE / "styles" / (style_name + ".csl")
    csl = (args.csl or (bundled_style if bundled_style.is_file() else Path.home() / "Zotero" / "styles" / (style_name + ".csl"))).resolve()
    require(csl.is_file(), f"Missing installed Zotero style: {csl}. Pass --csl with its existing location.")
    cslroot = parse(csl.read_bytes())
    style_id = cslroot.find("{http://purl.org/net/xbiblio/csl}info/{http://purl.org/net/xbiblio/csl}id")
    require(style_id is not None and style_id.text == template["style"], "CSL style does not match the template's Zotero settings")
    tracked[csl] = digest(csl)
    bib_path = source.parent / "references.bib"
    if bib_path.is_file():
        tracked[bib_path] = digest(bib_path)
    latex = preprocess(expand_source(source, tracked))
    ast = json.loads(run_pandoc(str(binary), ["-f", "latex", "-t", "json"], data=latex, cwd=source.parent))
    transform_ast(ast)
    ignored = check_ast(ast)
    assets = collect_assets(ast, source.parent, tracked)
    citations = list(nodes(ast, "Cite"))
    keys = {c["citationId"] for cite in citations for c in cite["c"][0]}
    require(keys, "No citations found in manuscript")
    print(f"Resolving {len(keys)} citation keys against Zotero (read-only)...", flush=True)
    resolved = Zotero().resolve(keys)
    expected = Counter(x["t"] for x in nodes(ast)); expected["citations"] = len(citations)
    report = {"source": str(source), "template": str(template_path), "output": str(output), "pandoc": version, "style": template["style"], "ignored_layout": ignored, "resolved": [{k: v for k, v in resolved[key].items() if k != "itemData"} for key in sorted(resolved)]}
    if args.check:
        print(f"Preflight passed: {len(keys)} references, {len(citations)} citation occurrences, {expected['Table']} tables, {expected['Image']} figures. No output written.")
        return report
    output.parent.mkdir(parents=True, exist_ok=True)
    # Use the output filesystem so the final rename is atomic.
    with tempfile.TemporaryDirectory(prefix=".compile-word-", dir=output.parent) as tmp:
        work = Path(tmp)
        print("Rendering figures and composing Word document...", flush=True)
        render_assets(ast, assets, work, template["width"])
        refs = work / "references.json"
        refs.write_text(json.dumps([r["itemData"] for r in resolved.values()], ensure_ascii=False))
        ast["meta"]["lang"] = {"t": "MetaString", "c": template["locale"]}
        rendered = json.loads(run_pandoc(str(binary), ["-f", "json", "-t", "json", "--citeproc", f"--csl={csl}", f"--bibliography={refs}"], data=json.dumps(ast), cwd=source.parent))
        report["citations"] = embed_citations(rendered, resolved)
        docx = work / "manuscript.docx"
        run_pandoc(str(binary), ["-f", "json", "-t", "docx", f"--reference-doc={template_path}", "--syntax-highlighting=none", "-o", str(docx)], data=json.dumps(rendered), cwd=source.parent)
        finish_docx(docx, template)
        report["validation"] = validate_docx(docx, expected, resolved, template)
        for path, before in tracked.items():
            require(digest(path) == before, f"Input changed during build: {path}; output not published")
        report["input_sha256"] = {str(p): v for p, v in tracked.items()}
        report["output_sha256"] = digest(docx)
        report_file = work / "build-report.json"
        report_file.write_text(json.dumps(report, indent=2, ensure_ascii=False) + "\n")
        os.replace(docx, output)
        os.replace(report_file, output.with_suffix(".build.json"))
    print(f"Created {output}\nValidated {len(citations)} live Zotero citations, {len(keys)} references, {expected['Table']} tables, and {expected['Image']} figures.")
    return report


def build_parser():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, required=True, help="Main LaTeX file (content); may be outside this folder")
    parser.add_argument("--template", type=Path, default=HERE / "templates" / "JMIR.docx", help="Word formatting and Zotero settings reference (default: bundled JMIR template)")
    parser.add_argument("--output", type=Path, help="Generated Word file (default: output/<source-stem>.docx inside the compiler folder)")
    parser.add_argument("--pandoc", help="Pandoc executable (otherwise PANDOC or PATH)")
    parser.add_argument("--csl", type=Path, help="CSL file matching the template's Zotero style (default: bundled style, then Zotero's styles folder)")
    parser.add_argument("--check", action="store_true", help="Read-only source, dependency, asset, and Zotero preflight")
    return parser


def main(argv=None):
    try:
        compile_document(build_parser().parse_args(argv))
    except (BuildError, OSError, ValueError, ET.XMLSyntaxError) as exc:
        print(f"Word build failed: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
