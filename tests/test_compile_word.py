"""Offline regression tests; integration tests need the supplied template and Pandoc."""
from copy import deepcopy
import json
import os
from pathlib import Path
import runpy
import shutil
import sys
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import compile_word as c


def record(key, item="ABCDEFGH", library="groups/5808658"):
    return {"key": item, "data": {"citationKey": key}, "links": {"alternate": {"href": f"https://www.zotero.org/{library}/items/{item}"}}}


def reference(key):
    uri = f"http://zotero.org/groups/5808658/items/{key.upper():0<8}"
    return {"key": key, "uri": uri, "item_key": key.upper(), "library": "/api/groups/5808658",
            "itemData": {"id": key, "citation-key": key, "type": "article-journal", "title": "Test " + key,
                         "author": [{"family": "Author" + key, "given": "A"}], "issued": {"date-parts": [[2020]]},
                         "container-title": "Test Journal", "volume": "1", "page": "1-2"}}


class SourceTests(unittest.TestCase):
    def test_bundled_reference_is_blank_and_generic(self):
        template = c.read_template(c.HERE / "templates" / "JMIR.docx")
        doc = template["doc"]
        body = doc.find("w:body", c.NS)
        self.assertEqual([element.tag for element in body], [c.q("p"), c.q("sectPr")])
        self.assertFalse(doc.findall(".//w:t", c.NS))
        self.assertFalse(doc.findall(".//w:instrText", c.NS))
        self.assertFalse(any("/media/" in name for name in template["parts"]))
        props = c.parse(template["parts"]["docProps/custom.xml"])
        self.assertTrue(all(p.get("name").startswith("ZOTERO_PREF_") for p in props))
        styles = c.parse(template["parts"]["word/styles.xml"])
        self.assertTrue({"Normal", "Heading1", "JMIRTable", "TableCaption", "ImageCaption", "Bibliography"}.issubset({s.get(c.q("styleId")) for s in styles}))

    def test_layout_wrappers_preserve_nested_content(self):
        source = r"\makebox[\textwidth][c]{\resizebox{2cm}{!}{\textbf{A} & B}}"
        self.assertEqual(c.preprocess(source), r"\textbf{A} & B")
        self.assertEqual(c.preprocess(r"A \\ \cline{2-5} B \\ \cmidrule(lr){1-2}"), "A \\\\  B \\\\ ")

    def test_verbatim_not_interpreted(self):
        source = "\\begin{Verbatim}\n100% literal \\input{not-a-file}\n\\makebox{literal}\n\\end{Verbatim}"
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "main.tex"
            path.write_text(source)
            self.assertEqual(c.preprocess(c.expand_source(path, {})), source)

    def test_nested_inputs_comments_and_cycles(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "main.tex").write_text("% \\input{missing}\n\\input{section}\n25\\%")
            (root / "section.tex").write_text("content")
            tracked = {}
            self.assertIn("content\n25\\%", c.expand_source(root / "main.tex", tracked))
            self.assertEqual(len(tracked), 2)
            (root / "section.tex").write_text(r"\input{main}")
            with self.assertRaisesRegex(c.BuildError, "Circular"):
                c.expand_source(root / "main.tex", {})

    def test_missing_asset(self):
        ast = {"t": "Image", "c": [["", [], []], [], ["does-not-exist.pdf", ""]]}
        with self.assertRaisesRegex(c.BuildError, "Missing figure"):
            c.collect_assets(ast, Path.cwd(), {})

    def test_broken_reference(self):
        ast = {"meta": {}, "blocks": [c.header(c.BIB_MARK), {"t": "Para", "c": [{"t": "Link", "c": [["", [], [["reference-type", "ref"]]], [], ["#missing", ""]]}]}]}
        with self.assertRaisesRegex(c.BuildError, "Unresolved cross-reference"):
            c.transform_ast(ast)

    def test_appendix_and_bibliography_order(self):
        ast = {"meta": {}, "blocks": [c.header(c.BIB_MARK), c.header(c.APP_MARK), c.header("Prompt", 1, "app:prompt")]}
        c.transform_ast(ast)
        self.assertEqual(c.text(ast["blocks"][0]["c"][2]), "References")
        self.assertEqual(c.text(ast["blocks"][2]["c"][2]), "Appendix A. Prompt")

    def test_unsupported_semantic_command(self):
        with self.assertRaisesRegex(c.BuildError, "Unsupported content"):
            c.check_ast({"t": "RawInline", "c": ["latex", r"\unknown{important text}"]})


class ZoteroTests(unittest.TestCase):
    def test_group_preferred_and_personal_fallback(self):
        z = c.Zotero()
        paths = []
        def inventory(path):
            paths.append(path)
            return [record("group")] if "groups" in path else [record("personal", "PERSONAL", "users/123"), record("group", "DUPLICAT", "users/123")]
        def get(path, **params):
            key = "personal" if "PERSONAL" in path else "group"
            return [reference(key)["itemData"]]
        with patch.object(z, "inventory", side_effect=inventory), patch.object(z, "get", side_effect=get):
            result = z.resolve({"group", "personal"})
        self.assertEqual(paths, ["/api/groups/5808658", "/api/users/0"])
        self.assertIn("/groups/5808658/", result["group"]["uri"])
        self.assertIn("/users/123/", result["personal"]["uri"])

    def test_no_personal_request_when_group_complete(self):
        z = c.Zotero()
        with patch.object(z, "inventory", return_value=[record("a")]) as inv, patch.object(z, "get", return_value=[reference("a")["itemData"]]):
            z.resolve({"a"})
        self.assertEqual(inv.call_count, 1)

    def test_duplicate_exact_key_fails(self):
        z = c.Zotero()
        with patch.object(z, "inventory", return_value=[record("a"), record("a", "OTHERKEY")]):
            with self.assertRaisesRegex(c.BuildError, "Ambiguous"):
                z.resolve({"a"})

    def test_missing_key_is_not_fuzzy_matched(self):
        z = c.Zotero()
        with patch.object(z, "inventory", return_value=[record("ab")]):
            with self.assertRaisesRegex(c.BuildError, "No exact"):
                z.resolve({"a"})

    def test_connection_failure_does_not_fall_back(self):
        z = c.Zotero()
        with patch.object(z, "inventory", side_effect=c.BuildError("unavailable")) as inv:
            with self.assertRaisesRegex(c.BuildError, "unavailable"):
                z.resolve({"a"})
        self.assertEqual(inv.call_count, 1)

    def test_inventory_paginates(self):
        z = c.Zotero()
        with patch.object(z, "get", side_effect=[[record("a")] * 100, [record("b")]]) as get:
            self.assertEqual(len(z.inventory("/api/groups/5808658")), 101)
        self.assertEqual(get.call_args_list[1].kwargs["start"], 100)

    def test_request_is_get_and_disables_redirects(self):
        z = c.Zotero()
        with patch.object(z.opener, "open") as op:
            op.return_value.__enter__.return_value.read.return_value = b"[]"
            z.get("/api/groups/5808658/items/top", limit=100)
        self.assertEqual(op.call_args.args[0].get_method(), "GET")
        with self.assertRaisesRegex(c.BuildError, "redirected"):
            c.NoRedirect().redirect_request(None, None, 302, "", {}, "https://example.com")


class WordIntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        sibling = Path(sys.executable).with_name("pandoc")
        cls.pandoc = os.environ.get("PANDOC") or shutil.which("pandoc") or (str(sibling) if sibling.is_file() else None)
        cls.template = c.HERE / "templates" / "JMIR.docx"
        cls.csl = c.HERE / "styles" / "journal-of-medical-internet-research.csl"
        if not cls.pandoc or not cls.template.exists() or not cls.csl.exists():
            raise unittest.SkipTest("Set PANDOC and retain the bundled template and style for integration tests")

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name).resolve()
        self.source = self.root / "main.tex"
        self.output = self.root / "result.docx"
        self.source.write_text(r"""\documentclass{article}
\title{Test manuscript}\author{Anonymous Authors}
\begin{document}\maketitle
\begin{abstract}Preserve this abstract.\end{abstract}
\section{Introduction}
First \cite{b}. Group \cite{a,b}. Repeat \cite{b}.
\bibliography{references}
\appendix\section{Prompt}\label{app:prompt}
\begin{verbatim}
100% literal and {braces}
second line
\end{verbatim}
\end{document}
""")
        self.args = c.argparse.Namespace(source=self.source, template=self.template, output=self.output, pandoc=self.pandoc, csl=self.csl, check=False)

    def build(self):
        with patch.object(c.Zotero, "resolve", side_effect=lambda keys: {k: reference(k) for k in keys}):
            return c.compile_document(self.args)

    def test_native_fields_reorder_group_repeat_and_new_reference(self):
        before = c.digest(self.template)
        report = self.build()
        self.assertEqual([x["text"] for x in report["citations"]], ["[1]", "[1,2]", "[1]"])
        self.assertEqual(report["validation"]["citation_fields"], 3)
        self.assertEqual(c.digest(self.template), before)
        with ZipFile(self.output) as z:
            doc = c.parse(z.read("word/document.xml"))
            visible = "".join(doc.xpath(".//w:t/text()", namespaces=c.NS))
        self.assertIn("100% literal and {braces}", visible)
        self.assertLess(visible.index("References"), visible.index("Appendix A."))
        self.source.write_text(self.source.read_text().replace("First \\cite{b}.", "First \\cite{a}. New \\cite{c}."))
        report = self.build()
        self.assertEqual(report["validation"]["references"], 3)
        self.assertEqual(report["citations"][-1]["text"], "[3]")

    def test_failed_build_preserves_existing_output(self):
        self.output.write_bytes(b"previous successful output")
        with patch.object(c.Zotero, "resolve", side_effect=c.BuildError("unavailable")):
            with self.assertRaises(c.BuildError):
                c.compile_document(self.args)
        self.assertEqual(self.output.read_bytes(), b"previous successful output")
        with patch.object(c.Zotero, "resolve", side_effect=lambda keys: {k: reference(k) for k in keys}), patch.object(c, "validate_docx", side_effect=c.BuildError("invalid DOCX")):
            with self.assertRaisesRegex(c.BuildError, "invalid DOCX"):
                c.compile_document(self.args)
        self.assertEqual(self.output.read_bytes(), b"previous successful output")

    def test_check_writes_nothing(self):
        self.args.check = True
        before = sorted(p.name for p in self.root.iterdir())
        self.build()
        self.assertEqual(sorted(p.name for p in self.root.iterdir()), before)

    def test_relocated_folder_builds_with_external_source(self):
        relocated = self.root / "moved compiler"
        relocated.mkdir()
        shutil.copy2(c.HERE / "compile_word.py", relocated)
        for folder in ("templates", "styles"):
            shutil.copytree(c.HERE / folder, relocated / folder)
        module = runpy.run_path(str(relocated / "compile_word.py"))
        args = module["build_parser"]().parse_args(["--source", str(self.source), "--pandoc", self.pandoc])
        self.assertEqual(args.template, relocated / "templates" / "JMIR.docx")
        with patch.object(module["Zotero"], "resolve", side_effect=lambda keys: {k: reference(k) for k in keys}):
            report = module["compile_document"](args)
        self.assertEqual(Path(report["output"]), relocated / "output" / "main.docx")
        self.assertEqual(report["validation"]["citation_fields"], 3)
        self.assertTrue((relocated / "output" / "main.build.json").is_file())

    def test_table_layout_commands_and_verbatim_alignment(self):
        table = r"""\begin{table}
\caption{A small table}\label{tab:test}
\makebox[\textwidth][c]{\begin{tabular}{lrr}
\hline
Model & \multicolumn{2}{c}{Scores} \\
\cline{2-3}
 & First & Second \\
\hline
A & 0.71 & 0.82 \\
\hline
\end{tabular}}
\end{table}
See Table~\ref{tab:test}.
"""
        self.source.write_text(self.source.read_text().replace(r"\bibliography", table + r"\bibliography"))
        report = self.build()
        self.assertEqual(report["validation"]["tables"], 1)
        with ZipFile(self.output) as z:
            doc = c.parse(z.read("word/document.xml"))
        tbl = doc.find(".//w:tbl", c.NS)
        table_text = "".join(tbl.xpath(".//w:t/text()", namespaces=c.NS))
        self.assertNotIn("2-3", table_text)
        self.assertIn("0.71", table_text)
        self.assertEqual(len(tbl.findall("w:tr", c.NS)), 3)
        self.assertEqual(tbl.find("w:tblPr/w:tblStyle", c.NS).get(c.q("val")), "JMIRTable")
        self.assertEqual(tbl.find(".//w:r/w:rPr/w:sz", c.NS).get(c.q("val")), "16")
        code = doc.find(".//w:pPr/w:pStyle[@w:val='SourceCode']", c.NS)
        self.assertEqual(code.getparent().find("w:jc", c.NS).get(c.q("val")), "left")


if __name__ == "__main__":
    unittest.main()
