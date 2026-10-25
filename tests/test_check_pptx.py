"""Synthetic contract tests for the unmodified v6 PPTX inspector (standard library only)."""

from pathlib import Path
import hashlib
import json
import locale
import platform
import runpy
import subprocess
import sys
import tempfile
import unittest
import warnings
import zipfile


HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
SCRIPT = ROOT / "skills/research-ppt-design-v6/scripts/check_pptx.py"
EXPECTED_SHA256 = "511ada5b75ba4777f7e31b33c8e7acc10e34f5e30cf921f3c6adde51b881690b"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def slide(text="Example / not research results", hidden=False, objects=False):
    flag = ' show="0"' if hidden else ""
    shape = f"<p:sp><p:txBody><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>" if text is not None else ""
    extras = "<p:pic/><p:cxnSp/><p:graphicFrame/>" if objects else ""
    return f'<p:sld xmlns:p="{P_NS}" xmlns:a="{A_NS}"{flag}><p:cSld><p:spTree>{shape}{extras}</p:spTree></p:cSld></p:sld>'


def rels(entries=()):
    return f'<Relationships xmlns="{REL_NS}">' + "".join(entries) + "</Relationships>"


def base_parts():
    return {
        "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>',
        "_rels/.rels": rels(['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>']),
        "ppt/presentation.xml": f'<p:presentation xmlns:p="{P_NS}"/>',
        "ppt/_rels/presentation.xml.rels": rels(['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>']),
        "ppt/slides/slide1.xml": slide(),
    }


class InspectorContract(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        assert digest(SCRIPT) == EXPECTED_SHA256
        cls.script_before = (digest(SCRIPT), SCRIPT.stat().st_mtime_ns)
        cls.temporary = tempfile.TemporaryDirectory(prefix="synthetic-", dir=HERE)
        cls.directory = Path(cls.temporary.name)
        cls.inspect = staticmethod(runpy.run_path(str(SCRIPT), run_name="contract_import")["inspect"])
        cls.sequence = 0
        cls.readonly_checks = 0

    @classmethod
    def tearDownClass(cls):
        assert (digest(SCRIPT), SCRIPT.stat().st_mtime_ns) == cls.script_before
        cls.temporary.cleanup()

    def make(self, parts=None, duplicate=None):
        type(self).sequence += 1
        path = self.directory / f"synthetic-{self.sequence}.pptx"
        with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_STORED) as z:
            for name, content in (base_parts() if parts is None else parts).items():
                z.writestr(name, content)
            if duplicate:
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore", UserWarning)
                    z.writestr(duplicate, "duplicate synthetic payload")
        return path

    def call(self, path, *args):
        before = {p.name: (digest(p), p.stat().st_mtime_ns) for p in self.directory.iterdir() if p.is_file()}
        proc = subprocess.run([sys.executable, str(SCRIPT), str(path), *args], capture_output=True, text=True, timeout=20)
        after = {p.name: (digest(p), p.stat().st_mtime_ns) for p in self.directory.iterdir() if p.is_file()}
        self.assertEqual(before, after, "Inspector changed input content, timestamps or created files")
        type(self).readonly_checks += 1
        return proc

    def json_call(self, path, expected=0, *args):
        proc = self.call(path, "--json", *args)
        self.assertEqual(proc.returncode, expected, proc.stderr or proc.stdout)
        self.assertEqual(proc.stderr, "")
        return json.loads(proc.stdout)

    def test_01_source_hash(self):
        self.assertEqual(digest(SCRIPT), EXPECTED_SHA256)

    def test_02_cli_help(self):
        proc = self.call(self.directory / "unused.pptx", "--help")
        self.assertEqual(proc.returncode, 0)
        self.assertIn("--max-bytes", proc.stdout)
        self.assertIn("--json", proc.stdout)

    def test_03_json_pass_and_direct_inspect(self):
        path = self.make()
        report = self.json_call(path)
        self.assertTrue(report["structural_check_passed"])
        self.assertEqual(report["slide_part_count"], 1)
        self.assertEqual(report["slides"][0]["text_runs"], 1)
        self.assertEqual(report["slides"][0]["text_characters"], len("Example / not research results"))
        self.assertEqual(report["issues"], [])
        self.assertIsNone(report["limit_bytes"])
        self.assertIsNone(report["remaining_bytes"])
        before = (digest(path), path.stat().st_mtime_ns)
        self.assertEqual(self.inspect(path), report)
        self.assertEqual((digest(path), path.stat().st_mtime_ns), before)

    def test_04_default_has_no_size_budget(self):
        parts = base_parts()
        parts["ppt/media/synthetic.bin"] = b"x" * (1024 * 1024 + 1)
        report = self.json_call(self.make(parts))
        self.assertTrue(report["structural_check_passed"])
        self.assertGreater(report["file_bytes"], 1024 * 1024)
        self.assertIsNone(report["limit_bytes"])

    def test_05_budget_exact_pass(self):
        path = self.make()
        report = self.json_call(path, 0, "--max-bytes", str(path.stat().st_size))
        self.assertEqual(report["remaining_bytes"], 0)

    def test_06_budget_headroom_pass(self):
        path = self.make()
        report = self.json_call(path, 0, "--max-bytes", str(path.stat().st_size + 123))
        self.assertEqual(report["remaining_bytes"], 123)

    def test_07_budget_exceeded(self):
        path = self.make()
        report = self.json_call(path, 1, "--max-bytes", str(path.stat().st_size - 1))
        self.assertEqual(report["remaining_bytes"], -1)
        self.assertEqual(report["issues"], ["File exceeds budget by 1 bytes."])

    def test_08_missing_required_parts(self):
        for required in ("[Content_Types].xml", "_rels/.rels", "ppt/presentation.xml"):
            with self.subTest(required=required):
                parts = base_parts()
                del parts[required]
                report = self.json_call(self.make(parts), 1)
                self.assertIn(f"Required part missing: {required}", report["issues"])

    def test_09_missing_declared_relationship_target(self):
        parts = base_parts()
        parts["ppt/slides/_rels/slide1.xml.rels"] = rels(['<Relationship Id="missing" Target="../media/absent.bin"/>'])
        report = self.json_call(self.make(parts), 1)
        self.assertIn("Missing relationship target: ppt/slides/_rels/slide1.xml.rels -> ../media/absent.bin", report["issues"])

    def test_10_external_relationship_is_not_fetched(self):
        parts = base_parts()
        parts["ppt/slides/_rels/slide1.xml.rels"] = rels(['<Relationship Id="external" Target="https://example.invalid/not-fetched" TargetMode="External"/>'])
        self.assertTrue(self.json_call(self.make(parts))["structural_check_passed"])

    def test_11_absolute_and_fragment_target_resolution(self):
        parts = base_parts()
        parts["ppt/_rels/presentation.xml.rels"] = rels(['<Relationship Id="absolute" Target="/ppt/slides/slide1.xml#example"/>'])
        self.assertTrue(self.json_call(self.make(parts))["structural_check_passed"])

    def test_12_no_slide_parts(self):
        parts = base_parts()
        del parts["ppt/slides/slide1.xml"]
        parts["ppt/_rels/presentation.xml.rels"] = rels()
        report = self.json_call(self.make(parts), 1)
        self.assertIn("No slide XML parts found.", report["issues"])

    def test_13_no_native_text_warning_is_not_failure(self):
        parts = base_parts()
        parts["ppt/slides/slide1.xml"] = slide(None)
        report = self.json_call(self.make(parts))
        self.assertEqual(len(report["warnings"]), 1)
        self.assertIn("no native text runs", report["warnings"][0])

    def test_14_embedded_fonts_and_files_warning(self):
        parts = base_parts()
        parts["ppt/embeddings/synthetic.bin"] = b"example only"
        parts["ppt/fonts/synthetic.fntdata"] = b"not an actual font"
        report = self.json_call(self.make(parts))
        self.assertEqual(len(report["embedded_parts"]), 2)
        self.assertIn("Embedded files/fonts present", report["warnings"][0])

    def test_15_object_counts_hidden_flag_and_numeric_slide_order(self):
        parts = base_parts()
        parts["ppt/slides/slide10.xml"] = slide("Ten")
        parts["ppt/slides/slide2.xml"] = slide("Two", hidden=True, objects=True)
        report = self.json_call(self.make(parts))
        self.assertEqual([s["part"] for s in report["slides"]], ["ppt/slides/slide1.xml", "ppt/slides/slide2.xml", "ppt/slides/slide10.xml"])
        second = report["slides"][1]
        for field in ("shapes", "pictures", "connectors", "graphic_frames"):
            self.assertEqual(second[field], 1)
        self.assertTrue(second["hidden"])

    def test_16_duplicate_zip_part(self):
        report = self.json_call(self.make(duplicate="ppt/media/duplicate.bin"))
        self.assertTrue(report["structural_check_passed"])
        parts = base_parts()
        parts["ppt/media/duplicate.bin"] = b"first synthetic payload"
        report = self.json_call(self.make(parts, duplicate="ppt/media/duplicate.bin"), 1)
        self.assertIn("ZIP contains duplicate entry names.", report["issues"])

    def test_17_corrupted_media_crc(self):
        payload = b"UNIQUE_SYNTHETIC_CRC_PAYLOAD_123456789"
        parts = base_parts()
        parts["ppt/media/crc.bin"] = payload
        path = self.make(parts)
        raw = path.read_bytes()
        self.assertEqual(raw.count(payload), 1)
        path.write_bytes(raw.replace(payload, b"X" + payload[1:]))
        report = self.json_call(path, 1)
        self.assertIn("ZIP CRC failure: ppt/media/crc.bin", report["issues"])

    def test_18_bad_zip(self):
        path = self.directory / "not-a-zip.pptx"
        path.write_bytes(b"Synthetic invalid ZIP file")
        proc = self.call(path, "--json")
        self.assertEqual(proc.returncode, 2)
        self.assertIn("Cannot inspect file:", proc.stderr)
        self.assertEqual(proc.stdout, "")

    def test_19_malformed_slide_and_relationship_xml(self):
        for name in ("ppt/slides/slide1.xml", "ppt/_rels/presentation.xml.rels"):
            with self.subTest(name=name):
                parts = base_parts()
                parts[name] = "<invalid"
                proc = self.call(self.make(parts), "--json")
                self.assertEqual(proc.returncode, 2)
                self.assertIn("Cannot inspect file:", proc.stderr)

    def test_20_missing_input_file(self):
        proc = self.call(self.directory / "missing.pptx", "--json")
        self.assertEqual(proc.returncode, 2)
        self.assertIn("Cannot inspect file:", proc.stderr)

    def test_21_nonpositive_budget_rejected(self):
        path = self.make()
        for limit in ("0", "-1"):
            with self.subTest(limit=limit):
                proc = self.call(path, "--max-bytes", limit)
                self.assertEqual(proc.returncode, 2)
                self.assertIn("--max-bytes must be positive", proc.stderr)

    def test_22_human_readable_mode(self):
        proc = self.call(self.make())
        self.assertEqual(proc.returncode, 0)
        self.assertIn("no size limit set", proc.stdout)
        self.assertIn("Structural/budget result: PASS", proc.stdout)
        self.assertIn("require separate review", proc.stdout)

    def test_23_documented_boundary_not_full_opc_validation(self):
        parts = base_parts()
        parts["[Content_Types].xml"] = "not valid XML"
        parts["ppt/presentation.xml"] = "not valid XML"
        report = self.json_call(self.make(parts))
        self.assertTrue(report["structural_check_passed"])
        self.assertIn("Target-application editability and compatibility", report["not_checked"])


if __name__ == "__main__":
    suite = unittest.defaultTestLoader.loadTestsFromTestCase(InspectorContract)
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    report = {
        "python": platform.python_version(),
        "platform": platform.system(),
        "cli_text_encoding": locale.getpreferredencoding(False),
        "source_relative_path": "skills/research-ppt-design-v6/scripts/check_pptx.py",
        "source_sha256": digest(SCRIPT),
        "source_hash_matches": digest(SCRIPT) == EXPECTED_SHA256,
        "tests_run": result.testsRun,
        "failures": len(result.failures),
        "errors": len(result.errors),
        "skipped": len(result.skipped),
        "cli_readonly_checks": InspectorContract.readonly_checks,
        "passed": result.wasSuccessful(),
        "fixtures": "Generated locally from original synthetic strings; removed after the suite.",
        "boundaries": [
            "Fixtures test the inspector contract; they are not rendered presentation examples.",
            "A pass does not establish OPC/XML schema conformance or that PowerPoint can open the file.",
            "Content_Types and presentation XML are checked for existence, not parsed.",
            "Slide XML parts are discovered by filename, not the presentation's relationship/order list.",
            "Visual quality, science, citations, fonts, application editability and external URLs are not validated.",
        ],
    }
    (HERE / "test-results.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    sys.exit(0 if result.wasSuccessful() else 1)
