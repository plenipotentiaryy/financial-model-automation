from copy import deepcopy

import pytest
from openpyxl import load_workbook

import finmodel.recalculate as adapter
from finmodel.samples import ACME
from finmodel.schemas import ModelError
from finmodel.service import build_company
from test_workbook import formulas


def test_engine_unavailable_is_explicit(project, make_company, config, monkeypatch):
    monkeypatch.setattr(adapter, "detect_engine", lambda: None)
    output = project / "model.xlsx"
    manifest = build_company(make_company(), project / "samples/template.xlsx", output, config, calculate=True)
    assert manifest["outcome"] == "FAILED" and not output.exists()
    assert manifest["calculation"]["status"] == "FAILED"
    assert any(i.code == "RECALCULATION_FAILED" and "unavailable" in i.message for i in manifest["issues"])


def test_engine_timeout_is_explicit(project, make_company, config, monkeypatch):
    import subprocess

    def timeout(*args, **kwargs):
        raise subprocess.TimeoutExpired("soffice", 60)
    monkeypatch.setattr(adapter, "detect_engine", lambda: "/fake/soffice")
    monkeypatch.setattr(adapter.subprocess, "run", timeout)
    output = project / "model.xlsx"
    manifest = build_company(make_company(), project / "samples/template.xlsx", output, config, calculate=True)
    assert manifest["outcome"] == "FAILED" and not output.exists()
    assert manifest["calculation"]["status"] == "FAILED"


@pytest.mark.engine
@pytest.mark.skipif(adapter.detect_engine() is None, reason="LibreOffice unavailable; Excel-on-open mode still tested")
def test_real_engine_sample_and_changed_input(project, make_company, config):
    template = project / "samples/template.xlsx"
    output = project / "model.xlsx"
    manifest = build_company(make_company(), template, output, config, calculate=True)
    assert manifest["outcome"] != "FAILED", [i.message for i in manifest["issues"]]
    assert manifest["calculation"]["status"] == "RECALCULATED"
    baseline = load_workbook(template)
    saved = load_workbook(output)
    assert formulas(saved) == formulas(baseline)
    values = load_workbook(output, data_only=True)
    for sheet, addresses in [("Calculations", ["B2", "B3", "B4", "B5"]), ("Report", ["B7", "B8", "B9", "B10"])]:
        assert [values[sheet][address].value for address in addresses] == pytest.approx([0.25, 0.2, 30000000, 1.5])
    assert values["Report"]["B38"].value == "RECALCULATED"
    assert values["Report"]["B11"].value in {None, ""} and values["Report"]["B12"].value in {None, ""}
    assert not any(i.code == "FORMULA_NOT_RECALCULATED" for i in manifest["issues"])
    # Inspect native XML: formulas and styles stay exactly as populated by us.
    assert saved["Inputs"]["C7"]._style == baseline["Inputs"]["C7"]._style
    rows = deepcopy(ACME)
    rows[1][7] = 120
    manifest2 = build_company(make_company(rows), template, output, config, calculate=True, overwrite=True)
    assert manifest2["calculation"]["status"] == "RECALCULATED"
    changed = load_workbook(output, data_only=True)
    assert changed["Report"]["B7"].value == pytest.approx(0.5)
    assert changed["Report"]["B8"].value == pytest.approx(1 / 6)
    assert changed["Report"]["B9"].value == 30000000


@pytest.mark.engine
@pytest.mark.skipif(adapter.detect_engine() is None, reason="LibreOffice unavailable")
@pytest.mark.parametrize("index,value,blanks", [(8, "", ["B8", "B10"]), (8, -2, ["B10"]), (7, 0, ["B8"])])
def test_engine_missing_and_denominator_guards(project, make_company, config, index, value, blanks):
    rows = deepcopy(ACME)
    rows[1][index] = value
    output = project / "missing.xlsx"
    manifest = build_company(make_company(rows), project / "samples/template.xlsx", output, config, calculate=True)
    assert manifest["outcome"] != "FAILED", [i.message for i in manifest["issues"]]
    wb = load_workbook(output, data_only=True)
    for address in blanks:
        assert wb["Report"][address].value in {None, ""}


def test_bad_cache_cannot_be_marked_fresh(project):
    original = project / "samples/template.xlsx"
    with pytest.raises(ModelError):
        adapter.transplant_caches(original, original)


@pytest.mark.engine
@pytest.mark.skipif(adapter.detect_engine() is None, reason="LibreOffice unavailable")
def test_cache_transplant_preserves_formula_xml_and_other_zip_parts(project, make_company, config):
    import re
    from zipfile import ZipFile
    from openpyxl.styles import Font

    output = project / "original.xlsx"
    template = project / "samples/template.xlsx"
    wb = load_workbook(template)
    wb["Calculations"]["A50"].font = Font(name="Arial", bold=True)
    wb["Calculations"]["B50"] = "=7*8"
    wb.save(template)
    build_company(make_company(), template, output, config)
    with ZipFile(output) as archive:
        before = {name: archive.read(name) for name in archive.namelist()}
    adapter.recalculate(output)
    with ZipFile(output) as archive:
        after = {name: archive.read(name) for name in archive.namelist()}
    assert before.keys() == after.keys()
    changed_parts = []
    for name in before:
        if before[name] == after[name]:
            continue
        changed_parts.append(name)
        assert name.startswith("xl/worksheets/")
        assert re.findall(rb'<f\b[^>]*>.*?</f>', before[name]) == re.findall(rb'<f\b[^>]*>.*?</f>', after[name])
        def without_caches(xml):
            xml = re.sub(rb'<v\b[^>]*(?<!/)>.*?</v>|<v\s*/>', b"", xml)
            return re.sub(rb'\s+t="(?:n|str)"', b"", xml)
        assert without_caches(before[name]) == without_caches(after[name])
    assert len(changed_parts) == 3
    assert load_workbook(output, data_only=True)["Calculations"]["B50"].value == 56
