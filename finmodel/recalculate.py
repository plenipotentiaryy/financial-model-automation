"""Optional LibreOffice adapter. Preserve original formulas and styles in the ZIP."""

import os
import posixpath
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from .schemas import ModelError

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"s": MAIN}


def detect_engine() -> str | None:
    candidate = shutil.which("soffice") or shutil.which("libreoffice")
    mac_path = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    return candidate or (str(mac_path) if mac_path.is_file() else None)


def sheet_paths(archive):
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {r.attrib["Id"]: r.attrib["Target"] for r in rels}
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    result = {}
    for sheet in workbook.find("s:sheets", NS):
        target = targets[sheet.attrib[f"{{{REL}}}id"]]
        result[sheet.attrib["name"]] = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
    return result


def _canonical_formula(text):
    # LibreOffice may add sheet quotes and absolute-address markers.
    # Only these harmless serialisation changes are accepted.
    return re.sub(r"'([A-Za-z_][A-Za-z_0-9]*)'!", r"\1!", text or "").replace("$", "")


def transplant_caches(original: Path, calculated: Path):
    """Validate the calculation copy, then copy only <v>/<is> and result type."""
    modified = {}
    with ZipFile(original) as baseline, ZipFile(calculated) as engine:
        original_sheets, engine_sheets = sheet_paths(baseline), sheet_paths(engine)
        strings = []
        if "xl/sharedStrings.xml" in engine.namelist():
            for item in ET.fromstring(engine.read("xl/sharedStrings.xml")):
                strings.append("".join(item.itertext()))
        count = 0
        for name, xml_path in original_sheets.items():
            original_xml = baseline.read(xml_path)
            root = ET.fromstring(original_xml)
            formula_cells = [c for c in root.findall(".//s:c", NS) if c.find("s:f", NS) is not None]
            if not formula_cells:
                continue
            if name not in engine_sheets:
                raise ModelError(f"Recalculation removed worksheet {name}")
            recalculated = ET.fromstring(engine.read(engine_sheets[name]))
            engine_cells = {c.attrib["r"]: c for c in recalculated.findall(".//s:c", NS)}
            replacements = {}
            for cell in formula_cells:
                address = cell.attrib["r"]
                result = engine_cells.get(address)
                if result is None or result.find("s:f", NS) is None:
                    raise ModelError(f"Recalculation lost formula {name}!{address}")
                formula = cell.find("s:f", NS)
                engine_formula = result.find("s:f", NS)
                if any(key in f.attrib for f in (formula, engine_formula) for key in ("ref", "si")) or any(f.get("t", "normal") != "normal" for f in (formula, engine_formula)):
                    raise ModelError("Headless recalculation supports scalar formulas, not shared/array formulas")
                if _canonical_formula(formula.text) != _canonical_formula(engine_formula.text):
                    raise ModelError(f"Recalculation changed formula semantics: {name}!{address}")
                result_type = result.get("t")
                if result_type == "e":
                    raise ModelError(f"Spreadsheet error at {name}!{address}: {result.findtext('s:v', namespaces=NS)}")
                cached = result.find("s:v", NS)
                inline = result.find("s:is", NS)
                if cached is None and inline is None and result_type not in {"str", "inlineStr"}:
                    raise ModelError(f"Missing formula cache at {name}!{address}")
                if result_type not in {"str", "inlineStr"} and cached is not None and cached.text is None:
                    raise ModelError(f"Empty numeric formula cache at {name}!{address}")
                if result_type == "s":
                    if cached is None or cached.text is None:
                        raise ModelError(f"Missing shared-string cache at {name}!{address}")
                    try:
                        index = int(cached.text)
                        if index < 0:
                            raise IndexError("Negative string index")
                        text = strings[index]
                    except (ValueError, IndexError) as exc:
                        raise ModelError("Invalid shared-string formula cache") from exc
                    cached = ET.Element(f"{{{MAIN}}}v")
                    cached.text = text
                    result_type = "str"
                elif result_type == "inlineStr":
                    cached = ET.Element(f"{{{MAIN}}}v")
                    cached.text = "".join(inline.itertext()) if inline is not None else ""
                    result_type = "str"
                elif cached is None:
                    cached = ET.Element(f"{{{MAIN}}}v")
                    cached.text = ""
                replacements[address] = (result_type, cached)
                count += 1
            # Patch each formula cell in the ORIGINAL XML, preserving namespaces,
            # all other cells, original formula bytes and every other ZIP part.
            # Exclude self-closing styled blanks: consuming one would swallow
            # the following formula cell and leave its cache unpatched.
            pattern = rb'<c\b[^>]*\br="([^"]+)"[^>]*(?<!/)>.*?</c>'
            patched = set()

            def patch(match):
                address = match.group(1).decode()
                if address not in replacements:
                    return match.group(0)
                patched.add(address)
                value_type, cached = replacements[address]
                raw = match.group(0)
                tag_end = raw.index(b">") + 1
                header = re.sub(rb'\s+t="[^"]*"', b"", raw[:tag_end])
                if value_type and value_type != "n":
                    header = header[:-1] + f' t="{value_type}">'.encode()
                body = re.sub(rb'<v\b[^>]*(?<!/)>.*?</v>|<v\s*/>|<is\b[^>]*(?<!/)>.*?</is>', b"", raw[tag_end:-4], flags=re.S)
                # XML-escape cache text without adding a namespace prefix.
                new_value = ET.Element("v")
                new_value.text = cached.text
                return header + body + ET.tostring(new_value, encoding="utf-8") + b"</c>"

            modified[xml_path] = re.sub(pattern, patch, original_xml, flags=re.S)
            if patched != set(replacements):
                raise ModelError(f"Not all formula caches could be transplanted on {name}")
        if count == 0:
            raise ModelError("No formulas found to recalculate")
        temp_path = original.with_suffix(".cached.xlsx")
        try:
            with ZipFile(temp_path, "w") as output:
                for member in baseline.infolist():
                    output.writestr(member, modified.get(member.filename, baseline.read(member.filename)))
        except Exception:
            temp_path.unlink(missing_ok=True)
            raise
    os.replace(temp_path, original)
    return count


def recalculate(path: Path, timeout: int = 60):
    executable = detect_engine()
    if executable is None:
        raise ModelError("--recalculate requested but LibreOffice is unavailable")
    with tempfile.TemporaryDirectory(prefix="finmodel-calc-") as directory:
        directory = Path(directory)
        source_dir, output_dir = directory / "input", directory / "output"
        source_dir.mkdir()
        output_dir.mkdir()
        source = source_dir / "model.xlsx"
        shutil.copyfile(path, source)
        command = [executable, f"-env:UserInstallation={(directory / 'profile').as_uri()}", "--headless", "--convert-to", "xlsx", "--outdir", str(output_dir), str(source)]
        try:
            process = subprocess.run(command, capture_output=True, text=True, timeout=timeout, check=False)
        except (OSError, subprocess.TimeoutExpired) as exc:
            raise ModelError(f"LibreOffice recalculation failed: {exc}") from exc
        output = output_dir / source.name
        if process.returncode != 0 or not output.is_file():
            detail = (process.stderr or process.stdout).strip()[:500]
            raise ModelError(f"LibreOffice did not produce a recalculated workbook: {detail}")
        count = transplant_caches(path, output)
    return {"status": "RECALCULATED", "engine": executable, "cached_formula_cells": count}
