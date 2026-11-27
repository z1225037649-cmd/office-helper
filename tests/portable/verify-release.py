"""Opt-in artifact verification; never builds, registers hosts, or changes trust."""

import datetime
import hashlib
import io
import json
import pathlib
import posixpath
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET


ROOT = pathlib.Path(__file__).resolve().parents[2]
NAME = "OfficeHelper-0.2.2-win-x64"
ARCHIVE = ROOT / "output/portable" / (NAME + ".zip")
MANIFEST = ROOT / "output/verification/five-gaps-022-dist.json"
PRE = ROOT / "output/verification/portable-022-prepackage.json"
REPORT = ROOT / "output/verification/portable-022-artifact.json"
OFFICE_ID = "f5202c69-1f59-4e09-b510-829a3748a31b"
NS = {
    "o": "http://schemas.microsoft.com/office/appforoffice/1.1",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "we": "http://schemas.microsoft.com/office/webextensions/webextension/2010/11",
    "r": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def sha(data):
    return hashlib.sha256(data).hexdigest()


def inventory(folder):
    result = {}
    for path in sorted(folder.rglob("*")):
        assert not path.is_symlink() and not path.is_junction(), "Unexpected link"
        if path.is_file():
            result[path.relative_to(folder).as_posix()] = sha(path.read_bytes())
    return result


def safe_content(data, label):
    patterns = [
        r"native-check\.html|__native-report|x-helper-test|OFFICE_HELPER_NATIVE_TEST|native-token",
        r"extended-check\.html|__extended-report|OFFICE_HELPER_EXTENDED_TEST|extended-token",
        r"sourceMappingURL\s*=|[?&]token\s*=",
        r"-----BEGIN (?:[A-Z0-9]+ )*PRIVATE KEY-----",
        r"(?:^|[^a-f0-9])[a-f0-9]{48}(?![a-f0-9])",
    ]
    for encoding in ["utf-8", "utf-16-le"]:
        text = data.decode(encoding, errors="ignore")
        assert not any(re.search(p, text, re.I) for p in patterns), (
            "Forbidden diagnostic or private material: " + label
        )


def safe_archive(archive, prefix=None):
    names = archive.namelist()
    assert len(names) == len(set(names)), "Duplicate ZIP entry"
    assert len(names) == len({n.casefold() for n in names}), "Case-colliding ZIP entry"
    for entry in archive.infolist():
        name = entry.filename
        path = pathlib.PurePosixPath(name)
        assert "\\" not in name and ":" not in name and "\x00" not in name
        assert not path.is_absolute() and ".." not in path.parts
        assert (entry.external_attr >> 16) & 0o170000 != 0o120000, "ZIP symlink"
        assert not entry.flag_bits & 1, "Encrypted ZIP entry"
        if prefix:
            assert name.startswith(prefix), "Unexpected ZIP root"
    assert archive.testzip() is None, "ZIP CRC32 failed"


def sample(data, extension):
    with zipfile.ZipFile(io.BytesIO(data)) as nested:
        safe_archive(nested)
        contents = {n: nested.read(n) for n in nested.namelist() if not n.endswith("/")}
    for name, value in contents.items():
        safe_content(value, extension + ":" + name)
        if name.endswith((".xml", ".rels")):
            ET.fromstring(value)
        if name.endswith(".rels"):
            for rel in ET.fromstring(value):
                if rel.get("TargetMode") == "External":
                    continue
                target = rel.get("Target")
                base = posixpath.dirname(posixpath.dirname(name))
                # OPC relationships may start at the package root with '/'.
                resolved = posixpath.normpath(posixpath.join(base, target)).lstrip("/")
                assert not resolved.startswith("../") and resolved in contents, (
                    "Broken sample relationship: " + name
                )
    refs = []
    for name, value in contents.items():
        if name.endswith("/webextension.xml"):
            refs.extend(ET.fromstring(value).findall("we:reference", NS))
    assert len(refs) == 1
    assert refs[0].get("id") == OFFICE_ID and refs[0].get("version") == "1.0.2.2"
    if extension == "docx":
        document = ET.fromstring(contents["word/document.xml"])
        paragraphs = document.findall("w:body/w:p", NS)
        assert len(paragraphs) == 2
        for paragraph in paragraphs:
            spacing = paragraph.find("w:pPr/w:spacing", NS)
            indent = paragraph.find("w:pPr/w:ind", NS)
            assert spacing is not None and indent is not None
            assert spacing.get("{" + NS["w"] + "}line") == "240"
            assert spacing.get("{" + NS["w"] + "}lineRule") == "auto"
            assert indent.get("{" + NS["w"] + "}firstLine") == "0"
        return {"crc32": True, "relationships": True, "paragraphs": 2,
                "line": 240, "lineRule": "auto", "firstLine": 0,
                "embeddedManifestVersion": "1.0.2.2"}
    worksheets = [n for n in contents if re.fullmatch(r"xl/worksheets/[^/]+\.xml", n)]
    assert len(worksheets) == 1
    sheet = ET.fromstring(contents[worksheets[0]])
    rows = sheet.findall("s:sheetData/s:row", NS)
    cells = sheet.findall("s:sheetData/s:row/s:c", NS)
    expected = [f"{column}{row}" for row in range(1, 6) for column in "AB"]
    assert len(rows) == 5 and [c.get("r") for c in cells] == expected
    values = [float(c.find("s:v", NS).text) for c in cells if c.find("s:v", NS) is not None]
    assert values == [120, 85, 45.5, 200]
    assert sum(values) == 450.5 and sum(values) / len(values) == 112.625
    dimension = sheet.find("s:dimension", NS)
    assert dimension is None or dimension.get("ref") == "A1:B5"
    return {"crc32": True, "relationships": True, "range": "A1:B5", "rows": 5,
            "cells": 10, "values": values, "sum": sum(values), "average": 112.625,
            "dimensionPresent": dimension is not None, "embeddedManifestVersion": "1.0.2.2"}


pre = json.loads(PRE.read_text(encoding="utf-8"))
assert sha(MANIFEST.read_bytes()) == pre["distManifestSha256"]
expected = {e["path"].replace("\\", "/"): e["sha256"].lower()
            for e in json.loads(MANIFEST.read_text(encoding="utf-8-sig"))}
assert inventory(ROOT / "dist") == expected == pre["dist"], "Frozen dist mismatch"
old = ROOT / "output/portable/OfficeHelper-0.2.1-win-x64"
assert inventory(old) == pre["old021"]["folder"], "Old directory changed"
assert sha(pathlib.Path(str(old) + ".zip").read_bytes()) == pre["old021"]["zip"]
assert sha(pathlib.Path(str(old) + ".zip.sha256").read_bytes()) == pre["old021"]["checksumFile"]
payload = ARCHIVE.read_bytes()
checksum = pathlib.Path(str(ARCHIVE) + ".sha256").read_text().strip()
assert checksum == sha(payload) + "  " + NAME + ".zip"
with zipfile.ZipFile(io.BytesIO(payload)) as archive:
    safe_archive(archive, NAME + "/")
    files = {n[len(NAME) + 1:]: archive.read(n) for n in archive.namelist() if not n.endswith("/")}
    fixed = {
        "build-info.json", "使用说明.txt", "办公小帮手.exe", "办公小帮手.exe.config",
        "app/certificates.ps1", "app/controller.ps1", "app/helper.ico", "app/server.mjs",
        "app/site-policy.mjs", "manifests/office.xml", "runtime/node.exe",
        "samples/OfficeHelper-Sample.docx", "samples/OfficeHelper-Sample.xlsx",
    }
    for license_name in ["Node", "react", "react-dom", "scheduler", "lucide-react",
                         "office-addin-dev-settings"]:
        fixed.add("licenses/" + license_name + "-LICENSE.txt")
    fixed.add("licenses/THIRD-PARTY.txt")
    assert set(files) == fixed | {"site/" + n for n in expected}, "Unexpected payload file list"
    assert inventory(ROOT / "output/portable" / NAME) == {n: sha(b) for n, b in files.items()}
    info = json.loads(files["build-info.json"])
    assert info["version"] == "0.2.2" and info["platform"] == "win32-x64"
    site_hashes = {f["path"]: f["sha256"] for f in info["siteFiles"]}
    assert len(info["siteFiles"]) == len(site_hashes) == len(expected)
    transforms = []
    for name in expected:
        source = (ROOT / "dist" / name).read_bytes()
        expected_bytes = source
        component = name.split("/")[1] if name.startswith("wps/") else ""
        if component in ["writer", "sheet"] and name.endswith("/ribbon.xml"):
            expected_bytes = source.decode("utf-8").replace("OfficeHelper", "OfficeHelperPortable").replace(
                'label="小帮手"', 'label="小帮手（便携版）"', 1).encode("utf-8")
            transforms.append(name)
        elif component in ["writer", "sheet"] and name.endswith("/manifest.xml"):
            host = "OfficeHelperPortable" + ("Writer" if component == "writer" else "Sheet")
            expected_bytes = ('<?xml version="1.0" encoding="UTF-8"?><JsPlugin><ApiVersion>1.0.0</ApiVersion>'
                              f'<Name>{host}</Name><Description>办公小帮手便携版</Description></JsPlugin>').encode("utf-8")
            transforms.append(name)
        assert files["site/" + name] == expected_bytes, "Site mismatch: " + name
        assert sha(expected_bytes) == site_hashes[name], "Build-info mismatch: " + name
    for name in ["server.mjs", "site-policy.mjs", "controller.ps1", "certificates.ps1"]:
        assert files["app/" + name] == (ROOT / "portable" / name).read_bytes()
    assert files["使用说明.txt"] == (ROOT / "portable/使用说明.txt").read_bytes()
    assert "候选试用" in files["使用说明.txt"].decode("utf-8")
    for name, data in files.items():
        if name.endswith((".js", ".mjs", ".ps1", ".html", ".xml", ".txt", ".json", ".config", ".css", ".svg")):
            safe_content(data, name)
    office = ET.fromstring(files["manifests/office.xml"])
    assert office.find("o:Version", NS).text == "1.0.2.2"
    assert office.find("o:Id", NS).text == OFFICE_ID
    locations = [e.get("DefaultValue") for e in office.iter() if e.get("DefaultValue", "").startswith("https:")]
    assert locations and all(v.startswith("https://localhost:18766/") for v in locations)
    assert not re.search(rb"localhost:(3443|5173)\b", files["manifests/office.xml"])
    word = sample(files["samples/OfficeHelper-Sample.docx"], "docx")
    excel = sample(files["samples/OfficeHelper-Sample.xlsx"], "xlsx")
    extraction = pathlib.Path(tempfile.mkdtemp(prefix="OfficeHelper-022-中文 空格-"))
    archive.extractall(extraction)

report = {
    "checkedAt": datetime.datetime.now(datetime.timezone.utc).isoformat(),
    "archive": str(ARCHIVE), "bytes": len(payload), "sha256": sha(payload),
    "files": len(files), "crc32": True, "checksumMatches": True,
    "zipPathsSafeAndUnique": True, "folderMatchesZip": True,
    "frozenDistManifest": str(MANIFEST), "frozenDistManifestSha256": sha(MANIFEST.read_bytes()),
    "frozenDistMatchesBeforeAndAfter": True, "siteFiles": len(expected),
    "matchingAssets": len([n for n in expected if n.startswith("assets/")]),
    "allowedWpsTransforms": sorted(transforms), "buildInfoHashesMatch": True,
    "excludedDevelopmentAndPrivateMaterial": True, "old021Preserved": True,
    "manifest": {"version": "1.0.2.2", "id": OFFICE_ID, "httpsPort": 18766},
    "word": word, "excel": excel, "buildInfo": info,
    "extractedRoot": str(extraction / NAME),
    "scope": "Candidate trial; no host registration, certificate changes or real Office/WPS launches.",
}
with REPORT.open("x", encoding="utf-8") as stream:
    json.dump(report, stream, indent=2, ensure_ascii=False)
print(json.dumps({k: v for k, v in report.items() if k != "buildInfo"}, indent=2, ensure_ascii=False))
