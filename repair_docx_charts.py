# -*- coding: utf-8 -*-
"""Repair chart XML namespace prefixes in generated DOCX files."""

from __future__ import annotations

import argparse
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET


NAMESPACES = {
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "c14": "http://schemas.microsoft.com/office/drawing/2007/8/2/chart",
    "c15": "http://schemas.microsoft.com/office/drawing/2012/chart",
    "c16": "http://schemas.microsoft.com/office/drawing/2014/chart",
    "wps": "https://web.wps.cn/et/2018/main",
}


def repair_docx(input_docx: Path, output_docx: Path) -> Path:
    for prefix, uri in NAMESPACES.items():
        ET.register_namespace(prefix, uri)

    with ZipFile(input_docx, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}

    for name in list(members):
        if name.startswith("word/charts/chart") and name.endswith(".xml"):
            root = ET.fromstring(members[name])
            members[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(output_docx, "w", ZIP_DEFLATED) as zout:
        for name, blob in members.items():
            zout.writestr(name, blob)
    return output_docx


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input_docx", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    output = args.output or args.input_docx.with_name(args.input_docx.stem + "_修复图表包.docx")
    print(repair_docx(args.input_docx, output))


if __name__ == "__main__":
    main()
