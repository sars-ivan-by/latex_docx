"""Minimal Word footnotes part. python-docx has no footnote API."""
from __future__ import annotations

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from lxml import etree

_EMPTY = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
</w:footnotes>
"""


class Footnotes:
    def __init__(self, document):
        self._document = document
        self._part = None
        self._root = None

    def add(self, paragraph, text: str) -> None:
        root = self._ensure()
        note_id = self._next_id(root)
        root.append(_note(note_id, text))
        self._part._blob = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        paragraph._p.append(_reference(note_id))

    def _ensure(self):
        if self._root is not None:
            return self._root
        part = self._existing_part()
        if part is None:
            part = Part(
                PackURI("/word/footnotes.xml"),
                CT.WML_FOOTNOTES,
                _EMPTY.encode("utf-8"),
                self._document.part.package,
            )
            self._document.part.relate_to(part, RT.FOOTNOTES)
        self._part = part
        self._root = parse_xml(part.blob)
        return self._root

    def _existing_part(self):
        try:
            return self._document.part.part_related_by(RT.FOOTNOTES)
        except KeyError:
            return None

    @staticmethod
    def _next_id(root) -> int:
        used = []
        for note in root.findall(qn("w:footnote")):
            raw = note.get(qn("w:id"))
            if raw and raw.lstrip("-").isdigit():
                used.append(int(raw))
        positive = [n for n in used if n > 0]
        return (max(positive) if positive else 0) + 1


def _reference(note_id: int):
    run = OxmlElement("w:r")
    props = OxmlElement("w:rPr")
    align = OxmlElement("w:vertAlign")
    align.set(qn("w:val"), "superscript")
    props.append(align)
    run.append(props)
    ref = OxmlElement("w:footnoteReference")
    ref.set(qn("w:id"), str(note_id))
    run.append(ref)
    return run


def _note(note_id: int, text: str):
    note = OxmlElement("w:footnote")
    note.set(qn("w:id"), str(note_id))
    paragraph = OxmlElement("w:p")
    mark = OxmlElement("w:r")
    props = OxmlElement("w:rPr")
    align = OxmlElement("w:vertAlign")
    align.set(qn("w:val"), "superscript")
    props.append(align)
    mark.append(props)
    mark.append(OxmlElement("w:footnoteRef"))
    body = OxmlElement("w:r")
    node = OxmlElement("w:t")
    node.set(qn("xml:space"), "preserve")
    node.text = " " + text
    body.append(node)
    paragraph.append(mark)
    paragraph.append(body)
    note.append(paragraph)
    return note
