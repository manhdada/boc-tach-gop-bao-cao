from __future__ import annotations

import copy
import re
import unicodedata
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from pypdf import PdfReader


ROMAN_RE = re.compile(r"^([IVXLCDM]+)[.\)]\s+(.+)$")
COMPOUND_RE = re.compile(r"^(\d+(?:\.\d+)+)[.\)]?\s+(.+)$")
DECIMAL_RE = re.compile(r"^(\d+)[.\)]\s+(.+)$")
LETTER_RE = re.compile(r"^([a-zđ])[.\)]\s+(.+)$", re.I)
MANUAL_PREFIX_RE = re.compile(
    r"^(?:(?:[IVXLCDM]+|\d+(?:\.\d+)*|[a-zđ])[.\)])\s+", re.I
)


@dataclass
class RunData:
    text: str
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None


@dataclass
class Block:
    kind: str
    text: str
    runs: list[RunData] = field(default_factory=list)
    alignment: int | None = None
    left_indent_pt: float | None = None
    first_line_indent_pt: float | None = None
    table_element: Any | None = None


@dataclass
class Occurrence:
    id: str
    source_id: str
    source_name: str
    source_order: int
    title: str
    clean_title: str
    norm_title: str
    level: int
    marker: str
    marker_kind: str
    sibling_index: int
    blocks: list[Block] = field(default_factory=list)

    @property
    def content_text(self) -> str:
        return "\n".join(b.text for b in self.blocks if b.text.strip()).strip()


@dataclass
class SourceReport:
    id: str
    path: Path
    name: str
    order: int
    occurrences: list[Occurrence] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


@dataclass
class SectionGroup:
    id: str
    title: str
    norm_title: str
    level: int
    parent_id: str | None
    order: int
    marker_kind: str
    occurrences: list[Occurrence] = field(default_factory=list)
    enabled: bool = True


def new_id() -> str:
    return uuid.uuid4().hex


def strip_marker(text: str) -> tuple[str, str, str, int | None]:
    text = " ".join(text.strip().split())
    for regex, kind, level in (
        (ROMAN_RE, "roman", 1),
        (COMPOUND_RE, "compound", 3),
        (DECIMAL_RE, "decimal", 2),
        (LETTER_RE, "letter", 3),
    ):
        match = regex.match(text)
        if match:
            return match.group(2).strip(), match.group(1), kind, level
    return text, "", "none", None


def normalize_title(text: str) -> str:
    clean, _, _, _ = strip_marker(text)
    clean = clean.rstrip(".:;- ")
    clean = clean.replace("Đ", "D").replace("đ", "d")
    clean = unicodedata.normalize("NFKD", clean)
    clean = "".join(ch for ch in clean if not unicodedata.combining(ch)).lower()
    replacements = {
        r"\btthc\b": "thu tuc hanh chinh",
        r"\bubnd\b": "uy ban nhan dan",
        r"\bhdnd\b": "hoi dong nhan dan",
        r"\btpvhcc\b": "trung tam phuc vu hanh chinh cong",
        r"\b02\b": "2",
    }
    for pattern, value in replacements.items():
        clean = re.sub(pattern, value, clean)
    clean = re.sub(r"[^a-z0-9]+", " ", clean)
    clean = re.sub(r"\b(doi voi|cac|cua|tai)\b", " ", clean)
    return " ".join(clean.split())


def _bold_ratio(paragraph: Paragraph) -> float:
    total = sum(len(run.text.strip()) for run in paragraph.runs)
    bold = sum(len(run.text.strip()) for run in paragraph.runs if run.bold)
    return bold / total if total else 0.0


def _direct_num_pr(paragraph: Paragraph):
    ppr = paragraph._p.pPr
    return ppr.numPr if ppr is not None else None


class NumberingResolver:
    def __init__(self, document):
        self.num_to_abs: dict[str, str] = {}
        self.levels: dict[tuple[str, int], tuple[str, str, int]] = {}
        self.counters: dict[str, list[int]] = {}
        root = document.part.numbering_part.element
        for num in root.findall(qn("w:num")):
            num_id = num.get(qn("w:numId"))
            abs_el = num.find(qn("w:abstractNumId"))
            if num_id and abs_el is not None:
                self.num_to_abs[num_id] = abs_el.get(qn("w:val"))
        for abstract in root.findall(qn("w:abstractNum")):
            abs_id = abstract.get(qn("w:abstractNumId"))
            for lvl in abstract.findall(qn("w:lvl")):
                ilvl = int(lvl.get(qn("w:ilvl"), "0"))
                fmt_el = lvl.find(qn("w:numFmt"))
                text_el = lvl.find(qn("w:lvlText"))
                start_el = lvl.find(qn("w:start"))
                fmt = fmt_el.get(qn("w:val"), "decimal") if fmt_el is not None else "decimal"
                template = text_el.get(qn("w:val"), f"%{ilvl + 1}.") if text_el is not None else f"%{ilvl + 1}."
                start = int(start_el.get(qn("w:val"), "1")) if start_el is not None else 1
                self.levels[(abs_id, ilvl)] = (fmt, template, start)

    @staticmethod
    def _roman(value: int) -> str:
        pairs = ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"),
                 (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"),
                 (5, "V"), (4, "IV"), (1, "I"))
        out = []
        for number, symbol in pairs:
            while value >= number:
                out.append(symbol)
                value -= number
        return "".join(out)

    @classmethod
    def _format(cls, value: int, fmt: str) -> str:
        if fmt == "lowerLetter":
            return chr(ord("a") + max(0, value - 1) % 26)
        if fmt == "upperLetter":
            return chr(ord("A") + max(0, value - 1) % 26)
        if fmt == "lowerRoman":
            return cls._roman(value).lower()
        if fmt == "upperRoman":
            return cls._roman(value)
        if fmt == "bullet":
            return "•"
        return str(value)

    def next_label(self, paragraph: Paragraph) -> str:
        numpr = _direct_num_pr(paragraph)
        if numpr is None or numpr.numId is None:
            return ""
        num_id = str(numpr.numId.val)
        ilvl = int(numpr.ilvl.val) if numpr.ilvl is not None else 0
        abstract = self.num_to_abs.get(num_id)
        if abstract is None:
            return ""
        counters = self.counters.setdefault(num_id, [0] * 9)
        _, _, start = self.levels.get((abstract, ilvl), ("decimal", f"%{ilvl + 1}.", 1))
        counters[ilvl] = counters[ilvl] + 1 if counters[ilvl] else start
        for deeper in range(ilvl + 1, len(counters)):
            counters[deeper] = 0
        fmt, template, _ = self.levels.get((abstract, ilvl), ("decimal", f"%{ilvl + 1}.", 1))
        label = template
        for idx in range(1, 10):
            value = counters[idx - 1] or 1
            level_fmt = self.levels.get((abstract, idx - 1), (fmt, "", 1))[0]
            label = label.replace(f"%{idx}", self._format(value, level_fmt))
        return label.strip()


def _heading_from_paragraph(paragraph: Paragraph, display_text: str, number_label: str):
    raw = " ".join(display_text.split())
    if not raw or len(raw) > 280:
        return None
    clean, marker, marker_kind, regex_level = strip_marker(raw)
    style_name = paragraph.style.name.lower() if paragraph.style else ""
    outline = None
    if paragraph._p.pPr is not None:
        outline_el = paragraph._p.pPr.find(qn("w:outlineLvl"))
        if outline_el is not None:
            outline = int(outline_el.get(qn("w:val"), "0")) + 1
    style_level = None
    style_match = re.search(r"heading\s*(\d+)", style_name)
    if style_match:
        style_level = int(style_match.group(1))
    has_emphasis = _bold_ratio(paragraph) >= 0.18 or style_level is not None or outline is not None
    if regex_level is not None and has_emphasis:
        return clean, marker, marker_kind, regex_level
    if number_label and _bold_ratio(paragraph) >= 0.55:
        clean, marker, marker_kind, regex_level = strip_marker(f"{number_label} {paragraph.text}")
        level = regex_level or min(3, int(_direct_num_pr(paragraph).ilvl.val if _direct_num_pr(paragraph).ilvl is not None else 0) + 2)
        return clean, marker, marker_kind, level
    level = style_level or outline
    if level is not None and level <= 4:
        return clean, marker, marker_kind, level
    return None


def _paragraph_block(paragraph: Paragraph, display_text: str) -> Block:
    runs = [RunData(run.text, run.bold, run.italic, run.underline) for run in paragraph.runs]
    run_text = "".join(run.text for run in runs)
    if run_text != paragraph.text:
        runs = [RunData(paragraph.text)]
    prefix = ""
    if display_text.endswith(paragraph.text) and display_text != paragraph.text:
        prefix = display_text[: -len(paragraph.text)]
    if prefix:
        runs.insert(0, RunData(prefix))
    fmt = paragraph.paragraph_format
    return Block(
        kind="paragraph",
        text=display_text,
        runs=runs,
        alignment=int(paragraph.alignment) if paragraph.alignment is not None else None,
        left_indent_pt=fmt.left_indent.pt if fmt.left_indent is not None else None,
        first_line_indent_pt=fmt.first_line_indent.pt if fmt.first_line_indent is not None else None,
    )


def _table_block(table: Table) -> Block:
    rows = []
    for row in table.rows:
        rows.append("\t".join(cell.text.strip() for cell in row.cells))
    return Block(
        kind="table",
        text="\n".join(rows),
        table_element=copy.deepcopy(table._tbl),
    )


def _is_signature_table(table: Table) -> bool:
    text = " ".join(cell.text for row in table.rows for cell in row.cells).lower()
    normalized = normalize_title(text)
    return (
        "noi nhan" in normalized
        or "kt chu tich" in normalized
        or "tm uy ban nhan dan" in normalized
    )


def parse_docx(path: Path, source_id: str, source_name: str, source_order: int) -> SourceReport:
    document = Document(path)
    resolver = NumberingResolver(document)
    report = SourceReport(source_id, path, source_name, source_order)
    current: Occurrence | None = None
    started = False
    sibling_counters = [0] * 8
    for block in document.iter_inner_content():
        if isinstance(block, Paragraph):
            label = resolver.next_label(block)
            display = f"{label} {block.text}".strip() if label else block.text.strip()
            heading = _heading_from_paragraph(block, display, label)
            if heading:
                clean, marker, marker_kind, level = heading
                if level == 1:
                    started = True
                if not started:
                    continue
                sibling_counters[level] += 1
                for idx in range(level + 1, len(sibling_counters)):
                    sibling_counters[idx] = 0
                current = Occurrence(
                    id=new_id(),
                    source_id=source_id,
                    source_name=source_name,
                    source_order=source_order,
                    title=display,
                    clean_title=clean,
                    norm_title=normalize_title(clean),
                    level=level,
                    marker=marker,
                    marker_kind=marker_kind,
                    sibling_index=sibling_counters[level],
                )
                report.occurrences.append(current)
            elif started and current is not None and display:
                current.blocks.append(_paragraph_block(block, display))
        elif isinstance(block, Table) and started and current is not None:
            if _is_signature_table(block):
                break
            current.blocks.append(_table_block(block))
    if not report.occurrences:
        report.warnings.append("Không phát hiện được tiêu đề phần/mục.")
    return report


def _pdf_heading(line: str):
    clean, marker, kind, level = strip_marker(line)
    if level is not None and len(line) <= 280:
        return clean, marker, kind, level
    return None


def parse_pdf(path: Path, source_id: str, source_name: str, source_order: int) -> SourceReport:
    reader = PdfReader(str(path))
    report = SourceReport(source_id, path, source_name, source_order)
    current: Occurrence | None = None
    started = False
    sibling_counters = [0] * 8
    for page in reader.pages:
        text = page.extract_text() or ""
        for raw_line in text.splitlines():
            line = " ".join(raw_line.split())
            if not line:
                continue
            heading = _pdf_heading(line)
            if heading:
                clean, marker, marker_kind, level = heading
                if level == 1:
                    started = True
                if not started:
                    continue
                sibling_counters[level] += 1
                for idx in range(level + 1, len(sibling_counters)):
                    sibling_counters[idx] = 0
                current = Occurrence(
                    new_id(), source_id, source_name, source_order, line, clean,
                    normalize_title(clean), level, marker, marker_kind,
                    sibling_counters[level], []
                )
                report.occurrences.append(current)
            elif started and current is not None:
                current.blocks.append(Block("paragraph", line, [RunData(line)]))
    if not report.occurrences:
        report.warnings.append("PDF không có lớp text hoặc không phát hiện được tiêu đề.")
    return report


def parse_report(path: str | Path, source_order: int, source_name: str | None = None) -> SourceReport:
    path = Path(path)
    source_id = new_id()
    name = source_name or path.stem
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return parse_docx(path, source_id, name, source_order)
    if suffix == ".pdf":
        return parse_pdf(path, source_id, name, source_order)
    raise ValueError(f"Định dạng không được hỗ trợ: {path.suffix}")


class ReportModel:
    def __init__(self):
        self.sources: list[SourceReport] = []
        self.groups: list[SectionGroup] = []
        self.aliases: dict[str, str] = {}

    def analyze(self, files: Iterable[tuple[str, str | None]]) -> None:
        self.sources = [parse_report(path, idx, name) for idx, (path, name) in enumerate(files)]
        self._build_groups()

    def _children(self, parent_id: str | None) -> list[SectionGroup]:
        return sorted((g for g in self.groups if g.parent_id == parent_id), key=lambda g: g.order)

    def _build_groups(self) -> None:
        self.groups = []
        for source in self.sources:
            stack: dict[int, str] = {}
            for occ in source.occurrences:
                parent_id = stack.get(occ.level - 1) if occ.level > 1 else None
                candidates = self._children(parent_id)
                alias_target = self.aliases.get(occ.norm_title)
                match = next((g for g in candidates if g.norm_title == (alias_target or occ.norm_title)), None)
                if match is None and occ.marker and occ.level > 1:
                    match = next(
                        (g for g in candidates if any(o.marker.lower() == occ.marker.lower() for o in g.occurrences)),
                        None,
                    )
                if match is None:
                    match = next((g for g in candidates if g.order == occ.sibling_index - 1), None)
                if match is None:
                    match = SectionGroup(
                        new_id(), occ.clean_title, occ.norm_title, occ.level,
                        parent_id, len(candidates), occ.marker_kind, []
                    )
                    self.groups.append(match)
                match.occurrences.append(occ)
                stack[occ.level] = match.id
                for deeper in list(stack):
                    if deeper > occ.level:
                        del stack[deeper]

    def group(self, group_id: str) -> SectionGroup:
        return next(g for g in self.groups if g.id == group_id)

    def children(self, parent_id: str | None) -> list[SectionGroup]:
        return self._children(parent_id)

    def roots(self) -> list[SectionGroup]:
        return self._children(None)

    def rename_group(self, group_id: str, title: str) -> None:
        group = self.group(group_id)
        group.title = title.strip()
        group.norm_title = normalize_title(title)

    def toggle_group(self, group_id: str) -> None:
        group = self.group(group_id)
        group.enabled = not group.enabled

    def merge_groups(self, group_ids: list[str]) -> str:
        selected = [self.group(group_id) for group_id in group_ids]
        if len(selected) < 2:
            raise ValueError("Chọn ít nhất hai mục để gộp.")
        if len({g.level for g in selected}) != 1:
            raise ValueError("Chỉ có thể gộp các mục cùng cấp.")
        target = selected[0]
        for other in selected[1:]:
            target.occurrences.extend(other.occurrences)
            for child in self.children(other.id):
                child.parent_id = target.id
            for occ in other.occurrences:
                self.aliases[occ.norm_title] = target.norm_title
            self.groups.remove(other)
        self._merge_duplicate_children(target.id)
        self._reindex(target.parent_id)
        return target.id

    def _merge_duplicate_children(self, parent_id: str) -> None:
        seen: dict[str, SectionGroup] = {}
        for child in list(self.children(parent_id)):
            existing = seen.get(child.norm_title)
            if existing is None:
                seen[child.norm_title] = child
                continue
            existing.occurrences.extend(child.occurrences)
            for grandchild in self.children(child.id):
                grandchild.parent_id = existing.id
            self.groups.remove(child)
            self._merge_duplicate_children(existing.id)
        self._reindex(parent_id)

    def split_group(self, group_id: str) -> list[str]:
        group = self.group(group_id)
        if len(group.occurrences) < 2:
            raise ValueError("Mục này chỉ có một nguồn nên không thể tách.")
        occurrences = sorted(group.occurrences, key=lambda o: o.source_order)
        siblings = self.children(group.parent_id)
        insert_at = siblings.index(group)
        group.occurrences = [occurrences[0]]
        created = [group.id]
        source_to_parent = {occurrences[0].source_id: group.id}
        created_groups = [group]
        for offset, occ in enumerate(occurrences[1:], 1):
            new_group = SectionGroup(
                new_id(), occ.clean_title, occ.norm_title, group.level,
                group.parent_id, group.order + offset, occ.marker_kind, [occ], group.enabled
            )
            self.groups.append(new_group)
            created.append(new_group.id)
            created_groups.append(new_group)
            source_to_parent[occ.source_id] = new_group.id
        ordered = siblings[:insert_at] + created_groups + siblings[insert_at + 1:]
        for index, sibling in enumerate(ordered):
            sibling.order = index
        self._redistribute_children(group.id, source_to_parent)
        return created

    def _redistribute_children(self, old_parent_id: str, source_to_parent: dict[str, str]) -> None:
        old_children = list(self.children(old_parent_id))
        for child in old_children:
            grandchildren_parent = child.id
            self.groups.remove(child)
            source_to_child: dict[str, str] = {}
            sibling_counts: dict[str, int] = {}
            for source_id, parent_id in source_to_parent.items():
                occurrences = [occ for occ in child.occurrences if occ.source_id == source_id]
                if not occurrences:
                    continue
                order = sibling_counts.get(parent_id, 0)
                sibling_counts[parent_id] = order + 1
                clone = SectionGroup(
                    new_id(), child.title, child.norm_title, child.level,
                    parent_id, child.order, child.marker_kind, occurrences, child.enabled
                )
                self.groups.append(clone)
                source_to_child[source_id] = clone.id
            self._redistribute_children(grandchildren_parent, source_to_child)
        for parent_id in set(source_to_parent.values()):
            self._reindex(parent_id)

    def move_group(self, group_id: str, delta: int) -> None:
        group = self.group(group_id)
        siblings = self.children(group.parent_id)
        index = siblings.index(group)
        new_index = max(0, min(len(siblings) - 1, index + delta))
        if new_index == index:
            return
        siblings[index], siblings[new_index] = siblings[new_index], siblings[index]
        for idx, sibling in enumerate(siblings):
            sibling.order = idx

    def _reindex(self, parent_id: str | None) -> None:
        for idx, group in enumerate(self.children(parent_id)):
            group.order = idx

    def update_source_name(self, source_id: str, name: str) -> None:
        source = next(s for s in self.sources if s.id == source_id)
        source.name = name
        for group in self.groups:
            for occ in group.occurrences:
                if occ.source_id == source_id:
                    occ.source_name = name

    def alias_payload(self) -> dict[str, str]:
        payload = dict(self.aliases)
        for group in self.groups:
            for occ in group.occurrences:
                payload[occ.norm_title] = group.norm_title
        return payload
