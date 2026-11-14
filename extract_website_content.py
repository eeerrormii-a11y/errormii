#!/usr/bin/env python3
"""Extract website copy into a DOCX; requires python-docx.

Install the dependency with: python -m pip install python-docx
"""

from __future__ import annotations

import argparse
import ast
import html
import re
from collections import defaultdict
from dataclasses import dataclass, field
from html.parser import HTMLParser
from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt


SOURCE_EXTENSIONS = {".html", ".htm", ".jsx", ".tsx", ".vue", ".svelte", ".js", ".ts"}
SKIP_DIRS = {".git", "node_modules", "dist", "build", "coverage", "vendor"}
TEXT_TAGS = {
    "title", "h1", "h2", "h3", "h4", "h5", "h6", "p", "a", "button",
    "label", "li", "summary", "option", "figcaption", "blockquote", "dt", "dd",
}
ATTRIBUTE_COPY = {"aria-label", "alt", "placeholder", "title"}
VOID_TAGS = {
    "area", "base", "br", "col", "embed", "hr", "img", "input", "link",
    "meta", "param", "source", "track", "wbr",
}


@dataclass
class Element:
    tag: str
    attrs: dict[str, str]
    parent: Element | None = None
    children: list[Element | str] = field(default_factory=list)


class PageParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.root = Element("document", {})
        self.stack = [self.root]

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        element = Element(tag, {key.lower(): value or "" for key, value in attrs}, self.stack[-1])
        self.stack[-1].children.append(element)
        if tag not in VOID_TAGS:
            self.stack.append(element)

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        element = Element(tag, {key.lower(): value or "" for key, value in attrs}, self.stack[-1])
        self.stack[-1].children.append(element)

    def handle_endtag(self, tag: str) -> None:
        for index in range(len(self.stack) - 1, 0, -1):
            if self.stack[index].tag == tag:
                del self.stack[index:]
                return

    def handle_data(self, data: str) -> None:
        if self.stack[-1].tag not in {"script", "style", "template"}:
            self.stack[-1].children.append(data)


def element_text(element: Element) -> str:
    parts: list[str] = []

    def visit(node: Element | str) -> None:
        if isinstance(node, str):
            parts.append(node)
        elif node.tag not in {"script", "style", "template"}:
            for child in node.children:
                visit(child)

    visit(element)
    return re.sub(r"\s+", " ", " ".join(parts)).strip()


def descendants(element: Element):
    for child in element.children:
        if isinstance(child, Element):
            yield child
            yield from descendants(child)


def nearest_group(element: Element, fallback: str) -> str:
    current: Element | None = element
    while current is not None:
        if current.tag in {"article", "section"} or current.attrs.get("id"):
            heading = next(
                (
                    element_text(node)
                    for node in descendants(current)
                    if node.tag in {"h1", "h2", "h3", "h4"}
                    and element_text(node)
                ),
                "",
            )
            name = heading or current.attrs.get("aria-label") or current.attrs.get("id") or current.tag
            return f"{fallback} — {name}"
        current = current.parent
    return fallback


def parse_html(path: Path) -> dict[str, list[tuple[str, str]]]:
    parser = PageParser()
    source = path.read_text(encoding="utf-8", errors="replace")
    parser.feed(source)
    result: dict[str, list[tuple[str, str]]] = defaultdict(list)
    page_name = f"Page: {path.name}"

    for element in descendants(parser.root):
        if element.tag in TEXT_TAGS:
            text = element_text(element)
            if text:
                group = nearest_group(element, page_name)
                result[group].append(("Text", text))
        for attr in ATTRIBUTE_COPY:
            value = element.attrs.get(attr, "").strip()
            if value:
                group = nearest_group(element, page_name)
                result[group].append((attr.replace("-", " ").title(), value))
        if element.tag == "meta" and element.attrs.get("name", "").lower() == "description":
            value = element.attrs.get("content", "").strip()
            if value:
                result[f"Page metadata: {path.name}"].append(("Description", value))

    return result


def read_js_string(line: str, index: int) -> tuple[str, int]:
    quote = line[index]
    index += 1
    value: list[str] = []
    while index < len(line):
        char = line[index]
        if char == "\\" and index + 1 < len(line):
            value.append(line[index:index + 2])
            index += 2
        elif char == quote:
            literal = quote + "".join(value) + quote
            try:
                decoded = ast.literal_eval(literal)
            except (SyntaxError, ValueError):
                decoded = "".join(value).replace("\\'", "'").replace('\\"', '"')
            return str(decoded), index + 1
        else:
            value.append(char)
            index += 1
    return "".join(value), index


def extract_translations(source: str) -> dict[str, list[tuple[str, str]]]:
    locales: dict[str, list[tuple[str, str]]] = defaultdict(list)
    locale: str | None = None
    inside_dictionary = False

    for line in source.splitlines():
        if not inside_dictionary and re.search(r"\btranslations\s*=\s*\{", line):
            inside_dictionary = True
            continue
        if not inside_dictionary:
            continue
        locale_match = re.match(r"\s{2}(en|hu|sk)\s*:\s*\{", line)
        if locale_match:
            locale = locale_match.group(1)
            continue
        if locale and re.match(r"\s{2}\}\s*,?\s*$", line):
            locale = None
            continue
        if locale:
            match = re.match(r"\s*(['\"])", line)
            if not match:
                continue
            key, end = read_js_string(line, match.start(1))
            colon = re.match(r"\s*:\s*(['\"])", line[end:])
            if not colon:
                continue
            value_start = end + colon.end() - 1
            value, _ = read_js_string(line, value_start)
            locales[locale].append((key, value))
        elif re.match(r"\s*\};?\s*$", line):
            break
    return locales


def translation_group(key: str) -> str:
    if key.startswith("notYourTree"):
        return "Project: Not Your Tree"
    if key.startswith("tourism"):
        return "Project: Tourism Art"
    if key.startswith("rice"):
        return "Performance: Rice"
    if key.startswith("giveTake"):
        return "Performance: Give-Take"
    if key.startswith("trace"):
        return "Performance: Trace"
    if key.startswith("days"):
        return "Performance: Days of War"
    if key.startswith("kala"):
        return "Project: Kala Sutra"
    if key.startswith("principle"):
        return "About: Principles"
    if key.startswith(("bio", "artist")):
        return "About: Artist biographies"
    if key.startswith("about"):
        return "About"
    if key.startswith("contact"):
        return "Contact"
    if key.startswith("footer"):
        return "Footer"
    if key.startswith("map"):
        return "Not Your Tree: Map and dialogs"
    if key in {"formLabel", "openGallery", "clickLarger", "mapExpand"}:
        return "Controls and forms"

    prefix = key.split(".", 1)[0]
    groups = {
        "nav": "Navigation",
        "hero": "Hero",
        "works": "Works",
        "filter": "Work filters",
        "group": "Work categories",
        "workshop": "Workshop",
    }
    return groups.get(prefix, "Other interface and content")


def extract_dynamic_js(source: str) -> list[tuple[str, str]]:
    found: list[tuple[str, str]] = []
    patterns = (
        (r"\b(?:textContent|innerHTML)\s*=\s*(['\"])((?:\\.|(?!\1).)*?)\1", "Dynamic UI text"),
        (
            r"\b(?:textContent|innerHTML)\s*=[^;\n]*?\|\|\s*(['\"])((?:\\.|(?!\1).)*?)\1",
            "Dynamic UI fallback",
        ),
        (r"\b(?:name|caption)\s*:\s*(['\"])((?:\\.|(?!\1).)*?)\1", "Map location"),
        (r"\b(['\"])(Map|Satellite)\1\s*:", "Map control"),
        (r"\battribution\s*:\s*(['\"])((?:\\.|(?!\1).)*?)\1", "Map attribution"),
    )
    for pattern, kind in patterns:
        for match in re.finditer(pattern, source):
            value = match.group(2).replace("\\'", "'").replace('\\"', '"').strip()
            if value and not value.startswith("<"):
                found.append((kind, value))
    return found


def collect_content(root: Path) -> tuple[
    dict[str, dict[str, list[tuple[str, str]]]],
    dict[str, dict[str, list[tuple[str, str]]]],
]:
    pages: dict[str, dict[str, list[tuple[str, str]]]] = {}
    translations: dict[str, dict[str, list[tuple[str, str]]]] = {}
    js_sources: list[Path] = []

    for path in sorted(root.rglob("*")):
        if not path.is_file() or path.suffix.lower() not in SOURCE_EXTENSIONS:
            continue
        if any(part in SKIP_DIRS for part in path.relative_to(root).parts):
            continue
        if path.suffix.lower() in {".html", ".htm", ".jsx", ".tsx", ".vue", ".svelte"}:
            pages[str(path.relative_to(root))] = parse_html(path)
            source = path.read_text(encoding="utf-8", errors="replace")
            for src in re.findall(r"<script\b[^>]*\bsrc=['\"]([^'\"]+)['\"]", source, re.I):
                linked = (path.parent / html.unescape(src.split("?", 1)[0])).resolve()
                if linked.is_file() and root.resolve() in linked.parents:
                    js_sources.append(linked)
            js_sources.append(path)
        elif path.suffix.lower() in {".js", ".ts"}:
            js_sources.append(path)

    for path in sorted(set(js_sources)):
        source = path.read_text(encoding="utf-8", errors="replace")
        locale_data = extract_translations(source)
        for locale, entries in locale_data.items():
            groups = translations.setdefault(locale, defaultdict(list))
            for key, value in entries:
                groups[translation_group(key)].append((key, value))
        dynamic = extract_dynamic_js(source)
        if dynamic:
            file_key = f"Dynamic copy: {path.relative_to(root)}"
            pages.setdefault(file_key, {}).setdefault(file_key, []).extend(dynamic)

    return pages, translations


def make_document(
    pages: dict[str, dict[str, list[tuple[str, str]]]],
    translations: dict[str, dict[str, list[tuple[str, str]]]],
    output: Path,
) -> None:
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)
    document.styles["Normal"].font.name = "Arial"
    document.styles["Normal"].font.size = Pt(10)

    document.add_heading("Website Content Inventory", 0)
    document.add_paragraph("Extracted user-facing copy, labels, accessible names, and translations.")

    if pages:
        document.add_heading("Pages and Components", level=1)
        for file_name, groups in pages.items():
            document.add_heading(file_name, level=2)
            for group_name, entries in groups.items():
                document.add_heading(group_name, level=3)
                for kind, text in entries:
                    paragraph = document.add_paragraph(style="List Bullet")
                    paragraph.add_run(f"{kind}: ").bold = True
                    paragraph.add_run(text)

    locale_names = {"en": "English", "hu": "Hungarian", "sk": "Slovak"}
    for locale, groups in translations.items():
        document.add_page_break()
        document.add_heading(f"Translations — {locale_names.get(locale, locale)}", level=1)
        for group_name, entries in groups.items():
            document.add_heading(group_name, level=2)
            for key, text in entries:
                paragraph = document.add_paragraph(style="List Bullet")
                paragraph.add_run(f"{key}: ").bold = True
                paragraph.add_run(text)

    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(output)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parent)
    parser.add_argument(
        "--output",
        type=Path,
        default=Path(__file__).resolve().parent / "website_content.docx",
    )
    args = parser.parse_args()
    pages, translations = collect_content(args.root.resolve())
    if not pages and not translations:
        raise SystemExit(f"No UI content found under {args.root}")
    make_document(pages, translations, args.output.resolve())
    print(f"Saved website content to {args.output.resolve()}")


if __name__ == "__main__":
    main()
