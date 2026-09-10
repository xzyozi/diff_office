from __future__ import annotations

from dataclasses import dataclass
import os
import posixpath
from typing import Mapping
from urllib.parse import urlparse
import xml.etree.ElementTree as ET
import zipfile

WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NAMESPACES = {"r": OFFICE_RELATIONSHIP_NAMESPACE, "w": WORD_NAMESPACE}


@dataclass(frozen=True)
class LinkIntegrityIssue:
    code: str
    severity: str
    source_part: str
    reference_kind: str
    detail: str
    reference_id: str | None = None
    target: str | None = None


@dataclass(frozen=True)
class LinkIntegrityResult:
    path: str
    status: str
    checked_references: int
    issues: tuple[LinkIntegrityIssue, ...]


def validate_word_links(filepath: str | os.PathLike[str]) -> LinkIntegrityResult:
    """Validate Word package links without opening the document or using a network."""
    return WordLinkIntegrityChecker(filepath).check()


class WordLinkIntegrityChecker:
    """Read-only checker for hyperlinks, bookmarks, and package relationships."""

    def __init__(self, filepath: str | os.PathLike[str]):
        self.filepath = os.fspath(filepath)

    def check(self) -> LinkIntegrityResult:
        issues: list[LinkIntegrityIssue] = []
        try:
            with zipfile.ZipFile(self.filepath, "r") as package:
                package_files = set(package.namelist())
                if "word/document.xml" not in package_files:
                    return self._unreadable("MISSING_DOCUMENT", "word/document.xml が見つかりません。")

                parts = self._word_xml_parts(package_files)
                relationships = self._load_relationships(package, parts, package_files, issues)
                document = self._parse_xml(package, "word/document.xml", issues)
                if document is None:
                    return LinkIntegrityResult(self.filepath, "UNREADABLE", 0, tuple(issues))

                checked = self._check_hyperlinks(document, relationships.get("word/document.xml", {}), issues)
                checked += self._check_package_relationships(parts, package, relationships, package_files, issues)
                status = "OK" if not issues else "ISSUES_FOUND"
                return LinkIntegrityResult(self.filepath, status, checked, tuple(issues))
        except (OSError, zipfile.BadZipFile) as error:
            return self._unreadable("UNREADABLE_PACKAGE", str(error))

    def _unreadable(self, code: str, detail: str) -> LinkIntegrityResult:
        issue = LinkIntegrityIssue(code, "ERROR", "word/document.xml", "package", detail)
        return LinkIntegrityResult(self.filepath, "UNREADABLE", 0, (issue,))

    @staticmethod
    def _word_xml_parts(package_files: set[str]) -> list[str]:
        return sorted(
            name
            for name in package_files
            if name.startswith("word/") and name.endswith(".xml") and "/_rels/" not in name
        )

    def _parse_xml(self, package: zipfile.ZipFile, part: str, issues: list[LinkIntegrityIssue]) -> ET.Element | None:
        try:
            with package.open(part) as source:
                return ET.parse(source).getroot()
        except (KeyError, ET.ParseError, OSError) as error:
            issues.append(LinkIntegrityIssue("INVALID_XML", "ERROR", part, "xml", str(error)))
            return None

    def _load_relationships(
        self,
        package: zipfile.ZipFile,
        parts: list[str],
        package_files: set[str],
        issues: list[LinkIntegrityIssue],
    ) -> dict[str, dict[str, Mapping[str, str]]]:
        relationships: dict[str, dict[str, Mapping[str, str]]] = {}
        for part in parts:
            relationship_part = self._relationship_part_name(part)
            if relationship_part not in package_files:
                relationships[part] = {}
                continue
            root = self._parse_xml(package, relationship_part, issues)
            if root is None:
                relationships[part] = {}
                continue
            relationships[part] = {
                item.get("Id", ""): item.attrib
                for item in root.findall(f"{{{RELATIONSHIP_NAMESPACE}}}Relationship")
                if item.get("Id")
            }
        return relationships

    @staticmethod
    def _relationship_part_name(part: str) -> str:
        directory = posixpath.dirname(part)
        filename = posixpath.basename(part)
        return f"{directory}/_rels/{filename}.rels"

    def _check_hyperlinks(
        self,
        document: ET.Element,
        relationships: Mapping[str, Mapping[str, str]],
        issues: list[LinkIntegrityIssue],
    ) -> int:
        bookmarks = self._bookmarks(document)
        checked = 0
        for hyperlink in document.findall(".//w:hyperlink", NAMESPACES):
            relation_id = hyperlink.get(f"{{{OFFICE_RELATIONSHIP_NAMESPACE}}}id")
            anchor = hyperlink.get(f"{{{WORD_NAMESPACE}}}anchor")
            if relation_id:
                checked += 1
                self._check_external_hyperlink(relation_id, relationships, issues)
            if anchor:
                checked += 1
                if anchor not in bookmarks:
                    issues.append(
                        LinkIntegrityIssue(
                            "MISSING_BOOKMARK",
                            "ERROR",
                            "word/document.xml",
                            "bookmark",
                            "リンク先のブックマークがありません。",
                            target=anchor,
                        )
                    )
                elif bookmarks[anchor] > 1:
                    issues.append(
                        LinkIntegrityIssue(
                            "DUPLICATE_BOOKMARK",
                            "WARNING",
                            "word/document.xml",
                            "bookmark",
                            "同名のブックマークが複数あります。",
                            target=anchor,
                        )
                    )
        return checked

    @staticmethod
    def _bookmarks(document: ET.Element) -> dict[str, int]:
        bookmarks: dict[str, int] = {}
        for bookmark in document.findall(".//w:bookmarkStart", NAMESPACES):
            name = bookmark.get(f"{{{WORD_NAMESPACE}}}name")
            if name:
                bookmarks[name] = bookmarks.get(name, 0) + 1
        return bookmarks

    def _check_external_hyperlink(
        self,
        relation_id: str,
        relationships: Mapping[str, Mapping[str, str]],
        issues: list[LinkIntegrityIssue],
    ) -> None:
        relationship = relationships.get(relation_id)
        if relationship is None:
            issues.append(
                LinkIntegrityIssue(
                    "MISSING_RELATIONSHIP",
                    "ERROR",
                    "word/document.xml",
                    "hyperlink",
                    "r:id に対応するリレーションがありません。",
                    relation_id,
                )
            )
            return
        target = relationship.get("Target", "")
        relationship_type = relationship.get("Type", "")
        if not relationship_type.endswith("/hyperlink") or relationship.get("TargetMode") != "External":
            issues.append(
                LinkIntegrityIssue(
                    "INVALID_HYPERLINK_RELATIONSHIP",
                    "ERROR",
                    "word/document.xml",
                    "hyperlink",
                    "外部ハイパーリンクのリレーション形式が不正です。",
                    relation_id,
                    target,
                )
            )
        elif not self._is_external_target(target):
            issues.append(
                LinkIntegrityIssue(
                    "INVALID_EXTERNAL_TARGET",
                    "ERROR",
                    "word/document.xml",
                    "hyperlink",
                    "外部リンク先の形式が不正です。",
                    relation_id,
                    target,
                )
            )

    @staticmethod
    def _is_external_target(target: str) -> bool:
        parsed = urlparse(target)
        return bool(target and (parsed.scheme in {"http", "https", "mailto"}) and (parsed.netloc or parsed.path))

    def _check_package_relationships(
        self,
        parts: list[str],
        package: zipfile.ZipFile,
        relationships: Mapping[str, Mapping[str, Mapping[str, str]]],
        package_files: set[str],
        issues: list[LinkIntegrityIssue],
    ) -> int:
        checked = 0
        for part in parts:
            root = self._parse_xml(package, part, issues)
            if root is None:
                continue
            for element in root.iter():
                for attribute, relation_id in element.attrib.items():
                    if not attribute.startswith(f"{{{OFFICE_RELATIONSHIP_NAMESPACE}}}"):
                        continue
                    if (
                        part == "word/document.xml"
                        and element.tag == f"{{{WORD_NAMESPACE}}}hyperlink"
                        and attribute.endswith("}id")
                    ):
                        continue
                    checked += 1
                    self._check_package_relationship(
                        part, relation_id, relationships.get(part, {}), package_files, issues
                    )
        return checked

    def _check_package_relationship(
        self,
        part: str,
        relation_id: str,
        relationships: Mapping[str, Mapping[str, str]],
        package_files: set[str],
        issues: list[LinkIntegrityIssue],
    ) -> None:
        relationship = relationships.get(relation_id)
        if relationship is None:
            issues.append(
                LinkIntegrityIssue(
                    "MISSING_RELATIONSHIP",
                    "ERROR",
                    part,
                    "relationship",
                    "参照先のリレーションがありません。",
                    relation_id,
                )
            )
            return
        target = relationship.get("Target", "")
        if relationship.get("TargetMode") == "External":
            return
        resolved_target = self._resolve_internal_target(part, target)
        if resolved_target is None:
            issues.append(
                LinkIntegrityIssue(
                    "INVALID_INTERNAL_TARGET",
                    "ERROR",
                    part,
                    "relationship",
                    "パッケージ外を参照するTargetです。",
                    relation_id,
                    target,
                )
            )
        elif resolved_target not in package_files:
            issues.append(
                LinkIntegrityIssue(
                    "MISSING_PACKAGE_TARGET",
                    "ERROR",
                    part,
                    "relationship",
                    "ZIP内に参照先がありません。",
                    relation_id,
                    resolved_target,
                )
            )

    @staticmethod
    def _resolve_internal_target(part: str, target: str) -> str | None:
        if not target:
            return None
        target_path = target.lstrip("/") if target.startswith("/") else posixpath.join(posixpath.dirname(part), target)
        resolved_target = posixpath.normpath(target_path)
        if resolved_target == ".." or resolved_target.startswith("../"):
            return None
        return resolved_target
