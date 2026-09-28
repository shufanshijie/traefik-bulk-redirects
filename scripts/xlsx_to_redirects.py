#!/usr/bin/env python3
"""Convert an XLSX redirect mapping into the plugin JSON format."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import posixpath
import re
import sys
import tempfile
import zipfile
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator, Sequence
from urllib.parse import SplitResult, urlsplit, urlunsplit
from xml.etree import ElementTree


ALLOWED_STATUS_CODES = {301, 302, 303, 307, 308}
DEFAULT_MAX_JSON_BYTES = 16 << 20
MAX_ERROR_SAMPLES = 20

SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOCUMENT_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

CELL_REFERENCE = re.compile(r"^([A-Za-z]+)([0-9]+)$")

TRUE_VALUES = {"1", "true", "yes", "y", "enabled", "on", "是", "启用"}
FALSE_VALUES = {"0", "false", "no", "n", "disabled", "off", "否", "禁用"}
EXACT_MATCH_VALUES = {"exact", "exacturl", "exact_url", "精确url", "精准url"}

COLUMN_ALIASES = {
    "rule_id": {"rule_id", "ruleid", "规则id"},
    "source": {
        "source_url",
        "sourceurl",
        "匹配串(精确url或正则)",
        "匹配串（精确url或正则）",
        "页面url",
    },
    "target": {"target_url", "targeturl", "301目标(新站)", "301目标（新站）", "新网站映射链接"},
    "match_mode": {"match_mode", "matchmode", "匹配方式"},
    "status_code": {"status_code", "statuscode", "状态码"},
    "preserve_query": {"preserve_query_string", "preservequerystring", "保留query", "保留查询参数"},
    "subpath": {"subpath_matching", "subpathmatching", "子路径匹配"},
    "enabled": {"enabled", "启用"},
}


class ConversionError(Exception):
    """Raised when the workbook cannot be converted safely."""


@dataclass(frozen=True)
class SheetInfo:
    name: str
    path: str


@dataclass(frozen=True)
class ConversionOptions:
    sheet_name: str | None = None
    header_row: int = 1
    default_status_code: int = 301
    default_preserve_query: bool = False
    min_rules: int = 1
    max_rules: int = 30000
    max_json_bytes: int = DEFAULT_MAX_JSON_BYTES
    skip_unsupported: bool = False
    strip_source_fragments: bool = False
    deduplicate_same_target: bool = False
    skip_conflicting_targets: bool = False


@dataclass(frozen=True)
class ConversionResult:
    sheet_name: str
    redirects: list[dict[str, object]]
    disabled_rows: int
    unsupported_rows: int
    stripped_fragments: int
    deduplicated_rows: int
    conflicting_rows: int
    conflict_warnings: tuple[str, ...]


class Problems:
    def __init__(self) -> None:
        self.count = 0
        self.samples: list[str] = []

    def add(self, message: str, count: int = 1) -> None:
        self.count += count
        if len(self.samples) < MAX_ERROR_SAMPLES:
            self.samples.append(message)

    def raise_if_any(self) -> None:
        if self.count == 0:
            return
        omitted = self.count - len(self.samples)
        lines = [f"validation failed with {self.count} error(s)"]
        lines.extend(f"- {message}" for message in self.samples)
        if omitted > 0:
            lines.append(f"- ... {omitted} additional error(s) omitted")
        raise ConversionError("\n".join(lines))


def xml_name(namespace: str, name: str) -> str:
    return f"{{{namespace}}}{name}"


def normalize_header(value: object) -> str:
    return "".join(str(value).strip().lower().split())


def column_index(cell_reference: str) -> int:
    match = CELL_REFERENCE.match(cell_reference)
    if match is None:
        raise ConversionError(f"invalid XLSX cell reference {cell_reference!r}")
    index = 0
    for character in match.group(1).upper():
        index = index * 26 + ord(character) - ord("A") + 1
    return index - 1


def parse_number(value: str) -> int | float | str:
    try:
        if not any(character in value.lower() for character in (".", "e")):
            return int(value)
        return float(value)
    except ValueError:
        return value


class XlsxWorkbook:
    def __init__(self, path: Path) -> None:
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ConversionError("input must be an .xlsx or .xlsm file")
        try:
            self.archive = zipfile.ZipFile(path)
        except (OSError, zipfile.BadZipFile) as exc:
            raise ConversionError(f"unable to open workbook {path}: {exc}") from exc
        self.shared_strings = self._read_shared_strings()
        self.sheets = self._read_sheets()

    def close(self) -> None:
        self.archive.close()

    def __enter__(self) -> XlsxWorkbook:
        return self

    def __exit__(self, exc_type: object, exc: object, traceback: object) -> None:
        self.close()

    def _read_shared_strings(self) -> list[str]:
        try:
            with self.archive.open("xl/sharedStrings.xml") as handle:
                root = ElementTree.parse(handle).getroot()
        except KeyError:
            return []
        return [
            "".join(node.text or "" for node in item.iter(xml_name(SPREADSHEET_NS, "t")))
            for item in root.findall(xml_name(SPREADSHEET_NS, "si"))
        ]

    def _read_sheets(self) -> list[SheetInfo]:
        try:
            with self.archive.open("xl/workbook.xml") as handle:
                workbook = ElementTree.parse(handle).getroot()
            with self.archive.open("xl/_rels/workbook.xml.rels") as handle:
                relationships_root = ElementTree.parse(handle).getroot()
        except KeyError as exc:
            raise ConversionError(f"invalid XLSX workbook: missing {exc.args[0]}") from exc

        relationships = {
            relationship.attrib["Id"]: relationship.attrib["Target"]
            for relationship in relationships_root.findall(xml_name(PACKAGE_REL_NS, "Relationship"))
            if relationship.attrib.get("TargetMode") != "External"
        }

        sheets: list[SheetInfo] = []
        sheets_node = workbook.find(xml_name(SPREADSHEET_NS, "sheets"))
        if sheets_node is None:
            raise ConversionError("invalid XLSX workbook: no worksheets found")
        for sheet in sheets_node.findall(xml_name(SPREADSHEET_NS, "sheet")):
            relationship_id = sheet.attrib.get(xml_name(DOCUMENT_REL_NS, "id"))
            target = relationships.get(relationship_id or "")
            if target is None:
                continue
            if target.startswith("/"):
                sheet_path = target.lstrip("/")
            else:
                sheet_path = posixpath.normpath(posixpath.join("xl", target))
            if sheet_path.startswith("../"):
                raise ConversionError(f"invalid worksheet path {target!r}")
            sheets.append(SheetInfo(name=sheet.attrib["name"], path=sheet_path))
        if not sheets:
            raise ConversionError("invalid XLSX workbook: no readable worksheets found")
        return sheets

    def iter_rows(self, sheet: SheetInfo) -> Iterator[tuple[int, list[object | None]]]:
        try:
            handle = self.archive.open(sheet.path)
        except KeyError as exc:
            raise ConversionError(f"invalid XLSX workbook: missing worksheet {sheet.path}") from exc

        with handle:
            for _, element in ElementTree.iterparse(handle, events=("end",)):
                if element.tag != xml_name(SPREADSHEET_NS, "row"):
                    continue
                row_number = int(element.attrib.get("r", "0"))
                values: dict[int, object | None] = {}
                for cell in element.findall(xml_name(SPREADSHEET_NS, "c")):
                    reference = cell.attrib.get("r")
                    if not reference:
                        continue
                    values[column_index(reference)] = self._cell_value(cell)
                if values:
                    width = max(values) + 1
                    row = [None] * width
                    for index, value in values.items():
                        row[index] = value
                    yield row_number, row
                element.clear()

    def _cell_value(self, cell: ElementTree.Element) -> object | None:
        cell_type = cell.attrib.get("t", "n")
        if cell_type == "inlineStr":
            return "".join(node.text or "" for node in cell.iter(xml_name(SPREADSHEET_NS, "t")))

        value_node = cell.find(xml_name(SPREADSHEET_NS, "v"))
        if value_node is None or value_node.text is None:
            return None
        raw_value = value_node.text
        if cell_type == "s":
            try:
                return self.shared_strings[int(raw_value)]
            except (IndexError, ValueError) as exc:
                raise ConversionError(f"invalid shared string index {raw_value!r}") from exc
        if cell_type == "b":
            return raw_value == "1"
        if cell_type in {"str", "e"}:
            return raw_value
        return parse_number(raw_value)


def value_at(row: Sequence[object | None], index: int | None) -> object | None:
    if index is None or index >= len(row):
        return None
    return row[index]


def text_value(value: object | None) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def find_columns(headers: Sequence[object | None]) -> dict[str, int | None]:
    normalized = [normalize_header(value) for value in headers]
    columns: dict[str, int | None] = {}
    for field, aliases in COLUMN_ALIASES.items():
        indexes = [index for index, header in enumerate(normalized) if header in aliases]
        if len(indexes) > 1:
            raise ConversionError(f"multiple columns match {field}: {indexes}")
        columns[field] = indexes[0] if indexes else None
    missing = [field for field in ("source", "target") if columns[field] is None]
    if missing:
        raise ConversionError(f"missing required column(s): {', '.join(missing)}")
    return columns


def parse_boolean(value: object | None, default: bool, field: str, row_number: int) -> bool:
    if value is None or text_value(value) == "":
        return default
    if isinstance(value, bool):
        return value
    normalized = text_value(value).lower()
    if normalized in TRUE_VALUES:
        return True
    if normalized in FALSE_VALUES:
        return False
    raise ValueError(f"row {row_number}: {field} must be true or false, got {value!r}")


def parse_status_code(value: object | None, default: int, row_number: int) -> int:
    if value is None or text_value(value) == "":
        status_code = default
    else:
        try:
            numeric = float(text_value(value))
        except ValueError as exc:
            raise ValueError(f"row {row_number}: invalid status code {value!r}") from exc
        if not numeric.is_integer():
            raise ValueError(f"row {row_number}: invalid status code {value!r}")
        status_code = int(numeric)
    if status_code not in ALLOWED_STATUS_CODES:
        raise ValueError(f"row {row_number}: status code must be one of {sorted(ALLOWED_STATUS_CODES)}")
    return status_code


def parse_http_url(value: str, field: str, row_number: int) -> SplitResult:
    parsed = urlsplit(value)
    try:
        _ = parsed.port
    except ValueError as exc:
        raise ValueError(f"row {row_number}: invalid {field} port in {value!r}") from exc
    if parsed.scheme.lower() not in {"http", "https"} or not parsed.hostname:
        raise ValueError(f"row {row_number}: {field} must be an absolute HTTP(S) URL")
    if parsed.username is not None or parsed.password is not None:
        raise ValueError(f"row {row_number}: {field} must not contain credentials")
    return parsed


def source_key(parsed: SplitResult) -> tuple[str, str, str]:
    return ((parsed.hostname or "").lower(), parsed.path or "/", parsed.query)


def is_exact_match(value: object | None) -> bool:
    return normalize_header(value or "") in EXACT_MATCH_VALUES


def sheet_headers(workbook: XlsxWorkbook, sheet: SheetInfo, header_row: int) -> list[object | None]:
    for row_number, row in workbook.iter_rows(sheet):
        if row_number == header_row:
            return row
        if row_number > header_row:
            break
    raise ConversionError(f"sheet {sheet.name!r} does not contain header row {header_row}")


def select_sheet(workbook: XlsxWorkbook, sheet_name: str | None, header_row: int) -> SheetInfo:
    if sheet_name:
        for sheet in workbook.sheets:
            if sheet.name == sheet_name:
                return sheet
        available = ", ".join(sheet.name for sheet in workbook.sheets)
        raise ConversionError(f"sheet {sheet_name!r} not found; available sheets: {available}")

    candidates: list[SheetInfo] = []
    for sheet in workbook.sheets:
        try:
            find_columns(sheet_headers(workbook, sheet, header_row))
        except ConversionError:
            continue
        candidates.append(sheet)
    if not candidates:
        raise ConversionError("no worksheet contains recognized source and target URL columns")
    preferred = next((sheet for sheet in candidates if sheet.name == "最终规则总表"), None)
    return preferred or candidates[0]


def convert_workbook(input_path: Path, options: ConversionOptions) -> ConversionResult:
    problems = Problems()
    redirects: list[dict[str, object]] = []
    seen_sources: dict[tuple[str, str, str], tuple[int, str, str]] = {}
    seen_rule_ids: dict[str, int] = {}
    unsupported = Counter()
    unsupported_samples: list[tuple[int, str]] = []
    disabled_rows = 0
    stripped_fragments = 0
    deduplicated_rows = 0
    conflicting_rows = 0
    conflict_warnings: list[str] = []

    with XlsxWorkbook(input_path) as workbook:
        sheet = select_sheet(workbook, options.sheet_name, options.header_row)
        headers = sheet_headers(workbook, sheet, options.header_row)
        columns = find_columns(headers)

        for row_number, row in workbook.iter_rows(sheet):
            if row_number <= options.header_row or not any(value is not None for value in row):
                continue
            rule_id_value = text_value(value_at(row, columns["rule_id"]))
            if columns["rule_id"] is not None:
                if not rule_id_value:
                    problems.add(f"row {row_number}: rule ID is required")
                    continue
                if rule_id_value in seen_rule_ids:
                    problems.add(
                        f"row {row_number}: duplicate rule ID {rule_id_value!r}; "
                        f"first used by row {seen_rule_ids[rule_id_value]}"
                    )
                    continue
                seen_rule_ids[rule_id_value] = row_number
            rule_id = rule_id_value or f"row-{row_number}"

            try:
                enabled = parse_boolean(value_at(row, columns["enabled"]), True, "enabled", row_number)
            except ValueError as exc:
                problems.add(str(exc))
                continue
            if not enabled:
                disabled_rows += 1
                continue

            match_mode = value_at(row, columns["match_mode"])
            if columns["match_mode"] is not None and not is_exact_match(match_mode):
                mode_name = text_value(match_mode) or "<blank>"
                unsupported[mode_name] += 1
                if len(unsupported_samples) < 5:
                    unsupported_samples.append((row_number, rule_id))
                continue

            source_url = text_value(value_at(row, columns["source"]))
            target_url = text_value(value_at(row, columns["target"]))
            if not source_url:
                problems.add(f"row {row_number} ({rule_id}): source URL is required")
                continue
            if not target_url:
                problems.add(f"row {row_number} ({rule_id}): target URL is required")
                continue

            try:
                source = parse_http_url(source_url, "source URL", row_number)
                parse_http_url(target_url, "target URL", row_number)
                if source.fragment:
                    if not options.strip_source_fragments:
                        raise ValueError(
                            f"row {row_number} ({rule_id}): source URL contains fragment {source.fragment!r}; "
                            "use --strip-source-fragments only after reviewing this behavior"
                        )
                    source = source._replace(fragment="")
                    source_url = urlunsplit(source)
                    stripped_fragments += 1

                status_code = parse_status_code(
                    value_at(row, columns["status_code"]), options.default_status_code, row_number
                )
                preserve_query = parse_boolean(
                    value_at(row, columns["preserve_query"]),
                    options.default_preserve_query,
                    "preserveQueryString",
                    row_number,
                )
                subpath_matching = parse_boolean(
                    value_at(row, columns["subpath"]), False, "subpathMatching", row_number
                )
                if source.query and subpath_matching:
                    raise ValueError(
                        f"row {row_number} ({rule_id}): query source cannot use subpathMatching"
                    )
            except ValueError as exc:
                problems.add(str(exc))
                continue

            key = source_key(source)
            previous = seen_sources.get(key)
            if previous is not None:
                previous_row, previous_rule_id, previous_target = previous
                if options.deduplicate_same_target and previous_target == target_url:
                    deduplicated_rows += 1
                    continue
                if previous_target != target_url and options.skip_conflicting_targets:
                    conflicting_rows += 1
                    conflict_warnings.append(
                        f"row {row_number} ({rule_id}): conflicting target skipped for "
                        f"source {source_url!r}; kept row {previous_row} ({previous_rule_id}) "
                        f"target {previous_target!r}; skipped target {target_url!r}"
                    )
                    continue
                conflict = "conflicting targets" if previous_target != target_url else "duplicate source"
                problems.add(
                    f"row {row_number} ({rule_id}): {conflict} after host/path/query normalization; "
                    f"first used by row {previous_row} ({previous_rule_id})"
                )
                continue

            seen_sources[key] = (row_number, rule_id, target_url)
            redirects.append(
                {
                    "sourceURL": source_url,
                    "targetURL": target_url,
                    "statusCode": status_code,
                    "preserveQueryString": preserve_query,
                    "subpathMatching": subpath_matching,
                }
            )

    unsupported_rows = sum(unsupported.values())
    if unsupported_rows and not options.skip_unsupported:
        modes = ", ".join(f"{name}={count}" for name, count in sorted(unsupported.items()))
        samples = ", ".join(f"row {row} ({rule_id})" for row, rule_id in unsupported_samples)
        problems.add(
            f"unsupported match modes found ({modes}); examples: {samples}; "
            "use --skip-unsupported only when omitting them is intentional",
            count=unsupported_rows,
        )

    problems.raise_if_any()

    redirects.sort(key=lambda redirect: str(redirect["sourceURL"]))
    if len(redirects) < options.min_rules or len(redirects) > options.max_rules:
        raise ConversionError(
            f"generated rule count {len(redirects)} is outside allowed range "
            f"[{options.min_rules}, {options.max_rules}]"
        )
    return ConversionResult(
        sheet_name=sheet.name,
        redirects=redirects,
        disabled_rows=disabled_rows,
        unsupported_rows=unsupported_rows,
        stripped_fragments=stripped_fragments,
        deduplicated_rows=deduplicated_rows,
        conflicting_rows=conflicting_rows,
        conflict_warnings=tuple(conflict_warnings),
    )


def encode_result(result: ConversionResult, max_json_bytes: int) -> bytes:
    encoded = (
        json.dumps({"redirects": result.redirects}, ensure_ascii=False, separators=(",", ":")) + "\n"
    ).encode("utf-8")
    if len(encoded) > max_json_bytes:
        raise ConversionError(
            f"generated JSON is {len(encoded)} bytes, exceeding the plugin limit of {max_json_bytes} bytes"
        )
    return encoded


def write_atomic(output_path: Path, content: bytes) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path: str | None = None
    try:
        with tempfile.NamedTemporaryFile(dir=output_path.parent, delete=False) as temporary:
            temporary.write(content)
            temporary.flush()
            os.fsync(temporary.fileno())
            temporary_path = temporary.name
        os.chmod(temporary_path, 0o644)
        os.replace(temporary_path, output_path)
    finally:
        if temporary_path is not None and os.path.exists(temporary_path):
            os.unlink(temporary_path)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, required=True, help="input .xlsx or .xlsm workbook")
    parser.add_argument("--output", type=Path, required=True, help="output redirects.json path")
    parser.add_argument("--sheet", dest="sheet_name", help="worksheet name; auto-detected when omitted")
    parser.add_argument("--header-row", type=int, default=1)
    parser.add_argument("--status-code", type=int, default=301, choices=sorted(ALLOWED_STATUS_CODES))
    parser.add_argument("--preserve-query-string", action="store_true")
    parser.add_argument("--min-rules", type=int, default=1)
    parser.add_argument("--max-rules", type=int, default=30000)
    parser.add_argument("--max-json-bytes", type=int, default=DEFAULT_MAX_JSON_BYTES)
    parser.add_argument("--skip-unsupported", action="store_true")
    parser.add_argument("--strip-source-fragments", action="store_true")
    parser.add_argument("--deduplicate-same-target", action="store_true")
    parser.add_argument(
        "--skip-conflicting-targets",
        action="store_true",
        help="keep the first target and log later conflicting rows to standard error",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    if args.header_row < 1:
        print("FAILED: --header-row must be at least 1", file=sys.stderr)
        return 1
    if args.min_rules < 0 or args.max_rules < args.min_rules:
        print("FAILED: rule count range is invalid", file=sys.stderr)
        return 1
    if args.max_json_bytes < 1:
        print("FAILED: --max-json-bytes must be at least 1", file=sys.stderr)
        return 1
    if args.input.resolve() == args.output.resolve():
        print("FAILED: input and output paths must be different", file=sys.stderr)
        return 1
    options = ConversionOptions(
        sheet_name=args.sheet_name,
        header_row=args.header_row,
        default_status_code=args.status_code,
        default_preserve_query=args.preserve_query_string,
        min_rules=args.min_rules,
        max_rules=args.max_rules,
        max_json_bytes=args.max_json_bytes,
        skip_unsupported=args.skip_unsupported,
        strip_source_fragments=args.strip_source_fragments,
        deduplicate_same_target=args.deduplicate_same_target,
        skip_conflicting_targets=args.skip_conflicting_targets,
    )
    try:
        result = convert_workbook(args.input, options)
        encoded = encode_result(result, options.max_json_bytes)
        write_atomic(args.output, encoded)
    except (ConversionError, OSError) as exc:
        print(f"FAILED: {exc}", file=sys.stderr)
        return 1

    for warning in result.conflict_warnings:
        print(f"WARNING: {warning}", file=sys.stderr)
    print(f"sheet={result.sheet_name}")
    print(f"generated_rules={len(result.redirects)}")
    print(f"disabled_rows={result.disabled_rows}")
    print(f"unsupported_rows={result.unsupported_rows}")
    print(f"stripped_fragments={result.stripped_fragments}")
    print(f"deduplicated_rows={result.deduplicated_rows}")
    print(f"conflicting_rows={result.conflicting_rows}")
    print(f"json_bytes={len(encoded)}")
    print(f"json_sha256={hashlib.sha256(encoded).hexdigest()}")
    print(f"output={args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
