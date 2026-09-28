from __future__ import annotations

import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

sys.path.insert(0, str(Path(__file__).resolve().parent))

from xlsx_to_redirects import (
    ConversionError,
    ConversionOptions,
    convert_workbook,
    encode_result,
)


CHINESE_HEADERS = [
    "匹配顺序(自上而下首次命中)",
    "规则ID",
    "规则类型",
    "匹配方式",
    "匹配串(精确URL或正则)",
    "301目标(新站)",
    "原PV/整合条数",
    "备注",
]


def column_name(index: int) -> str:
    result = ""
    value = index + 1
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def write_workbook(path: Path, rows: list[list[object | None]], sheet_name: str = "最终规则总表") -> None:
    shared_strings: list[str] = []
    string_indexes: dict[str, int] = {}

    def shared_index(value: object) -> int:
        text = str(value)
        if text not in string_indexes:
            string_indexes[text] = len(shared_strings)
            shared_strings.append(text)
        return string_indexes[text]

    row_xml: list[str] = []
    for row_number, row in enumerate(rows, start=1):
        cells: list[str] = []
        for column, value in enumerate(row):
            if value is None:
                continue
            reference = f"{column_name(column)}{row_number}"
            index = shared_index(value)
            cells.append(f'<c r="{reference}" t="s"><v>{index}</v></c>')
        row_xml.append(f'<row r="{row_number}">{"".join(cells)}</row>')

    shared_xml = "".join(f"<si><t>{escape(value)}</t></si>" for value in shared_strings)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<sheets><sheet name="{escape(sheet_name)}" sheetId="1" r:id="rId1"/></sheets>'
            "</workbook>",
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            'Target="worksheets/sheet1.xml"/>'
            "</Relationships>",
        )
        archive.writestr(
            "xl/sharedStrings.xml",
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f"{shared_xml}</sst>",
        )
        archive.writestr(
            "xl/worksheets/sheet1.xml",
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<sheetData>{"".join(row_xml)}</sheetData></worksheet>',
        )


def exact_row(rule_id: str, source: str, target: str) -> list[object]:
    return [1, rule_id, "逐条保留", "精确URL", source, target, 10, "test"]


class XlsxToRedirectsTest(unittest.TestCase):
    def convert(self, rows: list[list[object | None]], **overrides: object):
        with tempfile.TemporaryDirectory() as temporary_directory:
            workbook_path = Path(temporary_directory) / "redirects.xlsx"
            write_workbook(workbook_path, [CHINESE_HEADERS, *rows])
            options = ConversionOptions(**overrides)
            return convert_workbook(workbook_path, options)

    def test_converts_chinese_columns(self) -> None:
        result = self.convert(
            [
                exact_row(
                    "K-00001",
                    "https://old.example.com/path?a=1&b=2",
                    "https://new.example.com/target",
                )
            ]
        )

        self.assertEqual("最终规则总表", result.sheet_name)
        self.assertEqual(1, len(result.redirects))
        self.assertEqual("https://old.example.com/path?a=1&b=2", result.redirects[0]["sourceURL"])
        self.assertEqual(301, result.redirects[0]["statusCode"])
        self.assertFalse(result.redirects[0]["preserveQueryString"])

    def test_converts_raw_source_sheet_as_exact_redirects(self) -> None:
        headers = [
            "品类",
            "页面URL",
            "浏览量(PV)",
            "访客数(UV)",
            "贡献下游浏览量",
            "退出页次数",
            "平均停留时长",
            "新网站映射链接",
        ]
        row = [
            "箱包",
            "https://old.example.com/list?page=2",
            100,
            80,
            20,
            60,
            "00:01:15",
            "https://new.example.com/bags",
        ]

        with tempfile.TemporaryDirectory() as temporary_directory:
            workbook_path = Path(temporary_directory) / "seo-mapping.xlsx"
            write_workbook(workbook_path, [headers, row], sheet_name="源表")
            result = convert_workbook(
                workbook_path,
                ConversionOptions(sheet_name="源表"),
            )

        self.assertEqual("源表", result.sheet_name)
        self.assertEqual(1, len(result.redirects))
        self.assertEqual("https://old.example.com/list?page=2", result.redirects[0]["sourceURL"])
        self.assertEqual("https://new.example.com/bags", result.redirects[0]["targetURL"])
        self.assertFalse(result.redirects[0]["subpathMatching"])

    def test_raw_query_order_is_distinct(self) -> None:
        result = self.convert(
            [
                exact_row("K-00001", "https://old.example.com/path?a=1&b=2", "https://new.example.com/a"),
                exact_row("K-00002", "https://old.example.com/path?b=2&a=1", "https://new.example.com/b"),
            ]
        )

        self.assertEqual(2, len(result.redirects))

    def test_unsupported_match_mode_requires_explicit_skip(self) -> None:
        regex_row = [2, "R-00001", "分页聚合", "正则", "^https://old.example.com/.*$", "https://new.example.com", 20, ""]
        with self.assertRaisesRegex(ConversionError, "unsupported match modes"):
            self.convert([regex_row])

        result = self.convert(
            [
                exact_row("K-00001", "https://old.example.com/path", "https://new.example.com/path"),
                regex_row,
            ],
            skip_unsupported=True,
        )
        self.assertEqual(1, len(result.redirects))
        self.assertEqual(1, result.unsupported_rows)

    def test_fragment_stripping_and_same_target_deduplication_are_explicit(self) -> None:
        rows = [
            exact_row("K-00001", "https://old.example.com/path#anchor", "https://new.example.com/path"),
            exact_row("K-00002", "http://OLD.example.com/path", "https://new.example.com/path"),
        ]
        with self.assertRaisesRegex(ConversionError, "source URL contains fragment"):
            self.convert(rows)

        result = self.convert(
            rows,
            strip_source_fragments=True,
            deduplicate_same_target=True,
        )
        self.assertEqual(1, len(result.redirects))
        self.assertEqual(1, result.stripped_fragments)
        self.assertEqual(1, result.deduplicated_rows)
        self.assertEqual("https://old.example.com/path", result.redirects[0]["sourceURL"])

    def test_conflicting_normalized_source_always_fails(self) -> None:
        with self.assertRaisesRegex(ConversionError, "conflicting targets"):
            self.convert(
                [
                    exact_row("K-00001", "https://old.example.com/path", "https://new.example.com/a"),
                    exact_row("K-00002", "http://OLD.example.com/path", "https://new.example.com/b"),
                ],
                deduplicate_same_target=True,
            )

    def test_duplicate_rule_id_fails(self) -> None:
        with self.assertRaisesRegex(ConversionError, "duplicate rule ID"):
            self.convert(
                [
                    exact_row("K-00001", "https://old.example.com/a", "https://new.example.com/a"),
                    exact_row("K-00001", "https://old.example.com/b", "https://new.example.com/b"),
                ]
            )

    def test_json_size_limit_is_enforced(self) -> None:
        result = self.convert(
            [exact_row("K-00001", "https://old.example.com/path", "https://new.example.com/path")]
        )
        with self.assertRaisesRegex(ConversionError, "exceeding the plugin limit"):
            encode_result(result, max_json_bytes=10)


if __name__ == "__main__":
    unittest.main()
