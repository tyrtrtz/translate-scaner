"""Match approved source sheets using a disk index; save baseline copies only."""

import os
import sqlite3
import tempfile
import shutil
import re
import json
from io import TextIOWrapper
from zipfile import ZipFile
from xml.sax.expatreader import create_parser
from xml.sax.handler import feature_external_ges
from xml.sax.saxutils import XMLGenerator, escape
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from ignored_sheets import sheet_key
from scanner import SUPPORTED, hidden_row_numbers, normalize, scan_folder


def match_key(value):
    return value.strip() if isinstance(value, str) else ""


def boundary_newlines(value):
    """Detect CR/LF in the whitespace before and after the baseline text."""
    return tuple(any(c in edge for c in "\r\n") for edge in (
        re.match(r"^\s*", value).group(), re.search(r"\s*$", value).group()))


def align_boundary_newlines(value, keep_start, keep_end):
    """Align boundary newline presence; retain spaces and internal newlines."""
    if not isinstance(value, str):
        return value
    without_newlines = lambda match: match.group().replace("\r", "").replace("\n", "")
    if not keep_start:
        value = re.sub(r"^\s+", without_newlines, value, count=1)
    elif value.strip() and not boundary_newlines(value)[0]:
        value = "\n" + value
    if not keep_end:
        value = re.sub(r"\s+$", without_newlines, value, count=1)
    elif value.strip() and not boundary_newlines(value)[1]:
        value += "\n"
    return value


def text_cell(sheet, row, column, value):
    cell = sheet.cell(row, column, value)
    if isinstance(value, str):
        cell.data_type = "s"  # A source string beginning with '=' is text, never a formula.
    return cell


@dataclass
class MatchReport:
    destination: Path
    check_report: object
    baseline_rows: int
    matched_rows: int
    source_rows: int
    source_matched: int
    skipped: int
    outputs: list
    missing_translations: int = 0
    conflict_texts: int = 0
    match_pairs: int = 0
    recommended_rows: int = 0
    baseline_files: int = 0


def baseline_sheets(workbook, path, header_rows):
    for sheet in workbook:
        if hasattr(sheet, "reset_dimensions"):
            sheet.reset_dimensions()
        candidates = []
        for number, row in enumerate(sheet.iter_rows(max_row=header_rows), 1):
            headers = {cell.column: normalize(cell.value) for cell in row if cell.value is not None}
            for column, title in headers.items():
                if title == "Text":
                    candidates.append((number, column, headers))
        location = f"{path.name} / {sheet.title}"
        if not candidates:
            raise ValueError(f"基准表头缺失：{location} 的第 1～{header_rows} 行未找到 Text（完整名称匹配，忽略首尾空白）")
        if len(candidates) > 1:
            positions = "、".join(f"{get_column_letter(column)}{row}" for row, column, _ in candidates)
            raise ValueError(f"基准表头不明确：{location} 找到多个 Text：{positions}，请保留唯一的 Text 表头")
        header_row, text_column, headers = candidates[0]
        if "匹配中文1" in headers.values() or "推荐状态" in headers.values():
            raise ValueError(f"{location} 已有检索结果，请提供未回填的基准文件")
        output_column = max(18, max(headers) + 1)
        if output_column + 15 > 16384:
            raise ValueError(f"{location} 的表头右侧不足 16 列，无法写入检索结果")
        yield sheet, header_row, text_column, output_column


class BaselineWriter(XMLGenerator):
    """Stream original worksheet XML unchanged in meaning, adding result cells."""

    def __init__(self, stream, connection, path, sheet, header_row, output_column):
        # Avoid the platform's automatic LF -> CRLF conversion in worksheet XML.
        stream = TextIOWrapper(stream, encoding="utf-8", newline="\n", write_through=True)
        super().__init__(stream, encoding="utf-8", short_empty_elements=True)
        self.connection, self.path, self.sheet = connection, str(path), sheet
        self.header_row = header_row
        self.column_shift = output_column - 18
        self.output_letters = {get_column_letter(c) for c in range(output_column, output_column + 16)}
        self.row = None
        self.protected_cell = False
        self.protected_value = False
        self.prefix = ""
        self.omit_dimension = False
        self.has_columns = False
        self.matched_rows = 0

    def startElement(self, name, attrs):
        local = name.rsplit(":", 1)[-1]
        if local == "worksheet":
            self.prefix = name[:-len(local)]
        if local == "dimension":
            # Optional metadata can be wrong in the customer's baseline files.
            # Without it readers determine dimensions from the actual cells.
            self.omit_dimension = True
            return
        if local == "cols":
            self.has_columns = True
        if local == "sheetData" and not self.has_columns:
            super().startElement(self.prefix + "cols", {})
            self.output_columns()
            super().endElement(self.prefix + "cols")
        if local == "row":
            self.row = int(attrs["r"])
        if local == "c":
            column = re.match(r"[A-Z]+", attrs.get("r", "")).group()
            self.protected_cell = column in self.output_letters
        if self.protected_cell and local == "f":
            self.existing_content()
        if self.protected_cell and local in ("v", "t"):
            self.protected_value = True
        if self.protected_cell:
            return
        super().startElement(name, attrs)

    def characters(self, content):
        if self.protected_cell and self.protected_value and content:
            self.existing_content()
        if not self.omit_dimension and not self.protected_cell:
            self.xml_characters(content)

    def xml_characters(self, content):
        # XML parsers normalize literal CR/CRLF; character references preserve
        # retained source newlines and the original baseline cell values.
        if "\r" in content:
            self._finish_pending_start_element()
            self._write(escape(content, {"\r": "&#13;"}))
        else:
            super().characters(content)

    def existing_content(self):
        first, last = (get_column_letter(c + self.column_shift) for c in (18, 33))
        raise ValueError(f"{Path(self.path).name} / {self.sheet} 的 {first}～{last} 列已有内容，请提供未回填的基准文件")

    def output_columns(self):
        for first, last, width in ((18, 26, 45), (27, 29, 18), (30, 32, 55), (33, 33, 24)):
            super().startElement(self.prefix + "col", {"min": str(first + self.column_shift), "max": str(last + self.column_shift), "width": str(width), "customWidth": "1"})
            super().endElement(self.prefix + "col")

    def output_cell(self, column, value):
        if value is None:
            return
        attrs = {"r": f"{get_column_letter(column + self.column_shift)}{self.row}"}
        if isinstance(value, str):
            attrs["t"] = "inlineStr"
        super().startElement(self.prefix + "c", attrs)
        if isinstance(value, str):
            super().startElement(self.prefix + "is", {})
            super().startElement(self.prefix + "t", {"xml:space": "preserve"})
            self.xml_characters(value)
            super().endElement(self.prefix + "t")
            super().endElement(self.prefix + "is")
        else:
            super().startElement(self.prefix + "v", {})
            super().characters(str(value))
            super().endElement(self.prefix + "v")
        super().endElement(self.prefix + "c")

    def endElement(self, name):
        local = name.rsplit(":", 1)[-1]
        if local == "dimension":
            self.omit_dimension = False
            return
        if local in ("v", "t"):
            self.protected_value = False
        if local == "c":
            omitted = self.protected_cell
            self.protected_cell = False
            if omitted:
                return
        if self.protected_cell:
            return
        if local == "cols":
            self.output_columns()
        if local == "row":
            if self.row == self.header_row:
                for column in range(18, 27):
                    title = ("匹配中文", "匹配韩语", "来源路径 / Sheet / 行号")[(column - 18) % 3]
                    self.output_cell(column, f"{title}{(column - 18) // 3 + 1}")
                self.output_cell(27, "匹配条数")
                self.output_cell(28, "缺译条数")
                self.output_cell(29, "译文冲突")
                for column, title in enumerate(("推荐韩语", "推荐来源", "推荐依据", "推荐状态"), 30):
                    self.output_cell(column, title)
            else:
                record = self.connection.execute("SELECT text,leading_newline,trailing_newline FROM baseline_rows WHERE path=? AND sheet=? AND row=?",
                                              (self.path, self.sheet, self.row)).fetchone()
                if record:
                    key = (record[0],)
                    keep_start, keep_end = record[1:]
                    summary = self.connection.execute("SELECT n,missing FROM counts WHERE text=?", key).fetchone()
                    count, missing = summary if summary else (0, 0)
                    self.matched_rows += bool(count)
                    if 0 < count <= 3:
                        hits = self.connection.execute("SELECT chinese,korean,path,sheet,row FROM hits WHERE text=? ORDER BY rowid LIMIT 3", key)
                        for index, (chinese, korean, filename, title, number) in enumerate(hits):
                            chinese = align_boundary_newlines(chinese, keep_start, keep_end)
                            korean = align_boundary_newlines(korean, keep_start, keep_end)
                            for offset, value in enumerate((chinese, korean, f"{filename}；Sheet：{title}；行号：{number}")):
                                self.output_cell(18 + index * 3 + offset, value)
                    self.output_cell(27, count)
                    self.output_cell(28, missing)
                    conflict = self.connection.execute("SELECT 1 FROM conflicts WHERE text=?", key).fetchone()
                    self.output_cell(29, "是" if conflict else "否" if count else "未匹配")
                    recommendation = self.connection.execute("""
                        SELECT s.korean,s.path,s.sheet,s.row,r.basis,r.status
                        FROM recommendations r LEFT JOIN source_rows s ON s.rowid=r.source_id
                        WHERE r.path=? AND r.sheet=? AND r.row=?
                    """, (self.path, self.sheet, self.row)).fetchone()
                    korean, filename, title, number, basis, status = recommendation
                    self.output_cell(30, align_boundary_newlines(korean, keep_start, keep_end))
                    if filename is not None:
                        self.output_cell(31, f"{filename}；Sheet：{title}；行号：{number}")
                    self.output_cell(32, basis)
                    self.output_cell(33, status)
            self.row = None
        super().endElement(name)


def write_baseline_copy(path, destination, connection, sheets):
    matched = 0
    with ZipFile(path) as source, ZipFile(destination, "w") as output:
        for item in source.infolist():
            with source.open(item) as original, output.open(item, "w") as copied:
                if item.filename in sheets:
                    writer = BaselineWriter(copied, connection, path, *sheets[item.filename])
                    parser = create_parser()
                    parser.setFeature(feature_external_ges, False)
                    parser.setContentHandler(writer)
                    parser.parse(original)
                    matched += writer.matched_rows
                else:
                    shutil.copyfileobj(original, copied, length=1024 * 1024)
    return matched


def write_detail_sheet(book, title, headers, rows, widths, progress=None, row_limit=1048576):
    """Stream data rows; split automatically at Excel's worksheet row limit."""
    sheet = None
    page = count = total = 0
    for values in rows:
        if sheet is None or count == row_limit:
            if sheet is not None:
                sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{count}"
            page += 1
            sheet = book.create_sheet(title if page == 1 else f"{title}_{page}")
            sheet.freeze_panes = "A2"
            for column, width in enumerate(widths, 1):
                sheet.column_dimensions[get_column_letter(column)].width = width
            header_cells = []
            for value in headers:
                cell = WriteOnlyCell(sheet, value=value)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="245B78")
                header_cells.append(cell)
            sheet.append(header_cells)
            count = 1
        cells = []
        for value in values:
            cell = WriteOnlyCell(sheet, value=value)
            if isinstance(value, str):
                cell.data_type = "s"
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if headers[len(cells)].endswith("率"):
                cell.number_format = "0.00%"
            cells.append(cell)
        sheet.append(cells)
        count += 1
        total += 1
        if progress and total % 5000 == 0:
            progress(f"写出{title}：{total} 条")
    if sheet is None:
        sheet = book.create_sheet(title)
        sheet.freeze_panes = "A2"
        sheet.append(headers)
        count = 1
        for column, width in enumerate(widths, 1):
            sheet.column_dimensions[get_column_letter(column)].width = width
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{count}"
    return total


def read_context(connection, table, path, sheet, sequence):
    # Table names are fixed by the two callers. Only six nearby records are read.
    records = list(connection.execute(f"""
        SELECT seq,row,text FROM {table} WHERE path=? AND sheet=? AND seq BETWEEN ? AND ? ORDER BY seq
    """, (path, sheet, sequence - 3, sequence + 3)))
    previous = [[row, text] for seq, row, text in reversed(records) if seq < sequence]
    following = [[row, text] for seq, row, text in records if seq > sequence]
    return previous, following


def context_text(records, direction):
    return "\n".join(f"{direction}{index}句 [行{row}] {text}" for index, (row, text) in enumerate(records, 1))


def build_recommendations(connection, progress=None):
    """Compare each baseline occurrence against its candidates; keep evidence on disk."""
    connection.executescript("""
        CREATE TABLE source_context (source_id INTEGER PRIMARY KEY, previous TEXT, following TEXT);
        CREATE TABLE baseline_context (path TEXT,sheet TEXT,row INTEGER,previous TEXT,following TEXT,
            PRIMARY KEY(path,sheet,row));
        CREATE TABLE comparisons (path TEXT,sheet TEXT,row INTEGER,source_id INTEGER,
            previous_equal INTEGER,following_equal INTEGER,full_sides INTEGER,
            PRIMARY KEY(path,sheet,row,source_id));
        CREATE TABLE recommendations (path TEXT,sheet TEXT,row INTEGER,source_id INTEGER,basis TEXT,status TEXT,
            PRIMARY KEY(path,sheet,row));
    """)
    for source_id, path, sheet, sequence in connection.execute("SELECT rowid,path,sheet,seq FROM hits"):
        previous, following = read_context(connection, "source_rows", path, sheet, sequence)
        connection.execute("INSERT INTO source_context VALUES (?,?,?)",
                           (source_id, json.dumps(previous, ensure_ascii=False), json.dumps(following, ensure_ascii=False)))
    recommended = 0
    for index, (path, sheet, row, text, sequence) in enumerate(connection.execute(
            "SELECT path,sheet,row,text,seq FROM baseline_rows ORDER BY path,sheet,row"), 1):
        previous, following = read_context(connection, "baseline_rows", path, sheet, sequence)
        connection.execute("INSERT INTO baseline_context VALUES (?,?,?,?,?)",
                           (path, sheet, row, json.dumps(previous, ensure_ascii=False), json.dumps(following, ensure_ascii=False)))
        count = connection.execute("SELECT n,missing,variants FROM counts WHERE text=?", (text,)).fetchone()
        best_score, best_id, best_key, tied = (-1, -1), None, None, False
        best_previous = best_following = 0
        candidates = connection.execute("""
            SELECT s.rowid,s.korean_key,c.previous,c.following FROM hits s
            JOIN source_context c ON c.source_id=s.rowid WHERE s.text=? ORDER BY s.rowid
        """, (text,))
        for source_id, korean_key, source_previous, source_following in candidates:
            sp, sf = json.loads(source_previous), json.loads(source_following)
            pe = sum(a[1] == b[1] for a, b in zip(previous, sp))
            fe = sum(a[1] == b[1] for a, b in zip(following, sf))
            full = int(len(previous) == len(sp) == pe == 3) + int(len(following) == len(sf) == fe == 3)
            connection.execute("INSERT INTO comparisons VALUES (?,?,?,?,?,?,?)", (path, sheet, row, source_id, pe, fe, full))
            if not korean_key:
                continue
            score = (full, pe + fe)
            if score > best_score:
                best_score, best_id, best_key, tied = score, source_id, korean_key, False
                best_previous, best_following = pe, fe
            elif score == best_score and korean_key != best_key:
                tied = True
        if count is None:
            status, basis, best_id = "未匹配", "没有中文全文匹配的源行", None
        elif count[2] == 0:
            status, basis, best_id = "全部缺译", "中文已匹配，但所有来源韩语均为空或全空白", None
        elif count[2] == 1:
            status, basis = "唯一译文", "全部非空韩语去除首尾空白后相同；选取上下文得分最高的来源，并列取首条"
        elif best_score[0] == 0:
            status, basis, best_id = "待确认", "存在多种译文，没有任一来源的完整前3句或完整后3句一致", None
        elif tied:
            status, basis, best_id = "待确认", "存在多种译文，最高上下文得分对应不同韩语，无法唯一推荐", None
        else:
            status = "上下文推荐"
            basis = f"前{best_previous}/3句一致，后{best_following}/3句一致；完整一致侧数{best_score[0]}，一致位置数{best_score[1]}"
        connection.execute("INSERT INTO recommendations VALUES (?,?,?,?,?,?)", (path, sheet, row, best_id, basis, status))
        recommended += best_id is not None
        if progress and index % 5000 == 0:
            progress(f"比较上下文并推荐译文：{index} 个基准行")
    connection.commit()
    return recommended


def write_baseline_statistics(destination, connection, baseline_parts):
    headers = ["非空中文行数", "中文匹配行数", "有译文行数", "推荐行数", "未匹配行数", "命中但全缺译行数",
               "冲突行数", "冲突已推荐行数", "冲突待确认行数", "中文匹配率", "译文覆盖率", "推荐率"]

    def counts(path, sheet=None):
        clause = "b.path=?" + (" AND b.sheet=?" if sheet is not None else "")
        arguments = (str(path), sheet) if sheet is not None else (str(path),)
        n, matched, available, recommended, missing, conflicts, resolved = connection.execute(f"""
            SELECT COUNT(*),COALESCE(SUM(c.n>0),0),COALESCE(SUM(c.variants>0),0),
                   COALESCE(SUM(r.source_id IS NOT NULL),0),COALESCE(SUM(c.variants=0),0),
                   COUNT(f.text),COALESCE(SUM(f.text IS NOT NULL AND r.source_id IS NOT NULL),0)
            FROM baseline_rows b LEFT JOIN counts c ON c.text=b.text
            LEFT JOIN conflicts f ON f.text=b.text
            JOIN recommendations r ON r.path=b.path AND r.sheet=b.sheet AND r.row=b.row WHERE {clause}
        """, arguments).fetchone()
        return [n, matched, available, recommended, n - matched, missing, conflicts, resolved, conflicts - resolved]

    def with_rates(values):
        n = values[0]
        return [*values, *(values[i] / n if n else None for i in (1, 2, 3))]

    book = Workbook(write_only=True)
    file_rows, sheet_rows, total = [], [], [0] * 9
    for path, parts in baseline_parts.items():
        values = counts(path)
        total = [a + b for a, b in zip(total, values)]
        file_rows.append([path.name, str(path), *with_rates(values)])
        for sheet, _, _ in parts.values():
            sheet_rows.append([path.name, str(path), sheet, *with_rates(counts(path, sheet))])
    write_detail_sheet(book, "文件统计", ["文件名", "完整路径", *headers], file_rows, [35, 65, *([20] * len(headers))])
    write_detail_sheet(book, "Sheet统计", ["文件名", "完整路径", "Sheet", *headers], sheet_rows, [35, 65, 25, *([20] * len(headers))])
    write_detail_sheet(book, "总体统计", ["基准文件数", "基准 Sheet 数", *headers],
                       [[len(baseline_parts), len(sheet_rows), *with_rates(total)]], [20] * (len(headers) + 2))
    book.save(destination)
    book.close()


def write_matching_details(destination, connection, progress=None):
    book = Workbook(write_only=True)
    matches = connection.execute("""
        SELECT b.path,b.sheet,b.row,s.path,s.sheet,s.row,s.chinese,s.korean,
               CASE WHEN s.korean_key='' THEN '是' ELSE '否' END,
               CASE WHEN c.text IS NOT NULL THEN '是' ELSE '否' END,
               bc.previous,bc.following,sc.previous,sc.following,
               e.previous_equal,e.following_equal,e.full_sides,
               CASE WHEN r.source_id=s.rowid THEN '是' ELSE '否' END,r.status,r.basis
        FROM source_rows s JOIN baseline_rows b ON b.text=s.text
        LEFT JOIN conflicts c ON c.text=s.text
        JOIN baseline_context bc ON bc.path=b.path AND bc.sheet=b.sheet AND bc.row=b.row
        JOIN source_context sc ON sc.source_id=s.rowid
        JOIN comparisons e ON e.path=b.path AND e.sheet=b.sheet AND e.row=b.row AND e.source_id=s.rowid
        JOIN recommendations r ON r.path=b.path AND r.sheet=b.sheet AND r.row=b.row
        ORDER BY s.rowid,b.path,b.sheet,b.row
    """)
    rows = ([Path(b).name,b,bs,br,Path(p).name,p,ss,sr,cn,ko,missing,conflict,
             context_text(json.loads(bp), "前"),context_text(json.loads(bf), "后"),
             context_text(json.loads(sp), "前"),context_text(json.loads(sf), "后"),
             pe,fe,full,pe+fe,selected,status,basis]
            for b,bs,br,p,ss,sr,cn,ko,missing,conflict,bp,bf,sp,sf,pe,fe,full,selected,status,basis in matches)
    pairs = write_detail_sheet(book, "匹配明细",
                               ["基准文件名","基准完整路径","基准 Sheet","基准行号","源文件名","源完整路径","源 Sheet","源行号","源中文","韩语","缺译","译文冲突",
                                "基准前3句","基准后3句","来源前3句","来源后3句","前句一致数","后句一致数","完整一致侧数","一致位置总数","推荐来源","推荐状态","推荐依据"],
                               rows, [30,65,25,16,35,65,25,16,55,55,12,16,65,65,65,65,18,18,20,20,16,24,70], progress)
    for title, condition in (("未匹配明细", "t.text IS NULL"), ("缺译明细", "s.korean_key=''"), ("译文冲突", "c.text IS NOT NULL")):
        # Conditions are fixed program constants, never user input.
        records = connection.execute(f"""
            SELECT s.path,s.sheet,s.row,s.chinese,s.korean,
                   CASE WHEN t.text IS NOT NULL THEN '是' ELSE '否' END,
                   CASE WHEN s.korean_key='' THEN '是' ELSE '否' END,
                   (SELECT COUNT(DISTINCT v.korean_key) FROM source_rows v WHERE v.text=s.text AND v.korean_key!='')
            FROM source_rows s LEFT JOIN targets t ON t.text=s.text
            LEFT JOIN conflicts c ON c.text=s.text WHERE {condition} ORDER BY s.text,s.rowid
        """)
        rows = ([Path(p).name,p,sheet,row,cn,ko,matched,missing,variants]
                for p,sheet,row,cn,ko,matched,missing,variants in records)
        write_detail_sheet(book, title, ["源文件名","源完整路径","源 Sheet","源行号","源中文","韩语","命中基准","缺译","非空韩语种类数"],
                           rows, [35,65,25,16,55,55,16,12,24], progress)
    book.save(destination)
    book.close()
    return pairs


def issue_counts(connection, path, sheet=None):
    clause = "s.path=?" + (" AND s.sheet=?" if sheet is not None else "")
    arguments = (str(path), sheet) if sheet is not None else (str(path),)
    return connection.execute(f"""
        SELECT COALESCE(SUM(s.korean_key=''),0),COUNT(DISTINCT c.text)
        FROM source_rows s LEFT JOIN conflicts c ON c.text=s.text WHERE {clause}
    """, arguments).fetchone()


def write_statistics(destination, file_rows, sheet_rows, skipped):
    book = Workbook()
    book.remove(book.active)
    for title, headers, rows, widths in (
        ("文件统计", ["文件名", "完整路径", "参与 Sheet 数", "跳过 Sheet / 文件项数", "非空中文行数", "匹配行数", "未匹配行数", "匹配率", "处理状态", "缺译行数", "冲突中文数"], file_rows, [38, 70, 18, 24, 18, 18, 18, 16, 24,18,18]),
        ("Sheet统计", ["文件名", "完整路径", "Sheet", "非空中文行数", "匹配行数", "未匹配行数", "匹配率", "缺译行数", "冲突中文数"], sheet_rows, [38, 70, 30, 18, 18, 18, 16,18,18]),
        ("跳过清单", ["文件名", "完整路径", "Sheet", "跳过原因"], skipped, [38, 70, 30, 85]),
    ):
        sheet = book.create_sheet(title)
        for row_number, values in enumerate([headers, *rows], 1):
            for column, value in enumerate(values, 1):
                cell = text_cell(sheet, row_number, column, value)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                if row_number == 1:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="245B78")
                elif headers[column - 1] == "匹配率":
                    cell.number_format = "0.00%"
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        sheet.row_dimensions[1].height = 30
        for column, width in enumerate(widths, 1):
            sheet.column_dimensions[get_column_letter(column)].width = width
    book.save(destination)
    book.close()


def match_folder(report, baseline_root, output_root, ignored_sheets=(), progress=None):
    baseline_root = Path(baseline_root).expanduser().resolve()
    output_root = Path(output_root).expanduser().resolve()
    source_root = report.root.resolve()
    if not baseline_root.is_dir():
        raise ValueError("请选择存在的基准文件夹")
    if source_root == baseline_root or source_root in baseline_root.parents or baseline_root in source_root.parents:
        raise ValueError("基准文件夹和源文件夹必须分开，不能互相包含")
    if any(output_root == root or root in output_root.parents for root in (source_root, baseline_root)):
        raise ValueError("输出文件夹不能位于源文件夹或基准文件夹内，以免再次扫描结果")
    baselines = sorted(p for p in baseline_root.rglob("*") if p.is_file()
                       and p.suffix.lower() in SUPPORTED and not p.name.startswith("~$"))
    if not baselines:
        raise ValueError("基准文件夹中没有 .xlsx 或 .xlsm 文件")
    ignored = set(ignored_sheets)
    previous_exceptions = {sheet_key(r.path, r.sheet): r for r in report.exceptions}
    # Recheck at execution time so changed/new sheets cannot bypass validation.
    fresh = scan_folder(source_root, report.header_rows, report.chinese_headers,
                        report.korean_headers, ignored_sheets=ignored,
                        recursive=report.recursive, include_hidden_sheets=report.include_hidden_sheets,
                        include_hidden_rows=report.include_hidden_rows,
                        progress=lambda i, n, p: progress(f"复查表头 {i}/{n}：{p.name}") if progress else None)
    candidates, skipped = {}, []
    for result in fresh.skipped_sheets:
        skipped.append([result.path.name, str(result.path), result.sheet, result.issue])
    for index, result in enumerate(fresh.results):
        key = sheet_key(result.path, result.sheet)
        if key in previous_exceptions:
            if result.passed:
                result = previous_exceptions[key]
                fresh.results[index] = result
        if result.passed and sheet_key(result.path, "") in previous_exceptions:
            result.issue = previous_exceptions[sheet_key(result.path, "")].issue
        if result.passed and sheet_key(result.path, "") not in previous_exceptions:
            candidates.setdefault(result.path, []).append(result)
        else:
            reason = result.issue or previous_exceptions[sheet_key(result.path, "")].issue
            skipped.append([result.path.name, str(result.path), result.sheet, "异常：" + reason])
    # Include only ignore records that actually exist in this source folder.
    for path in fresh.file_paths:
        records = [sheet for filename, sheet in ignored if filename == sheet_key(path, "")[0]]
        if records:
            try:
                book = load_workbook(path, read_only=True, keep_links=False)
                try:
                    for sheet in book.sheetnames:
                        if sheet_key(path, sheet) in ignored:
                            skipped.append([path.name, str(path), sheet, "本地已忽略"])
                finally:
                    book.close()
            except Exception:
                pass  # File-level read errors already appear in the header report.
    if not candidates:
        raise ValueError("没有可参与匹配的 Sheet；请先检查表头并处理异常或调整忽略记录")
    output_root.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="匹配临时_", dir=output_root) as temporary:
        staging = Path(temporary)
        connection = sqlite3.connect(staging / "index.sqlite3")
        baseline_rows = matched_rows = 0
        sheet_rows = []
        try:
            connection.executescript("""
                CREATE TABLE targets (text TEXT PRIMARY KEY);
                CREATE TABLE baseline_rows (path TEXT, sheet TEXT, row INTEGER, text TEXT, seq INTEGER,
                    leading_newline INTEGER,trailing_newline INTEGER,PRIMARY KEY (path, sheet, row));
                CREATE INDEX baseline_text ON baseline_rows(text);
                CREATE UNIQUE INDEX baseline_sequence ON baseline_rows(path,sheet,seq);
                CREATE TABLE source_rows (text TEXT, chinese TEXT, korean TEXT, path TEXT, sheet TEXT, row INTEGER, korean_key TEXT, seq INTEGER);
                CREATE INDEX source_text ON source_rows(text);
                CREATE INDEX source_position ON source_rows(path,sheet);
                CREATE UNIQUE INDEX source_sequence ON source_rows(path,sheet,seq);
                CREATE VIEW hits AS SELECT s.rowid AS rowid,s.* FROM source_rows s JOIN targets t ON t.text=s.text;
            """)
            baseline_parts = {}
            for path in baselines:
                if progress:
                    progress(f"读取基准中文：{path.name}")
                book = load_workbook(path, read_only=True, data_only=False, keep_links=False)
                try:
                    for sheet, header_row, text_column, output_column in baseline_sheets(book, path, report.header_rows):
                        baseline_parts.setdefault(path, {})[sheet._worksheet_path.lstrip("/")] = (sheet.title, header_row, output_column)
                        sequence = 0
                        for row in sheet.iter_rows(min_row=header_row + 1, min_col=text_column, max_col=text_column):
                            cell = row[0]
                            if cell.data_type == "f":
                                raise ValueError(f"基准中文不能是公式：{path.name} / {sheet.title} / {cell.coordinate}")
                            key = match_key(cell.value)
                            if key:
                                baseline_rows += 1
                                sequence += 1
                                connection.execute("INSERT OR IGNORE INTO targets VALUES (?)", (key,))
                                connection.execute("INSERT INTO baseline_rows VALUES (?,?,?,?,?,?,?)",
                                                   (str(path), sheet.title, cell.row, key, sequence, *boundary_newlines(cell.value)))
                finally:
                    book.close()
            if not baseline_rows:
                raise ValueError("基准文件的 Text 列没有非空中文文本")
            connection.commit()
            for index, (path, results) in enumerate(candidates.items(), 1):
                if progress:
                    progress(f"检索源文件 {index}/{len(candidates)}：{path.name}")
                book = None
                try:
                    book = load_workbook(path, read_only=True, data_only=False, keep_links=False)
                    for result in results:
                        connection.execute("SAVEPOINT sheet")
                        try:
                            sheet = book[result.sheet]
                            sheet.reset_dimensions()
                            hidden_rows = hidden_row_numbers(book, sheet) if not report.include_hidden_rows else set()
                            total = matched = 0
                            first, last = sorted((result.chinese_column, result.korean_column))
                            for number, row in enumerate(sheet.iter_rows(min_row=result.header_row + 1,
                                                                        min_col=first, max_col=last), result.header_row + 1):
                                chinese = row[result.chinese_column - first]
                                korean = row[result.korean_column - first]
                                if number in hidden_rows:
                                    continue
                                if chinese.data_type == "f" or korean.data_type == "f":
                                    raise ValueError(f"第 {number} 行中韩文本含公式，请先转成文本值")
                                key = match_key(chinese.value)
                                if not key:
                                    continue
                                total += 1
                                translation = None if korean.value is None else str(korean.value)
                                connection.execute("INSERT INTO source_rows VALUES (?,?,?,?,?,?,?,?)",
                                                   (key, chinese.value, translation, str(path), sheet.title, number, match_key(translation), total))
                                if connection.execute("SELECT 1 FROM targets WHERE text=?", (key,)).fetchone():
                                    matched += 1
                            connection.execute("RELEASE sheet")
                            sheet_rows.append([path.name, str(path), result.sheet, total, matched, total - matched,
                                               matched / total if total else None])
                        except sqlite3.Error:
                            raise
                        except Exception as error:
                            connection.execute("ROLLBACK TO sheet")
                            connection.execute("RELEASE sheet")
                            result.issue = f"正文读取失败：{type(error).__name__}: {error}"
                            skipped.append([path.name, str(path), result.sheet, result.issue])
                    connection.commit()
                except sqlite3.Error:
                    raise
                except Exception as error:
                    for result in results:
                        result.issue = f"文件读取失败：{type(error).__name__}: {error}"
                        skipped.append([path.name, str(path), result.sheet, result.issue])
                finally:
                    if book is not None:
                        book.close()
            if not sheet_rows:
                failures = [f"{r.path.name} / {r.sheet}：{r.issue}" for results in candidates.values() for r in results if r.issue]
                raise ValueError("所有候选 Sheet 正文读取失败，未生成匹配结果：\n" + "\n".join(failures[:10]))
            connection.executescript("""
                CREATE TABLE counts AS SELECT text, COUNT(*) AS n,SUM(korean_key='') AS missing,
                    COUNT(DISTINCT CASE WHEN korean_key!='' THEN korean_key END) AS variants FROM hits GROUP BY text;
                CREATE UNIQUE INDEX count_text ON counts(text);
                CREATE TABLE conflicts AS SELECT text,COUNT(DISTINCT korean_key) AS variants FROM source_rows
                    WHERE korean_key!='' GROUP BY text HAVING COUNT(DISTINCT korean_key)>1;
                CREATE UNIQUE INDEX conflict_text ON conflicts(text);
            """)
            missing_translations = connection.execute("SELECT COUNT(*) FROM source_rows WHERE korean_key='' ").fetchone()[0]
            conflict_texts = connection.execute("SELECT COUNT(*) FROM conflicts").fetchone()[0]
            if progress:
                progress("比较前后三句并生成推荐……")
            recommended_rows = build_recommendations(connection, progress)
            outputs = []
            for path in baselines:
                if progress:
                    progress(f"写入基准副本：{path.name}")
                relative = path.relative_to(baseline_root)
                destination = staging / "基准匹配结果" / relative
                destination.parent.mkdir(parents=True, exist_ok=True)
                matched_rows += write_baseline_copy(path, destination, connection, baseline_parts[path])
                outputs.append(Path("基准匹配结果") / relative)
            if progress:
                progress("写出匹配、未匹配、缺译及冲突明细……")
            match_pairs = write_matching_details(staging / "匹配明细.xlsx", connection, progress)
            outputs.append(Path("匹配明细.xlsx"))
            write_baseline_statistics(staging / "基准文件匹配统计.xlsx", connection, baseline_parts)
            outputs.append(Path("基准文件匹配统计.xlsx"))
            for row in sheet_rows:
                row.extend(issue_counts(connection, row[1], row[2]))
            file_rows = []
            for filename in sorted({r[1] for r in sheet_rows} | {r[1] for r in skipped}):
                rows = [r for r in sheet_rows if r[1] == filename]
                skip_count = sum(r[1] == filename for r in skipped)
                total = sum(r[3] for r in rows)
                matched = sum(r[4] for r in rows)
                file_rows.append([Path(filename).name, filename, len(rows), skip_count, total, matched, total - matched,
                                  matched / total if total else None, "部分跳过" if rows and skip_count else "已扫描" if rows else "全部跳过",
                                  *issue_counts(connection, filename)])
            write_statistics(staging / "源文件匹配统计.xlsx", file_rows, sheet_rows, skipped)
            outputs.append(Path("源文件匹配统计.xlsx"))
        finally:
            connection.close()
        (staging / "index.sqlite3").unlink()
        destination = output_root / ("匹配结果_" + datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + staging.name[-6:])
        os.rename(staging, destination)
    return MatchReport(destination, fresh, baseline_rows, matched_rows, sum(r[3] for r in sheet_rows),
                       sum(r[4] for r in sheet_rows), len(skipped), [destination / p for p in outputs],
                       missing_translations, conflict_texts, match_pairs, recommended_rows, len(baselines))
