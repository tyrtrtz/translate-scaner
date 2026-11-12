"""Stream a second baseline copy containing only translation-column changes."""

import re
import shutil
from pathlib import Path
from zipfile import ZipFile
from xml.etree.ElementTree import fromstring, iterparse
from xml.sax.expatreader import create_parser
from xml.sax.handler import feature_external_ges
from xml.sax.saxutils import XMLGenerator, escape

from openpyxl import Workbook
from openpyxl.utils import get_column_letter, column_index_from_string

from matcher import BaselineWriter, align_boundary_newlines, write_detail_sheet


def prepare_values(db):
    db.execute("CREATE TABLE fill_values (path TEXT,sheet TEXT,row INTEGER,value TEXT,PRIMARY KEY(path,sheet,row))")
    db.execute("CREATE TABLE fill_audit (path TEXT,sheet TEXT,row INTEGER,key TEXT,old TEXT,new TEXT,status TEXT,action TEXT,basis TEXT,source TEXT)")
    for path, sheet, row, key, old, formula, korean, start, end, status, basis, source in db.execute("""
        SELECT b.path,b.sheet,b.row,f.key,f.original,f.is_formula,COALESCE(t.korean,s.korean),
            b.leading_newline,b.trailing_newline,COALESCE(t.status,r.status),COALESCE(t.basis,r.basis),
            COALESCE(t.source,s.path || '；Sheet：' || s.sheet || '；行号：' || s.row,'')
        FROM baseline_rows b JOIN fill_rows f ON f.path=b.path AND f.sheet=b.sheet AND f.row=b.row
        JOIN recommendations r ON r.path=b.path AND r.sheet=b.sheet AND r.row=b.row
        LEFT JOIN source_rows s ON s.rowid=r.source_id
        LEFT JOIN structure_recommendations t ON t.path=b.path AND t.sheet=b.sheet AND t.row=b.row
        ORDER BY b.path,b.sheet,b.row
    """):
        value = align_boundary_newlines(korean, start, end)
        if not value or not value.strip():
            action, value = "未推荐，保留原译文", old
        elif formula:
            action, value = "译文列含公式，保留原公式待人工处理", old
        else:
            action = "已推荐，译文与原值一致" if value == old else "已回填"
            db.execute("INSERT INTO fill_values VALUES (?,?,?,?)", (path, sheet, row, value))
        db.execute("INSERT INTO fill_audit VALUES (?,?,?,?,?,?,?,?,?,?)", (path, sheet, row, key, old, value, status, action, basis, source))
    db.commit()


def prepare_strings(source, part, db, path):
    """Index only pending fill text on disk; keep existing SST indices stable."""
    db.execute("CREATE TEMP TABLE IF NOT EXISTS fill_strings (value TEXT PRIMARY KEY,string_index INTEGER)")
    db.execute("DELETE FROM fill_strings")
    db.execute("INSERT OR IGNORE INTO fill_strings(value) SELECT value FROM fill_values WHERE path=?", (str(path),))
    db.execute("INSERT OR IGNORE INTO fill_strings(value) SELECT new_header FROM fill_sheets WHERE path=? AND new_header!=''", (str(path),))
    size = 0
    with source.open(part) as stream:
        elements = iterparse(stream, events=("start", "end"))
        _, root = next(elements)
        namespace = root.tag[:-3]  # expanded name of sst, including the closing brace
        for event, element in elements:
            if event == "end" and element.tag == namespace + "si":
                # Do not reuse a rich-text entry: its run formatting belongs to
                # the old cells. Leave it and every old index intact.
                if len(element) == 1 and element[0].tag == namespace + "t":
                    value = element[0].text or ""
                    if not re.search(r"_x[0-9A-Fa-f]{4}_", value):
                        db.execute("UPDATE fill_strings SET string_index=? WHERE value=? AND string_index IS NULL", (size, value))
                size += 1
                root.remove(element)
    for index, (value,) in enumerate(db.execute("SELECT value FROM fill_strings WHERE string_index IS NULL ORDER BY value"), size):
        db.execute("UPDATE fill_strings SET string_index=? WHERE value=?", (index, value))
    return size


class SharedStringsWriter(XMLGenerator):
    """Stream the original SST, appending new plain strings before extLst."""

    def __init__(self, stream, db, original_size, count_delta):
        super().__init__(stream, encoding="utf-8")
        self.db, self.original_size, self.count_delta = db, original_size, count_delta
        self.prefix = ""
        self.appended = False

    def startElement(self, name, attrs):
        local = name.rsplit(":", 1)[-1]
        if local == "sst":
            self.prefix = name[:-3]
            attrs = dict(attrs)
            if "count" in attrs and "uniqueCount" in attrs:
                added = self.db.execute("SELECT COUNT(*) FROM fill_strings WHERE string_index>=?", (self.original_size,)).fetchone()[0]
                attrs.update(count=str(int(attrs["count"]) + self.count_delta), uniqueCount=str(self.original_size + added))
            else:
                # Both optional counts must be present together.
                attrs.pop("count", None)
                attrs.pop("uniqueCount", None)
        elif name == self.prefix + "extLst":
            self.append_strings()
        super().startElement(name, attrs)

    def characters(self, content):
        if "\r" in content:
            self._finish_pending_start_element()
            self._write(escape(content, {"\r": "&#13;"}))
        else:
            super().characters(content)

    def append_strings(self):
        if self.appended:
            return
        self.appended = True
        for (value,) in self.db.execute("SELECT value FROM fill_strings WHERE string_index>=? ORDER BY string_index", (self.original_size,)):
            super().startElement(self.prefix + "si", {})
            super().startElement(self.prefix + "t", {"xml:space": "preserve"})
            # Escape literal OOXML control tokens only in newly appended text.
            # Existing SST entries already have their original encoding.
            self.characters(re.sub(r"_(x[0-9A-Fa-f]{4}_)", r"_x005F_\1", value))
            super().endElement(self.prefix + "t")
            super().endElement(self.prefix + "si")

    def endElement(self, name):
        if name == self.prefix + "sst":
            self.append_strings()
        super().endElement(name)


class BackfillWriter(BaselineWriter):
    """Reuse literal-safe XML emission, replacing only the intended cell."""

    def __init__(self, stream, db, path, sheet, column, header_row, new_header, shared_strings=False):
        super().__init__(stream, db, path, sheet, header_row, column)
        self.target = column
        self.new_header = new_header
        self.suppressed = 0
        self.value = None
        self.written = False
        self.checking = False
        self.shared_strings = shared_strings
        self.shared_count_delta = 0

    def emit(self, attrs=None):
        attrs = dict(attrs or {})
        self.shared_count_delta += int(self.shared_strings and attrs.get("t") != "s")
        attrs.update(r=f"{get_column_letter(self.target)}{self.row}", t="s" if self.shared_strings else "inlineStr")
        XMLGenerator.startElement(self, self.prefix + "c", attrs)
        if self.shared_strings:
            index = self.connection.execute("SELECT string_index FROM fill_strings WHERE value=?", (self.value,)).fetchone()[0]
            XMLGenerator.startElement(self, self.prefix + "v", {})
            self.xml_characters(str(index))
            XMLGenerator.endElement(self, self.prefix + "v")
        else:
            XMLGenerator.startElement(self, self.prefix + "is", {})
            XMLGenerator.startElement(self, self.prefix + "t", {"xml:space": "preserve"})
            self.xml_characters(self.value)
            XMLGenerator.endElement(self, self.prefix + "t")
            XMLGenerator.endElement(self, self.prefix + "is")
        XMLGenerator.endElement(self, self.prefix + "c")
        self.written = True

    def startElement(self, name, attrs):
        local = name.rsplit(":", 1)[-1]
        if self.checking and local == "f":
            raise ValueError(f"{Path(self.path).name} / {self.sheet} 新增译文列已有公式，请补齐 TextTrans 表头")
        if self.suppressed:
            self.suppressed += 1
            if self.new_header and local == "f":
                raise ValueError(f"{Path(self.path).name} / {self.sheet} 新增译文列已有公式，未生成回填文件")
            return
        if local == "worksheet":
            self.prefix = name[:-len(local)]
        if local == "dimension":
            self.omit_dimension = True
            return
        if local == "row":
            self.row, self.value, self.written = int(attrs["r"]), None, False
            if self.row == self.header_row and self.new_header:
                self.value = self.new_header
            else:
                found = self.connection.execute("SELECT value FROM fill_values WHERE path=? AND sheet=? AND row=?", (self.path, self.sheet, self.row)).fetchone()
                if found:
                    self.value = found[0]
        if local == "c":
            column = column_index_from_string(re.match(r"[A-Z]+", attrs["r"])[0])
            if self.value is not None:
                if column == self.target:
                    self.emit(attrs)
                    self.suppressed = 1
                    return
                if column > self.target and not self.written:
                    self.emit()
            # An absent header must never authorize overwriting unlabelled data.
            if column == self.target and self.new_header:
                self.checking = True
        XMLGenerator.startElement(self, name, attrs)

    def characters(self, content):
        if self.checking and content.strip():
            raise ValueError(f"{Path(self.path).name} / {self.sheet} 新增译文列已有内容，请补齐 TextTrans 表头")
        if self.suppressed:
            if self.new_header and content.strip():
                raise ValueError(f"{Path(self.path).name} / {self.sheet} 新增译文列已有内容，请补齐 TextTrans 表头")
        elif not self.omit_dimension:
            self.xml_characters(content)

    def endElement(self, name):
        if self.suppressed:
            self.suppressed -= 1
            return
        local = name.rsplit(":", 1)[-1]
        if local == "c":
            self.checking = False
        if local == "dimension":
            self.omit_dimension = False
            return
        if local == "row":
            if self.value is not None and not self.written:
                self.emit()
            self.row = None
        XMLGenerator.endElement(self, name)


def write_copy(path, destination, db, parts):
    destination.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(path) as source, ZipFile(destination, "w") as output:
        shared_part = next((e.attrib["PartName"].lstrip("/") for e in fromstring(source.read("[Content_Types].xml"))
                            if e.attrib.get("ContentType", "").endswith(".sharedStrings+xml")), None)
        original_size = prepare_strings(source, shared_part, db, path) if shared_part else 0
        count_delta = 0
        for item in source.infolist():
            if item.filename == shared_part:
                continue
            with source.open(item) as original, output.open(item, "w") as copied:
                if item.filename in parts:
                    sheet = parts[item.filename][0]
                    column, header, new_header = db.execute("SELECT column_number,header_row,new_header FROM fill_sheets WHERE path=? AND sheet=?", (str(path), sheet)).fetchone()
                    writer = BackfillWriter(copied, db, path, sheet, column, header, new_header, bool(shared_part))
                    parser = create_parser()
                    parser.setFeature(feature_external_ges, False)
                    parser.setContentHandler(writer)
                    parser.parse(original)
                    count_delta += writer.shared_count_delta
                else:
                    shutil.copyfileobj(original, copied, length=1024 * 1024)
        if shared_part:
            with source.open(shared_part) as original, output.open(source.getinfo(shared_part), "w") as copied:
                writer = SharedStringsWriter(copied, db, original_size, count_delta)
                parser = create_parser()
                parser.setFeature(feature_external_ges, False)
                parser.setContentHandler(writer)
                parser.parse(original)


def write_log(destination, db, baseline_root, progress=None):
    book = Workbook(write_only=True)
    rows = ((*r, str(Path(r[0]).relative_to(baseline_root))) for r in db.execute("SELECT * FROM fill_audit ORDER BY path,sheet,row"))
    write_detail_sheet(book, "回填明细", ["基准完整路径", "Sheet", "行号", "键", "回填前", "回填后", "推荐状态", "回填处理", "推荐依据", "推荐来源", "基准相对路径"],
                       rows, [65, 25, 12, 45, 60, 60, 26, 45, 90, 80, 60], progress)
    write_detail_sheet(book, "说明", ["项目", "规则"], [
        ["范围", "纯回填结果与基准匹配结果保留相同目录、文件名、Sheet、行号；纯回填不追加匹配分析列。"],
        ["写入", "成功推荐（含待复核推荐）按推荐列相同的首尾换行规则写入 TextTrans／translate；未推荐、缺译保留原值。"],
        ["保护", "原件不修改；已有译文公式保留并记录。缺少 TextTrans 时在表头末尾新增，若该列已有未标名数据则停止生成，防止覆盖。"],
        ["字符串格式", "原基准有 sharedStrings 时，回填译文及新增表头也使用共享索引，保留原索引及富文本；原本为 inlineStr 的基准保持原方式。"],
        ["交叉核对", "先按键及文件/Sheet定位；双方均明确回填或确认且原始译文完全一致，可列为双方一致。双方都未回填、仅沿用原值，不能据此定稿。"],
        ["格式差异", "与客户工具比对时，须区分文字差异、标签、字面\\n与实际换行、|变体及自动字幕折行。此输出不擅自复制客户工具的格式改写。"],
    ], [24, 110])
    book.save(destination)
    book.close()
