"""Stream a second baseline copy containing only translation-column changes."""

import re
import shutil
from pathlib import Path
from zipfile import ZipFile
from xml.sax.expatreader import create_parser
from xml.sax.handler import feature_external_ges
from xml.sax.saxutils import XMLGenerator

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


class BackfillWriter(BaselineWriter):
    """Reuse literal-safe XML emission, replacing only the intended cell."""

    def __init__(self, stream, db, path, sheet, column, header_row, new_header):
        super().__init__(stream, db, path, sheet, header_row, column)
        self.target = column
        self.new_header = new_header
        self.suppressed = 0
        self.value = None
        self.written = False
        self.checking = False

    def emit(self, attrs=None):
        attrs = dict(attrs or {})
        attrs.update(r=f"{get_column_letter(self.target)}{self.row}", t="inlineStr")
        XMLGenerator.startElement(self, self.prefix + "c", attrs)
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
        for item in source.infolist():
            with source.open(item) as original, output.open(item, "w") as copied:
                if item.filename in parts:
                    sheet = parts[item.filename][0]
                    column, header, new_header = db.execute("SELECT column_number,header_row,new_header FROM fill_sheets WHERE path=? AND sheet=?", (str(path), sheet)).fetchone()
                    writer = BackfillWriter(copied, db, path, sheet, column, header, new_header)
                    parser = create_parser()
                    parser.setFeature(feature_external_ges, False)
                    parser.setContentHandler(writer)
                    parser.parse(original)
                else:
                    shutil.copyfileobj(original, copied, length=1024 * 1024)


def write_log(destination, db, baseline_root, progress=None):
    book = Workbook(write_only=True)
    rows = ((*r, str(Path(r[0]).relative_to(baseline_root))) for r in db.execute("SELECT * FROM fill_audit ORDER BY path,sheet,row"))
    write_detail_sheet(book, "回填明细", ["基准完整路径", "Sheet", "行号", "键", "回填前", "回填后", "推荐状态", "回填处理", "推荐依据", "推荐来源", "基准相对路径"],
                       rows, [65, 25, 12, 45, 60, 60, 26, 45, 90, 80, 60], progress)
    write_detail_sheet(book, "说明", ["项目", "规则"], [
        ["范围", "纯回填结果与基准匹配结果保留相同目录、文件名、Sheet、行号；纯回填不追加匹配分析列。"],
        ["写入", "成功推荐（含待复核推荐）按推荐列相同的首尾换行规则写入 TextTrans／translate；未推荐、缺译保留原值。"],
        ["保护", "原件不修改；已有译文公式保留并记录。缺少 TextTrans 时在表头末尾新增，若该列已有未标名数据则停止生成，防止覆盖。"],
        ["交叉核对", "先按键及文件/Sheet定位；双方均明确回填或确认且原始译文完全一致，可列为双方一致。双方都未回填、仅沿用原值，不能据此定稿。"],
        ["格式差异", "与客户工具比对时，须区分文字差异、标签、字面\\n与实际换行、|变体及自动字幕折行。此输出不擅自复制客户工具的格式改写。"],
    ], [24, 110])
    book.save(destination)
    book.close()
