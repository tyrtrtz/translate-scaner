"""Compare two local backfills, retaining exact strings and evidence of changes."""

import os
import sqlite3
import tempfile
import unicodedata
from collections import Counter
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

from key_matching import base, business_key, content_key, metadata_columns, records, value_at
from matcher import write_detail_sheet
from scanner import SUPPORTED, normalize


def identity(value):
    return unicodedata.normalize("NFC", str(value)).replace("\\", "/").casefold()


def index_files(db, side, root, header_rows, progress):
    paths = [root] if root.is_file() else sorted(p for p in root.rglob("*") if p.is_file() and p.suffix.lower() in SUPPORTED and not p.name.startswith("~$"))
    for path in paths:
        if path.suffix.lower() not in SUPPORTED:
            raise ValueError(f"不支持的对比文件：{path.name}")
        if progress:
            progress(f"读取{'我方' if side == 0 else '客户'}回填文件：{path.name}")
        rel = path.name if root.is_file() else str(path.relative_to(root))
        book = load_workbook(path, read_only=True, data_only=False)
        try:
            for sheet in book:
                sheet.reset_dimensions()
                headers = []
                for number, row in enumerate(sheet.iter_rows(max_row=header_rows), 1):
                    titles = {c.column: normalize(c.value) for c in row if c.value is not None}
                    text = [c for c, title in titles.items() if title == "Text" or (title == "chs" and "translationKey" in titles.values())]
                    if text:
                        columns = metadata_columns(row)
                        if side == 1:
                            results = [c for c, title in titles.items() if title.lower() == "html回填结果"]
                            if results:
                                # This named output takes priority over a retained original translation column.
                                columns["target"] = results[0] if len(results) == 1 else None
                        headers.append((number, text, columns))
                if not headers:
                    db.execute("INSERT INTO skipped VALUES (?,?,?,?)", (side, rel, sheet.title, f"前{header_rows}行无 Text 或 translationKey/chs 表头"))
                    continue
                if len(headers) != 1 or len(headers[0][1]) != 1 or not headers[0][2]["target"]:
                    raise ValueError(f"{path.name} / {sheet.title} 中文或译文表头缺失／不唯一，请先整理表头")
                number, text, columns = headers[0]
                for row in sheet.iter_rows(min_row=number + 1, max_col=max(text[0], columns["target"], columns["key"] or 0, columns["comment"] or 0)):
                    cn = row[text[0] - 1]
                    if cn.value is None or not str(cn.value).strip():
                        continue
                    ko = row[columns["target"] - 1]
                    db.execute("INSERT INTO entries VALUES (NULL,?,?,?,?,?,?,?,?,?,?,?)", (side, identity(rel), sheet.title, cn.row,
                        business_key(row, columns), str(cn.value), str(ko.value) if ko.value is not None else "",
                        int(cn.data_type == "f" or ko.data_type == "f"), rel, str(path), value_at(row, columns["key"]).strip()))
        finally:
            book.close()
    if not db.execute("SELECT 1 FROM entries WHERE side=? LIMIT 1", (side,)).fetchone():
        raise ValueError(f"{'我方' if side == 0 else '客户'}文件中未读到可对比的中文行，请检查所选目录和表头检查行数")


def difference(a, b):
    if base(a).strip() == base(b).strip():
        return "换行表示、首尾空白或 Unicode 写法不同；仍不视为完全一致"
    if content_key(a) == content_key(b):
        return "去标签、空白或标点后相同；需复核格式与标签"
    if "|" in a or "|" in b:
        return "译文内容或 | 变体不同，需人工复核"
    return "译文文字不同，结合双方来源和推荐依据复核"


def compare_backfills(ours, customer, *, header_rows=10, progress=None):
    ours, customer = Path(ours).expanduser().resolve(), Path(customer).expanduser().resolve()
    if ours.name == "纯回填结果":
        ours = ours.parent
    pure, log = ours / "纯回填结果", ours / "回填日志.xlsx"
    if not pure.is_dir() or not log.is_file():
        raise ValueError("请选择含“纯回填结果”文件夹及“回填日志.xlsx”的本次匹配结果目录")
    if not customer.exists() or customer == ours or customer == pure or customer in pure.parents or pure in customer.parents:
        raise ValueError("请选择独立的客户回填文件或文件夹，不能用我方结果和自身对比")
    if header_rows < 1:
        raise ValueError("表头检查行数最低为 1")
    with tempfile.TemporaryDirectory(prefix=".回填对比_", dir=ours) as temporary:
        staging = Path(temporary)
        db = sqlite3.connect(staging / "index.sqlite3")
        try:
            db.executescript("""
                CREATE TABLE entries (id INTEGER PRIMARY KEY,side INTEGER,rel TEXT,sheet TEXT,row INTEGER,key TEXT,text TEXT,ko TEXT,formula INTEGER,display TEXT,path TEXT,legacy_key TEXT);
                CREATE INDEX entry_key ON entries(side,key);
                CREATE INDEX entry_position ON entries(side,rel,sheet,row);
                CREATE TABLE audit (rel TEXT,sheet TEXT,row INTEGER,key TEXT,old TEXT,new TEXT,status TEXT,action TEXT,basis TEXT,source TEXT,PRIMARY KEY(rel,sheet,row));
                CREATE TABLE used (id INTEGER PRIMARY KEY);
                CREATE TABLE skipped (side INTEGER,path TEXT,sheet TEXT,reason TEXT);
                CREATE TABLE results (status TEXT,reason TEXT,ours TEXT,sheet TEXT,row INTEGER,key TEXT,text TEXT,old TEXT,ko TEXT,customer TEXT,customer_sheet TEXT,customer_row INTEGER,customer_text TEXT,customer_ko TEXT,recommendation TEXT,basis TEXT,source TEXT);
            """)
            book = load_workbook(log, read_only=True, data_only=False)
            try:
                sheets = [s for s in book if s.title == "回填明细" or s.title.startswith("回填明细_")]
                if not sheets:
                    raise ValueError("回填日志缺少“回填明细”")
                for sheet in sheets:
                    rows = sheet.iter_rows(values_only=True)
                    header = next(rows, ())
                    if len(header) < 11 or header[10] != "基准相对路径":
                        raise ValueError("回填日志格式不完整，请用新版重新生成匹配结果")
                    for row in rows:
                        r = list(row) + [None] * max(0, 11 - len(row))
                        db.execute("INSERT INTO audit VALUES (?,?,?,?,?,?,?,?,?,?)", (identity(r[10]), r[1], r[2], r[3] or "", r[4] or "", r[5] or "", r[6], r[7], r[8], r[9]))
            finally:
                book.close()
            index_files(db, 0, pure, header_rows, progress)
            index_files(db, 1, customer, header_rows, progress)
            counts = Counter()
            for a in records(db, "SELECT * FROM entries WHERE side=0 ORDER BY rel,sheet,row"):
                audit = next(records(db, "SELECT * FROM audit WHERE rel=? AND sheet=? AND row=?", (a["rel"], a["sheet"], a["row"])), None)
                if a["key"]:
                    candidates = list(records(db, "SELECT * FROM entries WHERE side=1 AND key=?", (a["key"],)))
                    local = [b for b in candidates if (b["rel"], b["sheet"]) == (a["rel"], a["sheet"])]
                    if local:
                        candidates = local
                else:
                    candidates = list(records(db, "SELECT * FROM entries WHERE side=1 AND rel=? AND sheet=? AND row=?", (a["rel"], a["sheet"], a["row"])))
                b = candidates[0] if len(candidates) == 1 else None
                status, reason = "无法对应", "客户结果中未找到相同键或无键行的相同文件/Sheet/行号" if not candidates else "客户同键候选不唯一，未按译文挑选位置"
                if b:
                    db.execute("INSERT OR IGNORE INTO used VALUES (?)", (b["id"],))
                    if a["text"] != b["text"]:
                        reason = "定位到的原文不同，可能使用了不同基准版本，不能据此定稿"
                    elif a["formula"] or b["formula"]:
                        status, reason = "待确认", "原文或译文含公式，未用公式文本代替计算结果"
                    elif a["ko"] != b["ko"]:
                        status, reason = "译文不一致", difference(a["ko"], b["ko"])
                    elif not audit or audit["new"] != a["ko"] or audit["key"] not in (a["key"], a["legacy_key"]):
                        status, reason = "待确认", "译文一致，但我方文件与回填日志不一致或缺少记录，不能证明本次推荐"
                    elif not a["ko"].strip():
                        status, reason = "待确认", "双方译文均为空，不视为定稿"
                    elif audit["action"] not in ("已回填", "已推荐，译文与原值一致"):
                        status, reason = "待确认", "双方译文一致，但我方未实际推荐，沿用原译文"
                    elif b["ko"] == audit["old"]:
                        status, reason = "待确认", "双方译文一致且等于原值，客户是否实际回填／确认无法仅凭 Excel 证明"
                    else:
                        status, reason = "一致可定稿", "双方原文、译文逐字符一致；我方有推荐记录，客户译文相对同一基准原值已改变"
                counts[status] += 1
                db.execute("INSERT INTO results VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (status, reason, a["display"], a["sheet"], a["row"], a["key"], a["text"], audit["old"] if audit else None, a["ko"],
                    b["path"] if b else "", b["sheet"] if b else "", b["row"] if b else None, b["text"] if b else "", b["ko"] if b else "", audit["status"] if audit else "", audit["basis"] if audit else "", audit["source"] if audit else ""))
            book = Workbook(write_only=True)
            headers = ["对比状态", "判定依据", "我方相对文件", "我方Sheet", "我方行号", "键", "我方原文", "基准原译文", "我方回填译文", "客户文件", "客户Sheet", "客户行号", "客户原文", "客户回填译文", "我方推荐状态", "我方推荐依据", "我方来源"]
            widths = [24, 85, 55, 22, 12, 45, 60, 60, 60, 65, 22, 12, 60, 60, 28, 90, 80]
            for status in ("一致可定稿", "译文不一致", "无法对应", "待确认"):
                write_detail_sheet(book, status, headers, db.execute("SELECT * FROM results WHERE status=? ORDER BY ours,sheet,row", (status,)), widths, progress)
            write_detail_sheet(book, "客户未对应", ["客户文件", "Sheet", "行号", "键", "原文", "译文"], db.execute("SELECT path,sheet,row,key,text,ko FROM entries WHERE side=1 AND id NOT IN (SELECT id FROM used) ORDER BY rel,sheet,row"), [65, 22, 12, 45, 60, 60], progress)
            write_detail_sheet(book, "跳过工作表", ["我方0_客户1", "文件", "Sheet", "原因"], db.execute("SELECT * FROM skipped"), [18, 65, 22, 80])
            write_detail_sheet(book, "说明与统计", ["项目", "说明／数量"], [
                ["我方结果", str(ours)], ["客户结果", str(customer)], *[[k, counts[k]] for k in ("一致可定稿", "译文不一致", "无法对应", "待确认")],
                ["使用前提", "客户工具必须从同一批、同一版本的原始基准生成结果；不要拿我方回填结果作为客户工具输入。"],
                ["定位", "有键按键定位，优先同相对文件/Sheet；跨文件同键须唯一。无键按相对文件/Sheet/行号对应。原文必须逐字符相同。"],
                ["一致", "译文逐字符比较，不忽略标签、空白、换行或标点；双方空白或仅沿用原值不自动定稿。此处定稿是双方结果一致，不代表已验证译文质量。"],
                ["用途", "只生成对比清单，不改两边工作簿；不一致时结合推荐依据与来源审核，客户来源需在客户报告中查阅。"],
            ], [28, 110])
            book.save(staging / "回填对比.xlsx")
            book.close()
        finally:
            db.close()
        (staging / "index.sqlite3").unlink()
        destination = ours / ("回填对比_" + datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + staging.name[-6:] + ".xlsx")
        os.rename(staging / "回填对比.xlsx", destination)
    return destination, dict(counts)
