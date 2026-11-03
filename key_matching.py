"""Last-pass script-to-key matching. Never changes an existing recommendation.

The translation table stays in SQLite. Only one source sheet and bounded
anchor windows are materialized; Korean is never used to locate a key.
"""

import re
import sqlite3
import unicodedata
from collections import Counter
from pathlib import Path


STORY_KEY = re.compile(r"^(story_(?:text|convenient))_(\d+)_Dialog_Detail")
OPTION_KEY = re.compile(r"^(story_(?:text|convenient)_\d+)_Option(\d+)_")
VOICE_KEY = re.compile(r"^char_voice_(?:words_[a-z]+_)?(\d{6})")
COMMENT_KEY = re.compile(r"(?:story_(?:(?:text|convenient)_)?\d+|char_voice_(?:words_[a-z]+_)?\d+)(?:_[A-Za-z0-9]+)+")
OPTION = re.compile(r"(?:^|\s)([A-Z]|\d{1,2})(?:[.．](?!\d)|[、:：;；,，])\s*")
PLAYER = {"玩家", "主角", "主人公", "玩家(男)", "玩家(女)", "{playername}", "유저", "플레이어", "주인공", "男主角", "女主角", "星临者"}
GENERIC_SPEAKERS = {"旁白", "系统", "系统说明", "系统文本", "系统提示", "群杂", "众人", "路人", "npc"}
EDGE_MATCH = "客户规则：单侧顺序补查"


def base(value):
    return unicodedata.normalize("NFC", str(value or "").replace("_x000D_", "").replace("\\n", "\n").replace("\r", ""))


def content_key(value, *, parentheses=False):
    from matcher import breath_key, plain_key
    text = plain_key(base(value), punctuation=False)
    if parentheses:
        # Keep malformed brackets literal rather than silently deleting content.
        text = breath_key(text)
    return "".join(c for c in text.replace("诶", "欸").lower() if c.isalnum())


def similarity(a, b):
    """Full-length LCS/Indel similarity, with linear memory and no 400-char cut."""
    if not a or not b:
        return float(a == b)
    # Bitset LCS avoids a quadratic Python matrix; the integer is O(len(a)).
    masks = {}
    for i, c in enumerate(a):
        masks[c] = masks.get(c, 0) | (1 << i)
    state = 0
    for c in b:
        x = masks.get(c, 0) | state
        state = x & ~(x - ((state << 1) | 1))
    return 2 * state.bit_count() / (len(a) + len(b))


def option_parts(value):
    text = base(value)
    marks = list(OPTION.finditer(text))
    if len(marks) < 2 or text[:marks[0].start()].strip():
        return []
    labels = [m[1] for m in marks]
    expected = [str(i + 1) for i in range(len(marks))] if labels[0].isdigit() else [chr(65 + i) for i in range(len(marks))]
    if labels != expected:
        return []
    parts = [text[m.end():marks[i + 1].start() if i + 1 < len(marks) else len(text)].strip() for i, m in enumerate(marks)]
    return parts if all(parts) else []


def speaker_variants(value):
    text = re.sub(r"\s+", "", unicodedata.normalize("NFKC", base(value))).lower()
    if not text or text in GENERIC_SPEAKERS:
        return set()
    if text in PLAYER:
        return {"@player"}
    values = {text, re.sub(r"\([^)]*\)", "", text), *re.findall(r"\(([^)]*)\)", text)}
    return {s for s in values if s and not re.fullmatch(r"[?？…]+", s)}


def speaker_compare(a, b):
    a, b = speaker_variants(a), speaker_variants(b)
    return 0 if not a or not b else 1 if a & b else -1


def filename_speaker(path):
    name = unicodedata.normalize("NFC", Path(path).stem)
    if not re.search(r"角色文档|캐릭터", name):
        return ""
    match = re.search(r"_P\d+_+([^_]+)", name) or re.search(r"([^_]+)_+角色文档", name)
    speaker = re.sub(r"[（(][^）)]*[）)]", "", match[1]).strip() if match else ""
    if speaker.lower() in PLAYER:
        gender = re.search(r"主人公_+(男|女)|주인공_+(남|여)", name)
        if gender:
            return "男主角" if next(v for v in gender.groups() if v) in ("男", "남") else "女主角"
    return speaker


def metadata_columns(header):
    def find(*names):
        names = {unicodedata.normalize("NFKC", n).lower().replace(" ", "") for n in names}
        cols = [c.column for c in header if isinstance(c.value, str) and
                unicodedata.normalize("NFKC", c.value).strip().lower().replace(" ", "") in names]
        # Ambiguous optional metadata is never guessed.
        return cols[0] if len(cols) == 1 else None
    return dict(key=find("translationKey", "Key", "TextKey", "ID"), comment=find("Comment"),
                speaker=find("说话角色名", "说话角色", "角色", "角色名", "speaker", "CharacterName", "Character"),
                dialogue=find("对话ID", "台词ID"), plot=find("剧情ID"),
                target=find("TextTrans", "translate"), note=find("特殊说明"))


def value_at(row, column):
    if not column or column > len(row):
        return ""
    cell = row[column - 1]
    return "" if cell.data_type == "f" or cell.value is None else str(cell.value)


def business_key(row, columns):
    """Some exports use Key for a numeric record ID and Comment for the game key."""
    key = value_at(row, columns["key"]).strip()
    comment = value_at(row, columns["comment"]).strip()
    if (not key or re.fullmatch(r"\d+(?:\.0+)?", key)) and COMMENT_KEY.fullmatch(comment):
        return comment
    return key


def numeric_id(value):
    # Unusual identifiers remain text keys, never overflow SQLite integers.
    return int(value) if value.isdigit() and len(value) <= 18 else None


def init_index(db):
    db.executescript("""
        CREATE TABLE key_baselines (id INTEGER PRIMARY KEY,path TEXT,sheet TEXT,row INTEGER,
            key TEXT,speaker TEXT,text TEXT,loose TEXT,family TEXT,number INTEGER,option_base TEXT,option_number INTEGER,
            UNIQUE(path,sheet,row));
        CREATE INDEX key_baseline_key ON key_baselines(key);
        CREATE INDEX key_baseline_text ON key_baselines(loose);
        CREATE INDEX key_baseline_number ON key_baselines(family,number);
        CREATE INDEX key_baseline_position ON key_baselines(path,sheet,row);
        CREATE INDEX key_baseline_options ON key_baselines(option_base,option_number);
        CREATE TABLE key_variants (baseline_id INTEGER,text TEXT,PRIMARY KEY(baseline_id,text));
        CREATE INDEX key_variant_text ON key_variants(text);
        CREATE TABLE key_sources (source_id INTEGER PRIMARY KEY,speaker TEXT,dialogue TEXT,plot TEXT,
            has_dialogue INTEGER,has_speaker INTEGER,char_voice INTEGER,excluded TEXT);
        CREATE TABLE key_audit (source_id INTEGER,baseline_id INTEGER,method TEXT,similarity REAL,outcome TEXT);
        CREATE TABLE key_proposals (baseline_id INTEGER,source_id INTEGER,korean TEXT,method TEXT,similarity REAL,
            part INTEGER,total INTEGER,PRIMARY KEY(baseline_id,source_id,part));
        CREATE TEMP TABLE key_edge_proposals AS SELECT * FROM key_proposals;
        CREATE TABLE fill_sheets (path TEXT,sheet TEXT,column_number INTEGER,header_row INTEGER,new_header TEXT,
            schema TEXT,PRIMARY KEY(path,sheet));
        CREATE TABLE fill_rows (path TEXT,sheet TEXT,row INTEGER,key TEXT,original TEXT,is_formula INTEGER,
            PRIMARY KEY(path,sheet,row));
    """)


def index_baseline(db, path, sheet, row, text, columns):
    key = business_key(row, columns)
    speaker = value_at(row, columns["speaker"]).strip()
    number = next((c.row for c in row if c.value is not None), None)
    story, option = STORY_KEY.match(key), OPTION_KEY.match(key)
    cur = db.execute("INSERT INTO key_baselines VALUES (NULL,?,?,?,?,?,?,?,?,?,?,?)",
                     (str(path), sheet, number, key, speaker, text, content_key(text, parentheses=True),
                      story[1] if story and numeric_id(story[2]) is not None else "", numeric_id(story[2]) if story else None,
                      option[1] if option else "", numeric_id(option[2]) if option else None))
    for part in re.split(r"\|(?!hpp\()", text):
        normalized = content_key(part, parentheses=True)
        if normalized:
            db.execute("INSERT OR IGNORE INTO key_variants VALUES (?,?)", (cur.lastrowid, normalized))
    target = columns["target"]
    cell = row[target - 1] if target and target <= len(row) else None
    db.execute("INSERT INTO fill_rows VALUES (?,?,?,?,?,?)", (str(path), sheet, number, key,
               str(cell.value) if cell is not None and cell.value is not None else "",
               int(cell is not None and cell.data_type == "f")))


def index_source(db, source_id, path, row, columns, chinese_column, korean_column, include_deprecated):
    speaker = value_at(row, columns["speaker"]).strip() if columns["speaker"] else filename_speaker(path)
    dialogue = re.sub(r"\.0+$", "", value_at(row, columns["dialogue"]).strip())
    plot = re.sub(r"\.0+$", "", value_at(row, columns["plot"]).strip())
    note = value_at(row, columns["note"]).strip()
    chinese, korean = value_at(row, chinese_column), value_at(row, korean_column)
    other = [str(c.value or "").strip() for number, c in enumerate(row, 1)
             if number not in (chinese_column, korean_column) and c.data_type != "f"]
    reason = ""
    if not include_deprecated and any(re.search(r"^(废弃|폐기)$|[（(]废弃[）)]", v) for v in other):
        reason = "客户补匹配排除：废弃台词"
    elif re.fullmatch(r"韩[语文](?:已经|已)收录", korean.strip()) or any(re.search(r"번역\s*불필요|중복\s*번역|无需翻译", v) for v in other):
        reason = "客户补匹配排除：无需翻译标记"
    elif note == "演出" or (columns["speaker"] and not speaker):
        reason = "客户补匹配排除：演出或有角色列但本行角色为空"
    elif not content_key(chinese, parentheses=True):
        reason = "客户补匹配排除：只有符号或括号说明"
    char_voice = bool(filename_speaker(path)) and not columns["dialogue"] and not columns["plot"]
    db.execute("INSERT INTO key_sources VALUES (?,?,?,?,?,?,?,?)", (source_id, speaker, dialogue, plot,
               bool(columns["dialogue"]), bool(columns["speaker"]), char_voice, reason))


def records(db, query, args=()):
    cursor = db.cursor()
    cursor.row_factory = sqlite3.Row
    return cursor.execute(query, args)


def ordered_alignment(sources, targets, compatible):
    """Monotonic LCS, preferring relative positions only inside bounded anchors."""
    m, n = len(sources), len(targets)
    scores = [[0] * (n + 1) for _ in range(m + 1)]
    def weight(i, j):
        return 1000 - round(100 * abs((i + 1) / (m + 1) - (j + 1) / (n + 1)))
    for i in range(m):
        for j in range(n):
            score = scores[i][j] + weight(i, j) if compatible(sources[i], targets[j]) else -1
            scores[i + 1][j + 1] = max(scores[i][j + 1], scores[i + 1][j], score)
    result = []
    i, j = m, n
    while i and j:
        if compatible(sources[i - 1], targets[j - 1]) and scores[i][j] == scores[i - 1][j - 1] + weight(i - 1, j - 1):
            result.append((sources[i - 1], targets[j - 1])); i -= 1; j -= 1
        elif scores[i][j] == scores[i - 1][j]:
            i -= 1
        else:
            j -= 1
    return list(reversed(result))


def unique_alignment(sources, targets, compatible):
    """Accept a complete ordered match only when earliest/latest positions agree."""
    def positions(reverse=False):
        indexes = range(len(targets))
        available = iter(reversed(indexes) if reverse else indexes)
        found = []
        for source in reversed(sources) if reverse else sources:
            target = next((i for i in available if compatible(source, targets[i])), None)
            if target is None:
                return []
            found.append(target)
        return list(reversed(found)) if reverse else found
    first, last = positions(), positions(True)
    return [(r, targets[i]) for r, i in zip(sources, first)] if first and first == last else []


def supplement(db, progress=None):
    """Resolve script positions independently, then fill only unlocked targets."""
    if progress:
        progress("保留 v1.1.15 推荐；按 ID、角色与台本顺序补查剩余位置……")
    aliases = set()
    # Learn aliases only from repeated unambiguous ID + identical Chinese evidence.
    pairs, totals, evidence = Counter(), Counter(), set()
    for r in records(db, "SELECT m.*,s.chinese FROM key_sources m JOIN source_rows s ON s.rowid=m.source_id WHERE m.dialogue!='' AND m.excluded=''"):
        if not r["dialogue"].isdigit() or len(content_key(r["chinese"], parentheses=True)) < 7:
            continue
        targets = list(records(db, "SELECT * FROM key_baselines WHERE key IN (?,?)", tuple(f"{f}_{r['dialogue']}_Dialog_Detail_v0" for f in ("story_text", "story_convenient"))))
        names = {e["speaker"] for e in targets if e["loose"] == content_key(r["chinese"], parentheses=True)}
        identity = (r["speaker"], r["dialogue"])
        if len(names) == 1 and r["speaker"] and identity not in evidence:
            evidence.add(identity)
            totals[r["speaker"]] += 1
            pairs[r["speaker"], next(iter(names))] += 1
    aliases.update((a, b) for (a, b), n in pairs.items() if b and n >= 3 and n / totals[a] > .5)

    def role(r, e):
        return 1 if (r["speaker"], e["speaker"]) in aliases else speaker_compare(r["speaker"], e["speaker"])

    voice_numbers = {}
    def allowed(r, e):
        if role(r, e) < 0:
            return False
        if not r["char_voice"]:
            return True
        voice = VOICE_KEY.match(e["key"])
        if not voice:
            return False
        name = r["speaker"]
        if name not in voice_numbers:
            voice_numbers[name] = ({"199002"} if name == "男主角" else {"199001"} if name == "女主角" else
                {m[1] for x in records(db, "SELECT key,speaker FROM key_baselines WHERE key LIKE 'char_voice_%'")
                 if role(r, x) == 1 and (m := VOICE_KEY.match(x["key"]))})
        return voice[1] in voice_numbers[name] if voice_numbers[name] else role(r, e) == 1

    def candidates(r, text=None):
        z = r["z"] if text is None else content_key(text, parentheses=True)
        result = list(records(db, "SELECT b.* FROM key_baselines b JOIN key_variants v ON v.baseline_id=b.id WHERE v.text=? ORDER BY b.id", (z,))) if z else []
        selected = [e for e in result if allowed(r, e)]
        if result and not selected and text is None:
            r["method"] = "同中文候选均不符合角色或角色语音范围，未自动采用"
        return selected

    def choose(r, e, method, score=1.0):
        if speaker_compare(r["speaker"], e["speaker"]) < 0 and (r["speaker"], e["speaker"]) in aliases:
            method += f"；角色别名（{r['speaker']}↔{e['speaker']}，{pairs[r['speaker'], e['speaker']]}个独立ID证据）"
        r.update(chosen=e, method=method, score=score)

    sheets = db.execute("SELECT DISTINCT path,sheet FROM source_rows ORDER BY path,sheet").fetchall()
    for index, (path, sheet) in enumerate(sheets, 1):
        sources = [dict(r) for r in records(db, """SELECT m.*,s.chinese,s.korean,s.row,s.seq,
                    EXISTS(SELECT 1 FROM matched_sources u WHERE u.source_id=s.rowid) AS used FROM key_sources m
                    JOIN source_rows s ON s.rowid=m.source_id WHERE s.path=? AND s.sheet=? ORDER BY s.row""", (path, sheet))]
        current_plot, plot_sequence, previous_id = "", 0, ""
        for r in sources:
            r.update(z=content_key(r["chinese"], parentheses=True), chosen=None, method="", score=None, cands=[], options=[], no_slot=False)
            parts = option_parts(r["chinese"])
            if r["plot"].isdigit() and r["plot"] != current_plot:
                current_plot, plot_sequence = r["plot"], 0
            elif r["plot"] and not r["plot"].isdigit():
                current_plot = ""
            if r["excluded"]:
                continue
            did = r["dialogue"] if numeric_id(r["dialogue"]) is not None else ""
            if not r["has_dialogue"] and numeric_id(current_plot) is not None and len(current_plot) <= 16:
                plot_sequence += 1
                did = str(int(current_plot) * 100 + plot_sequence)
            r["effective_id"] = did
            if parts:
                korean_parts = option_parts(r["korean"])
                groups = []
                opt_id = r["dialogue"] if r["dialogue"].isdigit() else previous_id
                if opt_id:
                    for family in ("story_text", "story_convenient"):
                        es = list(records(db, "SELECT * FROM key_baselines WHERE option_base=? ORDER BY option_number,id", (f"{family}_{opt_id}",)))
                        if es:
                            groups.append(es)
                if not groups:
                    bases = {e["option_base"] for e in candidates(r, parts[0]) if e["option_base"]}
                    groups = [list(records(db, "SELECT * FROM key_baselines WHERE option_base=? ORDER BY option_number,id", (b,))) for b in sorted(bases)]
                groups = [es for es in groups if len(es) == len(parts) == len(korean_parts) and
                          [e["option_number"] for e in es] in (list(range(len(parts))), list(range(1, len(parts) + 1))) and
                          all(allowed(r, e) and content_key(z, parentheses=True) == e["loose"] for z, e in zip(parts, es))]
                if len(groups) == 1:
                    r["options"] = list(zip(groups[0], korean_parts))
                    r["method"] = "客户规则：选项按 ID 键组及逐项中文定位" if opt_id else "客户规则：选项全部中文唯一键组"
                    r["method"] += f"；键编号 Option{groups[0][0]['option_number']}～Option{groups[0][-1]['option_number']}"
                else:
                    r["method"] = "选项组、编号或中韩段数不唯一，保留待确认"
                continue
            if did:
                previous_id = did
                direct = list(records(db, "SELECT * FROM key_baselines WHERE key IN (?,?) ORDER BY id", tuple(f"{f}_{did}_Dialog_Detail_v0" for f in ("story_text", "story_convenient"))))
                exact = [e for e in direct if allowed(r, e) and e["loose"] == r["z"]]
                if exact and len({e["number"] for e in exact}) == 1:
                    choose(r, exact[0], "客户规则：对话ID与中文一致" if r["has_dialogue"] else "客户规则：剧情ID生成键与中文一致")
                if not r["chosen"]:
                    near = list(records(db, "SELECT * FROM key_baselines WHERE number BETWEEN ? AND ? ORDER BY id", (int(did) - 5, int(did) + 5)))
                    same = [e for e in near if allowed(r, e) and e["loose"] == r["z"]]
                    if len(same) == 1:
                        choose(r, same[0], f"客户规则：ID ±5 原文修正（{did}）")
                if not r["chosen"] and direct:
                    own = max(similarity(r["z"], e["loose"]) for e in direct)
                    r["method"] = f"ID原文相似度 {own:.1%}，需另找可确定的中文及角色位置"
                    if own < .6 and len(r["z"]) >= 4:
                        near = list(records(db, "SELECT * FROM key_baselines WHERE family=? AND number BETWEEN ? AND ?", (direct[0]["family"], int(did) - 9, int(did) + 9)))
                        scored = sorted([(similarity(r["z"], e["loose"]), e) for e in near if allowed(r, e)], key=lambda x: x[0], reverse=True)
                        if scored and scored[0][0] >= .7 and scored[0][0] >= own + .2 and (len(scored) == 1 or scored[0][0] > scored[1][0]):
                            choose(r, scored[0][1], f"客户规则：ID ±9 相似原文修正（{did}）", scored[0][0])
                    if not r["chosen"]:
                        strong = [e for e in direct if allowed(r, e) and similarity(r["z"], e["loose"]) >= .9]
                        if len(strong) == 1:
                            choose(r, strong[0], "客户规则：ID对应原文写法差异", similarity(r["z"], strong[0]["loose"]))
            r["cands"] = candidates(r)
            if not r["chosen"] and len(r["cands"]) == 1 and len(r["z"]) >= 7:
                choose(r, r["cands"][0], "客户规则：去标点、括号提示及异体字后中文唯一")

        # A repeated short dialogue ID may denote a block whose real keys have
        # one or two extra digits. Require one table/sheet and explicit text evidence.
        blocks = {}
        for r in sources:
            if r.get("effective_id") and not r["chosen"] and not r["excluded"] and not r["options"]:
                blocks.setdefault(r["effective_id"], []).append(r)
        for did, group in blocks.items():
            es = list(records(db, """SELECT * FROM key_baselines WHERE family!=''
                AND CAST(number AS TEXT) LIKE ? AND length(CAST(number AS TEXT)) BETWEEN ? AND ? ORDER BY number,id""", (did + "%", len(did) + 1, len(did) + 2)))
            if not es or len(es) > 60 or len({(e["path"], e["sheet"], e["family"]) for e in es}) != 1:
                continue
            for r, e in ordered_alignment(group, es, lambda r, e: r["z"] == e["loose"] and allowed(r, e)):
                choose(r, e, f"客户规则：ID块 {did} 内原文及顺序一致")

        def firm(r):
            return r["chosen"] is not None

        def anchors(pos):
            a = next((r["chosen"] for r in reversed(sources[:pos]) if firm(r)), None)
            b = next((r["chosen"] for r in sources[pos + 1:] if firm(r)), None)
            return a, b

        def same(r, e):
            return bool(r["z"]) and r["z"] == e["loose"] and allowed(r, e)

        # Two confirmed anchors, never cross target workbook/sheet or a 60-row window.
        fixed = [i for i, r in enumerate(sources) if firm(r)]
        for start, stop in zip(fixed, fixed[1:]):
            a, b = sources[start]["chosen"], sources[stop]["chosen"]
            gap = [r for r in sources[start + 1:stop] if not r["excluded"] and not r["options"] and r["z"]]
            if not gap or len(gap) > 60 or (a["path"], a["sheet"]) != (b["path"], b["sheet"]):
                continue
            if a["family"] and a["family"] == b["family"] and 0 < b["number"] - a["number"] <= 60:
                es = list(records(db, "SELECT * FROM key_baselines WHERE path=? AND sheet=? AND family=? AND number>? AND number<? ORDER BY number,id", (a["path"], a["sheet"], a["family"], a["number"], b["number"])))
                method = "客户规则：前后确定ID之间按台本顺序对齐"
            elif 0 < b["row"] - a["row"] <= 60:
                es = list(records(db, "SELECT * FROM key_baselines WHERE path=? AND sheet=? AND row>? AND row<? AND option_base='' ORDER BY row", (a["path"], a["sheet"], a["row"], b["row"])))
                method = "客户规则：前后确定台词之间按表行顺序对齐"
            else:
                continue
            for r, e in ordered_alignment(gap, es, same):
                choose(r, e, method)

        for _ in range(3):
            changed = 0
            for pos, r in enumerate(sources):
                if firm(r) or r["excluded"] or r["options"] or r["no_slot"] or not r["cands"]:
                    continue
                a, b = anchors(pos)
                before = [x["z"] for x in sources[:pos] if not x["excluded"] and x["z"]][-3:]
                after = [x["z"] for x in sources[pos + 1:] if not x["excluded"] and x["z"]][:3]
                scores = []
                for e in r["cands"]:
                    if e["family"]:
                        context = list(records(db, "SELECT loose,number AS position FROM key_baselines WHERE path=? AND sheet=? AND family=? AND number BETWEEN ? AND ?", (e["path"], e["sheet"], e["family"], e["number"] - 5, e["number"] + 5)))
                        position = e["number"]
                    else:
                        context = list(records(db, "SELECT loose,row AS position FROM key_baselines WHERE path=? AND sheet=? AND row BETWEEN ? AND ?", (e["path"], e["sheet"], e["row"] - 5, e["row"] + 5)))
                        position = e["row"]
                    pe = {x["loose"] for x in context if x["position"] < position}
                    ne = {x["loose"] for x in context if x["position"] > position}
                    score = sum(x in pe for x in before) + sum(x in ne for x in after)
                    distances = [abs(e["number"] - x["number"]) for x in (a, b) if x and x["family"] and x["family"] == e["family"]]
                    distance = min(distances, default=10**20)
                    scores.append((score, role(r, e), -distance, e))
                scores.sort(key=lambda x: x[:3], reverse=True)
                top = scores[0]
                if top[0] > 0 and (len(scores) == 1 or top[:2] > scores[1][:2]):
                    window = "键编号±5" if top[3]["family"] else "邻近5行"
                    choose(r, top[3], f"客户规则：台本前后3句与表{window}重合{top[0]}句")
                elif -top[2] <= 30 and (len(scores) == 1 or top[2] - scores[1][2] >= 5):
                    choose(r, top[3], f"客户规则：邻近确定键编号距离{-top[2]}")
                elif a and b and (a["path"], a["sheet"]) == (b["path"], b["sheet"]) and 0 < b["row"] - a["row"] <= 30:
                    between = [e for e in r["cands"] if (e["path"], e["sheet"]) == (a["path"], a["sheet"]) and a["row"] < e["row"] < b["row"]]
                    if len(between) == 1:
                        choose(r, between[0], "客户规则：前后确定台词30行内唯一中文位置")
                    elif not between:
                        r["no_slot"] = True
                        r["method"] = "此处无键：前后确定台词之间不存在该中文，其他位置同文不代用"
                changed += int(firm(r))
            if not changed:
                break

        # Recheck only this new pass, never the locked v1.1.15 recommendations.
        fixed = [i for i, r in enumerate(sources) if firm(r) and
                 ("ID" in r["method"] or ("中文唯一" in r["method"] and len(r["z"]) >= 7))]
        for start, stop in zip(fixed, fixed[1:]):
            a, b = sources[start]["chosen"], sources[stop]["chosen"]
            gap = [r for r in sources[start + 1:stop] if not r["excluded"] and r["cands"] and not r["options"] and not r["no_slot"]]
            if not gap or len(gap) > 60 or (a["path"], a["sheet"]) != (b["path"], b["sheet"]) or not 0 < b["row"] - a["row"] <= 60:
                continue
            es = list(records(db, "SELECT * FROM key_baselines WHERE path=? AND sheet=? AND row>? AND row<? AND family!='' ORDER BY row", (a["path"], a["sheet"], a["row"], b["row"])))
            for r, e in ordered_alignment(gap, es, same):
                if not firm(r) or r["chosen"]["id"] != e["id"]:
                    choose(r, e, "客户规则：确定台词之间60行内流程校正")

        for pos, r in enumerate(sources):
            if firm(r) or r["excluded"] or r["options"] or r["cands"] or r["used"] or not r["z"] or option_parts(r["chinese"]):
                continue
            a, b = anchors(pos)
            pool, threshold, label = [], .9, "客户规则：全表相似原文"
            if a and a["family"] and len(r["z"]) >= 4:
                pool = list(records(db, "SELECT * FROM key_baselines WHERE path=? AND sheet=? AND family=? AND number BETWEEN ? AND ?", (a["path"], a["sheet"], a["family"], a["number"] - 21, a["number"] + 21)))
                threshold, label = .6, "客户规则：相邻ID ±21 相似原文"
            elif r["char_voice"] and r["speaker"] and len(r["z"]) >= 4:
                pool = list(records(db, "SELECT * FROM key_baselines WHERE key LIKE 'char_voice_%' AND speaker=?", (r["speaker"],)))
                threshold, label = .7, "客户规则：同角色语音相似原文"
            best = []
            for e in pool:
                if allowed(r, e) and (score := similarity(r["z"], e["loose"])) >= threshold:
                    best.append((score, e))
            if not best and len(r["z"]) >= 6:
                threshold, label = .9, "客户规则：全表相似原文"
                # Disk character index: no all-pairs scan or arbitrary top-60 cutoff.
                if not db.execute("SELECT 1 FROM sqlite_temp_master WHERE name='key_characters'").fetchone():
                    db.executescript("""CREATE TEMP TABLE key_characters (character TEXT,size INTEGER,id INTEGER,n INTEGER,PRIMARY KEY(character,size,id)) WITHOUT ROWID;
                        CREATE TEMP TABLE key_query (character TEXT PRIMARY KEY,n INTEGER) WITHOUT ROWID;""")
                    for e in records(db, """SELECT b.id,b.loose FROM key_baselines b JOIN recommendations r ON r.path=b.path AND r.sheet=b.sheet AND r.row=b.row
                        WHERE r.source_id IS NULL AND NOT EXISTS(SELECT 1 FROM structure_recommendations t WHERE t.path=b.path AND t.sheet=b.sheet AND t.row=b.row)"""):
                        db.executemany("INSERT INTO key_characters VALUES (?,?,?,?)", [(c, len(e["loose"]), e["id"], n) for c, n in Counter(e["loose"]).items()])
                length = len(r["z"])
                db.execute("DELETE FROM key_query")
                db.executemany("INSERT INTO key_query VALUES (?,?)", Counter(r["z"]).items())
                for e in records(db, """SELECT b.* FROM key_baselines b JOIN (
                    SELECT f.id FROM key_query q JOIN key_characters f ON f.character=q.character AND f.size BETWEEN ? AND ?
                    GROUP BY f.id HAVING 20*SUM(MIN(q.n,f.n)) >= 9*(f.size+?)) c ON c.id=b.id""", ((9 * length + 10) // 11, 11 * length // 9, length)):
                    if not allowed(r, e):
                        continue
                    if (score := similarity(r["z"], e["loose"])) >= .9:
                        best.append((score, e))
                if len({e["loose"] for _, e in best}) != 1:
                    best = []
            best.sort(key=lambda x: x[0], reverse=True)
            if best and (len(best) == 1 or best[0][0] > best[1][0]):
                choose(r, best[0][1], label, best[0][0])

        # The customer's edge-window pass also works with only one nearby anchor.
        # Run after existing rules and require a unique complete alignment.
        fixed = [-1, *[i for i, r in enumerate(sources) if firm(r)], len(sources)]
        occupied = {r["chosen"]["id"] for r in sources if firm(r)}
        for start, stop in zip(fixed, fixed[1:]):
            a = sources[start]["chosen"] if start >= 0 else None
            b = sources[stop]["chosen"] if stop < len(sources) else None
            gap = [r for r in sources[start + 1:stop] if not r["excluded"] and
                   not option_parts(r["chinese"]) and r["z"] and not r["no_slot"]]
            if not gap or len(gap) > 15:
                continue
            if a and b and a["family"] and (a["path"], a["sheet"], a["family"]) == (b["path"], b["sheet"], b["family"]) and 0 < b["number"] - a["number"] <= 60:
                continue  # Already handled by the two-anchor pass.
            anchor = a if a and a["family"] else b if b and b["family"] else None
            if anchor is None:
                continue
            forward = anchor is a
            low, high = ((anchor["number"] + 1, anchor["number"] + 15) if forward else
                         (anchor["number"] - 15, anchor["number"] - 1))
            es = [e for e in records(db, """SELECT * FROM key_baselines
                WHERE path=? AND sheet=? AND family=? AND number BETWEEN ? AND ? ORDER BY number,id""",
                (anchor["path"], anchor["sheet"], anchor["family"], low, high)) if e["id"] not in occupied]
            # A physically adjacent second anchor still bounds the scene even
            # when its ID switches families or jumps backwards.
            if a and b and (a["path"], a["sheet"]) == (b["path"], b["sheet"]):
                if b["row"] <= a["row"]:
                    continue
                es = [e for e in es if a["row"] < e["row"] < b["row"]]
            for r, e in unique_alignment(gap, es, same):
                choose(r, e, f"{EDGE_MATCH}；{'前' if forward else '后'}侧确定键 {anchor['key']}，"
                       f"{'向后' if forward else '向前'}15个ID内中文及顺序唯一对齐（{len(gap)}句）")
                occupied.add(e["id"])

        for r in sources:
            if r["excluded"]:
                db.execute("INSERT INTO key_audit VALUES (?,NULL,'',NULL,?)", (r["source_id"], r["excluded"]))
                continue
            proposals = [(r["chosen"], r["korean"], 0, 1)] if firm(r) else [(e, kr, i, len(r["options"])) for i, (e, kr) in enumerate(r["options"], 1)]
            if not proposals:
                reason = ("未定位：短句缺少上下文或位置依据" if len(r["z"]) <= 6 and r["cands"] else
                          "未定位：同中文候选不唯一" if r["cands"] else "未定位：没有满足中文、角色及相似度门槛的候选")
                db.execute("INSERT INTO key_audit VALUES (?,NULL,?,NULL,?)", (r["source_id"], r["method"], reason))
            for e, korean, part, total in proposals:
                table = "key_edge_proposals" if r["method"].startswith(EDGE_MATCH) else "key_proposals"
                db.execute(f"INSERT OR IGNORE INTO {table} VALUES (?,?,?,?,?,?,?)", (e["id"], r["source_id"], korean, r["method"], r["score"] if r["score"] is not None else 1.0, part, total))
        if progress and index % 5 == 0:
            progress(f"客户规则补匹配：已处理 {index}/{len(sheets)} 张台本 Sheet")
    # Propagate only same-type, compatible duplicate keys. Short lines also need
    # an identical long neighbour; keys claimed by other scripts are not stolen.
    db.execute("CREATE TEMP TABLE key_claimed AS SELECT DISTINCT baseline_id FROM key_proposals")
    for p in records(db, """SELECT p.*,b.key,b.loose,b.path,b.sheet,b.row,m.speaker,m.char_voice
        FROM key_proposals p JOIN key_baselines b ON b.id=p.baseline_id JOIN key_sources m ON m.source_id=p.source_id
        WHERE p.part=0 AND p.baseline_id IN (SELECT baseline_id FROM key_claimed) ORDER BY p.baseline_id,p.source_id"""):
        def family(key):
            return "dialog" if STORY_KEY.match(key) else "voice" if VOICE_KEY.match(key) else re.sub(r"\d+", "#", key)
        if not p["key"]:
            continue
        for e in records(db, "SELECT * FROM key_baselines WHERE loose=? AND id!=? AND option_base=''", (p["loose"], p["baseline_id"])):
            if not e["key"] or family(e["key"]) != family(p["key"]) or not allowed(p, e):
                continue
            if db.execute("SELECT 1 FROM key_claimed WHERE baseline_id=?", (e["id"],)).fetchone():
                continue
            if len(p["loose"]) <= 6:
                neighbours = [db.execute("SELECT loose FROM key_baselines WHERE path=? AND sheet=? AND row=?", (p["path"], p["sheet"], p["row"] + d)).fetchone() for d in (-1, 1)]
                if not any(n and len(n[0]) >= 7 and db.execute("SELECT 1 FROM key_baselines WHERE path=? AND sheet=? AND row=? AND loose=?", (e["path"], e["sheet"], e["row"] + d, n[0])).fetchone() for d, n in zip((-1, 1), neighbours)):
                    continue
            db.execute("INSERT OR IGNORE INTO key_proposals VALUES (?,?,?,?,?,0,1)", (e["id"], p["source_id"], p["korean"], p["method"] + "；同类型、同角色重复键同步", p["similarity"]))
    # Snapshot all old claims, including duplicate propagation, before inserting
    # edge proposals. Preserve old results across source files; keep every new
    # competing translation together rather than accepting whichever is first.
    db.execute("INSERT INTO key_claimed SELECT DISTINCT baseline_id FROM key_proposals WHERE baseline_id NOT IN (SELECT baseline_id FROM key_claimed)")
    db.execute("""INSERT INTO key_audit SELECT source_id,baseline_id,method,similarity,
        '未采用单侧补查：此键已有客户定位候选，保留原判定' FROM key_edge_proposals
        WHERE baseline_id IN (SELECT baseline_id FROM key_claimed)""")
    db.execute("INSERT OR IGNORE INTO key_proposals SELECT * FROM key_edge_proposals WHERE baseline_id NOT IN (SELECT baseline_id FROM key_claimed)")
    return apply_proposals(db)


def apply_proposals(db):
    added = 0
    for e in records(db, "SELECT DISTINCT b.* FROM key_baselines b JOIN key_proposals p ON p.baseline_id=b.id ORDER BY b.id"):
        key = (e["path"], e["sheet"], e["row"])
        proposals = list(records(db, "SELECT * FROM key_proposals WHERE baseline_id=? ORDER BY source_id,part", (e["id"],)))
        locked = db.execute("SELECT 1 FROM recommendations WHERE path=? AND sheet=? AND row=? AND source_id IS NOT NULL UNION ALL SELECT 1 FROM structure_recommendations WHERE path=? AND sheet=? AND row=?", (*key, *key)).fetchone()
        structural_conflict = db.execute("SELECT 1 FROM structure_pending WHERE path=? AND sheet=? AND row=?", key).fetchone()
        translations = {p["korean"].strip() for p in proposals if p["korean"] and p["korean"].strip()}
        outcome = ("保留 v1.1.15 已有推荐" if locked else "结构推荐竞争，仍需人工复核" if structural_conflict else
                   "定位到同一位置的台本译文冲突，未推荐" if len(translations) > 1 else "台本缺译，未推荐" if not translations else "已补推荐·待复核")
        for p in proposals:
            db.execute("INSERT INTO key_audit VALUES (?,?,?,?,?)", (p["source_id"], e["id"], p["method"], p["similarity"], outcome))
            if locked or structural_conflict or p["part"]:
                continue
            db.execute("INSERT OR IGNORE INTO comparisons VALUES (?,?,?,?,0,0,0,?,'[]','[]','[]','[]','',?,?)", (*key, p["source_id"], "客户键定位补匹配", p["method"] + "；未采用旧三句位置得分", p["similarity"]))
            db.execute("INSERT OR IGNORE INTO matched_sources VALUES (?)", (p["source_id"],))
        if locked or structural_conflict:
            continue
        if len(translations) == 1:
            p = next(p for p in proposals if p["korean"] and p["korean"].strip())
            basis = f"客户规则补匹配；{p['method']}；中文相似度 {p['similarity']:.1%}；键 {e['key'] or '无键，按文件/Sheet/行定位'}；角色无冲突；待复核；保留旧候选"
            if p["part"]:
                source = db.execute("SELECT path,sheet,row FROM source_rows WHERE rowid=?", (p["source_id"],)).fetchone()
                db.execute("INSERT INTO structure_recommendations VALUES (?,?,?,?,?,?,?,?,?,?)", (*key, e["text"], p["korean"], f"{source[0]}；Sheet：{source[1]}；行号：{source[2]}；选项{p['part']}", basis, "选项推荐·待复核", "选项", f"选项{p['source_id']}"))
                db.execute("INSERT INTO structure_sources VALUES (?,?,?,?,?,?,0)", (*key, p["source_id"], p["part"], p["total"]))
                db.execute("INSERT OR IGNORE INTO matched_sources VALUES (?)", (p["source_id"],))
                variants = db.execute("""SELECT COUNT(*) FROM (SELECT s.korean_key FROM comparisons c
                    JOIN source_rows s ON s.rowid=c.source_id WHERE c.path=? AND c.sheet=? AND c.row=? AND s.korean_key!=''
                    UNION SELECT ?)""", (*key, p["korean"].strip())).fetchone()[0]
                db.execute("INSERT INTO baseline_counts VALUES (?,?,?,1,0,?) ON CONFLICT(path,sheet,row) DO UPDATE SET n=n+1,variants=excluded.variants", (*key, variants))
            else:
                db.execute("UPDATE recommendations SET source_id=?,basis=?,status='键定位推荐·待复核',match_type='客户键定位补匹配' WHERE path=? AND sheet=? AND row=?", (p["source_id"], basis, *key))
            added += 1
        elif translations:
            db.execute("UPDATE recommendations SET basis=basis || ?,status='待确认' WHERE path=? AND sheet=? AND row=?", ("；客户规则：" + outcome + "；保留旧候选及客户定位证据", *key))
        if not db.execute("SELECT 1 FROM structure_recommendations WHERE path=? AND sheet=? AND row=?", key).fetchone():
            db.execute("""INSERT OR REPLACE INTO baseline_counts
                SELECT e.path,e.sheet,e.row,COUNT(*),SUM(s.korean_key=''),COUNT(DISTINCT NULLIF(s.korean_key,''))
                FROM comparisons e JOIN source_rows s ON s.rowid=e.source_id
                WHERE e.path=? AND e.sheet=? AND e.row=? GROUP BY e.path,e.sheet,e.row""", key)
    db.commit()
    return added
