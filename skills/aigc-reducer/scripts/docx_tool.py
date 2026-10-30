#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx_tool.py — 论文降 AIGC 率工具四件套

  split   docx -> 段落清单 + 章节树 + 分批方案 (JSON)
  batch   从 split 结果里取一批段落，拼成可直接发送的改写请求
  write   改写结果 -> 新 docx（保留原样式）+ 对照表
  check   改写前后硬指标自检

依赖：python-docx
    安装：python -m pip install python-docx

示例：
  python docx_tool.py split 论文.docx -o work/blocks.json --batch-chars 1200
  python docx_tool.py batch work/blocks.json --id 3 --prompt references/prompt.md -o work/batch3.txt
  python docx_tool.py write 论文.docx work/edits.json -o 论文_降AI版.docx --diff-md work/对照表.md
  python docx_tool.py check 论文.docx work/edits.json -o work/check.json
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime

try:
    from docx import Document
except ImportError:
    sys.exit(
        "缺少依赖 python-docx。\n"
        "请执行：python -m pip install python-docx"
    )

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ---------------------------------------------------------------- 段落分类

REF_MARKERS = ("参考文献", "references", "致谢", "acknowledg", "附录", "appendix")
MANUAL_HEADING = re.compile(r"^\s*(第[一二三四五六七八九十百]+[章节篇]|\d+(?:\.\d+)*[\.、\s])")
CAPTION = re.compile(r"^\s*(图|表|Fig|Table)\s*[\d\-—\.]")
BANNED_SUBJECTIVE = ("我觉得", "我认为", "在我看来", "窃以为", "笔者觉得")
AI_CLICHE = ("综上所述", "值得注意的是", "不难发现", "由此可见", "显著提升", "至关重要")

NUM_RE = re.compile(r"\d+(?:\.\d+)?")
CITE_RE = re.compile(r"[\[（(]\s*\d+(?:\s*[-,，–—~]\s*\d+)*\s*[\]）)]")

# 只有 para 允许被改写；其余类型一律拦截，避免标题、图注被正文覆盖
EDITABLE_KINDS = ("para",)
KIND_LABEL = {
    "para": "正文",
    "heading": "标题",
    "caption": "图表题注",
    "ref": "参考文献/致谢",
    "formula": "公式",
    "skip": "空白或图片",
}


def has_image(p):
    xml = p._element.xml
    return "<w:drawing" in xml or "<w:pict" in xml or "<pic:pic" in xml


def has_formula(p):
    xml = p._element.xml
    return "<m:oMath" in xml or "<m:oMathPara" in xml


def style_name(p):
    try:
        return p.style.name or ""
    except Exception:
        return ""


def is_heading_style(p):
    name = style_name(p)
    if not name:
        return False
    n = name.lower()
    return n.startswith("heading") or n.startswith("title") or "标题" in name


def looks_like_manual_heading(text):
    t = text.strip()
    if not t or len(t) > 50:
        return False
    if t.endswith(("。", "！", "？", ".", ";", "；")):
        return False
    return bool(MANUAL_HEADING.match(t))


def classify(p, in_refs):
    """返回 (kind, text)"""
    text = p.text or ""
    stripped = text.strip()

    if has_image(p) and not stripped:
        return "skip", text
    if has_formula(p):
        return "formula", text
    if not stripped:
        return "skip", text

    if is_heading_style(p):
        return "heading", text
    if CAPTION.match(stripped):
        return "caption", text
    if looks_like_manual_heading(stripped):
        return "heading", text
    if in_refs:
        return "ref", text
    return "para", text


def scan_blocks(doc):
    blocks = []
    in_refs = False
    for i, p in enumerate(doc.paragraphs):
        text = p.text or ""
        low = text.strip().lower()
        if not in_refs and is_heading_style(p) and any(m in low for m in REF_MARKERS):
            in_refs = True
        kind, txt = classify(p, in_refs)
        if kind != "ref" and len(text.strip()) < 20 and any(k in low for k in REF_MARKERS):
            in_refs = True
        blocks.append(
            {
                "id": i,
                "kind": kind,
                "style": style_name(p),
                "chars": len(txt.strip()),
                "runs": len(p.runs),
                "text": txt,
            }
        )
    return blocks


def build_sections(blocks):
    """章节树：每个 heading 后面到下一个 heading 之前的正文统计"""
    sections = []
    cur = None
    for b in blocks:
        if b["kind"] == "heading":
            cur = {
                "heading_id": b["id"],
                "heading": b["text"].strip(),
                "para_ids": [],
                "chars": 0,
            }
            sections.append(cur)
        elif b["kind"] == "para" and cur is not None:
            cur["para_ids"].append(b["id"])
            cur["chars"] += b["chars"]
    return sections


def build_batches(blocks, limit):
    batches = []
    cur_ids, cur_chars, cur_sec = [], 0, "（文首）"

    def flush():
        nonlocal cur_ids, cur_chars
        if cur_ids:
            batches.append(
                {
                    "batch_id": len(batches) + 1,
                    "section": cur_sec,
                    "para_ids": list(cur_ids),
                    "chars": cur_chars,
                }
            )
            cur_ids, cur_chars = [], 0

    for b in blocks:
        if b["kind"] == "heading":
            flush()
            cur_sec = b["text"].strip()[:60]
            continue
        if b["kind"] != "para":
            flush()
            continue
        if cur_ids and cur_chars + b["chars"] > limit:
            flush()
        cur_ids.append(b["id"])
        cur_chars += b["chars"]
    flush()
    return batches


# ---------------------------------------------------------------- split

def cmd_split(args):
    doc = Document(args.docx)
    blocks = scan_blocks(doc)
    sections = build_sections(blocks)
    batches = build_batches(blocks, args.batch_chars)

    eligible = [b for b in blocks if b["kind"] == "para"]
    tables = len(doc.tables)
    table_chars = 0
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                table_chars += len(cell.text or "")

    out = {
        "source": os.path.abspath(args.docx),
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "batch_chars": args.batch_chars,
        "summary": {
            "para_total": len(blocks),
            "eligible_paras": len(eligible),
            "eligible_chars": sum(b["chars"] for b in eligible),
            "heading_count": sum(1 for b in blocks if b["kind"] == "heading"),
            "ref_paras": sum(1 for b in blocks if b["kind"] == "ref"),
            "formula_paras": sum(1 for b in blocks if b["kind"] == "formula"),
            "caption_paras": sum(1 for b in blocks if b["kind"] == "caption"),
            "tables": tables,
            "table_chars": table_chars,
            "batches": len(batches),
        },
        "sections": sections,
        "batches": batches,
        "blocks": blocks,
    }

    write_json(args.output, out)
    s = out["summary"]
    print(f"[split] {os.path.basename(args.docx)}")
    print(f"  段落总数      {s['para_total']}")
    print(f"  可改写正文     {s['eligible_paras']} 段 / {s['eligible_chars']} 字")
    print(f"  标题 / 参考文献 / 公式 / 题注  {s['heading_count']} / {s['ref_paras']} / {s['formula_paras']} / {s['caption_paras']}")
    print(f"  表格          {s['tables']} 个（{s['table_chars']} 字，不参与改写）")
    print(f"  分批          {s['batches']} 批（每批 ≤{args.batch_chars} 字）")
    print(f"  输出          {args.output}")


# ---------------------------------------------------------------- batch

def read_prompt(path):
    """取第一个 --- 分隔线之后的全部内容（含「修改段落:」）"""
    with open(path, "r", encoding="utf-8") as f:
        raw = f.read()
    m = re.search(r"(?m)^---\s*$", raw)
    if not m:
        sys.exit(f"[错误] {path} 中没有找到 --- 分隔线，无法提取改写指令")
    body = raw[m.end():].lstrip("\n")
    if "修改段落" not in body:
        print(
            f"[警告] 从 {os.path.basename(path)} 提取的指令里没有「修改段落:」，"
            "请确认 --prompt 指向 references/prompt.md",
            file=sys.stderr,
        )
    return body


def match_section(sections, keyword):
    kw = keyword.strip()
    hits = [s for s in sections if s["heading"].strip() == kw]
    if not hits:
        hits = [s for s in sections if kw and (kw in s["heading"] or s["heading"] in kw)]
    if not hits:
        # 去编号后再试
        norm = re.sub(r"^[\d\.\s、]+", "", kw)
        if norm:
            hits = [s for s in sections if norm in s["heading"]]
    return hits


def cmd_batch(args):
    data = load_json(args.blocks)
    block_map = {b["id"]: b for b in data["blocks"]}
    batches = data["batches"]

    if args.id is not None:
        hit = [b for b in batches if b["batch_id"] == args.id]
        if not hit:
            sys.exit(f"[错误] 没有第 {args.id} 批，共 {len(batches)} 批")
        ids = hit[0]["para_ids"]
        section = hit[0]["section"]
        label = f"批次 {args.id} / 共 {len(batches)} 批"
    elif args.ids:
        raw = args.ids.replace("，", ",")
        try:
            ids = [int(x) for x in raw.split(",") if x.strip()]
        except ValueError:
            sys.exit(f"[错误] --ids 只接受逗号分隔的整数，收到：{args.ids}")
        section = ""
        label = "自定义段落"
    elif args.section:
        hits = match_section(data["sections"], args.section)
        if not hits:
            names = [s["heading"] for s in data["sections"] if s["heading"]]
            sys.exit(
                f"[错误] 没有匹配到章节「{args.section}」。\n"
                f"可用章节：\n  " + "\n  ".join(names[:40])
            )
        if len(hits) > 1:
            names = [s["heading"] for s in hits]
            print(f"[提示] 「{args.section}」匹配到 {len(hits)} 节，取第一节：{names[0]}", file=sys.stderr)
        ids = list(hits[0]["para_ids"])
        section = hits[0]["heading"]
        label = f"章节 {section}"
    else:
        sys.exit("需要 --id / --ids / --section 之一")

    missing = [i for i in ids if i not in block_map]
    if missing:
        sys.exit(f"[错误] blocks 里不存在这些段落 id：{missing}")
    blocked = [i for i in ids if block_map[i]["kind"] not in EDITABLE_KINDS]
    if blocked:
        detail = "\n  ".join(
            f"{i} -> {KIND_LABEL.get(block_map[i]['kind'], block_map[i]['kind'])}：{block_map[i]['text'][:30]}"
            for i in blocked
        )
        sys.exit(f"[错误] 这些 id 不是正文段落，不能参与改写：\n  {detail}")
    if not ids:
        sys.exit("[错误] 这一批没有可改写的正文段落")

    texts = [block_map[i]["text"].strip() for i in ids]
    chars = sum(len(t) for t in texts)
    body = "\n\n".join(texts)

    if args.prompt:
        payload = read_prompt(args.prompt).rstrip() + "\n" + body + "\n"
    else:
        payload = body + "\n"

    if args.output:
        os.makedirs(os.path.dirname(os.path.abspath(args.output)) or ".", exist_ok=True)
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(payload)

    print(f"[batch] {label}", file=sys.stderr)
    if section:
        print(f"  章节   {section}", file=sys.stderr)
    print(f"  段落   {len(ids)} 段 / {chars} 字", file=sys.stderr)
    print(f"  id 顺序 {','.join(str(i) for i in ids)}", file=sys.stderr)
    if args.output:
        print(f"  请求文本 -> {args.output}（可直接复制发送）", file=sys.stderr)
    else:
        print("  提示：加 -o 保存为纯请求文本，避免和上面的提示混在一起", file=sys.stderr)

    if not args.output:
        sys.stdout.write(payload)


# ---------------------------------------------------------------- write

def load_edits(path):
    data = load_json(path)
    if isinstance(data, dict):
        items = data
    elif isinstance(data, list):
        items = {}
        for it in data:
            if not isinstance(it, dict) or "id" not in it or "text" not in it:
                sys.exit('[错误] edits 列表模式下每项必须是 {"id": 段落号, "text": "新文本"}')
            items[str(it["id"])] = it["text"]
    else:
        sys.exit("[错误] edits 文件格式应为 {id: 新文本} 或 [{id, text}]")

    bad_keys = [k for k in items if not str(k).strip().lstrip("-").isdigit()]
    if bad_keys:
        sys.exit(
            "[错误] edits 的 key 必须是段落整数 id，这些不是："
            f"{bad_keys[:10]}\n"
            "段落 id 请从 blocks.json 的 blocks[].id 取，或先用 batch 子命令确认。"
        )
    bad_vals = [k for k, v in items.items() if not isinstance(v, str) or not v.strip()]
    if bad_vals:
        sys.exit(f"[错误] 这些段落的改写文本为空或不是字符串：{bad_vals[:10]}")
    return {int(k): v for k, v in items.items()}


def same_file(a, b):
    try:
        return os.path.samefile(a, b)
    except OSError:
        return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def check_targets(paras, edits, block_map, allow_non_para):
    """拦截非正文段落，避免把标题/图注/参考文献覆盖成正文"""
    out_of_range, blocked = [], []
    for pid in sorted(edits):
        if pid < 0 or pid >= len(paras):
            out_of_range.append(pid)
            continue
        kind = block_map.get(pid, {}).get("kind", "para")
        if kind not in EDITABLE_KINDS and not allow_non_para:
            blocked.append((pid, kind, block_map.get(pid, {}).get("text", "")))
    return out_of_range, blocked


def run_has_note_ref(r):
    """脚注/尾注引用所在的 run —— 改写时必须原样保留，否则脚注会消失"""
    xml = r._element.xml
    return "footnoteReference" in xml or "endnoteReference" in xml


def format_sig(r):
    """run 的格式签名，用于判断段落内是否存在多种格式"""
    f = r.font
    try:
        color = str(f.color.rgb) if f.color and f.color.rgb else None
    except Exception:
        color = None
    return (bool(r.bold), bool(r.italic), bool(r.underline), color, f.size)


def cmd_write(args):
    if same_file(args.docx, args.output):
        sys.exit(
            "[错误] 输出文件与源文件相同，会覆盖原稿。\n"
            "本工具不覆盖原文件，请换一个 -o 文件名，例如 论文_降AI版.docx"
        )

    doc = Document(args.docx)
    blocks = scan_blocks(doc)
    block_map = {b["id"]: b for b in blocks}
    edits = load_edits(args.edits)
    paras = doc.paragraphs

    out_of_range, blocked = check_targets(paras, edits, block_map, args.allow_non_para)
    for pid in out_of_range:
        print(f"  警告：段落 id {pid} 越界，已跳过", file=sys.stderr)
    if blocked:
        detail = "\n  ".join(
            f"{i} -> {KIND_LABEL.get(k, k)}：{t[:30]}" for i, k, t in blocked
        )
        sys.exit(
            f"[错误] {len(blocked)} 个 id 指向的不是正文段落，已全部拒绝回写：\n  {detail}\n"
            "这类段落（标题、图注、参考文献、公式）被正文覆盖后很难复原。\n"
            "请核对 blocks.json 的 blocks[].id；确需改写请加 --allow-non-para。"
        )

    applied, missing, no_run = 0, [], 0
    skipped_link, multi_fmt, note_kept, note_head = [], [], [], []
    written = set()

    for pid, new_text in sorted(edits.items()):
        if pid < 0 or pid >= len(paras):
            missing.append(pid)
            continue
        p = paras[pid]

        if "<w:hyperlink" in p._element.xml and not args.force:
            skipped_link.append(pid)
            continue

        runs = p.runs
        if not runs:
            p.add_run(new_text)
            no_run += 1
            applied += 1
            written.add(pid)
            continue

        if len({format_sig(r) for r in runs}) > 1:
            multi_fmt.append(pid)

        editable = [r for r in runs if not run_has_note_ref(r)]
        if not editable:
            p.add_run(new_text)
            no_run += 1
            applied += 1
            written.add(pid)
            continue

        if run_has_note_ref(runs[0]):
            note_head.append(pid)
        if len(editable) != len(runs):
            note_kept.append(pid)

        editable[0].text = new_text
        for r in editable[1:]:
            r.text = ""
        applied += 1
        written.add(pid)

    os.makedirs(os.path.dirname(os.path.abspath(args.output)) or ".", exist_ok=True)
    doc.save(args.output)

    print(f"[write] 已回写 {applied} 段 -> {args.output}")
    if no_run:
        print(f"  提示：{no_run} 段原本无可写 run，已新建（继承段落样式）")
    if missing:
        print(f"  警告：{len(missing)} 个段落 id 越界，已跳过：{missing[:10]}")
    if skipped_link:
        print(f"  [跳过] {len(skipped_link)} 段含超链接，未改写（强行改写会让链接文字重复）。需手工处理：{skipped_link[:10]}")
    if multi_fmt:
        print(f"  [提示] {len(multi_fmt)} 段内含多种字符格式（如局部加粗），改写后统一为段首格式，局部强调会丢失：{multi_fmt[:10]}")
    if note_kept:
        print(f"  [提示] {len(note_kept)} 段含脚注/尾注引用，引用已保留，但位置可能随语句调整而偏移：{note_kept[:10]}")
    if note_head:
        print(f"  [警告] {len(note_head)} 段脚注引用位于段首，改写文字可能排到引用之后，建议手工核对：{note_head[:10]}")

    if args.diff_md:
        real_edits = {k: v for k, v in edits.items() if k in written}
        write_diff_md(args.docx, real_edits, args.diff_md)
        print(f"  对照表 -> {args.diff_md}（{len(real_edits)} 段）")


def write_diff_md(src, edits, path):
    doc = Document(src)
    paras = doc.paragraphs
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    lines = [
        "# 改动对照表",
        "",
        f"- 源文件：{os.path.basename(src)}",
        f"- 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}",
        f"- 改动段落：{len(edits)} 段",
        "",
        "| 段落 | 原文字数 | 新文字数 |",
        "|---|---|---|",
    ]
    for pid in sorted(edits):
        if 0 <= pid < len(paras):
            old = (paras[pid].text or "").strip()
            new = edits[pid].strip()
            lines.append(f"| {pid} | {len(old)} | {len(new)} |")
    lines.append("")
    for pid in sorted(edits):
        if not (0 <= pid < len(paras)):
            continue
        old = (paras[pid].text or "").strip()
        new = edits[pid].strip()
        lines += [
            f"## 段落 {pid}",
            "",
            f"**原：** {old}",
            "",
            f"**新：** {new}",
            "",
        ]
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


# ---------------------------------------------------------------- check

def punct_count(t):
    return {
        "period": t.count("。"),
        "comma": t.count("，"),
        "semicolon": t.count("；"),
    }


def strip_cites(t):
    """算数字前后先去掉引用标记，避免 [1] 里的 1 被当成正文数字重复报警"""
    return CITE_RE.sub("", t)


def cmd_check(args):
    doc = Document(args.docx)
    blocks = scan_blocks(doc)
    block_map = {b["id"]: b for b in blocks}
    edits = load_edits(args.edits)
    paras = doc.paragraphs

    out_of_range, blocked = check_targets(paras, edits, block_map, args.allow_non_para)

    old_all, new_all = [], []
    per_para = []
    for pid in sorted(edits):
        if not (0 <= pid < len(paras)):
            continue
        if any(pid == b[0] for b in blocked):
            continue
        old = (paras[pid].text or "").strip()
        new = edits[pid].strip()
        old_all.append(old)
        new_all.append(new)
        po, pn = punct_count(old), punct_count(new)
        old_nums = set(NUM_RE.findall(strip_cites(old)))
        new_nums = set(NUM_RE.findall(strip_cites(new)))
        per_para.append(
            {
                "id": pid,
                "chars_old": len(old),
                "chars_new": len(new),
                "chars_pct": round((len(new) - len(old)) / max(len(old), 1) * 100, 1),
                "period_old": po["period"],
                "period_new": pn["period"],
                "numbers_lost": sorted(old_nums - new_nums),
                "cites_lost": sorted(set(CITE_RE.findall(old)) - set(CITE_RE.findall(new))),
                "banned": [w for w in BANNED_SUBJECTIVE if w in new],
            }
        )

    old_text, new_text = "".join(old_all), "".join(new_all)
    po, pn = punct_count(old_text), punct_count(new_text)

    all_nums_lost, all_cites_lost, banned_hits = [], [], []
    for r in per_para:
        all_nums_lost += r["numbers_lost"]
        all_cites_lost += r["cites_lost"]
        banned_hits += r["banned"]

    char_pct = round((len(new_text) - len(old_text)) / max(len(old_text), 1) * 100, 1)
    period_pct = (
        round((pn["period"] - po["period"]) / max(po["period"], 1) * 100, 1)
        if po["period"]
        else 0.0
    )

    issues, warnings = [], []
    if blocked:
        detail = "；".join(f"{i}({KIND_LABEL.get(k, k)})" for i, k, _ in blocked)
        issues.append(f"{len(blocked)} 个 id 不是正文段落，未计入自检：{detail}")
    if out_of_range:
        warnings.append(f"{len(out_of_range)} 个段落 id 越界，已忽略：{out_of_range[:10]}")
    if abs(char_pct) > 10:
        warnings.append(f"字数变化 {char_pct}%，超出 ±10% 口径")
    if all_cites_lost:
        issues.append(f"引用标记丢失 {len(all_cites_lost)} 处：{all_cites_lost[:8]}")
    if all_nums_lost:
        warnings.append(f"数字未原样出现 {len(all_nums_lost)} 处：{all_nums_lost[:8]}")
    if banned_hits:
        issues.append(f"残留主观表述：{sorted(set(banned_hits))}")
    if po["period"] and period_pct > -20:
        warnings.append(f"句号仅减少 {abs(period_pct)}%，提示词要求「减少句号使用」")
    if pn["comma"] <= po["comma"]:
        warnings.append("逗号数未增加，未体现「多用逗号衔接」")

    verdict = "fail" if issues else ("warn" if warnings else "pass")

    out = {
        "source": os.path.abspath(args.docx),
        "checked_at": datetime.now().isoformat(timespec="seconds"),
        "paras_edited": len(per_para),
        "chars": {"old": len(old_text), "new": len(new_text), "pct": char_pct},
        "punct": {
            "period": {"old": po["period"], "new": pn["period"], "pct": period_pct},
            "comma": {"old": po["comma"], "new": pn["comma"]},
            "semicolon": {"old": po["semicolon"], "new": pn["semicolon"]},
        },
        "numbers_lost": all_nums_lost,
        "cites_lost": all_cites_lost,
        "banned_hits": sorted(set(banned_hits)),
        "ai_cliche_left": [w for w in AI_CLICHE if w in new_text],
        "blocked_ids": [b[0] for b in blocked],
        "out_of_range_ids": out_of_range,
        "issues": issues,
        "warnings": warnings,
        "verdict": verdict,
        "per_para": per_para,
    }

    write_json(args.output, out)
    print(f"[check] {verdict.upper()}  改动 {len(per_para)} 段")
    print(f"  字数  {len(old_text)} -> {len(new_text)}  ({char_pct:+.1f}%)")
    print(f"  句号  {po['period']} -> {pn['period']}  ({period_pct:+.1f}%)   逗号 {po['comma']} -> {pn['comma']}")
    for x in issues:
        print(f"  [X] {x}")
    for x in warnings:
        print(f"  [!] {x}")
    print(f"  报告 -> {args.output}")


# ---------------------------------------------------------------- utils

def load_json(path):
    if not os.path.exists(path):
        sys.exit(f"[错误] 文件不存在：{path}")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def write_json(path, obj):
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


def main():
    ap = argparse.ArgumentParser(
        description="论文降 AIGC 率工具四件套：split / batch / write / check"
    )
    sub = ap.add_subparsers(dest="cmd", required=True)

    s1 = sub.add_parser("split", help="docx -> 段落清单 + 章节树 + 分批方案")
    s1.add_argument("docx")
    s1.add_argument("-o", "--output", default="work/blocks.json")
    s1.add_argument("--batch-chars", type=int, default=1200)
    s1.set_defaults(func=cmd_split)

    s2 = sub.add_parser("batch", help="取一批段落，拼成可直接发送的改写请求")
    s2.add_argument("blocks")
    s2.add_argument("--id", type=int, default=None, help="批次号（从 split 结果里取）")
    s2.add_argument("--ids", default=None, help="直接指定段落 id，逗号分隔，如 8,9,10")
    s2.add_argument("--section", default=None, help="按章节标题模糊匹配")
    s2.add_argument("--prompt", default=None, help="改写指令文件，如 references/prompt.md")
    s2.add_argument("-o", "--output", default=None, help="输出到文件（纯请求文本）")
    s2.set_defaults(func=cmd_batch)

    s3 = sub.add_parser("write", help="改写结果 -> 新 docx")
    s3.add_argument("docx")
    s3.add_argument("edits")
    s3.add_argument("-o", "--output", required=True)
    s3.add_argument("--diff-md", default=None)
    s3.add_argument("--force", action="store_true", help="强制改写含超链接的段落（会使链接文字重复，慎用）")
    s3.add_argument("--allow-non-para", action="store_true", help="允许改写标题/图注等非正文段落（默认拒绝）")
    s3.set_defaults(func=cmd_write)

    s4 = sub.add_parser("check", help="改写前后硬指标自检")
    s4.add_argument("docx")
    s4.add_argument("edits")
    s4.add_argument("-o", "--output", default="work/check.json")
    s4.add_argument("--allow-non-para", action="store_true", help="允许把非正文段落计入自检（默认忽略）")
    s4.set_defaults(func=cmd_check)

    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
