#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
📝 K Bank 看板 · 一键同步脚本
---------------------------
用法 1（交互式，最常用）：
    python3 sync.py

用法 2（从 JSON 文件批量同步）：
    python3 sync.py import 会议纪要.json

用法 3（先导出当前数据看看有什么）：
    python3 sync.py dump

数据直接写入 Firebase Realtime DB → 所有团队成员立刻看到 ✅
"""

import json
import os
import re
import sys
import urllib.request
import urllib.error
from datetime import datetime

# ===================== Firebase 配置 =====================
# （和 index.html 里的 firebaseConfig 一致，改这里就行）
FIREBASE_DB = "hisun--k-bank-default-rtdb"
API_BASE = f"https://{FIREBASE_DB}.asia-southeast1.firebasedatabase.app"
ITEMS_URL = f"{API_BASE}/kbank-items.json"
DICTS_URL = f"{API_BASE}/kbank-dicts.json"

# 看板条目（从 data.json 同步 seed 出来的）
# —— 如果模块名对不上，用 id 也行（查下面 data.json 的 id 列）
# =========================================================


# ---------- REST 工具 ----------
def fb_get(url):
    req = urllib.request.Request(url)
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read().decode())


def fb_put(url, data):
    body = json.dumps(data, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=body, method="PUT",
                                  headers={"Content-Type": "application/json; charset=utf-8"})
    with urllib.request.urlopen(req) as r:
        return r.status


def fb_patch(url, data):
    """部分更新（只改指定 id 的条目）"""
    body = json.dumps(data, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=body, method="PATCH",
                                  headers={"Content-Type": "application/json; charset=utf-8"})
    with urllib.request.urlopen(req) as r:
        return r.status


# ---------- 加载/保存 ----------
def load_items():
    """从 Firebase 拉全部条目"""
    print("📡  正在从云端拉取当前看板数据...")
    raw = fb_get(ITEMS_URL) or {}
    # Firebase 存的是 { "1": {...}, "2": {...} } 格式，转回 list
    if isinstance(raw, dict):
        items = [v for v in raw.values() if v]
    elif isinstance(raw, list):
        items = [v for v in raw if v]
    else:
        items = []
    print(f"   ✅ 云端共 {len(items)} 条模块")
    return items


# ---------- 模块名匹配（别名 + 模糊打分）----------
# 可选：在项目目录放 module-aliases.json 手动补别名，格式 {"正名片段": ["叫法1","叫法2"]}
ALIAS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "module-aliases.json")


def load_alias_overrides():
    try:
        with open(ALIAS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _name_keys(name):
    """把一个模块名拆成可匹配的候选键（含中文名、英文缩写、去括号全名）"""
    n = (name or "").replace("【里程碑】", "").strip()
    if not n:
        return set()
    keys = {n.lower()}
    # 括号里的英文缩写：Foreign Exchange (EX) 外汇 → ex
    for ab in re.findall(r"[\(（]([A-Za-z0-9/& ]{1,12})[\)）]", n):
        keys.add(ab.strip().lower())
    # 只保留中文：→ 外汇
    zh = re.sub(r"[A-Za-z0-9\(\)（）/&·\-\.]", " ", n)
    zh = re.sub(r"\s+", " ", zh).strip()
    if zh:
        keys.add(zh.lower())
        # 中文里再拆出有意义的词片段（如 "存款 TD" → "存款"）
        for seg in re.split(r"[·/、,，\s]+", zh):
            if len(seg) >= 2:
                keys.add(seg.lower())
    # 去掉括号内容的整串：foreign exchange 外汇
    keys.add(re.sub(r"[\(（].*?[\)）]", "", n).strip().lower())
    return {k for k in keys if k}


def _score(query, key):
    """匹配打分：完全相等最高；包含关系按命中长度"""
    if not query or not key:
        return 0
    if query == key:
        return 1000 + len(key)
    if query in key:
        return 100 + len(query)
    if key in query:
        return 50 + len(key)
    return 0


def find_item(items, name_or_id):
    """按 id / 模块名 / 别名 解析出唯一模块；解析不出返回 None"""
    q = str(name_or_id or "").strip()
    if not q:
        return None

    # 1) 纯数字优先当 id
    if re.fullmatch(r"\d+", q):
        for it in items:
            if it.get("id") == int(q):
                return it

    ql = q.lower()

    # 2) 手动别名覆盖（优先于自动匹配）
    overrides = load_alias_overrides()
    for canonical, aliases in overrides.items():
        cand = {str(canonical).lower(), *[str(a).lower() for a in aliases]}
        if ql in cand:
            for it in items:
                if str(canonical).lower() in (it.get("name") or "").lower():
                    return it

    # 3) 自动打分，取最高分（同分取名字更短的，避免"存款"误配长名）
    best, best_score, best_len = None, 0, 10**9
    for it in items:
        name_len = len(it.get("name") or "")
        for k in _name_keys(it.get("name")):
            s = _score(ql, k)
            if s > best_score or (s == best_score and s > 0 and name_len < best_len):
                best, best_score, best_len = it, s, name_len
    return best if best_score > 0 else None


def push_item(it):
    """把单个条目写回 Firebase"""
    url = f"{API_BASE}/kbank-items/{it['id']}.json"
    status = fb_put(url, it)
    return status == 200


# ---------- 合并逻辑（去重追加）----------
def merge_list(old, new_lines):
    """old = ["已有内容"], new_lines = ["新增1", "新增2"] → 合并去重"""
    if not new_lines:
        return old or []
    old_set = set(x.strip() for x in (old or []) if x and str(x).strip())
    merged = list(old or [])
    for line in new_lines:
        s = str(line).strip()
        if s and s not in old_set:
            merged.append(s)
            old_set.add(s)
    return merged


def merge_minutes(old, new_mins):
    """minutes 是对象数组，按 time+title 去重"""
    if not new_mins:
        return old or []
    old_keys = set((m.get("time") or "") + "|" + (m.get("title") or "")
                   for m in (old or []))
    merged = list(old or [])
    for m in new_mins:
        key = (m.get("time") or "") + "|" + (m.get("title") or "")
        if key not in old_keys:
            merged.append(m)
            old_keys.add(key)
    return merged


def now_time():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


# ---------- 交互模式 ----------
def prompt(prompt_text, default=""):
    if default:
        ans = input(f"{prompt_text} [{default}]: ").strip()
        return ans or default
    return input(f"{prompt_text}: ").strip()


def prompt_list(prompt_text):
    """多行输入，空行结束"""
    print(f"{prompt_text}（每行一条，空行结束）:")
    lines = []
    while True:
        line = input("  └─ ").strip()
        if not line:
            break
        lines.append(line)
    return lines


def interactive_add():
    """交互式：给某个模块添加会议内容"""
    items = load_items()
    if not items:
        print("❌ 云端没有数据，请先确认 Firebase 连接或在浏览器里打开一次看板")
        return

    # 列出所有模块
    print("\n📋 当前模块列表（选一个）:")
    for it in sorted(items, key=lambda x: (x.get("priority", ""), x.get("id", 0))):
        flag = "【里程碑】" if "里程碑" in it.get("name", "") else "         "
        print(f"  {it['id']:>2}. [{it.get('priority','')}] {flag} {it['name']}")
    print()

    target = prompt("输入模块 id 或名称（如 1 或 '外汇'）")
    it = find_item(items, target)
    if not it:
        print(f"❌ 没找到模块: {target}")
        return

    print(f"\n🎯 目标: [{it.get('priority','')}] {it['name']} (id={it['id']})")
    print(f"   当前状态: {it.get('status','plan')}")

    updated = False

    # 1) 会议纪要（对象格式）
    print("\n── 📝 会议纪要（可选）──")
    add_minutes = input("要添加本次会议的纪要吗？(y/N): ").strip().lower()
    if add_minutes == "y":
        title = prompt("纪要标题", f"{it['name']} Workshop")
        user = prompt("记录人", "PMO")
        time_str = prompt("时间", now_time())
        print("纪要内容（空行结束）:")
        content_lines = []
        while True:
            line = input("  └─ ").strip()
            if not line:
                break
            content_lines.append(line)
        content = "\n".join(content_lines)
        if content.strip():
            new_min = {"title": title, "user": user, "time": time_str, "content": content}
            it["minutes"] = merge_minutes(it.get("minutes", []), [new_min])
            updated = True
            print(f"  ✅ 已添加纪要: {title}")

    # 2) 各字段（空行跳过）
    fields = [
        ("confirmed",     "✅ 确认点 / 双方达成一致的点"),
        ("disputes",       "⚠️ 争议点 / 未达成一致需要后续跟进"),
        ("discussions",    "💬 讨论点 / 会议中讨论过的关键点"),
        ("gapConclusion",  "📋 差异需求结论（建议注明 GAP 编号与 FSD 章节）"),
        ("nextStep",       "⏭️ Next Step / 后续要做的事（含责任人 + 截止日更佳）"),
        ("deliverables",   "📦 交付物（文档、原型、代码等）"),
        ("remarks",        "📝 备注 / 风险 / 注意事项"),
    ]
    for key, label in fields:
        print(f"\n── {label}（跳过请直接回车）──")
        lines = prompt_list("输入内容（每行一条）")
        if lines:
            it[key] = merge_list(it.get(key, []), lines)
            updated = True
            print(f"  ✅ {key}: 新增 {len(lines)} 条")

    # 3) 状态更新（可选）
    print(f"\n── 🔄 状态更新（当前: {it.get('status','plan')}）──")
    new_status = prompt("新状态（留空不更新）", "")
    if new_status:
        it["status"] = new_status
        updated = True
        print(f"  ✅ 状态: {it.get('status','plan')} → {new_status}")

    # 4) changelog 审计记录
    if updated:
        changelog_entry = {
            "user": "PMO",
            "time": now_time(),
            "action": "sync.py 批量更新",
        }
        it["changelog"] = (it.get("changelog") or []) + [changelog_entry]

    # 推上去
    if updated:
        print(f"\n🚀  正在推送到 Firebase...")
        # 同时把 items 数组也 PUT 回去（兼容浏览器端读取）
        # 先把 {id: item} 格式准备好
        obj = {str(x["id"]): x for x in items}
        if fb_put(ITEMS_URL, obj) == 200:
            print(f"  ✅ 推送成功！所有团队成员现在就能看到更新了 🎉")
        else:
            print(f"  ❌ 推送失败，请检查网络或 Firebase 规则")
    else:
        print("\n⏭️  没有内容需要更新")


# ---------- import JSON 文件 ----------
FIELD_LABELS = {
    "confirmed": "确认点", "disputes": "争议点", "discussions": "讨论点",
    "gapConclusion": "差异结论", "nextStep": "待办", "deliverables": "交付物",
    "remarks": "备注", "minutes": "会议纪要",
}


def import_json(path, dry_run=False):
    """从 JSON 文件批量导入（纪要提炼结果）
    文件格式（两种都支持）:
      1) 数组: [ {"module": "...", "status": "ws", "add": {...}, "source": "..."} ]
      2) 对象: { "updates": [...], "unmatched": [ {"hint": "...", "content": "..."} ] }
    """
    with open(path, "r", encoding="utf-8") as f:
        raw = json.load(f)
    if isinstance(raw, dict):
        updates = raw.get("updates") or raw.get("items") or []
        unmatched = raw.get("unmatched") or []
    else:
        updates = raw or []
        unmatched = []

    if not isinstance(updates, list):
        print("❌ 文件格式不对：updates 必须是数组")
        return

    items = load_items()
    applied, skipped = 0, []

    print("\n" + "=" * 62)
    print("📋  待入库预览" if dry_run else "📋  开始入库")
    print("=" * 62)

    for u in updates:
        name = u.get("module") or u.get("name") or u.get("id")
        if not name:
            continue
        it = find_item(items, str(name))
        if not it:
            skipped.append(name)
            print(f"  ⚠️  未匹配模块「{name}」→ 已放入未识别清单")
            continue

        # 逐字段统计"新增多少条"（去重后）
        changes = []
        for key, val in (u.get("add") or {}).items():
            if not val:
                continue
            if key == "minutes":
                before = len(it.get("minutes") or [])
                merged = merge_minutes(it.get("minutes", []), val)
                added = len(merged) - before
                if not dry_run:
                    it["minutes"] = merged
            else:
                before = len(it.get(key) or [])
                merged = merge_list(it.get(key), val)
                added = len(merged) - before
                if not dry_run:
                    it[key] = merged
            if added > 0:
                changes.append(f"{FIELD_LABELS.get(key, key)} +{added}")

        status_note = ""
        if u.get("status") and u["status"] != it.get("status"):
            status_note = f"状态 {it.get('status')} → {u['status']}"
            if not dry_run:
                it["status"] = u["status"]

        if not changes and not status_note:
            print(f"  ⏭️  [{it['id']}] {it['name']} 无新增内容（已存在，跳过）")
            continue

        src = u.get("source") or ""
        desc = "、".join(changes + ([status_note] if status_note else []))
        print(f"  {'📝' if dry_run else '✅'} [{it['id']}] {it['name']}  {desc}" + (f"  ← {src}" if src else ""))

        if not dry_run:
            it["changelog"] = (it.get("changelog") or []) + [{
                "user": "PMO", "time": now_time(),
                "action": f"会议纪要入库 {os.path.basename(path)}" + (f"（{src}）" if src else "")
            }]
        applied += 1

    # 未识别内容单独存档，绝不静默丢弃
    if unmatched or skipped:
        review = {"unmatched": unmatched, "skipped_modules": skipped}
        out = "/tmp/kbank-unmatched.json"
        with open(out, "w", encoding="utf-8") as f:
            json.dump(review, f, ensure_ascii=False, indent=2)
        print(f"\n  🔎 未识别内容 {len(unmatched)} 条 / 未匹配模块 {len(skipped)} 个 → 已存 {out}")

    if dry_run:
        print(f"\n👀  预演结束：将更新 {applied} 个模块（未写入云端）")
        print("    确认无误后去掉 --dry-run 再执行一次即可入库")
        return

    if applied:
        obj = {str(x["id"]): x for x in items}
        if fb_put(ITEMS_URL, obj) == 200:
            print(f"\n🎉 成功推送 {applied} 个模块的更新！全员立即可见")
        else:
            print("\n❌ 推送失败")
    else:
        print("\n⏭️  没有新内容需要入库")


# ---------- dump ----------
def dump():
    items = load_items()
    out = "/tmp/kbank-dump.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=2)
    print(f"💾 已导出到 {out}（共 {len(items)} 条）")


# ---------- main ----------
if __name__ == "__main__":
    args = sys.argv[1:]
    try:
        if not args:
            interactive_add()
        elif args[0] == "dump":
            dump()
        elif args[0] == "import" and len(args) > 1:
            dry = "--dry-run" in args or "--preview" in args
            path = [a for a in args[1:] if not a.startswith("--")][0]
            import_json(path, dry_run=dry)
        elif args[0] in ("preview", "check") and len(args) > 1:
            import_json(args[1], dry_run=True)
        else:
            print(__doc__)
    except urllib.error.URLError as e:
        print(f"\n❌ 网络错误: {e}")
        print("   请检查: 1) 是否联网  2) Firebase Database URL 是否正确")
    except urllib.error.HTTPError as e:
        print(f"\n❌ Firebase 返回错误: HTTP {e.code}")
        if e.code == 401 or e.code == 403:
            print("   👉 请到 Firebase Console → Realtime Database → Rules")
            print('      把规则改为: {".read": true, ".write": true}')
    except KeyboardInterrupt:
        print("\n\n👋 已取消")
