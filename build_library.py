"""
重建选题库：
- 导入所有 8 份原始数据
- 与旧 topics.json 合并去重
- 规范化字段
- 写入 data/topics.json
"""
import json
import sys
from pathlib import Path

# 路径
ROOT = Path(__file__).parent
DATA_DIR = ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)

sys.path.insert(0, str(ROOT))
from topic_importer import import_file

# 原始数据文件（按届排）
RAW_FILES = [
    ("/Users/wangguang/Desktop/毕业论文助手/外国语学院2019届毕业论文选题来源（总表）.xls", "2019"),
    ("/Users/wangguang/Desktop/毕业论文助手/外国语学院2020届毕业论文（设计）选题情况汇总表.xls", "2020"),
    ("/Users/wangguang/Desktop/毕业论文助手/外国语学院2021届仲恺农业工程学院毕业设计（论文）汇总表 (10.18).xls", "2021"),
    ("/Users/wangguang/Desktop/毕业论文助手/外国语学院2022届毕业论文汇总表.xlsx", "2022"),
    ("/Users/wangguang/Desktop/毕业论文助手/外国语学院2023届毕业设计（论文）汇总表.xls", "2023"),
    ("/Users/wangguang/Desktop/毕业论文助手/外国语学院2024届本科毕业论文（设计）信息汇总表.xlsx", "2024"),
    ("/Users/wangguang/Desktop/毕业论文助手/外国语学院2025届本科毕业论文（设计）信息汇总表最新修改.xlsx", "2025"),
    ("/Users/wangguang/Desktop/毕业论文助手/外国语学院2026届毕业论文（设计）信息汇总表.xlsx", "2026"),
]


def main():
    all_topics = []
    summary = {}

    # 占位条目（应丢弃）
    PLACEHOLDER_TITLES = {
        "论文题目", "论文关键词", "题目", "选题", "（空）", "(空)",
        "无", "无标题", "", "Title", "title",
    }

    def _is_valid(t: dict) -> bool:
        """过滤占位条目"""
        title = (t.get("title") or "").strip()
        if title in PLACEHOLDER_TITLES:
            return False
        if title.startswith("论文题目"):
            return False
        if len(title) < 4:
            return False
        # 占位关键词
        kws = t.get("keywords") or []
        if kws and all(k.strip() in PLACEHOLDER_TITLES | {"论文关键词", "关键词"} for k in kws):
            return False
        return True

    for path, year in RAW_FILES:
        if not Path(path).exists():
            print(f"⚠️  文件不存在: {path}")
            continue
        topics = import_file(path)
        valid = []
        skipped = 0
        for i, t in enumerate(topics, 1):
            t["id"] = f"{year}-{i:03d}"
            t["title"] = t["title"].replace("\n", " ").replace("\r", " ").strip()
            # 规范化方向
            from topic_importer import normalize_direction
            if not t.get("direction") or t.get("direction") == "其他":
                t["direction"] = normalize_direction(t["title"])
            # 默认语言
            t.setdefault("language", "英语")
            t.setdefault("keywords", [])
            t.setdefault("origin", "")
            t.setdefault("nature", "")
            if _is_valid(t):
                valid.append(t)
            else:
                skipped += 1
        all_topics.extend(valid)
        summary[year] = (len(valid), skipped)
        print(f"✅ {year}届: {len(valid)} 条（已过滤 {skipped} 条占位）")

    # === 合并策略 ===
    # 1) 丢弃旧库中的占位条目
    # 2) 按 (year, title) 键去重，新数据优先（字段更完整）
    existing_file = DATA_DIR / "topics.json"
    if existing_file.exists():
        with open(existing_file, "r", encoding="utf-8") as f:
            old_data = json.load(f)

        # 先过滤旧库的占位
        old_topics = [t for t in old_data.get("topics", []) if _is_valid(t)]
        old_placeholder = len(old_data.get("topics", [])) - len(old_topics)
        print(f"\n旧库: {len(old_data.get('topics', []))} 条，过滤占位 {old_placeholder} 条 → {len(old_topics)} 条")

        # 用 (year, title) 做键，旧 → 新
        by_key: dict[tuple[str, str], dict] = {}
        for t in old_topics:
            key = (t.get("year", "?"), t["title"].strip())
            by_key[key] = t

        # 新数据覆盖同键旧数据
        replaced = 0
        added = 0
        for t in all_topics:
            key = (t.get("year", "?"), t["title"].strip())
            if key in by_key:
                # 新数据字段更完整：keywords/language/direction 都补全
                old_t = by_key[key]
                for k, v in t.items():
                    if v and (not old_t.get(k) or (isinstance(old_t.get(k), list) and len(old_t.get(k)) == 0)):
                        old_t[k] = v
                by_key[key] = old_t
                replaced += 1
            else:
                by_key[key] = t
                added += 1

        final_topics = list(by_key.values())
        print(f"合并: 旧库覆盖 {replaced} 条 + 新增 {added} 条 = {len(final_topics)} 条")
    else:
        final_topics = all_topics

    # ID 重新规范化（保证唯一 YYYY-XXX）
    year_seq: dict[str, int] = {}
    final_topics.sort(key=lambda t: (t.get("year", "9999"), t["title"]))
    for t in final_topics:
        yr = t.get("year", "unknown")
        year_seq[yr] = year_seq.get(yr, 0) + 1
        t["id"] = f"{yr}-{year_seq[yr]:03d}"

    # 方向分布
    from collections import Counter
    dir_count = Counter(t.get("direction", "其他") for t in final_topics)
    print("\n方向分布:")
    for d, c in dir_count.most_common():
        print(f"  {d}: {c}")

    # 年份分布
    year_count = Counter(t.get("year", "?") for t in final_topics)
    print("\n年份分布:")
    for y, c in sorted(year_count.items()):
        print(f"  {y}: {c}")

    # 写入
    data = {
        "metadata": {
            "school": "仲恺农业工程学院外国语学院",
            "major": "商务英语",
            "extracted_from": [Path(p).name for p, _ in RAW_FILES],
            "total": len(final_topics),
            "directions": dict(dir_count),
            "years": dict(year_count),
            "note": "包含 2019-2026 届真实选题，已过滤占位条目，按 (year, title) 去重"
        },
        "topics": final_topics,
    }
    with open(existing_file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"\n✅ 已写入 {existing_file}，共 {len(final_topics)} 条")


if __name__ == "__main__":
    main()
