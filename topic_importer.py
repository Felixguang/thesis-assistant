"""
选题库导入器
支持 .xls / .xlsx，自动识别各届格式，筛掉非商务英语专业
"""
import re
from pathlib import Path
from typing import Iterator

import openpyxl
import xlrd

# 各届方向字段映射到统一 4 大类（其他为兜底，不算方向）
DIRECTION_KEYWORDS = {
    "翻译": ["翻译", "英译", "汉译", "译本", "归化", "异化", "译介", "字幕", "译法", "创译", "翻译策略", "译配"],
    "跨文化": ["跨文化", "文化差异", "文化维度", "高低语境", "霍夫斯泰德", "Hofstede",
               "中外", "中美", "中俄", "中西方", "中韩", "中英", "文化适应", "文化价值",
               "文化传播", "文化营销", "文化图式", "本土化", "海外", "出海", "异质文化",
               "跨文化交际", "品牌文化", "IP联名", "符号消费", "文化符号", "文化外宣",
               "对外传播", "对外宣传", "对外商务", "跨文化营销", "中国形象",
               "消费者文化", "文化认同", "消费文化", "他者", "文化冲突", "文化距离",
               "文化整合", "文化语境", "岭南", "粤方言", "粤文化", "醒狮", "粿品",
               "中餐厅", "传统节日", "传统民乐", "古典音乐", "中韩校服", "新娘礼服",
               "苏绣", "服饰文化", "建筑风格"],
    "话语分析": ["话语分析", "多模态", "视觉语法", "批评话语", "discourse", "语篇",
                 "系统功能", "人际意义", "人际功能", "语用", "pragmatic", "言语行为",
                 "speech act", "礼貌原则", "合作原则", "面子理论", "元话语", "面子",
                 "模糊限制语", "话语策略", "修辞", "劝说", "说服", "互文性",
                 "评价理论", "态度系统", "语域", "体裁", "接受美学", "传播符号学",
                 "概念隐喻", "符号学", "语码转换", "叙事学", "拟剧论", "戏剧论",
                 "网络热词", "委婉语", "视觉语法", "纪录片", "短视频"],
    "商务英语习得": ["教学", "习得", "学习", "二语", "外语教学", "词汇习得", "词汇学习",
                  "词汇量", "词汇记忆", "教学法", "课程", "教材", "情感过滤",
                  "监控理论", "TPACK", "建构主义", "学习动机", "学习兴趣",
                  "输入假设", "Krashen", "词块学", "语言学习", "写作能力",
                  "口语", "POA", "产出导向法", "高中英语", "小学英语"],
}


def normalize_direction(text: str) -> str:
    """从方向字段或题目推断属于哪个大类"""
    if not text:
        return "其他"
    text_lower = text.lower()
    scores = {}
    for direction, keywords in DIRECTION_KEYWORDS.items():
        scores[direction] = sum(1 for kw in keywords if kw.lower() in text_lower)
    best = max(scores.items(), key=lambda x: x[1])
    return best[0] if best[1] > 0 else "其他"


def is_business_english(major: str, class_name: str = "") -> bool:
    """判断是否商务英语专业"""
    text = f"{major} {class_name}".lower()
    return any(kw in text for kw in ["商英", "商务英语", "business english"])


# ---------- 各届解析器 ----------

def parse_2019(rows: list[list[str]]) -> Iterator[dict]:
    """2019届：序号/学号/姓名/题目/校内指导/...
    用户说明：所有提供的选题都是商务英语专业，全部收
    """
    for row in rows[5:]:
        if len(row) >= 4 and row[3]:
            yield {"title": row[3], "year": "2019"}


def parse_2020(rows: list[list[str]]) -> Iterator[dict]:
    """2020届：序号/专业/题目/性质/来源"""
    for row in rows[3:]:
        if len(row) >= 4 and row[2]:
            yield {
                "title": row[2],
                "year": "2020",
                "nature": row[3] if len(row) > 3 else "",
                "origin": row[4] if len(row) > 4 else "",
                "direction": normalize_direction(row[2]),
            }


def parse_2021(rows: list[list[str]]) -> Iterator[dict]:
    """2021届：序号/专业班级/题目/性质/来源"""
    for row in rows[3:]:
        if len(row) >= 4 and row[2]:
            yield {
                "title": row[2],
                "year": "2021",
                "nature": row[3] if len(row) > 3 else "",
                "origin": row[4] if len(row) > 4 else "",
                "direction": normalize_direction(row[2]),
            }


def parse_2022(wb) -> Iterator[dict]:
    """2022届.xlsx：序号/专业班级/题目/性质/来源"""
    ws = wb.active
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not row or not row[2]:
            continue
        yield {
            "title": str(row[2]).strip(),
            "year": "2022",
            "nature": str(row[3] or "").strip(),
            "origin": str(row[4] or "").strip(),
            "direction": normalize_direction(str(row[2])),
        }


def parse_2023(rows: list[list[str]]) -> Iterator[dict]:
    """2023届：序号/题目/性质/来源（无专业列，全部纳入）"""
    for row in rows[3:]:
        if len(row) >= 2 and row[1]:
            yield {
                "title": row[1],
                "year": "2023",
                "nature": row[2] if len(row) > 2 else "",
                "origin": row[3] if len(row) > 3 else "",
                "direction": normalize_direction(row[1]),
            }


def parse_2024_2025(wb, year: str) -> Iterator[dict]:
    """2024/2025届：序号/专业名称/题目/语种/关键词/研究方向/来源/性质"""
    ws = wb.active
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not row or len(row) < 3:
            continue
        idx, major, title, lang, keywords, direction, origin, nature = (
            row[0], row[1], row[2], row[3] if len(row) > 3 else "",
            row[4] if len(row) > 4 else "",
            row[5] if len(row) > 5 else "",
            row[6] if len(row) > 6 else "",
            row[7] if len(row) > 7 else "",
        )
        if not title or not str(title).strip():
            continue
        kw_list = [k.strip() for k in str(keywords).replace("；", ";").split(";") if k.strip()]
        final_dir = str(direction).strip() if direction else ""
        if final_dir not in ["翻译", "跨文化", "话语分析", "商务英语习得"]:
            final_dir = normalize_direction(str(title) + " " + " ".join(kw_list))
        yield {
            "title": str(title).strip(),
            "year": year,
            "keywords": kw_list,
            "language": str(lang).strip() if lang else "英语",
            "direction": final_dir,
            "origin": str(origin).strip() if origin else "",
            "nature": str(nature).strip() if nature else "",
        }


def parse_2026(wb) -> Iterator[dict]:
    """2026届：学院/专业/题目/关键词/来源/研究方向/语种（全部纳入）"""
    ws = wb["论文（设计）信息汇总表"] if "论文（设计）信息汇总表" in wb.sheetnames else wb.active
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not row or not row[3]:
            continue
        title = str(row[3] or "").strip()
        if not title:
            continue
        keywords = str(row[4] or "")
        kw_list = [k.strip() for k in keywords.replace("；", ";").split(";") if k.strip()]
        direction = str(row[6] or "").strip() if len(row) > 6 else ""
        lang = str(row[7] or "英语").strip() if len(row) > 7 else "英语"
        if direction not in ["翻译", "跨文化", "话语分析", "商务英语习得"]:
            direction = normalize_direction(title + " " + " ".join(kw_list))
        yield {
            "title": title,
            "year": "2026",
            "keywords": kw_list,
            "language": lang,
            "direction": direction,
            "origin": str(row[5] or "").strip() if len(row) > 5 else "",
            "nature": "",
        }


# ---------- 主入口 ----------

def read_xls_rows(path: str) -> list[list[str]]:
    wb = xlrd.open_workbook(path)
    sh = wb.sheet_by_index(0)
    rows = []
    for r in range(sh.nrows):
        rows.append([str(sh.cell_value(r, c)).strip() for c in range(sh.ncols)])
    return rows


def read_xlsx_wb(path: str):
    return openpyxl.load_workbook(path, data_only=True)


def import_file(path: str) -> list[dict]:
    """根据文件名识别年份并解析，返回商务英语选题列表"""
    name = Path(path).name
    rows = []
    if path.endswith(".xls"):
        rows = read_xls_rows(path)
    elif path.endswith(".xlsx"):
        wb = read_xlsx_wb(path)
    else:
        return []

    if "2019" in name:
        return list(parse_2019(rows))
    elif "2020" in name:
        return list(parse_2020(rows))
    elif "2021" in name:
        return list(parse_2021(rows))
    elif "2022" in name:
        return list(parse_2022(wb))
    elif "2023" in name:
        return list(parse_2023(rows))
    elif "2024" in name:
        return list(parse_2024_2025(wb, "2024"))
    elif "2025" in name:
        return list(parse_2024_2025(wb, "2025"))
    elif "2026" in name:
        return list(parse_2026(wb))
    return []


def merge_with_existing(existing: list[dict], new_topics: list[dict]) -> list[dict]:
    """合并去重：相同 title 不重复加"""
    existing_titles = {t["title"].strip() for t in existing}
    merged = list(existing)
    added = 0
    for t in new_topics:
        title = t["title"].strip()
        if title and title not in existing_titles:
            existing_titles.add(title)
            merged.append(t)
            added += 1
    return merged, added


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("用法: python topic_importer.py <文件1> [文件2 ...]")
        sys.exit(1)
    total = 0
    for f in sys.argv[1:]:
        topics = import_file(f)
        print(f"{Path(f).name}: 解析 {len(topics)} 条商务英语选题")
        if topics:
            print(f"  示例: {topics[0]['title'][:60]}")
        total += len(topics)
    print(f"\n总计: {total} 条")
