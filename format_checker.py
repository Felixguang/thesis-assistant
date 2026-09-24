"""
格式检查模块 v0.4
依据：仲恺农业工程学院外国语学院毕业论文（设计）格式规范
- 封面/扉页/承诺书 = 前置独立区（不按正文规范）
- 正文 = 宋体(中)/Times New Roman(英) 小四(12pt) 1.5倍行距
- 一级标题 = Times New Roman 四号(14pt) 加粗
- 图表 = Times New Roman 五号(10.5pt)
- 参考文献 = [序号] 作者. 题名[J]. 出版地: 出版者, 出版年. 起止页码
- 附录/致谢 = 同正文（小四 1.5倍行距）
- 页面设置 = 上下左右 2.4cm, 1.5 倍行距, 页码居中, 摘要/目录不标正文页码
"""
import re
from collections import Counter, OrderedDict
from pathlib import Path
from typing import Optional, Tuple, List, Dict

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

# === 字号映射 ===
FONT_SIZE = {
    "初号": 42, "小初": 36,
    "一号": 26, "小一": 24,
    "二号": 22, "小二": 18,
    "三号": 16, "小三": 15,
    "四号": 14, "小四": 12,
    "五号": 10.5, "小五": 9,
}

# === 预编译正则（热路径，避免重复 JIT） ===
_RE_TOC_ENTRY = re.compile(r"\t\s*PAGEREF|\t\s*\d+$|\s+\d{1,3}\s*$")
_RE_NOTE_PREFIX = re.compile(r"^注[：:]\s*1\.")
_RE_GRADE_QUOTE = re.compile(r"^[“”]等级[“”][：:]")
_RE_ABSTRACT_ZH_SPACED = re.compile(r"^摘\s+要")
_RE_ABSTRACT_ZH = re.compile(r"^摘要")
_RE_ABSTRACT_EN = re.compile(r"^Abstract", re.I)
_RE_TOC_EN = re.compile(r"^Table of Contents", re.I)
_RE_TOC_ZH_SPACED = re.compile(r"^目\s+录")
_RE_TOC_ZH = re.compile(r"^目录")
_RE_REF_ZH = re.compile(r"^参考文献")
_RE_REF_EN = re.compile(r"^Bibliography", re.I)
_RE_APPENDIX = re.compile(r"^\s*Appendix(\s+[A-Z]|\s*[:：]|\s*$|\s+[A-Za-z一-鿿])")
_RE_APPENDIX_SOLO = re.compile(r"^\s*Appendix\s*$")
_RE_ACK_ZH = re.compile(r"^致谢")
_RE_ACK_EN = re.compile(r"^Acknowledgement", re.I)
_RE_EXAMPLE = re.compile(r"^\s*Example\s+\d+\s*[:：]")
_RE_DIAGRAM = re.compile(r"^\s*(Diagram|Figure|图|Picture|Example)\s*[\dA-Z]", re.I)
_RE_TABLE = re.compile(r"^\s*(Table|表)\s*\d+", re.I)
_RE_REF_TYPE_TAG = re.compile(r"\[(M|J|C|D|R|N|S|P|A|EB/OL|DB/CD|Z)\]")
_RE_HEADING_NUM = re.compile(r"^\s*(\d+(?:\.\d+){0,2})\s+[一-龥A-Za-z]")
_RE_PAGEREF_TAB = re.compile(r"\t\s*PAGEREF\b")
_RE_CHINESE_CHAR = re.compile(r"[一-鿿]")

# === 期望值（来自规范文档 + 优秀论文范文 + 规范示例截图）===
EXPECTED = {
    "body_font_zh": "宋体",
    "body_font_en": "Times New Roman",
    "body_size_pt": 12,       # 小四
    "h1_size_pt": 14,         # 四号
    "h2_size_pt": 12,         # 小四
    "h3_size_pt": 12,         # 小四
    "h1_bold": True,
    "h2_bold": True,
    "h3_bold": False,
    "h1_font": "Times New Roman",
    "h2_font": "Times New Roman",
    "h3_font": "Times New Roman",
    "line_spacing": 1.5,
    "page_margin_cm": 2.4,
    "caption_size_pt": 10.5,  # 五号
    "caption_font": "Times New Roman",
    "appendix_label_pt": 14,  # 四号
    # 摘要标题：黑体 14pt 加粗居中（规范文档 + 示例截图）
    "abstract_title_size_pt": 14,
    "abstract_title_font": "黑体",
    # 关键词标签：黑体 小四 左对齐
    "keyword_label_font": "黑体",
    "keyword_label_size_pt": 12,
    # 关键词内容：宋体 小四，分隔符为中文分号
    "keyword_content_font": "宋体",
    "keyword_content_size_pt": 12,
    "keyword_separator": "；",
    "keyword_min": 3,
    "keyword_max": 5,
    # 英文摘要标题
    "abstract_en_header_size_pt": 16,   # 三号
    "abstract_en_header_bold": True,
    "abstract_en_header_font": "Times New Roman",
    # 目录
    "toc_header_font": "Times New Roman",
    "toc_header_size_pt": 14,           # 四号
    "toc_header_bold": True,
    "toc_header_align_center": True,
    "toc_body_font": "Times New Roman",
    "toc_body_size_pt": 12,             # 小四
    # 参考文献
    "bib_size_pt": 10.5,                # 五号
    "bib_font_en": "Times New Roman",
    "bib_font_zh": "宋体",
    "bib_header_size_pt": 14,           # 四号
    "bib_header_bold": True,
    "bib_header_align_center": True,
    # 扉页（中英文对照的英文扉页）
    "frontpage_required_fields": [
        "Supervisor",
        "Zhongkai University of Agriculture and Engineering",
        "Guangzhou",
    ],
}

# === 严重度颜色映射（用于 UI 渲染）===
SEVERITY_COLOR = {
    "高": "#E63946",   # 红
    "中": "#F4A261",   # 橙
    "低": "#2A9D8F",   # 青
}


def _size_to_pt(size_val) -> Optional[float]:
    if size_val is None:
        return None
    try:
        return float(size_val.pt)
    except Exception:
        return None


def _pt_to_zh(pt: float) -> str:
    if pt is None:
        return "?"
    closest = min(FONT_SIZE.items(), key=lambda kv: abs(kv[1] - pt))
    return closest[0]


# === 区域识别 ===

def _detect_zones(doc) -> Dict[str, Tuple[int, int]]:
    """根据内容识别文档各区域。
    兼容段落和表格（封面常用表格布局）。
    返回 {区域名: (start_idx, end_idx_exclusive)}，索引基于所有 body 子元素。
    """
    body_children = list(doc.element.body)
    n = len(body_children)
    # 同时收集段落（按段落索引）和全文扫描
    paragraphs = doc.paragraphs
    zones: Dict[str, list] = {}
    boundaries = []  # (body_child_idx, zone_name)

    def _classify(text: str) -> Optional[str]:
        text = text.strip()
        if not text:
            return None
        # 目录项特征：包含 tab + 数字（页码），排除之
        is_toc_entry = bool(_RE_TOC_ENTRY.search(text))
        if is_toc_entry:
            return None
        # 成绩评定表注释：开头"注：1."或"等级"
        if _RE_NOTE_PREFIX.match(text) or _RE_GRADE_QUOTE.match(text):
            return "成绩评定表"
        head = text[:15]
        if "学生承诺书" in head and len(text) < 100:
            return "学生承诺书"
        if _RE_ABSTRACT_ZH_SPACED.match(text) or _RE_ABSTRACT_ZH.match(text):
            return "中文摘要"
        if _RE_ABSTRACT_EN.match(text) and len(text) < 100:
            return "英文摘要"
        if (_RE_TOC_EN.match(text) or _RE_TOC_ZH_SPACED.match(text) or _RE_TOC_ZH.match(text)) and len(text) < 100:
            return "目录"
        if (_RE_REF_ZH.match(text) or _RE_REF_EN.match(text)) and len(text) < 100:
            return "参考文献"
        # Appendix 标题：单独成段、正文、或带编号 "Appendix A: ..."
        if _RE_APPENDIX_SOLO.match(text) or _RE_APPENDIX.match(text):
            if len(text) < 100:
                return "附录"
        if _RE_ACK_ZH.match(text) or _RE_ACK_EN.match(text) and len(text) < 100:
            return "致谢"
        return None

    # 扫描每个 body 子元素
    for ci, child in enumerate(body_children):
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            text = ''.join(child.itertext())
            name = _classify(text)
            if name:
                boundaries.append((ci, name))
        elif tag == 'tbl':
            text = ''.join(child.itertext())
            name = _classify(text)
            if name:
                boundaries.append((ci, name))

    if boundaries:
        first = boundaries[0][0]
        zones["封面/扉页"] = [(0, first)]
        for j, (idx, name) in enumerate(boundaries):
            end = boundaries[j + 1][0] if j + 1 < len(boundaries) else n
            zones.setdefault(name, []).append((idx, end))
    else:
        zones["封面/扉页"] = [(0, n)]

    merged = {}
    for name, spans in zones.items():
        merged[name] = (spans[0][0], spans[-1][1])
    return merged


def _zone_of(body_idx: int, zones: Dict[str, Tuple[int, int]]) -> str:
    for name, (s, e) in zones.items():
        if s <= body_idx < e:
            return name
    return "未知"


def _get_effective_font(run, paragraph, has_chinese: Optional[bool] = None) -> Optional[str]:
    """获取 run 的有效字体。
    如果段落含中文字符，看 eastAsia（中文）；否则看 ascii/hAnsi（英文）。
    因为优秀论文设置惯例：ascii=Times New Roman / eastAsia=宋体 / hAnsi=Times New Roman

    has_chinese=None 时自动检测（按 run 文本是否含中文）。
    当 run 级 ascii/hAnsi 缺失时，会查段落 pPr 继承；都缺时查 pStyle 样式定义。
    """
    from docx.oxml.ns import qn

    # 自动检测：run 文本含中文 → True
    if has_chinese is None:
        has_chinese = bool(run.text) and any('一' <= ch <= '鿿' for ch in run.text)

    def _read_rFonts(rFonts):
        """从 rFonts 元素读对应字体（按 has_chinese 决定 ascii/eastAsia 优先级）。
        缺失字段返回 None，调用方决定是否回退到样式。"""
        if has_chinese:
            return rFonts.get(qn('w:eastAsia')) or rFonts.get(qn('w:ascii')) or rFonts.get(qn('w:hAnsi'))
        else:
            return rFonts.get(qn('w:ascii')) or rFonts.get(qn('w:hAnsi'))

    # 1. run 级 rFonts
    rPr = run._element.find(qn('w:rPr'))
    if rPr is not None:
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is not None:
            f = _read_rFonts(rFonts)
            if f:
                return f

    # 2. python-docx 解析的 font.name（仅当 ascii 非空时返回）
    if run.font.name:
        return run.font.name

    # 3. 段落 pPr.rFonts
    if paragraph is not None:
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            p_rPr = pPr.find(qn('w:rPr'))
            if p_rPr is not None:
                rFonts = p_rPr.find(qn('w:rFonts'))
                if rFonts is not None:
                    f = _read_rFonts(rFonts)
                    if f:
                        return f

            # 4. pStyle 样式定义里的 rFonts（如 toc 1 / heading 1 等）
            pStyle = pPr.find(qn('w:pStyle'))
            if pStyle is not None:
                style_id = pStyle.get(qn('w:val'))
                if style_id:
                    try:
                        style = paragraph.part.document.styles[style_id]
                        s_rPr = style.element.find(qn('w:rPr'))
                        if s_rPr is not None:
                            rFonts = s_rPr.find(qn('w:rFonts'))
                            if rFonts is not None:
                                f = _read_rFonts(rFonts)
                                if f:
                                    return f
                    except (KeyError, AttributeError):
                        pass
    return None


def _get_effective_size(run, paragraph) -> Optional[float]:
    """获取 run 的有效字号（pt）。
    优先 paragraph.pPr.rPr.sz（Word 通常把"官方"字号放在这里），
    其次 run.font.size，最后 run 内子元素的 sz。
    """
    from docx.oxml.ns import qn
    if paragraph is not None:
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            p_rPr = pPr.find(qn('w:rPr'))
            if p_rPr is not None:
                sz_el = p_rPr.find(qn('w:sz'))
                if sz_el is not None:
                    val = sz_el.get(qn('w:val'))
                    if val:
                        try:
                            return float(val) / 2
                        except Exception:
                            pass
    sz = _size_to_pt(run.font.size)
    if sz is not None:
        return sz
    if run is not None:
        rPr = run._element.find(qn('w:rPr'))
        if rPr is not None:
            sz_el = rPr.find(qn('w:sz'))
            if sz_el is not None:
                val = sz_el.get(qn('w:val'))
                if val:
                    try:
                        return float(val) / 2
                    except Exception:
                        pass
    return None


def _first_nonempty_run(paragraph):
    """返回段落中第一个**有文本**的 run，跳过空 run。
    Word 在段落开头常留一个空 run（如占位），其 font.size 可能与正文 run 不一致。
    用空 run 做字号/字体判定会产生误报。"""
    for r in paragraph.runs:
        if (r.text or "").strip():
            return r
    # 没有非空 run 时回退到第一个（保底）
    return paragraph.runs[0] if paragraph.runs else None


# === 标题 Title Case 检查（批次 2 / 3）===

_TITLE_CASE_STOPWORDS = {
    "a", "an", "the", "and", "or", "but", "nor",
    "in", "on", "at", "to", "for", "of", "with", "by", "from",
    "as", "into", "onto", "upon", "via", "vs",
    "between", "among", "over", "under", "within", "without",
}


def _check_title_case(text: str) -> Optional[str]:
    """检查英文标题 Title Case：剔除虚词后，实词首字母必须大写。
    返回 None 表示合规；返回 str 表示违规描述。
    仅适用于英文标题（包含 A-Z 且不含中文）。"""
    has_chinese = any('一' <= ch <= '鿿' for ch in text)
    if has_chinese:
        return None  # 中文标题不检查 Title Case
    # 切词：保留 hyphenated 词（如 "Cross-cultural"）和缩写（如 "'s"）
    # 用 [\s]+ 切词，避免按 hyphen 切碎
    tokens = re.split(r"\s+", text.strip())
    words = []
    for t in tokens:
        # 跳过纯标点
        if not re.search(r"[A-Za-z]", t):
            continue
        # 取每个 token 的"首字母单词"（包含 hyphen 的复合词）
        # 例如 "Cross-cultural" → 整体作为一个词
        first_word_match = re.match(r"([A-Za-z][A-Za-z]*)", t)
        if first_word_match:
            words.append(first_word_match.group(1))
    if len(words) < 2:
        return None  # 单词不检查
    bad = []
    for w in words:
        wl = w.lower()
        if wl in _TITLE_CASE_STOPWORDS:
            continue  # 虚词跳过
        # 实词首字母必须大写
        if not w[0].isupper():
            bad.append(w)
    if bad:
        return f"实词未大写：{', '.join(bad)}"
    return None


def _has_page_break_before(paragraph) -> bool:
    """检查段落前是否有分页符（<w:br type='page'/> 在前段末尾，
    或段落属性中有 <w:pageBreakBefore/>）。"""
    pPr = paragraph._p.find(qn("w:pPr"))
    if pPr is not None:
        if pPr.find(qn("w:pageBreakBefore")) is not None:
            return True
    prev = paragraph._p.getprevious()
    if prev is not None:
        for br in prev.iter(qn("w:br")):
            if br.get(qn("w:type")) == "page":
                return True
    return False


def _has_visual_break_before(paragraph) -> bool:
    """检查段落前是否有 ≥1 个完全空段（用作视觉分隔）。

    真实论文常靠段前空行模拟"另起一页"，不算正式分页但能接受。
    """
    count = 0
    prev = paragraph._p.getprevious()
    while prev is not None and count < 5:
        # 空段：text 为空且没有 page break
        text = "".join(prev.itertext()).strip()
        has_br = any(br.get(qn("w:type")) == "page" for br in prev.iter(qn("w:br")))
        if not text and not has_br:
            count += 1
            prev = prev.getprevious()
            continue
        break
    return count >= 1


def _is_run_superscript(run) -> bool:
    """检查 run 是否为上标（<w:vertAlign w:val='superscript'/>）。"""
    rPr = run._r.find(qn("w:rPr"))
    if rPr is None:
        return False
    valign = rPr.find(qn("w:vertAlign"))
    if valign is None:
        return False
    return valign.get(qn("w:val")) == "superscript"


def _is_paragraph_centered(paragraph) -> bool:
    """检查段落是否居中对齐（含样式继承）。"""
    try:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        if paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER:
            return True
        # 查样式继承
        pPr = paragraph._p.find(qn("w:pPr"))
        if pPr is not None:
            pStyle = pPr.find(qn("w:pStyle"))
            if pStyle is not None:
                style_id = pStyle.get(qn("w:val"))
                # 查 doc.styles
                try:
                    style = paragraph.part.document.styles[style_id]
                    if style.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                        return True
                except (KeyError, AttributeError):
                    pass
    except Exception:
        pass
    return False


def _is_run_bold_inherited(run, paragraph) -> Optional[bool]:
    """检查 run 是否加粗（含样式继承）。返回 True/False/None（不确定）。"""
    if run.font.bold is True:
        return True
    if run.font.bold is False:
        # 查样式继承
        pPr = paragraph._p.find(qn("w:pPr"))
        if pPr is not None:
            pStyle = pPr.find(qn("w:pStyle"))
            if pStyle is not None:
                style_id = pStyle.get(qn("w:val"))
                try:
                    style = paragraph.part.document.styles[style_id]
                    if style.font.bold is True:
                        return True
                except (KeyError, AttributeError):
                    pass
        return False
    return None  # 不确定


def _is_example_caption(text: str) -> bool:
    """匹配 Example N: 或 Example N：开头的例证标题。"""
    return bool(_RE_EXAMPLE.match(text))


# === 参考文献类型正则（批次 5）===
# 支持两种格式：
# A. GB/T 7714（带 [序号] 前缀）：[1]作者. 题名[M]. 出版地: 出版者, 出版年.
# B. ISO 690（无 [序号] 前缀）：作者. 题名[M]. 出版地: 出版者, 出版年.
# 正则只校验最关键的字段（类型标识 + 年份 + 出版地），避免误判。

REF_TYPE_PATTERNS = {
    # 专著 [M]：类型标识存在即可（GB/T 7714 中年份在条目尾部，不强制紧跟 [M]）
    "M":   r"\[M\]",
    # 期刊 [J]
    "J":   r"\[J\]",
    # 论文集 [C]
    "C":   r"\[C\]",
    # 学位论文 [D]
    "D":   r"\[D\]",
    # 报告 [R]
    "R":   r"\[R\]",
    # 报纸 [N]
    "N":   r"\[N\]",
    # 标准 [S]
    "S":   r"\[S\]",
    # 专利 [P]
    "P":   r"\[P\]",
    # 论文集析出 [A]
    "A":   r"\[A\]",
    # 电子文献 [EB/OL]
    "EB/OL": r"\[EB/OL\]",
    # 光盘 [DB/CD]
    "DB/CD": r"\[DB/CD\]",
    # 其他 [Z]
    "Z":   r"\[Z\]",
}


def _get_reference_type(text: str) -> Optional[str]:
    """从参考文献条目中提取类型标识（如 [M]/[J]/[EB/OL]）。"""
    m = _RE_REF_TYPE_TAG.search(text)
    if m:
        return m.group(1)
    return None


def _consensus_size(paragraph) -> Optional[float]:
    """取段落中**所有有文本的 run**的字号，按"多数决"决定。
    Word 在不同章节会留下不一致的元 run（比如段落级 pPr.rPr.sz 是 12pt，
    但实际标题 run 是 14pt；或反过来）。用 run-level 多数决比逐 run 级优先/段落级优先都稳。

    若所有有文本的 run 都未显式设置 font.size（继承自样式/默认），
    返回 None —— 不要回退到 pPr.rPr.sz 或 docDefaults，
    否则会误把"未设置"当成"小四（12pt）"。
    """
    from collections import Counter
    sizes: list = []
    for r in paragraph.runs:
        if not (r.text or "").strip():
            continue
        sz = _size_to_pt(r.font.size)
        if sz is not None:
            sizes.append(round(sz, 2))
    if not sizes:
        return None  # 不再回退到段落级，避免误判
    c = Counter(sizes)
    return c.most_common(1)[0][0]


# === 各级别判断 ===

def _detect_heading_level(text: str) -> Optional[int]:
    """根据文本模式判断章节层级：1 / 1.1 / 1.1.1"""
    m = _RE_HEADING_NUM.match(text)
    if not m:
        return None
    return m.group(1).count(".") + 1


def _is_reference_entry(text: str) -> bool:
    """判断是否为参考文献条目。
    仲恺规范支持两种格式：
      A. [序号] 作者. 题名[J/R/M/D]. 出版地: 出版者, 出版年. 起止页码
      B. 作者. 题名[J/R/M/D]. 出版地: 出版者, 出版年. 起止页码（无序号，ISO 690）
    """
    # 格式 A：[序号] 作者. 题名[文献类型]
    if re.match(r"^\s*\[\d+\]\s*[^.]+(\s+et al\.?|，\s*等)?\.\s*[^[]+\s*\[[A-Z]+\]", text):
        return True
    if re.match(r"^\s*\[\d+\]\s*.+\[[A-Z]+\]", text):
        return True
    # 格式 B：作者. 题名[文献类型]. 出版地/刊名...
    # 典型模式：<作者>. <题名>[J/M/R/D]. <出版地或刊名>...
    if re.match(r"^\s*([^.\n]{2,80})\.\s*([^.\n]{2,200}\])\.\s*", text):
        if re.search(r"\[[A-Z]\]", text):
            # 排除"附录"等误判（要满足题名长度）
            parts = text.split(".")
            if len(parts) >= 3 and len(parts[1].strip()) > 5:
                return True
    return False


def _is_figure_or_table_caption(text: str) -> Tuple[bool, Optional[str]]:
    """判断是否为图题/表题，返回 (是否, 类型)"""
    if _RE_DIAGRAM.match(text):
        return True, "图题"
    if _RE_TABLE.match(text):
        return True, "表题"
    return False, None


# === v2 识别：基于 Word 段落样式 + 文本特征 ===
# 兼容中英文样式名，比较前统一 lower()
_TOC_STYLE_NAMES = {
    "toc 1", "toc 2", "toc 3", "toc 4", "toc 5", "toc heading",
    "table of contents", "目录 1", "目录 2", "目录 3", "目录 4", "目录标题",
}
_HEADING_STYLE_NAMES = {
    "heading 1", "heading 2", "heading 3", "heading 4",
    "标题 1", "标题 2", "标题 3", "标题 4",
}

_HEADING_TEXT_RE = re.compile(r"^\s*(\d{1,2}(?:\.\d{1,2}){0,3})\s+\S")

# 仲恺商英规范约束（用户确认）：
# - L1 数字只能是 1-4（论文一共有 4 个一级标题：引言/理论方法/结果讨论/结论）
# - L2 数字段（主编号）只能是 1-9
# - L3 数字段（主编号）只能是 1-9
# 例如 "50 undergraduate students..."（主编号 50 > 4）即使符合标题正则也不算 L1。
_HEADING_NUM_LIMITS = {1: (1, 4), 2: (1, 9), 3: (1, 9)}


def _para_style_name(para) -> str:
    """容错读取段落样式名。"""
    if para is None:
        return ""
    try:
        return (para.style.name or "").strip()
    except Exception:
        return ""


def _is_toc_paragraph(para) -> bool:
    """判断是否为 TOC 条目。优先看样式名，否则看 tab + 页码 / PAGEREF。"""
    if para is None:
        return False
    sn = _para_style_name(para).lower()
    if sn in _TOC_STYLE_NAMES:
        return True
    text = para.text or ""
    if _RE_PAGEREF_TAB.search(text):
        return True
    # TOC 条目典型形式："...\t<页号>" 或 "...\t<页号>\n"
    if re.search(r"\t\s*\d{1,3}\s*$", text.rstrip()):
        return True
    return False


def _detect_headings(paragraphs) -> List[Dict]:
    """识别正文章节标题。三层策略：
    1. Word 样式名（heading 1/2/3/4 或 标题 1/2/3/4）
    2. 阿拉伯数字编号正则（1、1.1、1.1.1）
    3. TOC 模糊匹配兜底：短段落 Normal 样式，文本与 TOC 条目的标题部分相似
       （应对标题丢失编号、纯文本仅有 "Theory and Methodology" 等情况）

    返回 [{"level": 1|2|3|4, "idx": 段落号, "text": 截断后文本}, ...]
    """
    # 第一遍：扫描 TOC，提取 L1/L2/L3 标题文本（去掉编号和页码）
    toc_titles: Dict[int, List[str]] = {1: [], 2: [], 3: []}
    for para in paragraphs:
        if not _is_toc_paragraph(para):
            continue
        text = (para.text or "").strip()
        # 去掉尾部的 \t<页码>
        text = re.sub(r"\t\s*\d+\s*$", "", text).strip()
        # 提取编号前缀
        m = re.match(r"^\s*(\d+(?:\.\d+){0,2})\s+(.+)$", text)
        if not m:
            continue
        num = m.group(1)
        title = m.group(2).strip()
        dots = num.count(".")
        lvl = min(dots + 1, 3)
        toc_titles.setdefault(lvl, []).append(title)

    # 规范化函数：去空白 + lowercase，便于模糊匹配
    def _norm(s: str) -> str:
        return re.sub(r"\s+", "", s).lower()

    # 预计算规范化后的 TOC 标题集（按 level）
    toc_norm: Dict[int, set] = {
        lvl: {_norm(t) for t in titles} for lvl, titles in toc_titles.items()
    }

    out: List[Dict] = []
    for i, para in enumerate(paragraphs):
        text = (para.text or "").strip()
        if not text:
            continue
        sn = _para_style_name(para).lower()
        level = None
        tail_title = None  # 末尾提取的标题文本（如 "4 Conclusion"）

        # 1) 文本数字编号（最权威：1 / 1.1 / 1.1.1）
        #    Word 有时把 1/1.1/1.1.1 全部用 Heading 1 样式，但编号才是真正的级别。
        #    守卫：
        #    - 标题通常 ≤ 80 字符；超过则视为正文段（防 "50 undergraduate students..." 被误判为 L1）
        #    - 主编号按级别有范围限制（L1 1-4, L2/L3 1-9，符合仲恺商英规范）
        m_num = _HEADING_TEXT_RE.match(text)
        if m_num and not _is_toc_paragraph(para) and len(text) <= 80:
            num_str = m_num.group(1)
            dots = num_str.count(".")
            level = min(dots + 1, 3)
            # 检查主编号是否在合法范围内
            main_num = int(num_str.split(".")[0])
            limit_min, limit_max = _HEADING_NUM_LIMITS.get(level, (1, 99))
            if not (limit_min <= main_num <= limit_max):
                level = None  # 超出范围（如 "50 undergraduate..."），不算章节标题

        # 1b) 兜底：数字编号出现在段落末尾（典型场景：作者漏按 Enter，
        #     "4 Conclusion" 被拼到上一段正文末尾 "...cognitive load.4  Conclusion"）。
        #     触发条件：段落中除末尾标题外的部分是合法正文（>40 字符，且以句号/引号结尾），
        #     且末尾是 "数字  Title" 模式。
        if level is None and not _is_toc_paragraph(para) and len(text) > 50:
            # 找出最后一个 "\d+(\.\d+){0,2}\s+\S+..." 结尾
            m_end = re.search(r"(\d{1,2}(?:\.\d{1,2}){0,2})\s+([A-Za-z一-鿿][A-Za-z0-9\s,.\-:;'一-鿿]{2,60})$", text)
            if m_end:
                # 标题前缀必须是 "...。" "." 之类的句子结束符
                title_start_idx = m_end.start()
                prefix = text[:title_start_idx].rstrip()
                if len(prefix) > 30 and re.search(r"[。.!?！」』”\"']\s*$", prefix):
                    # 这是作者漏按 Enter 的笔误，不应作为有效章节标题。
                    # 后续 check_document 会单独检测"作者漏按 Enter"问题。
                    continue  # 跳过此段（不记录为标题）

        # 2) Word 样式名兜底（仅在没数字编号时使用）
        #    仅接受带 heading/标题 字样的样式，不接受 Normal
        #    守卫：跳过 back-matter 区域头（Bibliography / Acknowledgements 等）
        if level is None and re.search(r"heading|标题", sn):
            text_lc = text.lower().strip()
            is_back_matter_head = (
                text_lc in {"bibliography", "references"}
                or re.match(r"^致\s*谢", text) is not None
                or re.match(r"^acknowledg", text_lc) is not None
                or re.match(r"^appendix", text_lc) is not None
            )
            if not is_back_matter_head:
                if "1" in sn or sn.endswith(" 1"):
                    level = 1
                elif "2" in sn or sn.endswith(" 2"):
                    level = 2
                elif "3" in sn or sn.endswith(" 3"):
                    level = 3
                elif "4" in sn or sn.endswith(" 4"):
                    level = 4

        # 3) 短段落 Normal 样式 + TOC 标题模糊匹配（兜底：标题丢失编号）
        #    守卫：
        #    - 跳过 back-matter 区域头（Bibliography / Acknowledgements / 致谢 / Appendix 等）
        #    - 跳过图表标题（Table 2 / 表 2 / Figure 1 / 图 1）——它们常与 TOC 里的小节名重名（如
        #      "Normality Test" 同时是 L3 标题和表题的前缀）
        if level is None and sn in {"normal", "list paragraph", ""} and 4 <= len(text) <= 80:
            text_lc = text.lower().strip()
            is_back_matter_head = (
                text_lc in {"bibliography", "references"}
                or re.match(r"^致\s*谢", text) is not None
                or re.match(r"^acknowledg", text_lc) is not None
                or re.match(r"^appendix", text_lc) is not None
            )
            is_caption = bool(
                re.match(r"^\s*(Table|表|Figure|图|Picture|Example)\s*\d", text_lc, re.IGNORECASE)
                or re.match(r"^\s*(Table|表|Figure|图|Picture|Example)\s+[A-Z]", text_lc, re.IGNORECASE)
            )
            if is_back_matter_head or is_caption:
                continue  # 不当作章节标题
            nt = _norm(text)
            for lvl in (1, 2, 3):
                for tt in toc_norm.get(lvl, set()):
                    if not tt:
                        continue
                    if nt == tt or (len(nt) >= 8 and (nt in tt or tt in nt)):
                        level = lvl
                        break
                if level is not None:
                    break

        if level is not None:
            # 末尾标题场景：用干净标题文本覆盖（如 "4 Conclusion" 而非整段）
            display_text = tail_title if tail_title else text[:80]
            out.append({"level": level, "idx": i, "text": display_text})
    return out


def _detect_zones_v2(doc) -> "OrderedDict[str, Tuple[int, int]]":
    """基于 Word 段落样式 + 位置感知的版本。返回 {区域: (start, end_exclusive)}，
    索引基于所有 body 子元素（混合段落+表格）。

    关键改进：
    1. 目录结束位置 = 目录 header 之后第一个非空、非 TOC 条目的段落（自然限制目录长度）
    2. 显式构造"正文"zone（前后区域之间）
    3. 表格也参与扫描（封面常是表格布局）
    """
    body_children = list(doc.element.body)
    n = len(body_children)
    para_by_element = {p._element: p for p in doc.paragraphs}

    def _classify_zone_header(text: str, style_name: str) -> Optional[str]:
        text = (text or "").strip()
        if not text:
            return None
        # 成绩评定表注释
        if re.match(r"^注[：:]\s*1\.", text) or re.match(r"^[“”]等级[“”][：:]", text):
            return "成绩评定表"
        head = text[:15]
        if "学生承诺书" in head and len(text) < 100:
            return "学生承诺书"
        if re.match(r"^摘\s+要", text) or re.match(r"^摘要", text):
            return "中文摘要"
        if re.match(r"^Abstract", text, re.I) and len(text) < 100:
            return "英文摘要"
        if (re.match(r"^Table of Contents", text, re.I)
                or re.match(r"^目\s+录", text)
                or re.match(r"^目录", text)) and len(text) < 100:
            return "目录"
        if (re.match(r"^参考文献", text) or re.match(r"^Bibliography", text)) and len(text) < 100:
            return "参考文献"
        if re.match(r"^\s*Appendix\s+[A-Z]", text) and len(text) < 100:
            return "附录"
        if (re.match(r"^致谢", text)
                or re.match(r"^Acknowledgement", text, re.I)) and len(text) < 100:
            return "致谢"
        return None

    # 第一遍：扫描所有段落+表格，记录 zone-header 位置和 TOC 条目
    boundaries: List[Tuple[int, str]] = []  # (ci, zone_name)
    # 记录 sdt 容器（含 Word TOC 字段的 Structured Document Tag）
    sdt_with_toc: List[int] = []  # ci of sdt whose text contains "Table of Contents"
    for ci, child in enumerate(body_children):
        tag = child.tag.split("}")[-1]
        # 只取 <w:t> 节点文本（避免某些导出工具在 <w:p> 下放 orphan 文本节点导致重复）
        text = "".join(t.text or "" for t in child.iter(qn("w:t")))
        if tag == "p":
            para = para_by_element.get(child)
            # 关键防御：toc 1 样式的段落（如目录里的 "Bibliography\t44"）
            # 绝对不能被识别成"参考文献"边界。
            # 但 TOC 区域标题本身（"Table of Contents"）可能是 toc 1 样式，
            # 必须放行让 _classify_zone_header 分类为"目录"。
            # Word 的 TOC 字段渲染会重复文本（如 "Table of ContentsTable of Contents..."），
            # 所以用 search 而不是 match。
            text_stripped = text.strip()
            is_likely_header = bool(
                re.search(r"Table of Contents", text_stripped, re.I)
                or re.search(r"^目\s*录", text_stripped)
            )
            if _is_toc_paragraph(para) and not is_likely_header:
                continue
            name = _classify_zone_header(text, "")
            if name:
                boundaries.append((ci, name))
        elif tag == "tbl":
            # 表格里的封面标题、扉页英文标题等 — 也可能含 "参考文献" 字符但极少见
            name = _classify_zone_header(text, "")
            if name:
                boundaries.append((ci, name))
        elif tag == "sdt":
            # sdt = Word Structured Document Tag，常用于包装 TOC 字段。
            # 文本以 "Table of Contents" 开头（含 Word 渲染后的重复文本），
            # 视为目录区段起点（即使目录为空）。
            text_stripped = text.strip()
            if re.search(r"Table\s+of\s+Contents", text_stripped, re.I):
                sdt_with_toc.append(ci)
                boundaries.append((ci, "目录"))

    # 去重：同一 ci 只保留第一个匹配的 zone 名
    seen_ci = set()
    boundaries = [(ci, nm) for ci, nm in boundaries if not (ci in seen_ci or seen_ci.add(ci))]

    # 第二步：精确定位"目录"结束位置
    toc_start = next((ci for ci, nm in boundaries if nm == "目录"), None)
    # 是否 toc_start 是 SDT 容器（Word TOC 字段）？
    toc_start_is_sdt = (
        toc_start is not None
        and body_children[toc_start].tag.split("}")[-1] == "sdt"
    )
    if toc_start is not None:
        toc_end = None
        if toc_start_is_sdt:
            # SDT 容器是 Word 把所有 TOC 段落打包的结构，body_children 中只占一个 ci。
            # 因此目录结束位置就是 toc_start + 1（容器之后第一个 body child）。
            toc_end = toc_start + 1
        else:
            # 目录头是普通段落（如 "Table of Contents" 文本），需要向后扫描找正文起点
            for ci in range(toc_start + 1, n):
                tag = body_children[ci].tag.split("}")[-1]
                if tag == "tbl":
                    continue  # 跳过表格
                if tag == "sdt":
                    continue  # 跳过中间可能的 SDT
                para = para_by_element.get(body_children[ci])
                text = "".join(t.text or "" for t in body_children[ci].iter(qn("w:t"))).strip()
                if not text:
                    continue  # 跳过空段落
                if _is_toc_paragraph(para):
                    continue  # 跳过 TOC 条目（含 toc 1 样式的目录条目）
                sn = _para_style_name(para).lower() if para else ""
                if sn in {"normal", "list paragraph", ""} and _classify_zone_header(
                    text, sn
                ) in {"参考文献", "附录", "致谢", "成绩评定表"}:
                    break
                toc_end = ci
                break
            if toc_end is None:
                toc_end = n

        # 移除可能因同一 ci 被分类到多 zone 的旧边界（极少见，但防御性）
        boundaries = [(ci, nm) for ci, nm in boundaries if nm != "目录"]
        boundaries.append((toc_start, "目录"))
        boundaries.append((toc_end, "正文"))

    # 第三步：边界排序，构建 spans
    sorted_b = sorted(set(boundaries), key=lambda x: x[0])
    spans: "OrderedDict[str, List[Tuple[int, int]]]" = OrderedDict()

    first_classified = sorted_b[0][0] if sorted_b else n
    if first_classified > 0:
        spans["封面/扉页"] = [(0, first_classified)]

    for j, (idx, name) in enumerate(sorted_b):
        end = sorted_b[j + 1][0] if j + 1 < len(sorted_b) else n
        spans.setdefault(name, []).append((idx, end))

    # 第四步：保证"正文"span 存在（在所有前置区域结束后、所有后置区域开始前）
    front_names = {"学生承诺书", "中文摘要", "英文摘要", "目录"}
    back_names = {"参考文献", "附录", "致谢", "成绩评定表"}
    front_end = max(
        ([e for nm, lst in spans.items() if nm in front_names for (_, e) in lst]
         + [first_classified]),
        default=first_classified,
    )
    back_starts = [s for nm, lst in spans.items() if nm in back_names for (s, _) in lst]
    back_start = min(back_starts) if back_starts else n

    # 兜底：找第一个 Heading 1 样式的段落作为正文起点（处理无目录文档）
    # 必须跳过空 Heading 1 段落（Word 常留作占位符，无内容，不能作为正文起点）
    heading_start = None
    for ci, child in enumerate(body_children):
        tag = child.tag.split("}")[-1]
        if tag != "p":
            continue
        para = para_by_element.get(child)
        if not para:
            continue
        sn = _para_style_name(para).lower()
        if sn in {"heading 1", "标题 1"}:
            text = "".join(child.itertext()).strip()
            if text:
                heading_start = ci
                break

    # 如果正文起点（heading_start）早于 front_end，说明 front 区段被错拉长
    # 跨进了正文区域（典型：无目录文档，英文摘要覆盖了 Heading 1 之前的正文）。
    # 直接把 front 区段截断到 heading_start，并丢弃被夹扁的区段。
    if (heading_start is not None
            and heading_start < front_end
            and heading_start < back_start):
        for nm in list(spans.keys()):
            if nm in front_names:
                spans[nm] = [
                    (s, min(e, heading_start))
                    for (s, e) in spans[nm]
                    if s < min(e, heading_start)
                ]
                if not spans[nm]:
                    del spans[nm]
        front_end = heading_start

    if back_start > front_end and "正文" not in spans:
        spans["正文"] = [(front_end, back_start)]
    elif "正文" in spans:
        # 已有正文 span（来自 toc_end 注入），确保它在 front_end..back_start 范围内
        s0, e0 = spans["正文"][0]
        spans["正文"] = [(max(s0, front_end), min(e0, back_start))]
        if spans["正文"][0][0] >= spans["正文"][0][1]:
            # 被夹扁了，丢掉
            del spans["正文"]
    elif back_start <= front_end and heading_start is not None and heading_start < back_start:
        # 兜底场景：前置和后置区域在 body child 上相邻/重叠（典型：没目录文档，
        # 英文摘要结束后立刻接 Bibliography）。用第一个 Heading 1 作为正文起点。
        spans["正文"] = [(heading_start, back_start)]

    # 第五步：合并同名 span（非正文 collapse；正文严格保持单 span）
    merged: "OrderedDict[str, Tuple[int, int]]" = OrderedDict()
    for name, lst in spans.items():
        if name == "正文":
            if lst:
                merged[name] = lst[0]
        else:
            if lst:
                merged[name] = (lst[0][0], lst[-1][1])
    return merged


def _is_body_paragraph(text: str, para, zone: str) -> bool:
    """正文段落判定。比旧的 `len(text) > 30` 更宽松。"""
    if _is_toc_paragraph(para):
        return False
    if len(text) < 4:
        return False
    if _is_reference_entry(text):
        return False
    is_cap, _ = _is_figure_or_table_caption(text)
    if is_cap:
        return False
    text_strip = text.strip()
    # 区域标题本身（如 "Acknowledgements"、"致谢"）不应被当作正文段
    if text_strip in (
        "Table of Contents", "Bibliography", "Appendix",
        "Acknowledgements", "致谢", "参考文献", "目录",
    ):
        return False
    # 附录标题段（如 "Appendix A Screenshots of ..."、"Appendix B The Test ..."）
    # 不算正文段，单独走附录标题检查
    if re.match(r"^\s*Appendix\s+[A-Z]", text_strip, re.IGNORECASE):
        return False
    return True


def _count_inline_elements(paragraphs, body_range: Tuple[int, int]) -> Dict:
    """扫描正文段落，统计内文引用、图表题、例证。
    同步统计上标引用率（针对 [1]/[2] 等方括号引用是否使用上标）。"""
    start, end = body_range
    cite_pat = re.compile(r"\[\d+(?:[,\-，\s]\s*\d+)*\]")
    ex_pat = re.compile(r"^\s*例\s*\d+|^\s*例如")
    fig_pat = re.compile(r"^\s*(Figure|Fig\.?|图)\s*\d+", re.I)
    tab_pat = re.compile(r"^\s*(Table|Tab\.?|表)\s*\d+", re.I)
    counts = {"inline_citations": 0, "examples": 0, "figures": 0, "tables": 0}
    citation_paras = 0
    cite_brackets_total = 0
    cite_brackets_superscript = 0
    for i, para in enumerate(paragraphs):
        if not (start <= i < end):
            continue
        text = (para.text or "")
        if not text.strip():
            continue
        c = len(cite_pat.findall(text))
        if c:
            counts["inline_citations"] += c
            citation_paras += 1
        if ex_pat.match(text.strip()):
            counts["examples"] += 1
        if fig_pat.match(text.strip()):
            counts["figures"] += 1
        if tab_pat.match(text.strip()):
            counts["tables"] += 1
        # 检查上标率：逐 run 看 [数字] 模式所在 run 是否 superscript
        for run in para.runs:
            rt = run.text or ""
            if not rt:
                continue
            brackets = cite_pat.findall(rt)
            if brackets:
                cite_brackets_total += len(brackets)
                if _is_run_superscript(run):
                    cite_brackets_superscript += len(brackets)
    counts["citation_paragraphs"] = citation_paras
    counts["cite_brackets_total"] = cite_brackets_total
    counts["cite_brackets_superscript"] = cite_brackets_superscript
    return counts


def _check_page_numbers(doc, zones, _add_issue) -> None:
    """检查 sections 的 pgNumType 设置：
    - 目录区段应为 upperRoman（I, II, III…）
    - 正文区段应为 decimal（1, 2, 3…）
    - 摘要/承诺书区段通常无页码或特殊处理
    """
    try:
        sections = doc.sections
    except Exception:
        return
    if not sections:
        return

    # 把 zones 按起始 body-child 索引排序
    zone_order = sorted(
        [(name, span) for name, span in zones.items() if span and span[0] is not None],
        key=lambda x: x[1][0],
    )

    # 简化策略：检查每个 section 的 sectPr 里 pgNumType
    for sec_idx, section in enumerate(sections):
        sectPr = section._sectPr
        pgNumType = sectPr.find(qn("w:pgNumType"))
        fmt = pgNumType.get(qn("w:fmt")) if pgNumType is not None else None
        # 找这个 section 对应的 zone（按 body 子元素索引）
        # sections 在 doc 里按段落顺序排列，对应大致 zone 区间
        # 简化：基于 sec_idx 推断
        if sec_idx == 0:
            target_zone = "封面/扉页"
        else:
            target_zone = None
            # 通过统计累计子元素推断
            offset = sum(len(doc.element.body[i]) if False else 0 for i in range(0))
        # 实际简化：只发警告，不做严格 zone 匹配
        if fmt is None:
            continue  # 无明确设置，跳过
        # 检查正文区段（如 sec_idx 对应目录/正文）：应使用 decimal
        # 我们通过 zone 位置推断
        if sec_idx >= len(sections) - 2 and fmt != "decimal":
            # 后两个 section（一般是正文/附录）：应为 decimal
            _add_issue({
                "rule": "正文页码格式",
                "severity": "中",
                "location": f"Section {sec_idx + 1}",
                "expected": "正文页码应为阿拉伯数字（decimal）",
                "actual": f"当前格式 {fmt}",
                "suggestion": "在 Word 中：插入 > 页码 > 页码格式 > 1, 2, 3 (阿拉伯数字)",
                "zone": "正文",
            })
        if sec_idx == 1 and fmt not in ("upperRoman", "upper-roman"):
            # 通常第 2 个 section 对应目录：应为罗马数字
            _add_issue({
                "rule": "目录页码格式",
                "severity": "低",
                "location": f"Section {sec_idx + 1}",
                "expected": "目录页码应为罗马数字（I, II, III）",
                "actual": f"当前格式 {fmt}",
                "suggestion": "在 Word 中：插入 > 页码 > 页码格式 > I, II, III (罗马数字)",
                "zone": "目录",
            })


def _is_abstract_keyword_line(text: str, lang: str) -> bool:
    """判断是否为摘要关键词行（仅这一行需要查"3-5 个"）"""
    if lang == "zh":
        return bool(re.match(r"^\s*关键词[：:]", text))
    else:
        return bool(re.match(r"^\s*[Kk]eywords?[：:]", text))


def _check_abstract_keywords_count(text: str, lang: str) -> list:
    issues = []
    if lang == "en":
        m = re.search(r"[Kk]eywords?[：:]\s*(.+)", text)
    else:
        m = re.search(r"关键词[：:]\s*(.+)", text)
    if not m:
        return issues
    kws = m.group(1).strip()
    parts = [p.strip() for p in re.split(r"[;；，,]", kws) if p.strip()]
    if len(parts) < 3:
        issues.append(f"关键词过少（{len(parts)}个），规范要求 3-5 个")
    elif len(parts) > 5:
        issues.append(f"关键词过多（{len(parts)}个），规范要求 3-5 个")
    return parts, issues


# === 主体检查 ===

def check_document(docx_path: str) -> dict:
    doc = Document(docx_path)
    issues: List[Dict] = []
    stats = {}

    # ---- 1. 页面设置 ----
    section = doc.sections[0]
    margin_top = section.top_margin.cm if section.top_margin else 0
    margin_bottom = section.bottom_margin.cm if section.bottom_margin else 0
    margin_left = section.left_margin.cm if section.left_margin else 0
    margin_right = section.right_margin.cm if section.right_margin else 0
    stats["page_margin"] = {
        "top": round(margin_top, 2),
        "bottom": round(margin_bottom, 2),
        "left": round(margin_left, 2),
        "right": round(margin_right, 2),
    }
    for name, val in [("上", margin_top), ("下", margin_bottom),
                       ("左", margin_left), ("右", margin_right)]:
        if abs(val - EXPECTED["page_margin_cm"]) > 0.1:
            issues.append({
                "rule": "页面边距",
                "severity": "中",
                "location": "页面设置",
                "expected": f"上下左右各 {EXPECTED['page_margin_cm']}cm",
                "actual": f"{name}边距 {val}cm",
                "suggestion": f"Word → 布局 → 页边距 → 自定义边距 → {name}设为 {EXPECTED['page_margin_cm']}cm",
                "zone": "封面/扉页",
            })

    # ---- 2. 区域识别（v2：基于样式 + 位置感知）----
    zones = _detect_zones_v2(doc)
    stats["zones"] = {k: v for k, v in zones.items()}

    # 构建 body 子元素到"段落号"的映射（用于可视标注）
    # 在混合段落/表格布局下，paragraph 索引只对应 doc.paragraphs 列表
    para_idx_by_id = {p._element: i for i, p in enumerate(doc.paragraphs)}

    # ---- 2b. 章节标题识别（v2：样式名 + 阿拉伯数字正则）----
    headings = _detect_headings(doc.paragraphs)
    # 剔除位于"封面/扉页"zone 的伪标题（如封面副标题偶然与 TOC 标题模糊匹配）
    cover_span = zones.get("封面/扉页")
    if cover_span:
        cover_ci_start, cover_ci_end = cover_span
        # cover_span 是 body_child 索引；映射到 paragraph 索引
        all_children_list = list(doc.element.body)
        cover_para_idxs = set()
        for ci in range(cover_ci_start, cover_ci_end):
            if ci < len(all_children_list):
                child = all_children_list[ci]
                if child.tag.endswith("}p"):
                    pi = para_idx_by_id.get(child)
                    if pi is not None:
                        cover_para_idxs.add(pi)
        headings = [h for h in headings if h["idx"] not in cover_para_idxs]
    stats["headings"] = headings  # 详细列表（含 idx、level、text）
    heading_level_by_idx = {h["idx"]: h["level"] for h in headings}

    # body zone 的"段落号"范围（用于内文元素统计）
    body_zone_ci = zones.get("正文")
    if body_zone_ci:
        all_children_list = list(doc.element.body)
        body_ci_start, body_ci_end = body_zone_ci
        body_para_start = None
        body_para_end = None
        for ci in range(body_ci_start, body_ci_end):
            if ci >= len(all_children_list):
                break
            child = all_children_list[ci]
            tag = child.tag.split("}")[-1]
            if tag != "p":
                continue
            pi = para_idx_by_id.get(child)
            if pi is None:
                continue
            if body_para_start is None:
                body_para_start = pi
            body_para_end = pi
        if body_para_start is not None:
            # +1 是因为人类阅读段号从 1 开始
            stats["body_zone_range"] = (body_para_start + 1, (body_para_end or body_para_start) + 1)
        else:
            stats["body_zone_range"] = (0, 0)
    else:
        stats["body_zone_range"] = (0, 0)

    # ---- 3. 遍历 body 子元素（段落+表格，封面常是表格）----
    total_paragraphs = 0
    body_paragraphs = 0
    found_abstract_zh = False
    found_abstract_en = False
    found_toc = False
    found_references = False
    found_zh_title_on_cover = False
    found_en_title_on_cover = False
    found_pledge = False
    heading_count = Counter()

    # 目录可能是 Word SDT 容器（TOC 字段），其内部段落不出现于 doc.paragraphs，
    # 因此下面 paragraph 循环里 zone=="目录" 永远不会命中。
    # 提前检查 zones 中是否有"目录"区段。
    if zones.get("目录") and zones["目录"][1] > zones["目录"][0]:
        found_toc = True

    issue_per_para: Dict[int, List[Dict]] = {}

    def _add_issue(iss):
        issues.append(iss)
        loc = iss.get("location", "")
        m = re.match(r"第(\d+)段", loc)
        if m:
            pi = int(m.group(1)) - 1
            issue_per_para.setdefault(pi, []).append(iss)

    # 找出封面/扉页区域的所有 body 子元素（含表格）
    cover_range = zones.get("封面/扉页", (0, 0))
    cover_children = list(doc.element.body)[cover_range[0]:cover_range[1]]

    # ---- 3a. 扫描封面区找中文封面标题（包括表格里的）----
    def _extract_table_cells_clean(tbl_element) -> list[str]:
        """从表格元素中提取去重后的单元格文本列表"""
        cells_seen = set()
        result = []
        # 找所有 <w:tc> 单元格
        for tc in tbl_element.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
            pass
        # 更稳妥的方式：找所有 <w:tc>
        for tc in tbl_element.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc'):
            cell_text = ''.join(t.text or '' for t in tc.iter(
                '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'))
            cell_text = cell_text.strip()
            if cell_text and cell_text not in cells_seen:
                cells_seen.add(cell_text)
                result.append(cell_text)
        return result

    if not found_zh_title_on_cover:
        for child in cover_children:
            tag = child.tag.split('}')[-1]
            if tag == 'p':
                t = ''.join(child.itertext()).strip()
                if not t:
                    continue
                # 去重：合并单元格导致字符重复，如 "摘        要摘摘                要要"
                t_dedup = re.sub(r"(.)\1{2,}", r"\1", t)  # 3+ 连续重复字符压成 1 个
                t_dedup = re.sub(r"(\S{2,8}?)\1+", r"\1", t_dedup)  # 重复短语也压一下
                clean = re.sub(r"[—–-]+.*$", "", t_dedup).strip()
                has_chinese = bool(re.search(r"[一-龥]", clean))
                has_academic = bool(re.search(r"(研究|分析|探讨|翻译|调查|报告|设计|策略|应用|视角|基于)", clean))
                if (has_chinese and has_academic
                        and 6 < len(clean) < 60
                        and not re.search(r"\d+\s*$", clean)
                        and t not in ("学生承诺书",)):
                    found_zh_title_on_cover = True
                    stats["zh_cover_title"] = {"text": clean, "raw": t[:80]}
                    break
            elif tag == 'tbl':
                # 表格：按单元格去重后扫描
                cells = _extract_table_cells_clean(child)
                for cell_text in cells:
                    if cell_text in ("仲恺农业工程学院", "毕 业 论 文", "毕  业  论  文"):
                        continue
                    # 单元格内可能有 \n（如"认知负荷视角下大疆Neo 2用户手册\n英译研究"）
                    sub_lines = [ln.strip() for ln in cell_text.split('\n') if ln.strip()]
                    for line in sub_lines:
                        clean = re.sub(r"[—–-]+.*$", "", line).strip()
                        has_chinese = bool(re.search(r"[一-龥]", clean))
                        has_academic = bool(re.search(r"(研究|分析|探讨|翻译|调查|报告|设计|策略|应用|视角|基于)", clean))
                        if (has_chinese and has_academic
                                and 6 < len(clean) < 60
                                and not re.search(r"\d+\s*$", clean)):
                            found_zh_title_on_cover = True
                            stats["zh_cover_title"] = {"text": clean, "raw": cell_text[:80]}
                            break
                    if found_zh_title_on_cover:
                        break
            if found_zh_title_on_cover:
                break

    # ---- 3b. 扫描所有 body 元素（按出现顺序走完封面）----
    # 找出"封面结束"位置：封面区之后第一个非封面元素
    cover_end_ci = cover_range[1]  # body child index
    all_children = list(doc.element.body)

    for ci, child in enumerate(all_children):
        if ci >= cover_end_ci:
            break  # 过了封面/扉页区
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            t = ''.join(child.itertext()).strip()
            if not t:
                continue
            # 收集封面/扉页全部文本（去重后），用于字段完整性检查
            t_dedup = re.sub(r"(.)\1{2,}", r"\1", t)
            t_dedup = re.sub(r"(\S{2,8}?)\1+", r"\1", t_dedup)
            stats.setdefault("cover_all_text", "")
            stats["cover_all_text"] += t_dedup + "\n"
            if not found_en_title_on_cover:
                if (re.match(r"^[A-Z]", t_dedup)
                        and 20 < len(t_dedup) < 200
                        and re.search(r"\b(of|in|on|under|from|with|through|by)\b", t_dedup, re.I)
                        and not _is_figure_or_table_caption(t_dedup)[0]):
                    found_en_title_on_cover = True
                    stats["en_cover_title"] = {"text": t_dedup}
        elif tag == 'tbl':
            # 扉页可能在表格里
            if not found_en_title_on_cover:
                cells = _extract_table_cells_clean(child)
                for cell_text in cells:
                    stats.setdefault("cover_all_text", "")
                    stats["cover_all_text"] += cell_text + "\n"
                    if (re.match(r"^[A-Z]", cell_text)
                            and 20 < len(cell_text) < 200
                            and re.search(r"\b(of|in|on|under|from|with|through|by)\b", cell_text, re.I)):
                        found_en_title_on_cover = True
                        stats["en_cover_title"] = {"text": cell_text}
                        break

    # ---- 3c. 遍历所有段落（剩下的检查）----
    # 把 body 子元素索引和段落索引对应起来
    child_to_paraidx = {}
    para_counter = 0
    for ci, child in enumerate(all_children):
        if child.tag.endswith('}p'):
            child_to_paraidx[ci] = para_counter
            para_counter += 1

    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        total_paragraphs += 1
        # 找到该段落对应的 body child index
        body_ci = None
        for ci, pi in child_to_paraidx.items():
            if pi == idx:
                body_ci = ci
                break
        zone = _zone_of(body_ci, zones) if body_ci is not None else "未知"

        # 检测"漏按 Enter 拼接章节标题"问题：
        # 正文段末尾出现 "\d+\s+Title" 模式（典型： "...cognitive load.4 Conclusion"）
        if zone == "正文" and len(text) > 50:
            m_end = re.search(
                r"(\d{1,2}(?:\.\d{1,2}){0,2})\s+"
                r"([A-Za-z一-鿿][A-Za-z0-9\s,.\-:;'一-鿿]{2,60})$",
                text,
            )
            if m_end:
                title_start_idx = m_end.start()
                prefix = text[:title_start_idx].rstrip()
                if len(prefix) > 30 and re.search(r"[。.!?！」』”\"']\s*$", prefix):
                    tail_title = text[title_start_idx:].strip()
                    _add_issue({
                        "rule": "章节标题与正文未分段",
                        "severity": "高",
                        "location": f"第{idx+1}段：「{prefix[-30:]}」后接「{tail_title}」",
                        "expected": "章节标题独立成段（与正文之间用 Enter 分段）",
                        "actual": f"标题「{tail_title}」被拼接到正文末尾",
                        "suggestion": (
                            f"作者漏按 Enter 导致章节标题「{tail_title}」被拼接到正文末尾。"
                            f"请将光标定位到「{tail_title}」之前，按 Enter 单独成段，"
                            f"然后为该段应用一级标题格式（Times New Roman 四号加粗、左顶格）。"
                        ),
                        "zone": "正文",
                    })

        # 封面/扉页区已经在前面扫过，这里跳过
        if zone == "封面/扉页":
            continue

        # 承诺书判断
        if zone == "学生承诺书":
            found_pledge = True
            continue

        # ---- 中文摘要 ----
        if zone == "中文摘要":
            found_abstract_zh = True
            # 摘要标题检查："摘        要" 应该是 黑体 14pt 加粗居中
            if "摘" in text and "要" in text and len(text) < 15:
                run = _first_nonempty_run(para) if para.runs else None
                if run:
                    sz = _consensus_size(para) or _get_effective_size(run, para)
                    fn = _get_effective_font(run, para, has_chinese=True)
                    if sz and abs(sz - EXPECTED["abstract_title_size_pt"]) > 1.5:
                        _add_issue({
                            "rule": "中文摘要标题字号",
                            "severity": "中",
                            "location": f"第{idx+1}段：「{text[:10]}」",
                            "expected": "黑体 四号加粗居中",
                            "actual": f"当前字号 {_pt_to_zh(sz) if sz else '?'}",
                            "suggestion": "摘要标题应为黑体 四号加粗居中",
                            "zone": "中文摘要",
                        })
                    if fn and ("黑体" not in fn and "SimHei" not in fn):
                        _add_issue({
                            "rule": "中文摘要标题字体",
                            "severity": "中",
                            "location": f"第{idx+1}段：「{text[:10]}」",
                            "expected": EXPECTED["abstract_title_font"],
                            "actual": fn,
                            "suggestion": f"摘要标题用 {EXPECTED['abstract_title_font']}，更醒目",
                            "zone": "中文摘要",
                        })
                continue
            # 关键词行检查
            if _is_abstract_keyword_line(text, "zh"):
                if para.runs:
                    label_run = _first_nonempty_run(para) or para.runs[0]
                    label_fn = _get_effective_font(label_run, para)
                    label_sz = _get_effective_size(label_run, para)
                    if label_fn and ("黑体" not in label_fn and "SimHei" not in label_fn):
                        _add_issue({
                            "rule": "中文关键词标签字体",
                            "severity": "低",
                            "location": f"第{idx+1}段：「{text[:20]}」",
                            "expected": f"{EXPECTED['keyword_label_font']} 小四",
                            "actual": label_fn,
                            "suggestion": f"'关键词：'标签用 {EXPECTED['keyword_label_font']}，与正文宋体区分",
                            "zone": "中文摘要",
                        })
                    if label_sz and abs(label_sz - EXPECTED["keyword_label_size_pt"]) > 1.0:
                        _add_issue({
                            "rule": "中文关键词标签字号",
                            "severity": "低",
                            "location": f"第{idx+1}段：「{text[:20]}」",
                            "expected": "小四",
                            "actual": f"当前字号 {_pt_to_zh(label_sz)}",
                            "suggestion": "关键词标签用小四号",
                            "zone": "中文摘要",
                        })
                parts, kw_issues = _check_abstract_keywords_count(text, "zh")
                for k in kw_issues:
                    _add_issue({
                        "rule": "中文摘要关键词数量",
                        "severity": "中",
                        "location": f"第{idx+1}段",
                        "expected": f"{EXPECTED['keyword_min']}-{EXPECTED['keyword_max']} 个关键词（中文分号分隔）",
                        "actual": k,
                        "suggestion": f"关键词：词1{EXPECTED['keyword_separator']}词2{EXPECTED['keyword_separator']}词3（{EXPECTED['keyword_min']}-{EXPECTED['keyword_max']} 个）",
                        "zone": "中文摘要",
                    })
                if ";" in text and "；" not in text:
                    _add_issue({
                        "rule": "中文关键词分隔符",
                        "severity": "低",
                        "location": f"第{idx+1}段",
                        "expected": "中文分号 ；",
                        "actual": "英文分号 ;",
                        "suggestion": "中文摘要关键词之间用中文分号「；」分隔（规范要求全角标点）",
                        "zone": "中文摘要",
                    })
            # 摘要正文段落：宋体/Times New Roman 小四
            elif len(text) > 30 and not _is_abstract_keyword_line(text, "zh"):
                has_chinese = bool(re.search(r"[一-龥]", text))
                if para.runs:
                    run = _first_nonempty_run(para) or para.runs[0]
                    sz = _consensus_size(para)
                    fn = _get_effective_font(run, para, has_chinese)
                    expected_font = EXPECTED["body_font_zh"] if has_chinese else EXPECTED["body_font_en"]
                    if sz and abs(sz - EXPECTED["body_size_pt"]) > 1.0:
                        _add_issue({
                            "rule": "中文摘要正文字号",
                            "severity": "低",
                            "location": f"第{idx+1}段：「{text[:30]}...」",
                            "expected": "小四",
                            "actual": f"当前字号 {_pt_to_zh(sz)}",
                            "suggestion": "摘要正文用小四号。正确格式示例：「本研究采用文本分析法……」（宋体小四）",
                            "zone": "中文摘要",
                        })
                    if fn and fn != expected_font:
                        _add_issue({
                            "rule": "中文摘要正文字体",
                            "severity": "低",
                            "location": f"第{idx+1}段：「{text[:30]}...」",
                            "expected": f"{expected_font}（{'中文' if has_chinese else '英文/数字'}用）",
                            "actual": f"当前字体 {fn}",
                            "suggestion": f"摘要正文{'中文段落用宋体，英文/数字用 Times New Roman' if has_chinese else '中英文/数字用 Times New Roman'}。正确格式示例：「本研究采用文本分析法……」（宋体小四）",
                            "zone": "中文摘要",
                        })
                ls = para.paragraph_format.line_spacing
                if ls is not None and abs(ls - EXPECTED["line_spacing"]) > 0.1:
                    _add_issue({
                        "rule": "摘要行距",
                        "severity": "低",
                        "location": f"第{idx+1}段：「{text[:30]}...」",
                        "expected": f"{EXPECTED['line_spacing']} 倍",
                        "actual": f"{ls} 倍",
                        "suggestion": "摘要 1.5 倍行距",
                        "zone": "中文摘要",
                    })
            continue

        # ---- 英文摘要 ----
        if zone == "英文摘要":
            found_abstract_en = True
            if re.match(r"^\s*Abstract\s*$", text, re.I):
                run = _first_nonempty_run(para) if para.runs else None
                if run:
                    sz = _consensus_size(para)
                    fn = _get_effective_font(run, para)
                    if sz and abs(sz - EXPECTED["abstract_en_header_size_pt"]) > 1.5:
                        _add_issue({
                            "rule": "Abstract 标题字号",
                            "severity": "中",
                            "location": f"第{idx+1}段：「{text}」",
                            "expected": "Times New Roman 三号加粗居中",
                            "actual": f"当前字号 {_pt_to_zh(sz)}",
                            "suggestion": "Abstract 标题应用 Times New Roman 三号加粗居中。正确格式示例：「Abstract」（三号加粗居中）",
                            "zone": "英文摘要",
                        })
                    # 加粗（含样式继承）
                    if _is_run_bold_inherited(run, para) is False:
                        _add_issue({
                            "rule": "Abstract 标题加粗",
                            "severity": "中",
                            "location": f"第{idx+1}段：「{text}」",
                            "expected": "加粗",
                            "actual": "当前未加粗",
                            "suggestion": "Abstract 标题应加粗",
                            "zone": "英文摘要",
                        })
                    # 居中
                    if not _is_paragraph_centered(para):
                        _add_issue({
                            "rule": "Abstract 标题居中",
                            "severity": "低",
                            "location": f"第{idx+1}段：「{text}」",
                            "expected": "居中",
                            "actual": "未居中",
                            "suggestion": "Abstract 标题应居中",
                            "zone": "英文摘要",
                        })
                continue
            if _is_abstract_keyword_line(text, "en"):
                parts, kw_issues = _check_abstract_keywords_count(text, "en")
                for k in kw_issues:
                    _add_issue({
                        "rule": "英文摘要关键词数量",
                        "severity": "中",
                        "location": f"第{idx+1}段",
                        "expected": f"{EXPECTED['keyword_min']}-{EXPECTED['keyword_max']} 个 Keywords",
                        "actual": k,
                        "suggestion": f"Keywords: word1; word2; word3（{EXPECTED['keyword_min']}-{EXPECTED['keyword_max']} 个，英文分号分隔）",
                        "zone": "英文摘要",
                    })
                if para.runs:
                    label_run = _first_nonempty_run(para) or para.runs[0]
                    if _is_run_bold_inherited(label_run, para) is False:
                        _add_issue({
                            "rule": "英文关键词标签加粗",
                            "severity": "低",
                            "location": f"第{idx+1}段：「{text[:20]}」",
                            "expected": "Keywords: 加粗",
                            "actual": "未加粗",
                            "suggestion": "Keywords: 标签应加粗",
                            "zone": "英文摘要",
                        })
            elif len(text) > 30:
                has_chinese = bool(re.search(r"[一-龥]", text))
                if para.runs:
                    run = _first_nonempty_run(para) or para.runs[0]
                    sz = _consensus_size(para)
                    fn = _get_effective_font(run, para, has_chinese)
                    if sz and abs(sz - EXPECTED["body_size_pt"]) > 1.0:
                        _add_issue({
                            "rule": "英文摘要正文字号",
                            "severity": "低",
                            "location": f"第{idx+1}段：「{text[:30]}...」",
                            "expected": "小四",
                            "actual": f"当前字号 {_pt_to_zh(sz)}",
                            "suggestion": "英文摘要正文用小四号。正确格式示例：「This study adopts...」（Times New Roman 小四）",
                            "zone": "英文摘要",
                        })
                    if fn and fn != "Times New Roman":
                        _add_issue({
                            "rule": "英文摘要正文字体",
                            "severity": "低",
                            "location": f"第{idx+1}段：「{text[:30]}...」",
                            "expected": "Times New Roman",
                            "actual": f"当前字体 {fn}",
                            "suggestion": "英文摘要正文用 Times New Roman。正确格式示例：「This study adopts qualitative analysis...」（Times New Roman 小四）",
                            "zone": "英文摘要",
                        })
                ls = para.paragraph_format.line_spacing
                if ls is not None and abs(ls - EXPECTED["line_spacing"]) > 0.1:
                    _add_issue({
                        "rule": "英文摘要行距",
                        "severity": "低",
                        "location": f"第{idx+1}段：「{text[:30]}...」",
                        "expected": f"{EXPECTED['line_spacing']} 倍",
                        "actual": f"{ls} 倍",
                        "suggestion": "Abstract 1.5 倍行距",
                        "zone": "英文摘要",
                    })
            continue

        # ---- 目录 ----
        if zone == "目录":
            found_toc = True
            # 区分"目录标题"段（Table of Contents）和"目录条目"段
            is_toc_header = bool(re.match(r"^\s*(Table\s+of\s+Contents|目\s*录|目录)\s*$", text, re.I))
            if is_toc_header:
                # 标题：Times New Roman 四号 加粗 居中
                run = _first_nonempty_run(para)
                if run:
                    sz = _size_to_pt(run.font.size)
                    if sz and abs(sz - EXPECTED["toc_header_size_pt"]) > 1.5:
                        _add_issue({
                            "rule": "目录标题字号",
                            "severity": "中",
                            "location": f"第{idx+1}段：「{text[:30]}」",
                            "expected": "Times New Roman 四号（14pt）",
                            "actual": f"当前字号 {_pt_to_zh(sz)}",
                            "suggestion": "目录标题应为 Times New Roman 四号",
                            "zone": "目录",
                        })
                    font_name = _get_effective_font(_first_nonempty_run(para), para, has_chinese=False) or ""
                    if font_name and "times" not in font_name.lower():
                        _add_issue({
                            "rule": "目录标题字体",
                            "severity": "低",
                            "location": f"第{idx+1}段：「{text[:30]}」",
                            "expected": "Times New Roman",
                            "actual": font_name,
                            "suggestion": "目录标题字体应为 Times New Roman",
                            "zone": "目录",
                        })
                    if _is_run_bold_inherited(run, para) is False:
                        _add_issue({
                            "rule": "目录标题加粗",
                            "severity": "中",
                            "location": f"第{idx+1}段：「{text[:30]}」",
                            "expected": "加粗",
                            "actual": "未加粗",
                            "suggestion": "目录标题应加粗",
                            "zone": "目录",
                        })
                    if not _is_paragraph_centered(para):
                        _add_issue({
                            "rule": "目录标题居中",
                            "severity": "低",
                            "location": f"第{idx+1}段：「{text[:30]}」",
                            "expected": "居中",
                            "actual": "未居中",
                            "suggestion": "目录标题应居中",
                            "zone": "目录",
                        })
            else:
                # 目录条目：Times New Roman 小四
                run = _first_nonempty_run(para)
                if run and text.strip():
                    sz = _size_to_pt(run.font.size)
                    if sz and abs(sz - EXPECTED["toc_body_size_pt"]) > 1.0:
                        _add_issue({
                            "rule": "目录正文字号",
                            "severity": "低",
                            "location": f"第{idx+1}段：「{text[:40]}」",
                            "expected": "Times New Roman 小四（12pt）",
                            "actual": f"当前字号 {_pt_to_zh(sz)}",
                            "suggestion": "目录条目应小四",
                            "zone": "目录",
                        })
                    font_name = _get_effective_font(_first_nonempty_run(para), para, has_chinese=False) or ""
                    if font_name and "times" not in font_name.lower():
                        _add_issue({
                            "rule": "目录正文字体",
                            "severity": "低",
                            "location": f"第{idx+1}段：「{text[:40]}」",
                            "expected": "Times New Roman",
                            "actual": font_name,
                            "suggestion": "目录正文字体应为 Times New Roman",
                            "zone": "目录",
                        })
                    # 行距 1.5
                    ls = para.paragraph_format.line_spacing
                    if ls and abs(ls - EXPECTED["line_spacing"]) > 0.1:
                        _add_issue({
                            "rule": "目录正文行距",
                            "severity": "低",
                            "location": f"第{idx+1}段：「{text[:40]}」",
                            "expected": "1.5 倍",
                            "actual": f"{ls}",
                            "suggestion": "目录正文行距应为 1.5 倍",
                            "zone": "目录",
                        })
                    # 实词首字母大写（仅对英文标题条目）
                    tc_issue = _check_title_case(text)
                    if tc_issue:
                        _add_issue({
                            "rule": "目录条目实词大小写",
                            "severity": "低",
                            "location": f"第{idx+1}段：「{text[:50]}」",
                            "expected": "实词首字母大写（Title Case）",
                            "actual": tc_issue,
                            "suggestion": "标题各层实词首字母应大写，虚词小写",
                            "zone": "目录",
                        })
            continue

        # ---- 参考文献 ----
        if zone == "参考文献":
            found_references = True
            # Bibliography 标题检查（首段："Bibliography"）
            is_bib_header = bool(re.match(r"^\s*Bibliography\s*$", text, re.I))
            if is_bib_header:
                run = _first_nonempty_run(para)
                if run:
                    sz = _size_to_pt(run.font.size)
                    if sz and abs(sz - EXPECTED["bib_header_size_pt"]) > 1.5:
                        _add_issue({
                            "rule": "参考文献标题字号",
                            "severity": "中",
                            "location": f"第{idx+1}段：「{text}」",
                            "expected": "Times New Roman 四号（14pt）",
                            "actual": f"当前字号 {_pt_to_zh(sz)}",
                            "suggestion": "Bibliography 应为四号",
                            "zone": "参考文献",
                        })
                    if _is_run_bold_inherited(run, para) is False:
                        _add_issue({
                            "rule": "参考文献标题加粗",
                            "severity": "中",
                            "location": f"第{idx+1}段：「{text}」",
                            "expected": "加粗",
                            "actual": "未加粗",
                            "suggestion": "Bibliography 标题应加粗",
                            "zone": "参考文献",
                        })
                    if not _is_paragraph_centered(para):
                        _add_issue({
                            "rule": "参考文献标题居中",
                            "severity": "低",
                            "location": f"第{idx+1}段：「{text}」",
                            "expected": "居中",
                            "actual": "未居中",
                            "suggestion": "Bibliography 标题应居中",
                            "zone": "参考文献",
                        })
            else:
                # 参考文献正文：五号
                run = _first_nonempty_run(para)
                if run and text.strip():
                    sz = _size_to_pt(run.font.size)
                    if sz and abs(sz - EXPECTED["bib_size_pt"]) > 0.8:
                        _add_issue({
                            "rule": "参考文献正文字号",
                            "severity": "中",
                            "location": f"第{idx+1}段：「{text[:50]}」",
                            "expected": "五号（10.5pt），中文宋体五号、英文 Times New Roman 五号",
                            "actual": f"当前字号 {_pt_to_zh(sz)}",
                            "suggestion": "参考文献条目字号应为五号",
                            "zone": "参考文献",
                        })
                    # 检查末尾应为英文句号，不是中文句号
                    end = text.rstrip()
                    if end and end[-1] == "。":
                        _add_issue({
                            "rule": "参考文献末尾标点",
                            "severity": "低",
                            "location": f"第{idx+1}段：「{text[:50]}」",
                            "expected": "以英文句号 '.' 结束",
                            "actual": "以中文句号 '。' 结束",
                            "suggestion": "参考文献末尾应为英文句号",
                            "zone": "参考文献",
                        })
                    # 检查类型标识正则（宽松匹配，有问题发警告）
                    ref_type = _get_reference_type(text)
                    if ref_type and ref_type in REF_TYPE_PATTERNS:
                        pat = REF_TYPE_PATTERNS[ref_type]
                        if not re.search(pat, text):
                            _add_issue({
                                "rule": "参考文献类型格式",
                                "severity": "低",
                                "location": f"第{idx+1}段：「{text[:50]}」",
                                "expected": f"[{ref_type}] 类型标准格式（含必填字段）",
                                "actual": "字段可能缺失或不规范",
                                "suggestion": f"参照规范中的 [{ref_type}] 示例修正",
                                "zone": "参考文献",
                            })
            continue  # 参考文献不再走下面的通用正文逻辑

        # ---- 附录 ----
        if zone == "附录":
            # 检查 Appendix 标签字号/加粗
            if re.match(r"^\s*Appendix\s+[A-Z]", text):
                run = para.runs[0] if para.runs else None
                if run:
                    sz = _size_to_pt(run.font.size)
                    if sz and abs(sz - 14) > 1.5:
                        _add_issue({
                            "rule": "Appendix 标题字号",
                            "severity": "中",
                            "location": f"第{idx+1}段：「{text}」",
                            "expected": "Times New Roman 四号加粗",
                            "actual": f"当前字号 {_pt_to_zh(sz)}",
                            "suggestion": "附录标题应 Times New Roman 四号加粗",
                            "zone": "附录",
                        })
                    if _is_run_bold_inherited(run, para) is False:
                        _add_issue({
                            "rule": "Appendix 标题加粗",
                            "severity": "中",
                            "location": f"第{idx+1}段：「{text}」",
                            "expected": "加粗",
                            "actual": "未加粗",
                            "suggestion": "附录标题需加粗",
                            "zone": "附录",
                        })
            # 附录正文 = 同正文规范（小四 1.5倍行距）
            # 后续统一在"正文类"逻辑里检查

        # ---- 致谢 ----
        if zone == "致谢":
            # 致谢标题段检查（"Acknowledgements" 四号 加粗 居中 Times New Roman）
            is_ack_header = bool(re.match(r"^\s*Acknowledgement[s]?\s*$", text, re.I))
            if is_ack_header:
                run = _first_nonempty_run(para)
                if run:
                    sz = _size_to_pt(run.font.size)
                    if sz and abs(sz - EXPECTED["bib_header_size_pt"]) > 1.5:
                        _add_issue({
                            "rule": "致谢标题字号",
                            "severity": "中",
                            "location": f"第{idx+1}段：「{text}」",
                            "expected": "Times New Roman 四号（14pt）",
                            "actual": f"当前字号 {_pt_to_zh(sz)}",
                            "suggestion": "致谢标题应为四号",
                            "zone": "致谢",
                        })
                    if _is_run_bold_inherited(run, para) is False:
                        _add_issue({
                            "rule": "致谢标题加粗",
                            "severity": "中",
                            "location": f"第{idx+1}段：「{text}」",
                            "expected": "加粗",
                            "actual": "未加粗",
                            "suggestion": "致谢标题应加粗",
                            "zone": "致谢",
                        })
                    if not _is_paragraph_centered(para):
                        _add_issue({
                            "rule": "致谢标题居中",
                            "severity": "低",
                            "location": f"第{idx+1}段：「{text}」",
                            "expected": "居中",
                            "actual": "未居中",
                            "suggestion": "致谢标题应居中",
                            "zone": "致谢",
                        })
            # 致谢正文 = 同正文规范（已有通用逻辑处理）
            # 不 continue，让其走后面的正文类检查

        # ---- 成绩评定表：跳过所有检查 ----
        if zone == "成绩评定表":
            continue

        # ---- 章节标题（v2：用预计算的 heading_level_by_idx）----
        level = heading_level_by_idx.get(idx)
        if level == 1 and para.runs:
            heading_count[1] += 1
            run = _first_nonempty_run(para) or para.runs[0]
            sz = _consensus_size(para)
            if sz and abs(sz - EXPECTED["h1_size_pt"]) > 1.0:
                _add_issue({
                    "rule": "一级标题字号",
                    "severity": "中",
                    "location": f"第{idx+1}段：「{text[:40]}」",
                    "expected": "Times New Roman 四号加粗、左顶格",
                    "actual": f"当前字号 {_pt_to_zh(sz)}",
                    "suggestion": "一级标题应用 Times New Roman 四号加粗、左顶格。正确格式示例：「2  Theory and Methodology」（Times New Roman 四号加粗、左顶格）",
                    "zone": zone,
                })
            if _is_run_bold_inherited(run, para) is False:
                _add_issue({
                    "rule": "一级标题加粗",
                    "severity": "中",
                    "location": f"第{idx+1}段：「{text[:40]}」",
                    "expected": "Times New Roman 四号加粗",
                    "actual": "当前未加粗",
                    "suggestion": "一级标题必须加粗。正确格式示例：「2  Theory and Methodology」（Times New Roman 四号加粗、左顶格）",
                    "zone": zone,
                })
            # 字体检查
            font_name = _get_effective_font(run, para, has_chinese=False) or ""
            if font_name and "times" not in font_name.lower():
                _add_issue({
                    "rule": "一级标题字体",
                    "severity": "低",
                    "location": f"第{idx+1}段：「{text[:40]}」",
                    "expected": "Times New Roman",
                    "actual": font_name,
                    "suggestion": "一级标题字体应为 Times New Roman",
                    "zone": zone,
                })
            # Title Case 检查
            tc_issue = _check_title_case(text)
            if tc_issue:
                _add_issue({
                    "rule": "一级标题实词大小写",
                    "severity": "低",
                    "location": f"第{idx+1}段：「{text[:40]}」",
                    "expected": "实词首字母大写（Title Case）",
                    "actual": tc_issue,
                    "suggestion": "标题各层实词首字母应大写，虚词小写",
                    "zone": zone,
                })
            # 章节另起一页（仅 L1 检查；L2/L3 不要求）
            if not _has_page_break_before(para):
                has_visual = _has_visual_break_before(para)
                _add_issue({
                    "rule": "章节另起一页",
                    "severity": "低",
                    "location": f"第{idx+1}段：「{text[:40]}」",
                    "expected": "每章（一级标题）另起一页（规范要求）",
                    "actual": "未发现分页符"
                    + ("  [视觉上有空行分隔]" if has_visual else ""),
                    "suggestion": (
                        "在 L1 标题前插入分页符："
                        "1) Word 中光标定位到 L1 标题段开头"
                        "2) 快捷键 Ctrl+Enter 插入分页符"
                        "（注意：仅靠段前空行不算正式另起一页，"
                        "导师要求严格时仍需分页符）"
                    ),
                    "zone": zone,
                })
            # 一级标题应对齐：左顶格（不应居中）
            if _is_paragraph_centered(para):
                _add_issue({
                    "rule": "一级标题对齐",
                    "severity": "中",
                    "location": f"第{idx+1}段：「{text[:40]}」",
                    "expected": "左顶格",
                    "actual": "居中",
                    "suggestion": "一级标题应左顶格",
                    "zone": zone,
                })
        elif level == 2 and para.runs:
            heading_count[2] += 1
            run = _first_nonempty_run(para) or para.runs[0]
            sz = _consensus_size(para)
            if sz and abs(sz - EXPECTED["h2_size_pt"]) > 0.8:
                _add_issue({
                    "rule": "二级标题字号",
                    "severity": "中",
                    "location": f"第{idx+1}段：「{text[:40]}」",
                    "expected": "Times New Roman 小四（12pt）加粗",
                    "actual": f"当前字号 {_pt_to_zh(sz)}",
                    "suggestion": "二级标题应用小四加粗",
                    "zone": zone,
                })
            if _is_run_bold_inherited(run, para) is False:
                _add_issue({
                    "rule": "二级标题加粗",
                    "severity": "中",
                    "location": f"第{idx+1}段：「{text[:40]}」",
                    "expected": "加粗",
                    "actual": "当前未加粗",
                    "suggestion": "二级标题必须加粗",
                    "zone": zone,
                })
            font_name = _get_effective_font(run, para, has_chinese=False) or ""
            if font_name and "times" not in font_name.lower():
                _add_issue({
                    "rule": "二级标题字体",
                    "severity": "低",
                    "location": f"第{idx+1}段：「{text[:40]}」",
                    "expected": "Times New Roman",
                    "actual": font_name,
                    "suggestion": "二级标题字体应为 Times New Roman",
                    "zone": zone,
                })
            tc_issue = _check_title_case(text)
            if tc_issue:
                _add_issue({
                    "rule": "二级标题实词大小写",
                    "severity": "低",
                    "location": f"第{idx+1}段：「{text[:40]}」",
                    "expected": "实词首字母大写",
                    "actual": tc_issue,
                    "suggestion": "标题实词首字母应大写",
                    "zone": zone,
                })
            # 二级标题应对齐：左顶格（不应居中）
            if _is_paragraph_centered(para):
                _add_issue({
                    "rule": "二级标题对齐",
                    "severity": "中",
                    "location": f"第{idx+1}段：「{text[:40]}」",
                    "expected": "左顶格",
                    "actual": "居中",
                    "suggestion": "二级标题应左顶格",
                    "zone": zone,
                })
        elif level == 3 and para.runs:
            heading_count[3] += 1
            run = _first_nonempty_run(para) or para.runs[0]
            sz = _consensus_size(para)
            if sz and abs(sz - EXPECTED["h3_size_pt"]) > 0.8:
                _add_issue({
                    "rule": "三级标题字号",
                    "severity": "中",
                    "location": f"第{idx+1}段：「{text[:40]}」",
                    "expected": "Times New Roman 小四（12pt）",
                    "actual": f"当前字号 {_pt_to_zh(sz)}",
                    "suggestion": "三级标题应用小四",
                    "zone": zone,
                })
            if _is_run_bold_inherited(run, para) is True:
                _add_issue({
                    "rule": "三级标题加粗",
                    "severity": "低",
                    "location": f"第{idx+1}段：「{text[:40]}」",
                    "expected": "不加粗",
                    "actual": "当前加粗",
                    "suggestion": "三级标题不应加粗",
                    "zone": zone,
                })
            font_name = _get_effective_font(_first_nonempty_run(para), para, has_chinese=False) or ""
            if font_name and "times" not in font_name.lower():
                _add_issue({
                    "rule": "三级标题字体",
                    "severity": "低",
                    "location": f"第{idx+1}段：「{text[:40]}」",
                    "expected": "Times New Roman",
                    "actual": font_name,
                    "suggestion": "三级标题字体应为 Times New Roman",
                    "zone": zone,
                })
            tc_issue = _check_title_case(text)
            if tc_issue:
                _add_issue({
                    "rule": "三级标题实词大小写",
                    "severity": "低",
                    "location": f"第{idx+1}段：「{text[:40]}」",
                    "expected": "实词首字母大写",
                    "actual": tc_issue,
                    "suggestion": "标题实词首字母应大写",
                    "zone": zone,
                })
            # 三级标题应对齐：左顶格（不应居中）
            if _is_paragraph_centered(para):
                _add_issue({
                    "rule": "三级标题对齐",
                    "severity": "中",
                    "location": f"第{idx+1}段：「{text[:40]}」",
                    "expected": "左顶格",
                    "actual": "居中",
                    "suggestion": "三级标题应左顶格",
                    "zone": zone,
                })

        # ---- 图题/表题/例证 ----
        is_cap, cap_type = _is_figure_or_table_caption(text)
        is_example = _is_example_caption(text)
        if is_cap and para.runs:
            # 只在段落**显式**设置过字号时才检查。
            # 若所有 run 都未显式设 font.size（继承自样式/默认），
            # 不要用 pPr.rPr.sz 兜底（会把"未设置"误判为"小四"）。
            sz = _consensus_size(para)
            if sz is not None and abs(sz - EXPECTED["caption_size_pt"]) > 0.8:
                _add_issue({
                    "rule": f"{cap_type}字号",
                    "severity": "低",
                    "location": f"第{idx+1}段：「{text[:30]}」",
                    "expected": "Times New Roman 五号",
                    "actual": f"当前字号 {_pt_to_zh(sz)}",
                    "suggestion": f"{cap_type}应为 Times New Roman 五号居中。正确格式示例：「Table 1 标题内容」（五号居中）",
                    "zone": zone,
                })
            # 字体检查
            font_name = _get_effective_font(run, para, has_chinese=False) or ""
            if font_name and "times" not in font_name.lower():
                _add_issue({
                    "rule": f"{cap_type}字体",
                    "severity": "低",
                    "location": f"第{idx+1}段：「{text[:30]}」",
                    "expected": "Times New Roman",
                    "actual": font_name,
                    "suggestion": f"{cap_type}字体应为 Times New Roman",
                    "zone": zone,
                })
            # 居中
            if not _is_paragraph_centered(para):
                _add_issue({
                    "rule": f"{cap_type}对齐",
                    "severity": "低",
                    "location": f"第{idx+1}段：「{text[:30]}」",
                    "expected": "居中",
                    "actual": "未居中",
                    "suggestion": f"{cap_type}应居中（图题在图下方、表题在表上方）",
                    "zone": zone,
                })
        elif is_example and para.runs:
            # Example 例证检查（规范允许小四，不强制）
            pass  # 实际论文常用小四，强行报五号会大量误报

        # ---- 正文段落（v2：用 _is_body_paragraph 放宽过滤）----
        if level is None and _is_body_paragraph(text, para, zone):
            body_paragraphs += 1
            run = _first_nonempty_run(para) if para.runs else None
            has_chinese = bool(re.search(r"[一-龥]", text))
            if run:
                expected_font = EXPECTED["body_font_zh"] if has_chinese else EXPECTED["body_font_en"]
                actual_font = _get_effective_font(run, para, has_chinese)
                actual_size = _consensus_size(para)
                if actual_font and actual_font != expected_font:
                    _add_issue({
                        "rule": "正文字体",
                        "severity": "中",
                        "location": f"第{idx+1}段：「{text[:30]}...」",
                        "expected": f"{expected_font}（{'中文' if has_chinese else '英文/数字'}用）",
                        "actual": f"当前字体 {actual_font}",
                        "suggestion": f"正文{'中文段落应使用宋体，英文/数字使用 Times New Roman' if has_chinese else '段落中英文/数字应使用 Times New Roman'}。正确格式示例：「本研究采用文本分析法……」（宋体小四，中英文混排时英文自动用 Times New Roman）",
                        "zone": zone,
                    })
                if actual_size is not None and abs(actual_size - EXPECTED["body_size_pt"]) > 1.0:
                    _add_issue({
                        "rule": "正文字号",
                        "severity": "中",
                        "location": f"第{idx+1}段：「{text[:30]}...」",
                        "expected": "小四",
                        "actual": f"当前字号 {_pt_to_zh(actual_size)}",
                        "suggestion": "正文统一使用小四号。正确格式示例：「本研究采用文本分析法……」（宋体小四）",
                        "zone": zone,
                    })
            ls = para.paragraph_format.line_spacing
            if ls is not None and abs(ls - EXPECTED["line_spacing"]) > 0.1:
                _add_issue({
                    "rule": "行距",
                    "severity": "中",
                    "location": f"第{idx+1}段：「{text[:30]}...」",
                    "expected": f"{EXPECTED['line_spacing']} 倍行距",
                    "actual": f"{ls} 倍",
                    "suggestion": "正文统一 1.5 倍行距（图表除外）",
                    "zone": zone,
                })

    # ---- 3d. 内文元素统计（v2 新增）----
    body_range = stats.get("body_zone_range", (0, 0))
    if body_range[0] and body_range[1] > body_range[0]:
        # body_zone_range 是 1-based 段号，转回 0-based 段落索引
        inline = _count_inline_elements(
            doc.paragraphs,
            (body_range[0] - 1, body_range[1]),
        )
        stats["inline_citations_total"] = inline["inline_citations"]
        stats["inline_citation_paragraphs"] = inline["citation_paragraphs"]
        stats["figures_count"] = inline["figures"]
        stats["tables_count"] = inline["tables"]
        stats["examples_count"] = inline["examples"]
        stats["cite_brackets_total"] = inline["cite_brackets_total"]
        stats["cite_brackets_superscript"] = inline["cite_brackets_superscript"]
        # 上标引用率检查：< 30% 时警告
        total = inline["cite_brackets_total"]
        sup = inline["cite_brackets_superscript"]
        if total >= 5:  # 至少 5 个引用才检查，避免小样本误报
            rate = sup / total if total else 0
            if rate < 0.3:
                _add_issue({
                    "rule": "引用上标格式",
                    "severity": "中",
                    "location": f"正文（共 {total} 处方括号引用）",
                    "expected": "参考文献序号 [1] 应作为上标（superscript）",
                    "actual": f"仅 {sup}/{total}（{int(rate*100)}%）使用了上标",
                    "suggestion": "选中 [1] 这样的引用 → 字体 → 上标（或格式 → 文本效果 → 上标）",
                    "zone": "正文",
                })

    # ---- 4a. 页码设置检查 ----
    _check_page_numbers(doc, zones, _add_issue)

    # ---- 4. 参考文献专项检查 ----
    ref_range = zones.get("参考文献", (None, None))
    if ref_range[0] is not None:
        ref_count = 0
        no_type_marker = []
        # 遍历 body 元素（参考文献区可能跨表格）
        for ci in range(ref_range[0], ref_range[1]):
            if ci >= len(all_children):
                break
            child = all_children[ci]
            tag = child.tag.split('}')[-1]
            t = ''.join(child.itertext()).strip()
            if not t:
                continue
            if _is_reference_entry(t):
                ref_count += 1
                if not re.search(r"\[[A-Z]\]", t):
                    no_type_marker.append((ci, t[:60]))
        stats["reference_count"] = ref_count
        if ref_count < 15:
            _add_issue({
                "rule": "参考文献数量",
                "severity": "高",
                "location": f"参考文献区段（{ref_section_start + 1}段后）",
                "expected": "不少于 15 篇",
                "actual": f"检测到 {ref_count} 条",
                "suggestion": "请补充至至少 15 篇（涵盖近 5 年文献）",
                "zone": "参考文献",
            })
        if no_type_marker:
            for j, t in no_type_marker[:3]:
                _add_issue({
                    "rule": "参考文献缺少文献类型标识",
                    "severity": "中",
                    "location": f"第{j+1}段",
                    "expected": "题名后标注 [J]/[M]/[R]/[D] 等文献类型",
                    "actual": t,
                    "suggestion": "示例：作者. 题名[J]. 刊名, 年. / 作者. 题名[M]. 出版地: 出版者, 年.",
                    "zone": "参考文献",
                })

    # ---- 5. 必需段落检测 ----
    # 扉页字段完整性检查（按规范必填）
    cover_text = stats.get("cover_all_text", "")
    if found_en_title_on_cover and cover_text:
        for field in EXPECTED.get("frontpage_required_fields", []):
            if field not in cover_text:
                _add_issue({
                    "rule": "扉页必填字段",
                    "severity": "中",
                    "location": "扉页",
                    "expected": f"扉页必填字段「{field}」",
                    "actual": "未找到",
                    "suggestion": f"扉页（扉页为含英文论文题目的那页）应包含字段「{field}」",
                    "zone": "封面/扉页",
                })
    if not found_zh_title_on_cover:
        _add_issue({
            "rule": "中文封面标题",
            "severity": "高",
            "location": "封面（第1页）",
            "expected": "中文论文题目（封面第一页，宋体二号加粗居中）",
            "actual": "未检测到（封面只有英文标题）",
            "suggestion": "规范要求封面和扉页均包含中文/英文双标题。中文封面示例：「认知负荷视角下大疆Neo 2用户手册翻译研究」（宋体小二号加粗居中）",
            "zone": "封面/扉页",
        })
    if not found_en_title_on_cover:
        _add_issue({
            "rule": "英文封面标题",
            "severity": "高",
            "location": "封面",
            "expected": "英文论文题目（封面）",
            "actual": "未检测到",
            "suggestion": "商务英语专业必须有英文封面标题（Times New Roman 二号加粗居中）",
            "zone": "封面/扉页",
        })
    if not found_pledge:
        _add_issue({
            "rule": "学生承诺书",
            "severity": "高",
            "location": "前置部分",
            "expected": "学生承诺书",
            "actual": "未检测到",
            "suggestion": "封面后应包含学生承诺书（宋体三号加粗居中）",
            "zone": "封面/扉页",
        })
    if not found_abstract_zh:
        _add_issue({
            "rule": "中文摘要",
            "severity": "高",
            "location": "前置部分",
            "expected": "中文摘要（关键词 3-5 个）",
            "actual": "未检测到",
            "suggestion": "请确认包含'摘要'章节",
            "zone": "中文摘要",
        })
    if not found_abstract_en:
        _add_issue({
            "rule": "英文摘要",
            "severity": "高",
            "location": "前置部分",
            "expected": "Abstract（Keywords 3-5 个）",
            "actual": "未检测到",
            "suggestion": "请确认包含'Abstract'章节",
            "zone": "英文摘要",
        })
    if not found_toc:
        _add_issue({
            "rule": "目录",
            "severity": "中",
            "location": "前置部分",
            "expected": "Table of Contents / 目录",
            "actual": "未检测到",
            "suggestion": "用 Word 插入→引用→索引和目录自动生成",
            "zone": "目录",
        })
    if not found_references:
        _add_issue({
            "rule": "参考文献",
            "severity": "高",
            "location": "正文后",
            "expected": "参考文献 / Bibliography",
            "actual": "未检测到",
            "suggestion": "正文后必须有参考文献章节，不少于 15 篇",
            "zone": "参考文献",
        })

    # ---- 6. 字数统计 ----
    total_chars = 0
    body_chars = 0
    table_chars = 0
    skip_zones = {"封面/扉页", "学生承诺书", "中文摘要", "英文摘要",
                   "目录", "参考文献"}
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        if _is_reference_entry(text) or text in (
            "Bibliography", "参考文献", "Appendix", "Acknowledgements",
            "Table of Contents", "Abstract", "目录", "摘要", "学生承诺书"
        ):
            continue
        chars = len([c for c in text if c.isalnum() or '一' <= c <= '鿿'])
        total_chars += chars
        if _zone_of(i, zones) not in skip_zones:
            body_chars += chars
    # 也统计表格内字符（正文有时写在表格里）
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    t = p.text.strip()
                    if t:
                        table_chars += len([c for c in t if c.isalnum() or '一' <= c <= '鿿'])
    body_chars_total = body_chars + table_chars
    stats["approx_word_count"] = body_chars
    stats["approx_word_count_with_tables"] = body_chars_total
    stats["total_chars_including_front"] = total_chars
    if body_chars_total < 8000:
        _add_issue({
            "rule": "正文字数",
            "severity": "高",
            "location": "全文",
            "expected": "经管文法类正文 ≥ 10000 字 / 理论研究类 ≥ 15000 字",
            "actual": f"约 {body_chars_total} 字（段 {body_chars} + 表 {table_chars}，不含摘要/目录/参考文献）",
            "suggestion": "正文（不含摘要/目录/参考文献）需达到规定字数",
            "zone": "正文",
        })

    # ---- 7. 统计汇总 ----
    stats["total_paragraphs"] = total_paragraphs
    stats["body_paragraphs"] = body_paragraphs
    stats["headings_by_level"] = dict(heading_count)
    stats["issues_by_severity"] = dict(Counter(i["severity"] for i in issues))
    stats["total_issues"] = len(issues)
    stats["issue_per_para"] = {k: len(v) for k, v in issue_per_para.items()}

    # ---- 8. 评级 ----
    high = stats["issues_by_severity"].get("高", 0)
    medium = stats["issues_by_severity"].get("中", 0)
    low = stats["issues_by_severity"].get("低", 0)
    if high == 0 and medium <= 2:
        grade = "A 优秀"
    elif high == 0 and medium <= 5:
        grade = "B 良好"
    elif high <= 2:
        grade = "C 及格"
    else:
        grade = "D 需大幅修改"

    return {
        "file": docx_path,
        "school": "仲恺农业工程学院外国语学院",
        "stats": stats,
        "grade": grade,
        "issues": issues,
        "issue_per_para": issue_per_para,
        "zones": zones,
        "severity_color": SEVERITY_COLOR,
    }


# === 可视标注辅助函数 ===

def mark_doc_with_issues(docx_path: str, output_path: str, report: dict) -> str:
    """在 docx 中给有问题的段落加高亮背景 + 批注文字。
    每条批注包含两部分：
      ⚠️ [严重度] 问题名称：当前实际值
      ✅ 正确格式：建议的正确格式（含示例）

    返回输出文件路径。
    """
    doc = Document(docx_path)
    issue_per_para = report.get("issue_per_para", {})
    from docx.shared import RGBColor
    from docx.enum.text import WD_COLOR_INDEX

    for pi, iss_list in issue_per_para.items():
        if pi >= len(doc.paragraphs):
            continue
        para = doc.paragraphs[pi]
        # 取最高严重度
        sev_order = {"低": 1, "中": 2, "高": 3}
        top_sev = max((i["severity"] for i in iss_list), key=lambda s: sev_order.get(s, 0))
        color_map = {"高": WD_COLOR_INDEX.RED,
                     "中": WD_COLOR_INDEX.YELLOW,
                     "低": WD_COLOR_INDEX.GRAY_25}
        # 给该段所有 run 染色
        for run in para.runs:
            run.font.highlight_color = color_map[top_sev]
        # 在段落末尾追加红色问题描述
        problem_text = (
            f"  ⚠️ [{top_sev}] "
            + "; ".join(f"{i['rule']}（当前：{i['actual']}）" for i in iss_list)
        )
        # 紧接着追加绿色正确格式建议
        suggestion_text = (
            "  ✅ 正确格式："
            + " | ".join(i["suggestion"] for i in iss_list if i.get("suggestion"))
        )
        if problem_text:
            problem_run = para.add_run(problem_text)
            problem_run.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)
            problem_run.bold = True
        if suggestion_text:
            sug_run = para.add_run(suggestion_text)
            sug_run.font.color.rgb = RGBColor(0x2E, 0x7D, 0x32)
            sug_run.bold = False

    doc.save(output_path)
    return output_path


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("用法: python format_checker.py <docx文件>")
        sys.exit(1)
    result = check_document(sys.argv[1])
    print(f"\n{'='*60}")
    print(f"文件: {result['file']}")
    print(f"评级: {result['grade']}")
    print(f"问题总数: {result['stats']['total_issues']}")
    sev = result['stats']['issues_by_severity']
    print(f"  高: {sev.get('高', 0)}  中: {sev.get('中', 0)}  低: {sev.get('低', 0)}")
    print(f"\n区域识别:")
    for z, (s, e) in result['zones'].items():
        print(f"  {z}: 第{s+1}-{e}段")
    s = result['stats']
    print(f"\n统计: 总段落{s['total_paragraphs']}, "
          f"正文段{s['body_paragraphs']}, "
          f"正文字符约{s['approx_word_count']}, "
          f"参考文献{s.get('reference_count', 0)}篇")
    if s.get("body_zone_range") and s["body_zone_range"][1] > s["body_zone_range"][0]:
        br_s, br_e = s["body_zone_range"]
        print(f"  正文: 第{br_s}-{br_e}段")
    if "inline_citations_total" in s:
        print(f"  正文引用: {s['inline_citations_total']} 处 "
              f"(分布在 {s.get('inline_citation_paragraphs', 0)} 段)")
    if "figures_count" in s or "tables_count" in s:
        print(f"  图表: 图 {s.get('figures_count', 0)} 个 · "
              f"表 {s.get('tables_count', 0)} 个")
    if s.get("examples_count", 0) > 0:
        print(f"  例证: {s['examples_count']} 处")
    print(f"章节标题: {s['headings_by_level']}")
    if s.get("headings"):
        print(f"章节标题明细 (前 10):")
        shown = 0
        for lv in sorted({h['level'] for h in s['headings']}):
            for h in s['headings']:
                if h['level'] != lv:
                    continue
                print(f"  L{lv} 第{h['idx']+1}段: {h['text']}")
                shown += 1
                if shown >= 10:
                    break
            if shown >= 10:
                break
    print(f"\n{'='*60}\n问题清单（前25条）:\n")
    for i, iss in enumerate(result["issues"][:25], 1):
        print(f"{i}. [{iss['severity']}] {iss['rule']} | 区域: {iss.get('zone', '?')}")
        print(f"   位置: {iss['location']}")
        print(f"   期望: {iss['expected']}")
        print(f"   实际: {iss['actual']}")
        print(f"   建议: {iss['suggestion']}")
        print()
