"""
选题分析模块 v0.3
基于 910 条真实往届选题的统计规律重写：
- 理论表达模式：视角下/基于...理论/视域下/理论框架下
- 研究对象模式：以...为例/破折号+以...为例/直接放主标题
- 关键词热度
- 热门品牌/案例对象
"""
import json
import re
from collections import Counter
from pathlib import Path

from paths import resource_path

DATA_DIR = resource_path("data")
TOPICS_PATH = DATA_DIR / "topics.json"

# 商务英语领域词典（中文正向最大匹配）
DOMAIN_DICT = [
    # 理论框架（按真实热度排序）
    "目的论", "文化维度", "功能对等", "生态翻译学", "多模态话语分析",
    "关联理论", "顺应论", "接受美学", "言语行为", "高低语境",
    "礼貌原则", "霍夫斯泰德文化维度理论", "文化维度理论", "跨文化交际",
    "建构主义", "批评话语分析", "视觉语法", "元话语", "文化图式",
    "互文性", "语用学", "情感过滤假说", "TPACK理论", "监控理论",
    "认知负荷理论", "文化适应理论", "企业识别系统理论", "纽马克翻译二分法",
    "语境教学法", "形态学理论", "评价理论", "共情理论", "原型理论",
    "拉斯韦尔5W理论", "互文性理论", "语用学理论", "语言相对论",
    "言语行为理论", "功能对等理论", "目的论视角",
    # 商务场景
    "跨境电商", "电商直播", "直播带货", "商务谈判", "商务函电",
    "商务沟通", "商务话语", "公示语", "字幕翻译", "菜名翻译",
    "广告语", "广告翻译", "合同翻译", "产品描述", "用户手册",
    "说明书", "术语规范化", "英译", "汉译", "英译本",
    "归化", "异化", "创译法", "外宣翻译", "中医术语",
    "法律文本", "格式条款", "外宣", "字幕", "菜名", "广告",
    "说明书", "合同", "谈判",
    # 教学/学习
    "人才培养", "教学法", "二语习得", "词汇习得", "学习动机",
    "学习兴趣", "口语教学", "词汇教学", "语法教学", "英语教学",
    "外语教学", "批改网", "自动批改", "情感过滤", "商务英语专业",
    # 品牌/案例（精选 Top 30）
    "可口可乐", "苹果公司", "苹果", "耐克", "华为", "亚马逊",
    "阿里巴巴", "奔驰", "宝马", "麦当劳", "星巴克", "迪士尼",
    "Lululemon", "宜家", "美的", "青岛啤酒", "汉堡王", "农夫山泉",
    "广汽集团", "茶颜悦色", "王者荣耀", "原神", "黑神话", "英雄联盟",
    "张家界", "故宫", "博物馆", "孔子学院", "大疆", "DJI", "SHEIN",
    "TikTok", "淘宝", "天猫", "京东", "拼多多", "抖音", "小红书",
    # 体裁/对象
    "古典诗词", "国漫", "电影", "电视剧", "美剧", "影视作品",
    "外交部", "政府工作报告", "发言", "演讲", "品牌联名",
    # 行业/领域
    "一带一路", "粤港澳", "大湾区", "本土化", "国际化",
    "教学", "课程", "教材", "实习", "现状", "策略",
]
DOMAIN_DICT = list(set(DOMAIN_DICT))
DOMAIN_DICT.sort(key=len, reverse=True)

# 商务英语专业选题 4 大范围（扩充关键词）
SCOPE_CATEGORIES = {
    "商务环境中的语言研究": {
        "keywords": [
            "话语分析", "语用", "语篇", "批评话语", "多模态", "视觉语法",
            "元话语", "缓和语", "模糊语", "礼貌", "言语行为", "语用功能",
            "互文性", "人际意义", "系统功能", "人际功能", "概念功能",
            "语料库", "词汇特点", "话语策略", "语言策略", "语言特征",
            "语言景观", "商务话语", "企业话语", "营销话语", "广告话语",
            "演讲", "发言", "新闻发布会", "公关", "声明", "道歉",
            "新闻话语", "政治话语", "外交话语", "商标", "品牌话语",
            "多模态隐喻", "积极话语分析", "话语分析视角",
        ],
        "description": "聚焦商务语境下的语言使用特征与功能",
        "examples": [
            "批评话语分析视角下美的集团ESG报告的企业身份建构研究",
            "评价理论态度系统视角下马云演讲的积极话语分析",
            "视觉语法视角下耐克广告海报的多模态话语分析",
        ],
    },
    "商务环境中的翻译研究": {
        "keywords": [
            "翻译", "英译", "汉译", "译本", "归化", "异化", "创译",
            "翻译策略", "翻译方法", "外宣翻译", "字幕翻译", "公示语翻译",
            "合同翻译", "广告翻译", "产品翻译", "说明书", "用户手册",
            "菜名翻译", "术语翻译", "中医翻译", "法律翻译",
            "影视翻译", "字幕", "配音", "本地化", "国际传播",
            "功能对等", "目的论", "生态翻译学", "接受美学",
            "翻译特点", "翻译策略研究", "翻译实践", "翻译报告",
        ],
        "description": "聚焦商务文本/场景的翻译实践与策略",
        "examples": [
            "目的论视角下电影字幕英汉翻译策略研究——以《超能陆战队》为例",
            "生态翻译学视角下的翻译策略研究——以《茶经》中的茶文化翻译为例",
            "功能对等理论下鲁菜的英译方法研究",
        ],
    },
    "跨文化商务交际研究": {
        "keywords": [
            "跨文化", "跨文化交际", "跨文化传播", "跨文化管理",
            "跨文化适应", "跨文化营销", "跨文化谈判", "跨文化沟通",
            "文化差异", "文化维度", "高低语境", "霍夫斯泰德",
            "文化冲突", "文化适应", "文化整合", "文化负载词",
            "中美", "中俄", "中外", "文化图式", "文化定势",
            "国际商务", "海外市场", "品牌国际化", "本土化",
            "全球化", "国际市场", "中外合作", "国际化经营",
        ],
        "description": "聚焦跨文化情境下的商务沟通与传播",
        "examples": [
            "基于文化维度理论探究文化差异对中美商务谈判的影响与解决策略",
            "基于文化适应理论跨国企业在目标市场的文化整合策略研究——以宜家家居(中国)为例",
            "文化维度理论下探讨中美英雄电影的文化特点——对比分析《湄公河行动》和《碟中谍5》",
        ],
    },
    "商务英语学习与教学研究": {
        "keywords": [
            "教学", "教学法", "教学策略", "教学模式", "教学实践",
            "学习", "学习策略", "学习动机", "学习效果", "习得",
            "二语习得", "词汇习得", "口语习得", "课程", "教材",
            "TPACK", "建构主义", "情感过滤", "情感过滤假说",
            "监控理论", "批改网", "商务英语专业", "人才培养",
            "课程设置", "实训", "口语教学", "词汇教学",
            "语法教学", "翻译教学", "英语教育", "外语教育",
            "外语教学", "教学现状", "教学改革", "教师发展",
        ],
        "description": "聚焦商务英语教与学的理论与实践",
        "examples": [
            "商务英语专业学生技术知识培养现状及策略研究——以仲恺农业工程学院为例",
            "『克拉申情感过滤假说』下论学习兴趣对习得英语的重要性——以仲恺农业工程学院的学生为例",
            "监控理论视角下大学英语类专业语法教学探索",
        ],
    },
}

# 理论框架的常见表达模板（基于 910 条统计）
THEORY_PATTERNS = [
    (r"基于([一-龥A-Za-z]{2,15}(理论|学|假说))", "基于X理论/学"),
    (r"([一-龥A-Za-z]{2,15})理论(框架)?下", "X理论下"),
    (r"([一-龥A-Za-z]{2,15})视角(下)?", "X视角"),
    (r"([一-龥A-Za-z]{2,15})视域(下)?", "X视域"),
    (r"从([一-龥A-Za-z]{2,15})(角度|视角)", "从X角度/视角"),
]

# 研究对象的常见表达模板
OBJECT_PATTERNS = [
    (r"以([一-龥A-Za-z《（\(\)》《》\d]{2,30})为例", "以X为例"),
    (r"以([一-龥A-Za-z]{2,30})为语料", "以X为语料"),
    (r"以([一-龥A-Za-z]{2,30})为(研究对象|案例)", "以X为对象/案例"),
    (r"以([一-龥A-Za-z]{2,30})为(研究|分析)(材料|对象)", "以X为材料"),
    (r"对比分析(.+?)和(.+?)$", "对比分析A和B"),
    (r"——《(.+?)》", "破折号+书名号"),
]

STOPWORDS = {
    "的", "了", "是", "在", "和", "与", "及", "或", "一个", "这个", "那个",
    "如何", "怎么", "什么", "为什么", "是否", "a", "an", "the",
    "of", "in", "on", "at", "to", "for", "with", "from", "by",
    "and", "or", "study", "research", "analysis",
}


def _segment_chinese(text: str) -> list[str]:
    """正向最大匹配分词"""
    tokens = []
    i = 0
    n = len(text)
    while i < n:
        matched = None
        for word in DOMAIN_DICT:
            wlen = len(word)
            if i + wlen <= n and text[i:i + wlen] == word:
                matched = word
                break
        if matched:
            tokens.append(matched)
            i += len(matched)
        else:
            ch = text[i]
            if ch.isascii() and ch.isalnum():
                j = i
                while j < n and text[j].isascii() and text[j].isalnum():
                    j += 1
                tokens.append(text[i:j])
                i = j
            else:
                i += 1
    return tokens


def extract_keywords(title: str) -> list[str]:
    en_parts = re.findall(r"[A-Za-z][A-Za-z\s\-]*[A-Za-z]|[A-Za-z]", title)
    zh_only = re.sub(r"[A-Za-z][A-Za-z\s\-]*[A-Za-z]|[A-Za-z]", " ", title)
    zh_tokens = _segment_chinese(zh_only)
    en_tokens = []
    for ep in en_parts:
        words = ep.split()
        for w in words:
            wl = w.lower().strip("-")
            if wl in STOPWORDS or len(wl) < 3:
                continue
            en_tokens.append(w)
    all_tokens = zh_tokens + en_tokens
    seen, uniq = set(), []
    for t in all_tokens:
        tl = t.lower()
        if tl not in seen and tl not in STOPWORDS:
            seen.add(tl)
            uniq.append(t)
    return uniq[:10]


def detect_theory_expression(title: str) -> dict:
    """识别题目中的理论框架 + 表达方式
    返回：{has_theory: bool, theory_name: str, expression_pattern: str, examples: []}
    """
    hits = []
    title_lower = title

    # 已知理论名
    known_theories = [
        "功能对等理论", "功能对等", "目的论", "顺应论", "顺应理论",
        "礼貌原则", "跨文化交际", "跨文化传播", "多模态话语分析",
        "视觉语法", "批评话语分析", "批评话语", "语用学",
        "关联理论", "生态翻译学", "接受美学", "文化图式",
        "文化适应理论", "情感过滤假说", "情感过滤", "监控理论",
        "认知负荷理论", "TPACK", "TPACK理论", "建构主义",
        "言语行为理论", "言语行为", "企业识别系统",
        "纽马克", "语境教学法", "形态学理论", "评价理论",
        "共情理论", "原型理论", "高低语境", "文化维度",
        "霍夫斯泰德", "互文性", "元话语", "生态语言学",
    ]
    matched_theories = []
    for th in known_theories:
        if th in title:
            matched_theories.append(th)

    # 表达模式
    expression_patterns = []
    for pat, label in THEORY_PATTERNS:
        if re.search(pat, title):
            expression_patterns.append(label)

    # 提取理论主体（基于"基于X"模式）
    base_match = re.search(r"基于([一-龥A-Za-z]{2,15})", title)
    base_theory = base_match.group(1) if base_match else None

    perspective_match = re.search(r"([一-龥A-Za-z]{2,15})视角", title)
    perspective_theory = perspective_match.group(1) if perspective_match else None

    has_theory = bool(matched_theories or base_theory or perspective_theory)

    return {
        "has_theory": has_theory,
        "matched_theories": list(set(matched_theories)),
        "expression_patterns": list(set(expression_patterns)),
        "base_theory": base_theory,           # 从"基于X"提取的
        "perspective_theory": perspective_theory,  # 从"X视角"提取的
    }


def detect_object_expression(title: str) -> dict:
    """识别题目中的研究对象 + 表达方式
    返回：{has_object: bool, object_name: str, expression_pattern: str}
    """
    # 1. 以X为例
    case_match = re.search(r"以([^，。；、]{2,30}?)为例", title)
    # 2. 以X为语料/对象/案例
    corpus_match = re.search(r"以([^，。；、]{2,30}?)为(语料|对象|案例|研究材料)", title)
    # 3. 破折号后书名号
    book_match = re.search(r"《([^》]+)》", title)
    # 4. 末尾"——XX"（主标题+破折号+对象）
    dash_match = re.search(r"[—–-]{2}(.+?)$", title)

    obj = None
    pattern = None

    if case_match:
        obj = case_match.group(1).strip().rstrip("（(")
        pattern = "以X为例"
    elif corpus_match:
        obj = corpus_match.group(1).strip().rstrip("（(")
        pattern = "以X为语料/对象"
    elif book_match:
        obj = book_match.group(1)
        pattern = "书名号直接标注"
    elif dash_match:
        obj = dash_match.group(1).strip()
        pattern = "破折号引出对象"

    # 主标题末尾的对象（如《XX》书名/品牌名）
    if not obj:
        title_main = re.sub(r"[—–-]{2}.*$", "", title).strip()
        # 主标题里出现书名号也算
        if "《" in title_main and "》" in title_main:
            obj = re.search(r"《([^》]+)》", title_main).group(1)
            pattern = "主标题含书名号"

    return {
        "has_object": bool(obj),
        "object_name": obj,
        "expression_pattern": pattern or "未明确对象",
    }


def detect_direction(title: str) -> str:
    """基于 1032 条统计的方向识别（4 大类）"""
    title_lower = title.lower()
    scores = Counter()
    for direction, patterns in [
        ("翻译", ["翻译", "translation", "英译", "汉译", "译本",
                   "字幕", "归化", "异化", "外宣", "译介", "创译", "目的论",
                   "译配", "可译性"]),
        ("跨文化", ["跨文化", "cross-cultural", "高低语境", "文化维度",
                     "中外", "中美", "中俄", "海外", "本土化", "国际化",
                     "中西方", "中韩", "中英", "文化适应", "文化传播",
                     "文化营销", "品牌文化", "对外传播", "跨文化营销",
                     "文化冲突", "文化差异"]),
        ("话语分析", ["话语分析", "多模态", "视觉语法", "批评话语",
                       "discourse", "语篇", "系统功能", "人际意义",
                       "互文性", "多模态隐喻", "语用", "pragmatic",
                       "礼貌原则", "合作原则", "面子理论", "言语行为",
                       "元话语", "模糊限制语", "话语策略", "修辞",
                       "劝说", "说服", "概念隐喻", "符号学", "评价理论",
                       "态度系统", "语域"]),
        ("商务英语习得", ["教学", "习得", "学习", "二语", "外语教学",
                       "词汇习得", "词汇学习", "词汇记忆", "教学法", "课程",
                       "教材", "批改", "TPACK", "建构主义", "学习动机",
                       "学习兴趣", "输入假设", "Krashen", "POA",
                       "产出导向法", "写作能力"]),
    ]:
        for p in patterns:
            if p.lower() in title_lower:
                scores[direction] += 1
    if not scores:
        return "其他"
    return scores.most_common(1)[0][0]


def check_scope(title: str) -> dict:
    """4 大专业范围校验"""
    title_lower = title.lower()
    category_hits = {}
    for category, info in SCOPE_CATEGORIES.items():
        matched = []
        for kw in info["keywords"]:
            if kw.lower() in title_lower:
                matched.append(kw)
        if matched:
            category_hits[category] = matched
    sorted_cats = sorted(category_hits.items(), key=lambda x: -len(x[1]))
    if sorted_cats:
        top_cat, top_kws = sorted_cats[0]
        return {
            "in_scope": True,
            "matched_category": top_cat,
            "matched_keywords": top_kws,
            "assessment": f"符合「{top_cat}」",
            "icon": "✅",
            "description": SCOPE_CATEGORIES[top_cat]["description"],
        }
    return {
        "in_scope": False,
        "matched_category": "未匹配",
        "matched_keywords": [],
        "assessment": "⚠️ 不在 4 大专业范围内",
        "icon": "❌",
        "description": "题目未命中 4 大类关键词，建议调整方向",
    }


_LIB_CACHE = {"data": None, "mtime": None, "path": None}


def _load_from_disk():
    """从用户目录或内置路径读取库。"""
    from paths import user_data_path
    user_topics = user_data_path("topics.json")
    if user_topics.exists():
        try:
            return user_topics, json.loads(user_topics.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            pass  # 用户版损坏则回退到内置
    return TOPICS_PATH, json.loads(TOPICS_PATH.read_text(encoding="utf-8"))


def load_library():
    """优先读用户目录 ~/Documents/毕业论文助手/topics.json，回退到打包内置版。

    带文件 mtime 缓存：文件未变时直接返回内存副本，避免重复 IO+JSON parse。
    """
    from paths import user_data_path
    user_topics = user_data_path("topics.json")
    path = user_topics if user_topics.exists() else TOPICS_PATH
    try:
        mtime = path.stat().st_mtime
    except OSError:
        mtime = -1
    if _LIB_CACHE["data"] is not None and _LIB_CACHE["mtime"] == mtime and _LIB_CACHE["path"] == path:
        return _LIB_CACHE["data"]
    actual_path, data = _load_from_disk()
    _LIB_CACHE.update({"data": data, "mtime": mtime, "path": actual_path})
    return data


def invalidate_library_cache():
    """导入新数据后调用，让下次 load_library 重新读盘。"""
    _LIB_CACHE.update({"data": None, "mtime": None, "path": None})


def get_library_insights() -> dict:
    """从库里提取有用的统计信息，作为分析依据"""
    lib = load_library()
    topics = lib["topics"]

    # 关键词热度
    kw_counter = Counter()
    for t in topics:
        for kw in t.get("keywords", []):
            kw_counter[kw.strip()] += 1
    top_keywords = kw_counter.most_common(30)

    # 理论热度 + 方向归属
    theory_counter = Counter()
    theory_dir_map = {
        # 翻译方向（含较冷门的可作为推荐）
        "功能对等": "翻译", "目的论": "翻译", "生态翻译学": "翻译",
        "关联理论": "翻译",
        "译者主体性": "翻译", "翻译适应选择": "翻译", "变译理论": "翻译",
        "翻译补偿": "翻译", "文本世界": "翻译", "图式理论": "翻译",
        "改写理论": "翻译", "描写翻译学": "翻译",
        # 跨文化方向
        "跨文化交际": "跨文化", "文化维度": "跨文化", "霍夫斯泰德": "跨文化",
        "高低语境": "跨文化", "文化图式": "跨文化",
        "文化适应": "跨文化", "文化身份": "跨文化", "文化价值": "跨文化",
        "面子理论": "跨文化", "身份建构": "跨文化",
        # 话语分析方向
        "多模态话语分析": "话语分析", "视觉语法": "话语分析",
        "批评话语分析": "话语分析", "元话语": "话语分析",
        "互文性": "话语分析", "接受美学": "话语分析", "语用学": "话语分析",
        "评价理论": "话语分析", "态度系统": "话语分析", "及物性": "话语分析",
        "概念隐喻": "话语分析", "概念整合": "话语分析", "叙事学": "话语分析",
        "拟剧论": "话语分析", "语码转换": "话语分析", "符号学": "话语分析",
        # 商务英语习得方向
        "TPACK": "商务英语习得", "认知负荷": "商务英语习得",
        "情感过滤": "商务英语习得", "语境教学法": "商务英语习得",
        "建构主义": "商务英语习得", "产出导向法": "商务英语习得",
        "输入假设": "商务英语习得", "词块理论": "商务英语习得",
        # 通用（跨多个方向）
        "顺应论": "通用", "礼貌原则": "通用", "言语行为": "通用",
    }
    for t in topics:
        for kw in t.get("keywords", []):
            for th, th_dir in theory_dir_map.items():
                if th in kw:
                    theory_counter[th] += 1
                    break

    # 题目长度
    lengths = [len(t.get("title", "").replace(" ", "").replace("\n", ""))
               for t in topics]
    lengths = [l for l in lengths if 5 < l < 80]

    # 理论列表带上方向标签
    top_theories_with_dir = [
        (th, c, theory_dir_map.get(th, "通用"))
        for th, c in theory_counter.most_common()
    ]

    return {
        "total": len(topics),
        "top_keywords": top_keywords,
        "top_theories": top_theories_with_dir,
        "title_length_stats": {
            "median": sorted(lengths)[len(lengths)//2],
            "mean": round(sum(lengths)/len(lengths), 1),
            "min": min(lengths),
            "max": max(lengths),
        },
    }


def _check_length_strict(title: str) -> dict:
    """按 2025-12 修订规范检查主标题字数。
    中文：去掉所有空白后，约 20 字左右（常见 15-25）。
    英文：按空格分词后，约 15 词左右（常见 10-20）。
    返回：{kind: 'zh'|'en'|'mixed', zh_chars, en_words, assessment, details}
    """
    # 拆分中英文
    en_parts = re.findall(r"[A-Za-z][A-Za-z\-]*", title)
    en_text = " ".join(en_parts).strip()
    zh_text = re.sub(r"[A-Za-z0-9\s\-\.\(\)（）：:、，。《》!?]", "", title).strip()

    zh_chars = len(zh_text.replace(" ", ""))
    en_words = len([w for w in en_text.split() if w])

    # 类型判定：以中文字符为主 → zh；以英文为主 → en；混合 → mixed
    if zh_chars >= 5 and en_words <= 2:
        kind = "zh"
    elif en_words >= 5 and zh_chars <= 2:
        kind = "en"
    else:
        kind = "mixed"

    target_zh = 20
    target_en = 15

    if kind == "zh":
        if 15 <= zh_chars <= 25:
            assessment = "字数合适"
            details = f"主标题中文 {zh_chars} 字（规范约 20 字，常见 15-25）"
        elif zh_chars < 15:
            assessment = "偏短"
            details = f"主标题中文 {zh_chars} 字（规范约 20 字，常见 15-25，建议加具体对象或子方向）"
        else:
            assessment = "偏长"
            details = f"主标题中文 {zh_chars} 字（规范约 20 字，常见 15-25，建议删除冗余修饰或拆分主副标题）"
    elif kind == "en":
        if 10 <= en_words <= 20:
            assessment = "字数合适"
            details = f"主标题英文 {en_words} 词（规范约 15 词，常见 10-20）"
        elif en_words < 10:
            assessment = "偏短"
            details = f"主标题英文 {en_words} 词（规范约 15 词，常见 10-20）"
        else:
            assessment = "偏长"
            details = f"主标题英文 {en_words} 词（规范约 15 词，常见 10-20）"
    else:
        # 中英混合：分别评估
        zh_ok = 15 <= zh_chars <= 25
        en_ok = 10 <= en_words <= 20
        if zh_ok and en_ok:
            assessment = "字数合适"
            details = f"中文 {zh_chars} 字 / 英文 {en_words} 词，均在合理范围"
        else:
            assessment = "字数需调整"
            issues = []
            if not zh_ok:
                issues.append(f"中文 {zh_chars} 字（建议约 20）")
            if not en_ok:
                issues.append(f"英文 {en_words} 词（建议约 15）")
            details = "；".join(issues)

    return {
        "kind": kind,
        "zh_chars": zh_chars,
        "en_words": en_words,
        "assessment": assessment,
        "details": details,
    }


def _check_narrowness(title: str, theory_info: dict, object_info: dict) -> dict:
    """检查选题是否足够窄（有具体理论 + 具体研究对象 + 具体研究主题/问题）。
    三要素齐全 = 选题聚焦；缺一项都太宽。
    """
    has_theory = theory_info["has_theory"]
    has_object = object_info["has_object"]
    # 主题/问题关键词：翻译策略、特征、机制、影响、作用、建构、实践、应用
    theme_pat = re.compile(
        r"(翻译策略|翻译方法|话语策略|话语建构|话语特征|语言特征|建构研究|"
        r"应用研究|影响研究|作用研究|机制研究|实践研究|教学策略|教学模式|"
        r"传播策略|营销策略|本土化策略|跨文化策略|适应性研究|创新路径|"
        r"教学效果|学习效果|习得效果|接受效果|传播效果)",
        re.I,
    )
    has_theme = bool(theme_pat.search(title))

    missing = []
    if not has_theory:
        missing.append("具体理论框架")
    if not has_object:
        missing.append("具体研究对象")
    if not has_theme:
        missing.append("具体研究主题/问题")

    if not missing:
        level = "聚焦"
        assessment = "✅ 选题三要素齐全（理论 + 对象 + 主题），聚焦性好"
    elif len(missing) == 1:
        level = "较窄"
        assessment = f"🟡 选题基本聚焦，缺少：{missing[0]}"
    elif len(missing) == 2:
        level = "偏宽"
        assessment = f"🟠 选题偏宽，缺少：{'、'.join(missing)}"
    else:
        level = "过宽"
        assessment = "🔴 选题过宽：理论、对象、主题三者都缺，建议大幅收窄"

    return {
        "has_theory": has_theory,
        "has_object": has_object,
        "has_theme": has_theme,
        "missing": missing,
        "level": level,
        "assessment": assessment,
    }


def _build_reference_topics(
    lib, user_kw_set, direction, narrowness, year_window=("2023", "2024", "2025", "2026"),
    n=5,
):
    """推荐符合所有要求的近三年参考题目。
    要求：① 商务英语 4 大范围；② 选题三要素齐全（理论+对象+主题）；
         ③ 与用户题目关键词不撞；④ 优先近三年。
    """
    candidates = []
    for t in lib["topics"]:
        # 年份限制（近三年/四年）
        if t.get("year") not in year_window:
            continue
        title = t.get("title", "")
        if not title or title.startswith("论文题目"):  # 跳过模板条目
            continue
        # 范围判定（用 topic_analyzer 的 SCOPE_CATEGORIES 关键词）
        scope = check_scope(title)
        if not scope["in_scope"]:
            continue
        # 三要素齐全
        ti = detect_theory_expression(title)
        oi = detect_object_expression(title)
        nz = _check_narrowness(title, ti, oi)
        if nz["level"] != "聚焦":
            continue
        # 不撞关键词
        t_kws = set()
        for kw in t.get("keywords", []):
            for k in re.split(r"[\s,;]+", kw):
                k = k.strip()
                if k and k.lower() not in STOPWORDS and len(k) >= 2:
                    t_kws.add(k.lower())
        overlap = user_kw_set & t_kws
        if len(overlap) >= 2:
            continue
        # 排序分：年份越近越好
        yr_score = {"2026": 4, "2025": 3, "2024": 2, "2023": 1}.get(t.get("year", ""), 0)
        candidates.append((yr_score, len(overlap), t))

    candidates.sort(key=lambda x: (-x[0], x[1]))
    return [t for _, _, t in candidates[:n]]


def analyze_topic(user_title: str) -> dict:
    """对题目进行全面分析（v0.5：按 5 段式建议顺序重写）。"""
    lib = load_library()
    insights = get_library_insights()

    keywords = extract_keywords(user_title)
    theory_info = detect_theory_expression(user_title)
    object_info = detect_object_expression(user_title)
    direction = detect_direction(user_title)
    scope_result = check_scope(user_title)
    length_info = _check_length_strict(user_title)
    narrowness = _check_narrowness(user_title, theory_info, object_info)

    # === 撞题检测 ===
    matches = []
    user_kw_set = {k.lower() for k in keywords}
    for t in lib["topics"]:
        t_kws = set()
        for kw in t.get("keywords", []):
            for k in re.split(r"[\s,;]+", kw):
                k = k.strip()
                if k and k.lower() not in STOPWORDS and len(k) >= 2:
                    t_kws.add(k.lower())
        for tk in extract_keywords(t["title"]):
            t_kws.add(tk.lower())
        overlap = user_kw_set & t_kws
        if len(overlap) >= 2:
            matches.append({
                "id": t["id"],
                "title": t["title"],
                "overlap": list(overlap),
                "overlap_count": len(overlap),
                "direction": t.get("direction", "其他"),
                "year": t.get("year", "?"),
            })
    matches.sort(key=lambda x: x["overlap_count"], reverse=True)

    # === 创新性评估 ===
    if matches:
        top = matches[0]
        if top["overlap_count"] >= 3:
            innov_level = "🔴 高度撞题"
            innov_detail = (
                f"与 {top['year']} 届「{top['title'][:30]}...」"
                f"共享 {top['overlap_count']} 个关键词，方向/对象/主题可能重复"
            )
        elif top["overlap_count"] == 2:
            innov_level = "🟡 可能撞题"
            innov_detail = (
                f"与 {top['year']} 届「{top['title'][:30]}...」"
                f"共享 {top['overlap_count']} 个关键词，需明确差异化"
            )
        else:
            innov_level = "🟢 创新性可接受"
            innov_detail = f"最高撞题 {top['overlap_count']} 个关键词，差异化可行"
    else:
        innov_level = "🟢 创新性可接受"
        innov_detail = "未检测到与往届高重叠的选题"

    # === 风险评分（按重要性加权）===
    risk = 0
    risks = []

    # 0. 方向归到"其他"——与商务英语 4 大方向都不相关，不建议
    if direction == "其他":
        risk += 35
        risks.append("🚫 选题方向与商务英语 4 大方向（翻译/跨文化/话语分析/商务英语习得）均不相关，不建议")

    # 1. 专业范围不符（最致命）：40
    if not scope_result["in_scope"]:
        risk += 40
        risks.append("题目不在 4 大专业范围内")
    elif scope_result["matched_category"] == "商务英语学习与教学研究":
        # 教学类偏离商务环境，提示但不过度扣分
        risk += 8
        risks.append("选题偏向教学方向，需确认是否符合本专业的'商务环境'要求")

    # 2. 选题过宽：30
    if narrowness["level"] == "过宽":
        risk += 30
        risks.append("选题过宽（缺理论/对象/主题三要素）")
    elif narrowness["level"] == "偏宽":
        risk += 18
        risks.append(f"选题偏宽：{narrowness['assessment']}")
    elif narrowness["level"] == "较窄":
        risk += 6
        risks.append(f"选题基本聚焦：{narrowness['assessment']}")

    # 3. 撞题：20-40
    if matches:
        top = matches[0]
        if top["overlap_count"] >= 4:
            risk += 35
            risks.append(f"与往届高度撞题（共享 {top['overlap_count']} 个关键词）")
        elif top["overlap_count"] >= 3:
            risk += 22
            risks.append(f"与往届中度撞题（共享 {top['overlap_count']} 个关键词）")
        elif top["overlap_count"] == 2:
            risk += 10
            risks.append(f"与往届轻度撞题（共享 2 个关键词）")

    # 4. 字数不符：8
    if length_info["assessment"] in ("偏短", "偏长", "字数需调整"):
        risk += 8

    # 5. 表达不规范：5
    if theory_info["has_theory"] and theory_info["expression_patterns"]:
        good = {"基于X理论/学", "X理论下", "X视角", "X视域", "从X角度/视角"}
        if not any(p in good for p in theory_info["expression_patterns"]):
            risk += 5
            risks.append("理论表达方式不太规范")

    risk = min(risk, 100)

    # === 5 段式建议（按用户指定顺序）===
    suggestions = []

    # 顶部醒目提示：方向不属于 4 大类时
    if direction == "其他":
        suggestions.append("🚫🚫🚫 不建议使用此选题 🚫🚫🚫")
        suggestions.append("   本题与商务英语 4 大方向（翻译/跨文化/话语分析/商务英语习得）均不相关，")
        suggestions.append("   属于纯文学/跨学科研究（如女性主义、原型批评、生态批评、弗洛伊德心理学等）。")
        suggestions.append("   建议改用 4 大方向内的题目，或加上商务英语研究视角重新构建。")
        suggestions.append("")

    # === 段 1：是否符合本专业研究范围（强调商务环境）===
    suggestions.append("【1/5】专业研究范围（商务环境）")
    if scope_result["in_scope"]:
        cat = scope_result["matched_category"]
        suggestions.append(f"   ✅ 符合本专业「{cat}」方向")
        suggestions.append(f"   说明：{scope_result.get('description', '')}")
        suggestions.append(f"   ⚠️ 商务英语专业的核心要求：选题必须能体现『商务环境』（商务话语/商务文本/跨文化商务/商务教学）")
    else:
        suggestions.append("   ❌ 不在 4 大专业范围内")
        suggestions.append("   本专业 4 大方向：")
        for cat, info in SCOPE_CATEGORIES.items():
            suggestions.append(f"   • {cat}：{info['description']}")
        suggestions.append("   ⚠️ 必须体现『商务环境』要求")
    suggestions.append("")

    # === 段 2：选题是否足够窄 ===
    suggestions.append("【2/5】选题聚焦性（三要素检查）")
    suggestions.append(f"   {narrowness['assessment']}")
    suggestions.append(f"   • 具体理论框架：{'✅ ' + (theory_info['matched_theories'][0] if theory_info['matched_theories'] else '已识别') if narrowness['has_theory'] else '❌ 未识别'}")
    suggestions.append(f"   • 具体研究对象：{'✅ ' + (object_info['object_name'][:30] if object_info['object_name'] else '已识别') if narrowness['has_object'] else '❌ 未识别'}")
    suggestions.append(f"   • 具体研究主题/问题：{'✅ 已明确' if narrowness['has_theme'] else '❌ 未明确'}")
    if narrowness["missing"]:
        suggestions.append(f"   💡 需补：{'、'.join(narrowness['missing'])}")
    suggestions.append("")

    # === 段 3：选题是否太大众化或创新性不足 ===
    suggestions.append("【3/5】创新性 / 撞题检测")
    suggestions.append(f"   {innov_level}")
    suggestions.append(f"   {innov_detail}")
    if matches:
        suggestions.append("   撞题 Top 3：")
        for m in matches[:3]:
            suggestions.append(f"   • [{m['id']}] {m['year']}届·{m['direction']}：{m['title'][:50]}...")
            suggestions.append(f"     共享 {m['overlap_count']} 个关键词：{', '.join(list(m['overlap'])[:4])}")
    suggestions.append("")

    # === 段 4：选题是否符合字数要求和表达要求 ===
    suggestions.append("【4/5】字数 & 表达要求")
    suggestions.append(f"   📏 字数：{length_info['details']}")
    if length_info["assessment"] in ("偏短", "偏长", "字数需调整"):
        if length_info["kind"] == "zh":
            suggestions.append(f"      建议将主标题控制在 {15}-{25} 字（核心 20 字）")
        elif length_info["kind"] == "en":
            suggestions.append(f"      建议将主标题控制在 {10}-{20} 词（核心 15 词）")
    # 表达要求
    if theory_info["has_theory"]:
        if theory_info["expression_patterns"]:
            suggestions.append(f"   📝 表达：{', '.join(theory_info['expression_patterns'])}（已规范）")
        else:
            suggestions.append("   📝 表达：理论已识别但表达方式不规范")
            suggestions.append("      建议模板：「XX视角下」「基于XX理论」「XX理论框架下」")
    else:
        suggestions.append("   📝 表达：未识别到理论框架")
        suggestions.append("      建议加上高频模板：「XX视角下」「基于XX理论」")
    suggestions.append("")

    # === 段 5：具体改进建议 + 参考题目 ===
    suggestions.append("【5/5】具体改进建议 + 近三年参考题目")
    # 改进建议
    if not scope_result["in_scope"]:
        suggestions.append("   ① 改方向：参考下面 4 大方向的示例")
    if narrowness["level"] in ("过宽", "偏宽"):
        suggestions.append("   ② 加三要素：参考段 2/5 缺失项")
        # 把理论分两组：高频 vs 冷门
        # 高频 = 前 5 名（用 ≥ 24 次）：谨慎使用
        # 冷门 = 6 名之后但仍有 3+ 次使用：推荐（更创新）
        high_freq = insights["top_theories"][:5]
        all_theories = insights["top_theories"][5:]
        # 按用户当前方向过滤相关理论（无方向时显示全部冷门）
        if direction == "其他":
            relevant_cold = all_theories
        else:
            relevant_cold = [t for t in all_theories if t[2] in (direction, "通用")]
            if len(relevant_cold) < 3:
                # 相关不足时补全
                relevant_cold = all_theories[:5]

        suggestions.append("      ⚠️  谨慎使用（往届 ≥ 24 次，撞题风险高）：")
        for th, c, _ in high_freq:
            suggestions.append(f"      • {th}（往届用过 {c} 次）")
        suggestions.append("")
        suggestions.append(f"      ✨ 推荐使用（冷门且与「{direction}」方向相关，更具创新性）：")
        for th, c, d in relevant_cold[:5]:
            tag = f" [{d}]" if d != "通用" else ""
            suggestions.append(f"      • {th}（往届用过 {c} 次）{tag}")
    if matches and matches[0]["overlap_count"] >= 3:
        top_match = matches[0]
        suggestions.append(f"   ③ 去撞题：与 {top_match['year']} 届「{top_match['title'][:30]}」撞题严重")
        suggestions.append("      三种改法：换案例 / 换理论 / 缩子方向")
    if length_info["assessment"] in ("偏短", "偏长", "字数需调整"):
        suggestions.append("   ④ 调字数：参考段 4/5 建议范围")

    # 推荐题目：符合所有要求 + 近三年
    reference_topics = _build_reference_topics(
        lib, user_kw_set, direction, narrowness,
        year_window=("2023", "2024", "2025", "2026"), n=5,
    )

    if reference_topics:
        suggestions.append("")
        suggestions.append("   📚 参考题目（近三年+符合所有要求+不撞关键词）：")
        for t in reference_topics:
            suggestions.append(f"   • [{t['id']}] {t.get('year', '?')}届·{t.get('direction', '其他')}")
            suggestions.append(f"     {t['title'][:70]}")
            if t.get("keywords"):
                suggestions.append(f"     关键词：{', '.join(t['keywords'][:4])}")
    else:
        suggestions.append("")
        suggestions.append("   📚 未找到近三年完全匹配的参考题；可放宽年份/降低关键词不撞要求")

    return {
        "input_title": user_title,
        "extracted_keywords": keywords,
        "theory_info": theory_info,
        "object_info": object_info,
        "narrowness": narrowness,
        "suggested_direction": direction,
        "scope_result": scope_result,
        "length_info": length_info,
        "innovation": {
            "level": innov_level,
            "detail": innov_detail,
        },
        "direction_heat": {
            "direction": direction,
            "count_in_library": sum(1 for t in lib["topics"] if t.get("direction") == direction),
            "percentage": f"{sum(1 for t in lib['topics'] if t.get('direction') == direction) / max(1, len(lib['topics'])) * 100:.1f}%",
            "assessment": "热门" if sum(1 for t in lib["topics"] if t.get("direction") == direction) > len(lib["topics"]) * 0.3 else "中等" if sum(1 for t in lib["topics"] if t.get("direction") == direction) > len(lib["topics"]) * 0.1 else "冷门",
        },
        "risk_score": risk,
        "risk_level": "🔴 高" if risk >= 50 else "🟡 中" if risk >= 25 else "🟢 低",
        "risks": risks,
        "matches_in_library": matches[:5],
        "suggestions": suggestions,
        "reference_topics": reference_topics,
        "library_insights": insights,
    }


if __name__ == "__main__":
    test_titles = [
        "目的论视角下SHEIN跨境电商直播话语策略研究",
        "抖音直播带货话语分析",
        "汉诗英译策略研究",
        "商务英语专业学生实习现状研究",
        "中俄商务谈判中的跨文化交际策略研究",
        "TPACK理论框架下商务英语口语教学研究",
    ]
    for t in test_titles:
        print("\n" + "="*70)
        r = analyze_topic(t)
        print(f"题目: {t}")
        print(f"范围: {r['scope_result']['icon']} {r['scope_result']['assessment']} → {r['scope_result']['matched_category']}")
        print(f"风险: {r['risk_level']} ({r['risk_score']})")
        print(f"理论: {r['theory_info']}")
        print(f"对象: {r['object_info']}")
        print(f"方向: {r['suggested_direction']} ({r['direction_heat']['assessment']})")
        if r["risks"]:
            print(f"风险点: {r['risks'][0]}")
        print(f"建议数: {len(r['suggestions'])}, 参考题数: {len(r['reference_topics'])}")
