"""
商务英语毕业论文助手 - 桌面 GUI 主程序
"""
import io
import json
import sys

# Windows 默认 GBK 控制台，强制 UTF-8 避免中文 print/异常信息报错
if sys.platform == "win32":
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
        sys.stderr = sys.stdout
    except Exception:
        pass

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import Optional

from format_checker import check_document, mark_doc_with_issues, SEVERITY_COLOR
from paths import atomic_write, user_data_path
from topic_analyzer import analyze_topic, invalidate_library_cache, load_library, get_library_insights


# 跨平台中文字体检测（模块级缓存，避免反复探测）
_LINUX_CN_FONT: Optional[str] = None

# DPI 缩放因子：main() 探测后写入，_cn_font() 据此放大字号
# 1.0 = 96 DPI（普通屏）；1.25/1.5/2.0 = 高 DPI
# 计算公式：px_per_cm / 37.8（96 DPI = 37.8 px/cm）
_DPI_SCALE: float = 1.0


def _detect_linux_cn_font() -> Optional[str]:
    """探测当前 Linux 桌面可用的中文字体（fc-list 优先，回退 tk 探测）。

    结果模块级缓存，避免每次创建新 Tk root。
    """
    global _LINUX_CN_FONT
    if _LINUX_CN_FONT is not None:
        return _LINUX_CN_FONT if _LINUX_CN_FONT else None
    candidates = (
        "Noto Sans CJK SC", "Source Han Sans SC", "WenQuanYi Micro Hei",
        "WenQuanYi Zen Hei", "Droid Sans Fallback", "Sarasa Gothic SC",
        "Microsoft YaHei",  # WSL 下也可能用
    )
    # 优先 fc-list：避免新建 Tk 实例
    try:
        import subprocess
        r = subprocess.run(
            ["fc-list", ":lang=zh"],
            capture_output=True, text=True, timeout=2,
        )
        if r.returncode == 0:
            available = r.stdout
            for c in candidates:
                if c in available:
                    _LINUX_CN_FONT = c
                    return c
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
        pass
    # 回退 tk 探测（不创建 root，用 font families）
    try:
        import tkinter.font as tkfont
        families = set(tkfont.families())
        for c in candidates:
            if c in families:
                _LINUX_CN_FONT = c
                return c
    except Exception:
        pass
    _LINUX_CN_FONT = ""  # 已探测但无结果，避免重复探测
    return None


def _cn_font(size: int = 10, bold: bool = False) -> tuple:
    """根据操作系统返回可用的中文字体（字号已按 DPI 缩放）。

    scale_factor 由 main() 在创建 Tk root 后探测并写入 _DPI_SCALE，
    使 _cn_font() 在调用时按比例缩放字号，跨高 DPI 屏不显小。
    """
    weights = ("bold",) if bold else ()
    scaled = max(8, int(round(size * _DPI_SCALE)))
    if sys.platform == "darwin":
        return ("PingFang SC", scaled) + weights
    elif sys.platform == "win32":
        # Windows 优先用微软雅黑，找不到退回系统默认中文字体
        return ("Microsoft YaHei", scaled) + weights
    else:  # Linux / 其他
        f = _detect_linux_cn_font()
        if f:
            return (f, scaled) + weights
        return ("TkDefaultFont", scaled) + weights


def _mono_font(size: int = 10) -> tuple:
    """跨平台等宽字体（用于检查结果展示）。"""
    scaled = max(8, int(round(size * _DPI_SCALE)))
    if sys.platform == "darwin":
        return ("Menlo", scaled)
    elif sys.platform == "win32":
        return ("Consolas", scaled)
    else:
        return ("DejaVu Sans Mono", scaled)

# macOS 上 tk.Button 的 fg/bg 在 Aqua 主题下可能被忽略，
# 这里用一个 dict 来统一管理按钮配色，保证白字看得清
BUTTON_STYLES = {
    "primary":   {"bg": "#1976d2", "fg": "#ffffff", "activebackground": "#1565c0", "activeforeground": "#ffffff"},
    "success":   {"bg": "#2e7d32", "fg": "#ffffff", "activebackground": "#1b5e20", "activeforeground": "#ffffff"},
    "warning":   {"bg": "#ed6c02", "fg": "#ffffff", "activebackground": "#c55300", "activeforeground": "#ffffff"},
    "secondary": {"bg": "#455a64", "fg": "#ffffff", "activebackground": "#37474f", "activeforeground": "#ffffff"},
    "ghost":     {"bg": "#e0e0e0", "fg": "#212121", "activebackground": "#bdbdbd", "activeforeground": "#212121"},
    "example":   {"bg": "#f5f5f5", "fg": "#1565c0", "activebackground": "#e3f2fd", "activeforeground": "#0d47a1"},
}


def make_button(parent, text, command, style="primary", **kw):
    """用 tk.Label 模拟按钮（彻底解决 macOS Aqua 主题下 fg 失效的问题）
    支持的额外 kw: padx, pady, font, width
    """
    s = BUTTON_STYLES.get(style, BUTTON_STYLES["primary"])
    padx = kw.pop("padx", 14)
    pady = kw.pop("pady", 6)
    font = kw.pop("font", _cn_font(10, bold=True))
    width = kw.pop("width", None)

    # 容器 frame 用背景色填充（避免 Label 透明）
    btn_frame = tk.Frame(parent, bg=s["bg"], cursor="hand2", bd=0, highlightthickness=0)

    # 内部 Label 显示文字
    lbl = tk.Label(
        btn_frame,
        text=text,
        bg=s["bg"],
        fg=s["fg"],
        activebackground=s["activebackground"],
        activeforeground=s["activeforeground"],
        font=font,
        padx=padx,
        pady=pady,
        cursor="hand2",
        bd=0,
        highlightthickness=0,
    )
    lbl.pack()

    def on_click(e=None):
        if command:
            command()

    def on_enter(e):
        btn_frame.config(bg=s["activebackground"])
        lbl.config(bg=s["activebackground"])

    def on_leave(e):
        btn_frame.config(bg=s["bg"])
        lbl.config(bg=s["bg"])

    for widget in (btn_frame, lbl):
        widget.bind("<Button-1>", on_click)
        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)

    return btn_frame


class ThesisAssistantApp:
    def __init__(self, root):
        self.root = root
        root.title("商务英语毕业论文助手 v0.2 — 仲恺农业工程学院外国语学院")
        # 窗口尺寸 / 最小尺寸在 main() 里按物理 cm 设置（跨 DPI 一致）
        # 不再在 __init__ 硬编码像素值，避免被高分屏压缩成小窗

        self.status_var = tk.StringVar(value="就绪")

        # 顶部信息栏
        header = ttk.Frame(root)
        header.pack(fill="x")
        tk.Label(
            header, text="📚 商务英语毕业论文助手",
            fg="#2c3e50", font=_cn_font(16, bold=True),
            padx=20, pady=12,
        ).pack(side="left")
        tk.Label(
            header, text="仲恺农业工程学院外国语学院 · 本地版 · 零网络",
            fg="#7f8c8d", font=_cn_font(10),
        ).pack(side="left", padx=10)

        # 顶部右侧按钮：导入选题
        make_button(header, "📥 导入选题", self._on_import_topics,
                    style="primary", padx=14, pady=8,
                    font=_cn_font(10, bold=True)).pack(side="right", padx=10)

        # 主体：Notebook 三个标签页
        notebook = ttk.Notebook(root)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab_topic = ttk.Frame(notebook)
        self.tab_format = ttk.Frame(notebook)
        self.tab_library = ttk.Frame(notebook)

        notebook.add(self.tab_topic, text="🎯  选题建议")
        notebook.add(self.tab_library, text="📚  选题库浏览")
        notebook.add(self.tab_format, text="📋  格式检查")

        self._build_topic_tab()
        self._build_format_tab()
        self._build_library_tab()

        # 默认显示「选题建议」tab
        notebook.select(self.tab_topic)

        # 底部状态栏
        status = ttk.Label(root, textvariable=self.status_var, anchor="w", font=_cn_font(9))
        status.pack(side="bottom", fill="x")

    # ---------------- Tab 1: 选题建议 ----------------
    def _build_topic_tab(self):
        f = self.tab_topic

        # 顶部说明
        tk.Label(
            f, text="输入你的候选题目，工具会分析：撞题风险 / 理论框架 / 方向热度 / 是否符合本专业范围",
            font=_cn_font(10), fg="#7f8c8d",
        ).pack(anchor="w", padx=20, pady=(15, 5))

        # 输入区
        input_frame = ttk.LabelFrame(f, text=" 你的题目 ")
        input_frame.pack(fill="x", padx=20, pady=10)

        self.topic_entry = tk.Text(input_frame, height=3, font=_cn_font(11),
                                    wrap="word", relief="flat", bg="#fafafa")
        self.topic_entry.pack(fill="x", padx=10, pady=10)
        self.topic_entry.insert("1.0", "示例：目的论视角下跨境电商产品描述翻译策略研究——以SHEIN为例")

        # 操作按钮
        btn_frame = ttk.Frame(f)
        btn_frame.pack(fill="x", padx=20)

        make_button(
            btn_frame, "🔍 分析选题", self._on_analyze_topic,
            style="primary", padx=20, pady=8,
            font=_cn_font(11, bold=True),
        ).pack(side="left")

        make_button(
            btn_frame, "清空",
            lambda: self.topic_entry.delete("1.0", "end"),
            style="ghost", padx=15, pady=8,
            font=_cn_font(10),
        ).pack(side="left", padx=10)

        # 快捷示例（只保留 2 个，按钮变窄 + 完整显示）
        tk.Label(btn_frame, text="快速试用：",
                 font=_cn_font(9), fg="#95a5a6").pack(side="left", padx=(20, 5))

        examples = [
            "目的论视角下SHEIN跨境电商翻译",
            "直播带货话语分析",
        ]
        for ex in examples:
            make_button(
                btn_frame, ex,
                lambda e=ex: self.topic_entry.delete("1.0", "end") or self.topic_entry.insert("1.0", e),
                style="example", padx=10, pady=4,
                font=_cn_font(9),
            ).pack(side="left", padx=2)

        # 结果区
        result_frame = ttk.LabelFrame(f, text=" 分析结果 ")
        result_frame.pack(fill="both", expand=True, padx=20, pady=10)

        self.topic_result = scrolledtext.ScrolledText(
            result_frame, font=_mono_font(10), wrap="word",
            bg="#fafafa", relief="flat", state="disabled",
        )
        self.topic_result.pack(fill="both", expand=True, padx=10, pady=10)

    def _on_analyze_topic(self):
        title = self.topic_entry.get("1.0", "end").strip()
        if not title or title.startswith("示例"):
            messagebox.showwarning("提示", "请先输入你的题目")
            return
        try:
            result = analyze_topic(title)
            output = self._format_topic_result(result)
            self.topic_result.config(state="normal")
            self.topic_result.delete("1.0", "end")
            self.topic_result.insert("1.0", output)
            self.topic_result.config(state="disabled")
            self.status_var.set(f"选题分析完成 — 风险 {result['risk_level']}")
        except Exception as e:
            messagebox.showerror("错误", f"分析失败: {e}")

    def _format_topic_result(self, r) -> str:
        """按 5 段式输出：范围 → 聚焦性 → 创新性 → 字数/表达 → 改进建议+参考"""
        lines = []
        lines.append(f"{'='*70}")
        lines.append(f"📝 题目: {r['input_title']}")
        lines.append(f"{'='*70}\n")

        # 选题库统计摘要（基于真实正选题）
        # 已隐藏，避免冗余
        # lib_info = r.get("library_insights", {})
        # if lib_info:
        #     lines.append(f"📚 选题库: {lib_info.get('total', 0)} 条真实正选题（2019-2026 届）")
        #     lines.append("")

        # 顶部摘要
        lines.append(f"⚠️  风险评分: {r['risk_level']}  ({r['risk_score']} / 100)")
        sr = r["scope_result"]
        nz = r["narrowness"]
        li = r["length_info"]
        inn = r["innovation"]
        lines.append(f"   范围: {sr['icon']} {sr['matched_category']}")
        lines.append(f"   聚焦: {nz['level']} ({nz['assessment']})")
        lines.append(f"   字数: {li['assessment']}（{li['details']}）")
        lines.append(f"   创新: {inn['level']}")
        # 识别方向（仅显示商务英语 4 大方向）
        suggested = r.get("suggested_direction", "其他")
        if suggested == "其他":
            lines.append(f"   方向: 🚫 不属于商务英语 4 大方向，不建议")
        else:
            lines.append(f"   方向: {suggested}")
        # 方向热度（已隐藏，避免冗余）
        lines.append("")

        # 5 段建议（直接展示 r["suggestions"]）
        for sug in r["suggestions"]:
            lines.append(sug)
        lines.append("")

        # 撞题检测（详细列表）
        if r["matches_in_library"]:
            lines.append("📚 撞题明细:")
            for m in r["matches_in_library"]:
                lines.append(f"   • [{m['id']}] {m['year']}届·{m['direction']}")
                lines.append(f"     {m['title'][:70]}")
                lines.append(f"     共享 {m['overlap_count']} 个关键词：{', '.join(list(m['overlap'])[:4])}")
            lines.append("")

        # 风险点（简短）
        if r["risks"]:
            lines.append("⚠️  主要风险:")
            for risk in r["risks"]:
                lines.append(f"   • {risk}")
            lines.append("")

        lines.append(f"{'='*70}")
        if r["risk_score"] >= 50:
            lines.append("🔴 总体建议: 这个题目有较大风险，建议大幅修改后再与导师确认。")
        elif r["risk_score"] >= 25:
            lines.append("🟡 总体建议: 题目方向基本可取，但还需要根据上面的建议做小幅调整，再与导师讨论。")
        else:
            lines.append("🟢 总体建议: 题目方向清晰，可以进入下一步开题报告。")
        return "\n".join(lines)

    # ---------------- Tab 2: 格式检查 ----------------
    def _build_format_tab(self):
        f = self.tab_format

        tk.Label(
            f, text="选择你的论文 Word 文件（.docx），工具会按仲恺规范逐项检查格式",
            font=_cn_font(10), fg="#7f8c8d",
        ).pack(anchor="w", padx=20, pady=(15, 5))

        file_frame = ttk.Frame(f)
        file_frame.pack(fill="x", padx=20, pady=10)

        self.format_file_var = tk.StringVar()
        tk.Entry(file_frame, textvariable=self.format_file_var,
                 font=_cn_font(10), relief="solid", bd=1).pack(
            side="left", fill="x", expand=True, padx=(0, 10))

        make_button(file_frame, "📁 选择文件", self._on_choose_file,
                    style="primary").pack(side="left")
        make_button(file_frame, "🔍 开始检查", self._on_check_format,
                    style="success").pack(side="left", padx=5)
        make_button(file_frame, "📝 导出标注版 Word", self._on_export_marked_docx,
                    style="warning").pack(side="left", padx=5)
        make_button(file_frame, "导出报告", self._on_export_report,
                    style="secondary").pack(side="left", padx=5)

        result_frame = ttk.LabelFrame(f, text=" 检查结果 ")
        result_frame.pack(fill="both", expand=True, padx=20, pady=10)

        self.format_result = scrolledtext.ScrolledText(
            result_frame, font=_mono_font(10), wrap="word",
            bg="#fafafa", relief="flat", state="disabled",
        )
        self.format_result.pack(fill="both", expand=True, padx=10, pady=10)

    def _on_choose_file(self):
        path = filedialog.askopenfilename(
            title="选择论文 Word 文件",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
        )
        if path:
            self.format_file_var.set(path)

    def _on_check_format(self):
        path = self.format_file_var.get().strip()
        if not path or not Path(path).exists():
            messagebox.showwarning("提示", "请先选择有效的 .docx 文件")
            return
        try:
            result = check_document(path)
            self._last_format_result = result
            output = self._format_format_result(result)
            self.format_result.config(state="normal")
            self.format_result.delete("1.0", "end")
            self.format_result.insert("1.0", output)
            self.format_result.config(state="disabled")
            self.status_var.set(
                f"格式检查完成 — {result['grade']} — "
                f"高{result['stats']['issues_by_severity'].get('高',0)} "
                f"中{result['stats']['issues_by_severity'].get('中',0)}"
            )
        except Exception as e:
            messagebox.showerror("错误", f"检查失败: {e}\n\n请确认文件是 .docx 格式且未被加密。")

    def _format_format_result(self, r) -> str:
        lines = []
        lines.append(f"{'='*70}")
        lines.append(f"📄 文件: {Path(r['file']).name}")
        lines.append(f"🏫 规范依据: {r['school']} · 毕业论文格式规范（2025-12修订）")
        lines.append(f"📊 评级: {r['grade']}")
        lines.append(f"{'='*70}\n")

        # 区域识别
        lines.append(f"📍 文档结构（自动识别）:")
        for z, (s_idx, e_idx) in r.get("zones", {}).items():
            lines.append(f"   • {z}: 第{s_idx+1}–{e_idx}段")
        lines.append("")

        s = r["stats"]
        lines.append(f"📈 文档统计:")
        lines.append(f"   总段落数: {s['total_paragraphs']}")
        lines.append(f"   正文段数: {s.get('body_paragraphs', 0)}")
        # 正文 zone 段落范围（v2 新增）
        if s.get("body_zone_range") and s["body_zone_range"][1] > s["body_zone_range"][0]:
            br_s, br_e = s["body_zone_range"]
            br_count = br_e - br_s + 1
            lines.append(f"   正文: 第{br_s}–{br_e}段 ({br_count} 段)")
        # 内文引用 / 图表 / 例证（v2 新增）
        if "inline_citations_total" in s:
            lines.append(
                f"   正文引用: {s['inline_citations_total']} 处 "
                f"(分布在 {s.get('inline_citation_paragraphs', 0)} 段)"
            )
        if "figures_count" in s or "tables_count" in s:
            lines.append(
                f"   图表: 图 {s.get('figures_count', 0)} 个 · "
                f"表 {s.get('tables_count', 0)} 个"
            )
        if s.get("examples_count") is not None and s["examples_count"] > 0:
            lines.append(f"   例证: {s['examples_count']} 处")
        if "approx_word_count_with_tables" in s:
            lines.append(f"   正文字符数: 段 {s['approx_word_count']:,} + 表 "
                         f"{s['approx_word_count_with_tables'] - s['approx_word_count']:,} = "
                         f"{s['approx_word_count_with_tables']:,}")
        else:
            lines.append(f"   正文字符数: {s['approx_word_count']:,}")
        lines.append(f"   参考文献数量: {s.get('reference_count', 0)} 篇")
        lines.append(f"   章节标题分布: {s['headings_by_level']}")
        # 章节标题明细（前 10）
        if s.get("headings"):
            lines.append(f"   章节标题明细 (前 10):")
            shown = 0
            for lv in sorted({h["level"] for h in s["headings"]}):
                for h in s["headings"]:
                    if h["level"] != lv:
                        continue
                    lines.append(f"      L{lv} 第{h['idx']+1}段: {h['text']}")
                    shown += 1
                    if shown >= 10:
                        break
                if shown >= 10:
                    break
        if "page_margin" in s:
            m = s["page_margin"]
            lines.append(f"   页面边距: 上{m['top']} 下{m['bottom']} 左{m['left']} 右{m['right']} cm")
        lines.append("")

        sev = s["issues_by_severity"]
        lines.append(f"🔍 问题汇总: 共 {s['total_issues']} 项")
        lines.append(f"   🔴 高: {sev.get('高', 0)}   🟡 中: {sev.get('中', 0)}   🟢 低: {sev.get('低', 0)}")
        lines.append("")
        if s.get("issue_per_para"):
            lines.append(f"💡 提示: 点击工具栏「📝 导出标注版 Word」可在原文档中用颜色标注每个问题位置。")
            lines.append("")

        if not r["issues"]:
            lines.append("✅ 未发现问题，格式完全符合规范！")
            return "\n".join(lines)

        for severity, icon in [("高", "🔴"), ("中", "🟡"), ("低", "🟢")]:
            items = [i for i in r["issues"] if i["severity"] == severity]
            if not items:
                continue
            lines.append(f"{'─'*70}")
            lines.append(f"{icon} {severity}严重度问题 ({len(items)} 项)")
            lines.append(f"{'─'*70}")
            for idx, iss in enumerate(items, 1):
                zone = iss.get("zone", "")
                lines.append(f"\n{idx}. 【{iss['rule']}】 {f'({zone})' if zone else ''}")
                lines.append(f"   📍 位置: {iss['location']}")
                lines.append(f"   ✅ 规范: {iss['expected']}")
                lines.append(f"   ❌ 实际: {iss['actual']}")
                lines.append(f"   💡 建议: {iss['suggestion']}")

        lines.append(f"\n{'='*70}")
        return "\n".join(lines)

    def _on_export_report(self):
        if not hasattr(self, "_last_format_result"):
            messagebox.showwarning("提示", "请先运行一次格式检查")
            return
        path = filedialog.asksaveasfilename(
            title="保存检查报告",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt")],
        )
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self._format_format_result(self._last_format_result))
            self.status_var.set(f"报告已保存: {path}")
            messagebox.showinfo("成功", f"报告已保存到:\n{path}")

    def _on_export_marked_docx(self):
        """导出在原文档中标注了问题位置的 Word 文件（高亮 + 批注）"""
        if not hasattr(self, "_last_format_result"):
            messagebox.showwarning("提示", "请先运行一次格式检查")
            return
        result = self._last_format_result
        if not result.get("issue_per_para"):
            messagebox.showinfo("无问题", "未发现问题，无需标注。")
            return
        path = filedialog.asksaveasfilename(
            title="导出标注版 Word",
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")],
            initialfile=f"{Path(result['file']).stem}_标注版.docx",
        )
        if not path:
            return
        try:
            mark_doc_with_issues(result["file"], path, result)
            self.status_var.set(f"标注版已保存: {Path(path).name}")
            messagebox.showinfo(
                "成功",
                f"标注版 Word 已保存到:\n{path}\n\n"
                f"用 Word/WPS 打开后会看到:\n"
                f"  • 红/黄/灰 背景色 = 高/中/低 严重度问题\n"
                f"  • 段尾红色文字 = 具体问题说明\n\n"
                f"按 Ctrl+P 可截图保留标注版本。",
            )
        except Exception as e:
            messagebox.showerror("导出失败", f"错误: {e}")

    # ---------------- Tab 3: 选题库浏览 ----------------
    def _build_library_tab(self):
        f = self.tab_library

        top = ttk.Frame(f)
        top.pack(fill="x", padx=20, pady=(15, 5))

        # 选题总数从库实时读取
        try:
            _lib = load_library()
            _total = len(_lib.get("topics", []))
        except Exception:
            _total = 0
        self.lib_count_label = tk.Label(
            top, text=f"浏览往届选题库（共 {_total} 条）",
            font=_cn_font(10), fg="#7f8c8d",
        )
        self.lib_count_label.pack(side="left")

        filter_frame = ttk.Frame(f)
        filter_frame.pack(fill="x", padx=20, pady=10)

        tk.Label(filter_frame, text="搜索:", font=_cn_font(10)).pack(side="left")
        self.search_var = tk.StringVar()
        # 搜索框 debounce（每个键击一次，太频繁需要节流）
        self.search_var.trace_add("write", lambda *_: self._schedule_refresh())
        tk.Entry(filter_frame, textvariable=self.search_var,
                 font=_cn_font(10), width=30, relief="solid", bd=1).pack(side="left", padx=5)

        tk.Label(filter_frame, text="方向:", font=_cn_font(10)).pack(side="left", padx=(20, 0))
        self.direction_var = tk.StringVar(value="全部")
        directions = ["全部", "翻译", "跨文化", "话语分析", "商务英语习得"]
        self.direction_combo = ttk.Combobox(filter_frame, textvariable=self.direction_var,
                     values=directions, state="readonly", width=12
                     )
        self.direction_combo.pack(side="left", padx=5)
        # ttk.Combobox + state="readonly" + trace_add 在某些场景不触发，
        # 用 <<ComboboxSelected>> 虚拟事件更可靠
        self.direction_combo.bind("<<ComboboxSelected>>", lambda e: self._refresh_library())

        tk.Label(filter_frame, text="年份:", font=_cn_font(10)).pack(side="left", padx=(20, 0))
        self.year_var = tk.StringVar(value="全部")
        years = ["全部", "2026", "2025", "2024", "2023", "2022", "2021", "2020", "2019"]
        self.year_combo = ttk.Combobox(filter_frame, textvariable=self.year_var,
                     values=years, state="readonly", width=8
                     )
        self.year_combo.pack(side="left", padx=5)
        self.year_combo.bind("<<ComboboxSelected>>", lambda e: self._refresh_library())

        # 导出按钮（单独一行，靠右，避免被搜索/筛选控件挤窄）
        export_row = ttk.Frame(f)
        export_row.pack(fill="x", padx=20, pady=(4, 0))

        make_button(export_row, "📊 导出 Excel",
                    self._on_export_excel,
                    style="success", padx=18, pady=6,
                    font=_cn_font(10, bold=True)).pack(side="right", padx=(6, 0))
        make_button(export_row, "📄 导出 CSV",
                    self._on_export_csv,
                    style="secondary", padx=18, pady=6,
                    font=_cn_font(10, bold=True)).pack(side="right", padx=6)

        paned = tk.PanedWindow(f, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=20, pady=10)

        list_frame = tk.Frame(paned)
        paned.add(list_frame, width=560)

        tree = ttk.Treeview(
            list_frame,
            columns=("id", "title", "direction"),
            show="headings", height=20,
        )
        tree.heading("id", text="编号")
        tree.heading("title", text="题目")
        tree.heading("direction", text="方向")
        tree.column("id", width=70, anchor="center")
        tree.column("title", width=400)
        tree.column("direction", width=80, anchor="center")

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        tree.bind("<<TreeviewSelect>>", self._on_topic_select)
        self.library_tree = tree

        detail_frame = ttk.LabelFrame(paned, text=" 详情 ")
        paned.add(detail_frame, width=320)

        self.library_detail = scrolledtext.ScrolledText(
            detail_frame, font=_cn_font(10), wrap="word",
            bg="#fafafa", relief="flat", state="disabled",
        )
        self.library_detail.pack(fill="both", expand=True, padx=10, pady=10)

        self._refresh_library()

    def _schedule_refresh(self):
        """debounce：搜索/筛选变更后 200ms 再刷新，避免每个键击都重绘 Treeview。"""
        if hasattr(self, "_refresh_after_id") and self._refresh_after_id:
            try:
                self.after_cancel(self._refresh_after_id)
            except Exception:
                pass
        self._refresh_after_id = self.after(200, self._refresh_library)

    def _refresh_library(self):
        try:
            lib = load_library()
            keyword = self.search_var.get().strip().lower()
            direction = self.direction_var.get()
            year = self.year_var.get()
        except Exception as e:
            import traceback
            print(f"[_refresh_library] 初始化失败: {e}")
            traceback.print_exc()
            self.status_var.set(f"刷新失败: {e}")
            return

        # 收集筛选结果
        matched = []
        for t in lib["topics"]:
            if direction != "全部" and t.get("direction") != direction:
                continue
            if year != "全部" and str(t.get("year", "")) != year:
                continue
            if keyword and keyword not in t["title"].lower():
                continue
            matched.append(t)
        self._current_topics = matched

        # 按届别倒序、标题字典序，2019 在末尾、2026 在最前
        matched.sort(key=lambda t: (-int(str(t.get('year', '0')) or 0), t['title']))

        # 高效更新：先 detach 所有 → 重新 insert → 避免逐个 delete
        for item in self.library_tree.get_children():
            self.library_tree.delete(item)
        for t in matched:
            self.library_tree.insert("", "end", values=(
                t["id"], t["title"][:55] + ("…" if len(t["title"]) > 55 else ""),
                t.get("direction", "其他"),
            ))

        self.status_var.set(
            f"选题库: {lib['metadata']['total']} 条 | "
            f"当前筛选 {len(matched)} 条"
        )

    def _get_export_rows(self) -> list[dict]:
        """获取当前筛选结果，统一字段顺序"""
        rows = []
        for t in self._current_topics:
            rows.append({
                "编号": t.get("id", ""),
                "届别": t.get("year", ""),
                "题目": t.get("title", ""),
                "方向": t.get("direction", ""),
                "关键词": "; ".join(t.get("keywords", [])),
                "语言": t.get("language", ""),
                "选题来源": t.get("origin", ""),
                "选题性质": t.get("nature", ""),
            })
        return rows

    def _on_export_excel(self):
        if not getattr(self, "_current_topics", None):
            messagebox.showwarning("提示", "当前没有可导出的选题")
            return
        path = filedialog.asksaveasfilename(
            title="导出为 Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not path:
            return
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Alignment, Font, PatternFill
            rows = self._get_export_rows()
            wb = Workbook()
            ws = wb.active
            ws.title = "选题列表"
            # 表头
            headers = list(rows[0].keys())
            ws.append(headers)
            # 表头样式
            header_fill = PatternFill("solid", fgColor="2c3e50")
            header_font = Font(bold=True, color="FFFFFF", size=11)
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            # 数据
            for r in rows:
                ws.append([r[h] for h in headers])
            # 列宽
            col_widths = {"编号": 10, "届别": 8, "题目": 60, "方向": 12,
                          "关键词": 35, "语言": 8, "选题来源": 16, "选题性质": 12}
            for i, h in enumerate(headers, 1):
                ws.column_dimensions[chr(64 + i)].width = col_widths.get(h, 15)
            # 题目列自动换行
            for row in ws.iter_rows(min_row=2):
                row[2].alignment = Alignment(wrap_text=True, vertical="top")
            wb.save(path)
            self.status_var.set(f"已导出 {len(rows)} 条到 {Path(path).name}")
            messagebox.showinfo("导出成功",
                                f"已导出 {len(rows)} 条选题到：\n{path}\n\n"
                                f"可在 Excel/WPS/Numbers 中打开。")
        except Exception as e:
            messagebox.showerror("导出失败", f"错误: {e}")

    def _on_export_csv(self):
        if not getattr(self, "_current_topics", None):
            messagebox.showwarning("提示", "当前没有可导出的选题")
            return
        path = filedialog.asksaveasfilename(
            title="导出为 CSV",
            defaultextension=".csv",
            filetypes=[("CSV 文件", "*.csv")],
        )
        if not path:
            return
        try:
            import csv
            rows = self._get_export_rows()
            with open(path, "w", encoding="utf-8-sig", newline="") as f:
                # utf-8-sig 让 Excel 直接打开不乱码
                writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
                writer.writeheader()
                writer.writerows(rows)
            self.status_var.set(f"已导出 {len(rows)} 条到 {Path(path).name}")
            messagebox.showinfo("导出成功",
                                f"已导出 {len(rows)} 条选题到：\n{path}\n\n"
                                f"可直接用 Excel 打开。")
        except Exception as e:
            messagebox.showerror("导出失败", f"错误: {e}")

    def _on_topic_select(self, event):
        sel = self.library_tree.selection()
        if not sel:
            return
        item = self.library_tree.item(sel[0])
        tid = item["values"][0]
        lib = load_library()
        for t in lib["topics"]:
            if t["id"] == tid:
                self.library_detail.config(state="normal")
                self.library_detail.delete("1.0", "end")
                details = [
                    f"编号: {t['id']}",
                    f"题目: {t['title']}",
                    f"届别: {t.get('year', '?')}",
                    f"方向: {t.get('direction', '其他')}",
                    f"关键词: {', '.join(t.get('keywords', [])) or '（无）'}",
                    f"语言: {t.get('language', '英语')}",
                    f"来源: {t.get('origin', '')}",
                    f"性质: {t.get('nature', '')}",
                ]
                self.library_detail.insert("1.0", "\n".join(details))
                self.library_detail.config(state="disabled")
                break

    # ---------------- 导入选题 ----------------
    def _on_import_topics(self):
        path = filedialog.askopenfilename(
            title="选择选题文件 (.xls / .xlsx)",
            filetypes=[("Excel 文件", "*.xls *.xlsx"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            from topic_importer import import_file, merge_with_existing
            new_topics = import_file(path)
            if not new_topics:
                messagebox.showwarning(
                    "未导入",
                    f"文件解析成功但没有提取到选题。\n\n可能是文件格式不在支持范围内。\n文件: {Path(path).name}",
                )
                return
            # 合并到现有库（先算合并结果但不写）
            lib = load_library()
            merged, added = merge_with_existing(lib["topics"], new_topics)

            # 用户确认对话框（避免误操作）
            confirm = messagebox.askyesno(
                "确认导入",
                f"即将合并到本地选题库：\n\n"
                f"  文件: {Path(path).name}\n"
                f"  解析到: {len(new_topics)} 条\n"
                f"  新增（去重后）: {added} 条\n"
                f"  合并后总数: {len(merged)} 条\n\n"
                f"确定要写回本地库吗？",
            )
            if not confirm:
                self.status_var.set("导入已取消")
                return

            # 写到用户数据目录（不在源码目录）+ 原子写入
            lib["topics"] = merged
            lib["metadata"]["total"] = len(merged)
            user_topics = user_data_path("topics.json")
            atomic_write(
                user_topics,
                json.dumps(lib, ensure_ascii=False, indent=2),
            )
            invalidate_library_cache()

            # 刷新浏览页
            self._refresh_library()
            self.status_var.set(f"导入成功: {Path(path).name} → 新增 {added} 条, 总计 {len(merged)} 条")

            messagebox.showinfo(
                "导入成功",
                f"从 {Path(path).name} 导入：\n"
                f"  解析到: {len(new_topics)} 条\n"
                f"  新增（去重后）: {added} 条\n"
                f"  库总数: {len(merged)} 条\n\n"
                f"数据已保存到：\n{user_topics}\n\n"
                f"提示：可在『📚 选题库浏览』中查看新数据。",
            )
        except Exception as e:
            messagebox.showerror("导入失败", f"错误: {e}\n\n文件: {path}")


def _setup_high_dpi():
    """在创建 Tk 根之前开启 HiDPI 感知，避免打包后界面变小。

    - Windows: 设置进程 DPI 感知为 system-aware，让 tk 跟随系统缩放
    - Linux:   Qt/GTK 会自动处理；tk 默认不支持，留给 OS
    """
    if sys.platform == "win32":
        try:
            import ctypes
            # 0=unaware, 1=system-aware, 2=per-monitor
            # 用 1 跟随系统缩放（最稳，避免多显示器不一致）
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except (AttributeError, OSError):
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass


def main():
    global _DPI_SCALE
    _setup_high_dpi()
    root = tk.Tk()

    # === 跨平台 DPI 探测 + 缩放设置 ===
    # winfo_fpixels("1c") 返回 1 cm 对应的物理像素：
    #   96 DPI ≈ 37.8 px/cm;   120 DPI ≈ 47.2;   144 DPI ≈ 56.7;   192 DPI (Retina) ≈ 75.6
    # 基准 96 DPI → 1.0；比例 = px_per_cm / 37.8。
    # 这样窗口/字号都按物理尺寸缩放，Windows 高 DPI 屏不再"挤在小窗口里"，
    # macOS Retina 屏字号不显小。
    try:
        root.update_idletasks()
        px_per_cm = root.winfo_fpixels("1c")
        # 钳制在 [1.0, 2.5]：避免 4K 屏字号爆炸
        _DPI_SCALE = max(1.0, min(2.5, px_per_cm / 37.8))
        # macOS 强制 ≥1.25（Retina 一律按高分屏处理，避免文字偏小）
        # 设 1.25 是折中：Retina 上 12.5px（接近普通屏 13）、普通屏不放大爆炸
        if sys.platform == "darwin":
            _DPI_SCALE = max(_DPI_SCALE, 1.25)
        root.tk.call("tk", "scaling", _DPI_SCALE)
    except Exception:
        _DPI_SCALE = 1.0

    # === 全局 ttk Style：放大 Notebook 标签 + Treeview 行高 ===
    style = ttk.Style()
    # Treeview 字号随 DPI 缩放；行高 = 字号 + padding（避免行间挤压）
    tv_font_size = max(9, int(round(10 * _DPI_SCALE)))
    style.configure("Treeview", font=_cn_font(10), rowheight=tv_font_size + int(round(8 * _DPI_SCALE)))
    style.configure("Treeview.Heading", font=_cn_font(10, bold=True))
    # Notebook tab 内边距（左右各 12px，上下各 6px），让标签更醒目
    pad_x = int(round(12 * _DPI_SCALE))
    pad_y = int(round(6 * _DPI_SCALE))
    style.configure("TNotebook.Tab", padding=(pad_x, pad_y), font=_cn_font(11, bold=True))
    # LabelFrame 标题字号
    style.configure("TLabelframe.Label", font=_cn_font(10, bold=True))

    ThesisAssistantApp(root)

    # === 窗口尺寸按物理 cm 计算（跨平台一致）===
    # 默认 26 cm 宽 × 19 cm 高，最小 22 cm × 16 cm
    try:
        root.update_idletasks()
        px_per_cm = root.winfo_fpixels("1c")
        w = int(round(26 * px_per_cm))
        h = int(round(19 * px_per_cm))
        min_w = int(round(22 * px_per_cm))
        min_h = int(round(16 * px_per_cm))
        root.geometry(f"{w}x{h}")
        root.minsize(min_w, min_h)
    except Exception:
        root.geometry("960x720")
        root.minsize(800, 600)

    root.update_idletasks()
    root.update()
    root.after(100, lambda: root.update_idletasks())
    root.mainloop()


if __name__ == "__main__":
    main()
