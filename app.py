"""
商务英语毕业论文助手 - 桌面 GUI 主程序
"""
import io
import json
import math
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


def _cn_font(size: int = 10, bold: bool = False, italic: bool = False) -> tuple:
    """sans 字体（标题栏 / 标签 / location pill / 9-11pt 元信息用）"""
    weights = ()
    if bold: weights = weights + ("bold",)
    if italic: weights = weights + ("italic",)
    scaled = max(8, int(round(size * _DPI_SCALE)))
    if sys.platform == "darwin":
        return ("PingFang SC", scaled) + weights
    elif sys.platform == "win32":
        return ("Microsoft YaHei", scaled) + weights
    else:
        f = _detect_linux_cn_font()
        if f:
            return (f, scaled) + weights
        return ("TkDefaultFont", scaled) + weights


def _serif_font(size: int = 12, bold: bool = False, italic: bool = False) -> tuple:
    """衬线字体（学术笔记 v2 · 与 mockup Noto Serif SC 对齐）。
    - macOS: Songti SC (宋体)
    - Windows: SimSun (宋体)
    - Linux: Noto Serif CJK SC
    """
    weights = ()
    if bold: weights = weights + ("bold",)
    if italic: weights = weights + ("italic",)
    scaled = max(8, int(round(size * _DPI_SCALE)))
    if sys.platform == "darwin":
        return ("Songti SC", scaled) + weights
    elif sys.platform == "win32":
        return ("SimSun", scaled) + weights
    else:
        # Linux: 尝试 Noto Serif CJK SC，找不到退化为 sans
        for name in ("Noto Serif CJK SC", "Source Han Serif SC", "AR PL UMing CN"):
            try:
                import tkinter.font as tkf
                fams = tkf.families()
                if name in fams:
                    return (name, scaled) + weights
            except Exception:
                pass
        return ("TkDefaultFont", scaled) + weights


def _latin_font(size: int = 10, bold: bool = False, italic: bool = False) -> tuple:
    """英文衬线字体（grade letter / rule 名 / 等数字英文 · Times New Roman 系）"""
    weights = ()
    if bold: weights = weights + ("bold",)
    if italic: weights = weights + ("italic",)
    scaled = max(8, int(round(size * _DPI_SCALE)))
    return ("Times New Roman", scaled) + weights


def _mono_font(size: int = 10) -> tuple:
    """跨平台等宽字体（用于检查结果展示）。"""
    scaled = max(8, int(round(size * _DPI_SCALE)))
    if sys.platform == "darwin":
        return ("Menlo", scaled)
    elif sys.platform == "win32":
        return ("Consolas", scaled)
    else:
        return ("DejaVu Sans Mono", scaled)

# === 马卡龙设计 token ===
PALETTE = {
    # 背景 / 卡片
    "bg":            "#fefcf8",  # 极浅奶油白（窗口底）
    "card":          "#ffffff",  # 卡片白
    "bg_soft":       "#fcfaf5",  # 极淡奶油（隔行/输入底）
    "divider":       "#f5f2ea",  # 更淡的分隔

    # 马卡龙主交互/装饰
    "mint":          "#7eb8a2",  # 薄荷绿（主交互 ★）
    "mint_soft":     "#dcf0e6",  # 浅薄荷底
    "peach":         "#f4a896",  # 蜜桃粉
    "peach_soft":    "#fce0d7",  # 浅蜜桃底
    "butter":        "#f5d97a",  # 奶油黄
    "butter_soft":   "#fcf0c8",  # 浅奶油底
    "matcha":        "#7dcea0",  # 抹茶绿（成功）
    "matcha_soft":   "#dcf5e6",  # 浅抹茶底

    # 风险/警示（柔和版）
    "rose":          "#e27d8f",  # 玫瑰粉（高风险）
    "rose_soft":     "#fcdce6",  # 浅玫瑰底
    "apricot":       "#f5b97a",  # 杏色（中警示）
    "apricot_soft":  "#fce8cf",  # 浅杏底

    # 文字三级
    "text":          "#3d3d3d",
    "text2":         "#7a7a7a",
    "text3":         "#a8a8a8",

    # 边框
    "border":        "#ece8e0",  # 淡奶油边框
    "border_strong": "#d8d2c4",

    # === 学术笔记 v2 设计 token（mockup_mixed_v2.html 同步）===
    # 主体色彩 —— 不替换马卡龙主交互，仅用于 chrome / 数据视觉
    "paper":         "#faf7f0",  # 论文纸
    "paper_edge":    "#f3ede0",  # 论文纸边 / 未选中 tab 底色
    "ink":          "#2d2a24",  # 主文字（深墨）
    "ink_soft":     "#5a544a",  # 副文字（暖灰）
    "ink_faint":    "#8a8276",  # 极弱文字（米灰）
    "rule":         "#d8cfbc",  # 主分割线
    "moss":         "#3a5a40",  # 墨绿（与马卡龙 mint 区别：更深更稳）
    "moss_soft":    "#d4dccd",  # 浅墨绿
    "moss_tint":    "#e9efe2",  # 极浅墨绿底
    "ochre":        "#d97757",  # 赭石（暖橙强调）
    "ochre_soft":   "#f5e0d3",  # 浅赭石底
    "ochre_deep":   "#8a5a16",  # 深赭石（边框/阴影）
    "honey":        "#c89a4c",  # 蜂蜜色（中严重度）
    "sage":         "#6a8e58",  # 鼠尾草绿（低严重度 / 成功）
    "stone":        "#9c8e7e",  # 暖石（数据次要）
    "sepia":        "#9c8e7e",  # 同 stone（兼容 token）
    "rose_deep":    "#b56b6b",  # 深玫瑰（高严重度）
    "rose_tint":    "#f5dada",  # 浅玫瑰底
    "fail_tint":    "#f5dada",
}

# macOS 上 tk.Button 的 fg/bg 在 Aqua 主题下可能被忽略，
# 这里用一个 dict 来统一管理按钮配色，保证白字看得清
# 马卡龙版按钮样式 —— 单一薄荷绿主色，多个语义色辅助
BUTTON_STYLES = {
    "primary":   {"bg": "#3a5a40", "fg": "#ffffff",
                  "activebackground": "#2c4632", "activeforeground": "#ffffff"},
    "success":   {"bg": "#7dcea0", "fg": "#ffffff",
                  "activebackground": "#6ab88d", "activeforeground": "#ffffff"},
    "warning":   {"bg": "#f5b97a", "fg": "#3d3d3d",
                  "activebackground": "#e8a865", "activeforeground": "#3d3d3d"},
    "secondary": {"bg": "#ffffff", "fg": "#3d3d3d",
                  "activebackground": "#f5f2ea", "activeforeground": "#3d3d3d",
                  "outline": "#d8d2c4"},  # border_strong — 轮廓按钮形态
    "ghost":     {"bg": "#fcfaf5", "fg": "#7a7a7a",
                  "activebackground": "#f5f2ea", "activeforeground": "#3d3d3d"},
    "example":   {"bg": "#e9efe2", "fg": "#3a5a40",
                  "activebackground": "#d4dfca", "activeforeground": "#2c4632"},
    "danger":    {"bg": "#e27d8f", "fg": "#ffffff",
                  "activebackground": "#c66b7c", "activeforeground": "#ffffff"},
    "pink":      {"bg": "#f0a8b8", "fg": "#ffffff",
                  "activebackground": "#d995a5", "activeforeground": "#ffffff"},
    # 学术笔记 v2 · 墨色方块按钮（选择文件用，btn-ink 风）
    "ink":         {"bg": "#2f2a26", "fg": "#faf7f0",
                  "activebackground": "#1f1b18", "activeforeground": "#faf7f0"},
    # 学术笔记 v2 · 赭石方块按钮（导出标注版 Word 用）
    "export-marked": {"bg": "#d97757", "fg": "#faf7f0",
                      "activebackground": "#c66849", "activeforeground": "#faf7f0",
                      "outline": "#8a5a16"},
}


def make_button(parent, text, command, style="primary", **kw):
    """马卡龙圆角按钮（v2 真圆版）：
    - 用 tk.Canvas.create_rectangle(radius=r) 画圆角矩形底（tk 8.6+ 稳定支持）
    - tk.Label 居中覆盖文字（place(relx=0.5, rely=0.5, anchor="center")）
    - hover/click 事件绑到 canvas + label 都触发
    - 支持 kw: radius（圆角像素）、outline（描边色）、padx、pady、font、width

    tk < 8.6 时回退到 v1 的 tk.Frame 矩形实现（不带圆角），
    保证所有 Python 版本都能用。
    """
    s = BUTTON_STYLES.get(style, BUTTON_STYLES["primary"])
    padx = kw.pop("padx", 14)
    pady = kw.pop("pady", 6)
    font = kw.pop("font", _cn_font(10, bold=True))
    width = kw.pop("width", None)
    radius = kw.pop("radius", 8)
    # 默认无描边 —— 用户反馈边框太难看；显式传 outline= 可加
    outline = kw.pop("outline", None)

    # 圆角半径按 DPI 缩放
    r = max(2, int(round(radius * _DPI_SCALE)))

    # tk Canvas create_rectangle 的 radius= 参数支持不一致
    # （部分 8.6.x 默认未启用），改用 create_polygon 自绘圆角矩形
    # —— 跨平台最稳定的方案

    def _draw_round_rect(canvas, x0, y0, x1, y1, r, **kw):
        """画圆角矩形（用 polygon + 4 段 arc）。跨平台稳定。
        每段 arc 16 点（约 5.6° / 点）→ 圆角平滑 + 不过度平滑（避免小尺寸变形）。
        polygon 坐标缩进 0.5 像素，避免 outline 渲染溢出。
        """
        outline_w = kw.get("width", 0)
        # 缩进 outline 半个像素，让描边画在 polygon 内
        if outline_w:
            shrink = outline_w / 2.0
            x0 += shrink
            y0 += shrink
            x1 -= shrink
            y1 -= shrink
        r = min(r, (x1 - x0) // 2, (y1 - y0) // 2)
        points = []
        STEPS = 16  # 圆角每段采样点数（9 会有毛刺，32 在小尺寸会变形）
        # 顶边 (从 x0+r 到 x1-r, y0)
        points.extend([(x0 + r, y0), (x1 - r, y0)])
        # 右上圆弧 (中心 x1-r, y0+r, r)
        for i in range(STEPS + 1):
            ang = -i * 90.0 / STEPS  # 0° → -90°
            cx, cy = x1 - r, y0 + r
            x = cx + r * math.cos(math.radians(ang))
            y = cy + r * math.sin(math.radians(ang))
            points.append((x, y))
        # 右边 (x1, y0+r → y1-r)
        points.extend([(x1, y0 + r), (x1, y1 - r)])
        # 右下圆弧 (中心 x1-r, y1-r)
        for i in range(STEPS + 1):
            ang = 90 - i * 90.0 / STEPS
            cx, cy = x1 - r, y1 - r
            x = cx + r * math.cos(math.radians(ang))
            y = cy + r * math.sin(math.radians(ang))
            points.append((x, y))
        # 底边
        points.extend([(x1 - r, y1), (x0 + r, y1)])
        # 左下圆弧
        for i in range(STEPS + 1):
            ang = 180 - i * 90.0 / STEPS
            cx, cy = x0 + r, y1 - r
            x = cx + r * math.cos(math.radians(ang))
            y = cy + r * math.sin(math.radians(ang))
            points.append((x, y))
        # 左边
        points.extend([(x0, y1 - r), (x0, y0 + r)])
        # 左上圆弧
        for i in range(STEPS + 1):
            ang = 270 - i * 90.0 / STEPS
            cx, cy = x0 + r, y0 + r
            x = cx + r * math.cos(math.radians(ang))
            y = cy + r * math.sin(math.radians(ang))
            points.append((x, y))
        flat = [coord for pt in points for coord in pt]
        return canvas.create_polygon(flat, smooth=False, **kw)

    use_canvas_radius = True  # 始终用 Canvas 自绘

    def on_click(e=None):
        if command:
            command()

    if use_canvas_radius:
        # === Canvas 圆角真圆版（自绘 polygon，跨平台稳定）===
        # 计算按钮尺寸：Label 自然尺寸 + padx/pady + 描边预留
        lbl_tmp = tk.Label(parent, text=text, font=font, padx=padx, pady=pady)
        lbl_tmp.update_idletasks()
        bw = lbl_tmp.winfo_reqwidth()
        bh = lbl_tmp.winfo_reqheight()
        lbl_tmp.destroy()
        if width is not None:
            bw = max(bw, width * 7)
        # 描边预留
        outline_w = 1 if outline else 0
        # canvas 实际尺寸（额外 +2 让圆角弧线不贴边）
        cw = bw + 2 * outline_w + 4
        ch = bh + 2 * outline_w + 4

        # 容器 frame（用于 pack）—— bg 用 parent 的 bg 或 PALETTE 默认
        try:
            parent_bg = parent.cget("bg")
        except Exception:
            parent_bg = PALETTE["bg"]
        btn_frame = tk.Frame(parent, bg=parent_bg, bd=0, highlightthickness=0)

        canvas = tk.Canvas(
            btn_frame, width=cw, height=ch,
            bg=parent_bg, highlightthickness=0, bd=0,
            cursor="hand2",
        )
        canvas.pack()
        # 自绘圆角矩形（在 canvas 中央 +1,+1 给圆角留呼吸）
        rect_kwargs = dict(fill=s["bg"])
        if outline:
            rect_kwargs["outline"] = outline
            rect_kwargs["width"] = 1
        else:
            # 不画描边（用户反馈边框太难看）—— macOS Aqua Tk 下 polygon
            # 即使 width=0 也会画 1px 黑边，必须显式指定 outline= 与 fill 同色
            rect_kwargs["outline"] = s["bg"]
            rect_kwargs["width"] = 0
        # polygon 画在 (1, 1, cw-1, ch-1) 范围，描边不会溢出 canvas
        rect_id = _draw_round_rect(canvas, 1, 1, cw - 1, ch - 1, r, **rect_kwargs)
        # 文字 Label 覆盖在 canvas 上方（用 place 居中）
        lbl = tk.Label(
            btn_frame, text=text,
            bg=s["bg"], fg=s["fg"],
            activebackground=s["activebackground"], activeforeground=s["activeforeground"],
            font=font, padx=padx, pady=pady,
            cursor="hand2", bd=0, highlightthickness=0,
        )
        lbl.place(relx=0.5, rely=0.5, anchor="center")

        def on_enter(e):
            canvas.itemconfig(rect_id, fill=s["activebackground"])
            if outline:
                canvas.itemconfig(rect_id, outline=s["activebackground"])
            lbl.config(bg=s["activebackground"], fg=s["activeforeground"])

        def on_leave(e):
            canvas.itemconfig(rect_id, fill=s["bg"])
            if outline:
                canvas.itemconfig(rect_id, outline=outline)
            lbl.config(bg=s["bg"], fg=s["fg"])

        for widget in (canvas, lbl):
            widget.bind("<Button-1>", on_click)
            widget.bind("<Enter>", on_enter)
            widget.bind("<Leave>", on_leave)
    else:
        # === 回退到 v1 tk.Frame 直角版（理论上不会到达）===
        if outline:
            btn_frame = tk.Frame(
                parent, bg=outline, cursor="hand2", bd=0, highlightthickness=0
            )
            inner = tk.Frame(btn_frame, bg=s["bg"], cursor="hand2", bd=0, highlightthickness=0)
            inner.pack(padx=1, pady=1)
            btn_target = inner
        else:
            btn_frame = tk.Frame(parent, bg=s["bg"], cursor="hand2", bd=0, highlightthickness=0)
            btn_target = btn_frame
        lbl = tk.Label(
            btn_target, text=text,
            bg=s["bg"], fg=s["fg"],
            activebackground=s["activebackground"], activeforeground=s["activeforeground"],
            font=font, padx=padx, pady=pady,
            cursor="hand2", bd=0, highlightthickness=0,
        )
        lbl.pack()

        def on_enter(e):
            btn_target.config(bg=s["activebackground"])
            lbl.config(bg=s["activebackground"])
            if outline:
                btn_frame.config(bg=s["activebackground"])

        def on_leave(e):
            btn_target.config(bg=s["bg"])
            lbl.config(bg=s["bg"])
            if outline:
                btn_frame.config(bg=outline)

        for widget in (btn_frame, btn_target, lbl):
            widget.bind("<Button-1>", on_click)
            widget.bind("<Enter>", on_enter)
            widget.bind("<Leave>", on_leave)

    return btn_frame


class ThesisAssistantApp:
    def __init__(self, root):
        self.root = root
        root.title("商务英语毕业论文助手 v0.3 — 仲恺农业工程学院外国语学院")
        # 窗口尺寸 / 最小尺寸在 main() 里按物理 cm 设置（跨 DPI 一致）
        # 不再在 __init__ 硬编码像素值，避免被高分屏压缩成小窗

        self.status_var = tk.StringVar(value="就绪")

        # === 外层 chrome（4px 墨绿竖条 + spine + 主体）===
        # 整窗不再直接 pack 到 root，而是放在一个 chrome Frame 内；
        # chrome 左侧用 Frame 当作 4px 墨绿书脊条
        self.chrome = tk.Frame(root, bg=PALETTE["paper"])
        self.chrome.pack(fill="both", expand=True)
        self.chrome.columnconfigure(0, minsize=4)   # 左侧墨绿书脊条
        self.chrome.columnconfigure(1, weight=1)     # 主体
        self.chrome.rowconfigure(0, minsize=0)      # 标题栏已删除
        self.chrome.rowconfigure(1, minsize=68)     # spine
        self.chrome.rowconfigure(2, weight=1)       # 主体
        # 强制 root 先 layout 一次（保证后续 grid 子能拿到 chrome 的实际宽度）
        root.update_idletasks()

        # 左：4px 墨绿书脊竖条（贯穿标题栏 + spine + 主体）
        self._spine_strip = tk.Frame(self.chrome, bg=PALETTE["moss"], width=4)
        self._spine_strip.grid(row=0, column=0, rowspan=3, sticky="nsew")
        self._spine_strip.grid_propagate(False)

        # 先建占位 tab_frame，让 header 能引用；之后再替换为真的 ttk.Frame
        self.tab_topic = tk.Frame(self.chrome)
        self.tab_library = tk.Frame(self.chrome)
        self.tab_format = tk.Frame(self.chrome)
        self._build_header(self.chrome)

        # 主体：Notebook 三个标签页（视觉上隐藏 tab 标签，用自定义 TabSwitcher）
        # 保留 ttk.Notebook 是因为它有现成的 select() / add() API，
        # tab 标签设为空字符串 + 隐藏 tab 位置，仅保留 widget 引用
        self.notebook = ttk.Notebook(self.chrome)
        self.notebook.grid(row=2, column=1, sticky="nsew", padx=0, pady=(0, 10))

        # 替换占位为真的 ttk.Frame（保留 self.tab_xxx 引用）
        for placeholder in (self.tab_topic, self.tab_library, self.tab_format):
            placeholder.destroy()
        self.tab_topic = ttk.Frame(self.notebook)
        self.tab_format = ttk.Frame(self.notebook)
        self.tab_library = ttk.Frame(self.notebook)

        # tab 标签留空（自定义 TabSwitcher 已接管视觉）
        self.notebook.add(self.tab_topic, text="")
        self.notebook.add(self.tab_library, text="")
        self.notebook.add(self.tab_format, text="")
        # 隐藏 ttk.Notebook 默认 tab 栏
        style = ttk.Style()
        try:
            style.configure("TNotebook", tabmargins=(0, 0, 0, 0))
            style.layout("TNotebook.Tab", [])
        except Exception:
            pass

        self._build_topic_tab()
        self._build_format_tab()
        self._build_library_tab()

        # 默认显示「选题建议」tab
        self.notebook.select(self.tab_topic)
        self._select_tab("tab1")  # 同步自定义 TabSwitcher 的选中态

        # 底部状态栏
        status_bar = tk.Frame(root, bg=PALETTE["card"], height=24)
        status_bar.pack(side="bottom", fill="x")
        # 1px 上边框
        tk.Frame(status_bar, bg=PALETTE["border"], height=1).pack(side="top", fill="x")
        tk.Label(
            status_bar, textvariable=self.status_var,
            anchor="w", font=_cn_font(9),
            bg=PALETTE["card"], fg=PALETTE["text2"],
            padx=20,
        ).pack(side="left", fill="y")

    # ---------------- 顶部 chrome（学术笔记 v2） ----------------
    def _build_header(self, parent):
        """spine (68px) —— brand-block + tabs 居中（避开 brand）"""

        # ============ Spine（68px）—— pack 布局 ============
        spine = tk.Frame(parent, bg=PALETTE["paper"], height=76)
        spine.grid(row=1, column=1, sticky="nsew")
        spine.pack_propagate(False)
        # 1px 底边 —— 先 pack bottom 抢 1px 高度
        tk.Frame(spine, bg=PALETTE["rule"], height=1).pack(side="bottom", fill="x")

        # 左：brand-block（260px 固定 —— 给 ZHONGKAI · B.E. THESIS 留足显示空间）
        brand_block = tk.Frame(spine, bg=PALETTE["paper"], width=260)
        brand_block.pack(side="left", fill="y")
        brand_block.pack_propagate(False)
        # 1px border-right
        tk.Frame(brand_block, bg=PALETTE["rule"], width=1).pack(side="right", fill="y")

        brand_inner = tk.Frame(brand_block, bg=PALETTE["paper"])
        brand_inner.pack(side="left", fill="both", padx=(32, 0))
        brand_pad = tk.Frame(brand_inner, bg=PALETTE["paper"])
        brand_pad.pack(anchor="w", pady=8)
        # logo
        logo_w, logo_h = 36, 44
        logo_canvas = tk.Canvas(brand_pad, width=logo_w, height=logo_h,
                                bg=PALETTE["paper"], highlightthickness=0, bd=0)
        logo_canvas.pack(side="left", padx=(0, 12))
        self._draw_book_spine(logo_canvas, logo_w, logo_h)
        text_block = tk.Frame(brand_pad, bg=PALETTE["paper"])
        text_block.pack(side="left")
        tk.Label(
            text_block, text="毕业论文助手",
            font=_cn_font(13, bold=True),
            bg=PALETTE["paper"], fg=PALETTE["ink"],
            anchor="w", padx=0, pady=0,
        ).pack(anchor="w")
        tk.Label(
            text_block, text="ZHONGKAI · B.E. THESIS",
            font=_cn_font(8),
            bg=PALETTE["paper"], fg=PALETTE["ink_faint"],
            anchor="w", padx=0, pady=0,
        ).pack(anchor="w", pady=(1, 0))

        # 中：tabs_area —— pack side=left fill=both expand=True
        tabs_area = tk.Frame(spine, bg=PALETTE["paper"], height=67)
        tabs_area.pack(side="left", fill="both", expand=True)
        tabs_area.pack_propagate(False)

        self.tab_labels = {}
        tabs = [
            ("选题分析", "tab1", self.tab_topic),
            ("选题库浏览", "tab2", self.tab_library),
            ("格式检查", "tab3", self.tab_format),
        ]
        self.tab_ribbon_frame = tk.Frame(tabs_area, bg=PALETTE["paper"])
        self.tab_ribbon_frame.place(relx=0.5, rely=1.0, anchor="s", y=-1)
        for label, key, tab_widget in tabs:
            entry = self._build_tab_ribbon(self.tab_ribbon_frame, label, key)
            entry["container"].pack(side="left", padx=3)
            self.tab_labels[key] = entry

    def _draw_book_spine(self, canvas, w, h):
        """画 v1 几何书脊：moss 矩形 + inset 阴影 + 米白横线（无字母）"""
        # 圆角：(0,2), (2,0), (w-1,0), (w-1,h-6), (w-7,h-1), (0,h-1)
        c = PALETTE["moss"]
        pts = [
            0, 2, 2, 0, w-1, 0, w-1, h-6, w-7, h-1, 0, h-1
        ]
        canvas.create_polygon(pts, fill=c, outline="")
        # inset -3px 0 0 rgba(0,0,0,0.15) —— 书脊左内侧深一档
        canvas.create_rectangle(0, 2, 4, h-2, fill="#2c4632", outline="")
        # 内嵌米白横线（3 条，模拟装订线）
        ink = PALETTE["moss_tint"]
        canvas.create_rectangle(6, 10, w-2, 12, fill=ink, outline="")
        canvas.create_rectangle(6, 36, w-2, 38, fill=ink, outline="")
        canvas.create_rectangle(6, 44, w-2, 46, fill=ink, outline="")

    def _build_tab_ribbon(self, parent, label, key):
        """单条 v2 卷边 ribbon：Canvas + 6px 上圆下平 + 底部赭石下划线 + 12pt 衬线"""
        container = tk.Frame(parent, bg=PALETTE["paper"])
        cw, ch = 132, 38
        canvas = tk.Canvas(
            container, width=cw, height=ch,
            bg=PALETTE["paper"], highlightthickness=0, bd=0,
        )
        canvas.pack(side="top")
        # 6px 上圆角，下平贴 spine（border-bottom: none）
        r = 6
        pts = [
            0, r, r, 0,
            cw-1-r, 0, cw-1, r,
            cw-1, ch, 0, ch,
        ]
        # 默认未选中：「导出报告」一致的 secondary 配色：白底深字 + 棕描边
        bg_color = "#ffffff"
        outline_color = "#d8d2c4"
        canvas.create_polygon(pts, fill=bg_color, outline=outline_color, width=1)
        # 文字（衬线 11pt，留白更舒展）
        lbl = tk.Label(
            container, text=label,
            font=("PingFang SC", 11),
            bg=bg_color, fg="#3d3d3d",
            cursor="hand2", bd=0, highlightthickness=0,
            padx=0, pady=0,
        )
        lbl.place(in_=canvas, relx=0.5, rely=0.5, anchor="center")
        # 事件
        for w in (canvas, lbl):
            w.bind("<Button-1>", lambda e, k=key: self._select_tab(k))
            w.bind("<Enter>", lambda e, k=key, ent=True: self._on_tab_hover(k, ent))
            w.bind("<Leave>", lambda e, k=key, ent=False: self._on_tab_hover(k, ent))
        return {"key": key, "container": container, "canvas": canvas, "lbl": lbl}

    def _redraw_tab_ribbon(self, key, active):
        """重绘单个 ribbon 的选中/未选中态（学术笔记 v2）"""
        entry = self.tab_labels.get(key)
        if not entry:
            return
        canvas = entry["canvas"]
        lbl = entry["lbl"]
        canvas.delete("all")
        cw = canvas.winfo_width() or 132
        ch = canvas.winfo_height() or 38
        # 上圆下平（仅顶部 6px 圆角）
        r = 6
        pts = [
            0, r, r, 0,
            cw-1-r, 0, cw-1, r,
            cw-1, ch, 0, ch,
        ]
        if active:
            # 选中：白底 + moss 加粗字 + 赭石下划线（保持 secondary 主调，仅下划线暗示选中）
            bg = "#ffffff"
            fg = PALETTE["moss"]
            lbl.config(bg=bg, fg=fg, font=("PingFang SC", 11, "bold"))
        else:
            # 未选中：「导出报告」一致的 secondary 配色
            bg = "#ffffff"
            fg = "#3d3d3d"
            lbl.config(bg=bg, fg=fg, font=("PingFang SC", 11))
        canvas.create_polygon(pts, fill=bg, outline="#d8d2c4", width=1)
        if active:
            # 底部 36×3 赭石下划线
            ul_w, ul_h = 36, 3
            canvas.create_rectangle(
                (cw - ul_w) // 2, ch - ul_h,
                (cw + ul_w) // 2, ch,
                fill=PALETTE["ochre"], outline=""
            )
        lbl.place(in_=canvas, relx=0.5, rely=0.5, anchor="center")

    def _select_tab(self, key):
        """切换 tab：调 ttk.Notebook.select + 重绘自定义 ribbon + 同步标题"""
        tab_widget_map = {
            "tab1": self.tab_topic,
            "tab2": self.tab_library,
            "tab3": self.tab_format,
        }
        title_map = {
            "tab1": "毕业论文助手 — 选题分析",
            "tab2": "毕业论文助手 — 选题库浏览",
            "tab3": "毕业论文助手 — 格式检查",
        }
        target = tab_widget_map.get(key)
        if target is not None:
            try:
                self.notebook.select(target)
            except Exception:
                pass
        for k in self.tab_labels:
            self._redraw_tab_ribbon(k, active=(k == key))
        try:
            self.notebook.update_idletasks()
        except Exception:
            pass

    def _on_tab_hover(self, key, entering):
        """未选中 ribbon hover 时 fg 变 moss"""
        # 当前是否已选中（通过 fg 判断 —— moss 表示 active）
        active_key = None
        for k in self.tab_labels:
            try:
                if self.tab_labels[k]["lbl"].cget("foreground") == PALETTE["moss"]:
                    active_key = k
                    break
            except Exception:
                continue
        if active_key == key:
            return
        entry = self.tab_labels.get(key)
        if not entry:
            return
        if entering:
            entry["lbl"].config(fg=PALETTE["moss"])
        else:
            entry["lbl"].config(fg=PALETTE["ink_soft"])

    # ---------------- Tab 1: 选题建议 ----------------
    def _build_topic_tab(self):
        f = self.tab_topic

        # lead 引子（学术笔记 v2）：10.5pt 衬线 + ink-soft + 3px 左 ochre 竖条（紧凑）
        lead_row = tk.Frame(f, bg=PALETTE["bg"])
        lead_row.pack(anchor="w", padx=20, pady=(10, 3), fill="x")
        # 左：3px ochre 竖条
        tk.Frame(lead_row, bg=PALETTE["ochre"], width=3).pack(side="left", fill="y", padx=(0, 9))
        intro_lbl = tk.Label(
            lead_row,
            text="在下方输入你的候选题目，工具会从风险评分 · 范围归属 · 聚焦性 · 字数与表达 · 创新度 五个维度分析，并给出撞题明细和改进建议。",
            font=_serif_font(10.5), fg=PALETTE["ink_soft"], bg=PALETTE["bg"],
            justify="left", anchor="w", padx=0, pady=0,
        )
        intro_lbl.pack(side="left", fill="x", expand=True)
        # 跟随窗口缩放动态换行
        def _resize_intro(event):
            intro_lbl.config(wraplength=max(200, event.width - 40))
        f.bind("<Configure>", _resize_intro)

        # 输入区（白卡 + 1px 淡奶油边框）
        input_card = tk.Frame(
            f, bg=PALETTE["card"],
            highlightthickness=1, highlightbackground=PALETTE["border"],
            bd=0,
        )
        input_card.pack(fill="x", padx=20, pady=(2, 6))

        # 卡内标题
        tk.Label(
            input_card, text="你的题目",
            font=_cn_font(10), fg=PALETTE["text3"],
            bg=PALETTE["card"],
        ).pack(anchor="w", padx=16, pady=(8, 0))

        self.topic_entry = tk.Text(
            input_card, height=2, font=_cn_font(11),
            wrap="word", relief="flat", bd=0,
            bg=PALETTE["bg_soft"], fg=PALETTE["text"],
            highlightthickness=0,
            padx=12, pady=8,
        )
        self.topic_entry.pack(fill="x", padx=12, pady=(4, 8))
        self.topic_entry.insert("1.0", "示例：目的论视角下跨境电商产品描述翻译策略研究——以SHEIN为例")

        # 操作按钮
        btn_frame = tk.Frame(f, bg=PALETTE["bg"])
        btn_frame.pack(fill="x", padx=20, pady=(0, 2))

        make_button(
            btn_frame, "🔍 分析选题", self._on_analyze_topic,
            style="primary", padx=18, pady=6, radius=8,
            font=_cn_font(11, bold=True),
        ).pack(side="left")

        make_button(
            btn_frame, "清空",
            lambda: self.topic_entry.delete("1.0", "end"),
            style="ghost", padx=14, pady=6, radius=8,
            font=_cn_font(10),
        ).pack(side="left", padx=10)

        # 快捷示例
        tk.Label(btn_frame, text="快速试用：",
                 font=_cn_font(9), fg=PALETTE["text3"], bg=PALETTE["bg"]).pack(side="left", padx=(20, 5))

        examples = [
            "目的论视角下SHEIN跨境电商翻译",
            "直播带货话语分析",
        ]
        for ex in examples:
            make_button(
                btn_frame, ex,
                lambda e=ex: self.topic_entry.delete("1.0", "end") or self.topic_entry.insert("1.0", e),
                style="example", padx=10, pady=4, radius=6,
                font=_cn_font(9),
            ).pack(side="left", padx=2)

        # 结果区（5 列引文卡 + 改进建议便签块的容器）
        self._topic_overview = tk.Frame(f, bg=PALETTE["bg"])
        self._topic_overview.pack(fill="both", expand=True, padx=20, pady=(2, 10))
        # 占位提示
        tk.Label(
            self._topic_overview,
            text="  ↑ 在上方输入题目，点击「分析选题」",
            font=_cn_font(10), fg=PALETTE["text3"], bg=PALETTE["bg"],
            justify="left", anchor="w",
        ).pack(anchor="nw", padx=4, pady=6)

    def _on_analyze_topic(self):
        title = self.topic_entry.get("1.0", "end").strip()
        if not title or title.startswith("示例"):
            messagebox.showwarning("提示", "请先输入你的题目")
            return
        try:
            result = analyze_topic(title)
            self._last_topic_result = result
            parsed = self._parse_topic_result(result)
            self._render_topic_visual(parsed, result)
            self.status_var.set(f"选题分析完成 — 风险 {result['risk_level']}")
        except Exception as e:
            messagebox.showerror("错误", f"分析失败: {e}")

    # ---- Tab1 · 引文卡 + 改进建议 ----
    def _parse_topic_result(self, r):
        """把 analyze_topic 返回值拆给渲染层。
        包含：5 维度摘要、关键词、风险明细、撞题 Top 3、方向热度、字数细节、改进建议。"""
        sr = r.get("scope_result", {})
        nz = r.get("narrowness", {})
        li = r.get("length_info", {})
        inn = r.get("innovation", {})
        ti = r.get("theory_info", {})
        oi = r.get("object_info", {})
        dh = r.get("direction_heat", {})
        risk_level = r.get("risk_level", "🟢 低")
        risk_score = r.get("risk_score", 0)
        risks = r.get("risks", []) or []
        matches = r.get("matches_in_library", []) or []
        ref_topics = r.get("reference_topics", []) or []
        keywords = r.get("extracted_keywords", []) or []

        # ---- 理论热度榜（从 library_insights 取具体理论名）----
        insights = r.get("library_insights", {}) or {}
        all_top_theories = insights.get("top_theories", []) or []
        # 前 5 = 高频（撞题风险高），6+ = 冷门（按方向过滤）
        hot_theories = all_top_theories[:5]
        cold_all = all_top_theories[5:]
        # 复用 analyzer 中的方向过滤逻辑
        user_dir = (r.get("suggested_direction") or "其他")
        if user_dir == "其他":
            cold_theories = cold_all
        else:
            cold_theories = [t for t in cold_all if (len(t) >= 3 and t[2] in (user_dir, "通用"))]
            if len(cold_theories) < 3:
                cold_theories = cold_all[:5]
        cold_theories = cold_theories[:5]

        # ---- 抽取操作性建议 ----
        clean = []
        skip_patterns = (
            "【", "• [", "共享 ", "撞题 Top", "说明：", "📋 ",
            "📚 ",   # 参考题目
        )
        # 三轮共用的「理论热度榜标题」排除项
        theory_header_prefixes = ("⚠️", "✨")

        for s in r.get("suggestions", []) or []:
            for ln in s.splitlines():
                ln = ln.strip()
                if not ln or len(ln) < 6 or len(ln) > 50:
                    continue
                if any(ln.startswith(p) for p in skip_patterns):
                    continue
                if ln.startswith(theory_header_prefixes):
                    continue
                # 兜底：去掉 emoji 后以文本开头识别
                if ln.lstrip("⚠️✨ ").startswith("谨慎使用") or ln.lstrip("⚠️✨ ").startswith("以下理论在往届"):
                    continue
                if ln.startswith(("✅ ", "🟡 ", "🔴 ", "🟢 ", "📏 ", "📝 ", "⚠️ ", "💡", "• ", "📚 ")):
                    continue
                if ln.startswith("💡"):
                    txt = ln.lstrip("•·-— 💡").strip()
                    if txt and txt not in clean:
                        clean.append(txt)
            if len(clean) >= 4:
                break
        if len(clean) < 4:
            for s in r.get("suggestions", []) or []:
                for ln in s.splitlines():
                    ln = ln.strip()
                    if not ln or len(ln) < 6 or len(ln) > 50:
                        continue
                    if any(ln.startswith(p) for p in skip_patterns):
                        continue
                    if ln.startswith(theory_header_prefixes):
                        continue
                    if ln.lstrip("⚠️✨ ").startswith("谨慎使用") or ln.lstrip("⚠️✨ ").startswith("以下理论在往届"):
                        continue
                    if ln.startswith(("✅ ", "🟡 ", "🔴 ", "🟢 ", "📏 ", "📝 ", "⚠️ ", "💡", "• ", "📚 ")):
                        continue
                    if any(k in ln for k in ("改法", "可考虑", "可改", "可换", "建议改", "建议：", "建议补",
                                              "去撞题", "三种", "换案例", "换理论", "缩子方向")):
                        txt = ln.lstrip("•·-— ").strip()
                        if txt and txt not in clean:
                            clean.append(txt)
                    if len(clean) >= 4:
                        break
                if len(clean) >= 4:
                    break
        if len(clean) < 4:
            for s in r.get("suggestions", []) or []:
                for ln in s.splitlines():
                    ln = ln.strip()
                    if not ln or len(ln) < 6 or len(ln) > 50:
                        continue
                    if any(ln.startswith(p) for p in skip_patterns):
                        continue
                    if ln.startswith(("✅ ", "🟡 ", "🔴 ", "🟢 ", "📏 ", "📝 ", "⚠️", "✨", "💡", "• ", "📚 ")):
                        continue
                    # 排除理论热度榜的两行标题（理论名已在「理论热度榜」区块展示）
                    if ln.startswith("谨慎使用") or ln.startswith("以下理论在往届"):
                        continue
                    if any(k in ln for k in ("谨慎", "可考虑", "建议", "可选", "改", "换", "少用")):
                        txt = ln.lstrip("•·-— ⚠️✨").strip()
                        if txt and txt not in clean:
                            clean.append(txt)
                    if len(clean) >= 4:
                        break
                if len(clean) >= 4:
                    break
        if len(clean) < 4:
            defaults = [
                "明确具体研究对象（如某行业/某文本类型）",
                "聚焦一个明确的小切面（去除冗余修饰）",
                "在题目中点出使用的核心理论",
                "与现有往届选题拉开关键词差异",
            ]
            for d in defaults:
                if d not in clean:
                    clean.append(d)
                if len(clean) >= 4:
                    break
        clean = clean[:4]

        return {
            "risk": {"score": risk_score, "level": risk_level, "items": risks},
            "scope": {"icon": sr.get("icon", ""), "category": sr.get("matched_category", "—"),
                      "description": sr.get("description", "")},
            "narrow": {"level": nz.get("level", "—"), "assessment": nz.get("assessment", ""),
                       "has_theory": nz.get("has_theory", False),
                       "has_object": nz.get("has_object", False),
                       "has_theme": nz.get("has_theme", False),
                       "missing": nz.get("missing", [])},
            "length": {"assessment": li.get("assessment", "—"), "details": li.get("details", ""),
                       "kind": li.get("kind", "")},
            "innovation": {"level": inn.get("level", "—"), "detail": inn.get("detail", "")},
            "theory": ti,
            "object": oi,
            "direction_heat": dh,
            "extracted_keywords": keywords,
            "matches": matches[:3],
            "reference_topics": ref_topics[:3],
            "suggestions": clean,
            "hot_theories": [(t[0], t[1]) for t in hot_theories],   # (name, count)
            "cold_theories": [(t[0], t[1], t[2]) for t in cold_theories],  # (name, count, direction)
        }

    def _render_topic_visual(self, p, full_result=None):
        """学术笔记 v2 简化：上下滚动文本，包含详细分维分析。"""
        from tkinter import scrolledtext
        # 清空旧组件
        for w in self._topic_overview.winfo_children():
            w.destroy()

        # 配色
        GREEN_FG = PALETTE["moss"]
        YELLOW_FG = PALETTE["ochre"]
        ROSE_FG = PALETTE["rose_deep"]
        INK = PALETTE["ink"]
        INK_SOFT = PALETTE.get("ink_soft", PALETTE.get("text2", "#5a544a"))
        INK_FAINT = PALETTE.get("ink_faint", PALETTE.get("text3", "#8a847a"))
        OCHRE = PALETTE["ochre"]
        OCHRE_SOFT = PALETTE["ochre_soft"]
        RULE = PALETTE["rule"]

        # 创建 ScrolledText
        txt = scrolledtext.ScrolledText(
            self._topic_overview,
            wrap="word",
            bg=PALETTE["bg"], fg=INK,
            font=_cn_font(11),
            relief="flat", borderwidth=0,
            highlightthickness=1, highlightbackground=RULE,
            padx=16, pady=10,
            spacing1=0, spacing2=0, spacing3=0,
        )
        txt.pack(fill="both", expand=True)

        # 分隔线：固定 40 字符，在小窗口下也不会换行
        DIVIDER = "─" * 40

        # ===== tag 配置 =====
        # 紧凑字号：h1/h2 略小，body 11pt，key 10pt，val 11.5pt
        txt.tag_configure("h1", font=_cn_font(12, bold=True), foreground=OCHRE)
        txt.tag_configure("h2", font=_cn_font(11, bold=True), foreground=OCHRE)
        txt.tag_configure("h2_line", foreground=OCHRE_SOFT)
        txt.tag_configure("key", font=_cn_font(10, bold=True), foreground=INK_FAINT)
        txt.tag_configure("val_g", font=_cn_font(11), foreground=GREEN_FG)
        txt.tag_configure("val_y", font=_cn_font(11), foreground=YELLOW_FG)
        txt.tag_configure("val_r", font=_cn_font(11), foreground=ROSE_FG)
        txt.tag_configure("body", font=_cn_font(11), foreground=INK_SOFT)
        txt.tag_configure("body_emph", font=_cn_font(11, bold=True), foreground=INK)
        txt.tag_configure("bullet", font=_cn_font(11), foreground=OCHRE)
        txt.tag_configure("item_id", font=_cn_font(10), foreground=INK_FAINT)
        txt.tag_configure("item_year", font=_cn_font(10, bold=True), foreground=INK)
        txt.tag_configure("item_title", font=_cn_font(11), foreground=INK)
        txt.tag_configure("item_meta", font=_cn_font(10), foreground=INK_SOFT)
        txt.tag_configure("sep", foreground=RULE)
        txt.tag_configure("match_overlap", font=_cn_font(10), foreground=OCHRE)
        txt.tag_configure("risk_item", font=_cn_font(11), foreground=ROSE_FG)
        txt.tag_configure("kw_chip", font=_cn_font(11), foreground=INK)
        txt.tag_configure("kw_sep", font=_cn_font(11), foreground=RULE)
        # 三要素 mark —— 用纯文本符号 + 相同字号，避免 emoji 撑大
        txt.tag_configure("mark_ok", font=_cn_font(10, bold=True), foreground=GREEN_FG)
        txt.tag_configure("mark_no", font=_cn_font(10, bold=True), foreground=ROSE_FG)
        txt.tag_configure("warn", font=_cn_font(11, bold=True), foreground=ROSE_FG)
        txt.tag_configure("tip", font=_cn_font(11, bold=True), foreground=YELLOW_FG)

        def line(parts):
            """parts = list of (text, tag_name) tuples"""
            for t, tag in parts:
                txt.insert("end", t, tag)
            txt.insert("end", "\n")

        def section(title):
            line([(f"  {title}", "h2")])
            line([(DIVIDER, "h2_line")])

        def gap(n=1):
            for _ in range(n):
                line([("", "body")])

        def kv(key, val, kind="g"):
            """key 用字间空格模拟 letter-spacing；val 用 kind 配色。
            长 val 自动拆为「摘要 + 缩进明细」两行，避免窄窗口下换行后行间留白。"""
            tag = f"val_{kind}"
            # 标题用字间空格
            key_disp = " ".join(list(key))
            # 估算视觉宽度（中文 2 个 ASCII 算 1）
            def visual_width(s):
                w = 0
                for c in s:
                    if ord(c) > 0x2E80:  # 中日韩
                        w += 2
                    else:
                        w += 1
                return w
            # 标题列宽度：8 个 ASCII（"  " + "字"间空格*4 + "│" 占位）≈ 14
            KEY_COL_W = 16   # 标题列视觉宽度（含缩进）
            val_w = visual_width(val)
            # 如果 val 视觉宽度 > 36（约 18 个汉字），尝试按"·"拆分
            if val_w > 36 and "·" in val:
                parts = val.split("·", 1)
                head = parts[0].strip()
                tail = "·" + parts[1].strip() if len(parts) > 1 else ""
                line(
                    [("  ", "body"), (key_disp, "key"),
                     ("    ", "body"), (head, tag)]
                )
                if tail:
                    line(
                        [("              ", "body"), (tail, tag)]
                    )
            else:
                line(
                    [("  ", "body"), (key_disp, "key"),
                     ("    ", "body"), (val, tag)]
                )

        # ===== 1. 输入题目 =====
        if full_result:
            title_text = full_result.get("input_title", "")
            line([("分 析 对 象", "h1")])
            line([(DIVIDER, "h2_line")])
            line([("  ", "body"), (title_text, "body_emph")])
            gap()

        # ===== 2. 五维度摘要 =====
        risk_emoji = p["risk"]["level"]
        scope_icon = p["scope"]["icon"]
        narrow_lv = p["narrow"]["level"]
        narrow_assess = p["narrow"]["assessment"].replace("✅ ", "").replace("⚠️ ", "").strip()
        li_assess = p["length"]["assessment"]
        li_details = p["length"]["details"]
        inn_lv = p["innovation"]["level"]
        inn_short = inn_lv.replace("🟢 ", "").replace("🟡 ", "").replace("🔴 ", "")

        risk_kind = "r" if "🔴" in risk_emoji else "y" if "🟡" in risk_emoji else "g"
        scope_kind = "g" if scope_icon == "✅" else "r"
        narrow_kind = "g" if narrow_lv == "聚焦" else "y"
        li_kind = "g" if ("合适" in li_assess or "通过" in li_assess) else "y"
        inn_kind = "r" if "高度撞题" in inn_lv or "🔴" in inn_lv else \
                   "y" if "可能撞题" in inn_lv or "🟡" in inn_lv else "g"

        section("五 维 度 摘 要")
        kv("风险评分", f"{risk_emoji}  {p['risk']['score']} / 100", risk_kind)
        kv("范围归属", f"{scope_icon}  商务英语 · {p['scope']['category']}", scope_kind)
        kv("聚焦性",   f"{narrow_lv} · {narrow_assess}", narrow_kind)
        kv("字数·表达", f"{li_assess} · {li_details}", li_kind)
        kv("创新度",   inn_short, inn_kind)
        gap()

        # ===== 3. 关键词 + 理论 + 对象 =====
        section("提 取 的 关 键 词")
        kws = p.get("extracted_keywords", [])
        if kws:
            for i, kw in enumerate(kws):
                if i > 0:
                    txt.insert("end", "  •  ", "kw_sep")
                else:
                    txt.insert("end", "  ", "kw_sep")
                txt.insert("end", kw, "kw_chip")
            txt.insert("end", "\n")
        else:
            line([("  （未提取到关键词）", "body")])

        # 理论 + 对象细节
        ti = p.get("theory", {}) or {}
        oi = p.get("object", {}) or {}
        gap()
        section("理 论 与 对 象 识 别")
        kv("理论框架", ", ".join(ti.get("matched_theories", []) or []) or "未识别", "y" if not ti.get("matched_theories") else "g")
        kv("表达方式", ", ".join(ti.get("expression_patterns", []) or []) or "未识别", "y" if not ti.get("expression_patterns") else "g")
        kv("研究对象", (oi.get("object_name") or "未识别")[:30], "y" if not oi.get("object_name") else "g")
        kv("方向分类", f"{p.get('direction_heat', {}).get('direction', '—')} · "
                       f"{p.get('direction_heat', {}).get('assessment', '—')} · "
                       f"{p.get('direction_heat', {}).get('percentage', '—')}", "g")
        gap()

        # ===== 4. 字数与表达详情 =====
        section("字 数 与 表 达 详 情")
        li_kind_str = p["length"].get("kind", "")
        if li_kind_str == "zh":
            range_text = "规范 15-25 字（核心 20 字）"
        elif li_kind_str == "en":
            range_text = "规范 10-20 词（核心 15 词）"
        else:
            range_text = "—"
        kv("字数范围", range_text, "body")
        kv("当前评估", li_assess, li_kind)
        kv("详情", li_details, li_kind)
        # 聚焦性三要素（用纯文本符号而非 emoji，避免行高突变）
        miss = p["narrow"].get("missing", [])
        miss_text = "、".join(miss) if miss else "齐备"
        line([("  · 三要素：理论 ", "body"),
              ("[OK]" if p["narrow"].get("has_theory") else "[NO]", "mark_ok" if p["narrow"].get("has_theory") else "mark_no"),
              (" · 对象 ", "body"),
              ("[OK]" if p["narrow"].get("has_object") else "[NO]", "mark_ok" if p["narrow"].get("has_object") else "mark_no"),
              (" · 主题 ", "body"),
              ("[OK]" if p["narrow"].get("has_theme") else "[NO]", "mark_ok" if p["narrow"].get("has_theme") else "mark_no")])
        # 缺失项单独一行
        line([("      ", "body"),
              (f"（缺：{miss_text}）" if miss else "（齐备）", "body")])
        gap()

        # ===== 5. 风险明细 =====
        risk_items = p["risk"].get("items", [])
        if risk_items:
            section("风 险 明 细")
            for r in risk_items[:8]:
                line([("  !  ", "bullet"), (r, "risk_item")])
            gap()

        # ===== 6. 撞题 Top 3 =====
        matches = p.get("matches", [])
        if matches:
            section("撞 题 Top 3")
            for m in matches:
                line([("  -  ", "bullet"),
                      (f"[{m['id']}]", "item_id"),
                      (f" {m['year']}届·", "item_year"),
                      (f"{m['direction']}", "item_year")])
                line([("      ", "body"), (m["title"][:70], "item_title")])
                # 文字重合度（高亮颜色）
                text_pct = m.get("text_overlap", 0)
                if text_pct >= 70:
                    text_tag = "val_r"
                    label = "重合高"
                elif text_pct >= 50:
                    text_tag = "val_y"
                    label = "重合中"
                else:
                    text_tag = "item_meta"
                    label = "重合低"
                line([("      题目文字重合 ", "item_meta"),
                      (f"{text_pct:.0f}%", text_tag),
                      (f"（{label}）", text_tag if text_pct >= 50 else "item_meta")])
                gap()

        # ===== 7. 理论热度榜 =====
        hot = p.get("hot_theories", [])
        cold = p.get("cold_theories", [])
        if hot or cold:
            section("理 论 热 度 榜")
            if hot:
                line([("  !  ", "bullet"),
                      ("谨慎使用（往届高频前 5 名，撞题风险高）：", "warn")])
                for th, c in hot:
                    line([("      -  ", "bullet"),
                          (th, "item_title"),
                          (f"（往届用过 {c} 次）", "item_meta")])
                gap()
            if cold:
                line([("  >  ", "bullet"),
                      ("以下理论在往届选题中较少使用，可考虑。", "tip"),
                      ("建议自行寻找合适理论。", "tip")])
                for th, c, d in cold:
                    tag = f" [{d}]" if d != "通用" else ""
                    line([("      -  ", "bullet"),
                          (th, "item_title"),
                          (f"（往届用过 {c} 次）", "item_meta"),
                          (tag, "item_meta")])
                gap()

        # ===== 8. 改进建议 =====
        section("改 进 建 议")
        for i, s in enumerate(p["suggestions"][:4], start=1):
            line([("  >  ", "bullet"), (s, "body")])
            gap()

        # ===== 9. 参考题目 =====
        ref = p.get("reference_topics", [])
        if ref:
            section("近三年参考题目")
            for t in ref:
                line([("  -  ", "bullet"),
                      (f"[{t['id']}]", "item_id"),
                      (f" {t.get('year', '?')}届·", "item_year"),
                      (f"{t.get('direction', '其他')}", "item_year")])
                line([("      ", "body"), (t["title"][:70], "item_title")])
                kws_text = ", ".join((t.get("keywords") or [])[:4])
                if kws_text:
                    line([("      关键词：", "item_meta"), (kws_text, "match_overlap")])
                gap()

        txt.configure(state="disabled")

    def _make_quote_card(self, parent, col, key, val, bg, fg):
        """单张引文卡（学术笔记 v2）：
        - 右上 36pt 衬线 引号（赭石浅色）
        - key 行 9.5pt sans UPPERCASE weight 700 letter-spacing 2px 模拟（中文用空格填充）
        - 1px rule 底边
        - val 12pt 衬线 + 软底色徽章
        """
        card = tk.Frame(
            parent, bg=PALETTE["paper"],
            highlightthickness=1, highlightbackground=PALETTE["rule"],
            bd=0,
        )
        card.grid(row=0, column=col, padx=4, sticky="nsew")
        card.grid_propagate(False)
        card.configure(height=120)
        # 右上装饰引号 —— 36pt 衬线 + 赭石浅色（模拟 0.18 opacity）
        # tkinter 不支持 text opacity，用略浅的赭石 (#e89b7e) 模拟
        tk.Label(
            card, text="”",
            font=_cn_font(36), fg="#e8a589",
            bg=PALETTE["paper"],
        ).place(relx=0.97, rely=0.02, anchor="ne")
        inner = tk.Frame(card, bg=PALETTE["paper"])
        inner.pack(fill="both", expand=True, padx=12, pady=(10, 10))
        # key 行：9.5pt sans weight 700，letter-spacing 用 Font 对象的 slant 不可，
        # 改用字号 9.5pt + bold + 字间多打空格模拟
        key_display = " ".join(list(key))  # 字间插空格 → 视觉 letter-spacing
        tk.Label(
            inner, text=key_display,
            font=_cn_font(9, bold=True), fg=PALETTE["ink_faint"],
            bg=PALETTE["paper"],
        ).pack(anchor="w")
        # 1px rule 底边
        tk.Frame(inner, bg=PALETTE["rule"], height=1).pack(fill="x", pady=(3, 6))
        # val 软底色徽章：12pt 衬线 + wraplength（响应 card 宽度）
        pill = tk.Frame(inner, bg=bg)
        pill.pack(anchor="w", pady=(0, 2), fill="x")
        lbl = tk.Label(
            pill, text=val,
            font=_serif_font(12), fg=fg, bg=bg,
            padx=8, pady=4,
            justify="left", anchor="w",
            wraplength=140,
        )
        lbl.pack(fill="x", anchor="w")
        # 随 card 宽度变化重新设置 wraplength
        def _resize_card(event):
            try:
                w = max(80, event.width - 24)
                lbl.config(wraplength=w)
            except Exception:
                pass
        card.bind("<Configure>", _resize_card)
        # 卡底部 1px 半透明 rule 收口（card::before 模拟）
        tk.Frame(card, bg=PALETTE["rule"], height=1).pack(side="bottom", fill="x", padx=8)

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
                lines.append(f"     题目文字重合 {m.get('text_overlap', 0):.0f}%")
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

        # lead 引子（学术笔记 v2 · 与 Tab1 一致）：3px 左 ochre + 10.5pt ink-soft（紧凑）
        lead_row = tk.Frame(f, bg=PALETTE["bg"])
        lead_row.pack(anchor="w", padx=20, pady=(10, 4), fill="x")
        tk.Frame(lead_row, bg=PALETTE["ochre"], width=3).pack(side="left", fill="y", padx=(0, 9))
        self.fmt_lead_lbl = tk.Label(
            lead_row,
            text="选择你的论文 Word 文件（.docx），工具会按仲恺商英毕业论文格式规范（2025-12 修订）逐项检查格式。",
            font=_serif_font(10.5), fg=PALETTE["ink_soft"], bg=PALETTE["bg"],
            justify="left", anchor="w",
        )
        self.fmt_lead_lbl.pack(side="left", fill="x", expand=True)
        def _resize_fmt_lead(event):
            self.fmt_lead_lbl.config(wraplength=max(200, event.width - 40))
        f.bind("<Configure>", _resize_fmt_lead, add="+")

        file_frame = tk.Frame(f, bg=PALETTE["bg"])
        file_frame.pack(fill="x", padx=20, pady=10)

        self.format_file_var = tk.StringVar()

        # 按钮组先 pack（左侧固定宽度），Entry 后 pack 拿走剩余空间
        btns = tk.Frame(file_frame, bg=PALETTE["bg"])
        btns.pack(side="right")
        make_button(btns, "📁 选择文件", self._on_choose_file,
                    style="secondary", radius=8, padx=10, pady=6,
                    font=_cn_font(10)).pack(side="left")
        make_button(btns, "🔍 开始检查", self._on_check_format,
                    style="secondary", radius=8, padx=10, pady=6,
                    font=_cn_font(10)).pack(side="left", padx=4)
        make_button(btns, "📝 导出标注 Word", self._on_export_marked_docx,
                    style="secondary", radius=8, padx=10, pady=6,
                    font=_cn_font(10)).pack(side="left", padx=4)
        make_button(btns, "📄 导出报告", self._on_export_report,
                    style="secondary", radius=8, padx=10, pady=6,
                    font=_cn_font(10)).pack(side="left", padx=4)

        entry = tk.Entry(
            file_frame, textvariable=self.format_file_var,
            font=_cn_font(10), relief="flat", bd=0,
            bg=PALETTE["card"], fg=PALETTE["text"],
            highlightthickness=1, highlightbackground=PALETTE["border"],
        )
        entry.pack(side="left", fill="x", expand=True, padx=(0, 10), ipady=6)

        # 结果区（白卡 + 1px 边框）
        result_card = tk.Frame(
            f, bg=PALETTE["card"],
            highlightthickness=1, highlightbackground=PALETTE["border"],
            bd=0,
        )
        result_card.pack(fill="both", expand=True, padx=20, pady=10)
        tk.Label(
            result_card, text="检查结果",
            font=_cn_font(10), fg=PALETTE["text3"],
            bg=PALETTE["card"],
        ).pack(anchor="w", padx=16, pady=(12, 0))

        # === 视觉概览容器（圆环 + 严重度 pill + 问题清单）===
        # 初始为空，_render_format_visual() 检查完成后填充
        self._format_overview = tk.Frame(result_card, bg=PALETTE["card"])
        self._format_overview.pack(fill="x", padx=12, pady=(6, 6))

        # === 完整文本（ScrolledText 保留为详情）===
        self.format_result = scrolledtext.ScrolledText(
            result_card, font=_mono_font(10), wrap="word",
            bg=PALETTE["bg_soft"], fg=PALETTE["text"],
            relief="flat", bd=0, highlightthickness=0,
            padx=12, pady=10, state="disabled",
        )
        self.format_result.pack(fill="both", expand=True, padx=12, pady=(0, 12))

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
            # 视觉概览（圆环 + pill + 问题清单）
            self._render_format_visual(result)
            # 完整文本（保留为详情）
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

    # --- 圆环 + pill + 问题清单 ---
    GRADE_META = {
        "A": ("#7dcea0", "#dcf5e6", "优秀", 95),   # 抹茶
        "B": ("#7eb8a2", "#dcf0e6", "良好", 87),   # 薄荷
        "C": ("#f5b97a", "#fce8cf", "及格", 75),   # 杏
        "D": ("#e27d8f", "#fcdce6", "需大幅修改", 50),  # 玫瑰
    }
    SEVERITY_META = {
        "高": ("#e27d8f", "#fcdce6", "🔴"),
        "中": ("#f5b97a", "#fce8cf", "🟡"),
        "低": ("#7dcea0", "#dcf5e6", "🟢"),
    }

    def _render_format_visual(self, r):
        """渲染视觉概览：等级圆环 + 严重度 pill + 问题清单。
        写入 self._format_overview 容器；调用前 _format_overview 应已建立。
        """
        # 清空旧组件
        for w in self._format_overview.winfo_children():
            w.destroy()

        # 解析 grade 字符串
        grade_str = r.get("grade", "B 良好")
        letter = grade_str.split()[0] if grade_str else "B"
        meta = self.GRADE_META.get(letter, self.GRADE_META["B"])
        ring_color, ring_soft, level_name, score = meta

        sev = r.get("stats", {}).get("issues_by_severity", {})
        issues = r.get("issues", [])

        # === 主行：左侧 1/5（圆环 + pill 纵向堆叠）+ 右侧 4/5（问题清单）===
        main_row = tk.Frame(self._format_overview, bg=PALETTE["card"])
        main_row.pack(fill="x", pady=(4, 6))
        main_row.columnconfigure(0, weight=1, minsize=180)   # 左栏：圆环 + pill（约 1/5）
        main_row.columnconfigure(1, weight=4)                  # 右栏：问题清单（约 4/5）
        main_row.rowconfigure(0, weight=1)

        # ---- 左栏：等级圆 + 严重度分布（学术笔记 v2：140 实心环 + 56pt 字母 + 22pt 分数）----
        left_col = tk.Frame(main_row, bg=PALETTE["card"])
        left_col.grid(row=0, column=0, sticky="nsew", padx=(8, 4), pady=4)
        left_col.columnconfigure(0, weight=1)

        # 等级圆（solid moss ring + 1px outer halo @ 0.3 opacity + 56pt letter inside + 22pt score below）
        grade_size = 110
        # 外圈容器（容纳 halo + 主圆 + 字母 + 标签）
        grade_box = tk.Frame(left_col, bg=PALETTE["card"])
        grade_box.grid(row=0, column=0, pady=(0, 4))
        grade_box.columnconfigure(0, weight=1)

        # 主圆 canvas
        ring_canvas = tk.Canvas(
            grade_box, width=grade_size, height=grade_size,
            bg=PALETTE["card"], highlightthickness=0, bd=0,
        )
        ring_canvas.grid(row=0, column=0)
        self._draw_grade_circle(ring_canvas, grade_size, ring_color, letter, level_name)

        # 18pt ink serif 分数（圆下方 · letter-spacing）
        score_row = tk.Frame(grade_box, bg=PALETTE["card"])
        score_row.grid(row=1, column=0, pady=(4, 0))
        tk.Label(
            score_row, text=str(score),
            font=("Times New Roman", 18, "bold"),
            fg=PALETTE["ink"], bg=PALETTE["card"],
        ).pack(side="left")
        tk.Label(
            score_row, text=" 分",
            font=_cn_font(10),
            fg=PALETTE["ink_faint"], bg=PALETTE["card"],
        ).pack(side="left", anchor="s", pady=(0, 2))

        # 严重度小标签（纵向 3 行紧凑,填满左栏宽度）
        tk.Label(
            left_col, text="问题严重度分布",
            font=_cn_font(8), fg=PALETTE["text3"], bg=PALETTE["card"],
        ).grid(row=1, column=0, sticky="ew", pady=(4, 2))
        sev_box = tk.Frame(left_col, bg=PALETTE["card"])
        sev_box.grid(row=2, column=0, sticky="ew")
        sev_box.columnconfigure(0, weight=1)
        for i, sev_key in enumerate(("高", "中", "低")):
            count = sev.get(sev_key, 0)
            self._draw_sev_chip(sev_box, sev_key, count, i)

        # ---- 右栏：问题清单（占 4/5, 高度更大）----
        # 学术笔记 v2 简化：使用 ScrolledText 上下滚动文本渲染
        if issues:
            from tkinter import scrolledtext
            list_section = tk.Frame(
                main_row,
                bg=PALETTE["paper"],
                highlightthickness=1, highlightbackground=PALETTE["rule"],
                bd=0,
            )
            list_section.grid(row=0, column=1, sticky="nsew", padx=(4, 8), pady=4)
            list_section.columnconfigure(0, weight=1)
            list_section.rowconfigure(2, weight=1)

            # 标题栏
            header_row = tk.Frame(list_section, bg=PALETTE["paper_edge"])
            header_row.grid(row=0, column=0, sticky="ew")
            header_row.columnconfigure(0, weight=1)
            spaced_title = "  ".join(list("问题清单"))
            tk.Label(
                header_row, text=f"{spaced_title}　{len(issues)} 条 · 按严重度 ↓",
                font=_cn_font(13, bold=True), fg=PALETTE["ink"],
                bg=PALETTE["paper_edge"],
            ).grid(row=0, column=0, sticky="w", padx=18, pady=10)
            tk.Frame(list_section, bg=PALETTE["rule"], height=1).grid(
                row=1, column=0, sticky="ew"
            )

            # 滚动文本区域
            list_text = scrolledtext.ScrolledText(
                list_section,
                wrap="word",
                bg=PALETTE["paper"],
                fg=PALETTE["ink"],
                font=_cn_font(11),
                relief="flat",
                borderwidth=0,
                highlightthickness=0,
                padx=14, pady=10,
                spacing1=4, spacing3=8,  # 段前段后间距
            )
            list_text.grid(row=2, column=0, sticky="nsew")

            # tag 配置
            list_text.tag_configure("sev_h", font=("Times New Roman", 10, "bold italic"),
                                    foreground="#b56b6b")  # rose
            list_text.tag_configure("sev_m", font=("Times New Roman", 10, "bold italic"),
                                    foreground="#c89a4c")  # honey
            list_text.tag_configure("sev_l", font=("Times New Roman", 10, "bold italic"),
                                    foreground="#7dcea0")  # sage
            list_text.tag_configure("rule_marker", font=("Times New Roman", 12, "bold italic"),
                                    foreground=PALETTE["ochre"])
            list_text.tag_configure("rule_text", font=_serif_font(12, bold=True),
                                    foreground=PALETTE["ink"])
            list_text.tag_configure("loc_pill", font=_cn_font(10, bold=True),
                                    foreground=PALETTE["ink_soft"],
                                    background=PALETTE["paper_edge"])
            list_text.tag_configure("loc_section", font=_cn_font(10),
                                    foreground=PALETTE["ink_faint"])
            list_text.tag_configure("actual_marker", font=("Times New Roman", 10, "bold"),
                                    foreground=PALETTE["rose_deep"])
            list_text.tag_configure("actual_text", font=_serif_font(10),
                                    foreground=PALETTE["ink_soft"])
            list_text.tag_configure("expected_marker", font=("Times New Roman", 10, "bold"),
                                    foreground=PALETTE["sage"])
            list_text.tag_configure("expected_text", font=_serif_font(10),
                                    foreground=PALETTE["ink_soft"])
            list_text.tag_configure("fix_marker", font=_cn_font(10, bold=True),
                                    foreground=PALETTE["ochre"])
            list_text.tag_configure("fix_text", font=_serif_font(10, italic=True),
                                    foreground=PALETTE["ink_soft"])
            list_text.tag_configure("sep", foreground=PALETTE["rule"])

            sev_meta_v2 = {
                "高": ("HIGH", "sev_h"),
                "中": ("MID",  "sev_m"),
                "低": ("LOW",  "sev_l"),
            }
            sev_order = {"高": 0, "中": 1, "低": 2}
            sorted_issues = sorted(
                issues,
                key=lambda x: (sev_order.get(x.get("severity", "低"), 2),)
            )

            for idx, iss in enumerate(sorted_issues):
                sev_key = iss.get("severity", "低")
                sev_short, sev_tag = sev_meta_v2.get(sev_key, sev_meta_v2["低"])
                rule = iss.get("rule", "")
                location = iss.get("location", "")
                actual = iss.get("actual", "")
                expected = iss.get("expected", "")
                suggestion = iss.get("suggestion", "")
                loc_section = iss.get("section") or iss.get("chapter") or ""

                if idx > 0:
                    list_text.insert("end", "\n", "sep")
                    list_text.insert("end", "─" * 60 + "\n", "sep")
                    list_text.insert("end", "\n", "sep")

                # row 1: LOW § rule
                list_text.insert("end", f"  {sev_short}  ", sev_tag)
                list_text.insert("end", "§ ", "rule_marker")
                list_text.insert("end", f"{rule}\n", "rule_text")

                # row 2: location pill + section
                if location:
                    list_text.insert("end", f"  {location}  ", "loc_pill")
                    if loc_section:
                        list_text.insert("end", f"  「{loc_section}」", "loc_section")
                    list_text.insert("end", "\n")

                # row 3: actual ✗
                if actual:
                    list_text.insert("end", "✗ ", "actual_marker")
                    list_text.insert("end", f"{actual}\n", "actual_text")
                # row 4: expected ✓
                if expected and expected != actual:
                    list_text.insert("end", "✓ ", "expected_marker")
                    list_text.insert("end", f"{expected}\n", "expected_text")
                # row 5: 建议 · italic
                if suggestion:
                    list_text.insert("end", "建议 · ", "fix_marker")
                    list_text.insert("end", f"{suggestion}\n", "fix_text")

            list_text.configure(state="disabled")

    def _draw_grade_ring(self, canvas, size, score, color, letter, level_name, _score_unused):
        """画等级圆环：底层灰圆 + 上层彩色弧 + 中间字母 + 副文字
        字号/环宽按 size 缩放,小尺寸(80)也能清晰显示
        """
        canvas.delete("all")
        pad = max(4, size // 12)
        cx = cy = size // 2
        r = (size // 2) - pad
        # 环宽按 size 缩放：120→10, 80→7
        w = max(5, int(round(size * 0.085)))

        # 底层灰圆环
        canvas.create_arc(
            cx - r, cy - r, cx + r, cy + r,
            start=90, extent=-360, style="arc",
            outline=PALETTE["border_strong"], width=w,
        )
        # 上层彩色弧（顺时针，从顶部 90° 开始扫描）
        # 注意：tkinter extent 是逆时针为正；负值 = 顺时针
        extent = -score * 360 / 100
        if abs(extent) < 1:  # 0 时画不出来，强制小段
            extent = -2
        canvas.create_arc(
            cx - r, cy - r, cx + r, cy + r,
            start=90, extent=extent, style="arc",
            outline=color, width=w,
        )
        # 中间字母(按 size 缩放：120→28, 100→23)
        fs_letter = max(14, int(round(size * 0.23)))
        fs_sub = max(7, int(round(size * 0.085)))
        # 字母放中心,副文字在字母下方,二者间距 = fs_letter/2
        canvas.create_text(
            cx, cy - fs_letter * 0.15, text=letter,
            font=("Helvetica", fs_letter, "bold"),
            fill=color,
        )
        # 副文字（小字号）—— 与字母底部留 fs_letter/2 间距
        canvas.create_text(
            cx, cy + fs_letter * 0.7, text=f"{level_name} {score} 分",
            font=("Helvetica", fs_sub),
            fill=PALETTE["text2"],
        )

    def _draw_grade_circle(self, canvas, size, color, letter, level_name):
        """学术笔记 v2 等级圆：140px 主圆 + 1px 外halo(opacity 0.3 模拟) +
        56pt moss serif 字母 + 10pt label letter-spacing。
        分数(score)由外部 row 在圆下方单独显示(22pt ink serif)。
        """
        canvas.delete("all")
        cx = cy = size // 2
        # 主圆背景 = paper (内填白)
        canvas.create_oval(0, 0, size, size, fill=PALETTE["paper"], outline="")
        # 8px moss 实心环（按 size 比例）
        ring_w = max(5, int(round(size * 0.06)))
        canvas.create_oval(
            ring_w // 2, ring_w // 2,
            size - ring_w // 2 - 1, size - ring_w // 2 - 1,
            outline=color, width=ring_w,
        )
        # 1px 外 halo（比主圆大14px，用更浅色——moss 已带浅色,用 moss-tint 模拟）
        halo_extent = max(8, int(round(size * 0.1)))
        canvas.create_oval(
            -halo_extent, -halo_extent,
            size + halo_extent, size + halo_extent,
            outline=PALETTE["moss_tint"], width=1,
        )
        # 字母（按 size 比例缩放）
        fs_letter = max(36, int(round(size * 0.4)))
        canvas.create_text(
            cx, cy - max(4, int(round(size * 0.06))), text=letter,
            font=("Times New Roman", fs_letter, "bold"),
            fill=color,
        )
        # 副标签（按 size 比例缩放）
        spaced_label = "  ".join(level_name) if level_name else ""
        fs_label = max(8, int(round(size * 0.07)))
        canvas.create_text(
            cx, cy + max(18, int(round(size * 0.2))), text=spaced_label,
            font=_cn_font(fs_label),
            fill=PALETTE["ink_faint"],
        )

    def _draw_pill(self, parent, sev_key, count):
        """画一个严重度 pill：圆角胶囊 + 严重度标签 + 数字
        横向填满父容器(fill="x"),让 pill 长度一致且占满右侧空间
        """
        strong, soft, dot = self.SEVERITY_META[sev_key]
        # 外层 frame —— 横向撑满
        pill = tk.Frame(
            parent, bg=soft,
            highlightthickness=0, bd=0,
        )
        pill.pack(fill="x", pady=2, anchor="w")
        # Canvas 模拟圆角（用圆角矩形）
        # 这里简单用高度 + 足够水平 padx 来视觉上像胶囊
        lbl = tk.Label(
            pill, text=f"  {dot} {sev_key}严重度  {count}  ",
            font=_cn_font(10, bold=True), fg=strong, bg=soft,
            padx=14, pady=5,
        )
        lbl.pack(side="left", anchor="w")

    def _draw_sev_chip(self, parent, sev_key, count, idx):
        """左栏用的紧凑竖排 chip：学术笔记 v2 形态
        - paper-edge 底（macaron soft 替代）
        - 3px 左侧同色条
        - 11pt sans 标签 + 14pt bold 数字
        """
        # 严重度颜色映射（学术笔记 v2 · 主色 = rose/honey/sage）
        sev_colors = {
            "高": PALETTE["rose_deep"],
            "中": PALETTE["honey"],
            "低": PALETTE["sage"],
        }
        sev_color = sev_colors.get(sev_key, PALETTE["sage"])
        # 横向填满左栏
        chip = tk.Frame(
            parent, bg=PALETTE["paper_edge"],
            highlightthickness=0, bd=0,
        )
        chip.grid(row=idx, column=0, sticky="ew", pady=2)
        # 左 3px 色条
        tk.Frame(chip, bg=sev_color, width=3).pack(side="left", fill="y")
        # 严重度名
        inner = tk.Frame(chip, bg=PALETTE["paper_edge"])
        inner.pack(side="left", fill="both", expand=True, padx=10, pady=3)
        inner.columnconfigure(1, weight=1)
        tk.Label(
            inner, text=f"{sev_key} 严重度",
            font=_cn_font(10), fg=PALETTE["ink_soft"], bg=PALETTE["paper_edge"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        # 数字（13pt bold）
        tk.Label(
            inner, text=str(count),
            font=_cn_font(13, bold=True), fg=sev_color, bg=PALETTE["paper_edge"],
            anchor="e",
        ).grid(row=0, column=1, sticky="e", padx=(8, 0))

    def _draw_issue_row(self, parent, iss, idx):
        """画一条问题行：顶部一行（严重度pill + 规则名 + 位置标签）
        底部一整行：完整描述（位置/问题/建议），wrap 到多行
        """
        sev_key = iss.get("severity", "低")
        strong, soft, dot = self.SEVERITY_META.get(
            sev_key, self.SEVERITY_META["低"]
        )
        rule = iss.get("rule", "")
        location = iss.get("location", "")
        # 不截短 location —— 完整位置信息让用户能定位到文档中具体段落
        suggestion = iss.get("suggestion", iss.get("expected", ""))
        actual = iss.get("actual", "")
        expected = iss.get("expected", "")

        # 描述组合：位置 → 问题（actual） → 建议（suggestion）
        desc_parts = []
        if location:
            desc_parts.append(f"📍 位置：{location}")
        if actual:
            desc_parts.append(f"❌ 问题：{actual}")
        if expected and expected != actual:
            desc_parts.append(f"✅ 期望：{expected}")
        if suggestion:
            desc_parts.append(f"💡 建议：{suggestion}")
        desc_text = "\n".join(desc_parts) if desc_parts else rule

        # 行 frame（用 grid 分两行：top=top, desc=bottom）
        row = tk.Frame(
            parent, bg=PALETTE["card"] if idx % 2 == 0 else PALETTE["bg"],
            highlightthickness=0, bd=0,
        )
        row.pack(fill="x", pady=2, anchor="n")
        row.columnconfigure(0, weight=1)

        # 顶部行：严重度小 pill + 规则名（bold）+ 位置标签
        top = tk.Frame(row, bg=row.cget("bg"), highlightthickness=0, bd=0)
        top.grid(row=0, column=0, sticky="ew", padx=(8, 8), pady=(6, 0))
        top.columnconfigure(1, weight=1)

        # 严重度小 pill（最左）
        sev_pill = tk.Label(
            top, text=f" {sev_key} ", font=_cn_font(8, bold=True),
            fg=strong, bg=soft,
            padx=4, pady=2,
        )
        sev_pill.grid(row=0, column=0, padx=(0, 8), sticky="w")

        # 规则名（bold,占满中间）
        rule_lbl = tk.Label(
            top, text=rule, font=_cn_font(10, bold=True),
            fg=PALETTE["text"], bg=top.cget("bg"),
            anchor="w",
        )
        rule_lbl.grid(row=0, column=1, sticky="ew", padx=(0, 6))

        # 底部：完整描述（多行,占满整行宽度）
        desc_lbl = tk.Label(
            row, text=desc_text, font=_cn_font(9),
            fg=PALETTE["text2"], bg=row.cget("bg"),
            anchor="w", justify="left",
        )
        desc_lbl.grid(row=1, column=0, sticky="ew", padx=(8, 8), pady=(4, 8))

        # 父容器宽度变化时,精确计算 desc 实际可用宽度
        def _resize_desc(e, lbl=desc_lbl, parent=row):
            try:
                parent.update_idletasks()
                avail = max(200, parent.winfo_width() - 20)
                lbl.config(wraplength=avail)
            except Exception:
                pass
        row.bind("<Configure>", _resize_desc, add="+")

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

        # lead 引子（学术笔记 v2 · 与 Tab1 一致）：3px 左 ochre + 10.5pt ink-soft（紧凑）
        try:
            _lib = load_library()
            _total = len(_lib.get("topics", []))
        except Exception:
            _total = 0
        lead_row = tk.Frame(f, bg=PALETTE["bg"])
        lead_row.pack(anchor="w", padx=20, pady=(6, 2), fill="x")
        # 左：3px ochre 竖条
        tk.Frame(lead_row, bg=PALETTE["ochre"], width=3).pack(side="left", fill="y", padx=(0, 9))
        self.lib_lead_lbl = tk.Label(
            lead_row,
            text=f"浏览往届毕业论文选题库（共 {_total} 条 · 2020-2026 届），按关键词 / 方向 / 年份检索，查看每个选题的研究方向、关键词与同类选题。",
            font=_serif_font(10.5), fg=PALETTE["ink_soft"], bg=PALETTE["bg"],
            justify="left", anchor="w", padx=0, pady=0,
        )
        self.lib_lead_lbl.pack(side="left", fill="x", expand=True)
        # 跟随窗口缩放动态换行
        def _resize_lib_lead(event):
            self.lib_lead_lbl.config(wraplength=max(200, event.width - 40))
        f.bind("<Configure>", _resize_lib_lead, add="+")

        filter_frame = tk.Frame(f, bg=PALETTE["bg"])
        filter_frame.pack(fill="x", padx=20, pady=4)

        tk.Label(filter_frame, text="搜索:", font=_cn_font(10),
                 bg=PALETTE["bg"], fg=PALETTE["text2"]).pack(side="left")
        self.search_var = tk.StringVar()
        # 1) trace 触发（英文/数字键击）
        # 2) <Key> 触发（按键即触发，最稳）
        # 3) <KeyRelease> 补充（IME 上屏阶段）
        # 4) <<Commit>> IME 组字结束
        # 多重保险：单独输入关键词也能即时看到匹配结果。
        self.search_var.trace_add("write", lambda *_: self._schedule_refresh())
        search_entry = tk.Entry(
            filter_frame, textvariable=self.search_var,
            font=_cn_font(10), width=30, relief="flat", bd=0,
            bg=PALETTE["card"], fg=PALETTE["text"],
            highlightthickness=1, highlightbackground=PALETTE["border"],
        )
        search_entry.pack(side="left", padx=5, ipady=3)
        # 按下任何键都触发（最稳，包括 IME）
        search_entry.bind("<Key>", lambda e: self._schedule_refresh())
        search_entry.bind("<KeyRelease>", lambda e: self._schedule_refresh())
        search_entry.bind("<<Commit>>", lambda e: self._schedule_refresh())
        # 粘贴 / 剪切后内容变了但不一定触发键事件
        search_entry.bind("<<Paste>>", lambda e: self.root.after(50, self._refresh_library))
        search_entry.bind("<<Cut>>", lambda e: self.root.after(50, self._refresh_library))

        tk.Label(filter_frame, text="方向:", font=_cn_font(10),
                 bg=PALETTE["bg"], fg=PALETTE["text2"]).pack(side="left", padx=(20, 0))
        self.direction_var = tk.StringVar(value="全部")
        directions = ["全部", "翻译", "跨文化", "话语分析", "商务英语习得"]
        self.direction_combo = ttk.Combobox(filter_frame, textvariable=self.direction_var,
                     values=directions, state="readonly", width=12)
        self.direction_combo.pack(side="left", padx=5)
        # ttk.Combobox + state="readonly" + trace_add 在某些场景不触发，
        # 用 <<ComboboxSelected>> 虚拟事件更可靠
        self.direction_combo.bind("<<ComboboxSelected>>", lambda e: self._refresh_library())

        tk.Label(filter_frame, text="年份:", font=_cn_font(10),
                 bg=PALETTE["bg"], fg=PALETTE["text2"]).pack(side="left", padx=(20, 0))
        self.year_var = tk.StringVar(value="全部")
        years = ["全部", "2026", "2025", "2024", "2023", "2022", "2021", "2020", "2019"]
        self.year_combo = ttk.Combobox(filter_frame, textvariable=self.year_var,
                     values=years, state="readonly", width=8)
        self.year_combo.pack(side="left", padx=5)
        self.year_combo.bind("<<ComboboxSelected>>", lambda e: self._refresh_library())

        # 导入 / 导出按钮（同一行，导入靠左、导出靠右）
        export_row = tk.Frame(f, bg=PALETTE["bg"])
        export_row.pack(fill="x", padx=20, pady=(2, 0))

        make_button(export_row, "📥 导入选题",
                    self._on_import_topics,
                    style="primary", padx=10, pady=3, radius=6,
                    font=_cn_font(9, bold=True)).pack(side="left")
        make_button(export_row, "📊 导出 Excel",
                    self._on_export_excel,
                    style="secondary", padx=10, pady=3, radius=6,
                    font=_cn_font(9, bold=True)).pack(side="right", padx=(4, 0))
        make_button(export_row, "📄 导出 CSV",
                    self._on_export_csv,
                    style="secondary", padx=10, pady=3, radius=6,
                    font=_cn_font(9, bold=True)).pack(side="right", padx=4)

        # === 顶部统计条：5 块（方向 / 高低频理论 / 高低频关键词）===
        self._stats_strip = tk.Frame(
            f, bg=PALETTE["paper_edge"],
            highlightthickness=1, highlightbackground=PALETTE["rule"],
            bd=0,
        )
        self._stats_strip.pack(fill="x", padx=20, pady=(0, 4))
        for i in range(5):
            self._stats_strip.columnconfigure(i, weight=1, uniform="stat")

        paned = tk.PanedWindow(f, orient="horizontal", bg=PALETTE["bg"],
                                 sashwidth=4, sashrelief="flat",
                                 bd=0)
        paned.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        # 列表区（白卡 + 1px 边框）
        list_frame = tk.Frame(paned, bg=PALETTE["card"],
                              highlightthickness=1, highlightbackground=PALETTE["border"],
                              bd=0)
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
        tree.column("title", width=340, stretch=True)
        tree.column("direction", width=110, anchor="center")
        # 行字体固定 8pt（不随 DPI 缩放，避免 Retina 下变 20pt 过大）
        tree.tag_configure("row", font=_cn_font(8))

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        tree.bind("<<TreeviewSelect>>", self._on_topic_select)
        self.library_tree = tree

        # 详情区（白卡 + 1px 边框）—— 不做风险判定，仅信息呈现；可上下滚动
        detail_frame = tk.Frame(paned, bg=PALETTE["card"],
                                 highlightthickness=1, highlightbackground=PALETTE["border"],
                                 bd=0)
        paned.add(detail_frame, width=320)
        # 占位提示
        self._library_detail_placeholder = tk.Label(
            detail_frame, text="← 选中题目查看详情",
            font=_cn_font(10), fg=PALETTE["text3"], bg=PALETTE["card"],
        )
        self._library_detail_placeholder.place(relx=0.5, rely=0.5, anchor="center")
        # 详情滚动容器：tk.Text（wheel 在所有 Tk 平台上自动滚动）+ Scrollbar
        # 弃用 Canvas+Frame 子树（macOS TkAqua 下子 widget 的 wheel 经常失效）
        self._detail_text = tk.Text(
            detail_frame,
            bg=PALETTE["card"], fg=PALETTE["ink"],
            font=_serif_font(10),
            wrap="word", bd=0, highlightthickness=0,
            padx=12, pady=10, spacing1=2, spacing3=3,
            takefocus=0,
        )
        self._detail_text.place(x=0, y=0, relwidth=1, width=-14, relheight=1)
        self._detail_scroll = tk.Scrollbar(detail_frame, orient="vertical",
                                            command=self._detail_text.yview,
                                            width=14,
                                            troughcolor=PALETTE["paper"],
                                            bg=PALETTE["border"],
                                            activebackground=PALETTE["moss"],
                                            bd=0, highlightthickness=0,
                                            relief="flat")
        self._detail_scroll.place(x=0, y=0, relx=1, anchor="ne", relheight=1)
        self._detail_text.configure(yscrollcommand=self._detail_scroll.set)
        # 保留兼容引用（旧代码引用了这些名字）
        self._detail_canvas = self._detail_text  # 兼容 _bind_detail_scroll 等
        self.library_detail = self._detail_text

        # macOS 触控板惯性滚动 + 串扰：鼠标移走后 wheel 仍可能触发上一个 widget。
        # 用 Enter/Leave + Motion + 坐标检测三重保险维护 _in_detail flag。
        self._in_detail = False
        def _update_in_detail():
            """实时检测鼠标是否在 detail 区域内"""
            try:
                px, py = self.root.winfo_pointerxy()
                dx, dy = self._detail_text.winfo_rootx(), self._detail_text.winfo_rooty()
                dw, dh = self._detail_text.winfo_width(), self._detail_text.winfo_height()
                self._in_detail = bool(dx <= px <= dx + dw and dy <= py <= dy + dh)
            except Exception:
                pass

        def _on_detail_enter(event):
            self._in_detail = True
        def _on_detail_leave(event):
            self.root.after(50, _update_in_detail)  # 延迟 50ms 让 Enter 新 widget 先到
        def _on_detail_motion(event):
            _update_in_detail()
        def _on_detail_wheel(event):
            # flag=False 拦截：阻止 detail 在鼠标已移走后还滚
            if not self._in_detail:
                return "break"
        self._detail_text.bind("<Enter>", _on_detail_enter)
        self._detail_text.bind("<Leave>", _on_detail_leave)
        self._detail_text.bind("<Motion>", _on_detail_motion)
        self._detail_text.bind("<MouseWheel>", _on_detail_wheel, add="+")
        self._detail_text.bind("<Button-4>", _on_detail_wheel, add="+")
        self._detail_text.bind("<Button-5>", _on_detail_wheel, add="+")

        # Text tag 配置：颜色 + 字体（整体缩小一档，适配 222px 窄详情区）
        self._detail_text.tag_configure("chap", font=_cn_font(8, bold=True),
                                         foreground=PALETTE["ink_faint"], spacing1=6, spacing3=2)
        self._detail_text.tag_configure("id", font=_cn_font(10, italic=True),
                                         foreground=PALETTE["moss"])
        self._detail_text.tag_configure("title", font=_serif_font(12),
                                         foreground=PALETTE["ink"], spacing1=2, spacing3=6)
        self._detail_text.tag_configure("pill_dir", font=_cn_font(9, bold=True),
                                         foreground=PALETTE["moss"],
                                         background=PALETTE["moss_tint"],
                                         spacing1=2, spacing3=1)
        self._detail_text.tag_configure("pill_year", font=_cn_font(9, bold=True),
                                         foreground=PALETTE["ink_soft"],
                                         background=PALETTE["paper_edge"])
        self._detail_text.tag_configure("pill_theory", font=_cn_font(9, bold=True),
                                         foreground=PALETTE["ochre"],
                                         background=PALETTE["ochre_soft"])
        self._detail_text.tag_configure("block_title", font=_cn_font(8, bold=True),
                                         foreground=PALETTE["ink_faint"], spacing1=8, spacing3=2)
        self._detail_text.tag_configure("block_body", font=_cn_font(9),
                                         foreground=PALETTE["ink"], background=PALETTE["paper_edge"],
                                         spacing3=1, lmargin1=8, lmargin2=8)
        self._detail_text.tag_configure("sim_id", font=_cn_font(9, italic=True),
                                         foreground=PALETTE["ink_faint"])
        self._detail_text.tag_configure("sim_title", font=_cn_font(9),
                                         foreground=PALETTE["ink"])
        self._detail_text.tag_configure("ratio_label", font=_cn_font(9),
                                         foreground=PALETTE["ink_faint"])
        self._detail_text.tag_configure("ratio_pct", font=_cn_font(9, bold=True),
                                         foreground=PALETTE["ink_soft"])
        self._detail_text.tag_configure("sep", font=_cn_font(8),
                                         foreground=PALETTE["rule"], spacing1=6, spacing3=6)
        self._detail_text.tag_configure("keyword", font=_cn_font(10),
                                         foreground=PALETTE["moss"],
                                         background=PALETTE["moss_tint"],
                                         spacing1=2, spacing3=2)
        self._detail_text.config(state="disabled")

        self._refresh_library()
        self._render_stats_strip()

    def _schedule_refresh(self):
        """搜索/筛选变更后立即刷新（150ms debounce 已在测试中发现死循环风险，改为直接刷新）。

        Treeview 项数 ≤ 1000，刷新一次开销很小（< 50ms），无需 debounce。
        """
        self._refresh_library()

    # === KPI 卡相关 ===
    def _build_kpi_card(self, parent, label, value_var, accent, soft, sub_text="", sub_var=None):
        """构建一个 KPI 大数字卡（左侧 4px 强调色条 + 数字 + 副文字）。
        sub_var 给定时使用动态 StringVar；否则使用静态 sub_text。
        """
        card = tk.Frame(
            parent,
            bg=PALETTE["card"],
            highlightthickness=1,
            highlightbackground=PALETTE["border"],
            bd=0,
        )
        bar = tk.Canvas(card, width=4, height=64, bg=accent, highlightthickness=0, bd=0)
        bar.pack(side="left", fill="y")
        inner = tk.Frame(card, bg=PALETTE["card"])
        inner.pack(side="left", fill="both", expand=True, padx=(10, 8), pady=8)
        tk.Label(
            inner, text=label,
            font=_cn_font(9), fg=PALETTE["text3"], bg=PALETTE["card"],
        ).pack(anchor="w")
        tk.Label(
            inner, textvariable=value_var,
            font=_cn_font(20, bold=True), fg=accent, bg=PALETTE["card"],
        ).pack(anchor="w", pady=(2, 0))
        # 副文字（动态优先）
        if sub_var is not None:
            tk.Label(
                inner, textvariable=sub_var,
                font=_cn_font(9), fg=PALETTE["text2"], bg=PALETTE["card"],
            ).pack(anchor="w")
        elif sub_text:
            tk.Label(
                inner, text=sub_text,
                font=_cn_font(9), fg=PALETTE["text2"], bg=PALETTE["card"],
            ).pack(anchor="w")
        return card

    def _compute_kpi_cache(self):
        """从 load_library + get_library_insights 缓存 KPI 数据。"""
        try:
            lib = load_library()
            self._kpi_total = len(lib.get("topics", []))
        except Exception:
            self._kpi_total = 0
        try:
            from topic_analyzer import get_library_insights
            insights = get_library_insights()
            # 覆盖方向固定为「商务英语 4 大方向」(翻译/习得/跨文化/话语分析)
            # —— 即使库里出现其它方向(如教学法),也不计入,保持口径一致
            self._kpi_dirs = 4
            tops = insights.get("top_theories", [])
            if tops:
                self._kpi_top_theory = str(tops[0][0]) if tops[0] else "—"
                self._kpi_top_theory_count = int(tops[0][1]) if len(tops[0]) > 1 else 0
            else:
                self._kpi_top_theory = "—"
                self._kpi_top_theory_count = 0
        except Exception:
            self._kpi_dirs = 4
            self._kpi_top_theory = "—"
            self._kpi_top_theory_count = 0
        self._risk_high_count = 0  # 由 _refresh_library 实时更新

    def _refresh_kpi_display(self, risk_high_count=None):
        """更新 KPI 卡显示。risk_high_count 来自 _refresh_library 实时统计。"""
        if risk_high_count is None:
            risk_high_count = getattr(self, "_risk_high_count", 0)
        total = getattr(self, "_kpi_total", 0)
        dirs_n = getattr(self, "_kpi_dirs", 0)
        theory = getattr(self, "_kpi_top_theory", "—")
        theory_n = getattr(self, "_kpi_top_theory_count", 0)
        # 显示值
        self._kpi_total_var.set(str(total))
        self._kpi_dirs_var.set(str(dirs_n))
        self._kpi_top_theory_var.set(theory[:6] if theory else "—")  # 截短防溢出
        # 副文字（高频理论：XX 次）
        try:
            # 通过 _kpi_widgets 列表的第 3 个（高频理论）卡内的 inner 的副 label 替换
            # 为简化，副文字用统一接口
            self._kpi_top_theory_sub.set(f"{theory_n} 次")
        except Exception:
            pass
        self._kpi_risk_var.set(str(risk_high_count))

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
        # 关键词匹配：标题 + id + 方向 + 年份（多字段联合，提升搜索体验）
        matched = []
        for t in lib["topics"]:
            if direction != "全部" and t.get("direction") != direction:
                continue
            if year != "全部" and str(t.get("year", "")) != year:
                continue
            if keyword:
                haystack = " ".join([
                    str(t.get("title", "")).lower(),
                    str(t.get("id", "")).lower(),
                    str(t.get("direction", "")).lower(),
                    str(t.get("year", "")).lower(),
                ])
                if keyword not in haystack:
                    continue
            matched.append(t)
        self._current_topics = matched

        # 按届别倒序、标题字典序，2019 在末尾、2026 在最前
        matched.sort(key=lambda t: (-int(str(t.get('year', '0')) or 0), t['title']))

        # 高效更新：先 detach 所有 → 重新 insert → 避免逐个 delete
        for item in self.library_tree.get_children():
            self.library_tree.delete(item)

        # v3: 不做风险判定（用户提供往届信息用）
        # 仅保留 zebra striping
        try:
            self.library_tree.tag_configure("row_even", background=PALETTE["card"])
            self.library_tree.tag_configure("row_odd",  background=PALETTE["bg_soft"])
        except Exception:
            pass

        for i, t in enumerate(matched):
            title = t["title"]
            display_title = self._truncate_with_keyword(title, keyword, width=28)
            zebra = "row_even" if i % 2 == 0 else "row_odd"
            self.library_tree.insert(
                "", "end",
                values=(t["id"], display_title,
                        t.get("direction", "其他")),
                tags=("row", zebra),
            )

        self._risk_high_count = 0

        # 状态栏：明确反馈搜索结果（包含关键词命中位置提示）
        kw_msg = ""
        if keyword:
            kw_msg = f"，关键词「{keyword}」命中 {len(matched)} 条"
        self.status_var.set(
            f"选题库: {lib['metadata']['total']} 条 | "
            f"当前筛选 {len(matched)} 条{kw_msg}"
        )

    @staticmethod
    def _truncate_with_keyword(text: str, keyword: str, width: int = 38) -> str:
        """智能截断：若有关键词，保留关键词所在位置的上下文；否则截首部。

        例如 '基于目的论探讨跨境电商营销的翻译策略' + 关键词 '跨境电商' + width=20
        → '…探讨跨境电商营销的翻译策略…'
        """
        if len(text) <= width:
            return text
        if not keyword:
            return text[:width] + "…"

        k_lower = keyword.lower()
        t_lower = text.lower()
        pos = t_lower.find(k_lower)
        if pos < 0:
            return text[:width] + "…"
        # 让关键词尽量居中：左 padding = (width - len(keyword)) // 2
        kw_len = len(keyword)
        if kw_len >= width:
            return "…" + text[max(0, pos):][:width]
        left_pad = max(0, (width - kw_len) // 2)
        # 计算 start：让关键词在 [start, start+width] 内
        start = max(0, pos - left_pad)
        end = start + width
        if end > len(text):
            start = max(0, len(text) - width)
            end = len(text)
        snippet = text[start:end]
        prefix = "…" if start > 0 else ""
        suffix = "…" if end < len(text) else ""
        return prefix + snippet + suffix

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
                self._render_library_detail(t)
                break

    # ---------------- 详情渲染（学术笔记 v2） ----------------
    def _bind_detail_scroll(self, widget):
        """跨平台鼠标滚轮 + 触控板手势。

        Tk 的 wheel 事件不会沿 widget 层级冒泡，且 macOS (TkAqua) 上子 widget
        上的 binding 经常被吞掉。改用 **Enter/Leave flag 模式**：
        - 给 detail 子树里每个 widget 绑 <Enter>/<Leave> 维护 _in_detail flag
        - 给 root bind_all <MouseWheel>，回调里查 flag 而不是 event.widget
        """
        canvas = self._detail_canvas
        self._in_detail = False

        def _on_enter(event):
            self._in_detail = True
        def _on_leave(event):
            # 鼠标离开 detail 内任一子 widget
            try:
                px, py = self.root.winfo_pointerxy()
                dx, dy = canvas.winfo_rootx(), canvas.winfo_rooty()
                dw, dh = canvas.winfo_width(), canvas.winfo_height()
                if not (dx <= px <= dx + dw and dy <= py <= dy + dh):
                    self._in_detail = False
            except Exception:
                self._in_detail = False

        def _on_wheel(event):
            if not canvas.winfo_exists():
                return
            if not self._in_detail:
                return
            if event.num == 4:
                canvas.yview_scroll(-1, "units")
                return "break"
            if event.num == 5:
                canvas.yview_scroll(1, "units")
                return "break"
            if event.delta > 0:
                canvas.yview_scroll(-1, "units")
            elif event.delta < 0:
                canvas.yview_scroll(1, "units")
            return "break"

        # 递归绑 Enter/Leave + wheel 到子树
        def _bind_recursive(w):
            try:
                w.bind("<Enter>", _on_enter)
                w.bind("<Leave>", _on_leave)
                w.bind("<MouseWheel>", _on_wheel)
                w.bind("<Button-4>", _on_wheel)
                w.bind("<Button-5>", _on_wheel)
            except Exception:
                pass
            for child in w.winfo_children():
                _bind_recursive(child)

        _bind_recursive(widget)
        # bind_all 也绑上（覆盖 wheel 发到 root 的情况，作为兜底）
        self.root.bind_all("<MouseWheel>", _on_wheel)
        self.root.bind_all("<Button-4>", _on_wheel)
        self.root.bind_all("<Button-5>", _on_wheel)

    def _bind_detail_scroll_realtime(self):
        """每次重新渲染 detail 内容后调用，重新递归绑定 wheel 到新子 widgets。
        旧 widgets 已 destroy，无需 unbind。"""
        self._bind_detail_scroll(self.library_detail)

    def _render_library_detail(self, t):
        """右侧详情区 —— 用 tk.Text 渲染（wheel 在所有 Tk 平台自动滚动）"""
        text = self._detail_text
        text.config(state="normal")
        text.delete("1.0", "end")

        # 章节标题（左侧「选 中 详 情」+ 右侧选题 ID）
        text.insert("end", " ".join(list("选 中 详 情")), "chap")
        text.insert("end", "\t\t" + t["id"] + "\n", "id")
        # 题目
        text.insert("end", t["title"] + "\n", "title")

        # pills：direction / year / theory
        pills = [f" DIR · {t.get('direction', '其他')} "]
        year_str = str(t.get("year") or "")
        if year_str:
            pills.append(f" YEAR · {year_str} ")
        theory = t.get("theory")
        if theory and str(theory).strip():
            t_short = str(theory)
            if len(t_short) > 10:
                t_short = t_short[:10] + "…"
            pills.append(f" 理 论 · {t_short} ")
        text.insert("end", "  ".join(pills) + "\n", "pill_dir")

        # 关键词 block
        kws = t.get("keywords") or []
        if kws:
            text.insert("end", "\n" + " ".join(list("关 键 词")) + "\n", "block_title")
            kw_text = "  ".join(f" {kw} " for kw in kws[:7])
            text.insert("end", kw_text + "\n", "keyword")

        # 类似选题
        try:
            from topic_analyzer import extract_keywords
        except Exception:
            extract_keywords = None
        similar = []
        if extract_keywords:
            try:
                lib = load_library()
                t_kw = set(extract_keywords(t.get("title", "")))
                for other in lib.get("topics", []):
                    if other["id"] == t["id"]:
                        continue
                    if other.get("direction") == t.get("direction"):
                        o_kw = set(extract_keywords(other.get("title", "")))
                        if t_kw & o_kw:
                            similar.append(other)
                            if len(similar) >= 5:
                                break
            except Exception:
                similar = []
        if similar:
            text.insert("end", "\n" + " ".join(list(f"类 似 选 题 · 在库共 {len(similar)} 篇")) + "\n",
                        "block_title")
            for s in similar[:5]:
                st = s.get("title", "")
                # 不截断，让 Text 的 wrap="word" 自动换行显示完整标题
                text.insert("end", f"{s['id']}   ", "sim_id")
                text.insert("end", f"{st}\n", "sim_title")
            # 方向占比
            try:
                lib2 = load_library()
                total = len(lib2.get("topics", []))
                same = sum(1 for x in lib2.get("topics", [])
                           if x.get("direction") == t.get("direction"))
            except Exception:
                total, same = 0, 0
            ratio_pct = (same / total * 100) if total else 0
            bar_w = int((same / total) * 20) if total else 0
            bar = "█" * bar_w + "░" * (20 - bar_w)
            text.insert("end", "\n" + " ".join(list("方 向 占 比")) + "\n", "block_title")
            text.insert("end", f"{bar}  {same} / {total} · {ratio_pct:.1f}%\n", "block_body")

        text.config(state="disabled")
        # 滚到顶部
        text.yview_moveto(0)

    def _render_detail_block(self, parent, title, content_widget):
        """渲染 detail-block：paper-edge 底 + 3px 左 ochre 条 + 大写 letter-spacing 标题"""
        block = tk.Frame(parent, bg=PALETTE["paper_edge"], bd=0)
        block.pack(fill="x", anchor="w", pady=(10, 0))
        # 左侧 3px ochre 条
        tk.Frame(block, bg=PALETTE["ochre"], width=3).pack(side="left", fill="y")
        inner = tk.Frame(block, bg=PALETTE["paper_edge"])
        inner.pack(side="left", fill="both", expand=True, padx=10, pady=8)
        # title
        title_disp = " ".join(list(title))
        tk.Label(
            inner, text=title_disp,
            font=_cn_font(9, bold=True), fg=PALETTE["ink_faint"],
            bg=PALETTE["paper_edge"],
        ).pack(anchor="w", pady=(0, 4))
        # content
        content_widget.config(bg=PALETTE["paper_edge"])
        content_widget.pack(fill="x", anchor="w", in_=inner)

    def _build_kw_chips(self, kws):
        """构造关键词 chip row widget"""
        row = tk.Frame(bg=PALETTE["paper_edge"])
        for i, k in enumerate(kws):
            chip = tk.Frame(row, bg=PALETTE["moss_tint"])
            chip.pack(side="left", padx=(0, 5) if i > 0 else 0, pady=2)
            tk.Label(
                chip, text=f" {k} ",
                font=_cn_font(11), fg=PALETTE["moss"], bg=PALETTE["moss_tint"],
            ).pack()
        return row

    def _add_pill(self, parent, text, bg, fg):
        chip = tk.Frame(parent, bg=bg)
        chip.pack(side="left", padx=(0, 4), pady=2)
        tk.Label(
            chip, text=f" {text} ",
            font=_cn_font(9, bold=True), fg=fg, bg=bg,
            padx=2, pady=1,
        ).pack()

    # ---------------- 顶部统计条 ----------------
    def _render_stats_strip(self):
        """5 块：方向(4 类) / 高低频理论 TOP3 / 高低频关键词 TOP3"""
        for w in self._stats_strip.winfo_children():
            w.destroy()

        try:
            from topic_analyzer import get_library_insights
            insights = get_library_insights()
        except Exception:
            insights = {}

        direction_stats = insights.get("direction_stats", {})
        # 商务英语 4 大方向，按 count 降序取前 5
        # 显示用短标签（保持一致长度），key 用 data 里的全称
        direction_label_map = {
            "翻译": "翻译", "跨文化": "跨文化", "话语分析": "话语分析",
            "商务英语习得": "习得",
        }
        dir_rows = sorted(
            [(direction_label_map.get(d, d), direction_stats.get(d, 0))
             for d in direction_label_map],
            key=lambda x: -x[1],
        )[:5]
        top_theories = insights.get("top_theories", [])
        low_theories = list(reversed(top_theories[-5:])) if top_theories else []
        top_kws = insights.get("top_keywords", [])
        low_kws = list(reversed(top_kws[-5:])) if top_kws else []

        blocks = [
            ("方向·4类", dir_rows, PALETTE["moss"], "moss"),
            ("高频理论 TOP5", top_theories[:5],  PALETTE["moss"], "moss"),
            ("低频理论 TOP5", low_theories,      PALETTE["stone"], "stone"),
            ("高频关键词 TOP5", top_kws[:5],   PALETTE["ochre"], "ochre"),
            ("低频关键词 TOP5", low_kws,       PALETTE["stone"], "stone"),
        ]

        for col, (title, items, color, kind) in enumerate(blocks):
            blk = tk.Frame(self._stats_strip, bg=PALETTE["paper_edge"],
                           highlightthickness=1, highlightbackground=PALETTE["rule"],
                           bd=0)
            blk.grid(row=0, column=col, sticky="nsew", padx=1, pady=1)
            inner = tk.Frame(blk, bg=PALETTE["paper_edge"])
            inner.pack(fill="both", expand=True, padx=6, pady=2)
            # label：7pt sans weight 700
            tk.Label(
                inner, text=title,
                font=_cn_font(7, bold=True), fg=PALETTE["ink_faint"],
                bg=PALETTE["paper_edge"],
            ).pack(anchor="w", pady=(0, 1))
            # 每个 item 一行
            norm = [(x[0], x[1]) for x in items]
            max_v = max((v for _, v in norm), default=1) or 1
            for name, val in norm:
                row = tk.Frame(inner, bg=PALETTE["paper_edge"])
                row.pack(fill="x", pady=0)
                tk.Label(
                    row, text=str(name),
                    font=_cn_font(9), fg=PALETTE["ink"],
                    bg=PALETTE["paper_edge"], anchor="w",
                ).pack(side="left", padx=(0, 3))
                # bar（2px 高）
                bar_wrap = tk.Frame(row, bg=PALETTE["paper"], height=2,
                                    highlightthickness=1, highlightbackground=PALETTE["rule"],
                                    bd=0)
                bar_wrap.pack(side="left", fill="x", expand=True, padx=2)
                bar_fg = tk.Frame(bar_wrap, bg=color)
                rel_w = (val / max_v) if max_v else 0
                bar_fg.place(relx=0, rely=0, relheight=1,
                             relwidth=max(0.04, rel_w))
                tk.Label(
                    row, text=str(val),
                    font=_cn_font(8, bold=True), fg=color,
                    bg=PALETTE["paper_edge"],
                ).pack(side="right")

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

    # === 全局 ttk Style：Treeview 行高 + 马卡龙主题 ===
    style = ttk.Style()
    # ttk.Frame / ttk.LabelFrame 背景：与 PALETTE["bg"] 一致
    style.configure("TFrame", background=PALETTE["bg"])
    style.configure("TLabelframe", background=PALETTE["bg"])
    style.configure("TLabelframe.Label", background=PALETTE["bg"],
                     foreground=PALETTE["text"], font=_cn_font(10, bold=True))
    # Treeview 字号：固定 8pt，不随 DPI 缩放（保证 Retina 下不会变太大）
    # 行高 = 字号 × 2.2 + 6px 留白
    tv_font_size = 8
    row_h = int(round(tv_font_size * 2.2 + 6))
    style.configure("Treeview",
                     font=_cn_font(tv_font_size),
                     rowheight=row_h,
                     background=PALETTE["card"],
                     fieldbackground=PALETTE["card"],
                     foreground=PALETTE["text"])
    style.configure("Treeview.Heading",
                     font=_cn_font(tv_font_size, bold=True),
                     background=PALETTE["bg_soft"],
                     foreground=PALETTE["text"])
    # 选中行：薄荷绿软底 + 薄荷绿字
    style.map("Treeview",
               background=[("selected", PALETTE["mint_soft"])],
               foreground=[("selected", PALETTE["mint"])])
    # Notebook tab 完全隐藏（视觉上由自定义 TabSwitcher 接管）
    try:
        style.configure("TNotebook", tabmargins=(0, 0, 0, 0), background=PALETTE["bg"])
        style.layout("TNotebook.Tab", [])
    except Exception:
        pass
    # Combobox 主题：白底 + 薄荷绿聚焦边框
    style.configure("TCombobox",
                     fieldbackground=PALETTE["card"],
                     background=PALETTE["card"],
                     foreground=PALETTE["text"],
                     arrowcolor=PALETTE["text2"])
    style.map("TCombobox",
               fieldbackground=[("readonly", PALETTE["card"])],
               selectbackground=[("readonly", PALETTE["card"])],
               selectforeground=[("readonly", PALETTE["text"])])

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
