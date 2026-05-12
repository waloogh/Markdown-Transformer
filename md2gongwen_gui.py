#!/usr/bin/env python3
"""
md2gongwen_gui.py — Markdown → 标准公文格式 DOCX 图形界面
零依赖：Python 3.8+ 标准库（Tkinter + md2gongwen 模块）

用法:
    python3 md2gongwen_gui.py
    pyinstaller --onefile --windowed md2gongwen_gui.py
"""

import os
import sys
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path

# 导入核心转换模块（同目录）
import md2gongwen as gw

# ── 常量 ──────────────────────────────────────────────────────────
APP_TITLE = "Markdown → 公文格式转换器"
TEMPLATE_PATH = Path.home() / ".gongwen_templates.json"
CONFIG_PATH = Path.home() / ".gongwen_gui_config.json"
DEFAULT_TEMPLATES = {
    "标准通知": "# 关于开展XXX工作的通知\n\n> 各相关部门：\n\n正文内容...\n\n## 一、背景\n\n## 二、工作要求\n\n---\n发文机关署名\n2026年1月1日\n",
    "请示报告": "# 关于XXX的请示\n\n> XXX部门：\n\n正文内容...\n\n## 一、基本情况\n\n## 二、请示事项\n\n---\n申请单位署名\n2026年1月1日\n",
    "会议纪要": "# XXX会议纪要\n\n## 一、会议时间地点\n\n## 二、参会人员\n\n## 三、会议内容\n\n## 四、议定事项\n\n---\n记录人署名\n2026年1月1日\n",
}

ROLES = ["标题", "主送单位", "正文", "一级标题", "二级标题", "三级标题", "附件", "落款"]

# 字号映射：中文名 → pt
SIZE_NAMES = {
    "初号": 42, "小初": 36,
    "一号": 26, "小一": 24,
    "二号": 22, "小二": 18,
    "三号": 16, "小三": 15,
    "四号": 14, "小四": 12,
    "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5,
    "七号": 5.5, "八号": 5,
}
SIZE_PT_TO_NAME = {v: k for k, v in SIZE_NAMES.items()}
DEFAULT_SIZES = {"标题": 22, "主送单位": 16, "正文": 16, "一级标题": 16,
                  "二级标题": 16, "三级标题": 16, "附件": 16, "落款": 16}

ROLE_DEFAULTS = {
    "标题":     {"font": "方正小标宋简体", "size": 22, "bold": False, "center": True,  "indent": 0},
    "主送单位": {"font": "仿宋_GB2312",    "size": 16, "bold": False, "center": False, "indent": 0},
    "正文":     {"font": "仿宋_GB2312",    "size": 16, "bold": False, "center": False, "indent": 2},
    "一级标题": {"font": "黑体",           "size": 16, "bold": False, "center": False, "indent": 2},
    "二级标题": {"font": "楷体_GB2312",    "size": 16, "bold": False, "center": False, "indent": 2},
    "三级标题": {"font": "仿宋_GB2312",    "size": 16, "bold": True,  "center": False, "indent": 2},
    "附件":     {"font": "仿宋_GB2312",    "size": 16, "bold": False, "center": False, "indent": 4},
    "落款":     {"font": "仿宋_GB2312",    "size": 16, "bold": False, "center": False, "indent": 0},
}


# ═══════════════════════════════════════════════════════════════════
# 配置持久化
# ═══════════════════════════════════════════════════════════════════

def _load_config():
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text("utf-8"))
        except Exception:
            pass
    return {}


def _save_config(cfg):
    CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), "utf-8")


def _load_user_templates():
    """加载用户自定义模板（不含默认）"""
    if TEMPLATE_PATH.exists():
        try:
            return json.loads(TEMPLATE_PATH.read_text("utf-8"))
        except Exception:
            pass
    return {}


def _get_all_templates():
    """合并默认 + 用户模板（用户覆盖同名默认）"""
    templates = dict(DEFAULT_TEMPLATES)
    templates.update(_load_user_templates())
    return templates


def _save_user_template(name, content):
    """保存单个用户模板"""
    user = _load_user_templates()
    user[name] = content
    TEMPLATE_PATH.write_text(json.dumps(user, ensure_ascii=False, indent=2), "utf-8")


def _delete_user_template(name):
    """删除用户模板覆盖（回退到默认）"""
    user = _load_user_templates()
    if name in user:
        del user[name]
        TEMPLATE_PATH.write_text(json.dumps(user, ensure_ascii=False, indent=2), "utf-8")
        return True
    return False


# ═══════════════════════════════════════════════════════════════════
# 首次启动引导
# ═══════════════════════════════════════════════════════════════════

class FirstRunWizard(tk.Toplevel):
    """首次启动引导窗口"""

    def __init__(self, parent, config, on_done):
        super().__init__(parent)
        self.config = config
        self.on_done = on_done
        self.title("首次配置 — 公文格式转换器")
        self.geometry("550x480")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self._build()
        self.protocol("WM_DELETE_WINDOW", self._finish)

    def _build(self):
        main = ttk.Frame(self, padding=20)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text="欢迎使用公文格式转换器",
                  font=("", 16, "bold")).pack(pady=(0, 5))
        ttk.Label(main, text="首次使用，请完成以下配置：").pack(pady=(0, 15))

        # ── 输出目录 ──
        out_frame = ttk.LabelFrame(main, text="1. 默认输出目录", padding=10)
        out_frame.pack(fill="x", pady=(0, 10))
        self.out_dir_var = tk.StringVar(
            value=self.config.get("output_dir", str(Path.home() / "Documents"))
        )
        ttk.Entry(out_frame, textvariable=self.out_dir_var, width=40).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(out_frame, text="浏览...",
                   command=self._browse_dir).pack(side="left", padx=(5, 0))

        # ── 字体扫描 ──
        font_frame = ttk.LabelFrame(main, text="2. 字体扫描", padding=10)
        font_frame.pack(fill="x", pady=(0, 10))

        self.font_text = scrolledtext.ScrolledText(
            font_frame, height=8, width=50,
            font=("Menlo", 10), state="disabled"
        )
        self.font_text.pack(fill="x")
        ttk.Button(font_frame, text="重新扫描",
                   command=self._scan_fonts).pack(pady=(5, 0))

        # ── 按钮 ──
        ttk.Button(main, text="✔ 完成配置，进入编辑器",
                   command=self._finish).pack(pady=(10, 0))

        # 初始扫描
        self._scan_fonts()

    def _browse_dir(self):
        d = filedialog.askdirectory(
            title="选择默认输出目录",
            initialdir=self.out_dir_var.get()
        )
        if d:
            self.out_dir_var.set(d)

    def _scan_fonts(self):
        self.font_text.configure(state="normal")
        self.font_text.delete(1.0, tk.END)

        self.font_text.insert(tk.END, "正在扫描系统字体...\n\n")
        self.font_text.update()
        try:
            fonts = gw.scan_fonts()
            if fonts:
                rec = gw.fuzzy_match_fonts(fonts)
                for role in ROLES:
                    matched = rec.get(role, "未匹配")
                    status = "✅" if matched != "请手动指定" else "⚠️ "
                    self.font_text.insert(tk.END, f"  {status} {role:　<5} → {matched}\n")
                self.font_text.insert(tk.END, f"\n共发现 {len(fonts)} 个中文字体")
            else:
                self.font_text.insert(tk.END, "❌ 未检测到系统字体")
        except Exception as e:
            self.font_text.insert(tk.END, f"❌ 扫描失败: {e}")

        self.font_text.configure(state="disabled")

    def _finish(self):
        self.config["output_dir"] = self.out_dir_var.get()
        self.config["first_run_done"] = True
        _save_config(self.config)
        self.destroy()
        if self.on_done:
            self.on_done()


# ═══════════════════════════════════════════════════════════════════
# 字体配置对话框
# ═══════════════════════════════════════════════════════════════════

class FontConfigDialog(tk.Toplevel):
    """字体/格式配置对话框"""

    def __init__(self, parent, role_config, page_config, on_save, sys_fonts=None):
        super().__init__(parent)
        self.role_config = dict(role_config)
        self.page_config = dict(page_config)
        self.on_save = on_save
        self.sys_fonts = sys_fonts or []
        self.title("格式配置")
        self.geometry("520x540")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self._build()

    def _build(self):
        main = ttk.Frame(self, padding=15)
        main.pack(fill="both", expand=True)

        # ── 各元素配置 ──
        nb = ttk.Notebook(main)
        nb.pack(fill="x", pady=(0, 10))

        role_labels = {
            "标题": "方正小标宋简体 二号 居中",
            "主送单位": "仿宋_GB2312 三号 顶格",
            "正文": "仿宋_GB2312 三号 首行缩进2字",
            "一级标题": "黑体 三号 缩进2字",
            "二级标题": "楷体_GB2312 三号 缩进2字",
            "三级标题": "仿宋_GB2312 三号 加粗 缩进2字",
            "附件": "仿宋_GB2312 三号 顶格 缩进4字",
            "落款": "仿宋_GB2312 三号 右对齐",
        }

        self.role_vars = {}
        for role in ROLES:
            tab = ttk.Frame(nb, padding=10)
            nb.add(tab, text=f"  {role}  ")

            cfg = self.role_config.setdefault(role, dict(ROLE_DEFAULTS[role]))
            vars_dict = {}

            # 字体（下拉列表 + 可手动输入）
            ttk.Label(tab, text="字体:").grid(row=0, column=0, sticky="w", pady=2)
            v_font = tk.StringVar(value=cfg.get("font", ""))
            font_combo = ttk.Combobox(
                tab, textvariable=v_font, width=18,
                values=self.sys_fonts[:200] if self.sys_fonts else [],
                state="normal"  # 可编辑
            )
            font_combo.grid(row=0, column=1, sticky="w", pady=2)
            # 滚动时筛选
            font_combo.bind("<KeyRelease>", lambda e, cb=font_combo, fl=self.sys_fonts:
                cb.configure(values=[f for f in fl if e.widget.get().lower() in f.lower()][:100]))
            vars_dict["font"] = v_font

            # 字号
            ttk.Label(tab, text="字号 (pt):").grid(row=1, column=0, sticky="w", pady=2)
            v_size = tk.IntVar(value=cfg.get("size", 16))
            ttk.Spinbox(tab, from_=8, to=72, textvariable=v_size, width=6).grid(
                row=1, column=1, sticky="w", pady=2
            )
            vars_dict["size"] = v_size

            # 加粗
            v_bold = tk.BooleanVar(value=cfg.get("bold", False))
            ttk.Checkbutton(tab, text="加粗", variable=v_bold).grid(
                row=2, column=0, columnspan=2, sticky="w", pady=2
            )
            vars_dict["bold"] = v_bold

            # 居中
            v_center = tk.BooleanVar(value=cfg.get("center", False))
            ttk.Checkbutton(tab, text="居中", variable=v_center).grid(
                row=3, column=0, columnspan=2, sticky="w", pady=2
            )
            vars_dict["center"] = v_center

            # 首行缩进
            v_indent = tk.BooleanVar(value=cfg.get("indent", True))
            ttk.Checkbutton(tab, text="首行缩进 2 字符", variable=v_indent).grid(
                row=4, column=0, columnspan=2, sticky="w", pady=2
            )
            vars_dict["indent"] = v_indent

            # 描述
            ttk.Label(tab, text=role_labels.get(role, ""),
                      foreground="gray").grid(
                row=5, column=0, columnspan=2, sticky="w", pady=(8, 0)
            )

            self.role_vars[role] = vars_dict

        # ── 页面设置 ──
        page_frame = ttk.LabelFrame(main, text="页面设置 (mm)", padding=10)
        page_frame.pack(fill="x", pady=(0, 10))

        self.page_vars = {}
        margins = [
            ("上边距", "margin_top", 37),
            ("下边距", "margin_bottom", 35),
            ("左边距", "margin_left", 28),
            ("右边距", "margin_right", 26),
        ]
        for i, (label, key, default) in enumerate(margins):
            ttk.Label(page_frame, text=label).grid(
                row=0, column=i * 2, sticky="w", padx=(0, 5)
            )
            v = tk.IntVar(value=self.page_config.get(key, default))
            ttk.Spinbox(page_frame, from_=0, to=100, textvariable=v, width=5).grid(
                row=0, column=i * 2 + 1, sticky="w", padx=(0, 10)
            )
            self.page_vars[key] = v

        # 行距
        ttk.Label(page_frame, text="行距 (pt):").grid(row=1, column=0, sticky="w", pady=(8, 0))
        v_line_spacing = tk.IntVar(value=self.page_config.get("line_spacing", 28))
        ttk.Spinbox(page_frame, from_=10, to=60, textvariable=v_line_spacing, width=5).grid(
            row=1, column=1, sticky="w", pady=(8, 0)
        )
        self.page_vars["line_spacing"] = v_line_spacing

        # ── 按钮 ──
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_frame, text="恢复默认",
                   command=self._reset).pack(side="left")
        ttk.Button(btn_frame, text="取消",
                   command=self.destroy).pack(side="right", padx=(5, 0))
        ttk.Button(btn_frame, text="保存",
                   command=self._save).pack(side="right")

    def _reset(self):
        if messagebox.askyesno("确认", "恢复所有格式设置为默认值？"):
            self.role_config = {r: dict(ROLE_DEFAULTS[r]) for r in ROLES}
            self.destroy()
            d = FontConfigDialog(self.master, self.role_config, self.page_config, self.on_save)

    def _save(self):
        for role, vars_dict in self.role_vars.items():
            self.role_config[role] = {
                "font": vars_dict["font"].get(),
                "size": vars_dict["size"].get(),
                "bold": vars_dict["bold"].get(),
                "center": vars_dict["center"].get(),
                "indent": vars_dict["indent"].get(),
            }
        for key, var in self.page_vars.items():
            self.page_config[key] = var.get()
        self.destroy()
        if self.on_save:
            self.on_save(self.role_config, self.page_config)


# ═══════════════════════════════════════════════════════════════════
# 主窗口
# ═══════════════════════════════════════════════════════════════════

class GongwenApp:
    """公文格式转换器 GUI 主窗口"""

    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1200x750")

        # ── 状态 ──
        self.config = _load_config()
        self.templates = _get_all_templates()
        self.is_custom_template = {n: (n in _load_user_templates()) for n in DEFAULT_TEMPLATES}
        self.current_md_path = None
        self.current_template_name = None
        self.role_config = {r: dict(ROLE_DEFAULTS[r]) for r in ROLES}
        self.page_config = {
            "margin_top": gw.MARGIN_TOP,
            "margin_bottom": gw.MARGIN_BOTTOM,
            "margin_left": gw.MARGIN_LEFT,
            "margin_right": gw.MARGIN_RIGHT,
            "line_spacing": gw.LINE_SPACING_PT,
        }
        # 加载已保存的格式配置
        saved_roles = self.config.get("role_config")
        if saved_roles:
            self.role_config.update(saved_roles)
        saved_page = self.config.get("page_config")
        if saved_page:
            self.page_config.update(saved_page)

        # ── 构建 UI ──
        self.status_var = tk.StringVar(value="就绪")
        self._build_menu()
        self._build_body()
        self._build_status()

        # ── 绑定快捷键 ──
        self.root.bind("<Command-o>", lambda e: self._open_file())
        self.root.bind("<Command-s>", lambda e: self._save_md())
        self.root.bind("<Command-b>", lambda e: self._convert())
        self.root.bind("<Command-1>", lambda e: self._load_template_by_key(1))
        self.root.bind("<Command-2>", lambda e: self._load_template_by_key(2))
        self.root.bind("<Command-3>", lambda e: self._load_template_by_key(3))

        # ── 首次启动 ──
        if not self.config.get("first_run_done"):
            self.root.after(300, self._show_first_run)

    # ── 菜单栏 ──────────────────────────────────────────────────

    def _build_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # 文件
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="打开 .md 文件...", command=self._open_file,
                              accelerator="Cmd+O")
        file_menu.add_command(label="保存 .md", command=self._save_md,
                              accelerator="Cmd+S")
        file_menu.add_command(label="另存为...", command=self._save_md_as)
        file_menu.add_separator()
        file_menu.add_command(label="转换并导出 .docx", command=self._convert,
                              accelerator="Cmd+B")
        file_menu.add_command(label="打开输出文件...", command=self._open_output)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit,
                              accelerator="Cmd+Q")
        menubar.add_cascade(label="文件", menu=file_menu)

        # 模板
        self.template_menu = tk.Menu(menubar, tearoff=0)
        self._rebuild_template_menu()
        menubar.add_cascade(label="模板", menu=self.template_menu)

        # 格式
        format_menu = tk.Menu(menubar, tearoff=0)
        format_menu.add_command(label="字体 & 页面设置...",
                                command=self._open_font_config)
        menubar.add_cascade(label="格式", menu=format_menu)

        # 帮助
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="关于", command=self._about)
        menubar.add_cascade(label="帮助", menu=help_menu)

    def _rebuild_template_menu(self):
        """重建模板菜单"""
        self.template_menu.delete(0, tk.END)
        for i, name in enumerate(self.templates):
            key = f"Cmd+{i + 1}" if i < 9 else ""
            suffix = " (自定义)" if name in _load_user_templates() else ""
            label = f"{name}{suffix}"
            self.template_menu.add_command(
                label=label,
                command=lambda n=name: self._load_template(n),
                accelerator=key,
            )
        # 恢复默认子菜单（仅对已修改的默认模板）
        custom_names = [n for n in DEFAULT_TEMPLATES if n in _load_user_templates()]
        if custom_names:
            self.template_menu.add_separator()
            restore_menu = tk.Menu(self.template_menu, tearoff=0)
            for n in custom_names:
                restore_menu.add_command(
                    label=f"{n} → 恢复默认",
                    command=lambda name=n: self._reset_template(name)
                )
            self.template_menu.add_cascade(label="🔄 恢复默认模板", menu=restore_menu)
        self.template_menu.add_separator()
        self.template_menu.add_command(label="💾 保存当前为模板...",
                                       command=self._save_as_template)
    def _build_body(self):
        """左右分割面板"""
        self.paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.paned.pack(fill="both", expand=True, padx=5, pady=5)

        # 左侧编辑区
        self.editor_frame = ttk.Frame(self.paned)
        self._build_editor(self.editor_frame)
        self.paned.add(self.editor_frame, weight=6)

        # 右侧功能区
        self.output_frame = ttk.Frame(self.paned)
        self._build_output(self.output_frame)
        self.paned.add(self.output_frame, weight=4)

        # 加载默认模板（两个面板都就绪后）
        if self.templates:
            self._load_template(list(self.templates.keys())[0], silent=True)

    def _build_editor(self, parent):
        """左侧 Markdown 编辑区"""
        # 工具栏
        toolbar = ttk.Frame(parent)
        toolbar.pack(fill="x", pady=(0, 3))

        ttk.Button(toolbar, text="📂 打开", command=self._open_file).pack(side="left", padx=2)
        ttk.Button(toolbar, text="💾 保存", command=self._save_md).pack(side="left", padx=2)
        ttk.Button(toolbar, text="🔁 转换", command=self._convert).pack(side="left", padx=2)

        # 模板快速加载
        ttk.Label(toolbar, text="模板:").pack(side="left", padx=(10, 2))
        self.template_combo = ttk.Combobox(
            toolbar, values=list(self.templates.keys()),
            state="readonly", width=12
        )
        self.template_combo.pack(side="left", padx=2)
        self.template_combo.bind("<<ComboboxSelected>>",
                                 lambda e: self._on_template_combo())
        if self.templates:
            self.template_combo.set(list(self.templates.keys())[0])

        # 编辑区
        self.editor = scrolledtext.ScrolledText(
            parent, wrap="word", undo=True,
            font=("Menlo", 13), padx=10, pady=10,
            bg="#fafafa", insertbackground="black"
        )
        self.editor.pack(fill="both", expand=True)
        self.editor.bind("<KeyRelease>", lambda e: self._update_status())

    def _build_output(self, parent):
        """右侧预览/输出区"""
        # 输出设置
        out_frame = ttk.LabelFrame(parent, text="输出设置", padding=8)
        out_frame.pack(fill="x", pady=(0, 5))

        ttk.Label(out_frame, text="输出目录:").pack(anchor="w")
        dir_frame = ttk.Frame(out_frame)
        dir_frame.pack(fill="x")
        self.out_dir_var = tk.StringVar(
            value=self.config.get("output_dir", str(Path.home() / "Documents"))
        )
        ttk.Entry(dir_frame, textvariable=self.out_dir_var).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(dir_frame, text="浏览",
                   command=self._browse_out_dir).pack(side="left", padx=(5, 0))

        ttk.Label(out_frame, text="输出文件名:").pack(anchor="w", pady=(5, 0))
        name_frame = ttk.Frame(out_frame)
        name_frame.pack(fill="x")
        self.out_name_var = tk.StringVar(value="output.docx")
        ttk.Entry(name_frame, textvariable=self.out_name_var).pack(
            side="left", fill="x", expand=True
        )
        ttk.Label(name_frame, text=".docx").pack(side="left", padx=2)

        # 转换按钮
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill="x", pady=(8, 0))
        ttk.Button(btn_frame, text="🔁 转换为 .docx", command=self._convert,
                   ).pack(fill="x", ipady=5)

        # 模板管理
        tmpl_frame = ttk.LabelFrame(parent, text="模板管理", padding=8)
        tmpl_frame.pack(fill="x", pady=(8, 0))

        cur_tmpl = ttk.Frame(tmpl_frame)
        cur_tmpl.pack(fill="x")
        ttk.Label(cur_tmpl, text="当前:").pack(side="left")
        self.cur_template_label = ttk.Label(cur_tmpl, text="标准通知",
                                            foreground="#007AFF")
        self.cur_template_label.pack(side="left", padx=(5, 0))

        btn_tmpl = ttk.Frame(tmpl_frame)
        btn_tmpl.pack(fill="x", pady=(5, 0))
        ttk.Button(btn_tmpl, text="➕ 新建模板",
                   command=self._new_template_dialog).pack(side="left", padx=2)
        ttk.Button(btn_tmpl, text="✏️ 编辑模板",
                   command=self._edit_current_template).pack(side="left", padx=2)
        ttk.Button(btn_tmpl, text="🗑 删除模板",
                   command=self._delete_template_dialog).pack(side="left", padx=2)
        ttk.Button(btn_tmpl, text="🔄 重置",
                   command=self._reset_template_dialog).pack(side="left", padx=2)

        # 格式速览
        fmt_frame = ttk.LabelFrame(parent, text="格式速览", padding=8)
        fmt_frame.pack(fill="x", pady=(8, 0))

        # 表头
        headers = ["角色", "字体", "字号", "加粗", "居中", "缩进"]
        col_widths = [7, 14, 5, 4, 4, 4]
        for j, (h, w) in enumerate(zip(headers, col_widths)):
            ttk.Label(fmt_frame, text=h, font=("", 9, "bold"),
                      width=w, anchor="center").grid(row=0, column=j, padx=1, pady=1)

        # 格式行（动态更新）
        self.fmt_labels = {}
        for i, role in enumerate(ROLES):
            row = i + 1
            cfg = self.role_config.get(role, ROLE_DEFAULTS[role])
            indent_val = cfg.get("indent", 0)
            indent_text = f"缩{indent_val}字" if indent_val else ""
            vals = [
                role,
                cfg["font"][:12],
                self._size_name(cfg["size"]),
                "✓" if cfg["bold"] else "",
                "✓" if cfg["center"] else "",
                indent_text,
            ]
            row_labels = {}
            for j, v in enumerate(vals):
                lbl = ttk.Label(fmt_frame, text=v, width=col_widths[j],
                                anchor="center" if j > 0 else "w",
                                foreground="#333")
                lbl.grid(row=row, column=j, padx=1, pady=1)
                row_labels[j] = lbl
            self.fmt_labels[role] = row_labels

        # 调整按钮
        ttk.Button(fmt_frame, text="⚙️ 调整格式...",
                   command=self._open_font_config).grid(
            row=len(ROLES) + 1, column=0, columnspan=6, pady=(5, 0), sticky="e"
        )

        # 输出结果
        result_frame = ttk.LabelFrame(parent, text="转换结果", padding=8)
        result_frame.pack(fill="both", expand=True, pady=(8, 0))

        self.result_text = scrolledtext.ScrolledText(
            result_frame, height=8, state="disabled",
            font=("Menlo", 11), padx=5, pady=5
        )
        self.result_text.pack(fill="both", expand=True)

        # 打开按钮
        self.open_btn = ttk.Button(
            result_frame, text="📂 在 Finder 中打开",
            command=self._open_output, state="disabled"
        )
        self.open_btn.pack(pady=(5, 0))

    # ── 状态栏 ──────────────────────────────────────────────────

    def _build_status(self):
        self.status_var = tk.StringVar(value="就绪")
        self.status_bar = ttk.Label(
            self.root, textvariable=self.status_var,
            relief="sunken", anchor="w", padding=(8, 2)
        )
        self.status_bar.pack(fill="x", side="bottom")

    def _update_status(self):
        text = self.editor.get(1.0, "end-1c")
        chars = len(text)
        lines = text.count("\n") + 1
        self.status_var.set(
            f"字符: {chars} | 行: {lines}    "
            f"{self.current_md_path or '未保存'}"
        )

    # ── 文件操作 ────────────────────────────────────────────────

    def _open_file(self, path=None):
        if not path:
            path = filedialog.askopenfilename(
                title="打开文件（支持任意格式，内容按纯文本读取）",
                filetypes=[
                    ("所有文件", "*.*"),
                    ("Markdown", "*.md *.markdown"),
                    ("文本文件", "*.txt"),
                ]
            )
        if path and os.path.isfile(path):
            try:
                content = Path(path).read_text("utf-8")
                self.editor.delete(1.0, tk.END)
                self.editor.insert(1.0, content)
                self.current_md_path = path
                self._update_status()
                # 自动识别输出文件名
                stem = Path(path).stem
                self.out_name_var.set(f"{stem}.docx")
            except Exception as e:
                messagebox.showerror("错误", f"无法读取文件:\n{e}")

    def _save_md(self):
        if self.current_md_path:
            self._write_md(self.current_md_path)
        else:
            self._save_md_as()

    def _save_md_as(self):
        path = filedialog.asksaveasfilename(
            title="保存 Markdown 文件",
            defaultextension=".md",
            filetypes=[
                ("Markdown files", "*.md"),
                ("Text files", "*.txt"),
            ]
        )
        if path:
            self._write_md(path)

    def _write_md(self, path):
        try:
            text = self.editor.get(1.0, "end-1c")
            Path(path).write_text(text, "utf-8")
            self.current_md_path = path
            self._update_status()
            self._log_result(f"✅ 已保存: {path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败:\n{e}")

    def _browse_out_dir(self):
        d = filedialog.askdirectory(initialdir=self.out_dir_var.get())
        if d:
            self.out_dir_var.set(d)
            self.config["output_dir"] = d
            _save_config(self.config)

    # ── 模板 ─────────────────────────────────────────────────────

    def _on_template_combo(self):
        """模板下拉选择事件（仅响应用户手动选择）"""
        if getattr(self, '_suppress_combo', False):
            self._suppress_combo = False
            return
        name = self.template_combo.get()
        if name and name != self.current_template_name:
            self._load_template(name)

    def _load_template(self, name, silent=False):
        if name in self.templates:
            self.current_template_name = name
            self.editor.delete(1.0, tk.END)
            self.editor.insert(1.0, self.templates[name])
            self._suppress_combo = True
            self.template_combo.set(name)
            self.cur_template_label.configure(text=name)
            self._update_status()
            if not silent:
                self._log_result(f"📋 加载模板: {name}")

    def _refresh_templates(self):
        """刷新模板列表（菜单 + 下拉 + 标签）"""
        self.templates = _get_all_templates()
        self._rebuild_template_menu()
        self.template_combo["values"] = list(self.templates.keys())
        if self.current_template_name in self.templates:
            self._suppress_combo = True
            self.template_combo.set(self.current_template_name)
            self.cur_template_label.configure(text=self.current_template_name)

    def _new_template_dialog(self):
        """新建空白模板"""
        name = tk.simpledialog.askstring(
            "新建模板", "输入模板名称:",
            parent=self.root
        )
        if name and name.strip():
            name = name.strip()
            self.current_template_name = name
            self.editor.delete(1.0, tk.END)
            self._suppress_combo = True
            self.template_combo.set("")
            self.cur_template_label.configure(text=f"{name} (未保存)")
            self._log_result(f"📝 新建模板: {name}（请编辑后保存）")

    def _edit_current_template(self):
        """编辑当前模板 → 直接弹出保存对话框"""
        if not self.current_template_name:
            messagebox.showinfo("提示", "请先选择或新建一个模板")
            return
        # 直接弹出保存对话框，名称预填当前模板名
        self._save_as_template()

    def _delete_template_dialog(self):
        """删除模板对话框"""
        name = self.current_template_name
        if not name:
            messagebox.showinfo("提示", "请先选择要删除的模板")
            return
        if name not in _load_user_templates():
            messagebox.showinfo(
                "提示",
                f"「{name}」是系统默认模板，不能删除。\n"
                "如需恢复原始内容，请使用「重置」按钮。"
            )
            return
        if messagebox.askyesno("确认删除", f"确定删除自定义模板「{name}」？"):
            _delete_user_template(name)
            self._refresh_templates()
            # 加载第一个可用模板
            if self.templates:
                first = list(self.templates.keys())[0]
                self._load_template(first, silent=True)
            self._log_result(f"🗑 已删除模板: {name}")

    def _reset_template_dialog(self):
        """重置当前模板为默认"""
        name = self.current_template_name
        if not name:
            return
        if name not in DEFAULT_TEMPLATES:
            messagebox.showinfo("提示", f"「{name}」不是系统默认模板，无法重置")
            return
        if messagebox.askyesno("确认重置", f"将「{name}」恢复为系统默认内容？"):
            self._reset_template(name)
            self._load_template(name, silent=True)

    def _load_template_by_key(self, n):
        keys = list(self.templates.keys())
        if 0 <= n - 1 < len(keys):
            self._load_template(keys[n - 1])

    def _save_as_template(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("保存模板")
        dialog.geometry("350x150")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="模板名称:").pack(pady=(15, 5))
        name_var = tk.StringVar(value=self.current_template_name or "")
        entry = ttk.Entry(dialog, textvariable=name_var, width=30)
        entry.pack(pady=(0, 5))
        entry.select_range(0, tk.END)
        entry.focus()

        hint_var = tk.StringVar()
        hint_label = ttk.Label(dialog, textvariable=hint_var, foreground="gray")
        hint_label.pack()

        def on_name_change(*args):
            name = name_var.get().strip()
            if name in DEFAULT_TEMPLATES:
                hint_var.set(f"⚠️ 将覆盖默认模板「{name}」，可通过菜单恢复")
            else:
                hint_var.set("将创建新模板")

        name_var.trace_add("write", on_name_change)
        if self.current_template_name:
            on_name_change()

        def do_save():
            name = name_var.get().strip()
            if not name:
                return
            text = self.editor.get(1.0, "end-1c")
            _save_user_template(name, text)
            self._refresh_templates()
            self.current_template_name = name
            self.cur_template_label.configure(text=name)
            self._log_result(f"💾 模板已保存: {name}")
            dialog.destroy()

        ttk.Button(dialog, text="保存", command=do_save).pack(pady=(5, 0))
        entry.bind("<Return>", lambda e: do_save())
        dialog.bind("<Escape>", lambda e: dialog.destroy())

    def _reset_template(self, name):
        """恢复模板为系统默认"""
        if _delete_user_template(name):
            self._refresh_templates()
            self._log_result(f"🔄 已恢复默认模板: {name}")
            messagebox.showinfo("已恢复", f"模板「{name}」已恢复为系统默认。")

    # ── 转换 ─────────────────────────────────────────────────────

    def _build_font_map(self):
        """从配置构建 font_map"""
        fm = {}
        for role in ROLES:
            cfg = self.role_config.get(role, ROLE_DEFAULTS[role])
            fm[role] = cfg["font"]
        return fm

    def _convert(self):
        """执行转换"""
        text = self.editor.get(1.0, "end-1c")
        if not text.strip():
            messagebox.showwarning("提示", "请先输入 Markdown 内容")
            return

        out_dir = self.out_dir_var.get()
        out_name = self.out_name_var.get()
        if not out_name.endswith(".docx"):
            out_name += ".docx"
        out_path = os.path.join(out_dir, out_name)

        # 确保输出目录存在
        os.makedirs(out_dir, exist_ok=True)

        self._log_result("⏳ 正在转换...")
        self.root.update()

        try:
            font_map = self._build_font_map()
            gw.convert_md_to_docx(text, out_path, font_map)
            self.last_output = out_path
            self.open_btn.configure(state="normal")
            self._log_result(f"✅ 转换完成!\n\n📄 输出: {out_path}")

            # 验证
            result = gw.verify_docx(out_path)
            if result["errors"]:
                self._log_result(f"\n⚠️ 验证警告: {'; '.join(result['errors'])}")
            else:
                self._log_result(f"\n📊 段落数: {result['paragraphs']} | "
                                 f"字体: {', '.join(result['fonts_used'])}")
        except Exception as e:
            self._log_result(f"❌ 转换失败: {e}")
            messagebox.showerror("转换失败", str(e))

        # 确保编辑器始终可用且聚焦
        self.editor.configure(state="normal")
        self.editor.focus_set()
        self.root.update()

    def _open_output(self):
        if hasattr(self, 'last_output') and os.path.isfile(self.last_output):
            import subprocess
            if sys.platform == "darwin":
                subprocess.run(["open", "-R", self.last_output])
            elif sys.platform == "win32":
                subprocess.run(["explorer", "/select,", self.last_output])
            else:
                subprocess.run(["xdg-open", os.path.dirname(self.last_output)])

    # ── 格式配置 ────────────────────────────────────────────────

    def _open_font_config(self):
        # 扫描系统字体（缓存）
        if not hasattr(self, '_cached_fonts'):
            try:
                self._cached_fonts = gw.scan_fonts()
            except Exception:
                self._cached_fonts = []
        FontConfigDialog(
            self.root, self.role_config, self.page_config,
            on_save=self._on_font_config_saved,
            sys_fonts=self._cached_fonts,
        )

    def _on_font_config_saved(self, role_config, page_config):
        self.role_config = role_config
        self.page_config = page_config
        self.config["role_config"] = role_config
        self.config["page_config"] = page_config
        _save_config(self.config)
        self._refresh_format_overview()
        self._log_result("✅ 格式配置已保存")

    def _size_name(self, pt):
        """pt → 中文字号名，如 22→'二号'"""
        return SIZE_PT_TO_NAME.get(pt, f"{pt}pt")

    def _refresh_format_overview(self):
        """刷新格式速览表"""
        if not hasattr(self, 'fmt_labels'):
            return
        for role in ROLES:
            cfg = self.role_config.get(role, ROLE_DEFAULTS[role])
            indent_val = cfg.get("indent", 0)
            indent_text = f"缩{indent_val}字" if indent_val else ""
            vals = [
                role,
                cfg["font"][:12],
                self._size_name(cfg["size"]),
                "✓" if cfg["bold"] else "",
                "✓" if cfg["center"] else "",
                indent_text,
            ]
            if role in self.fmt_labels:
                for j, v in enumerate(vals):
                    self.fmt_labels[role][j].configure(text=v)

    # ── 首次启动 ────────────────────────────────────────────────

    def _show_first_run(self):
        FirstRunWizard(self.root, self.config, on_done=self._on_first_run_done)

    def _on_first_run_done(self):
        self.config = _load_config()
        self.out_dir_var.set(self.config.get("output_dir", ""))
        self._log_result("✅ 配置完成，开始使用吧！")

    # ── 其他 ─────────────────────────────────────────────────────

    def _log_result(self, msg):
        self.result_text.configure(state="normal")
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, msg)
        self.result_text.configure(state="disabled")

    def _about(self):
        messagebox.showinfo(
            "关于",
            f"{APP_TITLE}\n\n"
            "基于 md2gongwen.py 纯标准库实现\n"
            "公文格式遵循 GB/T 9704—2012\n\n"
            "GitHub: (your repo)"
        )


# ═══════════════════════════════════════════════════════════════════
# 入口
# ═══════════════════════════════════════════════════════════════════

def _get_icon_path():
    """获取图标路径（兼容源码运行和 pyinstaller 打包）"""
    # pyinstaller 打包后用 sys._MEIPASS
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    candidates = [
        os.path.join(base, "icon.png"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.png"),
    ]
    for p in candidates:
        if os.path.isfile(p):
            return p
    return None


def main():
    root = tk.Tk()

    # 加载图标
    icon_path = _get_icon_path()
    if icon_path:
        try:
            img = tk.PhotoImage(file=icon_path)
            root.iconphoto(True, img)
            # macOS Dock 图标（需额外的 tk 调用）
            if sys.platform == "darwin":
                try:
                    root.tk.call(
                        "::tk::unsupported::MacWindowStyle", "appearance", "system"
                    )
                except Exception:
                    pass
        except Exception:
            pass

    GongwenApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
