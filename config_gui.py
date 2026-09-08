"""模流分析报告配置中心 (Tkinter GUI)。

T27/T28 重构要点:
- 主行动栏置顶, 模式单选一行化, 模板/画质等一次性配置收进"高级设置"折叠区;
- 点"生成"后窗口转为运行状态: 实时 tail temp\run.log (GBK 安全分块解码) + 取消按钮
  (taskkill /T /F, 带 poll() 守卫);
- 配置文件损坏时兜底默认配置 + 错误弹窗, 不再静默崩溃;
- 由 VBS (Moldflow 宏) 拉起时 (环境变量 MLFXBG_FROM_GUI=1) 只保存配置并退出,
  绝不二次拉起 VBS, 杜绝双重运行。
"""

import codecs
import json
import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "report_config.json")
RUN_LOG = os.path.join(SCRIPT_DIR, "temp", "run.log")
AVAILABLE_PLOTS_FILE = os.path.join(SCRIPT_DIR, "temp", "available_plots.json")
LIST_PLOTS_VBS = os.path.join(SCRIPT_DIR, "list_plots.vbs")

MODE_LABELS = [
    ("B", "方案 B (推荐): 标尺与模型分开导出后智能拼合, 支持后台导出"),
    ("A", "方案 A: 视口直接抓取 (所见即所得, 运行时 Moldflow 窗口需可见)"),
    ("BOTH", "双方案对比: 方案 A 与 B 各生成一份独立报告"),
]


def load_config():
    """读取配置; 损坏/缺失 → ({}, 错误信息), 绝不让 GUI 静默崩溃。"""
    if not os.path.exists(CONFIG_FILE):
        return {}, None
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return {}, "配置根元素不是对象"
        return data, None
    except Exception as e:
        return {}, str(e)


def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def find_duplicate_slides(plots):
    """
    勾选且页码相同的结果项 → {页码: [结果名, ...]} (只列撞页的)。
    未勾选 / 无页码 / 页码非法的跳过。
    """
    seen = {}
    for p in plots:
        if not p.get("enabled"):
            continue
        try:
            slide = int(p.get("slide"))
        except (TypeError, ValueError):
            continue
        if slide <= 0:
            continue
        seen.setdefault(slide, []).append(str(p.get("plot_name") or p.get("key") or ""))
    return {s: names for s, names in seen.items() if len(names) > 1}


def match_plots_to_available(plots, available_names):
    """
    结果名与方案内实际结果清单精确比对 → {key: 是否命中}。
    只认一字不差 (用户裁决: 别名/模糊匹配彻底移除, 绝不猜)。
    """
    available = set(available_names or [])
    return {p.get("key"): p.get("plot_name") in available for p in plots}


def extract_new_log_text(path, offset, decoder_state=None):
    """读取日志文件自 offset 起的新增内容, 返回 (text, new_offset)。

    用 codecs 增量解码器处理 GBK 双字节字符跨块: 不完整的尾字符留在
    解码器内部状态, 下次读取续上; 文件被重建 (变小) 时重置解码器从头读。
    decoder_state: 调用方持有的 IncrementalDecoder (首次传 None 自动创建)。
    """
    try:
        size = os.path.getsize(path)
    except OSError:
        return "", offset, decoder_state
    if size < offset:
        offset = 0
        decoder_state = None
    if size == offset:
        return "", offset, decoder_state
    try:
        with open(path, "rb") as f:
            f.seek(offset)
            raw = f.read()
    except OSError:
        return "", offset, decoder_state
    if decoder_state is None:
        decoder_state = codecs.getincrementaldecoder("gbk")(errors="replace")
    text = decoder_state.decode(raw)
    return text, offset + len(raw), decoder_state


class ConfigApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Moldflow 2023 模流分析报告配置中心 (多网格自适应版)")
        self.root.geometry("780x860")
        self.root.minsize(700, 720)

        self.style = ttk.Style()
        try:
            self.style.theme_use("vista")
        except Exception:
            pass

        self.cfg, cfg_err = load_config()
        if cfg_err:
            messagebox.showerror(
                "配置文件损坏",
                f"report_config.json 无法读取: {cfg_err}\n"
                f"已用默认配置启动; 保存时将覆盖损坏文件。",
            )
        self.plot_vars = {}
        self.slide_vars = {}
        self.plot_checkbuttons = {}
        self.plot_label_base = {}
        self.adv_open = False

        self.run_proc = None
        self.log_offset = 0
        self.log_decoder = None

        self.create_widgets()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.bind("<Return>", self.on_enter_key)
        self.root.after(400, self.startup_env_check)
        self.root.after(600, self.refresh_available_plots_async)

    def startup_env_check(self):
        """T13/T29: 启动时环境预检 (依赖/配置/模板标记), 问题只提示不阻断保存。"""
        try:
            sys.path.insert(0, SCRIPT_DIR)
            import check_env as ce

            problems, infos = [], []
            p, i = ce.check_python()
            infos += i
            p, i = ce.check_dependencies()
            problems += p
            infos += i
            p, i = ce.check_config(CONFIG_FILE)
            problems += p
            infos += i
            if not problems and self.cfg.get("template_pptx"):
                p, i = ce.check_template(self.cfg.get("template_pptx"))
                problems += p
            if problems:
                messagebox.showwarning(
                    "环境预检发现问题",
                    "以下问题可能影响报告生成:\n\n"
                    + "\n".join(problems)
                    + "\n\n(生成前请修正; 详见 check_env.py)",
                )
        except Exception as e:
            print(f"[Notice] 环境预检跳过: {e}")

    # ---------------- 界面搭建 (T28: 行动栏 → 模式 → 勾选区 → 高级设置折叠) ----------------

    def create_widgets(self):
        main = ttk.Frame(self.root, padding="10")
        main.pack(fill=tk.BOTH, expand=True)

        # 1. 常驻行动栏 (用户第一眼看到的就是主行动)
        action_bar = ttk.Frame(main)
        action_bar.pack(fill=tk.X, pady=(0, 6))
        self.run_btn = tk.Button(
            action_bar,
            text="★ 保存配置并立即生成报告",
            bg="#0066CC",
            fg="white",
            font=("Microsoft YaHei", 10, "bold"),
            padx=14,
            pady=5,
            relief=tk.RAISED,
            command=self.save_and_run,
        )
        self.run_btn.pack(side=tk.LEFT)
        ttk.Button(action_bar, text="仅保存配置", command=self.save_only).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(action_bar, text="打开运行日志", command=self.open_log).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(action_bar, text="关闭", command=self.on_close).pack(side=tk.RIGHT)

        # 2. 方案选择 (一行一条, 工具文案)
        mode_group = ttk.LabelFrame(main, text=" 截图方案 ", padding="6")
        mode_group.pack(fill=tk.X, pady=(0, 6))
        self.mode_var = tk.StringVar(value=self.cfg.get("screenshot_mode", "B"))
        for value, label in MODE_LABELS:
            ttk.Radiobutton(
                mode_group, text=label, value=value, variable=self.mode_var
            ).pack(anchor=tk.W, pady=1)

        # 3. 结果项勾选区 (占主要空间)
        list_group = ttk.LabelFrame(
            main, text=" 分析结果导出项目选择 (支持 30+ 种结果) ", padding="6"
        )
        list_group.pack(fill=tk.BOTH, expand=True, pady=(0, 6))
        btn_bar = ttk.Frame(list_group)
        btn_bar.pack(fill=tk.X, pady=(0, 4))
        ttk.Button(
            btn_bar, text="★ 恢复默认常用 13 项", command=self.reset_default_plots
        ).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_bar, text="全部勾选", width=9, command=self.select_all).pack(
            side=tk.LEFT, padx=3
        )
        ttk.Button(btn_bar, text="全部清空", width=9, command=self.deselect_all).pack(
            side=tk.LEFT, padx=3
        )
        ttk.Button(
            btn_bar,
            text="⟳ 从 Moldflow 刷新结果",
            command=self.refresh_available_plots_async,
        ).pack(side=tk.LEFT, padx=3)
        self.avail_label = ttk.Label(
            btn_bar, text="结果清单: 刷新中...", foreground="#888888"
        )
        self.avail_label.pack(side=tk.RIGHT)
        self.notebook = ttk.Notebook(list_group)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        categories = [
            ("core", "★ 默认常用 13 项"),
            ("flow", "流动与充填分析"),
            ("pack", "保压与缩痕分析"),
            ("warp", "翘曲与各向变形"),
            ("cool", "冷却系统分析"),
            ("fiber", "加纤取向分析"),
        ]
        plots = self.cfg.get("plots", [])
        for p in plots:
            self.plot_vars[p["key"]] = tk.BooleanVar(value=p.get("enabled", False))
        for cat_id, cat_title in categories:
            tab_frame = ttk.Frame(self.notebook, padding="5")
            self.notebook.add(tab_frame, text=cat_title)
            self.populate_tab(tab_frame, cat_id, plots)

        # 4. 高级设置折叠区 (一次性配置)
        self.adv_btn = ttk.Button(
            main,
            text="▸ 高级设置 (模板 / 输出目录 / 画质 / 动图 / 数据策略)",
            command=self.toggle_advanced,
        )
        self.adv_btn.pack(fill=tk.X, pady=(0, 4))
        self.adv_frame = ttk.LabelFrame(main, text=" 高级设置 ", padding="8")

        path_group = ttk.Frame(self.adv_frame)
        path_group.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(path_group, text="PPT 模板:").grid(row=0, column=0, sticky=tk.W)
        self.tpl_var = tk.StringVar(value=self.cfg.get("template_pptx", ""))
        ttk.Entry(path_group, textvariable=self.tpl_var).grid(
            row=0, column=1, sticky=tk.EW, padx=5, pady=2
        )
        ttk.Button(path_group, text="浏览...", command=self.browse_template).grid(
            row=0, column=2
        )
        ttk.Label(path_group, text="输出目录:").grid(row=1, column=0, sticky=tk.W)
        self.out_var = tk.StringVar(value=self.cfg.get("output_dir", ""))
        ttk.Entry(path_group, textvariable=self.out_var).grid(
            row=1, column=1, sticky=tk.EW, padx=5, pady=2
        )
        ttk.Button(path_group, text="浏览...", command=self.browse_output).grid(
            row=1, column=2
        )
        path_group.columnconfigure(1, weight=1)

        img_cfg = self.cfg.get("image_settings", {})
        anim_cfg = self.cfg.get("animation_settings", {})
        quality_group = ttk.Frame(self.adv_frame)
        quality_group.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(quality_group, text="截图画质:").grid(
            row=0, column=0, sticky=tk.W, pady=2
        )
        self.res_var = tk.StringVar(
            value=f"{img_cfg.get('width', 1920)}x{img_cfg.get('height', 1080)}"
        )
        ttk.Combobox(
            quality_group,
            textvariable=self.res_var,
            values=[
                "1920x1080 (1080P 高清)",
                "2560x1440 (2K 超高清推荐)",
                "3840x2160 (4K 极清)",
                "0x0 (当前视口原生尺寸)",
            ],
            state="readonly",
            width=26,
        ).grid(row=0, column=1, sticky=tk.W, padx=5)
        ttk.Label(quality_group, text="充填动画帧数:").grid(
            row=1, column=0, sticky=tk.W, pady=2
        )
        self.frames_var = tk.IntVar(value=anim_cfg.get("frames", 50))
        ttk.Spinbox(
            quality_group,
            from_=10,
            to=120,
            increment=5,
            textvariable=self.frames_var,
            width=8,
        ).grid(row=1, column=1, sticky=tk.W, padx=5)
        ttk.Label(quality_group, text="(推荐 30~80 帧)").grid(
            row=1, column=2, sticky=tk.W
        )
        ttk.Label(quality_group, text="动图播放间隔:").grid(
            row=2, column=0, sticky=tk.W, pady=2
        )
        self.delay_var = tk.IntVar(value=anim_cfg.get("delay_ms", 80))
        ttk.Spinbox(
            quality_group,
            from_=30,
            to=300,
            increment=10,
            textvariable=self.delay_var,
            width=8,
        ).grid(row=2, column=1, sticky=tk.W, padx=5)
        ttk.Label(quality_group, text="ms/帧 (默认 80)").grid(
            row=2, column=2, sticky=tk.W
        )
        self.keep_view_var = tk.BooleanVar(value=img_cfg.get("keep_view", True))
        ttk.Checkbutton(
            quality_group,
            text="保留操作界面模型摆放视角 (防止与左侧色带重叠)",
            variable=self.keep_view_var,
        ).grid(row=3, column=0, columnspan=3, sticky=tk.W)

        policy_group = ttk.Frame(self.adv_frame)
        policy_group.pack(fill=tk.X, pady=(0, 4))
        machine_row = ttk.Frame(policy_group)
        machine_row.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(machine_row, text="注塑机最大锁模力(吨):").pack(side=tk.LEFT)
        cur_ton = (self.cfg.get("clamp_force_settings") or {}).get("machine_max_ton")
        self.machine_ton_var = tk.StringVar(
            value=""
            if not isinstance(cur_ton, (int, float)) or cur_ton <= 0
            else str(cur_ton)
        )
        ttk.Entry(machine_row, textvariable=self.machine_ton_var, width=10).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Label(
            machine_row,
            text="(设备规格, 按本次产品实际机台填写; 留空则报告标'数据缺失')",
            foreground="#888888",
        ).pack(side=tk.LEFT, padx=4)
        self.open_var = tk.BooleanVar(value=self.cfg.get("open_after_export", True))
        ttk.Checkbutton(
            policy_group, text="生成后自动打开 PPT 报告", variable=self.open_var
        ).pack(anchor=tk.W)
        self.show_gui_var = tk.BooleanVar(
            value=self.cfg.get("show_gui_before_run", False)
        )
        ttk.Checkbutton(
            policy_group,
            text="每次点击宏运行时都弹出此窗口",
            variable=self.show_gui_var,
        ).pack(anchor=tk.W)
        self.on_missing_var = tk.StringVar(
            value=self.cfg.get("on_missing_data", "annotate")
        )
        ttk.Radiobutton(
            policy_group,
            text="数据缺失时: 标注缺失并继续生成 (推荐)",
            value="annotate",
            variable=self.on_missing_var,
        ).pack(anchor=tk.W)
        ttk.Radiobutton(
            policy_group,
            text="数据缺失时: 拒绝生成 (需补齐数据后重跑)",
            value="fail",
            variable=self.on_missing_var,
        ).pack(anchor=tk.W)

        # 5. 运行状态区 (T27: 点生成后显示, 含取消)
        self.status_frame = ttk.LabelFrame(main, text=" 运行状态 ", padding="6")
        self.log_text = tk.Text(
            self.status_frame,
            height=10,
            state=tk.DISABLED,
            font=("Consolas", 9),
            bg="#1e1e1e",
            fg="#d4d4d4",
        )
        log_scroll = ttk.Scrollbar(
            self.status_frame, orient=tk.VERTICAL, command=self.log_text.yview
        )
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        status_bar = ttk.Frame(self.status_frame)
        status_bar.pack(fill=tk.X, pady=(4, 0))
        self.status_label = ttk.Label(status_bar, text="正在启动自动化流程...")
        self.status_label.pack(side=tk.LEFT)
        ttk.Button(status_bar, text="打开日志", command=self.open_log).pack(
            side=tk.RIGHT
        )
        self.cancel_btn = tk.Button(
            status_bar,
            text="取消生成",
            bg="#c00000",
            fg="white",
            command=self.cancel_run,
        )
        self.cancel_btn.pack(side=tk.RIGHT, padx=6)

    def toggle_advanced(self):
        if self.adv_open:
            self.adv_frame.pack_forget()
            self.adv_btn.configure(
                text="▸ 高级设置 (模板 / 输出目录 / 画质 / 动图 / 数据策略)"
            )
        else:
            self.adv_frame.pack(fill=tk.X, before=self.notebook.master, pady=(0, 6))
            self.adv_btn.configure(
                text="▾ 高级设置 (模板 / 输出目录 / 画质 / 动图 / 数据策略)"
            )
        self.adv_open = not self.adv_open

    def populate_tab(self, parent, cat_id, plots):
        canvas = tk.Canvas(parent, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        scroll_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        for p in plots:
            k = p["key"]
            slide_no = p.get("slide")
            cat = p.get("category", "flow")
            if cat_id == "core":
                if not slide_no or slide_no > 16:
                    continue
            elif cat != cat_id:
                continue
            desc = p.get("desc", p["plot_name"])
            label_text = f"[{'P' + str(slide_no) if slide_no else '备选'}] {p['plot_name']}  ——  {desc}"
            self.plot_label_base.setdefault(k, label_text)
            row = ttk.Frame(scroll_frame)
            row.pack(anchor=tk.W, pady=3, padx=6)
            cb = ttk.Checkbutton(
                row, text=self.plot_label_base[k], variable=self.plot_vars[k]
            )
            cb.pack(side=tk.LEFT)
            self.plot_checkbuttons.setdefault(k, []).append(cb)
            # 页码可改 (用户裁决): 1~16, 撞页在保存时拦截; 备选项无页码不显示
            if slide_no:
                if k not in self.slide_vars:
                    self.slide_vars[k] = tk.StringVar(value=str(slide_no))
                    self.slide_vars[k].trace_add(
                        "write", lambda *_a, key=k: self._update_plot_label(key)
                    )
                ttk.Label(row, text="页码").pack(side=tk.LEFT, padx=(10, 2))
                ttk.Spinbox(
                    row,
                    from_=1,
                    to=16,
                    width=3,
                    textvariable=self.slide_vars[k],
                    validate="key",
                    validatecommand=(
                        self.root.register(lambda v: v.isdigit() or v == ""),
                        "%P",
                    ),
                ).pack(side=tk.LEFT)

    def _update_plot_label(self, key):
        """页码 Spinbox 变化时同步行标签前缀 [Pn]。"""
        base = self.plot_label_base.get(key, "")
        body = "] ".join(base.split("] ")[1:]) if "] " in base else base
        sv = self.slide_vars.get(key)
        prefix = (
            f"[P{sv.get().strip()}]" if sv and sv.get().strip().isdigit() else "[?]"
        )
        for w in self.plot_checkbuttons.get(key, []):
            w.configure(text=f"{prefix} {body}")

    # ---------------- 配置读写 ----------------

    def browse_template(self):
        f = filedialog.askopenfilename(
            title="选择 PPT 模流分析报告模板",
            filetypes=[("PowerPoint Files", "*.pptx"), ("All Files", "*.*")],
        )
        if f:
            self.tpl_var.set(f)

    def browse_output(self):
        d = filedialog.askdirectory(title="选择报告输出目录")
        if d:
            self.out_var.set(d)

    def select_all(self):
        for var in self.plot_vars.values():
            var.set(True)

    def deselect_all(self):
        for var in self.plot_vars.values():
            var.set(False)

    def reset_default_plots(self):
        for p in self.cfg.get("plots", []):
            slide_no = p.get("slide")
            enabled = bool(slide_no and 4 <= slide_no <= 16)
            self.plot_vars[p["key"]].set(enabled)

    def collect_config(self):
        res_str = self.res_var.get().split(" ")[0]
        if "x" in res_str:
            w, h = map(int, res_str.split("x"))
        else:
            w, h = 1920, 1080

        self.cfg["template_pptx"] = self.tpl_var.get().strip()
        self.cfg["output_dir"] = self.out_var.get().strip()
        self.cfg["open_after_export"] = self.open_var.get()
        self.cfg["show_gui_before_run"] = self.show_gui_var.get()
        self.cfg["screenshot_mode"] = self.mode_var.get()
        self.cfg["on_missing_data"] = self.on_missing_var.get()

        if "image_settings" not in self.cfg:
            self.cfg["image_settings"] = {}
        self.cfg["image_settings"]["width"] = w
        self.cfg["image_settings"]["height"] = h
        self.cfg["image_settings"]["keep_view"] = self.keep_view_var.get()

        if "animation_settings" not in self.cfg:
            self.cfg["animation_settings"] = {}
        self.cfg["animation_settings"]["frames"] = self.frames_var.get()
        self.cfg["animation_settings"]["delay_ms"] = self.delay_var.get()

        # 注塑机吨位: 本次产品实际机台规格; 留空/非法 → null (报告标数据缺失, 绝不沿用旧值)
        try:
            ton_val = float(self.machine_ton_var.get())
            machine_ton = ton_val if ton_val > 0 else None
        except ValueError:
            machine_ton = None
        self.cfg.setdefault("clamp_force_settings", {})["machine_max_ton"] = machine_ton

        for p in self.cfg.get("plots", []):
            k = p["key"]
            if k in self.plot_vars:
                p["enabled"] = self.plot_vars[k].get()
            if k in self.slide_vars and p.get("slide"):
                try:
                    p["slide"] = int(self.slide_vars[k].get().strip())
                except (TypeError, ValueError):
                    pass  # 非法页码保留原值; 撞页校验兜底
        return self.cfg

    def _duplicate_slides_error(self):
        """撞页检查 → 错误弹窗文案; 无撞页返回 None。"""
        dups = find_duplicate_slides(self.collect_config().get("plots", []))
        if not dups:
            return None
        lines = [f"  P{s}: " + "、".join(names) for s, names in sorted(dups.items())]
        return (
            "以下页码被多个勾选结果同时占用:\n"
            + "\n".join(lines)
            + "\n\n怎么办: 修改其中一项的页码 (每页只能放一个结果)"
        )

    def save_only(self):
        dup_err = self._duplicate_slides_error()
        if dup_err:
            messagebox.showerror("页码冲突", dup_err)
            return
        try:
            save_config(self.collect_config())
        except Exception as e:
            messagebox.showerror(
                "保存失败",
                f"问题: 配置写入失败\n原因: {e}\n怎么办: 检查文件是否被占用后重试",
            )
            return
        messagebox.showinfo("成功", "配置已成功保存！")

    # ---------------- 运行 (T27: 状态窗 + 取消) / 校验 (T28 前置校验) ----------------

    def validate_before_run(self):
        problems = []
        dup_err = self._duplicate_slides_error()
        if dup_err:
            problems.append(dup_err)
        template = self.tpl_var.get().strip()
        if not template:
            problems.append("PPT 模板路径为空 → 在高级设置中选择模板文件")
        elif not os.path.exists(template):
            problems.append(f"模板文件不存在: {template} → 重新选择有效模板")
        if not any(v.get() for v in self.plot_vars.values()):
            problems.append("未勾选任何分析结果项 → 至少勾选一项再生成")
        if problems:
            messagebox.showerror(
                "无法启动生成", "问题: 配置未通过预检\n" + "\n".join(problems)
            )
            return False
        return True

    def save_and_run(self):
        if self.run_proc is not None and self.run_proc.poll() is None:
            messagebox.showwarning("正在运行", "当前已有一次生成在运行中。")
            return
        if not self.validate_before_run():
            return
        try:
            save_config(self.collect_config())
        except Exception as e:
            messagebox.showerror(
                "保存失败",
                f"问题: 配置写入失败\n原因: {e}\n怎么办: 检查文件是否被占用后重试",
            )
            return

        vbs_driver = os.path.join(SCRIPT_DIR, "AutoReport.vbs")
        if not os.path.exists(vbs_driver):
            messagebox.showerror(
                "缺少 AutoReport.vbs",
                f"问题: 自动化主脚本缺失\n原因: 未找到 {vbs_driver}\n"
                f"怎么办: 确认部署完整 (AutoReport.vbs 应与本程序同目录)",
            )
            return

        env = dict(os.environ, MLFXBG_FROM_GUI="1")
        try:
            self.run_proc = subprocess.Popen(
                ["wscript", vbs_driver], cwd=SCRIPT_DIR, env=env
            )
        except Exception as e:
            messagebox.showerror(
                "启动失败",
                f"问题: 无法启动自动化脚本\n原因: {e}\n怎么办: 确认 wscript 可用",
            )
            return
        self.enter_run_mode()

    def enter_run_mode(self):
        self.run_btn.configure(state=tk.DISABLED)
        self.cancel_btn.configure(state=tk.NORMAL)
        self.status_label.configure(text="正在运行... (Moldflow 导出与报告构建)")
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)
        self.log_offset = os.path.getsize(RUN_LOG) if os.path.exists(RUN_LOG) else 0
        self.log_decoder = None
        self.root.after(600, self.poll_run)

    def append_log(self, text):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, text)
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def poll_run(self):
        if self.run_proc is None:
            return
        rc = self.run_proc.poll()
        if os.path.exists(RUN_LOG):
            text, self.log_offset, self.log_decoder = extract_new_log_text(
                RUN_LOG, self.log_offset, self.log_decoder
            )
            if text:
                self.append_log(text)
        if rc is None:
            self.root.after(700, self.poll_run)
        else:
            self.finish_run(rc)

    def cancel_run(self):
        if self.run_proc is None or self.run_proc.poll() is not None:
            return
        if not messagebox.askyesno(
            "取消生成",
            "确定要终止本次生成吗?\n已导出的临时文件将在下次运行时自动清理。",
        ):
            return
        if sys.platform == "win32" and self.run_proc.poll() is None:
            subprocess.run(
                ["taskkill", "/PID", str(self.run_proc.pid), "/T", "/F"],
                capture_output=True,
            )
        try:
            self.run_proc.wait(timeout=10)
        except Exception:
            pass
        self.invalidate_partial_output()
        self.append_log("\n[已由用户取消]\n")
        self.finish_run(self.run_proc.poll() if self.run_proc else -1, cancelled=True)

    def invalidate_partial_output(self):
        """取消后失效本次产物指针; 完整性由下次运行的陈旧产物清理兜底。"""
        try:
            marker = os.path.join(SCRIPT_DIR, "temp", "last_output_path.txt")
            if os.path.exists(marker):
                os.remove(marker)
        except Exception as e:
            print(f"[Notice] Could not invalidate last_output_path.txt: {e}")

    def finish_run(self, rc, cancelled=False):
        self.run_proc = None
        self.run_btn.configure(state=tk.NORMAL)
        self.cancel_btn.configure(state=tk.DISABLED)
        self.status_label.configure(
            text=("已取消" if cancelled else f"运行结束 (退出码 {rc})")
        )
        if cancelled:
            return
        if rc == 0:
            # T1/T13: 完成弹窗附数据完整性清单 (缺失字段要求操作者知情确认)
            missing_info = ""
            try:
                mf = os.path.join(SCRIPT_DIR, "temp", "missing_fields.json")
                if os.path.exists(mf):
                    with open(mf, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    items = data.get("missing") or []
                    if items:
                        shown = "\n".join(f"  - {m}" for m in items[:8])
                        if len(items) > 8:
                            shown += f"\n  ...共 {len(items)} 项 (详见 temp\\missing_fields.json)"
                        missing_info = (
                            f"\n\n⚠ 数据完整性: {len(items)} 项缺失"
                            f" (报告内已标注并盖章):\n{shown}"
                        )
            except Exception as e:
                print(f"[Notice] 读取缺失清单失败: {e}")
            messagebox.showinfo(
                "生成完成",
                "模流分析报告已生成。\n如未自动打开, 请查看输出目录。" + missing_info,
            )
        else:
            meaning = {1: "环境问题 (模板缺失/Python 依赖缺失)", 2: "构建失败"}.get(
                rc, "未知错误"
            )
            messagebox.showerror(
                "报告生成失败",
                f"问题: 自动化流程异常退出 (退出码 {rc} = {meaning})\n"
                f"原因: 详见运行日志 (点'打开日志'查看尾部)\n"
                f"怎么办: 常见原因 — Moldflow 未打开/方案未打开/模板路径无效",
            )

    def refresh_available_plots_async(self):
        """后台跑 list_plots.vbs 从 Moldflow 拿结果清单, 不卡界面。"""
        threading.Thread(target=self._refresh_available_worker, daemon=True).start()

    def _refresh_available_worker(self):
        result = {"ok": False, "study": "", "plots": [], "error": ""}
        try:
            subprocess.run(
                ["wscript", "//nologo", LIST_PLOTS_VBS],
                cwd=SCRIPT_DIR,
                capture_output=True,
                timeout=30,
            )
            with open(AVAILABLE_PLOTS_FILE, "r", encoding="utf-8-sig") as f:
                data = json.load(f)
            if isinstance(data, dict):
                result["study"] = str(data.get("study", "") or "")
                plots = data.get("plots")
                result["plots"] = (
                    [str(x) for x in plots] if isinstance(plots, list) else []
                )
                result["error"] = str(data.get("error", "") or "")
                result["ok"] = bool(result["plots"])
            else:
                result["error"] = "清单格式异常"
        except Exception as e:
            result["error"] = str(e)
        self.root.after(0, lambda: self.apply_available_plots(result))

    def apply_available_plots(self, result):
        """方案结果清单应用到界面: 默认13项按精确名匹配, 命中才勾选 (用户裁决:
        动态列表 + 不猜测)。备选项保持现状; 匹配不上的行灰显提示。"""
        if not result.get("ok"):
            hint = result.get("error") or "未取到"
            self.avail_label.configure(
                text=f"结果清单: 未取到 ({hint}; Moldflow 没开着?)",
                foreground="#b0b0b0",
            )
            return
        names = result.get("plots", [])
        matched = match_plots_to_available(self.cfg.get("plots", []), names)
        default_plots = [
            p
            for p in self.cfg.get("plots", [])
            if p.get("slide") and 4 <= p.get("slide") <= 16
        ]
        n13_hit = sum(1 for p in default_plots if matched.get(p.get("key")))
        self.avail_label.configure(
            text=(
                f"方案 {result.get('study', '')}: 共 {len(names)} 个结果, "
                f"默认13项命中 {n13_hit}/{len(default_plots)}"
            ),
            foreground="#1a7f37",
        )
        for p in default_plots:
            k = p.get("key")
            hit = matched.get(k, False)
            self.plot_vars[k].set(hit)
            base = self.plot_label_base.get(k, "")
            for w in self.plot_checkbuttons.get(k, []):
                w.configure(
                    text=base if hit else f"{base}   （方案中未找到）",
                    foreground=None if hit else "#b0b0b0",
                )

    def open_log(self):
        if os.path.exists(RUN_LOG):
            os.startfile(RUN_LOG)
        else:
            messagebox.showinfo("暂无日志", "尚未产生运行日志。")

    def on_enter_key(self, event):
        if self.run_proc is None or self.run_proc.poll() is not None:
            self.run_btn.invoke()
        return "break"

    def on_close(self):
        if self.run_proc is not None and self.run_proc.poll() is None:
            if messagebox.askyesno(
                "正在运行", "生成仍在运行中。关闭窗口将终止本次生成, 确定吗?"
            ):
                self.cancel_run()
                return
            return
        self.root.destroy()


def main():
    root = tk.Tk()
    app = ConfigApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
