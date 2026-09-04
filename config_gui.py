import os
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "report_config.json")

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

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

        self.cfg = load_config()
        self.plot_vars = {}

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="12")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. 顶部：模板与路径
        path_group = ttk.LabelFrame(main_frame, text=" 模板与输出路径 ", padding="8")
        path_group.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(path_group, text="PPT 模板:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.tpl_var = tk.StringVar(value=self.cfg.get("template_pptx", ""))
        ttk.Entry(path_group, textvariable=self.tpl_var, width=58).grid(row=0, column=1, padx=5, pady=2, sticky=tk.EW)
        ttk.Button(path_group, text="浏览...", command=self.browse_template).grid(row=0, column=2, pady=2)

        ttk.Label(path_group, text="输出目录:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.out_var = tk.StringVar(value=self.cfg.get("output_dir", ""))
        ttk.Entry(path_group, textvariable=self.out_var, width=58).grid(row=1, column=1, padx=5, pady=2, sticky=tk.EW)
        ttk.Button(path_group, text="浏览...", command=self.browse_output).grid(row=1, column=2, pady=2)
        path_group.columnconfigure(1, weight=1)

        # 2. 核心画质与动图参数栏
        param_group = ttk.LabelFrame(main_frame, text=" 画质与充填动图 (GIF) 参数设置 ", padding="8")
        param_group.pack(fill=tk.X, pady=(0, 8))

        img_cfg = self.cfg.get("image_settings", {})
        anim_cfg = self.cfg.get("animation_settings", {})

        ttk.Label(param_group, text="截图画质:").grid(row=0, column=0, sticky=tk.W, pady=3)
        self.res_var = tk.StringVar(value=f"{img_cfg.get('width', 1920)}x{img_cfg.get('height', 1080)}")
        res_combo = ttk.Combobox(param_group, textvariable=self.res_var, values=[
            "1920x1080 (1080P 高清推荐)",
            "2560x1440 (2K 超高清)",
            "3840x2160 (4K 极清)",
            "0x0 (当前视口原生尺寸)"
        ], state="readonly", width=24)
        res_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=3)

        self.keep_view_var = tk.BooleanVar(value=img_cfg.get("keep_view", True))
        ttk.Checkbutton(param_group, text="保留操作界面模型摆放视角 (彻底防止与左侧色带重叠)", variable=self.keep_view_var).grid(row=0, column=2, sticky=tk.W, padx=5)

        ttk.Label(param_group, text="充填动画帧数:").grid(row=1, column=0, sticky=tk.W, pady=3)
        self.frames_var = tk.IntVar(value=anim_cfg.get("frames", 50))
        frame_spin = ttk.Spinbox(param_group, from_=10, to=120, increment=5, textvariable=self.frames_var, width=10)
        frame_spin.grid(row=1, column=1, sticky=tk.W, padx=5, pady=3)
        ttk.Label(param_group, text="(推荐 30~80 帧，与充填时间属性-帧数挂钩，帧数越大动图越细腻)", foreground="#0066CC").grid(row=1, column=2, sticky=tk.W, padx=5)

        ttk.Label(param_group, text="动图播放间隔:").grid(row=2, column=0, sticky=tk.W, pady=3)
        self.delay_var = tk.IntVar(value=anim_cfg.get("delay_ms", 80))
        delay_spin = ttk.Spinbox(param_group, from_=30, to=300, increment=10, textvariable=self.delay_var, width=10)
        delay_spin.grid(row=2, column=1, sticky=tk.W, padx=5, pady=3)
        ttk.Label(param_group, text="ms/帧 (默认 80ms，数值越小播放越快)", foreground="gray").grid(row=2, column=2, sticky=tk.W, padx=5)

        # 2.1 截图标尺方案选择 (隔离与对比)
        mode_group = ttk.LabelFrame(main_frame, text=" 截图标尺方案与隔离选择 (解决缺少左侧彩条标尺问题) ", padding="8")
        mode_group.pack(fill=tk.X, pady=(0, 8))

        self.mode_var = tk.StringVar(value=self.cfg.get("screenshot_mode", "B"))
        ttk.Radiobutton(mode_group, text="★ 方案 B：双图独立导出再智能拼合 (默认推荐·标尺与模型像素级排版，绝不遮挡，支持后台导出)", 
                        value="B", variable=self.mode_var).pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(mode_group, text="方案 A：视口直接抓取 (所见即所得，截取视口真实画面，含左侧彩条标尺/刻度/单位，执行时窗口需可见)", 
                        value="A", variable=self.mode_var).pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(mode_group, text="双方案同时生成 (方案 A 与方案 B 各生成一份独立 PPT，方便直观对比)", 
                        value="BOTH", variable=self.mode_var).pack(anchor=tk.W, pady=2)

        # 3. 网格类型自适应说明标签
        grid_note = ttk.Frame(main_frame)
        grid_note.pack(fill=tk.X, pady=(0, 4))
        note_lbl = tk.Label(grid_note, text="💡 提示: 本脚本已完整支持【3D实体网格】、【双层面(Dual Domain)】与【中性面(Midplane)】，运行时自动识别并自适应网格数据！",
                            bg="#E6F2FF", fg="#004080", font=("Microsoft YaHei", 9), padx=6, pady=3, anchor="w")
        note_lbl.pack(fill=tk.X)

        # 4. 结果项分类勾选区域 (Notebook 分页)
        list_group = ttk.LabelFrame(main_frame, text=" 分析结果导出项目选择 (支持 30+ 种结果) ", padding="8")
        list_group.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        btn_bar = ttk.Frame(list_group)
        btn_bar.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(btn_bar, text="★ 恢复默认常用 13 项", command=self.reset_default_plots).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_bar, text="全部勾选", width=9, command=self.select_all).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_bar, text="全部清空", width=9, command=self.deselect_all).pack(side=tk.LEFT, padx=3)

        self.notebook = ttk.Notebook(list_group)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        categories = [
            ("core", "★ 默认常用 13 项"),
            ("flow", "流动与充填分析"),
            ("pack", "保压与缩痕分析"),
            ("warp", "翘曲与各向变形"),
            ("cool", "冷却系统分析"),
            ("fiber", "加纤取向分析")
        ]

        plots = self.cfg.get("plots", [])
        
        # 初始化变量
        for p in plots:
            k = p["key"]
            self.plot_vars[k] = tk.BooleanVar(value=p.get("enabled", False))

        for cat_id, cat_title in categories:
            tab_frame = ttk.Frame(self.notebook, padding="5")
            self.notebook.add(tab_frame, text=cat_title)
            self.populate_tab(tab_frame, cat_id, plots)

        # 5. 底部选项与按钮
        bottom_options = ttk.Frame(main_frame)
        bottom_options.pack(fill=tk.X, pady=(0, 8))

        self.open_var = tk.BooleanVar(value=self.cfg.get("open_after_export", True))
        ttk.Checkbutton(bottom_options, text="生成后自动打开 PPT 报告", variable=self.open_var).pack(side=tk.LEFT)

        self.show_gui_var = tk.BooleanVar(value=self.cfg.get("show_gui_before_run", False))
        ttk.Checkbutton(bottom_options, text="每次点击宏运行时都弹出此窗口", variable=self.show_gui_var).pack(side=tk.LEFT, padx=20)

        action_bar = ttk.Frame(main_frame)
        action_bar.pack(fill=tk.X, pady=4)

        run_btn = tk.Button(action_bar, text="★ 保存配置并立即生成报告", bg="#0066CC", fg="white", font=("Microsoft YaHei", 10, "bold"),
                            padx=16, pady=6, relief=tk.RAISED, command=self.save_and_run)
        run_btn.pack(side=tk.LEFT, padx=5)

        save_btn = ttk.Button(action_bar, text="仅保存配置", command=self.save_only)
        save_btn.pack(side=tk.LEFT, padx=5)

        close_btn = ttk.Button(action_bar, text="关闭", command=self.root.destroy)
        close_btn.pack(side=tk.RIGHT, padx=5)

    def populate_tab(self, parent, cat_id, plots):
        canvas = tk.Canvas(parent, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)

        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        for p in plots:
            k = p["key"]
            name = p["plot_name"]
            desc = p.get("desc", name)
            slide_no = p.get("slide")
            cat = p.get("category", "flow")

            # 筛选
            if cat_id == "core":
                if not slide_no or slide_no > 16:
                    continue
            else:
                if cat != cat_id:
                    continue

            var = self.plot_vars[k]
            label_text = f"[{'P' + str(slide_no) if slide_no else '备选'}] {name}  ——  {desc}"
            cb = ttk.Checkbutton(scroll_frame, text=label_text, variable=var)
            cb.pack(anchor=tk.W, pady=3, padx=6)

    def browse_template(self):
        f = filedialog.askopenfilename(
            title="选择 PPT 模流分析报告模板",
            filetypes=[("PowerPoint Files", "*.pptx"), ("All Files", "*.*")]
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
        # 仅默认勾选核心 13 项 (Slide 4 ~ 16)
        for p in self.cfg.get("plots", []):
            k = p["key"]
            slide_no = p.get("slide")
            if slide_no and 4 <= slide_no <= 16:
                self.plot_vars[k].set(True)
            else:
                self.plot_vars[k].set(False)

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

        if "image_settings" not in self.cfg:
            self.cfg["image_settings"] = {}
        self.cfg["image_settings"]["width"] = w
        self.cfg["image_settings"]["height"] = h
        self.cfg["image_settings"]["keep_view"] = self.keep_view_var.get()

        if "animation_settings" not in self.cfg:
            self.cfg["animation_settings"] = {}
        self.cfg["animation_settings"]["frames"] = self.frames_var.get()
        self.cfg["animation_settings"]["delay_ms"] = self.delay_var.get()

        for p in self.cfg.get("plots", []):
            k = p["key"]
            if k in self.plot_vars:
                p["enabled"] = self.plot_vars[k].get()

        return self.cfg

    def save_only(self):
        cfg = self.collect_config()
        save_config(cfg)
        messagebox.showinfo("成功", "配置已成功保存！")

    def save_and_run(self):
        cfg = self.collect_config()
        save_config(cfg)
        self.root.destroy()

        vbs_driver = os.path.join(SCRIPT_DIR, "AutoReport.vbs")
        if os.path.exists(vbs_driver):
            subprocess.Popen(["wscript", vbs_driver])

def main():
    root = tk.Tk()
    app = ConfigApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
