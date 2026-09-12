"""环境预检 — 全新机器跑通报告前的自助检查 (T29)。

用法: python check_env.py [--config report_config.json]
检查: Python 版本 / 运行时依赖 / 配置文件 / 模板 PPTX 存在性与幻灯片标记。
退出码 0 = 全部通过; 1 = 存在问题 (逐条列出怎么办)。
不依赖 Moldflow, 不修改任何文件。
"""

import json
import os
import sys

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))

REQUIRED_DEPS = [
    ("PIL", "Pillow", "12.3.0"),
    ("pptx", "python-pptx", "1.0.2"),
    ("numpy", "numpy", "2.5.2"),
]

# 模板契约标记 (依据 core/pptx_builder.py 的封面/网格页回填逻辑)
MARKERS_SLIDE1 = ("报告日期", "Moldflow")
MARKERS_SLIDE2 = ("实体计数", "三角")


def check_python():
    v = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
    return True, [f"Python {v} (本项目实测环境 3.14.6)"]


def check_dependencies():
    problems, infos = [], []
    for module_name, pkg_name, pinned in REQUIRED_DEPS:
        try:
            mod = __import__(module_name)
            version = getattr(mod, "__version__", "未知")
            infos.append(f"{pkg_name} {version}")
            if version != "未知" and version != pinned:
                problems.append(
                    f"{pkg_name} 版本 {version} 与锁定版本 {pinned} 不一致"
                    f" → 执行 pip install {pkg_name}=={pinned}"
                )
        except ImportError:
            problems.append(
                f"缺少 {pkg_name} → 执行 pip install {pkg_name}=={pinned}"
                f" (或 pip install -r requirements.txt)"
            )
    return problems, infos


def check_config(config_path):
    problems, infos = [], []
    if not os.path.exists(config_path):
        return [f"缺少配置文件 {config_path} → 部署三件套后重试"], []
    try:
        with open(config_path, "r", encoding="utf-8-sig") as f:
            cfg = json.load(f)
    except Exception as e:
        return [f"配置文件损坏 ({config_path}): {e} → 修正 JSON 语法"], []
    if not cfg.get("template_pptx"):
        problems.append(
            "配置缺少 template_pptx → 在 templates/ 放置模板并在配置中填写相对路径"
        )
    # output_dir 允许为空: 空值 = 默认输出到脚本同级目录 (项目根), 由 pptx_builder 解析
    if not cfg.get("output_dir"):
        infos.append("output_dir 为空 → 报告默认输出到脚本同级目录 (项目根)")
    infos.append(f"配置文件 OK: {config_path}")
    return problems, infos


def resolve_template_path(template_path):
    """模板相对路径 → 相对本项目根的绝对路径 (与 core/pptx_builder 同规则)。"""
    if not template_path:
        return template_path
    if os.path.isabs(str(template_path)):
        return template_path
    return os.path.join(REPO_ROOT, template_path)


def check_template(template_path):
    problems, infos = [], []
    template_path = resolve_template_path(template_path)
    if not template_path or not os.path.exists(template_path):
        return (
            [
                f"模板 PPTX 不存在: {template_path}"
                " → 把配置 template_pptx 指向正确的模板文件"
            ],
            [],
        )
    try:
        from pptx import Presentation
    except ImportError:
        return ["python-pptx 未安装, 无法检查模板 → pip install python-pptx"], []
    try:
        prs = Presentation(template_path)
    except Exception as e:
        return [f"模板无法打开: {template_path} ({e}) → 用 PowerPoint 另存修复"], []
    n = len(prs.slides)
    infos.append(f"模板可打开, 共 {n} 页幻灯片")
    if n < 2:
        problems.append(
            f"模板仅 {n} 页 → 本模板契约需要 ≥2 页 (封面/网格统计), 请核对模板"
        )
        return problems, infos

    def slide_text(idx):
        texts = []
        for shape in prs.slides[idx].shapes:
            if shape.has_text_frame:
                texts.append(shape.text_frame.text)
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    for cell in row.cells:
                        texts.append(cell.text)
        return "\n".join(texts)

    s1, s2 = slide_text(0), slide_text(1)
    for marker in MARKERS_SLIDE1:
        if marker not in s1:
            problems.append(
                f"模板第 1 页缺少标记文字「{marker}」→ 封面日期/版本回填依赖它, 请核对模板"
            )
    if not any(m in s2 for m in MARKERS_SLIDE2):
        problems.append(
            f"模板第 2 页缺少标记文字 {'/'.join(MARKERS_SLIDE2)}"
            " → 网格统计回填依赖它, 请核对模板"
        )
    return problems, infos


def main():
    argv = sys.argv[1:]
    if "--config" in argv:
        i = argv.index("--config")
        # --config 在末位且没给值时, 直接取 argv[i+1] 会 IndexError
        if i + 1 < len(argv) and not argv[i + 1].startswith("--"):
            config_path = argv[i + 1]
        else:
            print("[ERROR] --config 需要一个配置文件路径参数")
            sys.exit(2)
    else:
        config_path = os.path.join(REPO_ROOT, "report_config.json")
    all_problems = []

    ok, py_infos = check_python()
    for line in py_infos:
        print(f"[INFO] {line}")

    problems, dep_infos = check_dependencies()
    for line in dep_infos:
        print(f"[INFO] {line}")
    all_problems += problems

    problems, cfg_infos = check_config(config_path)
    for line in cfg_infos:
        print(f"[INFO] {line}")
    all_problems += problems

    if not problems:
        try:
            with open(config_path, "r", encoding="utf-8-sig") as f:
                template_path = json.load(f).get("template_pptx")
        except Exception:
            template_path = None
        problems, tpl_infos = check_template(template_path)
        for line in tpl_infos:
            print(f"[INFO] {line}")
        all_problems += problems

    if all_problems:
        print(f"\n[FAIL] 环境检查未通过 ({len(all_problems)} 项):")
        for p in all_problems:
            print(f"  - {p}")
        return 1
    print("\n[SUCCESS] 环境检查全部通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
