import io
import os
import sys
import json
import datetime
import argparse
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt, Inches
from PIL import Image, ImageDraw, ImageFont

try:
    from core.image_processor import (
        annotate_xy_curves,
        load_peaks,
        process_all_mode_b_plots,
        resolve_cover_image,
        trim_white_borders,
        load_truetype_font,
    )
except ImportError:
    try:
        from image_processor import (
            annotate_xy_curves,
            load_peaks,
            process_all_mode_b_plots,
            resolve_cover_image,
            trim_white_borders,
            load_truetype_font,
        )
    except ImportError:
        annotate_xy_curves = None
        load_peaks = None
        process_all_mode_b_plots = None
        resolve_cover_image = None
        trim_white_borders = None
        load_truetype_font = None

# 项目根 = core/ 的父目录 (与 config_gui.py / AutoReport.vbs 同级)。
# 所有相对路径 (模板/输出/数据目录) 一律相对项目根解析, 保证整目录拷贝到任何机器可运行。
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def resolve_template_path(value, project_root=PROJECT_ROOT):
    """解析模板 PPTX 路径: 绝对路径原样; 相对路径相对项目根拼接。

    解析后不存在 → 抛 FileNotFoundError (带实际尝试路径, 绝不静默回退)。
    """
    if not value:
        raise FileNotFoundError("template_pptx 配置为空, 未指定模板文件")
    candidate = str(value)
    if not os.path.isabs(candidate):
        candidate = os.path.join(project_root, candidate)
    if not os.path.exists(candidate):
        raise FileNotFoundError(f"Template PPTX not found at: {candidate}")
    return candidate


def resolve_output_dir(value, project_root=PROJECT_ROOT):
    """解析输出目录: 空/非法 → 项目根 (脚本同级目录); 相对路径 → 相对项目根。

    返回前确保目录存在 (makedirs), 失败异常上抛。
    """
    out_dir = str(value).strip() if value else ""
    if not out_dir:
        out_dir = project_root
    elif not os.path.isabs(out_dir):
        out_dir = os.path.join(project_root, out_dir)
    os.makedirs(out_dir, exist_ok=True)
    return out_dir


def resolve_output_dir_for_run(config, manifest, project_root=PROJECT_ROOT):
    """本次运行的输出目录决策:
    1. config.output_dir 显式配置优先 (相对路径相对项目根解析);
    2. 留空 → 输出到本次模型所在目录 (manifest.model_dir, 由 VBS 从
       Synergy.Project().Path 写入; 目录必须真实存在才采用);
    3. 均不可用 → 项目根 (兼容旧 manifest / 冒烟运行)。
    """
    explicit = str(config.get("output_dir") or "").strip()
    if explicit:
        return resolve_output_dir(explicit, project_root)
    model_dir = str(manifest.get("model_dir") or "").strip()
    if model_dir and os.path.isdir(model_dir):
        return model_dir
    return resolve_output_dir("", project_root)


# 数据缺失的统一呈现：宁可明确标注，绝不用任何 fabricated 默认值充数
MISSING_TEXT = "数据缺失"
MESH_SHORT = "—"

# 本次构建过程中登记的缺失数据字段（build_report 开头清空）
MISSING_FIELDS = []


def record_missing(label):
    if label not in MISSING_FIELDS:
        MISSING_FIELDS.append(label)


def enforce_missing_policy(missing, policy):
    """on_missing_data 策略: annotate 放行; fail 且存在缺失 → RuntimeError。"""
    if policy == "fail" and missing:
        raise RuntimeError("数据缺失 (on_missing_data=fail): " + "; ".join(missing))
    return None


def clamp_label_kind(text):
    """
    S9 锁模力表格标签识别 (精确匹配, 防误伤):
    只有短的标签单元格 ('CAE最大锁模力'/'注塑机最大锁模力') 返回 'cae'/'machine';
    '分析要求'等长说明文本即使包含关键词也返回 None
    (实发事故: 359.8T 被写进'分析要求'单元格)。
    """
    t = (text or "").strip()
    if len(t) > 12:
        return None
    if t == "CAE最大锁模力":
        return "cae"
    if t == "注塑机最大锁模力":
        return "machine"
    return None


def clear_stale_conclusion_rows(prs):
    """
    清空模板残留的'结果说明'类结论行 (值单元格)。
    模板是一份已填写的上个产品报告, 其结果说明 ('缩痕深度>0.03mm，本产品缩痕可见。'
    '体积收缩分布均匀。' 等) 是对上个产品结果的判断, 对本次产品即为臆造结论 → 清空待填。
    '分析要求' 行是通用验收标准 (与产品无关) → 保留。
    返回清空的行数。
    """
    cleared = 0
    for s_idx, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if not shape.has_table:
                continue
            tbl = shape.table
            for row in tbl.rows:
                label = row.cells[0].text.strip()
                is_conclusion = label == "结果说明" or (
                    label.endswith("说明") and 0 < len(label) <= 8
                )
                if not is_conclusion:
                    continue
                for c_idx in range(1, len(row.cells)):
                    if row.cells[c_idx].text.strip():
                        row.cells[c_idx].text = ""
                        cleared += 1
                print(
                    f"[Slide {s_idx}] Cleared stale template conclusion row: {label} (残留旧产品结论)"
                )
    return cleared


def fill_material_molding_range(slide, data_dir):
    """S8 '材料推荐成型范围' 单元格: 用本次真实材料的模温/熔温推荐范围回填。"""
    info_path = os.path.join(data_dir, "material_info.json")
    d = {}
    if os.path.exists(info_path):
        try:
            with open(info_path, "r", encoding="utf-8-sig", errors="ignore") as f:
                d = json.load(f)
        except Exception as e:
            print(f"[Notice] Error reading material_info.json: {e}")
    if not isinstance(d, dict):
        d = {}

    def num(key):
        try:
            v = float(d.get(key))
            return v if v >= 0 else None
        except (TypeError, ValueError):
            return None

    mold_min, mold_max = num("mold_temp_min"), num("mold_temp_max")
    melt_min, melt_max = num("melt_temp_min"), num("melt_temp_max")
    if not (
        mold_min is not None
        and mold_max is not None
        and melt_min is not None
        and melt_max is not None
    ):
        record_missing("材料: 推荐成型范围 (S8 表格)")
        return
    text = f"模温 {mold_min:g}~{mold_max:g}℃, 熔温 {melt_min:g}~{melt_max:g}℃"
    for shape in slide.shapes:
        if not shape.has_table:
            continue
        tbl = shape.table
        for row in tbl.rows:
            for c_idx, cell in enumerate(row.cells):
                if cell.text.strip() == "材料推荐成型范围" and c_idx + 1 < len(
                    row.cells
                ):
                    update_table_cell(row.cells[c_idx + 1], text, "微软雅黑", 12)
                    print(
                        f"[Slide 8] Filled 材料推荐成型范围 from real material: {text}"
                    )
                    return


# 材料字段描述 → 报告键 映射 (按顺序首个命中生效; 大小写不敏感, 中英描述均覆盖)。
# 字段 ID 由 VBS 通过官方 GetFirstField/GetNextField 枚举 + GetFieldDescription 描述导出,
# 绝不硬编码字段 ID (旧代码 1999/1998 等无出处猜测是材料文本字段全空的根因)。
MATERIAL_FIELD_RULES = [
    ("family_name", ("系列", "family")),
    ("trade_name", ("牌号", "贸易名称", "商品名", "trade")),
    ("manufacturer", ("制造商", "厂家", "manuf")),
    ("abbreviation", ("缩写", "abbrev")),
    ("material_type", ("材料类型", "material type")),
    ("data_source", ("数据来源", "source")),
    ("date_tested", ("测试日期", "date test")),
    ("date_modified", ("上次修改", "修改日期", "modif")),
    ("data_status", ("数据状态", "status")),
    ("grade_code", ("等级代码", "grade")),
    ("supplier_code", ("供应商代码", "supplier")),
    ("fiber_filler", ("填充", "fiber", "filler")),
    ("mold_temp_range", ("模具温度范围", "mold temperature")),
    ("mold_temp_rec", ("模具表面温度", "推荐模具温度", "mold surface")),
    ("melt_temp_range", ("熔体温度范围", "melt temperature")),
    ("melt_temp_max_abs", ("绝对最大熔体", "absolute max")),
    ("melt_temp_rec", ("熔体温度", "melt temp")),
    ("ejection_temp", ("顶出温度", "ejection")),
    ("max_shear_stress", ("剪切应力", "shear stress")),
    ("max_shear_rate", ("剪切速率", "shear rate")),
]

# mold_temp_range / melt_temp_range 值为数对 [min, max] — Moldflow 材料对话框的
# 推荐范围字段; 仅当解析出两个数值时才拆分, 单值/文本不猜。
RANGE_KEYS = {
    "mold_temp_range": ("mold_temp_min", "mold_temp_max"),
    "melt_temp_range": ("melt_temp_min", "melt_temp_max"),
}

# 已知 ID 回退 (第二优先级): 描述映射未命中时, 按字段 ID 补数值字段。
# 收录标准: 实机输出经交叉验证 — 通用 PP 与 Novodur HH-106 (ABS) 两种材料的同 ID
# 取值均落在各自推荐区间内 (2026-09-08 实机取证 temp/probe5_fields.log):
#   1808→模温范围 (20|80 / 60|80), 1800→熔温范围 (180|260 / 230|260),
#   1504→顶出温度 (124 / 105), 1804→最大剪切应力 (0.25 / 0.5),
#   1806→最大剪切速率 (100000 / 60000), 1801→绝对最大熔体温度 (300 / 300,
#   双材料一致高于推荐上限的安全天花板, 旧"越线存疑"判定有误, 现予采信),
#   11002→推荐熔体温度 (220 / 250), 11108→推荐模具温度 (50 / 70)。
LEGACY_FIELD_IDS = {
    1808: "mold_temp_range",
    1800: "melt_temp_range",
    1807: "mold_temp_rec",
    1801: "melt_temp_max_abs",
    11002: "melt_temp_rec",
    11108: "mold_temp_rec",
    1504: "ejection_temp",
    1804: "max_shear_stress",
    1806: "max_shear_rate",
}

# 已知 ID 回退 (第三优先级, 文本字段): GetFieldDescription 在 Moldflow 2023 实机
# 不支持 (枚举 128 字段描述全空, 取证 2026-09-08), 描述映射通道失效 → 文本字段
# 按项目原始字段表 ID 回退。仅接受非空且非纯数字的文本 (错位 ID 塞不进数值),
# 每次填充都写入审计日志; 待用户指定真实商料后按 material_fields.json 转储逐项核实锁定。
LEGACY_TEXT_IDS = {
    1999: "family_name",
    1998: "trade_name",
    1997: "manufacturer",
    1996: "data_source",
    1995: "material_type",
    1994: "abbreviation",
    1991: "date_tested",
    1990: "date_modified",
    1989: "data_status",
    1988: "grade_code",
    1987: "supplier_code",
    20031: "fiber_filler",
}


def _field_values_to_list(raw):
    """VBS 导出的 values 为 '|' 分隔字符串 → 数值列表 (非数值段忽略)。"""
    if raw is None:
        return []
    parts = str(raw).split("|")
    vals = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        try:
            vals.append(float(p))
        except ValueError:
            continue
    return vals


def build_material_info(data_dir):
    """
    material_fields.json (VBS 官方字段枚举的原始导出) → material_info.json。
    三级映射, 全部来自本次导出的真实数据, 绝不猜值:
    1. 描述关键字匹配 (GetFieldDescription 官方描述; 2023 实机不支持 → 描述全空, 通道自然跳过);
    2. 已知数值 ID 回退 (LEGACY_FIELD_IDS, 仅限上一版本实机验证过取值范围的 6 个数值字段);
    3. 已知文本 ID 回退 (LEGACY_TEXT_IDS, 仅接受非空非纯数字文本, 逐项审计日志)。
    material_fields.json 不存在时删除陈旧 material_info.json (防上个产品数据漏入)。
    返回写入的 dict。
    """
    src = os.path.join(data_dir, "material_fields.json")
    dst = os.path.join(data_dir, "material_info.json")
    if not os.path.exists(src):
        if os.path.exists(dst):
            os.remove(dst)
            print("[Material] material_fields.json 缺失, 已删除陈旧 material_info.json")
        return {}

    try:
        with open(src, "r", encoding="utf-8-sig", errors="ignore") as f:
            raw = json.load(f)
    except Exception as e:
        print(f"[Material] Error reading material_fields.json: {e}")
        if os.path.exists(dst):
            os.remove(dst)
        return {}
    if not isinstance(raw, dict):
        if os.path.exists(dst):
            os.remove(dst)
        return {}

    # material_id: VBS 导出的材料属性 ID (本方案属性表内, 经求解日志锁定真材料);
    # 无真实值则留空 → 报告标缺失
    mat_id_raw = str(raw.get("material_id", "") or "").strip()
    if mat_id_raw in ("-1", "0"):
        mat_id_raw = ""
    info = {
        "data_complete": bool(mat_id_raw),
        "material_id": mat_id_raw,
    }
    prop_name = str(raw.get("material_name", "") or "") or str(
        raw.get("prop_name", "") or ""
    )

    # 材料全名 "牌号 : 制造商" (如 "Novodur HH-106 : INEOS Styrolution") 拆分;
    # VBS 已拆分导出 trade_name/manufacturer 时优先用之 (取值源头一致, 不重复推断)
    if not info.get("trade_name") and prop_name:
        parts = prop_name.rsplit(" : ", 1)
        if len(parts) == 2 and parts[0].strip() and parts[1].strip():
            info["trade_name"] = parts[0].strip()
            info["manufacturer"] = parts[1].strip()
    for key, raw_key in (
        ("trade_name", "trade_name"),
        ("manufacturer", "manufacturer"),
    ):
        v = str(raw.get(raw_key, "") or "").strip()
        if v and not info.get(key):
            info[key] = v

    def apply_key(key, raw_vals, route):
        if key in RANGE_KEYS:
            vals = _field_values_to_list(raw_vals)
            if len(vals) == 2:
                k_min, k_max = RANGE_KEYS[key]
                info[k_min] = vals[0]
                info[k_max] = vals[1]
                return True
            return False
        if key in (
            "family_name",
            "trade_name",
            "manufacturer",
            "abbreviation",
            "material_type",
            "data_source",
            "date_tested",
            "date_modified",
            "data_status",
            "grade_code",
            "supplier_code",
            "fiber_filler",
        ):
            text = str(raw_vals).split("|")[0].strip()
            if text:
                info[key] = text
                return True
            return False
        vals = _field_values_to_list(raw_vals)
        if vals:
            info[key] = vals[0]
            return True
        return False

    fields = [f for f in raw.get("fields", []) if isinstance(f, dict)]
    desc_hit, legacy_hit = [], []
    for field in fields:
        desc_l = str(field.get("desc", "") or "").lower()
        raw_vals = field.get("values", "")
        if not desc_l:
            continue
        for key, keywords in MATERIAL_FIELD_RULES:
            if any(kw.lower() in desc_l for kw in keywords):
                if key not in info and apply_key(key, raw_vals, "desc"):
                    desc_hit.append(key)
                break  # 一个字段只映射到第一个命中的键

    # 第二级: 已知 ID 回退 (仅填描述映射未命中的键)
    for field in fields:
        try:
            fid = int(field.get("id"))
        except (TypeError, ValueError):
            continue
        key = LEGACY_FIELD_IDS.get(fid)
        if key is None or key in info or key in desc_hit:
            continue
        if apply_key(key, field.get("values", ""), "legacy"):
            legacy_hit.append(key)

    # 第三级: 文本字段已知 ID 回退 (仅非空非纯数字文本, 带审计)
    for field in fields:
        try:
            fid = int(field.get("id"))
        except (TypeError, ValueError):
            continue
        key = LEGACY_TEXT_IDS.get(fid)
        if key is None or info.get(key):
            continue
        text = str(field.get("values", "") or "").split("|")[0].strip()
        if not text or text.replace(".", "").replace("-", "").isdigit():
            continue
        info[key] = text
        legacy_hit.append(f"{key}(id={fid})")

    if desc_hit or legacy_hit:
        print(
            f"[Material] Mapped fields via desc: {desc_hit}; via known-id fallback: {legacy_hit}"
        )

    # 官方 Prop.Name 作为牌号兜底 (真实属性名, 非臆造)
    if not info.get("trade_name") and prop_name:
        info["trade_name"] = prop_name

    try:
        with open(dst, "w", encoding="utf-8") as f:
            json.dump(info, f, ensure_ascii=False, indent=2)
        print(
            f"[Material] Mapped material_fields.json -> material_info.json "
            f"({sum(1 for k, v in info.items() if v not in ('', None))} 项有效)"
        )
    except Exception as e:
        print(f"[Material] Could not write material_info.json: {e}")
    return info


def resolve_clamp_force(config, data_dir):
    """
    锁模力数据单源：CAE 计算值只认 clamp_force_info.json（VBS 导出），
    机台吨位只认 config 的 machine_max_ton（用户录入的设备规格）。
    缺失返回 None，绝不回退到任何硬编码吨位。
    """
    cf_cfg = config.get("clamp_force_settings", {})
    cae_ton = None
    mach_ton = None
    cf_info_path = os.path.join(data_dir, "clamp_force_info.json")
    if os.path.exists(cf_info_path):
        try:
            with open(cf_info_path, "r", encoding="utf-8-sig") as f:
                cae_ton = json.load(f).get("cae_max_clamp_force")
            if not isinstance(cae_ton, (int, float)) or cae_ton <= 0:
                cae_ton = None
        except Exception as e:
            print(f"[Notice] Error reading clamp_force_info.json: {e}")
            cae_ton = None
    mach_val = cf_cfg.get("machine_max_ton")
    if isinstance(mach_val, (int, float)) and mach_val > 0:
        mach_ton = mach_val
    return cae_ton, mach_ton


def resolve_peaks(peaks, config=None):
    """
    XY 探针峰值单源解析: 只认本次导出数据 (peak_values.json / *_curve_data.txt,
    由 VBS 从当前方案曲线导出)。值或时刻任一缺失 → None (跳过探针标注)。
    config 手工值回退已删除: 换产品后陈旧 config 数值会冒充本次结果 (实发事故)。
    """
    out = {}
    peaks = peaks or {}
    for curve in ("clamp_force", "inj_pressure"):
        entry = peaks.get(curve) or {}
        pv = entry.get("peak_value")
        pt = entry.get("peak_time")
        if (
            isinstance(pv, (int, float))
            and pv > 0
            and isinstance(pt, (int, float))
            and pt > 0
        ):
            out[curve] = (float(pv), float(pt))
        else:
            out[curve] = None
    return out


def sanitize_filename(name, max_len=80):
    """
    文件名净化 (T15): 去除 Windows 非法字符与路径分隔符 (防 ..\\ 目录逃逸),
    压缩连续下划线/空白, 限制长度。保留中文可读性。
    """
    import re

    name = str(name)
    name = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", name)
    name = name.replace("..", "_")
    name = re.sub(r"[_\s]{2,}", "_", name).strip("_ ")
    if len(name) > max_len:
        name = name[:max_len].rstrip("_ ")
    return name or "方案"


def ensure_missing_plot_placeholder(data_dir):
    """
    生成/复用 '（图像缺失）' 占位图。缺图时替换模板里的旧图 blob,
    绝不让模板中上一个零件的截图混入本次报告。
    """
    path = os.path.join(data_dir, "_missing_plot.png")
    try:
        font = (
            load_truetype_font(36)
            if load_truetype_font is not None
            else ImageFont.load_default()
        )
        im = Image.new("RGB", (600, 400), "#f5f5f5")
        draw = ImageDraw.Draw(im)
        draw.rectangle([4, 4, 595, 395], outline="#b0b0b0", width=2)
        text = "（图像缺失）"
        tw = int(draw.textlength(text, font=font))
        draw.text(((600 - tw) // 2, 170), text, fill="#808080", font=font)
        im.save(path, "PNG")
    except Exception as e:
        print(f"[Notice] Could not create missing-plot placeholder: {e}")
    return path


# 各幻灯片自适应安全放置区域 (left, top, width, height, allow_cover_title)
# Slide 尺寸：10.00 x 7.50 英寸 (9144000 x 6858000 EMU)
SLIDE_SAFE_BOXES = {
    # 1: 封面纯 CAD 模型 (居中，避让顶部标题与底部表格)
    1: (Inches(2.00), Inches(0.50), Inches(6.00), Inches(4.45), False),
    # 2: 产品网格模型 (左侧安全区，避让顶部和右侧文本框)
    2: (Inches(0.12), Inches(1.35), Inches(5.15), Inches(5.35), False),
    # 4: GIF 充填动画 (锁定用户手动缩放及位置，严格保留右下角文本框避让)
    4: (109728, 970643, 8924544, 4642394, False),
    # 5: 压力 (避让右上标题与底部表格 top=6.69)
    5: (Inches(0.12), Inches(0.55), Inches(9.76), Inches(6.10), False),
    # 6: V/P 切换时的压力
    6: (Inches(0.12), Inches(0.55), Inches(9.76), Inches(6.10), False),
    # 7: 注射位置处压力:XY 图 (满幅放置，物理覆盖右上角标题，但绝不删除标题文本)
    7: (Inches(0.12), Inches(0.12), Inches(9.76), Inches(6.53), True),
    # 8: 流动前沿温度 (避让底部表格 top=6.15)
    8: (Inches(0.12), Inches(0.55), Inches(9.76), Inches(5.55), False),
    # 9: 锁模力:XY 图 (满幅放置，物理覆盖右上角标题，但绝不删除标题文本)
    9: (Inches(0.12), Inches(0.12), Inches(9.76), Inches(6.53), True),
    # 10: 熔接痕 (避让底部表格 top=6.69)
    10: (Inches(0.12), Inches(0.55), Inches(9.76), Inches(6.10), False),
    # 11: 气穴 (避让底部表格 top=6.69)
    11: (Inches(0.12), Inches(0.55), Inches(9.76), Inches(6.10), False),
    # 12: 体积收缩率 (避让底部表格 top=6.69)
    12: (Inches(0.12), Inches(0.55), Inches(9.76), Inches(6.10), False),
    # 13: 变形，所有效应:变形 (跟随自适应算法，与 14~16 保持一致)
    13: (Inches(0.12), Inches(0.55), Inches(9.76), Inches(6.70), False),
    # 14: 变形，所有因素:X 分量 (无底部遮挡，超大幅面)
    14: (Inches(0.12), Inches(0.55), Inches(9.76), Inches(6.70), False),
    # 15: 变形，所有因素:Y 分量 (无底部遮挡，超大幅面)
    15: (Inches(0.12), Inches(0.55), Inches(9.76), Inches(6.70), False),
    # 16: 变形，所有因素:Z 分量 (无底部遮挡，超大幅面)
    16: (Inches(0.12), Inches(0.55), Inches(9.76), Inches(6.70), False),
}

# Slide 3 材料四联图象限定义 (顺序对应 material_basic/process/viscosity/pvt)。
# 阈值 4500000/3500000 EMU 沿用模板四联图的历史分界; 插入框按阈值分区再内缩边距,
# 仅在对应象限没有任何既有图片形状时用于 add_picture。
MATERIAL_QUADRANT_KEYS = (
    "material_basic",
    "material_process",
    "material_viscosity",
    "material_pvt",
)
MATERIAL_QUADRANT_LABELS = ("材料基本信息", "推荐工艺", "粘度曲线", "PVT 曲线")
MATERIAL_QUADRANT_X_SPLIT = 4500000
MATERIAL_QUADRANT_Y_SPLIT = 3500000
MATERIAL_QUADRANT_BOXES = (
    (150000, 400000, 4200000, 2950000),  # 左上
    (4650000, 400000, 4300000, 2950000),  # 右上
    (150000, 3650000, 4200000, 3000000),  # 左下
    (4650000, 3650000, 4300000, 3000000),  # 右下
)


def material_quadrant_index(x, y):
    """按形状左上角坐标判定 Slide 3 四联图象限 (0=左上 1=右上 2=左下 3=右下)。"""
    if y < MATERIAL_QUADRANT_Y_SPLIT:
        return 0 if x < MATERIAL_QUADRANT_X_SPLIT else 1
    return 2 if x < MATERIAL_QUADRANT_X_SPLIT else 3


def _encode_matching_format(new_image_path, target_ext):
    """按目标部件扩展名编码图片字节, 杜绝 'png 部件装 JPEG 字节' 的格式错配
    (该错配会导致部分 PowerPoint/WPS 版本封面图无法显示)。
    扩展名一致 → 原字节直通; 不一致 → 用 PIL 重编码为目标格式。
    GIF 部件只接受 GIF 源 (动画重编码会丢帧, 不一致直接拒绝)。
    """
    src_ext = os.path.splitext(new_image_path)[1].lower()
    target_ext = (target_ext or "").lower()
    if target_ext in (".jpg", ".jpe"):
        target_ext = ".jpeg"
    if src_ext == target_ext:
        with open(new_image_path, "rb") as f:
            return f.read(), None
    if target_ext == ".gif":
        return None, f"GIF 部件不接受非 GIF 源: {new_image_path}"
    try:
        with Image.open(new_image_path) as im:
            im = im.convert("RGB")
            buf = io.BytesIO()
            if target_ext == ".jpeg":
                im.save(buf, "JPEG", quality=92)
            else:
                im.save(buf, "PNG")
            return buf.getvalue(), None
    except Exception as e:
        return None, f"重编码为 {target_ext} 失败 ({new_image_path}): {e}"


def replace_picture_blob(pic_shape, new_image_path):
    """
    通过更新 SlidePart 中 Blip 关联的 ImagePart._blob，
    替换图片二进制内容。字节格式始终与部件扩展名保持一致。
    """
    if not os.path.exists(new_image_path):
        print(
            f"[Warning] Image file not found: {new_image_path}, skipping replacement."
        )
        return False

    try:
        rId = pic_shape._element.blip_rId
        rel = pic_shape.part.rels[rId]
        image_part = rel.target_part
        target_ext = os.path.splitext(str(image_part.partname))[1]

        new_blob, enc_err = _encode_matching_format(new_image_path, target_ext)
        if new_blob is None:
            print(
                f"[Warning] Failed to replace picture blob for {pic_shape.name}: {enc_err}"
            )
            return False
        image_part._blob = new_blob

        # 更新 content_type (与部件扩展名一致)
        if target_ext.lower() in (".jpg", ".jpeg", ".jpe"):
            image_part._content_type = "image/jpeg"
        elif target_ext.lower() == ".gif":
            image_part._content_type = "image/gif"
        else:
            image_part._content_type = "image/png"

        print(
            f"  [OK] Replaced blob of {pic_shape.name} with {os.path.basename(new_image_path)}"
        )
        return True
    except Exception as e:
        print(f"[Warning] Failed to replace picture blob for {pic_shape.name}: {e}")
        return False


def prepare_cover_jpeg(solid_img):
    """封面体积预算: PNG 转 JPEG q90。派生缓存每次无条件重生成 —
    旧的 exists 短路曾让上个产品的封面图冒充本次封面 (实发事故)。"""
    jpg_path = os.path.splitext(solid_img)[0] + "_cover.jpg"
    with Image.open(solid_img) as im_s:
        im_s.convert("RGB").save(jpg_path, "JPEG", quality=90)
    print(
        f"[Slide 1] 封面转 JPEG 控制体积 (本次重新生成): {os.path.basename(jpg_path)}"
    )
    return jpg_path


def compute_safe_box_fit(new_image_path, box_left, box_top, box_w, box_h):
    """计算图片等比适配安全框并居中的几何 (left, top, width, height)。
    1. 用 PIL 获取图片物理分辨率, scale = min(box_w/img_w, box_h/img_h) 严格锁定
       原生宽高比, 杜绝挤压变形; 计算结果居中于安全框内。
    2. 读图失败返回 None (调用方自行决定回退策略)。
    """
    try:
        with Image.open(new_image_path) as im:
            img_w, img_h = im.size
    except Exception as e:
        print(f"[Warning] Failed to read image size for {new_image_path}: {e}")
        return None
    scale = min(float(box_w) / float(img_w), float(box_h) / float(img_h))
    target_w = int(img_w * scale)
    target_h = int(img_h * scale)
    target_left = int(box_left + (box_w - target_w) / 2)
    target_top = int(box_top + (box_h - target_h) / 2)
    return target_left, target_top, target_w, target_h


def fit_picture_in_safe_box(
    slide,
    pic_shape,
    new_image_path,
    box_left,
    box_top,
    box_w,
    box_h,
    allow_cover_title=False,
):
    """
    自适应安全边界等比放大放置图片（几何由 compute_safe_box_fit 计算）：
    动态更新 pic_shape.left, pic_shape.top, pic_shape.width, pic_shape.height,
    再调用 replace_picture_blob 替换图片数据。
    注意：严格保留右上角标题文本框文字，绝不进行清空处理（允许图片自然覆盖或在旁展示）。
    """
    if not os.path.exists(new_image_path):
        return False

    geom = compute_safe_box_fit(new_image_path, box_left, box_top, box_w, box_h)
    if geom is None:
        return replace_picture_blob(pic_shape, new_image_path)

    pic_shape.left, pic_shape.top, pic_shape.width, pic_shape.height = geom

    # 替换图像二进制
    replace_picture_blob(pic_shape, new_image_path)
    return True


def get_main_picture(slide):
    """
    获取幻灯片中面积最大的主图片（防止误选角落的小坐标指示图标）
    """
    pics = [s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]
    if not pics:
        return None
    # 面积从大到小排序
    pics.sort(key=lambda s: s.width * s.height, reverse=True)
    return pics[0]


def insert_or_replace_picture(
    slide,
    new_image_path,
    box_left,
    box_top,
    box_w,
    box_h,
    allow_cover_title=False,
    log_tag="",
):
    """把结果图放到页面上：页面已有图片 → 替换面积最大的主图 (保留模板布局);
    页面无任何图片形状 → 按安全框等比居中插入新图。
    模板图片被用户清空时同样出图, 不再静默丢图 (清空模板导出无图回归)。
    返回 False 仅当图片放置失败 (文件缺失/读图失败), 供调用方登记缺失。
    """
    main_pic = get_main_picture(slide)
    if main_pic is not None:
        return fit_picture_in_safe_box(
            slide,
            main_pic,
            new_image_path,
            box_left,
            box_top,
            box_w,
            box_h,
            allow_cover_title,
        )
    geom = compute_safe_box_fit(new_image_path, box_left, box_top, box_w, box_h)
    if geom is None:
        return False
    slide.shapes.add_picture(new_image_path, *geom)
    print(
        f"{log_tag} 页面无图片形状, 已按安全框插入新图: "
        f"{os.path.basename(new_image_path)}"
    )
    return True


def update_table_cell(cell, text, font_name="微软雅黑", font_size_pt=12):
    """
    更新表格单元格文字并锁定字体格式（严格保留原模板东亚主题字体 +mn-ea）
    """
    if not cell.text_frame.paragraphs:
        cell.text = text
        return
    p = cell.text_frame.paragraphs[0]
    p.space_before = Pt(0)
    p.space_after = Pt(0)
    if p.runs:
        # 直接原位更新已有 run 的文字，100% 完整继承原模板的 <a:rPr>、<a:latin typeface="+mn-ea"/>、<a:ea typeface="+mn-ea"/> 及颜色属性
        p.runs[0].text = text
        for extra_r in p.runs[1:]:
            extra_r.text = ""
    else:
        r = p.add_run()
        r.text = text
        r.font.name = font_name
        r.font.size = Pt(font_size_pt)


def set_text_frame_keep_text_only(shape, lines, font_name="Arial", font_size_pt=8):
    """
    实现“只粘贴文本 (Keep Text Only)”效果并保留模板原始空行结构与排版：
    如果段落数量与目标行数匹配，原位逐行更新文本；若不足则追加，多出则清空。
    严格保持模板原生空行与段落格式，空行分明，杜绝挤压成一团。
    """
    tf = shape.text_frame
    tf.word_wrap = True

    for idx, line in enumerate(lines):
        if idx < len(tf.paragraphs):
            p = tf.paragraphs[idx]
            p.text = line
            for r in p.runs:
                r.font.name = font_name
                r.font.size = Pt(font_size_pt)
        else:
            p = tf.add_paragraph()
            p.text = line
            for r in p.runs:
                r.font.name = font_name
                r.font.size = Pt(font_size_pt)

    if len(tf.paragraphs) > len(lines):
        for p in tf.paragraphs[len(lines) :]:
            p.text = ""


def get_formatted_mesh_text(data_dir):
    """
    读取 mesh_summary.json，输出贴合模板的网格统计文本。
    字段缺失/损坏/null/负数哨兵 → 该字段渲染 '—' 占位并登记缺失，
    绝不使用任何 fabricated 默认值。返回 (lines, missing_labels)。
    """
    json_path = os.path.join(data_dir, "mesh_summary.json")
    d = {}
    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8-sig", errors="ignore") as f:
                d = json.load(f)
        except Exception as e:
            print(f"[Notice] Error reading mesh_summary.json: {e}")
            d = {}
    if not isinstance(d, dict):
        d = {}

    missing = []

    def get_val(key, cast, label):
        v = d.get(key)
        if v is None:
            missing.append(label)
            return None
        try:
            v = cast(v)
        except (ValueError, TypeError):
            missing.append(label)
            return None
        if v < 0:
            missing.append(label)
            return None
        return v

    triangles = get_val("triangles", int, "网格统计: 三角形")
    nodes = get_val("nodes", int, "网格统计: 已连接的节点")
    regions = get_val("connectivity_regions", int, "网格统计: 连通区域")
    unvis = get_val("unvisible_triangles", int, "网格统计: 不可见三角形")
    surf_val = get_val("surface_area", float, "网格统计: 表面面积")
    vol_val = get_val("volume", float, "网格统计: 体积")
    max_ar_val = get_val("max_aspect_ratio", float, "网格统计: 最大纵横比")
    ave_ar_val = get_val("ave_aspect_ratio", float, "网格统计: 平均纵横比")
    min_ar_val = get_val("min_aspect_ratio", float, "网格统计: 最小纵横比")
    free_edges = get_val("free_edges", int, "网格统计: 自由边")
    manifold_edges = get_val("manifold_edges", int, "网格统计: 共用边")
    non_manifold_edges = get_val("non_manifold_edges", int, "网格统计: 多重边")
    unoriented = get_val("unoriented", int, "网格统计: 配向不正确的单元")
    intersection = get_val("intersection_elements", int, "网格统计: 相交单元")
    overlap = get_val("overlap_elements", int, "网格统计: 完全重叠单元")
    m_val = get_val("match_ratio", float, "网格统计: 匹配百分比")
    r_val = get_val("reciprocal_match_ratio", float, "网格统计: 相互百分比")

    def num_or_dash(val, pattern):
        return pattern.format(val) if val is not None else MESH_SHORT

    surf_area = num_or_dash(surf_val, "{:.2f}")
    vol = num_or_dash(vol_val, "{:.3f}")
    max_ar = num_or_dash(max_ar_val, "{:.2f}")
    ave_ar = num_or_dash(ave_ar_val, "{:.2f}")
    min_ar = num_or_dash(min_ar_val, "{:.2f}")

    # 匹配率百分比归一化 (0<x<=1 视为小数比例); 0/缺失视为读数失败
    if m_val is not None and m_val > 0:
        if m_val <= 1.0:
            m_val = m_val * 100.0
        match_ratio = f"{m_val:.1f}%"
    else:
        match_ratio = MESH_SHORT
        if m_val is not None:
            missing.append("网格统计: 匹配百分比")
    if r_val is not None and r_val > 0:
        if r_val <= 1.0:
            r_val = r_val * 100.0
        reciprocal_ratio = f"{r_val:.1f}%"
    else:
        reciprocal_ratio = MESH_SHORT
        if r_val is not None:
            missing.append("网格统计: 相互百分比")

    mesh_type = d.get("mesh_type")
    if mesh_type and "dual" in str(mesh_type).lower():
        type_str = "Dual Domain"
    elif mesh_type and "mid" in str(mesh_type).lower():
        type_str = "Midplane"
    elif mesh_type:
        type_str = "3D"
    else:
        missing.append("网格统计: 网格类型")
        type_str = None

    def int_or_dash(val):
        return val if val is not None else MESH_SHORT

    lines = [
        "三角形 ",
        "----------------------------------------",
        "实体计数:",
        f"    三角形               {int_or_dash(triangles)}",
        f"    已连接的节点         {int_or_dash(nodes)}",
        f"    连通区域             {int_or_dash(regions)}",
        "",
        f"    不可见三角形            {int_or_dash(unvis)}",
        "",
        "面积:",
        "(不包括模具镶块和冷却管道)",
        f"    表面面积: \t{surf_area} cm^2",
        "",
        "按单元类型统计的体积:",
        f"    三角形： \t{vol} cm^3",
        "",
        "纵横比:",
        "    最大       平均        最小",
        f"     {max_ar:>5}       {ave_ar:>5}       {min_ar:>5}",
        "",
        "边细节:",
        f"    自由边\t\t\t {int_or_dash(free_edges)}",
        f"    共用边        \t\t {int_or_dash(manifold_edges)}",
        f"    多重边            \t\t {int_or_dash(non_manifold_edges)}",
        "",
        "取向细节:",
        f"    配向不正确的单元     \t {int_or_dash(unoriented)}",
        "",
        "交叉点细节:",
        f"    相交单元             \t {int_or_dash(intersection)}",
        f"    完全重叠单元              \t {int_or_dash(overlap)}",
        "",
        "匹配百分比:",
        f"    匹配百分比      \t\t {match_ratio}",
        f"    相互百分比           \t {reciprocal_ratio}",
        "",
    ]
    if type_str is not None:
        lines.append(f"适合 {type_str} 分析。")
    else:
        lines.append("数据不完整，结论仅供参考。")
    return lines, missing


def render_material_dialog_cards(data_dir, manifest=None):
    """
    100% 还原 Autodesk Moldflow 官方热塑性材料属性对话框与推荐工艺参数：
    1. material_basic.png (853 x 529): 完整 14 项基本属性 + 顶部选项卡栏（描述、推荐工艺等）
    2. material_process.png (555 x 453): 推荐工艺参数卡片（模具温度范围、熔体温度、红色告警参数）
    """
    if load_truetype_font is not None:
        f_tab = load_truetype_font(13)
        f_lbl = load_truetype_font(14)
        f_lbl_bold = load_truetype_font(14, bold=True)
        f_val = load_truetype_font(14)
        f_sub = load_truetype_font(13, bold=True)
    else:
        font_path = r"C:\Windows\Fonts\msyh.ttc"
        font_bold_path = (
            r"C:\Windows\Fonts\msyhbd.ttc"
            if os.path.exists(r"C:\Windows\Fonts\msyhbd.ttc")
            else font_path
        )
        f_tab = ImageFont.truetype(font_path, 13)
        f_lbl = ImageFont.truetype(font_path, 14)
        f_lbl_bold = ImageFont.truetype(font_bold_path, 14)
        f_val = ImageFont.truetype(font_path, 14)
        f_sub = ImageFont.truetype(font_bold_path, 13)

    # 读取 material_info.json
    info_file = os.path.join(data_dir, "material_info.json")
    d = {}
    if os.path.exists(info_file):
        try:
            with open(info_file, "r", encoding="utf-8-sig", errors="ignore") as f:
                d = json.load(f)
        except Exception as e:
            print(f"[Notice] Error reading material_info.json: {e}")
            d = {}
    if not isinstance(d, dict):
        d = {}

    missing = []

    def text_val(key, label):
        v = d.get(key)
        if v is None or (isinstance(v, str) and not v.strip()):
            missing.append(f"材料: {label}")
            return MISSING_TEXT
        return str(v)

    def num_val(key, label):
        v = d.get(key)
        if v is None:
            missing.append(f"材料: {label}")
            return None
        try:
            v = float(v)
        except (ValueError, TypeError):
            missing.append(f"材料: {label}")
            return None
        if v < 0:
            missing.append(f"材料: {label}")
            return None
        return v

    def num_str(val):
        return f"{val:g}" if val is not None else MISSING_TEXT

    trade_name = text_val("trade_name", "牌号")
    family_name = text_val("family_name", "系列")
    manufacturer = text_val("manufacturer", "制造商")
    abbrev = text_val("abbreviation", "材料名称缩写")
    mat_type = text_val("material_type", "材料类型")
    data_src = text_val("data_source", "数据来源")
    date_tested = text_val("date_tested", "测试日期")
    date_mod = text_val("date_modified", "上次修改日期")
    data_status = text_val("data_status", "数据状态")
    mat_id = text_val("material_id", "材料 ID")
    grade_code = text_val("grade_code", "等级代码")
    supp_code = text_val("supplier_code", "供应商代码")
    fiber = text_val("fiber_filler", "纤维/填充物")

    if manufacturer != MISSING_TEXT and trade_name != MISSING_TEXT:
        link = f"https://www.google.com/search?q={manufacturer}+{trade_name}"
    else:
        link = MISSING_TEXT

    mold_min = num_str(num_val("mold_temp_min", "模具温度范围最小值"))
    mold_max = num_str(num_val("mold_temp_max", "模具温度范围最大值"))
    mold_rec = num_str(num_val("mold_temp_rec", "模具表面温度"))
    melt_min = num_str(num_val("melt_temp_min", "熔体温度范围最小值"))
    melt_max = num_str(num_val("melt_temp_max", "熔体温度范围最大值"))
    melt_rec = num_str(num_val("melt_temp_rec", "熔体温度"))
    melt_abs = num_str(num_val("melt_temp_max_abs", "绝对最大熔体温度"))
    eject_t = num_str(num_val("ejection_temp", "顶出温度"))
    shear_s = num_str(num_val("max_shear_stress", "最大剪切应力"))
    shear_r = num_str(num_val("max_shear_rate", "最大剪切速率"))

    # ---------------- 1. 渲染基本属性卡片 (853 x 529) ----------------
    w1, h1 = 853, 529
    im1 = Image.new("RGB", (w1, h1), "#f8f9fa")
    draw1 = ImageDraw.Draw(im1)

    # 顶部标签栏
    tabs = [
        "描述",
        "推荐工艺",
        "流变属性",
        "热属性",
        "pvT 属性",
        "机械属性",
        "收缩属性",
        "填充物/纤维",
        "微孔发泡特性",
        "光学特性",
        "环境影响",
        "材料数...",
    ]
    tab_x = 8
    tab_h = 32
    draw1.line([(0, tab_h), (w1, tab_h)], fill="#dcdcdc", width=1)

    for i, t_name in enumerate(tabs):
        t_w = int(draw1.textlength(t_name, font=f_tab)) + 14
        if i == 0:
            # 当前选中的“描述”标签
            draw1.rectangle(
                [tab_x, 4, tab_x + t_w, tab_h], fill="#ffffff", outline="#dcdcdc"
            )
            draw1.line([tab_x, tab_h, tab_x + t_w, tab_h], fill="#ffffff", width=1)
            draw1.line([tab_x, 4, tab_x + t_w, 4], fill="#1890ff", width=2)
            draw1.text((tab_x + 7, 10), t_name, fill="#1890ff", font=f_lbl_bold)
        else:
            draw1.text((tab_x + 7, 10), t_name, fill="#495057", font=f_tab)
        tab_x += t_w + 4

    # 14 项完整基本参数
    basic_fields = [
        ("系列", family_name),
        ("牌号", trade_name),
        ("制造商", manufacturer),
        ("链接", link),
        ("材料名称缩写", abbrev),
        ("材料类型", mat_type),
        ("数据来源", data_src),
        ("上次修改日期", date_mod),
        ("测试日期", date_tested),
        ("数据状态", data_status),
        ("材料 ID", mat_id),
        ("等级代码", grade_code),
        ("供应商代码", supp_code),
        ("纤维/填充物", fiber),
    ]

    start_y = 44
    row_gap = 34
    box_x = 135
    box_w = w1 - box_x - 15

    for idx, (lbl, val) in enumerate(basic_fields):
        y = start_y + idx * row_gap
        draw1.text((16, y + 4), lbl, fill="#212529", font=f_lbl)
        draw1.rectangle(
            [box_x, y, box_x + box_w, y + 26],
            fill="#ffffff",
            outline="#ced4da",
            width=1,
        )
        v_color = "#c00000" if val == MISSING_TEXT else "#212529"
        draw1.text((box_x + 8, y + 4), str(val), fill=v_color, font=f_val)

    p1_path = os.path.join(data_dir, "material_basic.png")
    im1.save(p1_path, quality=98)

    # ---------------- 2. 渲染工艺参数卡片 (555 x 453) ----------------
    w2, h2 = 555, 453
    im2 = Image.new("RGB", (w2, h2), "#f8f9fa")
    draw2 = ImageDraw.Draw(im2)

    lbl_x2 = 18
    val_x2 = 200
    val_w2 = 290
    unit_x2 = 502

    def draw_proc_row(y, lbl, val, unit, is_red=False):
        color = "#c00000" if is_red else "#212529"
        f = f_lbl_bold if is_red else f_lbl
        draw2.text((lbl_x2, y + 4), lbl, fill=color, font=f)
        draw2.rectangle(
            [val_x2, y, val_x2 + val_w2, y + 26],
            fill="#ffffff",
            outline="#ced4da",
            width=1,
        )
        v_color = "#c00000" if str(val) == MISSING_TEXT else "#212529"
        draw2.text((val_x2 + 10, y + 4), str(val), fill=v_color, font=f_val)
        if str(val) != MISSING_TEXT:
            draw2.text((unit_x2, y + 4), unit, fill="#212529", font=f_lbl)

    cur_y = 18
    draw_proc_row(cur_y, "模具表面温度", mold_rec, "°C")
    cur_y += 38
    draw_proc_row(cur_y, "熔体温度", melt_rec, "°C")
    cur_y += 42

    # 分组 1: 模具温度范围(推荐)
    draw2.text((lbl_x2, cur_y), "模具温度范围(推荐)", fill="#495057", font=f_sub)
    draw2.line([(150, cur_y + 9), (w2 - 18, cur_y + 9)], fill="#dee2e6", width=1)
    cur_y += 24
    draw_proc_row(cur_y, "  最小值", mold_min, "°C")
    cur_y += 34
    draw_proc_row(cur_y, "  最大值", mold_max, "°C")
    cur_y += 42

    # 分组 2: 熔体温度范围(推荐)
    draw2.text((lbl_x2, cur_y), "熔体温度范围(推荐)", fill="#495057", font=f_sub)
    draw2.line([(150, cur_y + 9), (w2 - 18, cur_y + 9)], fill="#dee2e6", width=1)
    cur_y += 24
    draw_proc_row(cur_y, "  最小值", melt_min, "°C")
    cur_y += 34
    draw_proc_row(cur_y, "  最大值", melt_max, "°C")
    cur_y += 42

    draw_proc_row(cur_y, "绝对最大熔体温度", melt_abs, "°C")
    cur_y += 36
    draw_proc_row(cur_y, "顶出温度", eject_t, "°C")
    cur_y += 42

    # 红色告警参数
    draw_proc_row(cur_y, "最大剪切应力", shear_s, "MPa", is_red=True)
    cur_y += 36
    draw_proc_row(cur_y, "最大剪切速率", shear_r, "1/s", is_red=True)

    p2_path = os.path.join(data_dir, "material_process.png")
    im2.save(p2_path, quality=98)
    print(
        f"[Slide 3] Rendered authentic material cards: {os.path.basename(p1_path)} & {os.path.basename(p2_path)}"
    )
    return missing


def build_single_report(
    config,
    manifest,
    data_dir,
    template_path,
    today_str,
    study_name,
    part_name,
    mf_version,
    mode="B",
    output_path=None,
):
    """
    构建单一模式的 PPT 报告 (Mode A: 视口抓取; Mode B: 智能拼合)
    """
    mode = mode.upper()
    mode_tag = "方案A(视口抓取)" if mode == "A" else "方案B(智能拼合)"
    print(f"\n=======================================================")
    print(f"[PPTX Builder] 开始构建【{mode_tag}】报告...")
    print(f"=======================================================")

    prs = Presentation(template_path)
    total_slides = len(prs.slides)
    print(f"[PPTX Builder] 已载入模板，共 {total_slides} 页幻灯片。")

    # 模板残留旧产品结论清空 ('结果说明'等行是模板产品=上个产品的判断, 对本次即臆造)
    cleared_rows = clear_stale_conclusion_rows(prs)
    if cleared_rows:
        record_missing(
            f"结果说明: 模板残留结论已清空 {cleared_rows} 处 (待工程师按本次结果填写)"
        )

    # ------------------ Slide 1: 封面信息更新 (保持微软雅黑 12pt) ------------------
    if total_slides >= 1:
        s1 = prs.slides[0]
        for shape in s1.shapes:
            if shape.has_table:
                tbl = shape.table
                for row_idx, row in enumerate(tbl.rows):
                    for col_idx, cell in enumerate(row.cells):
                        text = cell.text.strip()
                        # 更新日期
                        if "报告日期" in text and col_idx + 1 < len(row.cells):
                            update_table_cell(
                                row.cells[col_idx + 1], today_str, "微软雅黑", 12
                            )
                        # 更新版本
                        if "Moldflow" in text and col_idx + 1 < len(row.cells):
                            update_table_cell(
                                row.cells[col_idx + 1], str(mf_version), "微软雅黑", 12
                            )
                        # 如果有真实零件名 (manifest 未提供时留空, 绝不写默认占位词)
                        if (
                            part_name
                            and part_name != "方案"
                            and "零件" in text
                            and "供应商" not in text
                            and col_idx + 1 < len(row.cells)
                        ):
                            if not row.cells[col_idx + 1].text.strip():
                                update_table_cell(
                                    row.cells[col_idx + 1], part_name, "微软雅黑", 12
                                )
        print("[Slide 1] Updated cover table info (严格保持微软雅黑 12pt 字体).")

        # 替换封面模型图 (纯模型本体截图，无网格、无节点)
        # 方案 A/B 共用同一条路径: 走 resolve_cover_image 选取最佳可用模型本体图
        # (优先 mode_b/model_pressure.png → mode_b/model_volumetric_shrinkage.png → solid_model.png,
        #  并对命中候选做非空/非割裂完整性校验 + 裁白边写回)。
        solid_img = None
        if resolve_cover_image is not None:
            try:
                solid_img = resolve_cover_image(data_dir)
            except Exception as e:
                print(f"[Slide 1] resolve_cover_image 调用异常: {e}")
                solid_img = None
        if not solid_img:
            solid_img = os.path.join(data_dir, "solid_model.png")
            if not os.path.exists(solid_img):
                cand_b = os.path.join(data_dir, "mode_b", "solid_model.png")
                if os.path.exists(cand_b):
                    solid_img = cand_b

        if solid_img and os.path.exists(solid_img):
            # T25 体积预算: 封面 CAD 渲染是摄影类大图, 转 JPEG q90 (PNG 保留给线条类拼合图);
            # 派生 jpg 每次无条件重生成, 防上个产品的旧封面混入 (prepare_cover_jpeg 内控)。
            # (封面角章按用户裁决移除: 缺失清单只在完成弹窗与 missing_fields.json 呈现)
            try:
                solid_img = prepare_cover_jpeg(solid_img)
            except Exception as e:
                print(f"[Notice] 封面 JPEG 转换跳过 (沿用 PNG): {e}")
            b_l, b_t, b_w, b_h, _ = SLIDE_SAFE_BOXES.get(
                1,
                (Inches(2.00), Inches(0.50), Inches(6.00), Inches(4.45), False),
            )
            if insert_or_replace_picture(s1, solid_img, b_l, b_t, b_w, b_h):
                print(
                    f"[Slide 1] Replaced cover model with pure CAD solid body (无网格、无节点): {os.path.basename(solid_img)}"
                )
            else:
                record_missing("封面模型图: 第 1 页放置封面图失败")
        else:
            record_missing(
                "封面模型图: solid_model.png 缺失 (VBS 导出失败且无重绘兜底)"
            )

    # ------------------ Slide 2: 网格模型图与统计文本 ------------------
    if total_slides >= 2:
        s2 = prs.slides[1]
        # 替换左侧网格模型图并自适应安全框
        mesh_img = os.path.join(data_dir, "mesh_model.png")
        if not os.path.exists(mesh_img):
            cand = os.path.join(data_dir, f"mode_{mode.lower()}", "mesh_model.png")
            if os.path.exists(cand):
                mesh_img = cand

        if os.path.exists(mesh_img):
            if trim_white_borders:
                try:
                    with Image.open(mesh_img) as m_im:
                        trimmed_m = trim_white_borders(m_im, border=8).convert("RGB")
                        cropped_mesh_path = os.path.join(
                            data_dir, "mesh_model_cropped.png"
                        )
                        trimmed_m.save(cropped_mesh_path, "PNG")
                        mesh_img = cropped_mesh_path
                except Exception as e:
                    print(f"[Notice] Error trimming mesh model: {e}")

            b_l, b_t, b_w, b_h, _ = SLIDE_SAFE_BOXES.get(
                2, (Inches(0.20), Inches(1.35), Inches(5.10), Inches(5.15), False)
            )
            insert_or_replace_picture(
                s2, mesh_img, b_l, b_t, b_w, b_h, log_tag="[Slide 2]"
            )
            print(
                "[Slide 2] Updated mesh model with proportional safe-box scaling and white border trimming."
            )

        # 替换右侧网格统计数据（严格保留模板原生空行排版与坐标，Arial 8pt，空行分明，杜绝挤压）
        formatted_lines, mesh_missing = get_formatted_mesh_text(data_dir)
        MISSING_FIELDS.extend(mesh_missing)
        for shape in s2.shapes:
            if shape.has_text_frame and (
                "实体计数" in shape.text_frame.text
                or "三角" in shape.text_frame.text
                or "文本框 2" in shape.name
            ):
                set_text_frame_keep_text_only(shape, formatted_lines, "Arial", 8)
                print(
                    "[Slide 2] Updated mesh statistics text frame (严格保留模板空行格式与排版，空行分明，无挤压)."
                )
                break

    # ------------------ Slide 3: 材料信息四图全部替换 ------------------
    if total_slides >= 3:
        s3 = prs.slides[2]
        ph_path = None
        # 按象限归集既有图片 (每象限取第一张); 无图的象限稍后按象限框插入新图,
        # 模板四联图被清空时同样出图
        quadrant_pics = {}
        for shape in s3.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                q = material_quadrant_index(shape.left, shape.top)
                if q not in quadrant_pics:
                    quadrant_pics[q] = shape
        for q, key in enumerate(MATERIAL_QUADRANT_KEYS):
            mat_img = os.path.join(data_dir, f"{key}.png")
            pic = quadrant_pics.get(q)
            if not os.path.exists(mat_img):
                # 缺图: 占位替换, 绝不让模板里上个产品的材料图混入
                if ph_path is None:
                    ph_path = ensure_missing_plot_placeholder(data_dir)
                record_missing(f"材料图缺失: {MATERIAL_QUADRANT_LABELS[q]}")
                print(
                    f"[Slide 3][ERROR] Image not found for {MATERIAL_QUADRANT_LABELS[q]}, 使用缺失占位图"
                )
                mat_img = ph_path
            if pic is not None:
                replace_picture_blob(pic, mat_img)
            else:
                b_l, b_t, b_w, b_h = MATERIAL_QUADRANT_BOXES[q]
                geom = compute_safe_box_fit(mat_img, b_l, b_t, b_w, b_h)
                s3.shapes.add_picture(mat_img, *(geom or (b_l, b_t, b_w, b_h)))
        print(
            "[Slide 3] Updated Material 4-panel figures (基本属性、推荐工艺、粘度、PVT 四图全部精准替换)."
        )

    # ------------------ Slide 4 ~ 16: 各项分析结果图与 GIF ------------------
    enabled_plots = [p for p in config.get("plots", []) if p.get("enabled", True)]
    fixed_plots = [
        p for p in enabled_plots if p.get("slide") and p.get("slide") <= total_slides
    ]
    extra_plots = [
        p for p in enabled_plots if not p.get("slide") or p.get("slide") > total_slides
    ]

    mode_sub = f"mode_{mode.lower()}"
    mode_dir = os.path.join(data_dir, mode_sub)

    for p_cfg in fixed_plots:
        slide_no = p_cfg["slide"]
        slide = prs.slides[slide_no - 1]  # 0-indexed
        key = p_cfg["key"]
        p_type = p_cfg.get("type", "image")

        # Slide 14~16: 模板契约保留页 (不放结果图)。用户显式勾选并分配到这三页时
        # 必须登记缺失, 不允许静默丢弃 (2026-09-09 体检 B-03)。
        if slide_no in [14, 15, 16]:
            record_missing(
                f"结果图未放置: {key} (第 {slide_no} 页为保留页 14-16, 请改页码或取消勾选)"
            )
            print(
                f"[Slide {slide_no}][WARN] 结果 {key} 分配在保留页 14-16, 未放置图片 (已登记缺失)"
            )
            continue

        # 安全区域配置
        safe_box_cfg = SLIDE_SAFE_BOXES.get(
            slide_no, (Inches(0.15), Inches(0.65), Inches(9.70), Inches(5.95), False)
        )
        b_l, b_t, b_w, b_h, allow_cover_title = safe_box_cfg

        if p_type == "gif":
            # GIF 动图 (Slide 4 充填时间)
            gif_path = os.path.join(data_dir, f"{key}.gif")
            if not os.path.exists(gif_path):
                gif_path = os.path.join(mode_dir, f"{key}.gif")
            if not os.path.exists(gif_path):
                gif_path = os.path.join(data_dir, "filling_animation.gif")

            if os.path.exists(gif_path):
                try:
                    try:
                        from core.gif_enhancer import optimize_existing_gif
                    except ImportError:
                        from gif_enhancer import optimize_existing_gif

                    target_delay = config.get("animation_settings", {}).get(
                        "delay_ms", 80
                    )
                    # T25 体积预算: GIF 是报告体积大头 — 帧数上限取配置帧数,
                    # 高度上限 720 (PPT 页内显示尺寸约 4.8 英寸 ≈ 720px 足够清晰)
                    n_frames = config.get("animation_settings", {}).get("frames")
                    optimize_existing_gif(
                        gif_path,
                        target_delay_ms=target_delay,
                        max_frames=int(n_frames) if n_frames else None,
                        max_height=720,
                    )
                except Exception as e:
                    print(f"[Notice] GIF optimization skipped: {e}")

                if insert_or_replace_picture(
                    slide,
                    gif_path,
                    b_l,
                    b_t,
                    b_w,
                    b_h,
                    allow_cover_title,
                    log_tag=f"[Slide {slide_no}]",
                ):
                    print(
                        f"[Slide {slide_no}] Updated GIF animation with safe-box fitting: {os.path.basename(gif_path)}"
                    )
                else:
                    record_missing(
                        f"充填动画占位失败: {key} (第 {slide_no} 页放置失败)"
                    )

                # 锁定用户微调后的 Shift+F5 播放提示文本框位置，并置于顶层杜绝遮挡
                for shape in slide.shapes:
                    if shape.has_text_frame and (
                        "F5" in shape.text_frame.text.upper()
                        or "播放" in shape.text_frame.text
                    ):
                        shape.left = 6932613
                        shape.top = 5373688
                        shape.width = 1503362
                        shape.height = 275590
                        try:
                            sp_elem = shape._element
                            sp_parent = sp_elem.getparent()
                            sp_parent.remove(sp_elem)
                            sp_parent.append(sp_elem)
                        except Exception as e:
                            print(
                                f"[Slide 4][Notice] F5 提示框置顶失败 (形状: {shape.name}): {e}"
                            )
                        print(
                            f"[Slide 4] Locked 'Shift+F5' prompt box coordinates to user-adjusted layout and brought to front."
                        )
            else:
                # 缺 GIF: 用占位图替换模板旧图, 绝不留上一次运行的动画
                record_missing(f"充填动画缺失: {key} (第 {slide_no} 页)")
                print(
                    f"[Slide {slide_no}][ERROR] GIF not found for {key}, 使用缺失占位图"
                )
                ph = ensure_missing_plot_placeholder(data_dir)
                insert_or_replace_picture(
                    slide,
                    ph,
                    b_l,
                    b_t,
                    b_w,
                    b_h,
                    allow_cover_title,
                    log_tag=f"[Slide {slide_no}]",
                )
        else:
            # 普通云图图片 (Slide 5 ~ 16)
            # 优先从该模式专属子目录寻找
            img_path = os.path.join(mode_dir, f"{key}.png")
            if not os.path.exists(img_path):
                # 回退查找模式 A 目录
                cand_a = os.path.join(data_dir, "mode_a", f"{key}.png")
                if os.path.exists(cand_a):
                    img_path = cand_a
                else:
                    # 回退到根临时目录查找
                    cand_root = os.path.join(data_dir, f"{key}.png")
                    if os.path.exists(cand_root):
                        img_path = cand_root

            if os.path.exists(img_path):
                insert_or_replace_picture(
                    slide,
                    img_path,
                    b_l,
                    b_t,
                    b_w,
                    b_h,
                    allow_cover_title,
                    log_tag=f"[Slide {slide_no}][{mode_tag}]",
                )
                tag_info = (
                    "极限全屏覆盖标题" if allow_cover_title else "自适应安全最大化"
                )
                print(
                    f"[Slide {slide_no}][{mode_tag}] Updated result plot: {p_cfg.get('desc', key)} ({tag_info})"
                )
            else:
                # 缺图: 用占位图替换模板旧图 (模板是填好的历史报告, 旧图=上个零件的云图)
                record_missing(f"结果图缺失: {key} (第 {slide_no} 页)")
                print(
                    f"[Slide {slide_no}][ERROR] Image not found for {key} in {mode_dir}, 使用缺失占位图"
                )
                ph = ensure_missing_plot_placeholder(data_dir)
                insert_or_replace_picture(
                    slide,
                    ph,
                    b_l,
                    b_t,
                    b_w,
                    b_h,
                    allow_cover_title,
                    log_tag=f"[Slide {slide_no}]",
                )

        # Slide 9: 自动回填注塑机最大锁模力与 CAE 最大锁模力
        if slide_no == 9:
            cae_ton, mach_ton = resolve_clamp_force(config, data_dir)
            if cae_ton is None:
                record_missing("锁模力: CAE 计算最大锁模力 (clamp_force_info.json)")
            if mach_ton is None:
                record_missing("锁模力: 注塑机最大吨位 (config machine_max_ton)")

            for shape in slide.shapes:
                if shape.has_table:
                    tbl = shape.table
                    for row in tbl.rows:
                        for c_idx, cell in enumerate(row.cells):
                            kind = clamp_label_kind(cell.text)
                            if kind and c_idx + 1 < len(row.cells):
                                val = cae_ton if kind == "cae" else mach_ton
                                update_table_cell(
                                    row.cells[c_idx + 1],
                                    f"{val:g}T" if val is not None else MISSING_TEXT,
                                    "微软雅黑",
                                    12,
                                )
                    print(
                        f"[Slide 9] Auto-filled clamp force table: CAE={cae_ton if cae_ton is not None else '数据缺失'}, Machine={mach_ton if mach_ton is not None else '数据缺失'}"
                    )

    # ------------------ Slide 8: 材料推荐成型范围回填 (真实材料数据) ------------------
    if total_slides >= 8:
        fill_material_molding_range(prs.slides[7], data_dir)

    # ------------------ Slide 14 ~ 16 (XYZ 变形): 不再剥离图片 ------------------
    # 历史原因: 旧模板 14-16 页内置了 XYZ 变形的占位图, 当时通过程序清空以避免错误结果展示。
    # 用户裁决 (2026-09-08): 模板本身已手动清空, 此类条目默认不再勾选 (见 report_config.json
    #   warpage_x/y/z.enabled=false), 因此运行时剥离块的副作用 (把页面正文图片误删) 反而不可接受。
    # 该块代码移除; 后续如需重新启用 XYZ 变形, 在 PPT 模板直接放对应图片即可, 勿再在程序层剥离。

    # 如果有用户勾选的额外结果项目且已导出图片，追加插入新页
    if extra_plots and total_slides >= 5:
        ref_slide = prs.slides[4]
        slide_layout = ref_slide.slide_layout
        for p_cfg in extra_plots:
            key = p_cfg["key"]
            img_path = os.path.join(mode_dir, f"{key}.png")
            if not os.path.exists(img_path):
                img_path = os.path.join(data_dir, f"{key}.png")

            if os.path.exists(img_path):
                new_slide = prs.slides.add_slide(slide_layout)
                if new_slide.shapes.title:
                    new_slide.shapes.title.text = p_cfg.get("plot_name", key)
                left = Inches(0.15)
                top = Inches(0.65)
                width = Inches(9.70)
                height = Inches(5.95)
                new_slide.shapes.add_picture(img_path, left, top, width, height)
                print(
                    f"[Extra Slide][{mode_tag}] Added new slide for: {p_cfg.get('plot_name', key)}"
                )
            else:
                # 追加页缺图同样登记缺失: 此前静默跳过, 报告少页无任何提示 (B-02)
                record_missing(f"结果图缺失: {key} (追加页, 未插入新页)")
                print(
                    f"[Extra Slide][{mode_tag}][ERROR] Image not found for {key}, 未追加新页"
                )

    # 确定输出路径: config.output_dir 显式配置优先, 留空默认输出到本次模型所在目录
    # (manifest.model_dir), 均不可用回退项目根 (相对路径相对项目根解析)
    if not output_path:
        out_dir = resolve_output_dir_for_run(config, manifest)
        filename = f"{study_name}-{today_str}-模流分析报告-{mode_tag}.pptx"
        output_path = os.path.join(out_dir, filename)

    if (
        str(config.get("on_missing_data", "annotate")).lower() == "fail"
        and MISSING_FIELDS
    ):
        raise RuntimeError(
            "数据缺失 (on_missing_data=fail): " + "; ".join(MISSING_FIELDS)
        )

    output_path = os.path.abspath(output_path)
    # 保存 PPTX（若文件已被独占打开，则自动追加时间戳，避免 PermissionError）
    saved = False
    original_path = output_path
    counter = 0
    while not saved and counter < 5:
        try:
            prs.save(output_path)
            saved = True
        except PermissionError:
            counter += 1
            base, ext = os.path.splitext(original_path)
            time_suffix = datetime.datetime.now().strftime("%H%M%S")
            output_path = f"{base}_{time_suffix}{ext}"
            print(
                f"[Notice] File locked by PowerPoint, saving as: {os.path.basename(output_path)}"
            )

    print(f"\n[SUCCESS] 【{mode_tag}】报告成功生成:")
    print(f"  {output_path}\n")
    return output_path


def build_report(config_path, data_dir, output_path=None, mode=None, no_open=False):
    with open(config_path, "r", encoding="utf-8-sig") as f:
        config = json.load(f)

    if mode is None:
        mode = config.get("screenshot_mode", "B")
    mode = str(mode).upper().strip()

    # 模板路径: 相对路径相对项目根解析 (templates/xxx.pptx), 绝对路径原样
    template_path = resolve_template_path(config.get("template_pptx"))

    # T26: 模板哈希校验 — 布局几何 (SLIDE_SAFE_BOXES/标记文字) 钉死在特定模板上,
    # 模板被更换/修订时提前告警, 防止布局静默错乱
    import hashlib

    hash_path = os.path.join(data_dir, "template_hash.txt")
    try:
        h = hashlib.sha256()
        with open(template_path, "rb") as f:
            for chunk in iter(lambda: f.read(1 << 20), b""):
                h.update(chunk)
        digest = h.hexdigest()
        if os.path.exists(hash_path):
            with open(hash_path, "r", encoding="utf-8") as f:
                stored = f.read().strip()
            if stored and stored != digest:
                print(
                    "[WARN] 模板文件与上次运行不同 (sha256 变化)。"
                    "页面布局按原模板硬编码, 新模板可能错位 — 请人工核对版式。"
                )
        with open(hash_path, "w", encoding="utf-8") as f:
            f.write(digest)
    except Exception as e:
        print(f"[Notice] 模板哈希校验跳过: {e}")

    # 读取 manifest / 元数据 (损坏 → 空表兜底 + 告警, 不崩)
    manifest_path = os.path.join(data_dir, "manifest.json")
    manifest = {}
    if os.path.exists(manifest_path):
        try:
            with open(manifest_path, "r", encoding="utf-8-sig") as f:
                manifest = json.load(f)
            if not isinstance(manifest, dict):
                manifest = {}
                print("[Notice] manifest.json 根元素不是对象, 已按空处理")
        except Exception as e:
            manifest = {}
            print(f"[Notice] Error reading manifest.json: {e}")

    today_str = datetime.date.today().strftime("%Y-%m-%d")
    study_name = sanitize_filename(manifest.get("study_name", "方案"))
    if study_name.lower().endswith(".sdy"):
        study_name = study_name[:-4]
    # 零件名: 只用真实值 (manifest 未提供则留空, 绝不写"方案"默认词冒充数据)
    part_name = manifest.get("part_name", "") or ""
    if part_name:
        part_name = sanitize_filename(part_name)
    mf_version = manifest.get("moldflow_version", "2023")

    # 数据缺失策略与登记 (annotate=标注后继续, fail=存在缺失则拒绝出报告)
    MISSING_FIELDS.clear()
    on_missing = str(config.get("on_missing_data", "annotate")).lower()

    # 材料数据: 先把 VBS 官方字段枚举原始导出映射为 material_info.json
    build_material_info(data_dir)

    # 预先生成 100% 仿真官方属性与工艺卡片
    mat_missing = render_material_dialog_cards(data_dir, manifest)
    MISSING_FIELDS.extend(mat_missing)

    # XY 探针峰值单源 (本次导出: peak_values.json + *_curve_data.txt 解析; 缺失登记)
    peaks_raw = load_peaks(data_dir) if load_peaks is not None else {}
    resolved_peaks = resolve_peaks(peaks_raw, config)
    for curve, label in (
        ("clamp_force", "锁模力 XY"),
        ("inj_pressure", "注射位置处压力 XY"),
    ):
        if resolved_peaks.get(curve) is None:
            record_missing(
                f"XY 探针: {label} 峰值数据缺失 (peak_values.json 与 config 均无)"
            )

    # 方案 A 的 XY 页同样要有峰值探针 (标注 Moldflow 原生截图; 已标注的图自动跳过)
    if annotate_xy_curves is not None:
        mode_a_dir = os.path.join(data_dir, "mode_a")
        if os.path.exists(mode_a_dir):
            annotate_xy_curves(mode_a_dir, data_dir, resolved_peaks)
        else:
            annotate_xy_curves(None, data_dir, resolved_peaks)

    generated_reports = []
    mode_errors = []

    # 1. 方案 A：视口直接抓取 (T16: 按方案隔离, 单方案失败不拖垮另一方案)
    if mode in ["A", "BOTH", "ALL"]:
        try:
            out_a = output_path if mode == "A" else None
            res_a = build_single_report(
                config,
                manifest,
                data_dir,
                template_path,
                today_str,
                study_name,
                part_name,
                mf_version,
                mode="A",
                output_path=out_a,
            )
            generated_reports.append(res_a)
        except Exception as e:
            mode_errors.append(f"方案 A: {e}")
            print(f"[ERROR] 方案 A 构建失败: {e}")

    # 2. 方案 B：双图独立导出再智能拼合
    if mode in ["B", "BOTH", "ALL"]:
        try:
            # 拼合画布期望高度 (T8): 取 image_settings.height; 0x0 视口原生模式
            # 语义定义为固定 1920 宽画布 (确定性输出, 配合陈旧产物清理)
            img_cfg_b = config.get("image_settings", {}) or {}
            img_h = img_cfg_b.get("height", 1080)
            canvas_target_h = int(img_h) if img_h and int(img_h) > 0 else 1215
            # SaveImage3 期望分辨率 (0x0 视口原生配置 → None 跳过尺寸校验)
            img_w_b = img_cfg_b.get("width", 1920)
            expected_model_size = (
                (int(img_w_b), int(img_h))
                if img_w_b and int(img_w_b) > 0 and img_h and int(img_h) > 0
                else None
            )
            # 预先执行 Mode B 图像拼合
            mode_b_dir = os.path.join(data_dir, "mode_b")
            if process_all_mode_b_plots and os.path.exists(mode_b_dir):
                print("\n>>> 执行方案 B 标尺与模型图像智能拼合...")
                process_all_mode_b_plots(
                    mode_b_dir,
                    peaks=resolved_peaks,
                    target_height=canvas_target_h,
                    expected_model_size=expected_model_size,
                )

            out_b = output_path if mode == "B" else None
            res_b = build_single_report(
                config,
                manifest,
                data_dir,
                template_path,
                today_str,
                study_name,
                part_name,
                mf_version,
                mode="B",
                output_path=out_b,
            )
            generated_reports.append(res_b)
        except Exception as e:
            mode_errors.append(f"方案 B: {e}")
            print(f"[ERROR] 方案 B 构建失败: {e}")

    if mode_errors and not generated_reports:
        raise RuntimeError("; ".join(mode_errors))
    if mode_errors:
        print(f"[WARN] 部分方案构建失败 ({len(mode_errors)}): {'; '.join(mode_errors)}")

    # 数据完整性清单 (供完成弹窗/封面章/排查使用)
    if MISSING_FIELDS:
        print(f"\n[数据完整性] 缺失字段 {len(MISSING_FIELDS)} 项:")
        for item in MISSING_FIELDS:
            print(f"  - {item}")
    try:
        with open(
            os.path.join(data_dir, "missing_fields.json"), "w", encoding="utf-8"
        ) as f:
            json.dump(
                {"on_missing_data": on_missing, "missing": list(MISSING_FIELDS)},
                f,
                ensure_ascii=False,
                indent=2,
            )
    except Exception as e:
        print(f"[Notice] Could not write missing_fields.json: {e}")

    # 记录最后生成路径
    try:
        with open(
            os.path.join(data_dir, "last_output_path.txt"), "w", encoding="utf-8-sig"
        ) as f:
            f.write("\n".join(generated_reports))
    except Exception as e:
        print(f"[Notice] Could not write last_output_path.txt: {e}")

    # 自动打开 PPT
    if config.get("open_after_export", True) and not no_open:
        for rep in generated_reports:
            try:
                os.startfile(rep)
            except Exception as e:
                print(f"[Notice] Could not auto-open presentation {rep}: {e}")

    return generated_reports


def _attach_file_logging(log_file):
    """T12: stdout/stderr 追加写入 GBK 日志 (VBS 传 --log-file; 全链路 GBK)。"""
    if not log_file:
        return
    os.makedirs(os.path.dirname(os.path.abspath(log_file)), exist_ok=True)

    class _GbkTee:
        def __init__(self, stream, path):
            self.stream = stream
            self.f = open(path, "ab")

        def write(self, s):
            self.stream.write(s)
            try:
                self.f.write(str(s).encode("gbk", errors="replace"))
            except Exception as e:
                print(f"[Notice] log write failed: {e}")
            return len(s)

        def flush(self):
            self.stream.flush()
            self.f.flush()

    sys.stdout = _GbkTee(sys.stdout, log_file)
    sys.stderr = _GbkTee(sys.stderr, log_file)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Build PPTX report from Moldflow results."
    )
    parser.add_argument(
        "--config", default="report_config.json", help="Path to config json file"
    )
    parser.add_argument(
        "--data-dir",
        default="temp",
        help="Directory containing exported plots and manifest",
    )
    parser.add_argument("--output", default=None, help="Output PPTX file path")
    parser.add_argument(
        "--mode",
        default=None,
        choices=["A", "B", "BOTH", "ALL"],
        help="Screenshot mode (A, B, BOTH, or ALL; default: config screenshot_mode=B). "
        "--output is only valid with a single mode",
    )
    parser.add_argument(
        "--no-open",
        action="store_true",
        help="Do not auto-open the generated PPTX (default is to open when config open_after_export=true; use --open to force)",
    )
    parser.add_argument(
        "--open",
        dest="force_open",
        action="store_true",
        help="Force auto-open regardless of config",
    )
    parser.add_argument(
        "--log-file",
        default=None,
        help="Append stdout/stderr to this GBK log file (used by AutoReport.vbs)",
    )
    args = parser.parse_args()

    if args.output and args.mode and args.mode.upper() in ("BOTH", "ALL"):
        parser.error(
            "--output 不能与 BOTH/ALL 组合 (会生成两份报告); 请指定单一 --mode 或省略 --output"
        )

    _attach_file_logging(args.log_file)
    # 退出码语义 (T12): 0=成功; 1=环境问题 (模板缺失等); 2=构建失败
    try:
        build_report(
            args.config,
            args.data_dir,
            args.output,
            mode=args.mode,
            no_open=args.no_open and not args.force_open,
        )
        sys.exit(0)
    except FileNotFoundError as e:
        print(f"[ERROR] 环境问题: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"[ERROR] 构建失败: {e}")
        sys.exit(2)
