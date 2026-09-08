"""零依赖回归测试 (stdlib only) — T1 假数据清除 / T2 峰值单源 / T3 模式分发。

运行: python tests/run_checks.py
退出码 0 = 全部通过; 1 = 存在失败。
函数级测试, 不需要 PPT 模板与 temp/ 数据。
"""

import json
import os
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from core import pptx_builder as pb

# 任何 fabricated 默认值出现即为失败 (来自旧版 get_formatted_mesh_text /
# render_material_dialog_cards 的硬编码假数据)
FAKE_MARKERS = [
    "61978",
    "30991",
    "9662.27",
    "1288.302",
    "94.7",
    "96.4",
    "15.73",
    "1.84",
    "1.16",
    "EP300H",
    "SABIC",
    "320.5",
]

REAL_MESH = {
    "mesh_type": "DualDomain",
    "triangles": 12345,
    "nodes": 6789,
    "connectivity_regions": 1,
    "unvisible_triangles": 3,
    "volume": 100.5,
    "surface_area": 750.25,
    "max_aspect_ratio": 10.2,
    "ave_aspect_ratio": 2.1,
    "min_aspect_ratio": 1.05,
    "free_edges": 4,
    "manifold_edges": 18000,
    "non_manifold_edges": 0,
    "unoriented": 0,
    "intersection_elements": 0,
    "overlap_elements": 0,
    "match_ratio": 92.3,
    "reciprocal_match_ratio": 95.1,
}


def _mesh_lines(lines):
    return "\n".join(lines)


def test_mesh_missing_file(tmp):
    """mesh_summary.json 缺失 → 全部占位, 绝无假数据。"""
    lines, missing = pb.get_formatted_mesh_text(tmp)
    joined = _mesh_lines(lines)
    for marker in FAKE_MARKERS:
        assert marker not in joined, f"假数据标记 {marker} 出现在网格文本中"
    assert missing, "缺失字段列表为空"
    assert any("数据不完整" in ln for ln in lines), "结论行未按缺失降级"
    assert not any(
        "适合 " in ln and "分析。" in ln and "数据不完整" not in ln for ln in lines
    )


def test_mesh_corrupt_json(tmp):
    """mesh_summary.json 损坏 → 同缺失处理。"""
    with open(os.path.join(tmp, "mesh_summary.json"), "w", encoding="utf-8") as f:
        f.write("{not valid json!!")
    lines, missing = pb.get_formatted_mesh_text(tmp)
    joined = _mesh_lines(lines)
    for marker in FAKE_MARKERS:
        assert marker not in joined, f"假数据标记 {marker} 出现在网格文本中"
    assert missing


def test_mesh_null_fields(tmp):
    """字段为 null → 该字段占位而非 0/假值。"""
    with open(os.path.join(tmp, "mesh_summary.json"), "w", encoding="utf-8") as f:
        json.dump({k: None for k in REAL_MESH}, f)
    lines, missing = pb.get_formatted_mesh_text(tmp)
    joined = _mesh_lines(lines)
    assert "数据缺失" in joined or "—" in joined
    assert missing


def test_mesh_real_data(tmp):
    """真实完整数据 → 数值如实渲染, 结论行正常。"""
    with open(os.path.join(tmp, "mesh_summary.json"), "w", encoding="utf-8") as f:
        json.dump(REAL_MESH, f)
    lines, missing = pb.get_formatted_mesh_text(tmp)
    joined = _mesh_lines(lines)
    assert "12345" in joined and "6789" in joined
    assert "750.25" in joined
    assert any("适合" in ln and "分析。" in ln for ln in lines)
    assert missing == [], f"完整数据不应报缺失: {missing}"


def test_material_missing_file(tmp):
    """material_info.json 缺失 → 材料卡逐字段'数据缺失', 绝无假材料。"""
    missing = pb.render_material_dialog_cards(tmp)
    assert isinstance(missing, list) and len(missing) >= 14, "材料缺失字段列表异常"
    assert any("牌号" in m for m in missing) and any("制造商" in m for m in missing)
    assert os.path.exists(os.path.join(tmp, "material_basic.png"))
    assert os.path.exists(os.path.join(tmp, "material_process.png"))


def test_material_sentinel_numerics(tmp):
    """数值字段为 -1 哨兵 (VBS COM 失败标记) → 视为缺失。"""
    info = {"trade_name": "ABC", "manufacturer": "某厂", "mold_temp_min": -1}
    with open(os.path.join(tmp, "material_info.json"), "w", encoding="utf-8") as f:
        json.dump(info, f)
    missing = pb.render_material_dialog_cards(tmp)
    assert any("模具温度" in m or "mold_temp_min" in m for m in missing)


def test_resolve_clamp_force(tmp):
    """CAE 锁模力只认 clamp_force_info.json; 机台吨位只认 config; 缺失为 None。"""
    cae, mach = pb.resolve_clamp_force({}, tmp)
    assert cae is None and mach is None
    with open(os.path.join(tmp, "clamp_force_info.json"), "w", encoding="utf-8") as f:
        json.dump({"cae_max_clamp_force": 280.3}, f)
    cae, mach = pb.resolve_clamp_force(
        {"clamp_force_settings": {"machine_max_ton": 350}}, tmp
    )
    assert cae == 280.3 and mach == 350
    cae, mach = pb.resolve_clamp_force({}, tmp)
    assert cae == 280.3 and mach is None


def test_resolve_peaks_exported(tmp):
    """导出数据优先: peak_values.json 的曲线峰值直接采用。"""
    peaks = {
        "clamp_force": {"peak_value": 320.5, "peak_time": 6.224},
        "inj_pressure": {"peak_value": 21.45, "peak_time": 5.092},
    }
    r = pb.resolve_peaks(peaks, {})
    assert r["clamp_force"] == (320.5, 6.224)
    assert r["inj_pressure"] == (21.45, 5.092)


def test_resolve_peaks_no_config_fallback(tmp):
    """导出缺失 → None。config 手工值 (上个产品的数据) 绝不允许再被静默采用。"""
    config = {
        "clamp_force_settings": {"cae_max_ton": 300.0, "peak_time_sec": 5.5},
        "inj_pressure_settings": {"max_pressure_mpa": 20.0, "peak_time_sec": 4.8},
    }
    r = pb.resolve_peaks({}, config)
    assert r["clamp_force"] is None, (
        "config 陈旧峰值回退未清除 (换产品后臆造上个产品数据)"
    )
    assert r["inj_pressure"] is None
    # 值有效但时刻缺失 (GetMaxValue 回退路径) → 仍无法标注, 不得回退 config
    r2 = pb.resolve_peaks(
        {"clamp_force": {"peak_value": 359.8, "peak_time": None}}, config
    )
    assert r2["clamp_force"] is None


def test_curve_data_peak_parse(tmp):
    """SaveXYPlotCurveData 导出的曲线 txt → 解析出 (峰值, 时刻), 覆盖 json 值。"""
    from core import image_processor as ip

    curve = (
        '"Time [s]","Pressure [MPa]"\n'
        "0.0000E+000,1.2000E+000\n"
        "2.5460E+000,1.8300E+001\n"
        "5.0920E+000,2.1450E+001\n"
        "6.1000E+000,1.9000E+001\n"
    )
    with open(os.path.join(tmp, "inj_pressure_xy_curve_data.txt"), "w") as f:
        f.write(curve)
    with open(os.path.join(tmp, "peak_values.json"), "w", encoding="utf-8") as f:
        json.dump({"inj_pressure": {"peak_value": None, "peak_time": None}}, f)
    peaks = ip.load_peaks(tmp)
    assert peaks["inj_pressure"]["peak_value"] == 21.45
    assert peaks["inj_pressure"]["peak_time"] == 5.092


def test_mode_a_xy_probe(tmp):
    """方案 A 的 XY 截图也要有峰值探针 (用户裁决回归): annotate_xy_curves 给
    mode_a 与 data_dir 两处的 Moldflow 原生 XY 图画黄色探针, 与方案 B 一致。"""
    import numpy as np
    from PIL import Image

    from core import image_processor as ip

    mode_a = os.path.join(tmp, "mode_a")
    os.makedirs(mode_a)
    xy_path = os.path.join(mode_a, "inj_pressure_xy.png")
    Image.new("RGB", (400, 300), (255, 255, 255)).save(xy_path)
    root_path = os.path.join(tmp, "clamp_force_xy.png")
    Image.new("RGB", (400, 300), (255, 255, 255)).save(root_path)
    peaks = {
        "inj_pressure": (21.45, 5.092),
        "clamp_force": (320.5, 6.224),
    }
    ip.annotate_xy_curves(mode_a, tmp, peaks)

    def yellow_count(path):
        arr = np.array(Image.open(path).convert("RGB"))
        return int(
            (
                (arr[:, :, 0] > 240)
                & (arr[:, :, 1] > 240)
                & (arr[:, :, 2] > 180)
                & (arr[:, :, 2] < 225)
            ).sum()
        )

    assert yellow_count(xy_path) > 100, "方案 A 目录的 XY 图未画探针"
    assert yellow_count(root_path) > 100, "根目录 XY 图未画探针"
    # 峰值缺失 → 跳过标注, 图保持原样 (绝不臆造)
    Image.new("RGB", (400, 300), (255, 255, 255)).save(xy_path)
    ip.annotate_xy_curves(mode_a, tmp, {"inj_pressure": None, "clamp_force": None})
    assert yellow_count(xy_path) == 0, "峰值缺失时不应画探针"


def test_find_duplicate_slides(tmp):
    """GUI 页码撞车检测: 勾选+同页 → 报; 未勾选/无页码/非法页码 → 跳过。"""
    plots = [
        {"key": "a", "plot_name": "压力", "slide": 9, "enabled": True},
        {"key": "b", "plot_name": "锁模力XY", "slide": 9, "enabled": True},
        {"key": "c", "plot_name": "温度", "slide": 8, "enabled": True},
        {"key": "d", "plot_name": "气穴", "slide": 9, "enabled": False},
        {"key": "e", "plot_name": "备选", "slide": None, "enabled": True},
        {"key": "f", "plot_name": "坏页码", "slide": "x", "enabled": True},
    ]
    dups = config_gui_mod().find_duplicate_slides(plots)
    assert dups == {9: ["压力", "锁模力XY"]}, dups
    # 解除一个勾选 → 无撞页
    plots[1]["enabled"] = False
    assert config_gui_mod().find_duplicate_slides(plots) == {}


def test_match_plots_to_available(tmp):
    """动态结果清单精确匹配 (用户裁决: 不猜测/不别名):
    "体积收缩率" 不得匹配 "顶出时的体积收缩率", 一字不差才算命中。"""
    plots = [
        {"key": "vs", "plot_name": "体积收缩率"},
        {"key": "vse", "plot_name": "顶出时的体积收缩率"},
        {"key": "wl", "plot_name": "熔接线"},
        {"key": "missing", "plot_name": "不存在的结果"},
    ]
    available = ["体积收缩率", "熔接线", "充填时间"]
    m = config_gui_mod().match_plots_to_available(plots, available)
    assert m == {"vs": True, "vse": False, "wl": True, "missing": False}
    # 空清单 → 全部不命中
    assert not any(config_gui_mod().match_plots_to_available(plots, []).values())


def config_gui_mod():
    """以模块方式加载 config_gui (不启动 Tk 窗口)。"""
    import importlib.util

    spec = importlib.util.spec_from_file_location(
        "config_gui_under_test", os.path.join(ROOT, "config_gui.py")
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def test_broken_model_falls_back_to_viewport(tmp):
    """SaveImage3 残缺导出回归 (实发: 模型只渲染一条横带, 内容占画布 <10%):
    拼合时必须弃用残缺模型图, 回退视口图直出 (完整正确保底)。"""
    import numpy as np
    from PIL import Image

    from core import image_processor as ip

    data_dir = tmp
    mode_a = os.path.join(data_dir, "mode_a")
    mode_b = os.path.join(data_dir, "mode_b")
    os.makedirs(mode_a)
    os.makedirs(mode_b)

    # 残缺模型图: 1200x900 白底, 模型只渲染一条 50px 高的横带 (面积占比 ~4%,
    # 对应实发案例 3.6%: SaveImage3 输出 3840x2160 但模型只渲染出一条横带)
    broken = np.full((900, 1200, 3), 255, dtype=np.uint8)
    broken[420:470, 100:1100] = (30, 180, 30)
    Image.fromarray(broken).save(os.path.join(mode_b, "model_pressure.png"))
    # 视口图: 完整模型 (占面积大) + 左侧红色数据条
    vp = np.full((1136, 2112, 3), 255, dtype=np.uint8)
    vp[100:1000, 400:2000] = (30, 30, 200)
    vp[50:1000, 30:150] = (200, 30, 30)
    Image.fromarray(vp).save(os.path.join(mode_a, "pressure.png"))
    # 正常色带图 (避免 scale 兜底)
    sc = np.full((1100, 200, 3), 255, dtype=np.uint8)
    Image.fromarray(sc).save(os.path.join(mode_b, "scale_pressure.png"))

    out = os.path.join(mode_b, "pressure.png")
    ok = ip.merge_scale_and_model(
        os.path.join(mode_b, "scale_pressure.png"),
        os.path.join(mode_b, "model_pressure.png"),
        out,
        viewport_path=os.path.join(mode_a, "pressure.png"),
        target_height=1000,
    )
    assert ok, "残缺回退路径未产出"
    arr = np.array(Image.open(out).convert("RGB"))
    # 输出应为视口图直出 (2112x1136), 而非残缺模型拼合
    assert arr.shape[:2] == (1136, 2112), f"未按视口图直出: {arr.shape}"
    blue = ((arr[:, :, 2] > 120) & (arr[:, :, 0] < 100)).sum()
    assert blue > 50000, "视口图内容缺失"

    # 正常模型图 (内容占画布 ~55%) 不触发回退, 走正常拼合
    good = np.full((800, 600, 3), 255, dtype=np.uint8)
    good[100:700, 100:500] = (30, 30, 200)
    Image.fromarray(good).save(os.path.join(mode_b, "model_pressure.png"))
    ok = ip.merge_scale_and_model(
        os.path.join(mode_b, "scale_pressure.png"),
        os.path.join(mode_b, "model_pressure.png"),
        out,
        viewport_path=os.path.join(mode_a, "pressure.png"),
        target_height=1000,
    )
    assert ok
    arr2 = np.array(Image.open(out).convert("RGB"))
    assert abs(arr2.shape[0] / arr2.shape[1] - 1 / 1.58) < 0.02, (
        f"正常模型不应走视口直出: {arr2.shape}"
    )

    # 渲染断带特征 (4K 实发: 内容高度占比 30%, 宽高比 2.66) → 回退视口图;
    # 同样内容但分辨率与配置一致且非断带 → 正常拼合
    band = np.full((2160, 3840, 3), 255, dtype=np.uint8)
    band[720:1370, 500:3340] = (
        30,
        30,
        200,
    )  # trim 后 ~2850x666: 高占比 31%, 宽高比 4.3
    Image.fromarray(band).save(os.path.join(mode_b, "model_pressure.png"))
    ok = ip.merge_scale_and_model(
        os.path.join(mode_b, "scale_pressure.png"),
        os.path.join(mode_b, "model_pressure.png"),
        out,
        viewport_path=os.path.join(mode_a, "pressure.png"),
        target_height=1000,
        expected_model_size=(3840, 2160),
    )
    assert ok
    arr_band = np.array(Image.open(out).convert("RGB"))
    assert arr_band.shape[:2] == (1136, 2112), f"断带图未回退视口图: {arr_band.shape}"

    # 分辨率不符 (实发: 配 2560x1440 实出 3840x2160 且渲染残缺) → 弃用回退视口图
    hi_res = np.full((2160, 3840, 3), 255, dtype=np.uint8)
    hi_res[500:1500, 800:3000] = (30, 30, 200)
    Image.fromarray(hi_res).save(os.path.join(mode_b, "model_pressure.png"))
    ok = ip.merge_scale_and_model(
        os.path.join(mode_b, "scale_pressure.png"),
        os.path.join(mode_b, "model_pressure.png"),
        out,
        viewport_path=os.path.join(mode_a, "pressure.png"),
        target_height=1000,
        expected_model_size=(2560, 1440),
    )
    assert ok
    arr3 = np.array(Image.open(out).convert("RGB"))
    assert arr3.shape[:2] == (1136, 2112), f"分辨率不符未回退视口图: {arr3.shape}"
    # 分辨率相符的正常大图 → 正常拼合 (不回退)
    ok = ip.merge_scale_and_model(
        os.path.join(mode_b, "scale_pressure.png"),
        os.path.join(mode_b, "model_pressure.png"),
        out,
        viewport_path=os.path.join(mode_a, "pressure.png"),
        target_height=1000,
        expected_model_size=(3840, 2160),
    )
    assert ok
    arr4 = np.array(Image.open(out).convert("RGB"))
    assert abs(arr4.shape[0] / arr4.shape[1] - 1 / 1.58) < 0.02, (
        f"分辨率相符不应回退: {arr4.shape}"
    )


def test_cover_jpeg_regenerated(tmp):
    """封面 JPEG 派生缓存必须每次重生成, 不得沿用上个产品的旧 jpg。"""
    from PIL import Image

    solid = os.path.join(tmp, "solid_model.png")
    Image.new("RGB", (60, 40), (10, 200, 30)).save(solid)
    jpg = os.path.join(tmp, "solid_model_cover.jpg")
    Image.new("RGB", (60, 40), (200, 10, 10)).save(jpg, "JPEG")  # 陈旧旧图
    out = pb.prepare_cover_jpeg(solid)
    assert out == jpg
    im = Image.open(jpg)
    # 重新编码后必须来自本次的 solid_model.png (绿色占主导)
    assert im.getpixel((30, 20))[1] > 150, "封面 jpg 未随 solid_model.png 重新生成"


def test_replace_blob_format_match(tmp):
    """目标图元部件是 .png 时, 即使新图是 JPEG 也必须按 PNG 编码写入 (格式错配会致图无法显示)。"""
    from PIL import Image
    from pptx import Presentation
    from pptx.util import Inches

    png_src = os.path.join(tmp, "a.png")
    Image.new("RGB", (40, 30), (1, 2, 3)).save(png_src, "PNG")
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    pic = slide.shapes.add_picture(png_src, Inches(1), Inches(1))
    tpl = os.path.join(tmp, "t.pptx")
    prs.save(tpl)

    jpg_src = os.path.join(tmp, "b.jpg")
    Image.new("RGB", (50, 40), (9, 8, 7)).save(jpg_src, "JPEG")

    prs2 = Presentation(tpl)
    pic2 = prs2.slides[0].shapes[0]
    assert pb.replace_picture_blob(pic2, jpg_src)
    part = pic2.part.rels[pic2._element.blip_rId].target_part
    blob = part.blob
    assert blob[:8] == b"\x89PNG\r\n\x1a\n", (
        "部件扩展名 png 却写入 JPEG 字节 (格式错配)"
    )
    assert part.content_type == "image/png"


def test_clamp_label_kind(tmp):
    """S9 锁模力表格标签必须精确匹配; '分析要求'长文本含关键词也不得命中。"""
    assert pb.clamp_label_kind("CAE最大锁模力") == "cae"
    assert pb.clamp_label_kind("注塑机最大锁模力") == "machine"
    assert (
        pb.clamp_label_kind(
            "最大压力<注塑机极限压力×70%；CAE最大锁模力＜实际注塑机最大锁模力*80%"
        )
        is None
    ), "长说明文本被误判为标签单元格 (359.8T 写入分析要求事故回归)"


def test_conclusion_rows_cleared(tmp):
    """模板残留的'结果说明'类结论 (对上个产品的判断) 必须清空, 分析要求等通用标准保留。"""
    from pptx import Presentation
    from pptx.util import Inches

    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    gf = s.shapes.add_table(3, 2, 0, 0, Inches(4), Inches(1)).table
    gf.cell(0, 0).text = "结果说明"
    gf.cell(0, 1).text = "缩痕深度>0.03mm，本产品缩痕可见。"
    gf.cell(1, 0).text = "分析要求"
    gf.cell(1, 1).text = "油漆件、电镀件<0.03mm，高光<0.01mm"
    gf.cell(2, 0).text = "锁模说明"
    gf.cell(2, 1).text = "注塑压力OK，请选择合适的注塑机台吨位。"
    cleared = pb.clear_stale_conclusion_rows(prs)
    assert cleared == 2
    assert gf.cell(0, 1).text.strip() == ""
    assert gf.cell(2, 1).text.strip() == ""
    assert "0.03mm" in gf.cell(1, 1).text, "分析要求(通用标准)不应被清空"


def test_material_fields_mapping(tmp):
    """material_fields.json (VBS 官方字段枚举原始导出) → material_info.json 键值映射。"""
    fields = {
        "prop_name": "Generic PP",
        "prop_type": 21000,
        "material_id": "30012",
        "material_file": "abc.mdb",
        "fields": [
            {"id": 1998, "desc": "贸易名称", "values": "PP K4220"},
            {"id": 1999, "desc": "系列名称", "values": "polypropylene"},
            {"id": 20001, "desc": "制造商", "values": "LyondellBasell"},
            {"id": 1808, "desc": "模具温度范围(推荐)", "values": "20|80"},
            {"id": 1807, "desc": "模具表面温度", "values": "60"},
            {"id": 1800, "desc": "熔体温度范围(推荐)", "values": "180|260"},
            {"id": 1801, "desc": "熔体温度", "values": "230"},
            {"id": 1805, "desc": "绝对最大熔体温度", "values": "280"},
            {"id": 1504, "desc": "顶出温度", "values": "124"},
            {"id": 1804, "desc": "最大剪切应力", "values": "0.25"},
            {"id": 1806, "desc": "最大剪切速率", "values": "100000"},
        ],
    }
    with open(os.path.join(tmp, "material_fields.json"), "w", encoding="utf-8") as f:
        json.dump(fields, f, ensure_ascii=False)
    info = pb.build_material_info(tmp)
    assert info["trade_name"] == "PP K4220"
    assert info["family_name"] == "polypropylene"
    assert info["manufacturer"] == "LyondellBasell"
    assert info["mold_temp_min"] == 20 and info["mold_temp_max"] == 80
    assert info["mold_temp_rec"] == 60
    assert info["melt_temp_min"] == 180 and info["melt_temp_max"] == 260
    assert info["melt_temp_rec"] == 230
    assert info["melt_temp_max_abs"] == 280
    assert info["ejection_temp"] == 124
    assert info["material_id"] == "30012"
    assert os.path.exists(os.path.join(tmp, "material_info.json"))


def test_material_legacy_ids_fallback(tmp):
    """描述缺失/不认识时 → 已知 ID 白名单回退补数值。
    1801 双材料实机交叉验证为绝对最大熔体温度 (通用 PP 与 Novodur 均为 300,
    高于各自推荐上限的安全天花板), 2026-09-08 起采信映射; 未知 ID 不得映射。"""
    fields = {
        "prop_name": "",
        "prop_type": 21000,
        "material_id": "",
        "fields": [
            {"id": 1808, "desc": "", "values": "20|80"},
            {"id": 1800, "desc": "unknown gibberish", "values": "180|260"},
            {"id": 1504, "desc": "", "values": "124"},
            {"id": 1801, "desc": "", "values": "300"},
            {"id": 11002, "desc": "", "values": "220"},
            {"id": 11108, "desc": "", "values": "50"},
            {"id": 9999, "desc": "", "values": "666"},
        ],
    }
    with open(os.path.join(tmp, "material_fields.json"), "w", encoding="utf-8") as f:
        json.dump(fields, f, ensure_ascii=False)
    info = pb.build_material_info(tmp)
    assert info["mold_temp_min"] == 20 and info["mold_temp_max"] == 80
    assert info["melt_temp_min"] == 180 and info["melt_temp_max"] == 260
    assert info["ejection_temp"] == 124
    assert info["melt_temp_max_abs"] == 300, "1801 (绝对最大熔体温度) 未映射"
    assert info["melt_temp_rec"] == 220, "11002 (推荐熔体温度) 未映射"
    assert info["mold_temp_rec"] == 50, "11108 (推荐模具温度) 未映射"
    assert info.get("material_id") == "" and info["data_complete"] is False


def test_material_name_split(tmp):
    """材料全名 "牌号 : 制造商" 拆分 (VBS 经求解日志锁定真材料后导出)。
    实机样本: "Novodur HH-106 : INEOS Styrolution" → 牌号/制造商各得其所。"""
    fields = {
        "prop_name": "Novodur HH-106 : INEOS Styrolution",
        "material_name": "Novodur HH-106 : INEOS Styrolution",
        "prop_type": 21000,
        "material_id": "2",
        "fields": [],
    }
    with open(os.path.join(tmp, "material_fields.json"), "w", encoding="utf-8") as f:
        json.dump(fields, f, ensure_ascii=False)
    info = pb.build_material_info(tmp)
    assert info["trade_name"] == "Novodur HH-106", (
        f"牌号拆分错误: {info.get('trade_name')}"
    )
    assert info["manufacturer"] == "INEOS Styrolution", (
        f"制造商拆分错误: {info.get('manufacturer')}"
    )
    assert info["material_id"] == "2" and info["data_complete"] is True

    # 旧格式 (无 material_name, 仅 prop_name) 同样拆分; 无分隔符则牌号=全名
    fields2 = {"prop_name": "单名材料", "prop_type": 21000, "fields": []}
    with open(os.path.join(tmp, "material_fields.json"), "w", encoding="utf-8") as f:
        json.dump(fields2, f, ensure_ascii=False)
    info2 = pb.build_material_info(tmp)
    assert info2["trade_name"] == "单名材料"
    assert "manufacturer" not in info2, "无分隔符材料名不应臆造制造商"


def test_material_id_sentinel(tmp):
    """material_id 为 -1/0 哨兵 → 视为缺失 (实发事故: 材料卡显示 -1)。"""
    fields = {"prop_name": "X", "prop_type": 21000, "material_id": "-1", "fields": []}
    with open(os.path.join(tmp, "material_fields.json"), "w", encoding="utf-8") as f:
        json.dump(fields, f, ensure_ascii=False)
    info = pb.build_material_info(tmp)
    assert info["material_id"] == "" and info["data_complete"] is False


def test_vbs_no_hardcoded_material(tmp):
    """VBS 材料链路静态断言: 实机挂起的 MaterialSelector 调用已禁用;
    材料锁定走属性表枚举 + 求解日志匹配 (GetFirstProperty 只取第一项 = 通用 PP 残留,
    实机取证 2026-09-08: 真材料 Novodur ID=2 在日志 097_1_____~1.out 中记载);
    材料曲线 CreateMaterialPlot 第二参 = 属性 ID (锁定值, 失败回退 1, 禁用 0);
    ShowPlot 后 Regenerate + 跳最后一帧 ShowPlotFrame (陈旧帧/数据条串台回归);
    模型导出前 Viewer.Fit + 等待 (割裂回归); XY 曲线走官方 SaveXYPlotCurveData。"""
    vbs = open(os.path.join(ROOT, "AutoReport.vbs"), "rb").read().decode("gbk")
    assert "MatID = 21000" not in vbs, "VBS 仍硬编码材料ID 21000 (臆造材料根因)"
    assert "MatID = 20030" not in vbs, "VBS 仍硬编码材料ID 20030"
    assert "Set MatSel" not in vbs, "MaterialSelector 调用未移除 (实机挂起 16 分钟)"
    assert "GetFirstProperty" in vbs, "材料属性未走 GetFirstProperty 官方枚举"
    assert "GetNextPropertyOfType" in vbs, (
        "未枚举完整属性表 (首项=通用PP残留, 取不到真材料)"
    )
    # 结果图查找只认精确名 (用户裁决 2026-09-08: 别名/模糊匹配/数据集ID猜测彻底移除)
    assert "pObj.aliases" not in vbs, "VBS 仍在使用别名列表查找结果图"
    assert "CreatePlotByDsID" not in vbs, "VBS 仍在按数据集 ID 猜测创建变形图"
    assert "FindDatasetByID" not in vbs, "VBS 仍在探测数据集 ID (猜测类查找)"
    assert vbs.count("FindPlotByName") >= 1, "结果图未走 FindPlotByName 精确查找"
    assert "FindMaterialInSolverLogs" in vbs, "缺少求解日志材料锁定 (材料取错根因)"
    assert "IsGenericDefaultName" in vbs, "缺少通用默认材料识别"
    assert "SaveXYPlotCurveData" in vbs, "XY 曲线未走官方 SaveXYPlotCurveData 数据导出"
    # 材料曲线: 首选锁定 ID, 失败回退 1; 0 会静默不产出 (实机), 禁止出现
    assert vbs.count("ExportMaterialPlotSafe(MatDBCode, MatUseIdx, 1310") == 1, (
        "粘度曲线未按锁定材料 ID 导出"
    )
    assert vbs.count("ExportMaterialPlotSafe(MatDBCode, MatUseIdx, 1004") == 1, (
        "PVT 曲线未按锁定材料 ID 导出"
    )
    assert "Array(CLng(idx), 1)" in vbs, "材料曲线缺少回退 ID=1 的保底路径"
    assert "CreateMaterialPlot(CLng(dbCode), CLng(attemptIdx), CLng(fieldId))" in vbs
    for zero in (
        "CreateMaterialPlot(MatDBCode, 0,",
        "CreateMaterialPlot(CLng(dbCode), 0,",
    ):
        assert zero not in vbs, "材料曲线出现 0 序号 (2023 实机静默无产出)"
    # 文件存在性校验在 ExportMaterialPlotSafe 内统一执行
    assert "If FSO.FileExists(outPath) Then ok = True" in vbs, (
        "材料曲线导出缺少文件存在性校验 (假日志回归)"
    )
    assert vbs.count("PlotObj.Regenerate") >= 2, (
        "ShowPlot 后未强制 Regenerate (视口截图陈旧帧/数据条串台回归)"
    )
    assert "ShowPlotFrame PlotObj, FrameTotal - 1" in vbs, (
        "未跳转到最后一帧 (官方导出模式, 数据条串台回归)"
    )
    assert "Viewer.Fit" in vbs, "模型图导出前未 Fit (封面局部裁切回归)"
    assert vbs.count("Call SleepSec(1)") >= 4, (
        "Fit/ShowPlot 后缺少视口重绘等待 (模型割裂/陈旧帧回归)"
    )
    # 数据 JSON 启动清理 (防上个产品数据漏入本次报告)
    for j in [
        "material_info.json",
        "material_fields.json",
        "mesh_summary.json",
        "peak_values.json",
        "clamp_force_info.json",
        "manifest.json",
        "solid_model_cover.jpg",
    ]:
        assert f'"{j}"' in vbs, f"VBS 启动清理缺少 {j}"


def test_resolve_peaks_none(tmp):
    """导出与 config 均无 → None (调用方跳过探针, 绝不臆造)。"""
    r = pb.resolve_peaks({}, {})
    assert r["clamp_force"] is None and r["inj_pressure"] is None


def test_probe_annotated_from_peaks(tmp):
    """像素级: 峰值传入 → 黄色探针绘制; 缺失 → 图面保持无探针。"""
    import numpy as np
    from PIL import Image

    from core import image_processor as ip

    data_dir = os.path.join(tmp, "data")
    os.makedirs(data_dir)
    os.makedirs(os.path.join(tmp, "mode_b"))  # 避免入口早退
    xy_path = os.path.join(data_dir, "inj_pressure_xy.png")
    Image.new("RGB", (400, 300), (255, 255, 255)).save(xy_path)

    def yellow_pixels(path):
        arr = np.array(Image.open(path).convert("RGB"))
        return int(
            (
                (arr[:, :, 0] > 240)
                & (arr[:, :, 1] > 240)
                & (arr[:, :, 2] > 180)
                & (arr[:, :, 2] < 225)
            ).sum()
        )

    peaks = {"clamp_force": None, "inj_pressure": (21.45, 5.092)}
    ip.process_all_mode_b_plots(
        os.path.join(tmp, "mode_b"), data_dir=data_dir, peaks=peaks
    )
    assert yellow_pixels(xy_path) > 100, "峰值传入但探针未绘制"

    Image.new("RGB", (400, 300), (255, 255, 255)).save(xy_path)
    peaks2 = {"clamp_force": None, "inj_pressure": None}
    ip.process_all_mode_b_plots(
        os.path.join(tmp, "mode_b"), data_dir=data_dir, peaks=peaks2
    )
    assert yellow_pixels(xy_path) == 0, "峰值缺失时不应绘制探针"


def test_vbs_mode_dispatch(tmp):
    """VBS 静态断言: BOTH/ALL 模式分发条件正确 (T3 回归)。

    GUI 只产生 A/B/BOTH; VBS 必须覆盖 BOTH (历史 bug: 重复条件致 BOTH 全跳过)。
    mode_a 视口抓取在所有模式下都导出 (方案 B 拼合依赖其提取色带/坐标系)。
    """
    vbs = open(os.path.join(ROOT, "AutoReport.vbs"), "rb").read().decode("gbk")
    assert vbs.count('ScreenshotMode = "B" Or ScreenshotMode = "B"') == 0, (
        "重复的 'B' 条件回归 (BOTH 模式将不导出 mode_b 图像)"
    )
    assert 'ScreenshotMode = "BOTH"' in vbs, "VBS 未覆盖 BOTH 模式"
    cond_a = 'If ScreenshotMode = "A" Or ScreenshotMode = "B" Or ScreenshotMode = "BOTH" Or ScreenshotMode = "ALL" Then'
    cond_b = 'If ScreenshotMode = "B" Or ScreenshotMode = "BOTH" Or ScreenshotMode = "ALL" Then'
    assert cond_a in vbs, "方案 A 视口导出条件缺 BOTH"
    assert cond_b in vbs, "方案 B 独立导出条件缺 BOTH"


def test_ref_triad_regenerated(tmp):
    """ref_triad.png 每次运行强制重生成, 不得沿用上次方案缓存。"""
    import numpy as np
    from PIL import Image

    from core import image_processor as ip

    data_dir = tmp
    mode_a = os.path.join(data_dir, "mode_a")
    mode_b = os.path.join(data_dir, "mode_b")
    os.makedirs(mode_a)
    os.makedirs(mode_b)
    # 视口图: 400x300 白底 + 右下角黑色像素 (触发布局提取)
    va = np.full((300, 400, 3), 255, dtype=np.uint8)
    va[280:295, 350:395] = 0
    Image.fromarray(va).save(os.path.join(mode_a, "pressure.png"))

    ip.process_all_mode_b_plots(mode_b, data_dir=data_dir)
    ref = os.path.join(mode_b, "ref_triad.png")
    assert os.path.exists(ref), "首次运行应生成 ref_triad"
    # 模拟"上次方案残留": 把 ref 覆盖成纯白
    Image.new("RGB", (50, 50), (255, 255, 255)).save(ref)
    ip.process_all_mode_b_plots(mode_b, data_dir=data_dir)
    arr = np.array(Image.open(ref).convert("RGB"))
    assert not (arr == 255).all(), "ref_triad 未重生成 (仍在用上次方案缓存)"


def test_legend_source_prefers_viewport(tmp):
    """数据条来源回归 (实发事故: SavePlotScaleImage 每次导出同一条, 串台):
    视口图 (含标题块的完整图例) 必须优先于 scale_*.png。"""
    import numpy as np
    from PIL import Image

    from core import image_processor as ip

    data_dir = tmp
    mode_a = os.path.join(data_dir, "mode_a")
    mode_b = os.path.join(data_dir, "mode_b")
    os.makedirs(mode_a)
    os.makedirs(mode_b)

    # 视口图: 白底 + 左侧红色图例带 + 右下黑色坐标块 (触发布局提取)
    va = np.full((400, 600, 3), 255, dtype=np.uint8)
    va[20:380, 20:200] = (200, 30, 30)  # 红色图例
    va[350:395, 500:590] = 0  # 坐标系
    Image.fromarray(va).save(os.path.join(mode_a, "pressure.png"))

    # scale 文件: 绿色条 (若被误用, 拼合图例区呈绿色)
    sc = np.full((1100, 200, 3), 255, dtype=np.uint8)
    sc[10:1090, 10:190] = (30, 180, 30)
    Image.fromarray(sc).save(os.path.join(mode_b, "scale_pressure.png"))

    # 模型图: 白底 + 中央蓝块
    md = np.full((800, 600, 3), 255, dtype=np.uint8)
    md[200:600, 200:400] = (30, 30, 200)
    Image.fromarray(md).save(os.path.join(mode_b, "model_pressure.png"))

    out = os.path.join(mode_b, "pressure.png")
    ok = ip.merge_scale_and_model(
        os.path.join(mode_b, "scale_pressure.png"),
        os.path.join(mode_b, "model_pressure.png"),
        out,
        viewport_path=os.path.join(mode_a, "pressure.png"),
        target_height=900,
    )
    assert ok
    comp = np.array(Image.open(out).convert("RGB"))
    # 拼合画布上数据条粘贴于 (35,35): 该区域应为红色 (视口源) 而非绿色 (scale 源)
    samples = [comp[120, 60], comp[200, 60], comp[300, 80]]
    red = sum(1 for r, g, b in samples if r > 120 and g < 100 and b < 100)
    green = sum(1 for r, g, b in samples if g > 120 and r < 100)
    assert red >= 2, f"数据条未取视口源 (取样点: {samples})"
    assert green == 0, "数据条错误地取了 scale_*.png 源"


def test_solid_blank_fallback(tmp):
    """封面回退 (round 2 改写): FUSION 隐藏 T 层致 VBS 导出全白空图时,
    旧版走灰色 CAD 重绘制兜底 (用户裁决该样式与真实实体不符, 已废弃);
    新版走 resolve_cover_image 直接选 mode_b/model_pressure.png 等候选直出,
    全白 + 候选缺失 → 返回 None (调用方登记缺失), 不得臆造任何图。"""
    import numpy as np
    from PIL import Image

    from core import image_processor as ip

    data_dir = tmp
    mode_b = os.path.join(data_dir, "mode_b")
    os.makedirs(mode_b)

    # 全白 solid_model.png + 无任何 model_*.png 候选 → resolve_cover_image 返回 None
    Image.new("RGB", (640, 360), (255, 255, 255)).save(
        os.path.join(data_dir, "solid_model.png")
    )
    assert ip.resolve_cover_image(data_dir) is None, (
        "全白 + 无候选, resolve_cover_image 不得臆造内容"
    )

    # 提供可用 model_pressure.png 候选 → 应被选中直出并写回 solid_model.png
    md = np.full((300, 400, 3), 255, dtype=np.uint8)
    md[80:220, 120:280] = (120, 130, 140)
    Image.fromarray(md).save(os.path.join(mode_b, "model_pressure.png"))
    chosen = ip.resolve_cover_image(data_dir)
    assert chosen, "有候选时 resolve_cover_image 必须返回路径"
    # solid_model.png 应被候选内容覆写 (不再依赖灰色重绘)
    arr = np.array(Image.open(os.path.join(data_dir, "solid_model.png")).convert("L"))
    assert (arr < 240).mean() > 0.005, "命中候选后 solid_model.png 未被正确覆写"

    # 旧 generate_solid_cad_model 已停用, 不得再被 process_all_mode_b_plots 调用
    import inspect

    src = inspect.getsource(ip.process_all_mode_b_plots)
    assert "generate_solid_cad_model" not in src, (
        "process_all_mode_b_plots 仍在调用已废弃的灰色重绘函数"
    )


def test_missing_placeholder_created(tmp):
    """缺图占位图可生成且非空白。"""
    import numpy as np
    from PIL import Image

    ph = pb.ensure_missing_plot_placeholder(tmp)
    assert os.path.exists(ph)
    arr = np.array(Image.open(ph).convert("RGB"))
    assert arr.shape[0] >= 300 and (arr < 200).any(), "占位图应含可见文字/边框"


def test_check_env(tmp):
    """check_env: 配置缺失报问题; 模板标记存在/缺失两种路径。"""
    import check_env as ce

    problems, _ = ce.check_config(os.path.join(tmp, "nope.json"))
    assert problems, "缺失配置应报问题"

    from pptx import Presentation

    p = Presentation()
    s1 = p.slides.add_slide(p.slide_layouts[6])
    s1.shapes.add_textbox(0, 0, 100, 50).text_frame.text = "报告日期 Moldflow"
    s2 = p.slides.add_slide(p.slide_layouts[6])
    s2.shapes.add_textbox(0, 0, 100, 50).text_frame.text = "实体计数 三角形"
    tpl = os.path.join(tmp, "tpl.pptx")
    p.save(tpl)
    problems, infos = ce.check_template(tpl)
    assert not problems, problems

    p2 = Presentation()
    p2.slides.add_slide(p2.slide_layouts[6])
    p2.slides.add_slide(p2.slide_layouts[6])
    tpl2 = os.path.join(tmp, "tpl2.pptx")
    p2.save(tpl2)
    problems2, _ = ce.check_template(tpl2)
    assert problems2, "缺标记的模板应报问题"


def test_load_config_corrupt(tmp):
    """配置损坏 → ({}, 错误信息) 兜底, 不崩溃 (T28)。"""
    import unittest.mock as mock

    import config_gui

    with mock.patch.object(config_gui, "CONFIG_FILE", os.path.join(tmp, "no.json")):
        cfg, err = config_gui.load_config()
    assert cfg == {} and err is None  # 缺失 = 正常默认

    poisoned = os.path.join(tmp, "poison.json")
    with open(poisoned, "w", encoding="utf-8") as f:
        f.write("{broken json")
    with mock.patch.object(config_gui, "CONFIG_FILE", poisoned):
        cfg, err = config_gui.load_config()
    assert cfg == {} and err, "损坏配置应返回错误信息而非崩溃"


def test_extract_new_log_text_split_gbk(tmp):
    """日志 tail 分块读取: GBK 双字节字符跨块不乱码不崩溃 (T27/A10)。"""
    import config_gui

    path = os.path.join(tmp, "run.log")
    full = "方案A截图完成".encode("gbk")
    with open(path, "wb") as f:
        f.write(full)
    # 模拟逐字节到达: 增量解码器持有跨块半字符, 拼接结果必须与整读一致
    collected = ""
    offset = 0
    decoder = None
    for _ in range(len(full)):
        text, offset, decoder = config_gui.extract_new_log_text(path, offset, decoder)
        collected += text
        assert offset <= len(full)
    assert collected == "方案A截图完成", f"分块拼接结果异常: {collected!r}"
    # 文件重建 (变小) → 解码器重置从头读
    with open(path, "wb") as f:
        f.write("新".encode("gbk"))
    text, offset, _ = config_gui.extract_new_log_text(path, len(full) + 10, decoder)
    assert text == "新" and offset == 2


def test_fail_policy(tmp):
    """on_missing_data=fail 且存在缺失 → 必须抛错; annotate → 放行。"""
    try:
        pb.enforce_missing_policy(["网格统计: 表面面积"], "fail")
        raise AssertionError("fail 模式未抛错")
    except RuntimeError:
        pass
    assert pb.enforce_missing_policy([], "fail") is None
    assert pb.enforce_missing_policy(["x"], "annotate") is None


def test_canvas_adaptive_height(tmp):
    """T8: 画布高度取期望值(1.58:1); 源不足降级绝不放大; 0x0 语义=1215。"""
    import numpy as np
    from PIL import Image

    from core import image_processor as ip

    data_dir = tmp
    mode_b = os.path.join(data_dir, "mode_b")
    os.makedirs(mode_b)
    # 800px 高的"模型"图 (白底+黑块) 与"色带"图 (带深色内容, 模拟真实图例文字;
    # 纯白图例会被空白数据条护栏拒用)
    Image.new("RGB", (600, 800), (255, 255, 255)).save(
        os.path.join(mode_b, "model_pressure.png")
    )
    scale_im = Image.new("RGB", (200, 1100), (255, 255, 255))
    from PIL import ImageDraw as _ID

    _ID.Draw(scale_im).rectangle([10, 10, 190, 300], fill=(30, 30, 30))
    scale_im.save(os.path.join(mode_b, "scale_pressure.png"))
    out = os.path.join(mode_b, "pressure.png")
    ok = ip.merge_scale_and_model(
        os.path.join(mode_b, "scale_pressure.png"),
        os.path.join(mode_b, "model_pressure.png"),
        out,
        target_height=1440,
    )
    assert ok
    with Image.open(out) as im:
        assert im.height == 870 and abs(im.width / im.height - 1.58) < 0.01, (
            f"源不足时应降为 800+70=870: got {im.size}"
        )

    # 源充足 → 画布 = 期望高度
    Image.new("RGB", (1200, 1400), (255, 255, 255)).save(
        os.path.join(mode_b, "model_pressure.png")
    )
    ok = ip.merge_scale_and_model(
        os.path.join(mode_b, "scale_pressure.png"),
        os.path.join(mode_b, "model_pressure.png"),
        out,
        target_height=1440,
    )
    with Image.open(out) as im:
        assert im.height == 1440 and im.width == int(1440 * 1.58)


def test_gif_global_palette(tmp):
    """T11: 全帧共用调色板 (逐帧量化在 Pillow 内部已由 palette= 保证一致),
    优化后的 GIF 帧数与循环标记正确。"""
    from PIL import Image

    from core import gif_enhancer as ge

    gif_in = os.path.join(tmp, "in.gif")
    frames = [
        Image.new("RGB", (80, 60), c) for c in [(255, 0, 0), (0, 255, 0), (0, 0, 255)]
    ]
    frames[0].save(
        gif_in, save_all=True, append_images=frames[1:], duration=100, loop=0
    )
    out = ge.optimize_existing_gif(gif_in, target_delay_ms=80)
    assert out == gif_in
    with Image.open(gif_in) as im:
        n = 1
        try:
            while True:
                im.seek(im.tell() + 1)
                n += 1
        except EOFError:
            pass
    assert n == 3, f"帧数丢失: {n}"


def test_sanitize_filename(tmp):
    """T15: 非法字符/路径逃逸/连续下划线净化, 保留中文。"""
    assert pb.sanitize_filename('a<b>c:"d"') == "a_b_c_d"
    assert "..\\evil" not in pb.sanitize_filename("..\\evil")
    r = pb.sanitize_filename("054_5_hf___2026(1)______方案")
    assert "__" not in r
    assert pb.sanitize_filename("模流方案A") == "模流方案A"
    assert pb.sanitize_filename("") == "方案"
    assert len(pb.sanitize_filename("x" * 200)) <= 80


def test_gif_caps(tmp):
    """T25: 帧数/高度上限生效。"""
    from PIL import Image

    from core import gif_enhancer as ge

    gif_in = os.path.join(tmp, "cap.gif")
    frames = [Image.new("RGB", (400, 300), (i * 5, 100, 200)) for i in range(10)]
    frames[0].save(gif_in, save_all=True, append_images=frames[1:], duration=50, loop=0)
    ge.optimize_existing_gif(gif_in, target_delay_ms=80, max_frames=4, max_height=100)
    with Image.open(gif_in) as im:
        assert im.n_frames == 4, f"帧数上限未生效: {im.n_frames}"
        assert im.size[1] == 100, f"高度上限未生效: {im.size}"


def test_vbs_syntax_compile(tmp):
    """VBS 全文编译门禁 (实机 800A03EA 事故回归)。

       静态特征: 行内无孤立 CR (Python
    转义残留) / 无未闭合字符串 / 字符串区外无 '--'。
       编译门禁: 复制脚本在 Option Explicit 后插入 WScript.Quit 0 交 cscript 解析 —
       VBScript 先整体编译再执行, 任何语法错误都会在执行 Quit 前报出, 纯语法检查不跑业务。
    """
    import subprocess

    vbs_path = os.path.join(ROOT, "AutoReport.vbs")
    body = open(vbs_path, "rb").read().decode("gbk")
    lines = body.split("\r\n")
    cr = "\r"
    for idx, l in enumerate(lines, 1):
        assert cr not in l, f"VBS L{idx} 行内嵌孤立 CR (Python 转义残留)"
        if l.strip().startswith("'"):
            continue
        in_str, outside, k = False, [], 0
        while k < len(l):
            c = l[k]
            if in_str:
                if c == '"':
                    if k + 1 < len(l) and l[k + 1] == '"':
                        k += 2
                        continue
                    in_str = False
            else:
                if c == '"':
                    in_str = True
                else:
                    outside.append(c)
            k += 1
        assert not in_str, f"VBS L{idx} 字符串未闭合"
        assert "--" not in "".join(outside), (
            f"VBS L{idx} 字符串区外含 '--' (VBS 引号转义损坏特征, 引号应双写)"
        )
    assert lines[7].strip() == "Option Explicit"
    wrapped = os.path.join(tmp, "syntax_check.vbs")
    with open(wrapped, "wb") as f:
        f.write(("\r\n".join(lines[:8] + ["WScript.Quit 0"] + lines[8:])).encode("gbk"))
    r = subprocess.run(
        ["cscript", "//nologo", wrapped], capture_output=True, timeout=60
    )
    out = (r.stdout + r.stderr).decode("gbk", errors="replace")
    assert r.returncode == 0, f"VBS 编译失败: {out[:300]}"


def test_left_bar_cutoff_drops_model_fragment(tmp):
    """方案 B 割裂回归 (2026-09-08 实发): 左侧 340px 数据条带混入模型左缘碎片,
    拼合后被贴回画布左端形成与主体分离的月牙。必须按色条定位 + 白隙切断。"""
    import numpy as np
    from PIL import Image

    from core import image_processor as ip

    w, h = 800, 600
    img = np.full((h, w, 3), 255, dtype=np.uint8)
    # 色条 (高饱和) x[60-100), y[50-500)
    for x in range(60, 100):
        img[50:500, x] = (255, int(255 * (x - 60) / 40), 0)
    # 数值标签 (深色) x[105-140)
    img[200:215, 105:140] = (0, 0, 0)
    img[480:495, 105:140] = (0, 0, 0)
    # 白隙 x[145-200)
    # 模型左缘碎片 (高饱和) x[205-330)
    img[150:420, 205:330] = (30, 90, 220)
    path = os.path.join(tmp, "viewport.png")
    Image.fromarray(img).save(path)

    left_bar, _triad, _title = ip.extract_viewport_components(path)
    assert left_bar is not None, "左侧数据条未提取到"
    arr = np.asarray(left_bar.convert("RGB")).astype(np.int16)
    sat_cols = ((arr.max(axis=2) - arr.min(axis=2)) > 60).mean(axis=0)
    tail = sat_cols[int(left_bar.width * 0.8) :]
    assert not (tail > 0.12).any(), (
        f"left_bar 尾部仍残留模型碎片 (宽 {left_bar.width}, 尾部饱和列 {int((tail > 0.12).sum())})"
    )
    assert left_bar.width < 205, f"切断失败 (left_bar 宽 {left_bar.width} 覆盖到碎片区)"


def test_curve_peak_axes_frame_locates_apex(tmp):
    """方案 B 探针回归 (2026-09-08 实发): 裁剪后峰顶落在 ROI 左侧之外 → 回退
    default_pos 使探针框浮空。轴框定位必须命中真实峰顶, 且不被标题文字干扰。"""
    import numpy as np
    from PIL import Image

    from core import image_processor as ip

    w, h = 800, 500
    img = np.full((h, w, 3), 255, dtype=np.uint8)
    # y 轴 / x 轴 (深灰: 灰度 100, 纯黑阈值 40 检不到)
    img[60:450, 100:102] = (100, 100, 100)
    img[450:452, 100:750] = (100, 100, 100)
    # 标题文字 (横跨绘图区顶线, 高于峰顶 — 3px 余量曾误检为此块)
    img[55:85, 400:470] = (20, 20, 20)
    # 轴外刻度文字 (左/下)
    img[200:212, 40:90] = (20, 20, 20)
    img[460:472, 200:260] = (20, 20, 20)
    # 曲线: 峰顶在 x=120 (ROI x 起点 0.18*800=144 之外, 复现实发失败场景)
    apex = (120, 150)
    for x in range(102, 400):
        y = 150 + int((x - 120) ** 2 / 90)
        if 60 <= y < 450:
            img[y : y + 2, x] = (100, 100, 100)

    px, py, method = ip.locate_curve_peak(
        Image.fromarray(img), (0.18, 0.40), (0.08, 0.32), (0.265, 0.113)
    )
    assert method == "axes-frame", f"应走轴框定位, 实际 method={method}"
    assert abs(px - apex[0]) <= 12 and abs(py - apex[1]) <= 12, (
        f"探针未落在峰顶: 得到 ({px},{py}), 期望约 {apex}"
    )


def test_no_4k_option(tmp):
    """4K 已移除 (SaveImage3 4K 断带缺陷): GUI 无 4K 选项且旧配置被钳到 1080P;
    VBS 侧同样有钳制守卫。"""
    gui_src = open(os.path.join(ROOT, "config_gui.py"), encoding="utf-8").read()
    assert "3840x2160 (4K" not in gui_src, "config_gui 仍提供 4K 选项"
    assert "w >= 3840 or h >= 2160" in gui_src, "config_gui 缺少 4K 钳制守卫"
    vbs_src = open(
        os.path.join(ROOT, "AutoReport.vbs"), encoding="gbk", errors="replace"
    ).read()
    assert "ImageWidth >= 3840 Or ImageHeight >= 2160" in vbs_src, (
        "AutoReport.vbs 缺少 4K 钳制守卫"
    )


def test_no_duplicate_rerun_after_completion(tmp):
    """二次触发回归 (严重缺陷): 关闭完成弹窗后不得再次拉起生成流程。
    - GUI: 无全局 <Return> 绑定 (Enter 曾重跑流水线) + 完成后写 gui_run_done.txt
    - VBS: 拉起 GUI 返回后检测该标记即退出 (外层实例曾整条流水线再跑一遍)"""
    gui_src = open(os.path.join(ROOT, "config_gui.py"), encoding="utf-8").read()
    assert 'self.root.bind("<Return>"' not in gui_src, "GUI 仍绑定 Enter 触发生成"
    assert "on_enter_key" not in gui_src.split("# 注: 不绑定全局")[0], (
        "on_enter_key 实现未清除"
    )
    assert "gui_run_done.txt" in gui_src, "GUI 未写完成标记 gui_run_done.txt"
    vbs_src = open(
        os.path.join(ROOT, "AutoReport.vbs"), encoding="gbk", errors="replace"
    ).read()
    assert "gui_run_done.txt" in vbs_src and "WScript.Quit 0" in vbs_src, (
        "VBS 未完成标记守卫 (外层实例仍会重复生成)"
    )


def test_mode_b_width_constraint_no_crop(tmp):
    """方案 B 拼合 width-only 缩放回归 (2026-09-08 round 2 实发):
    模型按 height 缩放到可用高度, 还要校验 width 不超过右侧剩余空间, 否则溢出
    右侧并被裁掉 (模型截图不完整)。"""
    import numpy as np
    from PIL import Image

    from core import image_processor as ip

    data_dir = tmp
    mode_a = os.path.join(data_dir, "mode_a")
    mode_b = os.path.join(data_dir, "mode_b")
    os.makedirs(mode_a)
    os.makedirs(mode_b)

    # 模拟窄高比 (横长方形) 的模型: 1600x600, 宽远大于高
    # 期望: 在 rem_w 限制下, 模型缩放后宽不应超过剩余空间
    model = np.full((600, 1600, 3), 255, dtype=np.uint8)
    model[100:500, 100:1500] = (30, 30, 200)
    Image.fromarray(model).save(os.path.join(mode_b, "model_pressure.png"))

    # 视口图: 含色条 + 坐标系, 触发完整路径
    vp = np.full((900, 1600, 3), 255, dtype=np.uint8)
    vp[50:850, 50:200] = (200, 30, 30)  # 色条
    vp[800:850, 1200:1450] = 0  # 坐标系
    Image.fromarray(vp).save(os.path.join(mode_a, "pressure.png"))

    sc = np.full((900, 200, 3), 255, dtype=np.uint8)
    Image.fromarray(sc).save(os.path.join(mode_b, "scale_pressure.png"))

    out = os.path.join(mode_b, "pressure.png")
    ok = ip.merge_scale_and_model(
        os.path.join(mode_b, "scale_pressure.png"),
        os.path.join(mode_b, "model_pressure.png"),
        out,
        viewport_path=os.path.join(mode_a, "pressure.png"),
        target_height=900,
    )
    assert ok
    arr = np.array(Image.open(out).convert("RGB"))
    # 拼合画布宽高应约为 1.58:1 (target_height=900 → 1422, 源不足会降级)
    assert abs(arr.shape[1] / arr.shape[0] - 1.58) < 0.02, (
        f"画布宽高比异常: {arr.shape}"
    )
    # 关键: 模型宽度约束生效, 缩放后宽不应超过 rem_w。
    # 像素级检测: 模型右缘外的右侧 1% 留白带不应有模型蓝块 (容差 ≤ 2 px 抗噪)。
    # 注意: 测试图模型本体已贴近右侧边界, 留白带本身为空; 若宽约束失效,
    # 右端会出现被裁的蓝块 → 检测到溢出。
    canvas_w = arr.shape[1]
    right_strip = arr[:, int(canvas_w * 0.99) :]
    blue_in_right = ((right_strip[:, :, 2] > 120) & (right_strip[:, :, 0] < 100)).sum()
    # 同时反查 rem_x 计算: lb_w 大约 = scale 高 900 * 0.541 = ~108, rem_w 大约 580-700
    # 缩放后模型宽度被 rem_w 限制, 最右蓝像素 col 应 < canvas_w * 0.99
    blue_cols = np.where(((arr[:, :, 2] > 120) & (arr[:, :, 0] < 100)).any(axis=0))[0]
    if len(blue_cols) > 0:
        rightmost_blue = int(blue_cols.max())
        # 允许在最后 1.5% (抗 trim/resize 抗锯齿溢出)
        assert rightmost_blue <= int(canvas_w * 0.985), (
            f"模型右缘溢出画布: 最右蓝像素 col={rightmost_blue}, 画布宽={canvas_w} "
            f"(width-only 缩放未生效, m_w 超出 rem_w 后被 paste 截断)"
        )
    assert blue_in_right < 200, (
        f"模型右侧溢出 (right_strip 蓝色像素 {blue_in_right} 个, 宽度约束未生效)"
    )


def test_resolve_cover_image_priority(tmp):
    """封面直出优先级 (用户裁决 2026-09-08 round 2): 优先 mode_b/model_pressure.png
    → mode_b/model_volumetric_shrinkage.png → solid_model.png;
    命中后裁白边写回 (返回路径恒为 solid_model.png, 内容来自最佳候选)。"""
    import numpy as np
    from PIL import Image

    from core import image_processor as ip
    from core.image_processor import _model_image_usable as _usable

    mode_b = os.path.join(tmp, "mode_b")
    os.makedirs(mode_b)
    fallback = os.path.join(tmp, "solid_model.png")

    def _non_blank_content(path):
        """检查路径对应图非空白 (>=0.5% 暗像素), 用于确认 resolve_cover_image
        选中的候选确实有内容被写回 fallback。"""
        with Image.open(path) as im:
            arr = np.array(im.convert("L"))
        return (arr < 240).mean() > 0.005

    # 缺省三个候选都缺失 → 返回 None
    assert ip.resolve_cover_image(tmp) is None

    # 仅 solid_model.png 存在 (有内容) → 命中, 返回 fallback 路径
    sm = np.full((400, 600, 3), 255, dtype=np.uint8)
    sm[100:300, 200:400] = (30, 30, 200)
    Image.fromarray(sm).save(fallback)
    chosen = ip.resolve_cover_image(tmp)
    assert chosen and os.path.normcase(chosen) == os.path.normcase(fallback)
    assert _non_blank_content(chosen)

    # mode_b/model_pressure.png 也存在且可用 → 应优先选它 (内容写回 fallback)
    # 先把 fallback 设成全白, 然后看它是否被 model_pressure 的内容覆盖
    Image.fromarray(np.full((400, 600, 3), 255, dtype=np.uint8)).save(fallback)
    mp = np.full((500, 700, 3), 255, dtype=np.uint8)
    mp[100:400, 200:500] = (60, 60, 200)  # 蓝色块
    Image.fromarray(mp).save(os.path.join(mode_b, "model_pressure.png"))
    chosen2 = ip.resolve_cover_image(tmp)
    assert chosen2 and os.path.normcase(chosen2) == os.path.normcase(fallback)
    # fallback 应被 model_pressure 覆写, 内容非白
    assert _non_blank_content(fallback), (
        "model_pressure.png 应被优先选中并写回 fallback, 但 fallback 仍为空白"
    )
    # 与原始 model_pressure 的 _model_image_usable 判据一致
    ok, _ = _usable(os.path.join(mode_b, "model_pressure.png"))
    assert ok, "model_pressure.png 自身应被判为可用"

    # 候选全白 (残缺) → 降级, 最终无内容返回 None
    Image.fromarray(np.full((400, 600, 3), 255, dtype=np.uint8)).save(
        os.path.join(mode_b, "model_pressure.png")
    )
    Image.fromarray(np.full((400, 600, 3), 255, dtype=np.uint8)).save(fallback)
    assert ip.resolve_cover_image(tmp) is None

    # model_volumetric_shrinkage 命中 → 应被选中 (作为 model_pressure 失效的降级)
    vs = np.full((400, 600, 3), 255, dtype=np.uint8)
    vs[100:300, 200:400] = (60, 200, 30)  # 绿色块
    Image.fromarray(vs).save(os.path.join(mode_b, "model_volumetric_shrinkage.png"))
    chosen3 = ip.resolve_cover_image(tmp)
    assert chosen3 and os.path.normcase(chosen3) == os.path.normcase(fallback)
    assert _non_blank_content(fallback), (
        "model_volumetric_shrinkage 应被选中作为降级, 但 fallback 内容未生效"
    )


def test_pptx_cover_uses_resolve_cover_image(tmp):
    """封面构建代码静态断言 (round 2): 必须调用 resolve_cover_image, 不再直接
    写死 solid_model.png (后者会被 Moldflow 图层隐藏场景下导出空图覆盖)。"""
    src = open(os.path.join(ROOT, "core", "pptx_builder.py"), encoding="utf-8").read()
    assert "resolve_cover_image(data_dir)" in src, (
        "build_single_report 封面段未调用 resolve_cover_image"
    )


def test_warp_page_stripping_block_removed(tmp):
    """warp 页剥离块已删除 (用户裁决 2026-09-08): 模板图片已手动清空, 运行时
    不再剥离; PPT 14-16 页保持模板自身内容。"""
    src = open(os.path.join(ROOT, "core", "pptx_builder.py"), encoding="utf-8").read()
    assert "[13, 14, 15]" not in src, (
        "warp 页图片剥离块残留 (用户已手动清空, 运行时不得再剥离)"
    )
    assert "Removed picture shape" not in src, "warp 页剥离日志残留"


def test_default_plot_keys_whitelist(tmp):
    """默认常用项白名单 (round 2): 共 10 项, 不含 warp_x/y/z (默认报告聚焦全部效应)。"""
    import config_gui

    assert len(config_gui.DEFAULT_PLOT_KEYS) == 10, (
        f"DEFAULT_PLOT_KEYS 应为 10 项: got {len(config_gui.DEFAULT_PLOT_KEYS)}"
    )
    for must in (
        "filling_animation",
        "pressure",
        "vp_switch_pressure",
        "inj_pressure_xy",
        "flow_front_temp",
        "clamp_force_xy",
        "weld_lines",
        "volumetric_shrinkage",
        "sink_marks",
        "warpage_all",
    ):
        assert must in config_gui.DEFAULT_PLOT_KEYS, f"缺失默认项: {must}"
    for forbid in ("warpage_x", "warpage_y", "warpage_z"):
        assert forbid not in config_gui.DEFAULT_PLOT_KEYS, (
            f"XYZ 分向变形不应在默认白名单: {forbid}"
        )


def test_gui_plot_rows_styled_not_foreground(tmp):
    """结果项行的更新必须统一走 _set_plot_hit_labels (ttk style 灰显)。

    回归 (实机崩溃): ttk.Checkbutton 没有 foreground 选项,
    w.configure(foreground=...) 即抛 TclError 'unknown option "-foreground"'。
    注: avail_label 是 ttk.Label, foreground 合法, 不在禁止范围。
    """
    src = open(os.path.join(ROOT, "config_gui.py"), encoding="utf-8").read()
    body = src.split("def apply_available_plots", 1)[1].split("\n    def ", 1)[0]
    assert "_set_plot_hit_labels" in body, "灰显更新未走 _set_plot_hit_labels 统一入口"
    assert "w.configure(" not in body, (
        "结果项行存在绕过 helper 的直配 (ttk.Checkbutton 无 foreground, 直配即崩)"
    )
    helper = src.split("def _set_plot_hit_labels", 1)[1].split("\n    def ", 1)[0]
    assert "foreground=" not in helper, "辅助方法仍配置颜色属性"
    assert "PlotMiss.TCheckbutton" in helper, "灰显样式未接入"
    assert 'self.style.configure("PlotMiss.TCheckbutton"' in src, (
        "PlotMiss.TCheckbutton 样式未在 ttk.Style 上定义"
    )


def test_compute_safe_box_fit_aspect(tmp):
    """等比适配几何: 严禁拉伸变形, 结果必须居中于安全框内。"""
    from PIL import Image

    wide = os.path.join(tmp, "wide.png")
    Image.new("RGB", (2000, 500), (1, 2, 3)).save(wide, "PNG")
    # 宽图 (4:1) 入 1000x500 框 → 宽受限: 1000x250, 垂直居中
    l, t, w, h = pb.compute_safe_box_fit(wide, 0, 0, 1000, 500)
    assert (l, t, w, h) == (0, 125, 1000, 250), f"宽图适配错误: {(l, t, w, h)}"

    tall = os.path.join(tmp, "tall.png")
    Image.new("RGB", (500, 2000), (4, 5, 6)).save(tall, "PNG")
    # 高图 (1:4) 入 1000x500 框 → 高受限: 125x500, 水平居中
    l2, t2, w2, h2 = pb.compute_safe_box_fit(tall, 0, 0, 1000, 500)
    assert (l2, t2, w2, h2) == (437, 0, 125, 500), f"高图适配错误: {(l2, t2, w2, h2)}"

    assert pb.compute_safe_box_fit(os.path.join(tmp, "缺.png"), 0, 0, 100, 100) is None


def test_insert_or_replace_into_blank_slide(tmp):
    """模板页无任何图片形状时必须插入新图 (清空模板导出无图回归),
    几何落在安全框内且保持纵横比; 有主图时仍走替换而非叠加。"""
    from PIL import Image
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from pptx.util import Emu

    img = os.path.join(tmp, "plot.png")
    Image.new("RGB", (2000, 1000), (10, 20, 30)).save(img, "PNG")

    prs = Presentation()
    prs.slide_width = Emu(9144000)
    prs.slide_height = Emu(6858000)
    blank = prs.slides.add_slide(prs.slide_layouts[6])
    box = (Emu(914400), Emu(914400), Emu(7315200), Emu(4572000))
    assert pb.insert_or_replace_picture(blank, img, *box, log_tag="[测试]"), (
        "空白页插入失败"
    )
    pics = [s for s in blank.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]
    assert len(pics) == 1, f"空白页应恰好插入 1 张图, 实际 {len(pics)}"
    l, t, w, h = pics[0].left, pics[0].top, pics[0].width, pics[0].height
    b_l, b_t, b_w, b_h = box
    assert b_l <= l and b_t <= t and l + w <= b_l + b_w and t + h <= b_t + b_h, (
        f"插入图越出安全框: {(l, t, w, h)}"
    )
    assert w == 2 * h, f"纵横比被破坏: {w}/{h}"

    # 已有图片的页面: 替换主图而不是新增形状
    prs2 = Presentation()
    prs2.slide_width = Emu(9144000)
    prs2.slide_height = Emu(6858000)
    with_pic = prs2.slides.add_slide(prs2.slide_layouts[6])
    with_pic.shapes.add_picture(img, Emu(1828800), Emu(914400))
    before = len(with_pic.shapes)
    assert pb.insert_or_replace_picture(with_pic, img, *box)
    assert len(with_pic.shapes) == before, "替换路径新增了多余形状"


def test_material_quadrant_index(tmp):
    """Slide 3 四联图象限判定沿用模板历史分界 (4500000/3500000 EMU)。"""
    assert pb.material_quadrant_index(100000, 100000) == 0, "左上"
    assert pb.material_quadrant_index(5000000, 100000) == 1, "右上"
    assert pb.material_quadrant_index(100000, 4000000) == 2, "左下"
    assert pb.material_quadrant_index(5000000, 4000000) == 3, "右下"
    assert len(pb.MATERIAL_QUADRANT_KEYS) == 4
    assert len(pb.MATERIAL_QUADRANT_BOXES) == 4


def test_resolve_output_dir_for_run(tmp):
    """输出目录决策: 显式配置 > manifest.model_dir (目录须存在) > 项目根。"""
    model_dir = os.path.join(tmp, "模型项目目录")
    os.makedirs(model_dir)

    assert (
        pb.resolve_output_dir_for_run({"output_dir": model_dir}, {"model_dir": "X"})
        == model_dir
    ), "显式配置必须优先于 manifest"
    assert pb.resolve_output_dir_for_run({}, {"model_dir": model_dir}) == model_dir, (
        "留空时应输出到模型所在目录"
    )
    assert (
        pb.resolve_output_dir_for_run({}, {"model_dir": os.path.join(tmp, "不存在")})
        == pb.PROJECT_ROOT
    ), "model_dir 不存在时必须回退项目根"
    assert pb.resolve_output_dir_for_run({}, {}) == pb.PROJECT_ROOT, (
        "无 manifest 时必须回退项目根"
    )


def test_output_defaults_to_model_dir_chain(tmp):
    """输出链静态断言: builder 使用 resolve_output_dir_for_run,
    VBS manifest 写入 model_dir 字段 (GBK 编码完好)。

    B-01 回归 (2026-09-09 体检): model_dir 断言必须钉到 manifest.json 写点。
    此前裸子串断言误命中 mesh_summary.json 写点 (AutoReport.vbs 的 MeshJson
    也含 model_dir), 而 manifest.json 漏写该字段, 输出到模型目录的决策链
    在真实 VBS→Python 链路上从未生效 (A-01)。"""
    src = open(os.path.join(ROOT, "core", "pptx_builder.py"), encoding="utf-8").read()
    assert "resolve_output_dir_for_run(config, manifest)" in src, (
        "输出路径决策未接入 manifest.model_dir"
    )
    vbs = open(os.path.join(ROOT, "AutoReport.vbs"), "rb").read().decode("gbk")
    assert "ModelDir = CStr(ProjTmp.Path)" in vbs, "VBS 未捕获 Synergy.Project().Path"
    # 断言钉到 ManifestText 赋值段 (而非全文/Dim 行), 防再误命中 MeshJson 写点
    seg_start = vbs.index('ManifestText = "{"')
    seg_end = vbs.index('WriteUtf8TextFile TempDir & "\\manifest.json"', seg_start)
    manifest_seg = vbs[seg_start:seg_end]
    assert '""model_dir"": """ & EscapeJson(ModelDir)' in manifest_seg, (
        "AutoReport.vbs manifest.json 未写入 model_dir (输出目录链断裂)"
    )


def test_builder_silent_skips_registered(tmp):
    """B-02/B-03 回归 (2026-09-09 体检): 启用结果被跳过时必须登记缺失,
    不允许静默丢图 —
    - 追加页 (extra_plots) 缺图: 此前无 else 分支, 报告少页且 missing 不记录;
    - 14-16 保留页: 此前 continue 无任何记录, 用户勾选分配到这三页即静默蒸发。"""
    src = open(os.path.join(ROOT, "core", "pptx_builder.py"), encoding="utf-8").read()
    extra_seg = src.split("if extra_plots and total_slides >= 5:", 1)[1].split(
        "# 确定输出路径", 1
    )[0]
    assert "record_missing" in extra_seg, "追加页缺图仍静默跳过 (B-02)"
    warp_seg = src.split("if slide_no in [14, 15, 16]:", 1)[1].split("continue", 1)[0]
    assert "record_missing" in warp_seg, "14-16 保留页仍静默跳过 (B-03)"


def test_cancel_writes_done_marker(tmp):
    """B-04 回归 (2026-09-09 体检): 取消生成也必须写 gui_run_done.txt。
    直接运行 AutoReport.vbs (show_gui_before_run=true) → GUI 内点生成 → 点
    "取消生成" → 关窗: 若取消不写标记, 外层 VBS 实例看不到标记, 会把
    Moldflow 导出 + PPT 生成整条流水线重跑一遍, 违背取消语义。"""
    src = open(os.path.join(ROOT, "config_gui.py"), encoding="utf-8").read()
    body = src.split("def finish_run", 1)[1].split("\n", 1)[1].split("\n    def ", 1)[0]
    assert "gui_run_done.txt" in body, "finish_run 未写完成标记"
    assert body.index("gui_run_done.txt") < body.index("\n        if cancelled:"), (
        "完成标记写入必须先于 cancelled 提前返回 (取消路径也要写标记, B-04)"
    )


def test_vbs_plot_notfound_vs_disabled_log(tmp):
    """B-05 回归 (2026-09-09 体检): "方案中未找到结果" 与 "未启用" 必须是
    两条不同日志。此前 Else 分支错位: 勾选了但方案中找不到该结果名时,
    日志打"未启用 (跳过)"误导排查; 真正未启用的条目反而零日志。"""
    vbs = open(os.path.join(ROOT, "AutoReport.vbs"), "rb").read().decode("gbk")
    assert "方案中未找到结果 (跳过)" in vbs, "未找到分支日志文案未修正 (B-05)"
    assert vbs.count("未启用 (跳过)") == 1, (
        "未启用日志应恰好一条 (挂在 pEnabled 的 Else 分支)"
    )


def test_vbs_root_copy_saveimage_guarded(tmp):
    """C-06 回归 (2026-09-09 体检): temp 根目录副本的 Viewer.SaveImage 此前
    未包 On Error 保护 (同函数内 mode_a/mode_b 导出均有保护), 一旦 COM 失败
    整个脚本中断且无日志; 现要求该块有错误保护与 WARN 日志。"""
    vbs = open(os.path.join(ROOT, "AutoReport.vbs"), "rb").read().decode("gbk")
    assert "根目录副本截图异常" in vbs, "根目录副本 SaveImage 缺错误保护 (C-06)"
    seg = vbs.split("根目录副本: mode_a", 1)[1].split("On Error GoTo 0", 1)[0]
    assert "On Error Resume Next" in seg and "Viewer.SaveImage TempDir" in seg


def test_gbk_tee_write_failure_no_recursion(tmp):
    """B-06 回归 (2026-09-09 体检): _GbkTee 挂到 sys.stdout 后, 日志文件持续
    写失败 (磁盘满/句柄被锁) 时只降级到控制台, 绝不在 except 里调 print
    (print 会再次进入本 write, 无界递归直至 RecursionError)。"""
    import io

    class _BadFile:
        def write(self, _b):
            raise OSError("disk full (模拟)")

        def flush(self):
            raise OSError("disk full (模拟)")

    log_path = os.path.join(tmp, "run.log")
    old_out, old_err = sys.stdout, sys.stderr
    opened = []
    try:
        pb._attach_file_logging(log_path)
        tee = sys.stdout  # _GbkTee 实例 (_attach_file_logging 内嵌套类)
        opened = [tee.f, sys.stderr.f]
        tee.stream = io.StringIO()  # 捕获"控制台"侧输出
        tee.f = _BadFile()  # 注入日志侧持续写失败
        print("hello")
        tee.flush()
    finally:
        sys.stdout, sys.stderr = old_out, old_err
        for f in opened:
            try:
                f.close()
            except Exception:
                pass
    out = tee.stream.getvalue()
    assert "hello" in out, "控制台输出丢失"
    assert "log write failed" in out, "写失败未降级提示"


def test_fail_policy_helper_used_in_build(tmp):
    """C-03 回归 (2026-09-09 体检): on_missing_data=fail 检查必须统一走
    enforce_missing_policy, 不允许 build_single_report 内联重复实现
    (双实现漂移风险 — helper 目前仅测试引用)。"""
    src = open(os.path.join(ROOT, "core", "pptx_builder.py"), encoding="utf-8").read()
    body = src.split("def build_single_report", 1)[1].split("\ndef ", 1)[0]
    assert "enforce_missing_policy(" in body, "fail 策略仍为内联实现 (C-03)"


def test_gui_traces_registered_once(tmp):
    """C-05/C-04④ 回归 (2026-09-09 体检): 结果项 StringVar/BooleanVar 的
    trace 只注册一次 — 变量跨 tab 重建持久, populate_tab 每次重建重复
    trace_add 会累积回调 (默认项在 core 页与分类页各渲染一行更易翻倍);
    同时清除该处 try/except Exception: pass 裸捕获。"""
    src = open(os.path.join(ROOT, "config_gui.py"), encoding="utf-8").read()
    pt = src.split("def populate_tab", 1)[1].split("\n    def ", 1)[0]
    assert "trace_add" not in pt, "populate_tab 仍每次重建重复挂 trace (C-05)"
    assert "except Exception" not in pt, "populate_tab 仍存在裸 except (R-3.1)"
    bpt = src.split("def _build_plot_tabs", 1)[1].split("\n    def ", 1)[0]
    assert "trace_add" in bpt, "trace 未统一收口到 _build_plot_tabs"


def test_xy_detection_strict_underscore(tmp):
    """C-04 回归 (2026-09-09 体检): XY 曲线判据收紧为 "_xy"。
    裸 "xy" in base_name 会误命中含 xy 子串的 key (如 oxygen → 动态结果
    "氧分布 (oxygen)"), 使普通 3D 云图被当成 2D 曲线图走错误管线;
    且原条件 "_xy" in b or "xy" in b 前件恒被后件包含, 属冗余。"""
    src = open(
        os.path.join(ROOT, "core", "image_processor.py"), encoding="utf-8"
    ).read()
    assert 'or "xy" in base_name' not in src, "is_xy 仍含宽匹配 (C-04)"
    assert '"_xy" in base_name' in src, "is_xy 判据缺失"
    fn = src.split("def locate_curve_peak", 1)[1].split("\ndef ", 1)[0]
    assert "arr = np.array(im)" not in fn, "locate_curve_peak 仍存在未用变量 arr"


def main():
    tests = [
        (name, fn)
        for name, fn in sorted(globals().items())
        if name.startswith("test_") and callable(fn)
    ]
    failed = 0
    for name, fn in tests:
        try:
            with tempfile.TemporaryDirectory() as tmp:
                fn(tmp)
            print(f"[PASS] {name}")
        except Exception as e:
            failed += 1
            print(f"[FAIL] {name}: {e}")
    print(f"\n{'=' * 50}\n{len(tests) - failed}/{len(tests)} passed")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
