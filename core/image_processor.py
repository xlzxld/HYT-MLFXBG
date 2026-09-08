"""
image_processor.py - 方案 B 专属图像智能拼合与全要素增强模块
功能：
1. 3D 结果云图：
   - 提取完整左侧数据条（包含顶部结果标题与时间戳、单位、纵向色条与清晰数值，严格对齐用户图 2）；
   - 提取三维定向坐标系（XYZ 轴及三行旋转角度数值，做透明背景处理，严格对齐用户图 3）；
   - 提取/渲染右上角工程方案名称标签；
   - 紧致裁剪 3D 模型实体四周留白，超大倍率放大，最大化利用画面有效区域；
   - 按照 1.58:1 黄金比例（契合 PPT 版面安全框）合成全要素高保真云图（严格对齐用户图 1）。
2. 2D 曲线图（XY 图）：
   - Moldflow 原生 XY 截图紧致裁切输出，锁模力/注射位置处压力曲线在最高峰值处
     绘制黄色探针标记框 (annotate_xy_curves, 方案 A/B 通用; 峰值数据由调用方
     从 peak_values.json/曲线 txt 单源传入, 缺失则跳过标注, 绝不使用硬编码值)。
3. 纯模型本体图 (封面)：
   - solid_model.png 优先采用 VBS 图层自适应后的 SaveImage3 真实导出;
     导出空白/缺失时回退 generate_solid_cad_model 灰色重绘 (FUSION 旧导出曾全白)。
"""

import os
import re
import glob
import json
import numpy as np
from PIL import Image, ImageChops, ImageFilter, ImageDraw, ImageFont


def load_truetype_font(size, bold=False):
    """
    安全加载中文字体，支持 Windows 常用字体平滑降级：
    simsun.ttc -> msyh.ttc -> simhei.ttf -> arial.ttf -> ImageFont.load_default()
    """
    candidates = []
    if bold:
        candidates.extend(
            [
                r"C:\Windows\Fonts\msyhbd.ttc",
                r"C:\Windows\Fonts\simhei.ttf",
                r"C:\Windows\Fonts\msyh.ttc",
                r"C:\Windows\Fonts\simsun.ttc",
            ]
        )
    else:
        candidates.extend(
            [
                r"C:\Windows\Fonts\simsun.ttc",
                r"C:\Windows\Fonts\msyh.ttc",
                r"C:\Windows\Fonts\simhei.ttf",
                r"C:\Windows\Fonts\arial.ttf",
            ]
        )
    for cp in candidates:
        if os.path.exists(cp):
            try:
                return ImageFont.truetype(cp, size)
            except Exception as e:
                print(f"[image_processor] 字体 {cp} 加载失败: {e}")
    return ImageFont.load_default()


def trim_white_borders(img, border=8):
    """
    自动裁剪图片四周的纯白/近白透明边距，保留有效实体或图表区域
    """
    if img.mode != "RGBA":
        img_rgba = img.convert("RGBA")
    else:
        img_rgba = img

    bg = Image.new("RGBA", img_rgba.size, (255, 255, 255, 255))
    diff = ImageChops.difference(img_rgba, bg)
    diff_gray = diff.convert("L")
    threshold = 12
    diff_thresh = diff_gray.point(lambda p: 255 if p > threshold else 0)
    bbox = diff_thresh.getbbox()

    if bbox:
        x1 = max(0, bbox[0] - border)
        y1 = max(0, bbox[1] - border)
        x2 = min(img.size[0], bbox[2] + border)
        y2 = min(img.size[1], bbox[3] + border)
        return img_rgba.crop((x1, y1, x2, y2))
    return img_rgba


def make_transparent_white(img, threshold=245):
    """
    将图片的纯白/近白背景转为透明通道 (用于坐标系等元素悬浮叠加)
    """
    rgba = img.convert("RGBA")
    arr = np.array(rgba)
    white_mask = (
        (arr[:, :, 0] > threshold)
        & (arr[:, :, 1] > threshold)
        & (arr[:, :, 2] > threshold)
    )
    arr[:, :, 3] = np.where(white_mask, 0, 255)
    return Image.fromarray(arr)


def extract_viewport_components(viewport_path):
    """
    从 Moldflow 视口抓取图 (如 mode_a/pressure.png) 中提取全要素：
    1. left_bar: 左侧完整数据条（含结果名、时间、单位、色条及数值，图2）
    2. triad: 右下角三维坐标系与旋转角度（图3）
    3. title: 顶部工程方案名
    """
    if not os.path.exists(viewport_path):
        return None, None, None

    try:
        im = Image.open(viewport_path).convert("RGB")
    except Exception as e:
        print(f"[image_processor] Error opening viewport {viewport_path}: {e}")
        return None, None, None

    w, h = im.size
    if w < 400 or h < 300:
        return None, None, None

    # 1. 提取左侧数据条 (图2) — 含标题块(结果名+时间)、单位、色条与数值
    # 实机包围盒 (2026-09-08, 2112x1136 视口): 标题 y29-59, 单位 y191,
    # 色条 y228-870, 数值 y1093-1120, 左起 x15 — 裁剪框必须完整包住,
    # 旧框 (50,20,350,…) 曾把标题左缘与顶部切掉 (用户反馈"截图不全")。
    # T8pre 结论仍成立: 保持固定框 (Moldflow 图例铬层为固定逻辑像素)。
    left_crop = im.crop((10, 8, min(w, 340), min(h, 1160)))
    gray_l = left_crop.convert("L")
    bbox_l = gray_l.point(lambda p: 255 if p < 240 else 0).getbbox()
    left_bar = None
    if bbox_l:
        x1 = max(0, bbox_l[0] - 2)
        y1 = max(0, bbox_l[1] - 2)
        x2 = min(left_crop.width, bbox_l[2] + 4)
        y2 = min(left_crop.height, bbox_l[3] + 4)
        fill = (
            (bbox_l[2] - bbox_l[0])
            * (bbox_l[3] - bbox_l[1])
            / (left_crop.width * left_crop.height)
        )
        if fill > 0.8:
            print(
                f"[image_processor] WARN: 左侧色带提取填充率 {fill:.0%} 异常偏高, "
                f"可能卷入模型/水印 (视口 {w}x{h} 偏离设计尺寸 1920x1080?)"
            )
        left_bar = left_crop.crop((x1, y1, x2, y2))

    # 2. 提取三维坐标系 (图3)
    # 视口右下角区域 x: w-260..w, y: h-260..h
    br_crop = im.crop((max(0, w - 260), max(0, h - 260), w, h))
    gray_br = br_crop.convert("L")
    bbox_br = gray_br.point(lambda p: 255 if p < 245 else 0).getbbox()
    triad = None
    if bbox_br:
        t_box = (
            max(0, bbox_br[0] - 2),
            max(0, bbox_br[1] - 2),
            min(br_crop.width, bbox_br[2] + 2),
            min(br_crop.height, bbox_br[3] + 2),
        )
        triad_crop = br_crop.crop(t_box)
        triad = make_transparent_white(triad_crop, threshold=245)

    # 3. 提取顶部工程方案名 (图1右上角)
    # 视口顶部区域 x: 800..1800, y: 20..80
    top_crop = im.crop((min(w - 100, 800), 20, min(w, 1800), min(h, 80)))
    gray_t = top_crop.convert("L")
    bbox_t = gray_t.point(lambda p: 255 if p < 240 else 0).getbbox()
    title_img = None
    if bbox_t:
        t_box = (
            max(0, bbox_t[0] - 2),
            max(0, bbox_t[1] - 2),
            min(top_crop.width, bbox_t[2] + 2),
            min(top_crop.height, bbox_t[3] + 2),
        )
        title_img = top_crop.crop(t_box)

    return left_bar, triad, title_img


def generate_solid_cad_model(model_source_path, output_path):
    """
    从纯模型着色图生成 100% 纯净、无网格、无节点的 CAD 实体本体图 (供 Slide 1 封面使用)
    采用 Autodesk 经典酷炫 CAD 金属灰/冷蓝灰材质着色与高光阴影
    [2026-09-08 起为封面回退路径: VBS 图层自适应导出 (Fusion 保留 T) 为首选,
     仅当其空白/缺失时调用本函数兜底]
    """
    if not os.path.exists(model_source_path):
        return False

    try:
        im_model = Image.open(model_source_path).convert("RGBA")
        arr_rgba = np.array(im_model)
        # 背景为纯白
        fg_mask = ~(
            (arr_rgba[:, :, 0] > 250)
            & (arr_rgba[:, :, 1] > 250)
            & (arr_rgba[:, :, 2] > 250)
        )
        if not np.any(fg_mask):
            return False

        im_gray = im_model.convert("L")
        arr_gray = np.array(im_gray).astype(np.float32) / 255.0

        out_arr = np.full_like(arr_rgba, 255)
        # 经典工程 CAD 实体配色 (优雅冷灰 RGB 168, 182, 198)
        base_r, base_g, base_b = 168, 182, 198
        for c, base in enumerate([base_r, base_g, base_b]):
            val = np.clip(base * (0.38 + 1.25 * arr_gray), 0, 255)
            out_arr[:, :, c] = np.where(fg_mask, val.astype(np.uint8), 255)

        out_im = Image.fromarray(out_arr).convert("RGB")
        # 适度裁剪留白
        out_trimmed = trim_white_borders(out_im, border=20).convert("RGB")
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        out_trimmed.save(output_path, "PNG", quality=98)
        print(
            f"[image_processor] Created pure CAD solid model (no mesh, no nodes) -> {output_path}"
        )
        return True
    except Exception as e:
        print(f"[image_processor] Error generating solid CAD model: {e}")
        return False


def annotate_curve_peak(
    img,
    peak_x_sec,
    peak_y_val,
    unit_str,
    roi_x=(0.15, 0.40),
    roi_y=(0.08, 0.32),
    default_pos=(0.25, 0.12),
    val_format=".1f",
):
    """
    通用 Moldflow 官方 XY 曲线峰值探针标注函数：
    在最高峰拐点处标记 x = ... [s], y = ... [单位] (黄色底色提示框 + 黑色引线 + 红色峰值点).
    """
    im = img.convert("RGB")
    w, h = im.size
    arr = np.array(im)

    # 检查是否已包含黄色提示框，避免重复叠加
    yellow_mask = (
        (arr[:, :, 0] > 240)
        & (arr[:, :, 1] > 240)
        & (arr[:, :, 2] > 180)
        & (arr[:, :, 2] < 225)
    )
    if np.sum(yellow_mask) > 100:
        return im

    # 智能定位峰值拐点
    black = (arr[:, :, 0] < 40) & (arr[:, :, 1] < 40) & (arr[:, :, 2] < 40)
    x_start, x_end = int(roi_x[0] * w), int(roi_x[1] * w)
    y_start, y_end = int(roi_y[0] * h), int(roi_y[1] * h)
    sub = black[y_start:y_end, x_start:x_end]
    ys, xs = np.where(sub)

    peak_x, peak_y = int(default_pos[0] * w), int(default_pos[1] * h)
    if len(ys) > 0:
        min_y_idx = np.argmin(ys)
        peak_y = y_start + int(ys[min_y_idx])
        peak_x = x_start + int(xs[min_y_idx])

    font_size = max(14, int(h * 0.016))
    f_callout = load_truetype_font(font_size)

    line1 = f"x = {peak_x_sec:.3f} [s]"
    formatted_y = f"{peak_y_val:{val_format}}"
    line2 = f"y = {formatted_y} [{unit_str}]"

    draw = ImageDraw.Draw(im)

    pad_x = int(font_size * 0.5)
    pad_y = int(font_size * 0.3)
    t1_w = int(draw.textlength(line1, font=f_callout))
    t2_w = int(draw.textlength(line2, font=f_callout))
    box_w = max(t1_w, t2_w) + pad_x * 2
    box_h = int(font_size * 2.6)

    # 框体置于峰值点右上方
    box_x1 = peak_x + int(font_size * 1.6)
    box_y1 = peak_y - int(font_size * 2.1)
    box_x2 = box_x1 + box_w
    box_y2 = box_y1 + box_h

    # 1. 绘制黑色引线
    draw.line(
        [(peak_x, peak_y), (box_x1, box_y1 + int(box_h * 0.45))],
        fill=(0, 0, 0),
        width=max(1, int(font_size * 0.08)),
    )

    # 2. 绘制峰值红色标记点
    r_dot = max(2, int(font_size * 0.15))
    draw.ellipse(
        [(peak_x - r_dot, peak_y - r_dot), (peak_x + r_dot, peak_y + r_dot)],
        fill=(255, 0, 0),
        outline=(0, 0, 0),
    )

    # 3. 绘制经典黄色探针框 (底色 #FFFFC8, 细黑边)
    draw.rectangle(
        [box_x1, box_y1, box_x2, box_y2], fill="#ffffc8", outline="#000000", width=1
    )

    # 4. 绘制文字
    draw.text((box_x1 + pad_x, box_y1 + pad_y), line1, fill=(0, 0, 0), font=f_callout)
    draw.text(
        (box_x1 + pad_x, box_y1 + pad_y + int(font_size * 1.2)),
        line2,
        fill=(0, 0, 0),
        font=f_callout,
    )

    return im


def annotate_clamp_force_peak(img, peak_x_sec, peak_y_ton):
    """
    还原 Moldflow 锁模力 XY 曲线查询探针标注 (公吨)。
    峰值数据由调用方从 peak_values.json/config 单源传入, 本函数不含任何默认值。
    """
    return annotate_curve_peak(
        img,
        peak_x_sec,
        peak_y_ton,
        "公吨",
        roi_x=(0.18, 0.40),
        roi_y=(0.08, 0.32),
        default_pos=(0.265, 0.113),
        val_format=".1f",
    )


def annotate_inj_pressure_peak(img, peak_x_sec, peak_y_mpa):
    """
    还原 Moldflow 注射位置处压力 XY 曲线查询探针标注 (MPa)。
    峰值数据由调用方从 peak_values.json/config 单源传入, 本函数不含任何默认值。
    """
    return annotate_curve_peak(
        img,
        peak_x_sec,
        peak_y_mpa,
        "MPa",
        roi_x=(0.15, 0.35),
        roi_y=(0.08, 0.30),
        default_pos=(0.215, 0.144),
        val_format=".2f",
    )


CURVE_DATA_FILES = {
    "clamp_force": "clamp_force_xy_curve_data.txt",
    "inj_pressure": "inj_pressure_xy_curve_data.txt",
}


def _parse_curve_peak(path):
    """
    解析 VBS 通过官方 Plot.SaveXYPlotCurveData 导出的曲线 txt,
    返回 (peak_value, peak_time) 或 None。
    列识别: 自变量列 = 首个近似单调不减的数值列; 因变量 = 最后一个数值列。
    解析不出两列有效数据 → None (调用方按缺失处理, 绝不臆造)。
    """
    try:
        with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
            lines = f.read().splitlines()
    except OSError as e:
        print(f"[image_processor] WARN: 读取曲线数据失败 ({path}): {e}")
        return None

    rows = []
    for ln in lines:
        parts = re.split(r"[\s,;\t]+", ln.strip())
        vals = []
        for p in parts:
            if not p:
                continue
            try:
                vals.append(float(p.replace(",", ".")))
            except ValueError:
                continue
        if len(vals) >= 2:
            rows.append(vals)
    if len(rows) < 2:
        return None

    n_cols = min(len(r) for r in rows)
    if n_cols < 2:
        return None

    def monotonic(col):
        inc = all(
            rows[i + 1][col] >= rows[i][col] - 1e-12 for i in range(len(rows) - 1)
        )
        dec = all(
            rows[i + 1][col] <= rows[i][col] + 1e-12 for i in range(len(rows) - 1)
        )
        return inc or dec

    x_col = None
    for c in range(n_cols):
        if monotonic(c):
            x_col = c
            break
    if x_col is None:
        x_col = 0
    y_col = n_cols - 1
    if y_col == x_col:
        return None

    peak_val, peak_time = None, None
    for r in rows:
        y = r[y_col]
        if peak_val is None or y > peak_val:
            peak_val = y
            peak_time = r[x_col]
    if peak_val is None or peak_time is None:
        return None
    return round(peak_val, 3), round(peak_time, 3)


def load_peaks(data_dir):
    """
    读取本次导出的 XY 峰值数据:
    1. peak_values.json (VBS GetMaxValue 值, 无时刻);
    2. {curve}_xy_curve_data.txt (官方 SaveXYPlotCurveData 曲线导出, 可同时得值+时刻)。
    txt 解析成功 → 覆盖 json (值+时刻齐全才能用于探针标注); 与 json 值差异大时告警。
    缺失/损坏 → {}。
    """
    base = data_dir or "."
    path = os.path.join(base, "peak_values.json")
    d = {}
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8-sig") as f:
                d = json.load(f)
            if not isinstance(d, dict):
                d = {}
        except Exception as e:
            print(f"[image_processor] WARN: 读取 peak_values.json 失败: {e}")
            d = {}

    for curve, fname in CURVE_DATA_FILES.items():
        cpath = os.path.join(base, fname)
        if not os.path.exists(cpath):
            continue
        parsed = _parse_curve_peak(cpath)
        if parsed is None:
            print(
                f"[image_processor] WARN: 曲线数据解析失败, 该曲线按缺失处理 ({fname})"
            )
            continue
        peak_val, peak_time = parsed
        entry = d.get(curve) if isinstance(d.get(curve), dict) else {}
        json_val = entry.get("peak_value")
        if (
            isinstance(json_val, (int, float))
            and json_val > 0
            and abs(peak_val - float(json_val)) / max(abs(float(json_val)), 1e-9) > 0.2
        ):
            print(
                f"[image_processor] WARN: {curve} 曲线解析峰值 {peak_val} 与 "
                f"GetMaxValue {json_val} 差异>20%, 请人工复核 ({fname})"
            )
        d[curve] = {"peak_value": peak_val, "peak_time": peak_time}
    return d


def merge_scale_and_model(
    scale_path,
    model_path,
    output_path,
    viewport_path=None,
    ref_triad_path=None,
    ref_title_path=None,
    clamp_peak=None,
    inj_peak=None,
    target_height=None,
    expected_model_size=None,
):
    """
    按照用户要求 (图 1、图 2、图 3) 执行高保真智能全要素拼合：
    expected_model_size: (w, h) SaveImage3 应输出的分辨率 (image_settings 配置值);
    实际输出不符 = 离屏导出布局错乱 (实发: 配 2560x1440 实出 3840x2160 且模型残缺),
    弃用并回退视口图。0x0 视口原生配置传 None 跳过该校验。
    1. 2D XY 曲线图：紧致裁切多余留白，如果是锁模力曲线图或注射位置处压力曲线图则在最高峰值处绘制黄色探针标记框
       (峰值数据 clamp_peak/inj_peak 由调用方单源传入, 缺失则跳过标注, 绝不使用硬编码值)；
    2. 3D 云图：
       - 左侧：完整数据条（结果名+时间戳+单位+色条及数值，图2）
       - 中右侧：超大裁剪模型实体
       - 右上角：工程方案名
       - 右下角：透明叠加三维坐标系与旋转角度（图3）
    target_height: 期望画布高度 (来自 image_settings.height, 1.58:1)。
    缺省 1200; 源模型分辨率不足时自动降至 model 高度+70, 绝不放大模糊 (T8)。
    """
    base_name = os.path.basename(output_path).lower()
    is_xy = "_xy" in base_name or "xy" in base_name

    # 1. 2D 曲线图处理
    if is_xy:
        src_path = model_path if os.path.exists(model_path) else scale_path
        if (
            not os.path.exists(src_path)
            and viewport_path
            and os.path.exists(viewport_path)
        ):
            src_path = viewport_path
        if not os.path.exists(src_path):
            return False
        try:
            img = Image.open(src_path).convert("RGB")
            trimmed = trim_white_borders(img, border=12).convert("RGB")
            if "clamp_force" in base_name:
                if clamp_peak:
                    trimmed = annotate_clamp_force_peak(
                        trimmed,
                        peak_x_sec=clamp_peak[1],
                        peak_y_ton=clamp_peak[0],
                    )
                else:
                    print(
                        f"[image_processor] WARN: clamp_force 峰值数据缺失, 跳过探针标注 ({os.path.basename(output_path)})"
                    )
            if "inj_pressure" in base_name or "pressure_xy" in base_name:
                if inj_peak:
                    trimmed = annotate_inj_pressure_peak(
                        trimmed,
                        peak_x_sec=inj_peak[1],
                        peak_y_mpa=inj_peak[0],
                    )
                else:
                    print(
                        f"[image_processor] WARN: inj_pressure 峰值数据缺失, 跳过探针标注 ({os.path.basename(output_path)})"
                    )
            os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
            trimmed.save(output_path, "PNG", quality=98)
            print(
                f"[image_processor] Exported clean XY plot -> {os.path.basename(output_path)}"
            )
            return True
        except Exception as e:
            print(f"[image_processor] Error processing XY plot {src_path}: {e}")
            return False

    # 2. 3D 云图全要素拼合
    # 2.1 数据条/坐标系/工程标题 — 唯一来源 = mode_a 视口截图 (含完整图例:
    #     标题块+时间+单位+色条+数值, 即用户要求的完整数据条)。
    # 实机取证 (2026-09-08): Regenerate 修复后视口逐图新鲜 (两两差异 22-24%);
    #     但 Viewer.SavePlotScaleImage 实测每次导出同一条 (两两差异仅 0.4-1.1%,
    #     且不含标题块) — scale_*.png 只能作最后兜底, 禁作主源。
    left_bar = None
    triad = None
    title_img = None

    if viewport_path and os.path.exists(viewport_path):
        left_bar, triad, title_img = extract_viewport_components(viewport_path)

    # 回退：若没有从视口抓取到坐标系，尝试使用全局参考坐标系
    if triad is None and ref_triad_path and os.path.exists(ref_triad_path):
        try:
            triad = Image.open(ref_triad_path).convert("RGBA")
        except Exception as e:
            print(f"[image_processor] 参考坐标系读取失败 ({ref_triad_path}): {e}")

    # 回退：若没有从视口抓取到工程标题，尝试使用全局参考标题
    if title_img is None and ref_title_path and os.path.exists(ref_title_path):
        try:
            title_img = Image.open(ref_title_path).convert("RGB")
        except Exception as e:
            print(f"[image_processor] 参考标题读取失败 ({ref_title_path}): {e}")

    # 最后回退：视口缺失时用 scale_*.png (已知缺陷: 可能不是本图专属条, 无标题块)
    if left_bar is None and os.path.exists(scale_path):
        try:
            with Image.open(scale_path) as im_s:
                if (np.array(im_s.convert("L")) < 240).mean() > 0.02:
                    left_bar = trim_white_borders(im_s.convert("RGB"), border=4)
        except Exception as e:
            print(f"[image_processor] 色带回退读取失败 ({scale_path}): {e}")

    # 检查模型文件
    if not os.path.exists(model_path):
        if viewport_path and os.path.exists(viewport_path):
            # 若没有独立模型图，则直接将视口图作为输出
            try:
                Image.open(viewport_path).save(output_path)
                return True
            except Exception as e:
                print(
                    f"[image_processor] 视口图直出失败 ({viewport_path} -> {output_path}): {e}"
                )
                return False
        return False

    try:
        im_model_raw = Image.open(model_path).convert("RGBA")
    except Exception as e:
        print(f"[image_processor] Error reading model {model_path}: {e}")
        return False

    # 紧致裁切模型多余白边
    im_model_trimmed = trim_white_borders(im_model_raw, border=8)
    if im_model_trimmed.height == 0 or im_model_trimmed.width == 0:
        return False

    # SaveImage3 离屏导出可靠性差 (实机取证 2026-09-08: 配 2560x1440 却输出
    # 3840x2160 且模型只渲染出局部 — "模型割裂/不完整"的直接来源)。
    # 校验 1 (主): 实际输出分辨率 != 配置分辨率 → 导出布局已错乱, 弃用;
    # 校验 2 (双保险): 内容占画布面积比过低 = 渲染残缺, 同样弃用。
    # 命中即回退视口图 (完整正确, 与方案 A 同级保底, 绝不输出残缺图)。
    size_mismatch = bool(
        expected_model_size
        and tuple(im_model_raw.size) != tuple(expected_model_size)
        and all(v > 0 for v in expected_model_size)
    )
    area_ratio = (im_model_trimmed.width * im_model_trimmed.height) / float(
        im_model_raw.width * im_model_raw.height
    )
    h_ratio = im_model_trimmed.height / float(im_model_raw.height)
    aspect = im_model_trimmed.width / float(im_model_trimmed.height)
    # 渲染断带特征 (实发 2026-09-08 4K 配置: 模型只剩一条横带, 高度占比 30%,
    # 宽高比 2.66): Fit 后的正常模型高度占比远高于此, 命中即视为残缺。
    # 权衡: 天然超宽扁零件 Fit 后也可能低占比, 误杀时回退的视口图仍完整可用。
    band_broken = h_ratio < 0.35 and aspect > 2.2
    if size_mismatch or area_ratio < 0.10 or band_broken:
        if size_mismatch:
            reason = (
                f"输出 {im_model_raw.width}x{im_model_raw.height} != 配置 "
                f"{expected_model_size[0]}x{expected_model_size[1]}"
            )
        elif band_broken:
            reason = f"渲染断带特征 (内容高度占比 {h_ratio:.0%}, 宽高比 {aspect:.2f})"
        else:
            reason = f"内容占画布 {area_ratio:.0%} < 10%"

        print(
            f"[image_processor] WARN: 模型图导出异常 ({reason}), "
            f"弃用 {os.path.basename(model_path)} 回退视口图 — 请检查 Moldflow 窗口状态"
        )
        if viewport_path and os.path.exists(viewport_path):
            try:
                Image.open(viewport_path).save(output_path)
                return True
            except Exception as e:
                print(f"[image_processor] 视口图直出失败 ({viewport_path}): {e}")
        return False

    # 若没有任何数据条，则单独输出模型
    if left_bar is None:
        im_model_trimmed.convert("RGB").save(output_path, "PNG", quality=98)
        return True

    # ---------------- 3. 画布构建与满幅黄金比例排版 ----------------
    # 目标：构建 1.58:1 宽高比画布，与 PPT 幻灯片安全区域完美吻合。
    # T8: 画布高度取调用方传入的期望值 (image_settings.height);
    # 源模型不足该高度时降到 model_h+70 (模型占 94% 高), 绝不放大模糊。
    target_h = int(target_height) if target_height else 1200
    max_useful_h = im_model_trimmed.height + 70
    if target_h > max_useful_h:
        print(
            f"[image_processor] 画布高度 {target_h} > 源模型可用高度 {max_useful_h}, "
            f"降为 {max_useful_h} 避免放大模糊 (提高导出分辨率可解除)"
        )
        target_h = max_useful_h
    target_w = int(target_h * 1.58)

    # 缩放模型：高度占用 ~94% 画布高度，实现震撼超大图幅
    model_avail_h = target_h - 70
    m_scale = model_avail_h / float(im_model_trimmed.height)
    m_w = int(im_model_trimmed.width * m_scale)
    m_h = int(im_model_trimmed.height * m_scale)
    model_scaled = im_model_trimmed.resize((m_w, m_h), Image.Resampling.LANCZOS)

    # 缩放左侧数据条：高度协调匹配
    lb_avail_h = target_h - 75
    lb_scale = lb_avail_h / float(left_bar.height)
    lb_w = int(left_bar.width * lb_scale)
    lb_h = int(left_bar.height * lb_scale)
    left_bar_scaled = left_bar.resize((lb_w, lb_h), Image.Resampling.LANCZOS)
    # T10: 仅当色带被放大时轻度锐化 (源足够时锐化只引入伪影)
    if lb_scale > 1.05:
        left_bar_scaled = left_bar_scaled.filter(
            ImageFilter.UnsharpMask(radius=1.2, percent=110, threshold=2)
        )

    # 创建纯白高清画布
    canvas = Image.new("RGB", (target_w, target_h), (255, 255, 255))

    # 1. 放置左侧数据条 (图2)
    lb_x = 35
    lb_y = 35
    canvas.paste(left_bar_scaled, (lb_x, lb_y))

    # 2. 放置超大模型 (居中偏右)
    rem_x = lb_x + lb_w + 20
    rem_w = target_w - rem_x - 20
    m_x = rem_x + max(0, (rem_w - m_w) // 2)
    m_y = (target_h - m_h) // 2
    canvas.paste(model_scaled, (m_x, m_y), model_scaled)

    # 3. 放置右上角工程方案名称 (图1右上角)
    if title_img is not None:
        t_x = target_w - title_img.width - 45
        t_y = 35
        canvas.paste(title_img, (t_x, t_y))

    # 4. 放置右下角三维坐标系与旋转角度 (图3)
    if triad is not None:
        tr_x = target_w - triad.width - 35
        tr_y = target_h - triad.height - 35
        canvas.paste(triad, (tr_x, tr_y), triad)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    canvas.save(output_path, "PNG", quality=98)
    print(
        f"[image_processor] Merged Mode B (图1/2/3规范) -> {os.path.basename(output_path)} ({target_w}x{target_h})"
    )
    return True


def process_all_mode_b_plots(
    mode_b_dir, data_dir=None, peaks=None, target_height=None, expected_model_size=None
):
    """
    批量处理方案 B 的全部图表：
    1. 生成/更新全局三维坐标系 (triad.png) 与工程标题 (study_title.png) 参考模板；
    2. 针对每个结果，从 mode_a 提取图 2 标准数据条并与 mode_b 纯净模型紧密拼合；
    3. 封面 solid_model.png: VBS 真实导出优先, 空白/缺失时灰色重绘兜底；
    4. 2D XY 图高清优化输出。
    peaks: {"clamp_force": (peak_value, peak_time) | None,
            "inj_pressure": (peak_value, peak_time) | None}
    由调用方从 peak_values.json/config 单源解析后传入; 缺失的曲线跳过探针标注。
    target_height: 拼合画布期望高度 (image_settings.height); None=默认 1200。
    """
    clamp_peak = (peaks or {}).get("clamp_force")
    inj_peak = (peaks or {}).get("inj_pressure")
    if not os.path.exists(mode_b_dir):
        return 0

    if data_dir is None:
        data_dir = os.path.dirname(os.path.abspath(mode_b_dir))

    mode_a_dir = os.path.join(data_dir, "mode_a")

    # 1. 先从 mode_a 中最具代表性的图片 (如 pressure.png) 提取全局基准坐标系与标题
    # (ref 图每次运行强制重生成, 杜绝继承上次方案的坐标/标题 — 陈旧基准图属于假数据类缺陷)
    ref_triad_path = os.path.join(mode_b_dir, "ref_triad.png")
    ref_title_path = os.path.join(mode_b_dir, "ref_title.png")

    for cand_name in [
        "pressure.png",
        "flow_front_temp.png",
        "volumetric_shrinkage.png",
    ]:
        p_cand = os.path.join(mode_a_dir, cand_name)
        if os.path.exists(p_cand):
            _, triad_cand, title_cand = extract_viewport_components(p_cand)
            if triad_cand:
                triad_cand.save(ref_triad_path)
            if title_cand:
                title_cand.save(ref_title_path)
            break

    # 2. 封面模型本体图 (solid_model.png): VBS 真实导出优先 (图层自适应, 见 AutoReport.vbs 4.1);
    #    导出缺失或全白时回退灰色重绘 (FUSION 隐藏 T 曾导出全白空图, 实机取证 2026-09-08)
    solid_out = os.path.join(data_dir, "solid_model.png")
    solid_ok = False
    if os.path.exists(solid_out):
        try:
            with Image.open(solid_out) as im_s:
                solid_ok = (np.array(im_s.convert("L")) < 240).mean() > 0.005
        except Exception as e:
            print(f"[image_processor] 读取 solid_model.png 失败: {e}")
            solid_ok = False
    if not solid_ok:
        model_cand = os.path.join(mode_b_dir, "model_pressure.png")
        if not os.path.exists(model_cand):
            model_cand = os.path.join(mode_b_dir, "model_volumetric_shrinkage.png")
        if os.path.exists(model_cand):
            if generate_solid_cad_model(model_cand, solid_out):
                print(
                    "[image_processor] VBS 实体导出空白/缺失, 已回退灰色重绘封面 (solid_model.png)"
                )
        else:
            print(
                "[image_processor][WARN] 封面无可用来源 (VBS 导出空白且无模型图), 报告封面将登记缺失"
            )

    # 3. 收集所有需拼合的 key
    all_keys = set()
    for s_path in glob.glob(os.path.join(mode_b_dir, "scale_*.png")):
        b = os.path.basename(s_path)
        all_keys.add(b[len("scale_") : -len(".png")])

    for m_path in glob.glob(os.path.join(mode_b_dir, "model_*.png")):
        b = os.path.basename(m_path)
        all_keys.add(b[len("model_") : -len(".png")])

    # 也检查 mode_a 中的 key
    if os.path.exists(mode_a_dir):
        for a_path in glob.glob(os.path.join(mode_a_dir, "*.png")):
            b = os.path.basename(a_path)
            all_keys.add(b[:-4])

    processed = 0
    for key in sorted(all_keys):
        s_path = os.path.join(mode_b_dir, f"scale_{key}.png")
        m_path = os.path.join(mode_b_dir, f"model_{key}.png")
        vp_path = os.path.join(mode_a_dir, f"{key}.png")
        if not os.path.exists(vp_path):
            cand_root = os.path.join(data_dir, f"{key}.png")
            if os.path.exists(cand_root):
                vp_path = cand_root

        out_path = os.path.join(mode_b_dir, f"{key}.png")
        if merge_scale_and_model(
            s_path,
            m_path,
            out_path,
            viewport_path=vp_path,
            ref_triad_path=ref_triad_path,
            ref_title_path=ref_title_path,
            clamp_peak=clamp_peak,
            inj_peak=inj_peak,
            target_height=target_height,
            expected_model_size=expected_model_size,
        ):
            processed += 1

    # 确保 mode_a 与 root 目录中的 2D XY 曲线图也进行峰值标注 (与方案 B 品质一致;
    # 峰值数据单源传入; 缺失则跳过标注并告警, 绝不使用硬编码值)
    annotate_xy_curves(mode_a_dir, data_dir, peaks)

    print(f"[image_processor] Total Mode B images successfully processed: {processed}")
    return processed


def annotate_xy_curves(mode_a_dir, data_dir, peaks=None):
    """
    给 Moldflow 原生 XY 截图 (mode_a/ 与 data_dir 根目录的 clamp_force_xy.png /
    inj_pressure_xy.png) 画峰值探针标注。方案 A/B 两种报告模式通用;
    已含探针的图 (黄色框检测) 自动跳过, 可安全重复调用。
    """
    for xy_target_dir in [mode_a_dir, data_dir]:
        if not (xy_target_dir and os.path.exists(xy_target_dir)):
            continue
        for xy_name, xy_key in [
            ("clamp_force_xy.png", "clamp_force"),
            ("inj_pressure_xy.png", "inj_pressure"),
        ]:
            xy_file = os.path.join(xy_target_dir, xy_name)
            if not os.path.exists(xy_file):
                continue
            xy_peak = (peaks or {}).get(xy_key)
            if not xy_peak:
                print(
                    f"[image_processor] WARN: {xy_key} 峰值数据缺失, 跳过探针标注 ({xy_file})"
                )
                continue
            try:
                im_xy = Image.open(xy_file).convert("RGB")
                if xy_key == "clamp_force":
                    im_xy_annot = annotate_clamp_force_peak(
                        im_xy,
                        peak_x_sec=xy_peak[1],
                        peak_y_ton=xy_peak[0],
                    )
                else:
                    im_xy_annot = annotate_inj_pressure_peak(
                        im_xy,
                        peak_x_sec=xy_peak[1],
                        peak_y_mpa=xy_peak[0],
                    )
                im_xy_annot.save(xy_file, "PNG", quality=98)
            except Exception as e:
                print(f"[image_processor] Error annotating {xy_file}: {e}")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Mode B Image Merging Processor")
    parser.add_argument(
        "--dir",
        default=r"c:\Users\5600\Documents\ZDH\MLFXBG\temp\mode_b",
        help="Mode B directory",
    )
    args = parser.parse_args()
    process_all_mode_b_plots(args.dir)
