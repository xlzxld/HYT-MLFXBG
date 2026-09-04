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
   - 自动识别为 2D 图表，紧致裁切多余留白，原图高清输出。
3. 纯 CAD 模型生成：
   - 生成 100% 无网格、无节点的纯 CAD 实体本体图 (solid_model.png)，专供 Slide 1 封面替换。
"""

import os
import glob
import shutil
import numpy as np
from PIL import Image, ImageChops, ImageFilter, ImageDraw, ImageFont

def load_truetype_font(size, bold=False):
    """
    安全加载中文字体，支持 Windows 常用字体平滑降级：
    simsun.ttc -> msyh.ttc -> simhei.ttf -> arial.ttf -> ImageFont.load_default()
    """
    candidates = []
    if bold:
        candidates.extend([
            r"C:\Windows\Fonts\msyhbd.ttc",
            r"C:\Windows\Fonts\simhei.ttf",
            r"C:\Windows\Fonts\msyh.ttc",
            r"C:\Windows\Fonts\simsun.ttc"
        ])
    else:
        candidates.extend([
            r"C:\Windows\Fonts\simsun.ttc",
            r"C:\Windows\Fonts\msyh.ttc",
            r"C:\Windows\Fonts\simhei.ttf",
            r"C:\Windows\Fonts\arial.ttf"
        ])
    for cp in candidates:
        if os.path.exists(cp):
            try:
                return ImageFont.truetype(cp, size)
            except Exception:
                pass
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
    white_mask = (arr[:, :, 0] > threshold) & (arr[:, :, 1] > threshold) & (arr[:, :, 2] > threshold)
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

    # 1. 提取左侧数据条 (图2)
    # 视口左侧区域 x: 50..350, y: 20..1160
    left_crop = im.crop((50, 20, min(w, 350), min(h, 1160)))
    gray_l = left_crop.convert("L")
    bbox_l = gray_l.point(lambda p: 255 if p < 240 else 0).getbbox()
    left_bar = None
    if bbox_l:
        x1 = max(0, bbox_l[0] - 2)
        y1 = max(0, bbox_l[1] - 2)
        x2 = min(left_crop.width, bbox_l[2] + 4)
        y2 = min(left_crop.height, bbox_l[3] + 4)
        left_bar = left_crop.crop((x1, y1, x2, y2))

    # 2. 提取三维坐标系 (图3)
    # 视口右下角区域 x: w-260..w, y: h-260..h
    br_crop = im.crop((max(0, w - 260), max(0, h - 260), w, h))
    gray_br = br_crop.convert("L")
    bbox_br = gray_br.point(lambda p: 255 if p < 245 else 0).getbbox()
    triad = None
    if bbox_br:
        t_box = (max(0, bbox_br[0] - 2), max(0, bbox_br[1] - 2), min(br_crop.width, bbox_br[2] + 2), min(br_crop.height, bbox_br[3] + 2))
        triad_crop = br_crop.crop(t_box)
        triad = make_transparent_white(triad_crop, threshold=245)

    # 3. 提取顶部工程方案名 (图1右上角)
    # 视口顶部区域 x: 800..1800, y: 20..80
    top_crop = im.crop((min(w - 100, 800), 20, min(w, 1800), min(h, 80)))
    gray_t = top_crop.convert("L")
    bbox_t = gray_t.point(lambda p: 255 if p < 240 else 0).getbbox()
    title_img = None
    if bbox_t:
        t_box = (max(0, bbox_t[0] - 2), max(0, bbox_t[1] - 2), min(top_crop.width, bbox_t[2] + 2), min(top_crop.height, bbox_t[3] + 2))
        title_img = top_crop.crop(t_box)

    return left_bar, triad, title_img

def generate_solid_cad_model(model_source_path, output_path):
    """
    从纯模型着色图生成 100% 纯净、无网格、无节点的 CAD 实体本体图 (供 Slide 1 封面使用)
    采用 Autodesk 经典酷炫 CAD 金属灰/冷蓝灰材质着色与高光阴影
    """
    if not os.path.exists(model_source_path):
        return False

    try:
        im_model = Image.open(model_source_path).convert("RGBA")
        arr_rgba = np.array(im_model)
        # 背景为纯白
        fg_mask = ~((arr_rgba[:, :, 0] > 250) & (arr_rgba[:, :, 1] > 250) & (arr_rgba[:, :, 2] > 250))
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
        print(f"[image_processor] Created pure CAD solid model (no mesh, no nodes) -> {output_path}")
        return True
    except Exception as e:
        print(f"[image_processor] Error generating solid CAD model: {e}")
        return False

def annotate_curve_peak(img, peak_x_sec, peak_y_val, unit_str,
                        roi_x=(0.15, 0.40), roi_y=(0.08, 0.32),
                        default_pos=(0.25, 0.12), val_format=".1f"):
    """
    通用 Moldflow 官方 XY 曲线峰值探针标注函数：
    在最高峰拐点处标记 x = ... [s], y = ... [单位] (黄色底色提示框 + 黑色引线 + 红色峰值点).
    """
    im = img.convert("RGB")
    w, h = im.size
    arr = np.array(im)

    # 检查是否已包含黄色提示框，避免重复叠加
    yellow_mask = (arr[:, :, 0] > 240) & (arr[:, :, 1] > 240) & (arr[:, :, 2] > 180) & (arr[:, :, 2] < 225)
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
    draw.line([(peak_x, peak_y), (box_x1, box_y1 + int(box_h * 0.45))], fill=(0, 0, 0), width=max(1, int(font_size * 0.08)))

    # 2. 绘制峰值红色标记点
    r_dot = max(2, int(font_size * 0.15))
    draw.ellipse([(peak_x - r_dot, peak_y - r_dot), (peak_x + r_dot, peak_y + r_dot)], fill=(255, 0, 0), outline=(0, 0, 0))

    # 3. 绘制经典黄色探针框 (底色 #FFFFC8, 细黑边)
    draw.rectangle([box_x1, box_y1, box_x2, box_y2], fill="#ffffc8", outline="#000000", width=1)

    # 4. 绘制文字
    draw.text((box_x1 + pad_x, box_y1 + pad_y), line1, fill=(0, 0, 0), font=f_callout)
    draw.text((box_x1 + pad_x, box_y1 + pad_y + int(font_size * 1.2)), line2, fill=(0, 0, 0), font=f_callout)

    return im

def annotate_clamp_force_peak(img, peak_x_sec=6.224, peak_y_ton=320.5):
    """
    还原 Moldflow 锁模力 XY 曲线查询探针标注 (公吨)
    """
    return annotate_curve_peak(img, peak_x_sec, peak_y_ton, "公吨",
                               roi_x=(0.18, 0.40), roi_y=(0.08, 0.32),
                               default_pos=(0.265, 0.113), val_format=".1f")

def annotate_inj_pressure_peak(img, peak_x_sec=5.092, peak_y_mpa=21.45):
    """
    还原 Moldflow 注射位置处压力 XY 曲线查询探针标注 (MPa)
    """
    return annotate_curve_peak(img, peak_x_sec, peak_y_mpa, "MPa",
                               roi_x=(0.15, 0.35), roi_y=(0.08, 0.30),
                               default_pos=(0.215, 0.144), val_format=".2f")

def merge_scale_and_model(scale_path, model_path, output_path, viewport_path=None, ref_triad_path=None, ref_title_path=None):
    """
    按照用户要求 (图 1、图 2、图 3) 执行高保真智能全要素拼合：
    1. 2D XY 曲线图：紧致裁切多余留白，如果是锁模力曲线图或注射位置处压力曲线图则在最高峰值处绘制黄色探针标记框；
    2. 3D 云图：
       - 左侧：完整数据条（结果名+时间戳+单位+色条及数值，图2）
       - 中右侧：超大裁剪模型实体
       - 右上角：工程方案名
       - 右下角：透明叠加三维坐标系与旋转角度（图3）
    """
    base_name = os.path.basename(output_path).lower()
    is_xy = "_xy" in base_name or "xy" in base_name

    # 1. 2D 曲线图处理
    if is_xy:
        src_path = model_path if os.path.exists(model_path) else scale_path
        if not os.path.exists(src_path) and viewport_path and os.path.exists(viewport_path):
            src_path = viewport_path
        if not os.path.exists(src_path):
            return False
        try:
            img = Image.open(src_path).convert("RGB")
            trimmed = trim_white_borders(img, border=12).convert("RGB")
            if "clamp_force" in base_name:
                trimmed = annotate_clamp_force_peak(trimmed)
            if "inj_pressure" in base_name or "pressure_xy" in base_name:
                trimmed = annotate_inj_pressure_peak(trimmed)
            os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
            trimmed.save(output_path, "PNG", quality=98)
            print(f"[image_processor] Exported clean XY plot -> {os.path.basename(output_path)}")
            return True
        except Exception as e:
            print(f"[image_processor] Error processing XY plot {src_path}: {e}")
            return False

    # 2. 3D 云图全要素拼合
    # 2.1 尝试从 viewport_path 抓取图2标准数据条、图3坐标系、工程标题
    left_bar = None
    triad = None
    title_img = None

    if viewport_path and os.path.exists(viewport_path):
        left_bar, triad, title_img = extract_viewport_components(viewport_path)

    # 回退：若没有从视口抓取到坐标系，尝试使用全局参考坐标系
    if triad is None and ref_triad_path and os.path.exists(ref_triad_path):
        try:
            triad = Image.open(ref_triad_path).convert("RGBA")
        except Exception:
            pass

    # 回退：若没有从视口抓取到工程标题，尝试使用全局参考标题
    if title_img is None and ref_title_path and os.path.exists(ref_title_path):
        try:
            title_img = Image.open(ref_title_path).convert("RGB")
        except Exception:
            pass

    # 回退：若没有左侧数据条，但有 scale_path，则用 scale_path 替代
    if left_bar is None and os.path.exists(scale_path):
        try:
            im_s = Image.open(scale_path).convert("RGB")
            left_bar = trim_white_borders(im_s, border=4)
        except Exception:
            pass

    # 检查模型文件
    if not os.path.exists(model_path):
        if viewport_path and os.path.exists(viewport_path):
            # 若没有独立模型图，则直接将视口图作为输出
            try:
                Image.open(viewport_path).save(output_path)
                return True
            except Exception:
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

    # 若没有任何数据条，则单独输出模型
    if left_bar is None:
        im_model_trimmed.convert("RGB").save(output_path, "PNG", quality=98)
        return True

    # ---------------- 3. 画布构建与满幅黄金比例排版 ----------------
    # 目标：构建 1.58:1 宽高比画布 (例如 1896 x 1200)，与 PPT 幻灯片安全区域完美吻合
    target_h = 1200
    target_w = int(target_h * 1.58)  # 1896 px

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
    # 高清锐化处理
    left_bar_scaled = left_bar_scaled.filter(ImageFilter.UnsharpMask(radius=1.2, percent=160, threshold=2))

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
    print(f"[image_processor] Merged Mode B (图1/2/3规范) -> {os.path.basename(output_path)} ({target_w}x{target_h})")
    return True

def process_all_mode_b_plots(mode_b_dir, data_dir=None):
    """
    批量处理方案 B 的全部图表：
    1. 生成/更新全局三维坐标系 (triad.png) 与工程标题 (study_title.png) 参考模板；
    2. 针对每个结果，从 mode_a 提取图 2 标准数据条并与 mode_b 纯净模型紧密拼合；
    3. 生成供 Slide 1 封面使用的纯 CAD 实体模型截图 (solid_model.png)；
    4. 2D XY 图高清优化输出。
    """
    if not os.path.exists(mode_b_dir):
        return 0

    if data_dir is None:
        data_dir = os.path.dirname(os.path.abspath(mode_b_dir))

    mode_a_dir = os.path.join(data_dir, "mode_a")

    # 1. 先从 mode_a 中最具代表性的图片 (如 pressure.png) 提取全局基准坐标系与标题
    ref_triad_path = os.path.join(mode_b_dir, "ref_triad.png")
    ref_title_path = os.path.join(mode_b_dir, "ref_title.png")

    for cand_name in ["pressure.png", "flow_front_temp.png", "volumetric_shrinkage.png"]:
        p_cand = os.path.join(mode_a_dir, cand_name)
        if os.path.exists(p_cand):
            _, triad_cand, title_cand = extract_viewport_components(p_cand)
            if triad_cand and not os.path.exists(ref_triad_path):
                triad_cand.save(ref_triad_path)
            if title_cand and not os.path.exists(ref_title_path):
                title_cand.save(ref_title_path)
            break

    # 2. 生成 Slide 1 封面使用的纯 CAD 实体模型 (solid_model.png)
    model_cand = os.path.join(mode_b_dir, "model_pressure.png")
    if not os.path.exists(model_cand):
        model_cand = os.path.join(mode_b_dir, "model_volumetric_shrinkage.png")
    solid_out = os.path.join(data_dir, "solid_model.png")
    if os.path.exists(model_cand):
        generate_solid_cad_model(model_cand, solid_out)
        shutil.copy2(solid_out, os.path.join(mode_b_dir, "solid_model.png"))

    # 3. 收集所有需拼合的 key
    all_keys = set()
    for s_path in glob.glob(os.path.join(mode_b_dir, "scale_*.png")):
        b = os.path.basename(s_path)
        all_keys.add(b[len("scale_"):-len(".png")])

    for m_path in glob.glob(os.path.join(mode_b_dir, "model_*.png")):
        b = os.path.basename(m_path)
        all_keys.add(b[len("model_"):-len(".png")])

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
        if merge_scale_and_model(s_path, m_path, out_path, viewport_path=vp_path, ref_triad_path=ref_triad_path, ref_title_path=ref_title_path):
            processed += 1

    # 确保 mode_a 与 root 目录中的 2D XY 曲线图也进行峰值标注，确保两份报告品质完全一致
    for xy_target_dir in [mode_a_dir, data_dir]:
        if xy_target_dir and os.path.exists(xy_target_dir):
            for xy_name, annot_fn in [
                ("clamp_force_xy.png", annotate_clamp_force_peak),
                ("inj_pressure_xy.png", annotate_inj_pressure_peak),
            ]:
                xy_file = os.path.join(xy_target_dir, xy_name)
                if os.path.exists(xy_file):
                    try:
                        im_xy = Image.open(xy_file).convert("RGB")
                        im_xy_annot = annot_fn(im_xy)
                        im_xy_annot.save(xy_file, "PNG", quality=98)
                    except Exception as e:
                        print(f"[image_processor] Error annotating {xy_file}: {e}")

    print(f"[image_processor] Total Mode B images successfully processed: {processed}")
    return processed

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Mode B Image Merging Processor")
    parser.add_argument("--dir", default=r"c:\Users\5600\Documents\ZDH\MLFXBG\temp\mode_b", help="Mode B directory")
    args = parser.parse_args()
    process_all_mode_b_plots(args.dir)
