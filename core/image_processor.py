"""
image_processor.py - 方案 B 专属图像智能拼合与全要素增强模块
功能：
1. 3D 结果云图：
   - 提取完整左侧数据条（包含顶部结果标题与时间戳、单位、纵向色条与清晰数值，
     按图例自身几何定位, 标题块与刻度数值一律不截断）；
   - 提取三维定向坐标系（XYZ 轴及三行旋转角度数值，做透明背景处理）；
   - 提取/渲染右上角工程方案名称标签；
   - 模型主体一律取自**方案 A 视口截图** (用户所见即所得, 几何完整、样式一致);
     离屏 SaveImage3 的 model_*.png 实机定案丢面/碎片/换样式, 仅作不可用时的兜底;
   - 按照 1.58:1 黄金比例（契合 PPT 版面安全框）合成全要素高保真云图。
2. 2D 曲线图（XY 图）：
   - Moldflow 原生 XY 截图紧致裁切输出，锁模力/注射位置处压力曲线在最高峰值处
     绘制黄色探针标记框 (annotate_xy_curves, 方案 A/B 通用; 峰值数据由调用方
     从 peak_values.json/曲线 txt 单源传入, 缺失则跳过标注, 绝不使用硬编码值)。
3. 纯模型本体图 (封面)：
   - 优先 VBS 图层自适应后的 SaveImage3 真实导出 solid_model.png (纯模型本体);
     离屏导出异常时依次回退视口模型裁剪 → 离屏结果渲染候选 (均启用割裂判据),
     全部落空返回 None, 由调用方登记缺失 (绝不臆造; 灰色重绘已停用)。
"""

import os
import re
import glob
import json
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


def _find_colour_bar(arr):
    """定位 Moldflow 左侧配色色条 (左区最高的高饱和竖列带)。

    返回 (bar_left, bar_right, bar_top, bar_bottom) 或 None。
    实机取证 (2026-09-15, 2112x1136 视口 pressure.png): 色条列 x 80..120 的高饱和
    像素占比 0.55, 其余列 (数值文字/白隙/模型) 占比 0.00 — 0.4 的判据有百倍余量。
    """
    h, w = arr.shape[:2]
    zone = arr[:, : max(1, int(w * 0.5))]
    sat = (zone.max(axis=2) - zone.min(axis=2)) > 60
    colored = sat.mean(axis=0) > 0.4
    if not colored.any():
        return None
    left = int(np.argmax(colored))
    # 含最左命中列的连续带 (允许 <=2px 破洞, 兼容色条内部分段)
    right, gap = left, 0
    for x in range(left, len(colored)):
        if colored[x]:
            right = x
            gap = 0
        else:
            gap += 1
            if gap > 2:
                break
    rows = np.where(sat[:, left : right + 1].any(axis=1))[0]
    if len(rows) == 0:
        return None
    return left, right, int(rows.min()), int(rows.max())


def _legend_right_edge(luma, bar_right):
    """图例块右缘 = 色条右侧第一条「整列全白」竖缝的左缘; 找不到返回 None。

    背景 (2026-09-15 实机取证, 2112x1136 视口): 旧实现把「色条右侧第一条白隙」
    当切断点, 实测该白隙落在**色条与数值标签之间** (band x≈130), 把标题块
    (右缘随结果名长度在 137..206 间变化) 与数值右半部一起切掉 —— 即用户反馈的
    「数据条上方的文字被截断」(pressure/volumetric_shrinkage/flow_front_temp/
    weld_lines/warpage_all/vp_switch_pressure 六张全部命中, cut 130~153 < 标题右缘)。
    新判据要求白缝宽 >= 8px **且缝右侧必须还有内容 (模型)**, 从而跳过色条与数值
    之间的小白隙, 停在图例与模型之间那条真正的分隔缝上 (实测 pressure 缝 158..289,
    vp_switch_pressure 缝 207..289, 两者右缘都恰好 = 对应标题块右缘 + 1)。

    扫描前屏蔽底部水印带 (y > 0.88h): 水印的非白像素会把每一列都算成「有内容」,
    任何白缝都检不出来 — 旧实现即因此退化为粗暴兜底 (bar_right + 45), 这正是
    截断的直接触发路径。
    """
    h, w = luma.shape
    band = luma[: max(1, int(h * 0.88))]
    nonwhite = (band < 245).mean(axis=0)
    min_gap = max(8, w // 150)
    run = 0
    for x in range(bar_right + 2, w):
        if nonwhite[x] < 0.005:
            run += 1
            continue
        if run >= min_gap and nonwhite[x : x + 5].max() > 0.02:
            return x - run
        run = 0
    return None


def legend_box(im):
    """左侧**完整**数据条包围盒 (x0,y0,x1,y1); 定位失败返回 None。

    纵向由色条界定: 上界取图例列内最上方的内容行 (结果名/时间/单位块), 下界取
    色条底部 + 8% 色条高 (包住末位刻度数值)。底部水印 (AUTODESK MOLDFLOW
    INSIGHT, 实测 y 1040..1120) 天然落在下界 (≈910) 之外 —— 旧实现把它裁进数据条,
    既让数据条白占 250px 高、又把刻度数值等比缩小。
    """
    luma = np.asarray(im.convert("L"))
    h, w = luma.shape
    arr = np.asarray(im.convert("RGB")).astype(np.int16)
    bar = _find_colour_bar(arr)
    if bar is None:
        # 无色条 (灰度色标/无图例): 退回保守区域, 只取左上角图例列
        cut = max(60, int(w * 0.16))
        y_bot = int(h * 0.86)
    else:
        bar_left, bar_right, bar_top, bar_bottom = bar
        edge = _legend_right_edge(luma, bar_right)
        if edge is None:
            bar_w = bar_right - bar_left + 1
            cut = min(w, bar_right + max(int(bar_w * 2.0), 60))
        else:
            # 缝内全是白列: 向缝内多取几像素只会多出白边, 绝不会截断文字
            cut = min(w, edge + min(8, max(2, w // 300)))
        y_bot = min(h, bar_bottom + max(10, int((bar_bottom - bar_top) * 0.08)))
    region = luma[:y_bot, :cut]
    mask = region < 240
    if not mask.any():
        return None
    ys, xs = np.where(mask)
    x0 = max(0, int(xs.min()) - 2)
    y0 = max(0, int(ys.min()) - 2)
    x1 = min(cut, int(xs.max()) + 5)
    y1 = min(y_bot, int(ys.max()) + 5)
    if x1 - x0 < 8 or y1 - y0 < 8:
        return None
    return x0, y0, x1, y1


def extract_viewport_components(viewport_path):
    """
    从 Moldflow 视口抓取图 (如 mode_a/pressure.png) 中提取全要素：
    1. left_bar: 左侧**完整**数据条（结果名+时间戳+单位+色条+全部刻度数值，图2）
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

    # 1. 提取左侧数据条 (图2) — 含标题块(结果名+时间)、单位、色条与全部数值。
    # 实机包围盒 (2026-09-08, 2112x1136 视口): 标题 y29-59, 单位 y191,
    # 色条 y228-870, 数值 y1093-1120, 左起 x15 — 裁剪框必须完整包住。
    # (2026-09-15 改写: 由固定 340px 带 + 白隙切断 改为 legend_box 自几何定位,
    #  原实现把标题块与数值右半部一并切掉, 详见 _legend_right_edge 注释。)
    left_bar = None
    box = legend_box(im)
    if box:
        bar_crop = im.crop(box)
        if bar_crop.width > w * 0.35:
            print(
                f"[image_processor] WARN: 数据条宽 {bar_crop.width} 超过画面 35%, "
                f"疑把模型左缘卷进图例 (视口 {w}x{h} 偏离设计尺寸?)"
            )
        left_bar = bar_crop

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


def _components_bbox(mask, stride):
    """非白掩码的 4 连通域列表 [(面积, x0, y0, x1, y1)] (stride 降采样后回标坐标)。

    降采样跑连通域: 2K 视口原图 240 万像素, 纯 Python 扫描过慢; stride 归一到
    200px 量级后仅约 4 万格, 且降采样会顺带抹平模型内部 1-2px 的浅色接缝
    (让同一实体更易判为单连通域), 对"找出模型主体 + 剔除视口铬层"足够。
    """
    small = mask[::stride, ::stride]
    sh, sw = small.shape
    seen = np.zeros_like(small, dtype=bool)
    out = []
    for sy in range(sh):
        if not small[sy].any():
            continue
        for sx in range(sw):
            if not small[sy, sx] or seen[sy, sx]:
                continue
            stack = [(sy, sx)]
            seen[sy, sx] = True
            n = 0
            y0 = y1 = sy
            x0 = x1 = sx
            while stack:
                cy, cx = stack.pop()
                n += 1
                if cy < y0:
                    y0 = cy
                elif cy > y1:
                    y1 = cy
                if cx < x0:
                    x0 = cx
                elif cx > x1:
                    x1 = cx
                for dy, dx in ((1, 0), (-1, 0), (0, 1), (0, -1)):
                    ny, nx = cy + dy, cx + dx
                    if (
                        0 <= ny < sh
                        and 0 <= nx < sw
                        and small[ny, nx]
                        and not seen[ny, nx]
                    ):
                        seen[ny, nx] = True
                        stack.append((ny, nx))
            out.append((n, x0, y0, x1, y1))
    return out


def _max_true_run(row):
    """一维布尔数组里最长连续 True 段的长度 (0 = 全 False)。"""
    if not row.any():
        return 0
    edges = np.flatnonzero(np.diff(np.concatenate(([0], row.view(np.int8), [0]))))
    return int((edges[1::2] - edges[::2]).max())


def erase_scale_bar(rgb, luma, y_from=0.85):
    """原地抹白视口底部 Moldflow 缩放比例尺铬层; 返回抹除的横线行数。

    实机取证 (2026-09-15, 2112x1136 视口): 比例尺横线占 y 1084-1085 两行, 单行最长
    连续非白段 957px (占区域宽 0.49), 上下各 3 行覆盖不到它的 55%; 它在多处被模型
    实体压住, 于是纯像素级连通域里与模型主体连成一体, 靠"剔除小块"去不掉 ——
    结果就是拼合图底边露出一段被切断的标尺 + 半截"缩放 (300 mm)"文字 (边界不清晰)。

    两步抹除 (只动无彩色像素, 模型实体的饱和色像素与灰底厚实体都原样保留):
    1) 比例尺主线: 按"细横线"识别 (整行有长连续段 + 上下行稀疏), 且该行像素纵向
       厚度 <= 4px 才算铬层 —— 模型实体在 ±4 行窗口内厚度必然 >=5px, 故被盖住的
       那几段比例尺不会被误当成模型, 模型轮廓也不受任何影响;
    2) 底部小尺寸无彩色块: 比例尺刻度线 / "缩放 (300 mm)" 说明文字 / 水印残尾。
    抹除后这些结构不再是连通域, 模型包围盒自然停在实体边界上。
    """
    h, w = luma.shape
    strict = luma < 245
    # 抹除用宽松阈值: 铬层线条带 1-2px 抗锯齿浅灰光晕 (实测 L≈246-252), 仍按严格
    # 阈值判结构会留下一条"淡灰细线"残影 —— 实测首版正是如此 (比例尺主线抹掉后
    # 底边仍可见 1px 淡线 + 断续文字残点)。
    loose = luma < 253
    sat = rgb.max(axis=2) - rgb.min(axis=2)
    achromatic = sat < 60
    y0 = max(0, int(h * y_from))
    removed_rows = 0

    for y in range(y0, h):
        cover = float(strict[y].mean())
        if cover < 0.06 or _max_true_run(strict[y]) < max(300, int(0.30 * w)):
            continue
        up = float(strict[max(0, y - 3) : y].mean())
        dn = float(strict[y + 1 : y + 4].mean())
        if up >= 0.55 * cover or dn >= 0.55 * cover:
            continue
        # 纵向厚度用宽松阈值量: 模型实体 (含抗锯齿边) 在 ±4 行窗口内必然 >=7px
        thin = loose[max(0, y - 4) : min(h, y + 5)].sum(axis=0) <= 6
        hit = False
        for yy in range(max(0, y - 1), min(h, y + 2)):
            sel = loose[yy] & achromatic[yy] & thin
            if sel.any():
                rgb[yy][sel] = 255
                strict[yy][sel] = False
                loose[yy][sel] = False
                hit = True
        if hit:
            removed_rows += 1

    strip = strict[y0:] & achromatic[y0:]
    if strip.any():
        stride = 2
        small_thr = max(64, int(0.005 * strip.size))
        for area, cx0, cy0, cx1, cy1 in _components_bbox(strip, stride):
            if (
                area * stride * stride >= small_thr
                or (cy1 - cy0 + 1) * stride >= 0.06 * h
            ):
                continue
            ay0 = max(y0, y0 + cy0 * stride - stride)
            ay1 = min(h, y0 + (cy1 + 1) * stride + stride)
            ax0 = max(0, cx0 * stride - stride)
            ax1 = min(w, (cx1 + 1) * stride + stride)
            block = loose[ay0:ay1, ax0:ax1] & achromatic[ay0:ay1, ax0:ax1]
            if block.any():
                rgb[ay0:ay1, ax0:ax1][block] = 255
                strict[ay0:ay1, ax0:ax1][block] = False
                loose[ay0:ay1, ax0:ax1][block] = False
    return removed_rows


def extract_viewport_model(viewport_path, exclude_x=None):
    """从方案 A 视口截图裁出**完整**的模型主体图; 取景不可信时返回 None。

    背景 (2026-09-15 实机取证, 配 2560x1440): SaveImage3 离屏导出的 model_*.png
    在本机丢面/留悬浮碎片 —— 同一方案 7 张里 pressure / volumetric_shrinkage /
    flow_front_temp / vp_switch_pressure 四张缺掉模型左下半张面 (连通域 = 主体
    97.7% + 悬浮碎片 2.2%), weld_lines 被渲染成半透明线框样式, 与视口所见
    (solid 实色) 不一致。这正是用户反复反馈、且此前多次修改仍未解决的
    "模型割裂/不完整"的直接来源 —— 离屏渲染路径不可信, 故方案 B 的模型一律
    取自视口截图 (用户所见即所得, 几何完整、样式一致)。

    实现: 在 exclude_x (图例右缘) 右侧做非白连通域分析, 保留主连通域与显著
    伴随连通域 (流道/水路等独立实体), 剔除视口铬层 (右侧工具栏图标、右下
    坐标系、底部比例尺等小块), 取并集包围盒裁剪。底部的缩放比例尺与模型实体
    在像素上粘连 (无法按连通域剔除), 先按"细横线"特征抹白 (见 erase_scale_bar)。
    可信度闸: 模型须占画面 >=25% 宽高且 >=3% 面积, 否则判为"取景不可信"
    返回 None, 由调用方回退离屏模型 (合成测试图/异常取景不会误裁出小块)。
    """
    if not viewport_path or not os.path.exists(viewport_path):
        return None
    try:
        im = Image.open(viewport_path).convert("RGB")
    except Exception as e:
        print(f"[image_processor] 视口模型读取失败 ({viewport_path}): {e}")
        return None
    w, h = im.size
    if exclude_x is None:
        box = legend_box(im)
        # 图例右缘 + 2px: 再窄也可能把图例右端文字切进模型裁剪区
        exclude_x = (box[2] + 2) if box else 0
    x_off = max(0, int(exclude_x))
    region = im.crop((x_off, 0, w, h))
    rw, rh = region.size
    if rw < 80 or rh < 80:
        return None

    arr = np.asarray(region).copy()
    luma = np.asarray(region.convert("L"))
    stripped = erase_scale_bar(arr, luma)
    if stripped:
        region = Image.fromarray(arr)
        luma = np.asarray(region.convert("L"))
        print(
            f"[image_processor] 已抹除视口底部比例尺铬层 {stripped} 行 (模型实色像素保留)"
        )
    mask = luma < 245
    stride = max(4, min(rw, rh) // 200)
    comps = _components_bbox(mask, stride)
    if not comps:
        return None
    comps.sort(key=lambda c: c[0], reverse=True)
    largest = comps[0][0]
    keep = []
    for area, cx0, cy0, cx1, cy1 in comps:
        if area < max(max(1, 3000 // (stride * stride)), int(largest * 0.003)):
            continue
        ax0 = cx0 * stride
        ay0 = cy0 * stride
        ax1 = min(rw, cx1 * stride + stride)
        ay1 = min(rh, cy1 * stride + stride)
        # 视口铬层: 右侧工具栏竖带 / 底部比例尺带。只判小块 (>=25% 主体的大块
        # 一律视为模型本体, 避免把贴合右侧的宽模型误删)。
        if ax0 >= 0.90 * rw and area < 0.25 * largest:
            continue
        if ay0 >= 0.92 * rh and area < 0.25 * largest:
            continue
        keep.append((ax0, ay0, ax1, ay1))
    if not keep:
        return None

    bx0 = min(k[0] for k in keep)
    by0 = min(k[1] for k in keep)
    bx1 = max(k[2] for k in keep)
    by1 = max(k[3] for k in keep)
    mw, mh = bx1 - bx0, by1 - by0
    if (
        mw < 0.25 * rw
        or mh < 0.25 * rh
        or mw * mh < 0.03 * rw * rh
        or mw < 2 * stride
        or mh < 2 * stride
    ):
        print(
            f"[image_processor] 视口模型取景不可信 ({mw}x{mh} 于 {rw}x{rh}), "
            "不裁用 (调用方回退离屏模型)"
        )
        return None
    if bx0 <= stride or by0 <= stride or bx1 >= rw - stride or by1 >= rh - stride:
        print(
            "[image_processor] WARN: 视口模型触及画面边缘, 可能被窗口裁切 — "
            "建议在 Moldflow 里缩小视图 (全部入框) 后重跑"
        )
    return region.crop((bx0, by0, bx1, by1))


def _model_image_usable(path, reject_fragments=False):
    """模型图完整性判据: 非空白且非断带 (与 merge 三重护栏同判据)。
    reject_fragments=True 时额外否决"主体 + 悬浮碎片"式割裂图
    (SaveImage3 离屏渲染的实发特征: 主体 97.7% + 碎片 2.2%, 且碎片与主体
    同为模型面片却颜色/位置错乱) —— 仅用于已知不可靠的离屏候选;
    VBS 纯模型直出与视口裁剪允许天然多实体 (流道/水路独立体)。
    返回 (ok, 原因)。"""
    try:
        with Image.open(path) as im:
            arr = np.array(im.convert("L"))
        if (arr < 240).mean() <= 0.005:
            return False, "空白图"
        h, w = arr.shape
        ys, xs = np.where(arr < 240)
        ch = int(ys.max() - ys.min()) + 1
        cw = int(xs.max() - xs.min()) + 1
        h_ratio = ch / float(h)
        aspect = cw / float(ch)
        if h_ratio < 0.35 and aspect > 2.2:
            return False, f"断带特征 (高度占比 {h_ratio:.0%}, 宽高比 {aspect:.2f})"
        if reject_fragments:
            comps = _components_bbox(arr < 240, max(4, min(w, h) // 200))
            if comps:
                comps.sort(key=lambda c: c[0], reverse=True)
                largest = comps[0][0]
                extra = sum(
                    a for a, *_ in comps[1:] if a >= max(1, int(largest * 0.01))
                )
                if extra > 0:
                    return False, (
                        f"割裂特征 (主体 {largest} px + 悬浮碎片 {extra} px, "
                        f"{len(comps)} 个连通域)"
                    )
        return True, ""
    except Exception as e:
        return False, str(e)


def resolve_cover_image(data_dir):
    """封面模型本体图解析。

    候选顺序 (2026-09-15 修订):
    1. data_dir/solid_model.png — VBS 用 SaveImage3 离屏直出的**纯模型本体图**
       (红色网格着色, 与分析模型一致)。实机取证: 该导出图连通域 = 1 个且占满
       画面 (完整), 是设计上的正解。
    2. mode_a/ 与 mode_b/ 下的同名副本 (VBS 导出后拷贝, 内容一致)。
    3. mode_a/<结果图> 的视口模型裁剪 — 完整但为结果着色 (封面次选)。
    4. mode_b/model_pressure.png / model_volumetric_shrinkage.png — 离屏结果渲染,
       实机定案不可靠 (丢面 + 悬浮碎片), 仅最后兜底并启用割裂判据。

    原实现把 4 提到第 1 位, 于是封面被残缺的离屏图覆写 — 用户反馈的
    "第一页模型本体图割裂" 直接来源。
    命中后裁白边写回 data_dir/solid_model.png 并返回其路径; 全部落空返回 None。
    """
    mode_a_dir = os.path.join(data_dir, "mode_a")
    mode_b_dir = os.path.join(data_dir, "mode_b")
    fallback = os.path.join(data_dir, "solid_model.png")

    def _accept(cand, reject_fragments, tag):
        ok, reason = _model_image_usable(cand, reject_fragments=reject_fragments)
        if not ok:
            print(
                f"[image_processor] 封面候选 {os.path.basename(cand)} 不可用: {reason}"
            )
            return None
        out = os.path.join(data_dir, "solid_model.png")
        if os.path.abspath(cand) == os.path.abspath(out):
            print(f"[image_processor] 封面模型本体直出: {tag}")
            return out
        trimmed = trim_white_borders(Image.open(cand), border=8)
        os.makedirs(data_dir, exist_ok=True)
        trimmed.convert("RGB").save(out, "PNG")
        print(f"[image_processor] 封面模型本体直出: {tag} -> solid_model.png")
        return out

    for cand, tag in [
        (fallback, "VBS 离屏纯模型导出 solid_model.png"),
        (os.path.join(mode_a_dir, "solid_model.png"), "mode_a/solid_model.png"),
        (os.path.join(mode_b_dir, "solid_model.png"), "mode_b/solid_model.png"),
    ]:
        if os.path.exists(cand):
            # 同样启用割裂判据: 这三个文件在正常一次运行里内容一致 (VBS 导出一份
            # 再拷贝), 但 data_dir/solid_model.png 可能被历史版本用残缺离屏图覆写过
            # —— 拒掉它才能让候选链继续走到 mode_a 的完整导出。
            got = _accept(cand, True, tag)
            if got:
                return got

    # 次选: 视口模型裁剪 (完整, 但为结果着色)
    for key in ("pressure", "volumetric_shrinkage", "flow_front_temp"):
        vp = os.path.join(mode_a_dir, f"{key}.png")
        if not os.path.exists(vp):
            continue
        vp_model = extract_viewport_model(vp)
        if vp_model is None:
            continue
        os.makedirs(data_dir, exist_ok=True)
        vp_model.convert("RGB").save(fallback, "PNG")
        print(
            f"[image_processor] 封面回退视口模型裁剪: mode_a/{key}.png -> solid_model.png"
        )
        return fallback

    # 最后兜底: 离屏结果渲染候选 (启用割裂判据)
    for cand in [
        os.path.join(mode_b_dir, "model_pressure.png"),
        os.path.join(mode_b_dir, "model_volumetric_shrinkage.png"),
    ]:
        if os.path.exists(cand):
            got = _accept(cand, True, os.path.basename(cand))
            if got:
                return got
    return None


def generate_solid_cad_model(model_source_path, output_path):
    """
    [2026-09-08 停用] 灰色 CAD 重着色封面 — 用户裁决该样式与真实实体不符, 为
    错误样式; 封面改走 resolve_cover_image (模型本体图直出)。函数保留但不再被
    调用 (如需恢复须显式确认)。

    从纯模型着色图生成 100% 纯净、无网格、无节点的 CAD 实体本体图 (供 Slide 1 封面使用)
    采用 Autodesk 经典酷炫 CAD 金属灰/冷蓝灰材质着色与高光阴影
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
        out_trimmed.save(output_path, "PNG")
        print(
            f"[image_processor] Created pure CAD solid model (no mesh, no nodes) -> {output_path}"
        )
        return True
    except Exception as e:
        print(f"[image_processor] Error generating solid CAD model: {e}")
        return False


def locate_curve_peak(img, roi_x, roi_y, default_pos):
    """统一的 XY 曲线峰值定位 (方案 A/B 共用, 逐级降级且每次降级打印原因)。

    背景 (2026-09-08 实机取证): 方案 B 对 model_*.png 先 trim_white_borders
    再标注, 裁剪后曲线峰顶 (实测 x≈0.145w) 落在按方案 A 视图标定的 ROI
    (x 起点 0.18) 之外 → 检测不到黑像素 → 回退 default_pos, 探针框浮空。

    降级链:
    1) 主: 坐标轴框定位 — 最长竖直黑线为 y 轴、最长水平黑线为 x 轴
       (连线长度 >= 短边 30% 才算), 在轴框内部 (内缩 margin, 天然排除
       框外左侧/底部的刻度文字与框上方标题) 自上而下取最高黑像素;
    2) 备: 调用方 ROI 内最高黑像素 (方案 A 视口图原逻辑, 保留兜底);
    3) 兜底: default_pos。
    返回 (peak_x, peak_y, method)。
    """
    im = img.convert("RGB")
    w, h = im.size
    # 曲线/轴线/刻度文字实测为深灰到黑 (灰度 0-130, 纯黑<40 检不到轴线),
    # 网格线为浅灰 (~203) 不入掩码。
    gray = np.array(im.convert("L"))
    dark = gray < 150
    return _locate_curve_peak_vec(dark, w, h, roi_x, roi_y, default_pos)


def _locate_curve_peak_vec(dark, w, h, roi_x, roi_y, default_pos):
    """locate_curve_peak 的向量化实现 (列/行最长连续黑段 + 框内最高点)。"""
    min_run = max(8, int(min(w, h) * 0.30))

    def longest_run_profile(mask_2d):
        """每行/列最长连续 True 段长度与起点 (向量化: 逐行 np.diff 分段)。"""
        n = mask_2d.shape[0]
        best_len = np.zeros(n, dtype=np.int64)
        best_start = np.zeros(n, dtype=np.int64)
        padded = np.concatenate(
            [np.zeros((n, 1), dtype=bool), mask_2d, np.zeros((n, 1), dtype=bool)],
            axis=1,
        )
        d = np.diff(padded.astype(np.int8), axis=1)
        for i in range(n):
            starts = np.where(d[i] == 1)[0]
            ends = np.where(d[i] == -1)[0]
            if len(starts) == 0:
                continue
            lens = ends - starts
            j = int(np.argmax(lens))
            best_len[i] = lens[j]
            best_start[i] = starts[j]
        return best_len, best_start

    # 注意方向: longest_run_profile(m) 返回 m 逐行(沿 axis=1)的最长连续段。
    # dark 形状为 (h, w): 行内段 = 水平线 → x 轴; dark.T 的行 = 图像列 → y 轴。
    row_len, row_start = longest_run_profile(dark)  # 水平段 (每行)
    col_len, col_start = longest_run_profile(dark.T)  # 竖直段 (每列)

    y_axis_x = int(np.argmax(col_len))
    x_axis_y = int(np.argmax(row_len))
    if int(col_len[y_axis_x]) < min_run or int(row_len[x_axis_y]) < min_run:
        print(
            f"[image_processor] 峰值定位: 轴线检测失败 "
            f"(竖线最长 {int(col_len.max())}px/横线最长 {int(row_len.max())}px, "
            f"阈值 {min_run}px), 降级 ROI 扫描"
        )
    else:
        y_top = int(col_start[y_axis_x])
        # 上边距用 2.5% 图高: Moldflow 标题 (如"锁模力:XY 图")横跨绘图区顶线,
        # 3px 余量会让标题字底漏进框内被当成峰顶 (2026-09-08 实测 (0.47w,0.06h) 误检)
        margin = max(3, int(h * 0.025))
        # 右界 = x 轴线右端: 视口图 x 轴外侧是工具栏图标 (深色小块),
        # 不限制右界会命中 (0.97w, 0.08h) 的图标像素 (2026-09-08 实测)
        x_axis_l = int(row_start[x_axis_y])
        x_axis_r = x_axis_l + int(row_len[x_axis_y])
        x0, x1 = y_axis_x + margin, x_axis_r - margin
        y0, y1 = y_top + margin, x_axis_y - margin
        if x1 > x0 and y1 > y0:
            frame = dark[y0:y1, x0:x1]
            if frame.any():
                # 按行密度剔除文字行: 标题/图例 (如"锁模力:XY 图") 的暗像素数
                # 远高于曲线本身, 逐行下扫时跳过密集行, 首个"稀疏行"即峰顶所在行
                # (2026-09-08: 仅靠上边距无法稳定排除横跨绘图区顶线的标题)
                row_counts = frame.sum(axis=1)
                dens_thr = max(12, int((x1 - x0) * 0.02))
                for i, c in enumerate(row_counts):
                    if c == 0 or c > dens_thr:
                        continue
                    xs = np.where(frame[i])[0]
                    if len(xs):
                        return x0 + int(xs[0]), y0 + i, "axes-frame"
        print("[image_processor] 峰值定位: 轴框内无黑像素, 降级 ROI 扫描")

    # 2) ROI 扫描 (方案 A 原逻辑)
    x_start, x_end = int(roi_x[0] * w), int(roi_x[1] * w)
    y_start, y_end = int(roi_y[0] * h), int(roi_y[1] * h)
    if x_end > x_start and y_end > y_start:
        sub = dark[y_start:y_end, x_start:x_end]
        ys, xs = np.where(sub)
        if len(ys) > 0:
            i = int(np.argmin(ys))
            return x_start + int(xs[i]), y_start + int(ys[i]), "roi"
    print("[image_processor] 峰值定位: ROI 内无黑像素, 降级 default_pos")

    # 3) default_pos
    return int(default_pos[0] * w), int(default_pos[1] * h), "default"


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

    # 智能定位峰值拐点 (方案 A/B 统一: 轴框定位 → ROI → default_pos 逐级降级)
    peak_x, peak_y, _method = locate_curve_peak(im, roi_x, roi_y, default_pos)

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
    # 判据收紧为 "_xy": 裸 "xy" 会误命中含 xy 子串的 key (如 oxygen),
    # 使普通 3D 云图被当成 2D 曲线图走错误管线 (C-04)
    is_xy = "_xy" in base_name

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
            trimmed.save(output_path, "PNG")
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

    # 2.2 模型来源 (2026-09-17 修订回: 视口优先)
    # 实机取证: 离屏导出 (SaveImage3) 屡次丢面/只渲染局部 (用户反馈 5/6/8/11 页
    # "图片不完整"全部来自离屏), 而"实体混入"问题已由 VBS 侧"临时隐藏层"在
    # 源头解决 (视口截图认图层隐藏), 所以视口裁剪恢复为主源 (恒完整),
    # 离屏仅在视口模型不可用时兜底, 兜底时仍跑全套完整性校验。
    im_model_raw = None
    if viewport_path and os.path.exists(viewport_path):
        vp_model = extract_viewport_model(viewport_path)
        if vp_model is not None:
            im_model_raw = vp_model.convert("RGBA")
            print(
                f"[image_processor] 模型取自视口完整画面 ({vp_model.width}x{vp_model.height}): "
                f"{os.path.basename(viewport_path)}"
            )

    if im_model_raw is None:
        if not os.path.exists(model_path):
            if viewport_path and os.path.exists(viewport_path):
                # 无独立模型图且视口模型不可用 → 直接将视口图作为输出
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

        # 兜底路径必须过完整性校验 (离屏导出可靠性差):
        # 校验 1 (主): 实际输出分辨率 != 配置分辨率 → 导出布局错乱, 弃用;
        # 校验 2: 内容占画布面积比过低 = 渲染残缺;
        # 校验 3: 断带/割裂特征 (主体 + 悬浮碎片)。
        # 命中即回退整张视口图 (完整正确, 与方案 A 同级保底, 绝不输出残缺图)。
        size_mismatch = bool(
            expected_model_size
            and tuple(im_model_raw.size) != tuple(expected_model_size)
            and all(v > 0 for v in expected_model_size)
        )
        im_model_trimmed_chk = trim_white_borders(im_model_raw, border=8)
        area_ratio = (im_model_trimmed_chk.width * im_model_trimmed_chk.height) / float(
            im_model_raw.width * im_model_raw.height
        )
        h_ratio = im_model_trimmed_chk.height / float(im_model_raw.height)
        aspect = im_model_trimmed_chk.width / float(im_model_trimmed_chk.height)
        band_broken = h_ratio < 0.35 and aspect > 2.2
        frag_ok, frag_reason = _model_image_usable(model_path, reject_fragments=True)
        if size_mismatch or area_ratio < 0.10 or band_broken or not frag_ok:
            if size_mismatch:
                reason = (
                    f"输出 {im_model_raw.width}x{im_model_raw.height} != 配置 "
                    f"{expected_model_size[0]}x{expected_model_size[1]}"
                )
            elif band_broken:
                reason = (
                    f"渲染断带特征 (内容高度占比 {h_ratio:.0%}, 宽高比 {aspect:.2f})"
                )
            elif not frag_ok:
                reason = frag_reason
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

    # 紧致裁切模型多余白边
    im_model_trimmed = trim_white_borders(im_model_raw, border=8)
    if im_model_trimmed.height == 0 or im_model_trimmed.width == 0:
        return False
        im_model_trimmed.convert("RGB").save(output_path, "PNG")
        return True

    # 若没有任何数据条，则单独输出模型
    if left_bar is None:
        im_model_trimmed.convert("RGB").save(output_path, "PNG")
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

    # 缩放左侧数据条：高度协调匹配 (先于模型缩放, 模型宽度约束依赖 lb_w)
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

    # 缩放模型：高度占用 ~94% 画布高度实现超大图幅；
    # 宽度双约束 (2026-09-08 修复"模型截图不完整"): 宽模型此前只按高度缩放,
    # m_w 超出可用宽度后从 rem_x 起粘贴、右侧溢出画布被裁 — 实发用户截图。
    # rem_w 与下方放置段保持同一几何 (lb_x=35, 间距 20/20)。
    model_avail_h = target_h - 70
    rem_x = 35 + lb_w + 20
    rem_w = target_w - rem_x - 20
    m_scale = min(
        model_avail_h / float(im_model_trimmed.height),
        rem_w / float(im_model_trimmed.width),
    )
    if m_scale < model_avail_h / float(im_model_trimmed.height):
        print(
            f"[image_processor] 模型受画布宽度约束缩放 "
            f"(x{m_scale:.3f} < 高度约束 x{model_avail_h / im_model_trimmed.height:.3f}), "
            "保证完整不入裁"
        )
    m_w = max(1, int(im_model_trimmed.width * m_scale))
    m_h = max(1, int(im_model_trimmed.height * m_scale))
    model_scaled = im_model_trimmed.resize((m_w, m_h), Image.Resampling.LANCZOS)

    # 创建纯白高清画布
    canvas = Image.new("RGB", (target_w, target_h), (255, 255, 255))

    # 1. 放置左侧数据条 (图2)
    lb_x = 35
    lb_y = 35
    canvas.paste(left_bar_scaled, (lb_x, lb_y))

    # 2. 放置超大模型 (居中偏右; m_scale 已含宽度约束, 恒完整)
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
    canvas.save(output_path, "PNG")
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

    # 2. 封面模型本体图: 已迁移至 resolve_cover_image (由 pptx_builder 在构建
    #    封面时调用, 方案 A/B 共用同一来源; 2026-09-08 用户二次裁决直出模型图)。
    #    此处不再处理, 避免与 build_single_report 双写。

    # 3. 收集所有需拼合的 key
    all_keys = set()
    for s_path in glob.glob(os.path.join(mode_b_dir, "scale_*.png")):
        b = os.path.basename(s_path)
        all_keys.add(b[len("scale_") : -len(".png")])

    for m_path in glob.glob(os.path.join(mode_b_dir, "model_*.png")):
        b = os.path.basename(m_path)
        all_keys.add(b[len("model_") : -len(".png")])

    # 也检查 mode_a 中的 key (排除 VBS 直接拷入的非结果图: solid_model/
    # mesh_model 走封面/网格专用链路, ref_* 是参照物 —— 它们没有 scale_/
    # model_ 对, 进来只会走"视口直出"原样复制一遍, 纯属重复 IO)
    NON_RESULT_PREFIXES = ("solid_model", "mesh_model", "ref_")
    if os.path.exists(mode_a_dir):
        for a_path in glob.glob(os.path.join(mode_a_dir, "*.png")):
            b = os.path.basename(a_path)
            if b.startswith(NON_RESULT_PREFIXES):
                continue
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
                im_xy_annot.save(xy_file, "PNG")
            except Exception as e:
                print(f"[image_processor] Error annotating {xy_file}: {e}")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Mode B Image Merging Processor")
    parser.add_argument(
        "--dir",
        default=None,
        help="Mode B directory",
    )
    args = parser.parse_args()
    if args.dir is None:
        args.dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "temp",
            "mode_b",
        )
    process_all_mode_b_plots(args.dir)
