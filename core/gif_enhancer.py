import os
import glob
from PIL import Image


# 充填动画色域广 (温度彩虹), 固定用中速中质量的自适应法; 显式常量而非误用重采样枚举
# 注: Pillow 的 Image.Quantize.MEDIANCUT == 0; 旧版无枚举时回退数值 0
# (曾误写 2 = FASTOCTREE, 与注释意图不符, 现代 Pillow 恒走枚举分支无实际影响)
_QUANT_METHOD = getattr(Image, "Quantize", None)
_QUANT_METHOD = _QUANT_METHOD.MEDIANCUT if _QUANT_METHOD else 0


def _load_frames(gif_path, max_height=None):
    # 显式关闭句柄: Windows 下同路径回写会被未关闭的文件锁挡住 (WinError 32)
    # max_height 下采样在载入时逐帧完成 —— 曾先全分辨率载入全部帧再统一缩放,
    # 2K/60 帧的峰值内存可超 1GB (每帧 RGB ~11MB, 量化副本还要再翻一倍)
    frames = []
    with Image.open(gif_path) as im:
        scale = 1.0
        if max_height and im.height > max_height:
            scale = max_height / im.height
        try:
            while True:
                fr = im.copy()
                if scale < 1.0:
                    fr = fr.resize(
                        (
                            max(1, int(fr.width * scale)),
                            max(1, int(fr.height * scale)),
                        ),
                        Image.Resampling.LANCZOS,
                    )
                frames.append(fr.convert("RGB"))
                im.seek(im.tell() + 1)
        except EOFError:
            pass
    return frames


def _global_palette(frames):
    """
    全局调色板: 均匀采样至多 5 帧纵向拼条后量化 256 色 (T11)。
    单中帧采样会让早帧色彩 (如充填初始的单一色) 欠采样, 造成跨帧劣化;
    拼条覆盖时间轴上的全部色域, 全帧共用同一色板 → 消除逐帧调色板跳变。
    """
    sample_idx = sorted(
        {round(i * (len(frames) - 1) / 4) for i in range(5)} if len(frames) > 1 else {0}
    )
    sampled = [frames[i] for i in sample_idx]
    strip = Image.new("RGB", (sampled[0].width, sampled[0].height * len(sampled)))
    for k, fr in enumerate(sampled):
        strip.paste(fr, (0, fr.height * k))
    return strip.quantize(colors=256, method=_QUANT_METHOD)


def create_gif_from_frames(frame_files, output_gif_path, delay_ms=80):
    """
    将一序列帧图片（PNG/JPG）合成高质量循环播放 GIF。
    注意: 当前管线不调用此函数 (VBS 直接导出 GIF 走 optimize_existing_gif)。
    """
    if not frame_files:
        raise ValueError("No frame files provided to create GIF")

    # 按文件名顺序排序
    frame_files = sorted(frame_files)

    images = [Image.open(f).convert("RGB") for f in frame_files]
    palette = _global_palette(images)
    quant_frames = [
        im.quantize(palette=palette, dither=Image.Dither.FLOYDSTEINBERG)
        for im in images
    ]

    quant_frames[0].save(
        output_gif_path,
        save_all=True,
        append_images=quant_frames[1:],
        duration=delay_ms,
        loop=0,
        disposal=2,
        optimize=True,
    )
    print(
        f"[GIF Enhancer] Generated high-quality GIF: {output_gif_path} ({len(quant_frames)} frames, {delay_ms}ms, global palette)"
    )
    return output_gif_path


def optimize_existing_gif(
    input_gif_path,
    output_gif_path=None,
    target_delay_ms=80,
    max_frames=None,
    max_height=None,
):
    """
    如果 Moldflow 直接导出了 GIF，对其进行循环属性（loop=0）和帧率标准化。
    T11: 全帧共用拼条采样的全局调色板, 消除逐帧独立量化导致的色彩跳动。
    T25 体积预算: max_frames 均匀抽帧 / max_height 等比降分辨率 (报告体积大头是 GIF)。
    """
    if not os.path.exists(input_gif_path):
        return None
    if output_gif_path is None:
        output_gif_path = input_gif_path

    try:
        frames = _load_frames(input_gif_path, max_height=max_height)
        if frames:
            if max_frames and len(frames) > max_frames:
                step = len(frames) / max_frames
                frames = [frames[int(i * step)] for i in range(max_frames)]
            palette = _global_palette(frames)
            quant_frames = [
                f.quantize(palette=palette, dither=Image.Dither.FLOYDSTEINBERG)
                for f in frames
            ]
            quant_frames[0].save(
                output_gif_path,
                save_all=True,
                append_images=quant_frames[1:],
                duration=target_delay_ms,
                loop=0,
                disposal=2,
                optimize=True,
            )
            size_mb = os.path.getsize(output_gif_path) / (1024 * 1024)
            print(
                f"[GIF Enhancer] Optimized GIF: {output_gif_path} ({len(quant_frames)} frames, global palette, {size_mb:.1f} MB)"
            )
            return output_gif_path
    except Exception as e:
        print(f"[GIF Enhancer] Error optimizing GIF: {e}")
    return input_gif_path


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 2:
        src = sys.argv[1]
        dst = sys.argv[2]
        if os.path.isdir(src):
            frames = glob.glob(os.path.join(src, "*.png"))
            create_gif_from_frames(frames, dst)
        elif os.path.isfile(src):
            optimize_existing_gif(src, dst)
