import os
import glob
from PIL import Image

def create_gif_from_frames(frame_files, output_gif_path, delay_ms=80):
    """
    将一序列帧图片（PNG/JPG）合成高质量循环播放 GIF
    """
    if not frame_files:
        raise ValueError("No frame files provided to create GIF")
    
    # 按文件名顺序排序
    frame_files = sorted(frame_files)
    
    images = []
    for f in frame_files:
        im = Image.open(f).convert("RGB")
        # 优化调色板为 256 色，避免色彩斑驳
        im_quant = im.quantize(colors=256, method=Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.LANCZOS)
        images.append(im_quant)
    
    if images:
        images[0].save(
            output_gif_path,
            save_all=True,
            append_images=images[1:],
            duration=delay_ms,
            loop=0,
            disposal=2,
            optimize=True
        )
        print(f"[GIF Enhancer] Generated high-quality GIF: {output_gif_path} ({len(images)} frames, {delay_ms}ms)")
        return output_gif_path
    return None

def optimize_existing_gif(input_gif_path, output_gif_path=None, target_delay_ms=80):
    """
    如果 Moldflow 直接导出了 GIF，对其进行循环属性（loop=0）和帧率标准化
    """
    if not os.path.exists(input_gif_path):
        return None
    if output_gif_path is None:
        output_gif_path = input_gif_path

    try:
        im = Image.open(input_gif_path)
        frames = []
        try:
            while True:
                frames.append(im.copy().convert("RGB"))
                im.seek(im.tell() + 1)
        except EOFError:
            pass

        if frames:
            quant_frames = [f.quantize(colors=256) for f in frames]
            quant_frames[0].save(
                output_gif_path,
                save_all=True,
                append_images=quant_frames[1:],
                duration=target_delay_ms,
                loop=0,
                disposal=2,
                optimize=True
            )
            print(f"[GIF Enhancer] Optimized GIF: {output_gif_path} ({len(frames)} frames)")
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
