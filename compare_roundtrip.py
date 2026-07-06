#!/usr/bin/env python3
"""
PPT2MD Roundtrip 对比工具。

将 PPT → MD → PPTX (roundtrip)，再把原始 PPT 和恢复的 PPTX 都导出为 PNG，
裁剪白边，拼接成对比图，供肉眼比对差异。

用法:
    python compare_roundtrip.py ori.pptx
    python compare_roundtrip.py ori.pptx --no-trim-strict
    python compare_roundtrip.py ori.pptx -s 2.0 --keep-emf
"""

import sys
import os
import subprocess
import shutil
import time
import argparse
from pathlib import Path

from PIL import Image
import numpy as np

Image.MAX_IMAGE_PIXELS = 500_000_000


# ====================================================================
#  参数
# ====================================================================

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="PPT2MD Roundtrip 对比工具：原始 PPT ↔ 恢复 PPT",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("input", help="输入的 .pptx 文件路径")
    parser.add_argument("-o", "--output", default="./compare_output",
                        help="输出目录 (默认: ./compare_output)")
    parser.add_argument("-s", "--scale", type=float, default=2.0,
                        help="PNG 缩放倍率 (默认: 2.0)")
    parser.add_argument("--dpi", type=int, default=300,
                        help="输出 DPI (默认: 300)")
    parser.add_argument("--no-trim-strict", action="store_false", dest="trim_strict",
                        default=True,
                        help="宽松模式：使用 >=248 阈值裁剪白边")
    parser.add_argument("--keep-emf", action="store_true",
                        help="保留中间 EMF 文件")
    return parser


# ====================================================================
#  白边裁剪 (移植自 emf2png/src/trim_whitespace.py)
# ====================================================================

def _is_white_background(arr: np.ndarray, strict: bool = True) -> bool:
    alpha = arr[:, :, 3]
    rgb = arr[:, :, :3]

    if strict:
        white_mask = (rgb[:, :, 0] == 255) & (rgb[:, :, 1] == 255) & (rgb[:, :, 2] == 255)
    else:
        white_mask = (rgb[:, :, 0] >= 248) & (rgb[:, :, 1] >= 248) & (rgb[:, :, 2] >= 248)
    white_mask |= (alpha == 0)

    if strict:
        rows_all_white = np.all(white_mask, axis=1)
        cols_all_white = np.all(white_mask, axis=0)
    else:
        rows_all_white = np.mean(white_mask, axis=1) >= 0.95
        cols_all_white = np.mean(white_mask, axis=0) >= 0.95

    return bool(rows_all_white[0] and rows_all_white[-1]
                and cols_all_white[0] and cols_all_white[-1])


def _find_crop_bounds(arr: np.ndarray, strict: bool = True):
    alpha = arr[:, :, 3]
    rgb = arr[:, :, :3]

    if strict:
        white_mask = (rgb[:, :, 0] == 255) & (rgb[:, :, 1] == 255) & (rgb[:, :, 2] == 255)
    else:
        white_mask = (rgb[:, :, 0] >= 248) & (rgb[:, :, 1] >= 248) & (rgb[:, :, 2] >= 248)
    white_mask |= (alpha == 0)

    if strict:
        rows_all_white = np.all(white_mask, axis=1)
        cols_all_white = np.all(white_mask, axis=0)
    else:
        white_ratio_rows = np.mean(white_mask, axis=1)
        white_ratio_cols = np.mean(white_mask, axis=0)
        rows_all_white = white_ratio_rows >= 0.995
        cols_all_white = white_ratio_cols >= 0.995

    if np.all(rows_all_white) or np.all(cols_all_white):
        return None

    top = int(np.argmax(~rows_all_white))
    bottom = int(len(rows_all_white) - np.argmax(~rows_all_white[::-1]))
    left = int(np.argmax(~cols_all_white))
    right = int(len(cols_all_white) - np.argmax(~cols_all_white[::-1]))
    return (left, top, right, bottom)


def trim_white_borders(png_path: str, strict: bool = True) -> str:
    img = Image.open(png_path).convert("RGBA")
    arr = np.array(img)

    if not _is_white_background(arr, strict=strict):
        return png_path

    bounds = _find_crop_bounds(arr, strict=strict)
    if bounds is None:
        return png_path

    left, top, right, bottom = bounds
    if left == 0 and top == 0 and right == img.width and bottom == img.height:
        return png_path

    cropped = img.crop((left, top, right, bottom))
    cropped.save(png_path)
    return png_path


# ====================================================================
#  PPT → EMF (PowerPoint COM)
# ====================================================================

def pptx_to_images(input_pptx: str, output_dir: str, scale: float, dpi: int, keep_emf: bool):
    """
    将 PPTX 的每一页导出为 PNG 图片。
    流程: PPTX → EMF (COM) → PNG (PIL + 缩放) → 裁剪白边。

    Returns:
        list[str]: PNG 文件路径列表
    """
    import win32com.client
    import pythoncom
    pythoncom.CoInitialize()

    ppt_path = str(Path(input_pptx).resolve())
    out = Path(output_dir).resolve()
    out.mkdir(parents=True, exist_ok=True)

    png_dir = out / "png"
    png_dir.mkdir(parents=True, exist_ok=True)

    app = None
    pres = None
    emf_files = []

    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.DisplayAlerts = False
        pres = app.Presentations.Open(ppt_path, WithWindow=False)

        total = pres.Slides.Count
        print(f"  幻灯片总数: {total}")

        for i in range(1, total + 1):
            slide = pres.Slides(i)
            emf_path = str(out / f"slide_{i:03d}.emf")
            print(f"  导出第 {i}/{total} 页 → {Path(emf_path).name}")
            slide.Export(emf_path, "EMF", 1920, 1080)
            emf_files.append(emf_path)

    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        app = None
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    # EMF → PNG
    png_files = []
    for i, emf_path in enumerate(emf_files, start=1):
        png_path = str(png_dir / f"slide_{i:03d}.png")
        print(f"  转换 EMF → PNG: {Path(emf_path).name}")

        img = Image.open(emf_path)
        w, h = img.size
        if scale != 1.0:
            new_size = (int(w * scale), int(h * scale))
            img = img.resize(new_size, Image.Resampling.LANCZOS)
        img.save(png_path, dpi=(dpi, dpi))
        png_files.append(png_path)

        if not keep_emf:
            try:
                os.unlink(emf_path)
            except OSError:
                pass

    return png_files


# ====================================================================
#  拼接对比图
# ====================================================================

def stitch_comparison(
    original_dir: str,
    roundtrip_dir: str,
    output_dir: str,
    trim_strict: bool,
):
    """
    将原始 PPT 和恢复 PPT 的同页图片并排拼接。

    输出:
        output_dir/slide_NNN_compare.png  — 每页独立对比图
        output_dir/compare_all.html       — 浏览器查看所有对比图
    """
    orig_dir = Path(original_dir)
    rt_dir = Path(roundtrip_dir)
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    # 获取图片列表
    orig_files = sorted(orig_dir.glob("slide_*.png"))
    rt_files = sorted(rt_dir.glob("slide_*.png"))

    if not orig_files:
        print("[!] 未找到原始 PPT 的 PNG 文件")
        return
    if not rt_files:
        print("[!] 未找到恢复 PPT 的 PNG 文件")
        return

    html_parts = []

    for orig_path, rt_path in zip(orig_files, rt_files):
        slide_num = orig_path.stem

        # 裁剪白边 (copy 一份再裁，不修改原始文件)
        orig_img = Image.open(orig_path).convert("RGBA")
        rt_img = Image.open(rt_path).convert("RGBA")

        # 临时保存以便裁剪
        tmp_orig = str(out / f"_tmp_orig_{slide_num}.png")
        tmp_rt = str(out / f"_tmp_rt_{slide_num}.png")
        orig_img.save(tmp_orig)
        rt_img.save(tmp_rt)

        trim_white_borders(tmp_orig, strict=trim_strict)
        trim_white_borders(tmp_rt, strict=trim_strict)

        orig_cropped = Image.open(tmp_orig)
        rt_cropped = Image.open(tmp_rt)

        # 统一高度
        max_h = max(orig_cropped.height, rt_cropped.height)
        if orig_cropped.height < max_h:
            pad = Image.new("RGBA", (orig_cropped.width, max_h), (255, 255, 255, 255))
            pad.paste(orig_cropped, (0, 0))
            orig_cropped = pad
        if rt_cropped.height < max_h:
            pad = Image.new("RGBA", (rt_cropped.width, max_h), (255, 255, 255, 255))
            pad.paste(rt_cropped, (0, 0))
            rt_cropped = pad

        # 并排: [原始 | 恢复]
        gap = 10
        total_w = orig_cropped.width + gap + rt_cropped.width
        combined = Image.new("RGBA", (total_w, max_h), (255, 255, 255, 255))
        combined.paste(orig_cropped, (0, 0))
        combined.paste(rt_cropped, (orig_cropped.width + gap, 0))

        # 添加标注
        from PIL import ImageDraw, ImageFont
        draw = ImageDraw.Draw(combined)
        try:
            font = ImageFont.truetype("arial.ttf", 28)
        except Exception:
            font = ImageFont.load_default()

        # 原始标签
        label_orig = "Original"
        label_rt = "Roundtrip"

        # 标签背景
        for label, x_pos in [(label_orig, 10), (label_rt, orig_cropped.width + gap + 10)]:
            bbox = draw.textbbox((0, 0), label, font=font)
            tw = bbox[2] - bbox[0]
            th = bbox[3] - bbox[1]
            draw.rectangle([x_pos - 4, 4, x_pos + tw + 4, th + 8], fill=(0, 0, 0, 180))
            draw.text((x_pos, 6), label, fill=(255, 255, 255), font=font)

        compare_path = str(out / f"{slide_num}_compare.png")
        combined.save(compare_path)
        print(f"  对比图: {compare_path}")

        # 清理临时文件
        try:
            os.unlink(tmp_orig)
            os.unlink(tmp_rt)
        except OSError:
            pass

        html_parts.append((slide_num, compare_path))

    # 生成 HTML
    html = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>PPT2MD Roundtrip 对比</title>
<style>
body { font-family: sans-serif; margin: 20px; background: #f5f5f5; }
h1 { color: #333; }
.slide { margin: 30px 0; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
.slide h2 { margin-top: 0; color: #555; font-size: 18px; }
.slide img { max-width: 100%; border: 1px solid #ddd; }
</style>
</head>
<body>
<h1>PPT2MD Roundtrip 对比</h1>
<p>左: 原始 PPT &nbsp;&nbsp;|&nbsp;&nbsp; 右: 恢复 PPT</p>
"""
    for slide_num, img_path in html_parts:
        rel = os.path.relpath(img_path, out)
        html += f'<div class="slide"><h2>{slide_num}</h2><img src="{rel}" alt="{slide_num}"></div>\n'

    html += "</body></html>"
    html_path = str(out / "compare_all.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"\nHTML 对比页面: {html_path}")


# ====================================================================
#  主流程
# ====================================================================

def main():
    parser = build_parser()
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"[ERR] 文件不存在: {input_path}")
        sys.exit(1)

    out = Path(args.output)
    out.mkdir(parents=True, exist_ok=True)

    # 完整输出路径
    orig_pptx = str(input_path.resolve())
    md_dir = str(out / "md")
    rt_pptx = str(out / "roundtrip.pptx")
    orig_png_dir = str(out / "original_png")
    rt_png_dir = str(out / "roundtrip_png")
    compare_dir = str(out / "compare")

    Path(md_dir).mkdir(parents=True, exist_ok=True)
    Path(compare_dir).mkdir(parents=True, exist_ok=True)

    start_time = time.time()

    # ----------------------------------------------------------------
    # Step 1: PPT → MD (ppt2md)
    # ----------------------------------------------------------------
    print("\n[1/4] PPT → MD")
    md_output = str(Path(md_dir) / "output.md")
    # 用 subprocess 调用 ppt2md 模块
    cmd = [
        sys.executable, "-m", "ppt2md",
        orig_pptx,
        "-o", md_dir,
        "--output-file", "output.md",
        "--no-frontmatter",
    ]
    print(f"  运行: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"[ERR] ppt2md 转换失败: {result.stderr}")
        sys.exit(1)
    print("  OK")

    # ----------------------------------------------------------------
    # Step 2: MD → PPTX (roundtrip)
    # ----------------------------------------------------------------
    print("\n[2/4] MD → PPTX (roundtrip)")
    cmd_rt = [
        sys.executable, "-c",
        f"import sys; sys.path.insert(0, '.'); "
        f"from ppt2md.converter.reverse import convert_md_to_pptx; "
        f"convert_md_to_pptx('{md_output}', '{rt_pptx}')"
    ]
    print(f"  运行: ppt2md 逆向转换")
    # 直接调用 convert module
    sys.path.insert(0, os.getcwd())
    from ppt2md.converter.reverse import convert_md_to_pptx
    convert_md_to_pptx(md_output, rt_pptx)
    print(f"  OK → {rt_pptx}")

    # ----------------------------------------------------------------
    # Step 3: 原始 PPT → PNG
    # ----------------------------------------------------------------
    print("\n[3/4] 原始 PPT → PNG")
    orig_pngs = pptx_to_images(orig_pptx, str(out / "_orig_export"), args.scale, args.dpi, args.keep_emf)
    # 移动到正式目录
    Path(orig_png_dir).mkdir(parents=True, exist_ok=True)
    for f in orig_pngs:
        shutil.copy2(f, str(Path(orig_png_dir) / Path(f).name))
    print(f"  → {len(orig_pngs)} 张图片")

    # ----------------------------------------------------------------
    # Step 4: 恢复 PPT → PNG
    # ----------------------------------------------------------------
    print("\n[4/4] 恢复 PPT → PNG")
    rt_pngs = pptx_to_images(rt_pptx, str(out / "_rt_export"), args.scale, args.dpi, args.keep_emf)
    Path(rt_png_dir).mkdir(parents=True, exist_ok=True)
    for f in rt_pngs:
        shutil.copy2(f, str(Path(rt_png_dir) / Path(f).name))
    print(f"  → {len(rt_pngs)} 张图片")

    # ----------------------------------------------------------------
    # Step 5: 拼接对比图
    # ----------------------------------------------------------------
    print("\n[5/4] 拼接对比图")
    stitch_comparison(orig_png_dir, rt_png_dir, compare_dir, args.trim_strict)

    # ----------------------------------------------------------------
    # 清理临时导出目录
    # ----------------------------------------------------------------
    for d in [str(out / "_orig_export"), str(out / "_rt_export")]:
        try:
            shutil.rmtree(d, ignore_errors=True)
        except Exception:
            pass

    elapsed = time.time() - start_time
    print(f"\n{'='*50}")
    print(f"[OK] 完成! 耗时: {elapsed:.1f}s")
    print(f"    MD 文件:        {md_output}")
    print(f"    恢复 PPTX:      {rt_pptx}")
    print(f"    原始 PNG:       {orig_png_dir}")
    print(f"    恢复 PNG:       {rt_png_dir}")
    print(f"    对比图:         {compare_dir}")
    print(f"    HTML 对比页面:  {Path(compare_dir) / 'compare_all.html'}")
    print(f"{'='*50}")
    print(f"\n在浏览器中打开 compare_all.html 查看每页的左右对比图。")
    print(f"注意差异，然后修复代码，再重新运行此脚本。直到两侧完全一致。")


if __name__ == "__main__":
    main()