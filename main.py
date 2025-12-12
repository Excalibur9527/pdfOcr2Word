"""
PDF OCR 转 Word 命令行工具

Author: Excalibur9527
"""

import argparse
import os
from typing import Optional

from tqdm import tqdm

from converter import convert_pdf_to_docx


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="将 PDF 通过 OCR 转成可编辑的 Word (.docx) 文档"
    )
    parser.add_argument("input_pdf", help="输入 PDF 文件路径")
    parser.add_argument("output_docx", help="输出 Word 文件路径（可不带 .docx 后缀）")
    parser.add_argument(
        "--mode",
        choices=["ocr", "text", "mac"],
        default="ocr",
        help="处理模式：ocr（默认）用 tesseract/paddle；text 直接提取文本层；mac 用 macOS 原生 Vision OCR（需装 PyMuPDF 与 ocrmac）。",
    )
    parser.add_argument(
        "--no-format",
        dest="auto_format",
        action="store_false",
        help="关闭简单格式化（默认开启）。",
    )
    parser.set_defaults(auto_format=True)
    parser.add_argument(
        "--engine",
        choices=["tesseract", "paddle"],
        default="tesseract",
        help="OCR 引擎：tesseract 或 paddle，默认 tesseract。中文建议用 paddle。",
    )
    parser.add_argument(
        "--lang",
        default=None,
        help="语言/模型代码：tesseract 如 chi_sim、eng；paddle 如 ch、en。不填则按引擎给默认值。",
    )
    parser.add_argument(
        "--remove-token",
        action="append",
        dest="remove_tokens",
        default=None,
        help="仅在 --mode text 时有效；用于删除页眉/页脚等干扰词，可多次传入。",
    )
    parser.add_argument(
        "--font-name",
        default="SimSun",
        help="输出文档使用的字体名称，默认 SimSun（宋体）。",
    )
    parser.add_argument(
        "--font-size",
        type=int,
        default=12,
        help="输出文档字号（磅值），默认 12。",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=300,
        help="PDF 渲染分辨率，默认 300，越高越清晰但越慢",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=None,
        help="OCR 并发线程数，默认不限制（由线程池自动决定）",
    )
    parser.add_argument(
        "--poppler-path",
        dest="poppler_path",
        default=None,
        help="可选，Poppler 可执行文件目录（Windows 常用）",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    input_pdf = os.path.abspath(args.input_pdf)
    output_docx = os.path.abspath(args.output_docx)

    print(f"输入 PDF: {input_pdf}")
    print(f"输出 Word: {output_docx}")
    print(f"模式: {args.mode}")
    print(f"字体: {args.font_name}，字号: {args.font_size}pt")
    if args.mode == "ocr":
        # 根据引擎设置默认语言
        if args.lang is None:
            if args.engine == "paddle":
                args.lang = "ch"  # PaddleOCR 中文
            else:
                args.lang = "chi_sim"  # Tesseract 简体中文
        print(f"OCR 引擎: {args.engine}")
        print(f"语言/模型: {args.lang}")
        if args.workers:
            print(f"并发线程数: {args.workers}")
        print(f"是否格式化: {'是' if args.auto_format else '否'}")
        print("开始 OCR 转换（可能需要一段时间，请耐心等待）...")
    elif args.mode == "text":
        if args.remove_tokens:
            print(f"去除干扰词: {args.remove_tokens}")
        print(f"是否格式化: {'是' if args.auto_format else '否'}")
        print("开始文本层提取（不经 OCR，速度更快）...")
    else:
        print("使用 macOS 原生 Vision OCR（依赖 PyMuPDF + ocrmac）")
        print(f"是否格式化: {'是' if args.auto_format else '否'}")

    pbar: Optional[tqdm] = None

    def progress_callback(current: int, total: int) -> None:
        nonlocal pbar
        if pbar is None:
            pbar = tqdm(total=total, unit="page", desc="OCR 进度")
        # 每调用一次代表完成一页，因此按 1 更新
        pbar.update(1)

    try:
        if args.mode == "ocr":
            result_path = convert_pdf_to_docx(
                pdf_path=input_pdf,
                output_path=output_docx,
                lang=args.lang,
                dpi=args.dpi,
                poppler_path=args.poppler_path,
                ocr_engine=args.engine,
                max_workers=args.workers,
                progress_callback=progress_callback,
                auto_format=args.auto_format,
                font_name=args.font_name,
                font_size_pt=args.font_size,
            )
        elif args.mode == "text":
            # text 模式：直接提取文本层，不用 OCR
            from converter import extract_pdf_text_to_docx

            result_path = extract_pdf_text_to_docx(
                pdf_path=input_pdf,
                output_path=output_docx,
                remove_tokens=args.remove_tokens,
                progress_callback=progress_callback,
                auto_format=args.auto_format,
                font_name=args.font_name,
                font_size_pt=args.font_size,
            )
        else:
            # mac 模式：macOS 原生 Vision OCR
            from converter import convert_pdf_to_docx_mac_vision

            result_path = convert_pdf_to_docx_mac_vision(
                pdf_path=input_pdf,
                output_path=output_docx,
                dpi=args.dpi,
                progress_callback=progress_callback,
                auto_format=args.auto_format,
                font_name=args.font_name,
                font_size_pt=args.font_size,
            )
    except Exception as e:
        if pbar is not None:
            pbar.close()
        print(f"转换失败: {e}")
        raise SystemExit(1)

    if pbar is not None:
        pbar.close()

    print(f"转换完成，文件已保存到：{result_path}")


if __name__ == "__main__":
    main()


