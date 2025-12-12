"""
PDF 转 Word 的核心逻辑包。

对外主要暴露：
- convert_pdf_to_docx(pdf_path: str, output_path: str, ...)
- extract_pdf_text_to_docx(pdf_path: str, output_path: str, ...)
- convert_pdf_to_docx_mac_vision(pdf_path: str, output_path: str, ...)

Author: Excalibur9527
"""

from .pdf_ocr_to_word import (
    convert_pdf_to_docx,
    extract_pdf_text_to_docx,
    convert_pdf_to_docx_mac_vision,
)

__all__ = [
    "convert_pdf_to_docx",
    "extract_pdf_text_to_docx",
    "convert_pdf_to_docx_mac_vision",
]


