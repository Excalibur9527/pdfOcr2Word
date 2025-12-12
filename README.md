# PDF OCR 转 Word 工具

> 一个功能强大的 PDF 转 Word 工具，支持多种 OCR 引擎和文本提取模式，适用于扫描版 PDF 和文字层 PDF。

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

---

## ✨ 特性

- 🚀 **多种处理模式**：OCR 识别、文本层直读、macOS 原生 Vision OCR
- 🎯 **多引擎支持**：Tesseract、PaddleOCR（中文优化）、macOS Vision
- ⚡ **多线程加速**：OCR 模式支持并发处理，大幅提升速度
- 📊 **实时进度条**：所有模式都支持进度显示和预计时间
- 🎨 **智能格式化**：自动合并断行、去除多余空格、统一字体字号
- 🔧 **灵活配置**：丰富的命令行参数，满足不同需求

---

## 📦 快速开始

### 安装依赖

#### 1. Python 依赖

```bash
pip install -r requirements.txt
```

#### 2. 系统依赖（OCR 模式需要）

**macOS (Homebrew):**
```bash
brew install tesseract tesseract-lang  # Tesseract OCR + 中文语言包
brew install poppler                    # PDF 渲染工具
```

**Ubuntu/Debian:**
```bash
sudo apt-get install tesseract-ocr tesseract-ocr-chi-sim
sudo apt-get install poppler-utils
```

#### 3. 可选依赖（按需安装）

```bash
# PaddleOCR（中文识别推荐）
pip install paddlepaddle paddleocr

# macOS Vision OCR（仅 macOS）
pip install ocrmac PyMuPDF
```

---

## 🎯 使用指南

### 基础用法

```bash
# 最简单的用法（使用默认 OCR 模式）
python main.py input.pdf output.docx

# 指定输出文件（自动添加 .docx 后缀）
python main.py book.pdf result
```

### 模式选择

项目支持 **3 种主要模式**，根据你的 PDF 类型选择：

| 模式 | 命令 | 适用场景 | 速度 | 质量 |
|------|------|---------|------|------|
| **OCR 模式** | `--mode ocr` | 扫描版 PDF、图片 PDF | 中等 | ⭐⭐⭐⭐ |
| **文本层模式** | `--mode text` | PDF 有内置文字层 | ⚡ 最快 | ⭐⭐⭐⭐⭐ |
| **macOS Vision** | `--mode mac` | macOS 系统，高质量 OCR | 较慢 | ⭐⭐⭐⭐⭐ |

---

## 📖 详细模式说明

### 1. OCR 模式 (`--mode ocr`，默认)

适用于扫描版 PDF 或图片型 PDF，支持两种 OCR 引擎。

#### 1.1 Tesseract OCR

**特点：** 跨平台、稳定、支持多语言混合

```bash
# 中文识别
python main.py input.pdf output.docx --mode ocr --engine tesseract --lang chi_sim

# 英文识别
python main.py input.pdf output.docx --mode ocr --engine tesseract --lang eng

# 中英混合识别
python main.py input.pdf output.docx --mode ocr --engine tesseract --lang chi_sim+eng

# 多线程加速（推荐 4-8 线程）
python main.py input.pdf output.docx --mode ocr --engine tesseract --lang chi_sim --workers 4
```

#### 1.2 PaddleOCR

**特点：** 中文识别效果更佳，推荐用于中文文档

```bash
# 中文识别（推荐）
python main.py input.pdf output.docx --mode ocr --engine paddle --lang ch

# 英文识别
python main.py input.pdf output.docx --mode ocr --engine paddle --lang en

# 多线程加速
python main.py input.pdf output.docx --mode ocr --engine paddle --lang ch --workers 4
```

**OCR 模式参数：**
- `--engine tesseract|paddle`：选择 OCR 引擎（默认 tesseract）
- `--lang`：语言代码
  - Tesseract: `chi_sim`（中文）、`eng`（英文）、`chi_sim+eng`（中英混合）
  - PaddleOCR: `ch`（中文）、`en`（英文）
- `--dpi 300`：PDF 渲染分辨率（默认 300，越高越清晰但越慢）
- `--workers 4`：并发线程数（默认自动，可手动指定加速）

---

### 2. 文本层直读模式 (`--mode text`)

**最快模式**，直接从 PDF 内置文本层提取文字，无需 OCR。

**适用场景：**
- ✅ PDF 有可选择的文字（不是纯图片）
- ✅ 需要快速提取（秒级完成）
- ✅ 希望保留原始格式

```bash
# 基础用法
python main.py input.pdf output.docx --mode text

# 去除页眉页脚干扰词（可多次使用）
python main.py input.pdf output.docx --mode text \
  --remove-token "人设、流量与成交" \
  --remove-token "王扬名"

# 关闭格式化（保留原始换行）
python main.py input.pdf output.docx --mode text --no-format
```

**文本层模式参数：**
- `--remove-token "关键词"`：可多次使用，删除页眉/页脚等干扰词
- `--no-format`：关闭自动格式化
- `--font-name` / `--font-size`：统一字体字号

---

### 3. macOS 原生 Vision OCR (`--mode mac`)

**macOS 专用**，使用系统自带的 Vision 框架进行 OCR，识别质量最高。

**适用场景：**
- ✅ 在 macOS 系统上运行
- ✅ 需要最高质量的 OCR 识别
- ✅ 可以接受较慢的处理速度

**系统要求：**
- macOS 系统
- 安装依赖：`pip install ocrmac PyMuPDF`

```bash
# 基础用法（自动识别中英文）
python main.py input.pdf output.docx --mode mac

# 自定义字体字号
python main.py input.pdf output.docx --mode mac \
  --font-name "PingFang SC" \
  --font-size 12

# 调整 DPI（默认 300）
python main.py input.pdf output.docx --mode mac --dpi 300
```

---

## 🎨 通用功能

### 统一字体字号

所有模式都支持统一设置输出文档的字体和字号。

```bash
# 使用苹方字体，11 磅
python main.py input.pdf output.docx --font-name "PingFang SC" --font-size 11

# 使用微软雅黑，12 磅
python main.py input.pdf output.docx --font-name "Microsoft YaHei" --font-size 12
```

**默认值：**
- 字体：宋体 (SimSun)
- 字号：12 磅

### 自动格式化（默认开启）

自动格式化功能包括：
- 去除中文之间的多余空格（如 "人 设" → "人设"）
- 智能合并被拆开的句子（根据标点符号判断）
- 按段落分行，提升可读性

**关闭格式化：**
```bash
python main.py input.pdf output.docx --no-format
```

### 进度条显示

所有模式都支持实时进度条，显示：
- 已处理页数 / 总页数
- 处理速度（页/秒）
- 预计剩余时间

---

## 📊 模式选择建议

| 你的 PDF 类型 | 推荐模式 | 命令示例 |
|-------------|---------|---------|
| **有文字层**（可复制粘贴） | `--mode text` | `python main.py book.pdf out.docx --mode text` |
| **扫描版中文** | `--mode ocr --engine paddle` | `python main.py scan.pdf out.docx --mode ocr --engine paddle --lang ch --workers 4` |
| **扫描版英文** | `--mode ocr --engine tesseract` | `python main.py scan.pdf out.docx --mode ocr --engine tesseract --lang eng` |
| **macOS 用户，追求质量** | `--mode mac` | `python main.py scan.pdf out.docx --mode mac` |
| **不确定类型** | 先试 `--mode text`，不行再用 `--mode ocr` | 从快到慢尝试 |

---

## 🚀 性能参考

以 100 页 PDF 为例（仅供参考）：

| 模式 | 预计耗时 | CPU 占用 | 推荐场景 |
|------|---------|---------|---------|
| `--mode text` | **5-10 秒** | 低 | 有文字层的 PDF |
| `--mode ocr --engine tesseract` | 2-5 分钟 | 中-高 | 扫描版，多语言 |
| `--mode ocr --engine paddle` | 3-8 分钟 | 高 | 扫描版中文 |
| `--mode mac` | 5-15 分钟 | 中 | macOS，高质量需求 |

**加速技巧：**
- OCR 模式使用 `--workers 4` 可显著提速（多核 CPU）
- 降低 `--dpi` 到 200 可提速，但识别率略降
- 文本层模式最快，优先尝试

---

## 📝 完整命令示例

```bash
# 场景1：快速提取有文字层的 PDF
python main.py book.pdf output.docx --mode text

# 场景2：扫描版中文 PDF，追求质量，多线程加速
python main.py scan.pdf output.docx \
  --mode ocr \
  --engine paddle \
  --lang ch \
  --workers 4

# 场景3：macOS 用户，高质量 OCR，自定义字体
python main.py scan.pdf output.docx \
  --mode mac \
  --font-name "PingFang SC" \
  --font-size 12

# 场景4：中英混合扫描版
python main.py mixed.pdf output.docx \
  --mode ocr \
  --engine tesseract \
  --lang chi_sim+eng \
  --workers 4

# 场景5：去除干扰词，自定义格式，关闭格式化
python main.py book.pdf output.docx \
  --mode text \
  --remove-token "页眉" \
  --remove-token "页脚" \
  --font-name "Microsoft YaHei" \
  --font-size 11 \
  --no-format
```

---

## 🔧 命令行参数完整列表

### 必需参数

- `input_pdf`：输入 PDF 文件路径
- `output_docx`：输出 Word 文件路径（可不带 .docx 后缀）

### 模式参数

- `--mode {ocr,text,mac}`：处理模式（默认 `ocr`）
  - `ocr`：OCR 识别模式
  - `text`：文本层直读模式
  - `mac`：macOS 原生 Vision OCR

### OCR 模式参数（`--mode ocr`）

- `--engine {tesseract,paddle}`：OCR 引擎（默认 `tesseract`）
- `--lang`：语言代码
  - Tesseract: `chi_sim`、`eng`、`chi_sim+eng`
  - PaddleOCR: `ch`、`en`
- `--dpi 300`：PDF 渲染分辨率（默认 300）
- `--workers`：并发线程数（默认自动）

### 文本层模式参数（`--mode text`）

- `--remove-token "关键词"`：删除干扰词（可多次使用）

### 通用参数

- `--font-name "SimSun"`：字体名称（默认宋体）
- `--font-size 12`：字号（默认 12 磅）
- `--no-format`：关闭自动格式化（默认开启）
- `--poppler-path`：Poppler 可执行文件目录（Windows 常用）

---

## 🏗️ 项目结构

```
pdfOcr2Word/
├── converter/              # 核心转换模块
│   ├── __init__.py
│   └── pdf_ocr_to_word.py  # OCR 和文本提取实现
├── main.py                 # 命令行入口
├── requirements.txt        # Python 依赖
└── README.md              # 项目文档
```

---

## 🔍 技术实现

### 核心流程

1. **PDF 渲染**：使用 `pdf2image` 或 `PyMuPDF` 将 PDF 页面转为图片
2. **OCR 识别**：根据选择的引擎进行文字识别
3. **文本格式化**：自动清理和合并文本
4. **Word 生成**：使用 `python-docx` 生成格式化的 Word 文档

### 支持的 OCR 引擎

- **Tesseract**：开源 OCR 引擎，跨平台支持
- **PaddleOCR**：百度开源，中文识别优化
- **macOS Vision**：Apple 系统原生 OCR，质量最高

---

## 🐛 常见问题

### Q: 识别效果不好怎么办？

A: 尝试以下方法：
1. 提高 DPI：`--dpi 400`（但会更慢）
2. 中文文档使用 PaddleOCR：`--engine paddle --lang ch`
3. macOS 用户尝试原生 OCR：`--mode mac`
4. 检查 PDF 质量，确保图片清晰

### Q: 处理速度太慢？

A: 优化建议：
1. 使用多线程：`--workers 4` 或 `--workers 8`
2. 有文字层的 PDF 优先用 `--mode text`
3. 降低 DPI：`--dpi 200`（会略降质量）

### Q: 如何判断 PDF 是否有文字层？

A: 在 PDF 阅读器中尝试复制文字，如果能复制，说明有文字层，可以使用 `--mode text`。

### Q: macOS Vision OCR 报错？

A: 确保：
1. 在 macOS 系统上运行
2. 已安装依赖：`pip install ocrmac PyMuPDF`
3. 系统版本支持 Vision 框架（macOS 10.13+）

---

## 📄 许可证

MIT License

---

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

---

## 📌 总结

项目目前支持 **3 种模式 × 2 种 OCR 引擎 = 5 种处理方式**：

1. ✅ **文本层直读**（最快，适用于有文字层的 PDF）
2. ✅ **Tesseract OCR**（跨平台，支持多语言）
3. ✅ **PaddleOCR**（中文优化，推荐中文文档）
4. ✅ **macOS Vision OCR**（质量最高，macOS 专用）
5. ✅ **统一字体字号、智能格式化、实时进度条**

按需选择，灵活组合！🎉
