# doc2md - AI Document Converter

> Convert `.doc` / `.docx` / `.pdf` documents to **Markdown + Images**, enabling AI (Cursor, Copilot, etc.) to fully read document content.

将 `.doc` / `.docx` / `.pdf` 需求文档转换为 **Markdown + 图片** 格式，让 AI (Cursor) 能完整阅读文档内容。

## 快速开始

### 1. 环境要求

- **Python 3.10 - 3.13**（推荐 3.13，MinerU 不支持 3.14+）
- Windows 11（GUI 已适配暗色主题）
- LibreOffice（仅 `.doc` 格式需要）

### 2. 安装

**方式一：一键安装（推荐）**

```bash
setup.bat
```

自动创建 Python 3.13 虚拟环境并安装所有依赖。

**方式二：手动安装**

```bash
py -3.13 -m venv .venv313
.venv313\Scripts\activate
pip install -r requirements.txt
```

### 3. 使用方式

**方式一：可视化界面（推荐）**

```bash
start_gui.bat
```

- 支持拖拽文件到窗口
- 支持批量添加文件/文件夹
- PDF 引擎选择（自动/MinerU/Docling/PyMuPDF4LLM）
- 实时进度条 + 转换日志
- 自动使用 `.venv313` 虚拟环境

**方式二：命令行**

```bash
convert.bat 需求文档.pdf                        # 自动选择最佳引擎
convert.bat 需求文档.pdf --engine mineru         # 指定 MinerU 引擎
convert.bat 需求文档.pdf --engine docling         # 指定 Docling 引擎
convert.bat 需求文档.docx -o ./output            # 指定输出目录
convert.bat ./docs/                              # 批量转换
convert.bat 文档.docx --embed-images             # 图片内嵌 base64
```

## PDF 多引擎架构

针对 PDF 转换的内容丢失问题，集成了 3 个引擎，按准确率自动降级：

| 引擎 | 总体准确率 | 表格提取 | 中文支持 | 速度 | 安装 |
|------|-----------|---------|---------|------|------|
| **MinerU** | 0.82 | **0.87** | 最佳 | 6s/页 | `pip install MinerU`（Python ≤3.13） |
| **Docling** | **0.86** | **0.89** | 良好 | 0.7s/页 | `pip install docling` |
| PyMuPDF4LLM | 0.57 | 0.40 | 需补全 | 0.09s/页 | `pip install pymupdf4llm` |

**引擎自动降级**：当选中的引擎运行失败时，自动尝试下一个引擎（MinerU → Docling → PyMuPDF4LLM），确保转换不会因单个引擎故障而中断。

- `auto` 模式自动检测已安装的最佳引擎：MinerU > Docling > PyMuPDF4LLM
- PyMuPDF4LLM 模式额外使用 fitz 底层 API 补全丢失的文本框内容（服务提供方、接口地址等）
- Docling 模式自动合并拆分的代码块

## 技术原理

### PDF 转换 — 基于计算机视觉模型（非大模型）

MinerU 和 Docling 使用的是**专用计算机视觉（CV）小模型**，不是 ChatGPT 那样的大语言模型（LLM）。

**为什么 PDF 需要模型？** PDF 本质上是一个"画面"（坐标 + 绘图指令），不包含结构信息。表格在底层只是线条和文字坐标，没有"这是表格"的标记，所以需要 AI 模型"看图识结构"。

MinerU 的处理流程：
```
PDF 页面（图像）
  → [YOLO 布局检测模型] 识别标题、表格、图片区域
  → [OCR 模型] 识别表格内文字
  → [公式识别模型] 识别数学公式
  → 输出 Markdown
```

### DOCX/DOC 转换 — 纯 Python 库解析（无模型）

`.docx` 本质是 ZIP 压缩包，内含结构化 XML，Python 库直接读取即可，不需要任何 AI 模型：
```
.docx (ZIP)
├── word/document.xml   → mammoth 读取正文
├── word/styles.xml     → python-docx 读取标题层级
├── word/footnotes.xml  → 提取脚注
└── word/media/         → 提取图片
```

### 对比总结

| | PDF | DOCX | DOC |
|---|---|---|---|
| 文件本质 | 画面（坐标+绘图指令） | 结构化 XML | 二进制旧格式 |
| 解析方式 | CV 模型"看图识字" | Python 库直接读 XML | 转成 docx 再解析 |
| 需要模型 | MinerU/Docling 需要 | 不需要 | 不需要 |
| 准确率 | 依赖模型质量 | 几乎 100% | 取决于转 docx 质量 |

## 磁盘占用与隔离性

### Python 库（.venv313 虚拟环境）

所有 Python 依赖包安装在项目目录下的 `.venv313/` 中，**完全隔离，不影响系统全局 Python 环境**。

```
<项目目录>/.venv313/    ← ~1.8 GB（Python 库 + 依赖）
```

删除 `.venv313` 文件夹即可完全卸载所有 Python 依赖，不留任何残余。

### AI 模型缓存（用户目录）

MinerU 和 Docling 的模型文件缓存在用户目录下，**首次运行时自动下载，后续直接使用缓存**：

```
C:\Users\<用户名>\.cache\huggingface\hub\
├── models--opendatalab--PDF-Extract-Kit-1.0\    ← ~1.9 GB（MinerU 布局检测 + OCR 模型）
├── models--docling-project--docling-layout-heron\ ← ~0.2 GB（Docling 布局模型）
└── models--docling-project--docling-models\       ← ~0.3 GB（Docling 表格模型）
```

总计约 **2.4 GB** 模型缓存。如需清理，直接删除 `C:\Users\<用户名>\.cache\huggingface\hub\` 即可（下次运行会重新下载）。

### 进程与内存

- **模型不会常驻后台** — 只在转换 PDF 时加载到内存，转换结束后进程退出，内存自动释放
- 关闭 GUI 或命令行窗口后，不会有任何残留进程
- 不会注册系统服务、启动项或后台任务

### 完整清理方法

如需完全卸载本工具的所有数据：

```bash
# 1. 删除虚拟环境（Python 库）
rd /s /q .venv313

# 2. 删除模型缓存（AI 模型）
rd /s /q C:\Users\%USERNAME%\.cache\huggingface\hub\models--opendatalab--PDF-Extract-Kit-1.0
rd /s /q C:\Users\%USERNAME%\.cache\huggingface\hub\models--docling-project--docling-layout-heron
rd /s /q C:\Users\%USERNAME%\.cache\huggingface\hub\models--docling-project--docling-models

# 3. 删除 Ultralytics 配置（几 KB）
rd /s /q C:\Users\%USERNAME%\AppData\Roaming\Ultralytics
```

## 输出结构

```
output/
├── 需求文档.md              # Markdown（AI 直接阅读）
└── 需求文档_images/          # 提取的图片
    ├── img_001.png
    ├── img_002.png
    └── ...
```

## 转换质量

### DOCX 转换

| 特性 | 说明 |
|------|------|
| 文档目录 | 按 Word 原始层级生成缩进列表 |
| 标题层级 | 精确匹配 Heading1/2/3 → #/##/### |
| 有序列表 | 自动修正编号（1,2,3...） |
| 表格 | 完整保留，修复空表头和合并单元格 |
| 图片位置 | 与原文位置一致 |
| 图片格式 | EMF/WMF/TIFF 自动转 PNG，**转换后自动更新 Markdown 引用** |
| 脚注 | 自动提取并追加到文末，**支持同一脚注多次引用** |
| 异常处理 | 所有异常输出 WARN 日志，便于排查 |
| 垃圾清理 | 自动删除 0KB 空文件、清理临时目录 |

### PDF 转换

| 特性 | 说明 |
|------|------|
| 多引擎支持 | MinerU / Docling / PyMuPDF4LLM 自动降级 |
| 引擎容错 | 引擎失败自动切换下一个，不会中断转换 |
| 表格提取 | MinerU/Docling 准确率 0.87-0.89 |
| 文本框内容 | 服务提供方、接口地址等灰色背景块完整保留 |
| 代码块 | 智能重建代码围栏（JSON/Java 等），自动合并拆分块 |
| 标题层级 | 自动推断并修正层级 |
| 中文空格 | 自动清理 PDF 提取时插入的多余空格 |

## 格式支持

| 格式 | 引擎 | 依赖 |
|------|------|------|
| `.docx` | mammoth + python-docx | 直接支持 |
| `.doc` | COM / doc2docx / LibreOffice → docx | 需安装 Office 或 [LibreOffice](https://www.libreoffice.org/download/download/) |
| `.pdf` | MinerU / Docling / PyMuPDF4LLM | 见上方引擎表 |

## 批处理脚本

| 脚本 | 用途 |
|------|------|
| `setup.bat` | 一键创建虚拟环境 + 安装依赖 |
| `start_gui.bat` | 启动 GUI（自动使用 .venv313） |
| `convert.bat` | 命令行转换（自动使用 .venv313） |

所有脚本自动优先使用 `.venv313` 虚拟环境（Python 3.13），确保 MinerU 等引擎可用。

## AI 阅读方式

转换后 AI 可以：
1. `Read` 工具读取 `.md` → 完整文字、表格、结构
2. `Read` 工具查看 `_images/*.png` → 截图、流程图、UI 设计图
3. 图片索引标注哪些图片 AI 可直接查看

## License

[MIT](LICENSE)
