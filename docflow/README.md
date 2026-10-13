# DocTools — 文档格式转换工具箱

本地运行的文档格式互转小工具，界面用原生 HTML/CSS/JS 写在 `web/`，通过
**pywebview** 包成本地桌面窗口（Windows 上用系统自带的 Edge WebView2，无需联网、无需下载浏览器内核）。

当前已实现：

- **Word 处理 → 转 PDF**（LibreOffice 无头）
- **PDF 处理 → 转 Word**（exactdoc 主引擎，带 LibreOffice 渲染闭环校验修正版面偏移；
  扫描件/图片型 PDF 自动经 OCR 前置；pdf2docx 作为兜底）

其余卡片（PDF/Word 转图片、PDF 转 Excel/PPT/合并/压缩/拆分等）为占位项（点击提示开发中）。

## 架构

- `converters.py`：转换核心。
  - `word_to_pdf` 调 LibreOffice 无头（带路径探测与 `-env:UserInstallation` 并发隔离）。
  - `pdf_to_word` 主引擎 exactdoc（直接产出高质量可编辑 docx + 渲染闭环校验）；扫描件先经
    `ocrmypdf` 前置 OCR 叠文字层；exactdoc 失败时回退 `pdf2docx`。内置两处补丁：
    exactdoc 的 CJK 页码崩溃修复、pdf2docx 的特殊色彩空间图片兼容。
- `logging_config.py`：全局日志（DEBUG 级，写 `logs/app.log`，跨天归档为 `logs/YYYY-MM-DD.log`）。
- `webview_app.py`：pywebview 启动器，暴露 `select_files / select_dir / check_engine /
  convert` 给前端 JS，并通过 `evaluate_js` 把逐文件进度推回界面。
- `web/`：现代界面。`index.html` + `style.css` + `app.js`。新增工具只需在
  `app.js` 的 `TOOLS` 里加一项，并实现 `converters.convert` 对应分支。

## 安装与环境

```bash
cd docflow
pip install -r requirements.txt
python webview_app.py
```

需要以下**系统软件**（非 pip 库，需单独安装）：

1. **LibreOffice**（Word→PDF 与 exactdoc 的 refine 闭环都依赖它）
   - Windows: `winget install --id TheDocumentFoundation.LibreOffice`
   - macOS:   `brew install --cask libreoffice`
   - Linux:   `sudo apt install libreoffice`

2. **Tesseract OCR + 中文语言包**（仅处理**扫描件/图片型 PDF** 时需要；纯数字 PDF 不需要）
   - 安装 Tesseract：
     - Windows: `winget install --id UB-Mannheim.TesseractOCR --exact --accept-package-agreements --accept-source-agreements`
     - macOS: `brew install tesseract tesseract-lang`
     - Linux: `sudo apt install tesseract-ocr tesseract-ocr-chi-sim`
   - 装完把 `C:\Program Files\Tesseract-OCR` 加入 PATH（重开终端生效）。
   - 中文语言包 `chi_sim.traineddata` 放到 Tesseract 的 tessdata 目录
     （Windows 默认 `C:\Program Files\Tesseract-OCR\tessdata\`）。
     下载：https://github.com/tesseract-ocr/tessdata_fast/raw/main/chi_sim.traineddata
     （若 GitHub 直连被截断，可走镜像 https://ghproxy.com/https://github.com/tesseract-ocr/tessdata_fast/raw/main/chi_sim.traineddata ）

缺少 LibreOffice 时界面顶部红条提示「未检测到 LibreOffice」；缺少 exactdoc 时提示「未检测到 exactdoc」。
Tesseract/OCR 缺失不影响数字 PDF 转换，仅在转换扫描件时报错提示。

## 打包（单 exe，后续）

```bash
pip install pyinstaller
pyinstaller -F -w webview_app.py
```

LibreOffice 体积大（~300MB），不塞进 exe；运行时若未检测到会提示先安装。

## 已知限制

- PDF→Word 对极复杂混排（无边框表格、SmartArt、公式）仍可能需人工校对；但 exactdoc 的渲染闭环已大幅优于 pdf2docx，且对 CJK 文档（带圈数字页码等）已做兼容。
- 扫描件需 OCR，识别质量取决于原图分辨率（建议 ≥ 200 DPI）。
- Word 艺术字 / SmartArt / 修订标记还原依赖 LibreOffice，可能不完全一致。
