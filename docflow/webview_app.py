"""DocTools —— Web 壳启动器（pywebview）。

界面是 web/index.html（原生 HTML/CSS/JS），通过 pywebview 包成本地桌面窗口。
Python 侧只负责三件事：
  1. 调本机文件/目录选择对话框（tkinter.filedialog）
  2. 探测转换引擎是否就绪（LibreOffice / pdf2docx）
  3. 逐文件调用 converters.convert 并把进度回报给前端（evaluate_js）

前端通过 window.pywebview.api.<方法> 调用这里暴露的函数。
"""

import os
import sys
import json
import platform
import tkinter as tk
from tkinter import filedialog
import webview
import converters
from logging_config import get_logger
logger = get_logger("webview_app")


# tkinter 的 Tk 根必须建/销在「同一个线程」里（pywebview 每个 API 调用可能跑在不同线程），
# 所以每次打开对话框都在当前调用线程内临时建一个隐藏 Tk 根，用完即销毁，避免跨线程访问 Tcl。
def _open_native_dialog(kind, filetypes=None):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)  # 盖在 pywebview 窗口之上
    try:
        if kind == "files":
            paths = filedialog.askopenfilenames(title="选择文件", filetypes=filetypes, parent=root)
            return list(paths)
        else:
            d = filedialog.askdirectory(title="选择输出目录", parent=root)
            return d or None
    finally:
        try:
            root.destroy()
        except Exception:
            pass


class Api:
    """暴露给前端 JavaScript 的接口。"""

    # ---------- 前端日志回传 ----------
    def log(self, msg):
        logger.debug("[frontend] %s", msg)

    # ---------- 选择对话框 ----------
    def select_files(self, tool_id):
        if tool_id == "word2pdf":
            filetypes = [("Word 文档", "*.docx *.doc"), ("所有文件", "*.*")]
        else:
            filetypes = [("PDF 文件", "*.pdf"), ("所有文件", "*.*")]
        logger.debug("select_files tool_id=%s filetypes=%s", tool_id, filetypes)
        try:
            result = _open_native_dialog("files", filetypes)
            logger.info("select_files 选择 %d 个文件", len(result))
            return result
        except Exception as e:
            logger.warning("tkinter 文件对话框异常，退回 pywebview: %s", e)
            # 兜底：退回 pywebview 自带对话框（部分环境可用）
            try:
                return list(WIN.create_file_dialog(
                    webview.OPEN_DIALOG, allow_multiple=True,
                    file_types=(("文件", "*.*"),))) or []
            except Exception as e2:
                logger.error("文件对话框全部失败: %s", e2)
                return []

    def select_dir(self):
        try:
            d = _open_native_dialog("dir")
            logger.info("select_dir 选择目录: %s", d)
            return d
        except Exception as e:
            logger.warning("tkinter 目录对话框异常，退回 pywebview: %s", e)
            try:
                r = WIN.create_file_dialog(webview.FOLDER_DIALOG)
                return r[0] if r else None
            except Exception as e2:
                logger.error("目录对话框全部失败: %s", e2)
                return None

    # ---------- 引擎探测 ----------
    def check_engine(self):
        lo = converters.find_libreoffice()
        exactdoc_ok = ocrmypdf_ok = False
        try:
            import exactdoc  # noqa: F401
            exactdoc_ok = True
        except Exception as e:
            logger.warning("exactdoc 不可用: %s", e)
        try:
            import ocrmypdf  # noqa: F401
            ocrmypdf_ok = True
        except Exception as e:
            logger.warning("ocrmypdf 不可用: %s", e)
        tesseract_ok = shutil.which("tesseract") is not None
        logger.info("引擎探测: libreoffice=%s, exactdoc=%s, ocrmypdf=%s, tesseract=%s",
                    lo, exactdoc_ok, ocrmypdf_ok, tesseract_ok)
        return {
            "libreoffice": lo,
            "exactdoc": exactdoc_ok,
            "ocrmypdf": ocrmypdf_ok,
            "tesseract": tesseract_ok,
        }

    # ---------- 转换（带进度回报）----------
    def convert(self, payload):
        files = payload.get("files", [])
        out_dir = payload.get("out_dir") or "./output"
        tool_id = payload.get("tool_id")
        conflict = payload.get("conflict", "skip")
        options = payload.get("options") or {}
        total = len(files)
        logger.info("convert 请求: tool_id=%s, 文件数=%d, out_dir=%s, conflict=%s, options=%s",
                    tool_id, total, out_dir, conflict, options)
        ok = 0
        for i, f in enumerate(files, 1):
            name = os.path.basename(f)
            try:
                out_path, status = converters.convert(tool_id, f, out_dir, conflict, options)
                if status == "done":
                    ok += 1
                logger.info("  [%d/%d] %s -> %s (%s)", i, total, name, out_path, status)
                self._js("onTask", {
                    "i": i, "total": total, "name": name,
                    "status": status, "out": out_path,
                })
            except Exception as e:  # 单文件失败不影响其余
                logger.error("  [%d/%d] %s 失败: %s", i, total, name, e)
                self._js("onTask", {
                    "i": i, "total": total, "name": name,
                    "status": "error", "out": str(e),
                })
        logger.info("convert 结束: %d/%d 成功", ok, total)
        self._js("onAllDone", {"total": total, "ok": ok})
        return {"total": total, "ok": ok}

    # ---------- 内部：把进度推给前端 ----------
    @staticmethod
    def _js(func, data):
        try:
            if WIN is not None:
                WIN.evaluate_js(f"{func}({json.dumps(data, ensure_ascii=False)})")
        except Exception as e:
            logger.warning("evaluate_js 失败 %s: %s", func, e)


WIN = None


def main():
    global WIN
    here = os.path.dirname(os.path.abspath(__file__))
    html = os.path.join(here, "web", "index.html")
    logger.info("==== DocTools 启动 ====")
    logger.info("Python: %s", sys.version.replace("\n", " "))
    logger.info("Platform: %s", platform.platform())
    logger.info("工作目录: %s", here)
    try:
        import pdf2docx as _p
        logger.info("pdf2docx: %s", getattr(_p, "__version__", "?"))
    except Exception as e:
        logger.warning("pdf2docx 导入失败: %s", e)
    try:
        import pymupdf as _m
        logger.info("pymupdf: %s", getattr(_m, "__version__", "?"))
    except Exception as e:
        logger.warning("pymupdf 导入失败: %s", e)
    try:
        import cv2 as _cv
        logger.info("opencv: %s", _cv.__version__)
    except Exception as e:
        logger.warning("opencv 导入失败: %s", e)
    try:
        import exactdoc as _e
        logger.info("exactdoc: %s", getattr(_e, "__version__", "?"))
    except Exception as e:
        logger.warning("exactdoc 导入失败: %s", e)
    try:
        import ocrmypdf as _o
        logger.info("ocrmypdf: %s", getattr(_o, "__version__", "?"))
    except Exception as e:
        logger.warning("ocrmypdf 导入失败: %s", e)
    logger.info("LibreOffice: %s", converters.find_libreoffice())
    api = Api()
    WIN = webview.create_window(
        "DocTools — 文档格式转换工具箱",
        html,
        js_api=api,
        width=1000,
        height=640,
        min_size=(780, 500),
    )
    webview.start(debug=False)


if __name__ == "__main__":
    main()
