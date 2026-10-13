"""docflow 全局日志配置。

设计目标（开发调试阶段）：
- 级别 DEBUG，信息量越大越好，便于排查转换/引擎/界面问题。
- 同时写文件 + 控制台。
- 写文件到 logs/app.log（当前活动日志）。
- 跨天轮转：每次进程启动时，若 app.log 已存在且其最后修改日期不是今天，
  把它重命名为 logs/YYYY-MM-DD.log（它那天的日期），再新建当日 app.log。
  这样每天的日志最终都归档成日期文件，当前活动日志始终是 app.log。

用法：
    from logging_config import get_logger
    logger = get_logger("converters")   # 得到 docflow.converters
    logger.debug("转换开始: %s", path)
"""
import os
import logging
from datetime import datetime

_LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
_CURRENT_LOG = os.path.join(_LOG_DIR, "app.log")
_configured = False


def _rollover_if_needed():
    """若 app.log 是旧日期的，重命名为 <日期>.log，避免把多天日志混在一个文件。"""
    if not os.path.exists(_CURRENT_LOG):
        return
    try:
        mtime = os.path.getmtime(_CURRENT_LOG)
        file_date = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d")
        today = datetime.now().strftime("%Y-%m-%d")
        if file_date == today:
            return
        archive = os.path.join(_LOG_DIR, f"{file_date}.log")
        if os.path.exists(archive):
            base, ext = os.path.splitext(archive)
            i = 1
            while os.path.exists(f"{base}_{i}{ext}"):
                i += 1
            archive = f"{base}_{i}{ext}"
        os.rename(_CURRENT_LOG, archive)
    except Exception:
        # 轮转失败绝不影响主流程
        pass


def setup_logging(level: int = logging.DEBUG):
    """配置 docflow 根日志器（幂等，只配置一次）。"""
    global _configured
    if _configured:
        return logging.getLogger("docflow")
    os.makedirs(_LOG_DIR, exist_ok=True)
    _rollover_if_needed()

    fmt = logging.Formatter(
        "%(asctime)s %(levelname)-7s [%(name)s.%(funcName)s:%(lineno)d] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    root = logging.getLogger("docflow")
    root.setLevel(level)
    for h in list(root.handlers):
        root.removeHandler(h)

    fh = logging.FileHandler(_CURRENT_LOG, encoding="utf-8")
    fh.setLevel(level)
    fh.setFormatter(fmt)
    root.addHandler(fh)

    sh = logging.StreamHandler()
    sh.setLevel(level)
    sh.setFormatter(fmt)
    root.addHandler(sh)

    # 第三方库日志也进文件，便于排查（pdf2docx / pymupdf / webview / tkinter）
    for name in ("pdf2docx", "fitz", "pymupdf", "webview", "tkinter"):
        logging.getLogger(name).setLevel(logging.DEBUG)

    _configured = True
    return root


def get_logger(name: str):
    """获取 docflow 子日志器，自动触发配置。"""
    setup_logging()
    if not name.startswith("docflow."):
        name = "docflow." + name
    return logging.getLogger(name)
