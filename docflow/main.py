"""DocTools — 文档格式转换工具箱 (Flet 界面)。

导航结构（两级）:
  Word 处理  -> 转 PDF (已实现) / 转图片 (占位)
  PDF  处理  -> 转 Word (已实现) / 转图片 (占位)

新增一个工具: 在 TOOLS 里加一条（含 icon），并实现 converters.convert 的分支即可。
"""

import flet as ft
from pathlib import Path

from converters import convert, find_libreoffice

BLUE = "#378ADD"
ORANGE = "#D85A30"
GREY = "#8A8A84"
SIDEBAR_BG = "#26262B"
PANEL_BG = "#FFFFFF"
PAGE_BG = "#ECEAE4"
INK = "#26262B"

# 工具注册表：加功能只需在此追加条目
TOOLS = [
    {"id": "word2pdf", "cat": "Word 处理", "label": "转 PDF", "title": "Word → PDF",
     "sub": "把 Word 文档批量转成 PDF", "engine": "LibreOffice 无头", "color": BLUE,
     "icon": ft.Icons.PICTURE_AS_PDF, "soon": False},
    {"id": "word2img", "cat": "Word 处理", "label": "转图片", "title": "Word → 图片",
     "sub": "把 Word 每页渲染成图片", "engine": "—", "color": GREY,
     "icon": ft.Icons.IMAGE, "soon": True},
    {"id": "pdf2word", "cat": "PDF 处理", "label": "转 Word", "title": "PDF → Word",
     "sub": "把 PDF 转成可编辑 Word", "engine": "pdf2docx", "color": ORANGE,
     "icon": ft.Icons.DESCRIPTION, "soon": False},
    {"id": "pdf2img", "cat": "PDF 处理", "label": "转图片", "title": "PDF → 图片",
     "sub": "把 PDF 每页渲染成图片", "engine": "—", "color": GREY,
     "icon": ft.Icons.IMAGE, "soon": True},
]


def _card_shadow():
    return ft.BoxShadow(blur_radius=10, color=ft.Colors.with_opacity(0.06, ft.Colors.BLACK),
                         offset=ft.Offset(0, 2))


class App:
    def __init__(self, page: ft.Page):
        self.page = page
        self.current = "word2pdf"
        self.files = []

        # 文件 / 目录选择器（flet 1.0 中 FilePicker 是 Service，构造时自动注册到 page；
        # 切勿加入 page.overlay，否则前端会把它当可视化控件渲染并报 "Unknown control: FilePicker"）
        self.pick_files = ft.FilePicker()
        self.pick_dir = ft.FilePicker()

        page.title = "DocTools — 文档格式转换工具箱"
        page.window.width = 920
        page.window.height = 580
        page.window.min_width = 740
        page.window.min_height = 480
        page.bgcolor = PAGE_BG
        page.padding = 0

        # 右侧内容容器
        self.content = ft.Column(expand=True, scroll=ft.ScrollMode.AUTO, spacing=0)
        self.content_wrap = ft.Container(
            content=self.content, bgcolor=PANEL_BG, expand=True,
            padding=ft.Padding.symmetric(horizontal=28, vertical=24))

        self.engine_text = ft.Text("", size=11, color=ft.Colors.GREY_500)
        self.sidebar = ft.Container(width=212, bgcolor=SIDEBAR_BG, content=self._build_sidebar())

        status = ft.Container(
            content=ft.Row(
                [ft.Text("本地运行 · 文件不上传", size=11, color=ft.Colors.GREY_500),
                 self.engine_text],
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            padding=ft.Padding.symmetric(horizontal=18, vertical=10),
            border=ft.Border.only(top=ft.BorderSide(0.5, ft.Colors.GREY_300)),
            bgcolor=PANEL_BG)

        body = ft.Row(expand=True, controls=[self.sidebar, ft.VerticalDivider(width=1), self.content_wrap])
        page.add(ft.Column([body, status], expand=True, spacing=0))

        self._select("word2pdf")

    # ---------- 侧栏 ----------
    def _build_sidebar(self):
        grouped = {}
        for t in TOOLS:
            grouped.setdefault(t["cat"], []).append(t)

        controls = [
            ft.Container(
                content=ft.Row([
                    ft.Container(width=32, height=32, border_radius=9, bgcolor="#534AB7",
                                 content=ft.Icon(ft.Icons.DOCUMENT_SCANNER, color=ft.Colors.WHITE, size=18),
                                 alignment=ft.Alignment(0, 0)),
                    ft.Text("DocTools", size=15, weight=ft.FontWeight.W_500, color=ft.Colors.WHITE),
                ], spacing=10),
                padding=ft.Padding.only(left=14, top=18, bottom=20, right=14)),
        ]

        for cat, items in grouped.items():
            controls.append(ft.Container(
                ft.Text(cat.upper(), size=10, color=ft.Colors.GREY_500, weight=ft.FontWeight.W_500),
                padding=ft.Padding.only(left=16, top=16, bottom=6)))
            for it in items:
                controls.append(self._nav_item(it))

        controls.append(ft.Container(expand=True))
        return ft.Column(controls, spacing=2, scroll=ft.ScrollMode.AUTO)

    def _nav_item(self, t):
        active = (t["id"] == self.current)
        color = t["color"]
        return ft.Container(
            content=ft.Row([
                ft.Container(width=3, height=20, border_radius=2,
                             bgcolor=color if active else ft.Colors.TRANSPARENT),
                ft.Icon(t["icon"], size=18, color=color if active else ft.Colors.GREY_500),
                ft.Text(t["label"], size=13,
                        color=ft.Colors.WHITE if active else ft.Colors.GREY_400,
                        weight=ft.FontWeight.W_500 if active else ft.FontWeight.NORMAL),
                ft.Container(
                    ft.Text("soon", size=10, color=ft.Colors.GREY_500),
                    bgcolor=ft.Colors.with_opacity(0.10, ft.Colors.GREY_500),
                    border_radius=4, padding=ft.Padding.symmetric(horizontal=6, vertical=1),
                    visible=t["soon"]),
            ], spacing=10),
            bgcolor=ft.Colors.with_opacity(0.10, ft.Colors.WHITE) if active else None,
            border_radius=9,
            padding=ft.Padding.symmetric(horizontal=13, vertical=7),
            ink=True,
            on_click=lambda e, tid=t["id"]: self._select(tid),
            tooltip=None if not t["soon"] else "即将推出")

    # ---------- 切换工具 ----------
    def _select(self, tool_id):
        self.current = tool_id
        t = next(x for x in TOOLS if x["id"] == tool_id)
        self.engine_text.value = f"引擎：{t['engine']}"
        self.sidebar.content = self._build_sidebar()
        self._build_panel(t)
        self.sidebar.update()

    # ---------- 右侧面板 ----------
    def _build_panel(self, t):
        self.content.controls.clear()
        color = t["color"]

        # 标题行
        icon = ft.Container(width=46, height=46, border_radius=12, bgcolor=color,
                            content=ft.Icon(t["icon"], color=ft.Colors.WHITE, size=24),
                            alignment=ft.Alignment(0, 0), shadow=_card_shadow())
        title_col = ft.Column(
            [ft.Text(t["title"], size=16, weight=ft.FontWeight.W_500, color=INK),
             ft.Text(t["sub"], size=12, color=ft.Colors.GREY_500)], spacing=3)
        engine_chip = ft.Container(
            content=ft.Text(f"引擎 · {t['engine']}", size=11, color=color),
            bgcolor=ft.Colors.with_opacity(0.10, color),
            border_radius=20, padding=ft.Padding.symmetric(horizontal=10, vertical=5))
        header = ft.Row([icon, title_col, ft.Container(expand=True), engine_chip],
                        alignment=ft.CrossAxisAlignment.CENTER, spacing=14)
        self.content.controls.append(ft.Container(header, padding=ft.Padding.only(top=4, bottom=22)))

        if t["soon"]:
            self.content.controls.append(ft.Container(
                content=ft.Column(
                    [ft.Icon(t["icon"], size=34, color=ft.Colors.GREY_400),
                     ft.Text(t["title"], size=15, weight=ft.FontWeight.W_500, color=INK),
                     ft.Text("正在开发中，敬请期待", size=13, color=ft.Colors.GREY_500)],
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=12),
                border=ft.Border.all(1, ft.Colors.GREY_300), border_radius=14,
                padding=64, alignment=ft.Alignment(0, 0), bgcolor="#FAFAFA"))
            self.page.update()
            return

        self._build_work_panel(t)
        self.page.update()

    def _build_work_panel(self, t):
        color = t["color"]
        self.file_label = ft.Text("未选择文件", size=12, color=ft.Colors.GREY_500)

        async def on_drop_click(e):
            await self._pick(t)

        drop = ft.Container(
            content=ft.Column(
                [ft.Icon(ft.Icons.UPLOAD_FILE_ROUNDED, size=30, color=color),
                 ft.Text("点击选择文件 / 文件夹", size=14, weight=ft.FontWeight.W_500, color=INK),
                 self.file_label],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=8),
            border=ft.Border.all(1.2, ft.Colors.GREY_300), border_radius=12,
            padding=34, alignment=ft.Alignment(0, 0), bgcolor="#FAFAFA", ink=True,
            on_click=on_drop_click, shadow=_card_shadow())
        self.content.controls.append(drop)

        # 选项
        self.out_dir = ft.TextField(value="./output", label="输出目录", expand=True,
                                    border=ft.OutlineInputBorder(side=ft.BorderSide(color=ft.Colors.GREY_300)))
        self.conflict = ft.Dropdown(
            label="重名处理", width=180,
            options=[ft.dropdown.Option("skip", "跳过已存在"),
                     ft.dropdown.Option("force", "覆盖"),
                     ft.dropdown.Option("ask", "询问")],
            value="skip", border=ft.OutlineInputBorder(side=ft.BorderSide(color=ft.Colors.GREY_300)))
        self.content.controls.append(ft.Container(
            ft.Row(
                [self.out_dir,
                 ft.IconButton(ft.Icons.FOLDER_OPEN, tooltip="选择输出目录", on_click=self._on_dir_click),
                 self.conflict], alignment=ft.MainAxisAlignment.START, spacing=10),
            margin=ft.Margin.only(top=18)))

        # 转换按钮
        self.convert_btn = ft.FilledButton("开始转换", bgcolor=color, color=ft.Colors.WHITE,
                                            on_click=self._on_convert,
                                            style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=10)))
        self.content.controls.append(ft.Container(self.convert_btn, margin=ft.Margin.only(top=18)))

        # 任务列表
        self.task_list = ft.Column(spacing=6)
        self.progress_text = ft.Text("0 / 0 完成", size=12, color=ft.Colors.GREY_500)
        self.content.controls.append(ft.Container(
            ft.Column(
                [ft.Row([ft.Text("任务列表", size=12, color=ft.Colors.GREY_500),
                         self.progress_text],
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                 self.task_list],
                spacing=8),
            border=ft.Border.all(0.5, ft.Colors.GREY_300), border_radius=12, padding=14,
            margin=ft.Margin.only(top=18), shadow=_card_shadow()))

    # ---------- 交互 ----------
    async def _pick(self, t):
        exts = ["docx", "doc"] if t["id"] == "word2pdf" else ["pdf"]
        try:
            files = await self.pick_files.pick_files(allow_multiple=True, allowed_extensions=exts)
        except Exception:
            self.page.snack_bar = ft.SnackBar(ft.Text("无法打开文件选择框，请重试"))
            self.page.snack_bar.open = True
            self.page.update()
            return
        if files:
            self.files = [f.path for f in files]
            self.file_label.value = f"已选择 {len(self.files)} 个文件"
        else:
            self.files = []
            self.file_label.value = "未选择文件"
        self.file_label.update()

    async def _on_dir_click(self, e):
        try:
            path = await self.pick_dir.get_directory_path()
        except Exception:
            self.page.snack_bar = ft.SnackBar(ft.Text("无法打开目录选择框，请重试"))
            self.page.snack_bar.open = True
            self.page.update()
            return
        if path:
            self.out_dir.value = path
            self.out_dir.update()

    def _on_convert(self, e):
        if not self.files:
            self.page.snack_bar = ft.SnackBar(ft.Text("请先选择文件"))
            self.page.snack_bar.open = True
            self.page.update()
            return
        if self.current == "word2pdf" and not find_libreoffice():
            self.page.snack_bar = ft.SnackBar(ft.Text("未检测到 LibreOffice，请先安装"))
            self.page.snack_bar.open = True
            self.page.update()
            return

        conflict = self.conflict.value
        out_dir = self.out_dir.value or "./output"
        tool_id = self.current
        files = list(self.files)
        self.task_list.controls.clear()
        self.convert_btn.disabled = True
        self.convert_btn.update()

        def worker():
            done = 0
            for f in files:
                name = Path(f).name
                row = ft.Row([ft.Text(name, size=13), ft.Text("", size=12)])
                self.task_list.controls.append(row)
                try:
                    _, status = convert(tool_id, f, out_dir, conflict)
                    row.controls[1].value = "完成 ✓" if status == "done" else "已跳过"
                    row.controls[1].color = ft.Colors.GREEN if status == "done" else ft.Colors.GREY
                except Exception as ex:
                    row.controls[1].value = "失败"
                    row.controls[1].color = ft.Colors.RED
                    row.controls.append(ft.Text(str(ex)[:60], size=11, color=ft.Colors.GREY,
                                                tooltip=str(ex)))
                done += 1
                self.progress_text.value = f"{done} / {len(files)} 完成"
                self.page.update()
            self.convert_btn.disabled = False
            self.convert_btn.update()

        self.page.run_thread(worker)


def main(page: ft.Page):
    App(page)


if __name__ == "__main__":
    ft.run(main)
