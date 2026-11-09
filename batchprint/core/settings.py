"""打印相关的数据模型：全局设置、单文档设置、打印任务。

DEVMODE 用 pywin32 的 PyDEVMODE 对象承载（不做持久化，仅在会话内内存中传递）。
单文档设置(DocSettings)中任何字段为 None 表示"继承全局"。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

# 图片排版模式
FIT_CENTER = "fit_center"    # 等比缩放 + 居中（默认）
FIT_FILL = "fill"            # 铺满纸张（可能裁切/变形，按等比铺满并居中裁切）
FIT_STRETCH = "stretch"      # 拉伸铺满（不保持比例）

# 打印任务状态
STATUS_PENDING = "待打印"
STATUS_PRINTING = "打印中"
STATUS_DONE = "完成"
STATUS_FAILED = "失败"
STATUS_SKIPPED = "已跳过"
STATUS_CONVERTING = "转换中"          # Word 文档正在转成 PDF
STATUS_CONVERT_FAILED = "转换失败"


@dataclass
class GlobalSettings:
    """全局打印设置，应用到队列里所有文档（除非文档单独覆盖）。"""

    printer_name: str = ""
    devmode: Optional[Any] = None       # PyDEVMODE，None = 打印机默认
    copies: int = 1
    image_fit: str = FIT_CENTER
    auto_orient: bool = True             # 按内容宽高比自动纵/横向
    print_markup: bool = False           # Word 文档是否打印修订和批注（默认不打印）


@dataclass
class DocSettings:
    """单文档设置，字段为 None 表示继承全局。"""

    devmode: Optional[Any] = None       # None = 继承全局 devmode
    copies: Optional[int] = None        # None = 继承全局份数
    page_range: Optional[str] = None    # None/"" = 全部；如 "1-3,5"（PDF / Word 有意义）
    print_markup: Optional[bool] = None  # None = 继承全局（仅 Word 有意义；界面暂未开放）

    def is_empty(self) -> bool:
        return (self.devmode is None and self.copies is None and not self.page_range
                and self.print_markup is None)


@dataclass
class PrintJob:
    """打印队列中的一项。"""

    path: Path
    kind: str                            # 'pdf' / 'image' / 'word'
    pages: Optional[int] = None          # PDF / Word 页数；图片为 None；Word 转换前为 None
    status: str = STATUS_PENDING
    checked: bool = True
    error: Optional[str] = None
    doc: DocSettings = field(default_factory=DocSettings)
    pdf_path: Optional[Path] = None      # Word 文档转换后的 PDF（缓存里）
    convert_error: Optional[str] = None  # Word 文档转换失败的原因


def type_label(job: "PrintJob") -> str:
    """类型显示文案：图片显示「扩展名 + 图片」（如 JPG 图片），
    其余显示「扩展名 + 文档」（如 PDF 文档、DOCX 文档）。"""
    ext = job.path.suffix.lstrip(".").upper()
    if job.kind == "image":
        return f"{ext} 图片" if ext else "图片"
    return f"{ext} 文档" if ext else "文档"


@dataclass
class EffectiveSettings:
    """把全局设置与单文档覆盖合并后的最终生效设置。"""

    printer_name: str
    devmode: Optional[Any]
    copies: int
    image_fit: str
    page_range: Optional[str]
    auto_orient: bool
    print_markup: bool


def effective_markup(global_s: GlobalSettings, job: PrintJob) -> bool:
    """Word 文档是否打印修订和批注。"""
    if job.doc.print_markup is not None:
        return job.doc.print_markup
    return global_s.print_markup


def resolve(global_s: GlobalSettings, job: PrintJob) -> EffectiveSettings:
    """合并全局设置与单文档覆盖，得到某个任务的最终生效设置。"""
    return EffectiveSettings(
        printer_name=global_s.printer_name,
        devmode=job.doc.devmode if job.doc.devmode is not None else global_s.devmode,
        copies=job.doc.copies if job.doc.copies is not None else global_s.copies,
        image_fit=global_s.image_fit,
        page_range=job.doc.page_range or None,
        auto_orient=global_s.auto_orient,
        print_markup=effective_markup(global_s, job),
    )
