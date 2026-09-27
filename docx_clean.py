#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx-junk-cleaner —— 清理 .docx 里的隐形垃圾与非标准标记。

清理内容
--------
1. U+FEFF (ZERO WIDTH NO-BREAK SPACE / BOM 残留)
   从网页、AI 回答、PDF 复制文本时最常见的残留；不可见，但会影响 Word 的断行。
2. U+200B–U+200D、U+2060、U+180E (零宽空格 / 零宽连接符 / word joiner / 蒙古文元音分隔符)
3. ``<w:cr/>`` -> ``<w:br/>``
   ``<w:cr/>`` 是 WPS / 旧版 WordprocessingML 的软回车；Word 自己的写法是 ``<w:br/>``。

只读原文件、只写新文件（``--in-place`` 除外，且会留 ``.bak``）；文档正文一个字都不改。

界面语言会跟随控制台：能显示中文就用中文，否则自动切英文
（也可用环境变量 ``DOCX_CLEAN_LANG=zh|en`` 强制）。

用法
----
    python docx_clean.py 文档.docx                 # 生成 文档_clean.docx
    python docx_clean.py 文档.docx -o out.docx     # 指定输出
    python docx_clean.py 文档.docx --dry-run       # 只报告，不写文件
    python docx_clean.py *.docx --in-place         # 就地清理（留 .bak）
    python docx_clean.py 文档.docx --json          # 机器可读输出
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import zipfile
from pathlib import Path

__version__ = "1.0.1"

FEFF = "\ufeff"
ZERO_WIDTH_RE = re.compile("[\u200b-\u200d\u2060\u180e]")
CR_RE = re.compile(r"<w:cr\s*/>")
# 只处理 word/ 下的 XML 部件（正文、页眉页脚、脚注尾注、批注、样式、主题、设置等）
XML_PART_RE = re.compile(r"^word/.*\.xml$", re.IGNORECASE)


# --------------------------------------------------------------------------
# 控制台：语言选择 + 永不因编码而崩溃
# --------------------------------------------------------------------------

def _detect_lang() -> str:
    forced = os.environ.get("DOCX_CLEAN_LANG", "").strip().lower()
    if forced in ("zh", "en"):
        return forced
    encoding = (getattr(sys.stdout, "encoding", None) or "ascii")
    try:
        "零".encode(encoding, "strict")
        return "zh"
    except Exception:
        return "en"


LANG = _detect_lang()


def t(zh: str, en: str) -> str:
    """按界面语言取词。"""
    return zh if LANG == "zh" else en


def _make_output_safe() -> None:
    """Windows 控制台可能是 cp1252/cp936：输出中文时不要抛 UnicodeEncodeError。

    保留控制台原有编码（中文 Windows 上显示正常），只是不再因为个别字符编码失败而崩掉。
    """
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass


# --------------------------------------------------------------------------
# 核心逻辑
# --------------------------------------------------------------------------

class Stats:
    """一次清理的统计结果。"""

    def __init__(self, src: Path) -> None:
        self.src = src
        self.feff = 0
        self.zero_width = 0
        self.cr = 0
        self.parts_touched: list[str] = []
        self.parts_scanned = 0
        self.skipped: list[str] = []
        self.output: Path | None = None

    @property
    def junk(self) -> int:
        return self.feff + self.zero_width + self.cr

    def as_dict(self) -> dict:
        return {
            "source": str(self.src),
            "output": str(self.output) if self.output else None,
            "feff": self.feff,
            "zero_width": self.zero_width,
            "cr": self.cr,
            "junk_total": self.junk,
            "parts_touched": self.parts_touched,
            "parts_scanned": self.parts_scanned,
            "skipped_parts": self.skipped,
        }

    def describe(self) -> str:
        lines = [
            t(
                f"U+FEFF={self.feff}  零宽字符={self.zero_width}  软回车<w:cr/>={self.cr}",
                f"U+FEFF={self.feff}  zero-width={self.zero_width}  <w:cr/> soft returns={self.cr}",
            )
        ]
        for name in self.parts_touched:
            lines.append(f"      {name}")
        if not self.junk:
            lines.append(t("      未发现需要清理的标记（文件本身是干净的）",
                           "      nothing to clean (file is already tidy)"))
        if self.skipped:
            lines.append(
                t(f"      跳过（非 UTF-8 文本部件）: {', '.join(self.skipped)}",
                  f"      skipped (not UTF-8 text): {', '.join(self.skipped)}")
            )
        return "\n".join(lines)


def _clean_text(text: str) -> tuple[str, int, int, int]:
    """返回 (清理后的文本, FEFF 数, 零宽字符数, 软回车数)。"""
    feff = text.count(FEFF)
    zero_width = len(ZERO_WIDTH_RE.findall(text))
    cr = len(CR_RE.findall(text))
    if feff:
        text = text.replace(FEFF, "")
    if zero_width:
        text = ZERO_WIDTH_RE.sub("", text)
    if cr:
        text = CR_RE.sub("<w:br/>", text)
    return text, feff, zero_width, cr


def clean_docx(
    src: Path,
    dst: Path | None = None,
    *,
    dry_run: bool = False,
    in_place: bool = False,
) -> Stats:
    """清理 ``src``，写出到 ``dst``（默认 ``<名字>_clean.docx``）。"""
    src = Path(src)
    if not src.is_file():
        raise FileNotFoundError(t(f"找不到文件: {src}", f"file not found: {src}"))
    if src.suffix.lower() != ".docx":
        raise ValueError(t(f"只支持 .docx: {src}", f"only .docx is supported: {src}"))

    explicit_output = dst is not None
    if in_place:
        dst = src
    elif dst is None:
        dst = src.with_name(f"{src.stem}_clean{src.suffix}")

    stats = Stats(src)
    stats.output = None if dry_run else Path(dst)

    with zipfile.ZipFile(src, "r") as zin:
        members: list[tuple[zipfile.ZipInfo, bytes]] = []
        for info in zin.infolist():
            data = zin.read(info.filename)
            if XML_PART_RE.match(info.filename):
                stats.parts_scanned += 1
                try:
                    text = data.decode("utf-8")
                except UnicodeDecodeError:
                    stats.skipped.append(info.filename)
                else:
                    new_text, feff, zw, cr = _clean_text(text)
                    if feff or zw or cr:
                        stats.feff += feff
                        stats.zero_width += zw
                        stats.cr += cr
                        stats.parts_touched.append(
                            t(f"{info.filename}: FEFF={feff} 零宽={zw} 软回车={cr}",
                              f"{info.filename}: FEFF={feff} zero-width={zw} <w:cr/>={cr}")
                        )
                        data = new_text.encode("utf-8")
            members.append((info, data))

    if dry_run:
        return stats
    # 文件本来就干净时：默认路径不产出副本，--in-place 直接跳过；
    # 只有用户显式指定 -o/--output 时才照写（方便统一产出规范副本）。
    if not stats.junk and not (explicit_output and not in_place):
        stats.output = None
        return stats

    target = Path(dst)
    if in_place:
        backup = target.with_suffix(target.suffix + ".bak")
        shutil.copy2(target, backup)

    tmp = target.with_name(target.name + ".tmp")
    try:
        with zipfile.ZipFile(tmp, "w") as zout:
            for info, data in members:
                zout.writestr(info, data)
        tmp.replace(target)
    except BaseException:
        tmp.unlink(missing_ok=True)
        raise
    return stats


def _iter_inputs(patterns: list[str]) -> list[Path]:
    files: list[Path] = []
    for pattern in patterns:
        p = Path(pattern)
        if p.is_dir():
            files.extend(sorted(p.glob("*.docx")))
        elif any(ch in pattern for ch in "*?["):
            files.extend(sorted(Path().glob(pattern)))
        else:
            files.append(p)
    return files


def main(argv: list[str] | None = None) -> int:
    _make_output_safe()
    parser = argparse.ArgumentParser(
        prog="docx_clean",
        description=t(
            "清理 .docx 里的隐形垃圾（U+FEFF、零宽字符）与非标准软回车 <w:cr/>",
            "Remove invisible junk from .docx (U+FEFF, zero-width chars) and WPS <w:cr/> soft returns",
        ),
        epilog=t(
            "示例: python docx_clean.py 报告.docx --dry-run",
            "example: python docx_clean.py report.docx --dry-run",
        ),
    )
    parser.add_argument("inputs", nargs="+", help=t("要清理的 .docx 文件、目录或通配符",
                                                    "one or more .docx files, directories or globs"))
    parser.add_argument("-o", "--output",
                        help=t("输出路径（仅单个输入时有效）", "output path (single input only)"))
    parser.add_argument("--in-place", action="store_true",
                        help=t("就地改写原文件（会生成 .bak 备份）",
                               "rewrite the source file, keeping a .bak backup"))
    parser.add_argument("--dry-run", action="store_true",
                        help=t("只报告发现的问题，不写文件", "report only, write nothing"))
    parser.add_argument("--json", action="store_true", help="JSON output")
    parser.add_argument("--quiet", action="store_true",
                        help=t("只输出出错信息", "only print errors"))
    parser.add_argument("--version", action="version", version=f"docx-junk-cleaner {__version__}")
    args = parser.parse_args(argv)

    files = _iter_inputs(args.inputs)
    if not files:
        print(t("没有匹配到任何 .docx 文件", "no .docx file matched"), file=sys.stderr)
        return 1
    if args.output and len(files) > 1:
        print(t("-o/--output 只能配合单个输入使用", "-o/--output requires a single input"),
              file=sys.stderr)
        return 2

    results: list[Stats] = []
    failed = 0
    for f in files:
        try:
            stats = clean_docx(
                f,
                Path(args.output) if args.output else None,
                dry_run=args.dry_run,
                in_place=args.in_place,
            )
        except (OSError, ValueError, zipfile.BadZipFile) as exc:
            failed += 1
            results.append(Stats(f))
            if args.json:
                print(json.dumps({"source": str(f), "error": str(exc)}, ensure_ascii=False))
            else:
                print(t(f"[失败] {f}", f"[FAILED] {f}"), file=sys.stderr)
                print(f"      {exc}", file=sys.stderr)
                if isinstance(exc, PermissionError):
                    print(t("      文件可能正被 Word/WPS 打开，请先关闭再试。",
                            "      the file may be open in Word/WPS - close it and retry."),
                          file=sys.stderr)
            continue

        results.append(stats)
        if args.json:
            print(json.dumps(stats.as_dict(), ensure_ascii=False))
        elif not args.quiet:
            if args.dry_run:
                print(t(f"[检查] {f.name}", f"[check] {f.name}"))
            elif stats.junk:
                print(t(f"[完成] {stats.output.name}", f"[done] {stats.output.name}"))
            else:
                print(t(f"[跳过] {f.name}（未发现需要清理的标记）",
                        f"[skip] {f.name} (nothing to clean)"))
                continue
            print(stats.describe())

    if args.json:
        print(
            json.dumps(
                {
                    "files": len(files),
                    "failed": failed,
                    "junk_total": sum(s.junk for s in results),
                },
                ensure_ascii=False,
            )
        )
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
