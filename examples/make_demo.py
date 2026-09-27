#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成一个"带病"的演示 docx：行尾 U+FEFF + WPS 非标准软回车 <w:cr/>。

案例可复现：
    python examples/make_demo.py                # 生成 demo_with_junk.docx
    python docx_clean.py demo_with_junk.docx --dry-run   # 体检
    python docx_clean.py demo_with_junk.docx             # 清理

生成的文件用 Word 打开通常看不出异常（是否发作取决于本机环境），
但用下面这行就能看到"字都在、只是标记脏"：

    python examples/make_demo.py --show         # 打印 XML 片段
"""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from tests import fixtures  # noqa: E402

DEMO_NAME = "demo_with_junk.docx"
ZERO_WIDTH_RE = re.compile("[\u200b-\u200d\u2060\u180e]")
CR_RE = re.compile(r"<w:cr\s*/>")


def show(document_xml: str) -> None:
    """把 <w:t> 里的隐形字符显形，方便肉眼看见问题。"""
    visible = (
        document_xml.replace("\ufeff", "⟨U+FEFF⟩")
        .replace("\u200b", "⟨U+200B⟩")
        .replace("\u200d", "⟨U+200D⟩")
        .replace("\u2060", "⟨U+2060⟩")
        .replace("<w:cr/>", "<w:cr/>  ← WPS 非标准软回车")
    )
    for run in re.findall(r"<w:t[^>]*>.*?</w:t>|<w:cr/>", visible):
        print("  " + run)
    feff = document_xml.count("\ufeff")
    zero_width = len(ZERO_WIDTH_RE.findall(document_xml))
    cr = len(CR_RE.findall(document_xml))
    print()
    print(f"  U+FEFF={feff}  零宽字符={zero_width}  软回车<w:cr/>={cr}")


def main() -> int:
    ap = argparse.ArgumentParser(description="生成带隐形垃圾的演示 docx")
    ap.add_argument("-o", "--output", default=DEMO_NAME, help=f"输出文件名（默认 {DEMO_NAME}）")
    ap.add_argument("--show", action="store_true", help="额外打印 document.xml 片段")
    args = ap.parse_args()

    out = fixtures.build_docx(Path(args.output))
    print(f"已生成: {out}（内容全部是占位文本，可安全公开）")
    if args.show:
        with zipfile.ZipFile(out) as z:
            show(z.read("word/document.xml").decode("utf-8"))
    else:
        print("提示：加 --show 可以看到隐形字符显形后的 XML 片段")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
