# docx-encoding-cleaner

> 清理 `.docx` 里的**隐形字符**与 **WPS 非标准软回车**，修掉"同一份文档，有人打开正常、有人打开少字/吞字"的问题。
>
> Remove invisible junk from `.docx`: BOM / zero-width characters (U+FEFF, U+200B–200D, U+2060, U+180E) and WPS-style `<w:cr/>` soft returns.

**纯标准库，零依赖** · Python 3.8+ 跨平台 · 另附 Windows 拖放脚本

---

## 案例：一份 23 页文档，在同事电脑上"少了两行"

### 症状

一份由 WPS 保存、用 Word 打开的小组文档（23 页）：

- 在**同事的电脑**上打开时，正文里连续两行（`人物 5：…` / `人物 6：…`）**整体消失**；
- 同一份文件，在**另一台电脑的 Word** 里完全正常；
- 用 **WPS** 打开也完全正常；
- **导出 PDF** 后文字也都在。

而且从截图实测：缺字处上下文字行的基线间距是均匀的（29–31 px），**没有多出来的空档**——说明那两行不是"排好了没画出来"，而是**根本没进入排版结果**。

```
同事电脑上的 Word（修复前）          修复后（同一台机器）
────────────────────────────         ────────────────────────────
人物 4：外星人                        人物 4：外星人
大愚若智。……所以也抖不起来包袱。       大愚若智。……所以也抖不起来包袱。
有些反对最近的 ai 绘画风潮……          人物 5：绘客 2026          ← 回来了
摇滚老炮，将 choose life 挂在嘴边……   有些反对最近的 ai 绘画风潮……
                                      人物 6：鹅                ← 回来了
                                      摇滚老炮，将 choose life 挂在嘴边……
```

> 上表文字用原案例的行数结构说明问题，实际公开仓库不含任何原文内容。

### 排查（关键动作与证据）

1. **解开 docx 直接看 `word/document.xml`**：那两行文字**都在**，而且字体、字号、加粗、下划线等属性与上下相邻的正常行**完全一致** → 排除"隐藏文字 / 白字 / 被删掉"。
2. **全量统计后发现了真正的异常**：这份 WPS 保存的文档里带着

   | 隐患 | 数量 | 说明 |
   |---|---|---|
   | `U+FEFF`（ZERO WIDTH NO-BREAK SPACE，BOM 残留） | 20 | 全部贴在每行文字末尾；不可见，但 Unicode 里属于"禁止在此断行"的字符 |
   | `<w:cr/>`（WPS / 旧版 WordprocessingML 的软回车） | 19 | Word 自己写文件用的是标准 `<w:br/>` |
   | 零宽字符 `U+200B/200C/200D/2060` | 2 | 藏在 `word/theme/theme1.xml` 等处 |

3. **清掉这两类标记后，同事那台机器恢复正常**（同一个 Word、同一份内容）。
4. **回归验证**：本工具输出的 `word/document.xml` 与当时人工修复的结果 **SHA-256 完全相同**。

### 为什么"换台电脑就好了"

Word 的断行与排版依赖**本机环境**：默认打印机的字体度量、字体是否被替换、Word 版本/build、第三方加载项……文件里的这套非标准标记是**触发条件**，本机环境是**扳机**——两个条件同时满足才会发作。WPS 的排版引擎不受这套标记影响，所以同一份文件在 WPS 里永远正常。本工具做的事，就是把**文件里的触发条件**彻底拿掉。

### 自己复现（不需要任何真实文档）

```console
$ python examples/make_demo.py --show
已生成: demo_with_junk.docx（内容全部是占位文本，可安全公开）
  <w:t>人物介绍</w:t>
  <w:t>人物 1：甲⟨U+FEFF⟩</w:t>
  <w:cr/>
  <w:t>待业青年⟨U+200B⟩，生活⟨U+200D⟩松⟨U+2060⟩弛。⟨U+FEFF⟩</w:t>
  <w:cr/>
  <w:t>人物 2：EXAMPLE-01</w:t>
  <w:cr/>
  <w:t>故事背景</w:t>

  U+FEFF=2  零宽字符=3  软回车<w:cr/>=3

$ python docx_clean.py demo_with_junk.docx --dry-run
[检查] demo_with_junk.docx
U+FEFF=2  零宽字符=3  软回车<w:cr/>=3
      word/document.xml: FEFF=2 零宽=3 软回车=3

$ python docx_clean.py demo_with_junk.docx
[完成] demo_with_junk_clean.docx
U+FEFF=2  零宽字符=3  软回车<w:cr/>=3
      word/document.xml: FEFF=2 零宽=3 软回车=3

$ python docx_clean.py demo_with_junk_clean.docx
[跳过] demo_with_junk_clean.docx（未发现需要清理的标记）
```

> 生成的文件用 Word 打开通常**看不出异常**——是否发作取决于本机环境；所以请用 `--show` / `--dry-run` 观察标记，而不是指望在每台机器上都复现丢字。

## 它做什么 / 不做什么

**做**（只碰标记，不动正文一个字）：

1. 删除 `U+FEFF`
2. 删除零宽字符 `U+200B`–`U+200D`、`U+2060`、`U+180E`
3. `<w:cr/>` → `<w:br/>`（规范成 Word 标准换行）

**不做**：不处理"对方没装这个字体""对方 Word 版本太老""默认打印机度量不同"这类**本机环境**问题——那类问题只能靠**同时发一份 PDF**兜底。本工具消灭的是文件里的触发条件。

## 使用（Python 版）

无需安装任何依赖，直接运行：

```bash
python docx_clean.py 报告.docx              # 生成 报告_clean.docx（原文件不动）
python docx_clean.py 报告.docx -o out.docx  # 指定输出
python docx_clean.py 报告.docx --dry-run    # 只体检，不写文件
python docx_clean.py 报告.docx --in-place   # 就地清理（留 .bak 备份）
python docx_clean.py 文件夹/                # 批量处理目录下的 docx
python docx_clean.py *.docx --json          # 机器可读输出
```

- 文件本来就干净时，默认**不会**产出多余副本，只提示"未发现需要清理的标记"
- 你的原文件永远只读不写（`--in-place` 除外，且会留 `.bak`）

## 使用（Windows 拖放版）

把 `.docx` 文件**拖到 `clean-docx.cmd` 上**，旁边就会生成 `原名_clean.docx`。只用系统自带的 PowerShell，不需要安装任何东西（`clean-docx.cmd` 与 `docx-clean.ps1` 需放在同一目录）。

## 它是怎么被验证的

- **17 个单元测试**（只用标准库）：计数正确、正文逐字不变、其余部件（图片/样式/`[Content_Types].xml`）字节不变、zip 结构可读、幂等、`--dry-run` 不落盘、`--in-place` 留备份、CLI 与 JSON 输出
- **真实文档回归**：对那份 23 页故障文档，清理后的 `word/document.xml` 与人工修复版 **SHA-256 完全相同**
- **幂等性**：对已清理的文件再跑一次，报告 0 处、不产出新文件

```bash
python -m unittest discover -s tests -t . -v
```

> 受限/沙箱环境若不允许删除临时目录，可设 `DOCX_CLEAN_TEST_TMP` 指向一个可写目录。

## 顺带一提：怎么避免产出这种文件

1. 粘贴网页 / AI / PDF 内容用 **Ctrl+Shift+V**（选择性粘贴 → 无格式文本），或先过一遍记事本
2. 一份文档别在 WPS 与 Word 之间反复"另存为"；跨软件流转时，中途用 Word 打开再另存一次
3. 正文只用通用字体（宋体/黑体/微软雅黑/Times New Roman/Arial）
4. 交付前：跑一遍本工具体检 → 导出 PDF 翻一遍关键页 → **PDF + docx 一起发**

## License

[MIT](LICENSE)

---

## English

`docx-encoding-cleaner` removes encoding-level junk from `.docx` files that makes Microsoft Word drop whole lines of text on some machines while rendering fine on others:

- `U+FEFF` (BOM / zero-width no-break space) left over from pasting web / AI / PDF content
- zero-width characters `U+200B`–`U+200D`, `U+2060`, `U+180E`
- `<w:cr/>` soft returns written by WPS Office, normalised to Word's standard `<w:br/>`

**Case in short:** a 23-page WPS-saved document showed two consecutive lines missing in Word on one colleague's PC, while the very same file rendered fine in Word on another PC, fine in WPS, and exported to PDF with every line present. The text was in the file all along (identical run properties to neighbouring, correctly rendered lines) — the 20 trailing `U+FEFF` characters and 19 `<w:cr/>` soft returns were the trigger. Removing them fixed that machine, and this tool's output is byte-identical to the hand-fixed file.

Word's line breaking depends on the local environment (default printer metrics, font substitution, build, add-ins), which is why the same file can behave differently on two machines. This tool removes the *file-side* trigger.

Pure Python standard library, no dependencies. It never touches the text content and never modifies the source file (except with `--in-place`, which keeps a `.bak`).

```bash
python docx_clean.py report.docx            # writes report_clean.docx
python docx_clean.py report.docx --dry-run  # report only
python examples/make_demo.py --show         # build a demo file and reveal its hidden junk
```

A Windows drag-and-drop wrapper (`clean-docx.cmd` + `docx-clean.ps1`) is included.

Verified with 17 unit tests and against a real 23-page document whose cleaned `word/document.xml` is byte-identical (same SHA-256) to a hand-fixed version that resolved the missing-lines symptom.

Note: this fixes the *file-side* trigger only. Cross-machine sharing should always include a PDF, since Word's layout also depends on local printer metrics, installed fonts and build version.
