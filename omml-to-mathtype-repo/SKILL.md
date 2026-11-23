---
name: omml-to-mathtype
description: Convert Word OMML (Office Math) equations into editable MathType OLE equations through a pandoc LaTeX relay, driving MathType's Toggle TeX command via UI Automation without touching the mouse. Use when a .docx contains OMML equations that must become MathType objects, especially when MathType's own "Convert Equations" reports "没有公式被发现" for OMML.
metadata:
  version: "1.0"
---

# OMML → MathType（TeX 中转）

把 docx 里的 Word 原生公式（OMML）换成可双击编辑的 MathType 公式对象。
不依赖 MathType 自带的 OMML 转换（在某些 Word 版本上它对 OMML 视而不见）。

## 环境前提

1. **pandoc**（经 `pip install pypandoc_binary` 安装，`pypandoc.get_pandoc_version()` 能返回版本）。
2. **Word + MathType 加载项**，且 Word 是**正常启动**的（`WINWORD.EXE 文件.docx`）。
   用 COM `/Automation` 启动的实例不会加载 MathType 功能区，`切换 TeX` 按钮找不到。
3. Word 里能打开目标文档；MathType 选项卡存在（标题栏出现 `MathType` 标签页）。
4. 若 `Word.Application` 被 WPS 抢注（`app.Path` 指向 `D:\WPS\...`），先看
   `scripts/word_com_override.ps1` 的说明，并在操作期间**完全退出 WPS**。

## 三步流程

```powershell
# 1) OMML -> LaTeX 文本（显示公式 \[...\]，行内公式 $...$；裸大写词自动包 \mathrm）
python scripts\omml_to_tex.py  输入.docx  中间件.docx

# 2) 在正常启动的 Word 里打开 中间件.docx，然后逐条触发 MathType 的「切换 TeX」
#    （UIA Invoke，不移动鼠标；自动处理段内偏移漂移，失败会回退到 Find）
powershell -File scripts\toggle_tex.ps1 -DocName 中间件.docx

# 3) 核验
python scripts\verify_docx.py 中间件.docx
```

`toggle_tex.ps1` 会把失败项逐条打印出来；对残留标记可以重跑一次（幂等）。

## 判定通过的硬指标

- `oMath` 计数 = **0**
- `OLEObject` 计数 = 原有 MathType 对象数 + 转换条数
- 残留 `\[` / `$` 标记 = 0

`verify_docx.py` 会直接打印这三项；再把 docx 导成 PDF 抽查两页，确认行内符号没有断行、
显示公式仍是居中加右侧编号。

## 已知坑（务必先读）

- MathType 的「切换 TeX」**只认** `\[…\]`（显示）和 `$…$`（行内）；
  `\(…\)` 一律失败，`$$…$$` 会在两侧留下美元符。
- **裸大写词**（`$MMF$`）会被拒；要写成 `${MMF}$` 或 `$\mathrm{MMF}$`。
  `omml_to_tex.py` 已自动处理两字母以上的纯大写。
- 段内有公式对象/域时，按字符偏移选区的坐标会漂移；脚本已内置
  「选区文本 != 标记 → 回退 Word Find」的兜底。
- 转换后 MathType 对象的高度与 OMML 略有差异，**分页会整体前移约一个公式块**，属正常重排。
- Word 另存会清掉 docx 里未被引用的孤立图片，文件体积骤降不是画质被压。

详细说明见 `references/pitfalls.md`，COM 注册表覆盖见 `scripts/word_com_override.ps1`。
