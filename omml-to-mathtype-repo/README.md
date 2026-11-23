# omml-to-mathtype

把 Word 文档里的 **OMML（Word 原生公式）** 批量转换成可双击编辑的 **MathType 公式对象**。

适用于 MathType 自带的「转换公式」在某些 Word 版本上对 OMML 完全失灵的情况：
本工具改为走 **pandoc → LaTeX → MathType「切换 TeX」** 的中转路线，全程用 UI Automation
驱动，不移动鼠标。

## 原理

1. `pandoc` 把每个 OMML 公式转成 LaTeX（比 Office 自带 XSLT 保真度高得多）。
2. 在 docx 里把公式替换为纯文本标记：显示公式用 `\[…\]`，行内公式用 `$…$`。
3. 在 Word 中逐条选中标记，调用 MathType 功能区的「切换 TeX」，生成 MathType OLE 对象。

## 环境要求

- Windows + Microsoft Word（实测 Word 2021 16.0.20326）
- MathType 7.x（带「MathType Commands」加载项，功能区有 MathType 选项卡）
- Python 3.9+，依赖：`pip install pypandoc_binary lxml`
- Word 需**正常启动**（`WINWORD.EXE 文件.docx`）；用 COM `/Automation` 启动的实例不会加载 MathType 功能区

## 快速开始

```powershell
# 1) OMML -> LaTeX 文本
python scripts\omml_to_tex.py 输入.docx 中间件.docx

# 2) 用 WORD 正常打开 中间件.docx，然后逐条触发「切换 TeX」
powershell -File scripts\toggle_tex.ps1 -DocName 中间件.docx

# 3) 核验（oMath 必须为 0，残留标记必须为 0）
python scripts\verify_docx.py 中间件.docx
```

## 脚本

| 文件 | 说明 |
|---|---|
| `scripts/omml_to_tex.py` | OMML → LaTeX 文本；显示/行内分别用 `\[…\]`、`$…$`；纯大写词自动包 `\mathrm{}` |
| `scripts/toggle_tex.ps1` | 依附运行中的 Word，用 UIA `Invoke()` 点「切换 TeX」；选区漂移时回退 Word Find |
| `scripts/verify_docx.py` | 统计 `oMath` / OLE 对象 / 残留标记，可导出 PDF |
| `scripts/word_com_override.ps1` | WPS 抢注 `Word.Application` 时的诊断与临时注册表覆盖 |

## 已知限制

- MathType 的「切换 TeX」**只认** `\[…\]` 和 `$…$`；`\(…\)` 会失败，`$$…$$` 会残留美元符。
- 裸大写词（如 `MMF`）需要 `\mathrm{}`，已自动处理。
- 转换后分页会整体前移约一个公式块（MathType 对象高度与 OMML 略有差异）。
- 仅支持 Windows + 桌面版 Word + MathType，不支持 WPS / macOS / 在线版 Word。

更多细节见 [`references/pitfalls.md`](references/pitfalls.md)。

## 作为 Claude Code Skill 安装

把仓库内容复制到 `~/.claude/skills/omml-to-mathtype/` 即可（`SKILL.md` 在根目录）。

## License

MIT
