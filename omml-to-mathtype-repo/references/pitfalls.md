# OMML → MathType 实测坑位（2026-11，Word 2021 + MathType 7.4.4）

## 1. MathType 自带的 OMML 转换在部分 Word 版本上失效

- 现象：`转换公式` 对话框勾选「Word 2007 及以上 (OMML) 公式」后报
  「没有公式被发现和/或更新」，而同一文档里勾「MathType 或 Equation Editor 公式」
  能正常识别并更新已有 MathType 对象。
- 证据：MathType 自己的统计键
  `HKCU\Software\Design Science\DSMT7\WCStats\CvtEqFromOMML` 在一次运行后不增长，
  说明它根本没尝试转换。
- 用 Word 自己生成的原生 OMML 文档同样为 0，排除文档来源问题。
- 结论：不要在这条路上调参数，直接走 TeX 中转。

## 2. Toggle TeX 只认两种定界符

| 写法 | 结果 |
|---|---|
| `\[…\]` | ✅ 显示公式（居中、独立段落） |
| `$…$` | ✅ 行内公式（不拆行） |
| `\(…\)` | ❌ 一律失败 |
| `$$…$$` | ❌ 转换后两侧残留美元符 |

## 3. 裸大写词会被拒

`$MMF$`、`\[MMF\]` 均失败；`${MMF}$`、`$\mathrm{MMF}$` 正常。
两字母以上的纯大写已在 `omml_to_tex.py` 中自动包 `\mathrm{}`。
单字母、带小写的词（`Dist`、`max`）不需要处理。

## 4. 段内字符偏移会漂移

当段落里已有公式对象（`\x01`）或域时，`paragraph.Range.Start + match.Index`
算出的选区可能偏前几十个字符。表现是选中的根本不是标记文本。
兜底：选区文本与标记不一致时改用 Word `Find`（非通配符、`Wrap=0`）定位再执行。
`toggle_tex.ps1` 已内置该回退。

## 5. Word 必须以普通方式启动

用 COM `/Automation` 启动的 Word 实例里，MathType 功能区标签页不存在，
UIA 找不到「切换 TeX」。先 `Start-Process WINWORD.EXE 文件.docx`，再用
`GetActiveObject` 挂上去做选区。

## 6. WPS 抢占 Word.Application

- 32 位注册表 `CLSID\{000209FF-…}` 指向 `wps.exe /Automation`；
  `New-Object -ComObject Word.Application` 返回的 `Name` 也叫 "Microsoft Word"，
  靠 `Path` 才能识破。
- WPS 进程活着时它持有 COM 类对象，改注册表也没用；必须先完全退出 WPS
  （含托盘进程），再写 HKCU 覆盖（见 `word_com_override.ps1`）。

## 7. UIA 而不是键盘 / 鼠标

- `PostMessage` 点击 VBA UserForm 的自绘控件无效；真实鼠标点击有效但会抢用户鼠标。
- `SendKeys("%\\")`（Alt+\）在焦点不确定时经常无效。
- 可靠做法：`Selection.Select()` 由 COM 完成，按钮用 UIA
  `InvokePattern.Invoke()`。全程不需要前台焦点。

## 8. 版面变化

MathType 对象的高度/基线略大于 OMML，转换后分页整体前移约一个公式块；
行内符号保持行内，显示公式仍居中并保留右侧编号。

## 9. Word 另存会清理孤立图片

如果 docx 是用脚本删段落生成的，媒体目录里可能残留未被引用的图片。
Word 保存后这些孤儿会被清掉，文件体积从数 MB 降到数百 KB，属正常现象，
不是重新压缩画质（引用中的图片字节不变）。

## 10. PowerShell 脚本编码

`powershell.exe` 5.1 按系统 ANSI 解析无 BOM 的 `.ps1`，中文（如按钮名
「切换 TeX」）会变乱码导致查找失败。脚本必须保存为 **UTF-8 with BOM**。

## 11. 依赖

- `pip install pypandoc_binary`（自带 pandoc 可执行文件）
- Python 侧只需 `lxml`、`pypandoc`
- UIA 由 PowerShell 的 `UIAutomationClient` 程序集完成，不需要 `uiautomation` 包
