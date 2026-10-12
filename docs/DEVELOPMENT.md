# 矢印 VecStamp 开发文档

本文面向想修改、构建或发布矢印的开发者。使用说明见 [README](../README.md)。

## 项目结构

```
VecStamp/
├── src/                              # PowerPoint 加载项源码
│   ├── VecStamp.bas                  # VBA 源码（UTF-8，可读；唯一需要手改的 VBA 文件）
│   ├── customUI14.xml                # 功能区界面（Office 2010+ 的 customUI 架构）
│   ├── icons/                        # 功能区图标、赞赏二维码卡片（make_icons.py 生成）
│   ├── helpers/xps2svg.ps1           # XPS → SVG 转换脚本（PowerPoint 2013 – 2021 的 SVG 导出；构建时嵌入 VBA）
│   ├── i18n/strings.json             # 加载项的界面文字：zh-CN / zh-TW / en / ja（构建时生成 LoadStrings）
│   ├── presets/presetShapeDefinitions.xml.gz  # ECMA-376 预设形状定义（构建时编译为 PresetSource）
│   └── vsdx/                         # VSDX 模板部件（PowerPoint 与 LibreOffice 共用）
├── libreoffice/extension/            # LibreOffice 扩展源码（Linux / 任意系统的 LibreOffice）
│   ├── vecstamp.py                   # UNO 组件 com.vluckyzhang.vecstamp.Exporter（XJobExecutor）
│   ├── pythonpath/vecstamp_core.py   # 导出引擎：格式、背景、设置，不含界面
│   ├── pythonpath/vecstamp_ui.py     # 对话框、文件选择、取色器、消息框
│   ├── pythonpath/vecstamp_vsdx.py   # VSDX 写出：形状 / 几何 / 文字 XML、打包
│   ├── pythonpath/vecstamp_shapes.py # LibreOffice 形状 → Visio 原生形状（自定义形状公式求值）
│   ├── pythonpath/vecstamp_i18n.py   # 界面语言：T(key, ...)、set_lang()
│   ├── pythonpath/vecstamp_strings.json  # 扩展的界面文字（四种语言，也用于 Addons.xcu 菜单标题）
│   ├── Addons.xcu                    # 菜单与工具栏（make_lo_extension.py 生成）
│   ├── description.xml, META-INF/, description/, icons/
├── build/
│   ├── build.py                      # 一键构建（见下）
│   ├── make_icons.py                 # Logo、横幅、功能区与工具栏图标（纯 SVG 绘制）
│   ├── make_promo.py                 # README 配图（assets/promo，另生成网站用的 *-web.png）
│   ├── make_website.py               # 网站图片、版本号、四种语言页面、SEO 标签与 sitemap
│   ├── build_bas.py                  # 生成纯 ASCII 的 dist/VecStamp.bas，并嵌入 VSDX 模板、预设形状、界面文字
│   ├── make_presets.py               # 把 presetShapeDefinitions.xml 编译成 VBA 用的紧凑格式
│   ├── i18n_tool.py                  # 字符串表检查（四种语言、占位符）与繁体初稿生成（OpenCC）
│   ├── make_readme_tw.py             # 由 README.md 生成 README.zh-TW.md
│   ├── promo_i18n.json               # README 配图的翻译表
│   ├── check_vba.py                  # VBA 静态检查（Win64 / Win32 / Mac 三条编译路径）
│   ├── check_ribbon.py               # 功能区回调、控件 id、图标检查（可选 XSD 校验）
│   ├── make_template.py              # 生成带功能区的 dist/VecStamp-template.pptm
│   └── make_lo_extension.py          # 生成 Addons.xcu 并打包 dist/VecStamp-LibreOffice.oxt
├── dist/                             # 构建产物（随项目发布，便于直接安装）
│   ├── VecStamp.bas, VecStamp-template.pptm, VecStamp-LibreOffice.oxt
│   └── VecStamp.ppam                 # 由 PowerPoint 生成，不纳入版本库
├── installer/
│   ├── windows/VecStamp-Setup.ps1    # UTF-8 BOM + CRLF
│   ├── macos/setup.sh, VecStamp.applescript
│   └── linux/setup.sh
├── tests/
│   ├── test_libreoffice.py           # LibreOffice 扩展端到端测试（含进度窗口与四种语言的对话框）
│   ├── test_i18n.py                  # 字符串表完整性、功能区文字覆盖、生成的 LoadStrings 逐条核对
│   ├── test_vsdx_presets.py          # 177 种预设形状：写成 VSDX 由 LibreOffice 渲染，与参考图比对
│   ├── preset_engine.py              # VBA 预设形状引擎的 Python 参考实现（也可画出全部形状的对照表）
│   ├── test_xps2svg.py               # XPS → SVG 转换脚本测试（需要 pwsh）
│   ├── make_xps_fixture.py           # 生成 tests/xps/sample.xps
│   └── xps/                          # sample.xps 与参考渲染 ref-1.png
├── website/                          # 静态宣传网站：index.html（简体中文源文件）、zh-TW/ en/ ja/（生成）、i18n.json
├── assets/                           # Logo、横幅、promo 配图、Linux 实机截图、赞赏二维码原图
├── docs/DEVELOPMENT.md
├── 安装或卸载-Windows.bat / 安装或卸载-macOS.command / 安装或卸载-Linux.sh
├── README.md, README.zh-TW.md, README.en.md, README.ja.md
├── 快速开始.txt, 快速開始.txt, Quick-Start.txt, クイックスタート.txt
├── CHANGELOG.md, LICENSE, THIRD_PARTY_NOTICES.md
```

## 构建

```bash
pip install cairosvg python-pptx pillow     # cairosvg 缺失时跳过图标，沿用仓库中的 PNG
pip install opencc-python-reimplemented     # 可选：只在生成繁体初稿时需要
python build/build.py                       # 加 --promo 同时重绘 README 与网站配图
```

`build.py` 依次执行：

1. `make_icons.py`：Logo、横幅、功能区图标、赞赏二维码卡片（由 `assets/donate/*-original.*` 裁切）、LibreOffice 工具栏图标；
2. `build_bas.py src/VecStamp.bas dist/VecStamp.bas`：非 ASCII 字符串编码为 `U("十六进制")`，并在文末的生成块中写入：`src/vsdx` → `VsdxPart` / `VsdxMap`；`src/helpers/xps2svg.ps1` → `HelperScript1..N` / `HelperScript`；ECMA-376 预设形状 → `PresetSrc1..N` / `PresetSource`（`make_presets.py`）；`src/i18n/strings.json` → `LoadStrings` 和每种语言的 `LoadStrings_<语言>_<n>`。每个过程都远低于 VBA 单个过程 64 KB 的上限；
3. `check_vba.py`：分别检查源码和生成的模块；
4. `check_ribbon.py`：每个回调在 VBA 中都存在、控件 id 唯一、图标文件齐全；
5. `make_template.py`：把功能区 XML 和图标写入 `dist/VecStamp-template.pptm`；
6. `make_lo_extension.py`：生成 `Addons.xcu`，复制 VSDX 模板，写入版本号，打包 `.oxt`；
7. `make_promo.py`（仅 `--promo`）：需要 Noto Sans CJK SC 和 DejaVu 字体；
8. `make_website.py`（仅 `--promo`）：把配图、二维码、图标复制到 `website/assets`，同步版本号，由 `index.html` 和 `i18n.json` 生成 `zh-TW/`、`en/`、`ja/` 三个页面，为四个页面写入 canonical / hreflang / Open Graph / JSON-LD，并生成 `sitemap.xml` 和 `robots.txt`（单独运行 `python build/make_website.py --pages` 只做页面部分）。

### 生成加载项

`VecStamp.ppam` 只能由 PowerPoint 生成：

- **Windows**：运行安装器选择 `4`（会征得同意后临时开启「信任对 VBA 工程对象模型的访问」，完成后恢复）。
  手动方式：用 PowerPoint 打开 `dist/VecStamp-template.pptm` → Alt+F11 → 「文件 → 导入文件」选 `dist/VecStamp.bas` → 「调试 → 编译 VBAProject」→ 关闭 VBE → 「另存为 → PowerPoint 加载项 (*.ppam)」，保存为 `dist/VecStamp.ppam`。
- **macOS**：同上，在「工具 → 宏 → Visual Basic 编辑器」中导入模块，另存为「PowerPoint 加载项 (.ppam)」。

## 开发约定（VBA）

- **只改 `src/VecStamp.bas`**。VBE 按系统 ANSI 代码页导入 `.bas`，所以发布用的 `dist/VecStamp.bas` 必须是纯 ASCII；`build_bas.py` 会把非 ASCII 字符串转换为 `U("…")`。注释请写英文。
- **界面文字一律用 `Tx("key", 参数…)`**，文字写在 `src/i18n/strings.json`（四种语言，占位符 `{0}`、`{1}`…）。VBA 源码中不再出现中文字面量（`tests/test_i18n.py` 会检查，语言名称「简体中文」等除外）。功能区文字由 `VS_GetLabel` / `VS_GetScreentip` / `VS_GetSupertip` / `VS_GetDescription` / `VS_GetTitle` 按控件 id 查 `rb.<id>`、`rb.<id>.tip`、`rb.<id>.sup`、`rb.<id>.desc`。`Tx` 会保留调用前的 `Err`，可以放心在错误处理程序中使用；函数名 `Tx` 不要被局部变量遮蔽。
- 所有 PowerPoint / Visio 对象后期绑定（`As Object`），枚举一律用数值常量（例如 `ppShapeFormatSVG = 6`、`ppSaveAsPDF = 32`）。
- Windows API 只能出现在 `#If Mac Then … #Else … #End If` 的 Windows 分支中，并使用 `PtrSafe` / `LongPtr`。macOS 上需要沙盒外权限的操作通过 `AppleScriptTask "VecStamp.scpt", "<handler>", "<参数>"` 完成，处理程序写在 `installer/macos/VecStamp.applescript`。
- 错误处理：在活动的错误处理程序中不要再写 `On Error Resume Next`；需要清理时使用 `On Error GoTo Failed` → `Failed:` → `Resume Cleanup` 模式，确保临时幻灯片、临时演示文稿和剪贴板状态被恢复。
- 功能区回调统一以 `VS_` 开头；新增控件时同时修改 `customUI14.xml` 与 `VecStamp.bas`，再运行 `python build/check_ribbon.py`。
- 设置在 Windows 上保存在注册表 `HKCU\Software\VecStamp`（`WScript.Shell.RegRead/RegWrite`，卸载时可一并删除），Mac 上使用 `GetSetting` / `SaveSetting`；读取时一律做范围限制（例如 DPI 1 – 3000、网格 2 – 200 磅、边距 0 – 500 磅）。
- `check_vba.py` 能发现未声明的变量、参数个数不符和 `Exit Sub/Function` 错配，但**不能代替 VBE 中的「调试 → 编译」**，提交前请在真实的 PowerPoint 中编译一次。
- 中文字符串经 `U("…")` 解码，所以**局部变量不要命名为 `u`**（VBA 不区分大小写，会遮蔽 `U()` 函数）。
- 长任务用 `ProgressBegin(标题, 项目数)` / `ProgressItem(i, 名称)` / `ProgressAt(本项进度 0–1, 当前步骤)` / `ProgressEnd` 包裹，在循环中调用 `CheckCancel`（取消时抛出 `ERR_CANCEL`）；入口宏用 `gBusy` 防止重入。进度窗口打开时用 `AskUser` 代替 `MsgBox`，它会先隐藏进度窗口并清掉之前按下的 Esc。
- 外部程序一律用 `RunHidden(命令行, 超时秒数[, 进度文件, 起点, 终点])` 或 `RunPowerShell(脚本, 超时秒数…)` 运行：`CreateProcessW`（`CREATE_NO_WINDOW`）+ `WaitForSingleObject` 轮询，期间刷新进度窗口，Esc 或超时用 `TerminateProcess` 结束；返回退出码或 `RUN_FAILED` / `RUN_CANCELLED` / `RUN_TIMEOUT`。不要再用阻塞的 `WScript.Shell.Run …, True`。

### 导出流程概览

```
DoExport ─ GetTarget（选区 / 整页）
         └ ExportItem ── 需要合成？（多个图形、背景、边距）
                       ├─ 是：AddTempSlide → BuildComposite（背景矩形 / 网格 / 粘贴）→ ExportShape → 删除临时页
                       └─ 否：ExportShape
ExportShape ─ EMF / WMF / PNG：Shape.Export（PNG 按 DPI 计算像素尺寸）
            ├ EMZ：EMF → GzipX
            ├ SVG：ExportSvg（原生导出 → FixSvgFile 校验 → XpsToSvg → Inkscape / LibreOffice 转换 → 询问改为 EMF）
            ├ PDF：ExportPdf（不可见的临时演示文稿，页面 = 图形大小，粘贴 EMF，SaveAs 32）
            └ VSDX：ExportVsdx（VsdxNative 开启：VxShape 遍历生成原生形状；关闭：内嵌 EMF）→ WriteVsdxPackage
```

### XPS → SVG（PowerPoint 2013 – 2021）

`XpsToSvg` 把图形粘贴到隐藏母版与背景的临时幻灯片上，用 `PrintOptions.Ranges` 只选这一页，`ExportAsFixedFormat` 导出 XPS（打印意图）并立即删除临时幻灯片，然后把内嵌的 `xps2svg.ps1` 写到临时文件夹（内容变化时重写），通过 `RunPowerShell` 在后台运行：

```
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -NonInteractive -Command "$ErrorActionPreference='Stop';
  & ([scriptblock]::Create([IO.File]::ReadAllText('<脚本>', [Text.Encoding]::UTF8))) -Xps … -Out … -Frame 'l,t,w,h' -Slide 'w,h' -Progress '<进度文件>' -MaxSeconds 175"
```

用 `scriptblock` 读入脚本，不需要修改执行策略，也不使用 `-EncodedCommand` / `-WindowStyle Hidden`（这两者容易被安全软件当作可疑命令拦截，是 1.2.1 中「一直卡在转换为 SVG」的原因之一）。脚本把进度写入进度文件（`百分比|步骤|a|b`，步骤为 `open` / `page` / `elem` / `font` / `write` / `done`），VBA 的 `ProgressFromFile` 读取后显示；超过 `-MaxSeconds` 时脚本以退出码 3 结束，VBA 端的总超时为 `SVG_TIMEOUT`（180 秒）。脚本兼容 Windows PowerShell 5.1 与 PowerShell 7，只用 .NET 自带的 `System.IO.Compression`：

- 读取 FixedDocumentSequence → FixedDocument → 第一页，解析 `Canvas` / `Path` / `Glyphs`、`RenderTransform`、`Clip`、`Opacity`，以及 `StaticResource` 引用的资源字典；
- 画刷：纯色、线性 / 径向渐变（含 `MappingMode` 与 `SpreadMethod`）、`ImageBrush`（转为 `<pattern>`）；
- 字体：按 XPS 规范用文件名中的 GUID 反序字节异或前 32 字节还原混淆字体，读取 `name` / `OS/2`，`cmap` / `hmtx` 按需解析，`Glyphs` 转为 `<text>`（`UnicodeString` + `Indices` 中的字距），不超过 `-MaxFontBytes`（6 MB）的字体以 `@font-face` 内嵌（`-NoEmbedFonts` 可关闭）；
- 去掉整页白色背景，按「图形框 ∪ 内容包围盒」裁剪，输出无 BOM 的 UTF-8；成功返回 0。

修改脚本后运行 `python3 tests/test_xps2svg.py`；样例由 `tests/make_xps_fixture.py` 生成。

### VSDX 原生形状

`ExportVsdx` → `VxShape` 递归遍历：组合写成 `Type='Group'`（子形状使用组合内的局部坐标），其余由 `VxNative` 尝试转换，失败时 `VxPicture` 导出 EMF 作为 `ForeignData` 形状。

- 几何：预设形状由 `VxPresetGeo` 计算。`build/make_presets.py` 把 ECMA-376 的 `presetShapeDefinitions.xml` 编译成每种形状一条紧凑字符串 `调整值|guide 公式|路径|文字框`（运算符编码 `m`=`*/`、`a`=`+-`、`d`=`+/`、`q`=`?:`、`b`=abs、`t`=at2、`c`=cat2、`o`=cos、`x`=max、`n`=min、`h`=mod、`p`=pin、`s`=sat2、`i`=sin、`r`=sqrt、`g`=tan、`v`=val；引用 `$n`=guide、`@n`=调整值、`!` 标记角度调整值），按 `AutoShapeType` 查表（177 种预设、182 个 AutoShape 类型）。PowerPoint 的 `Adjustments` 换算为 DrawingML 单位（线性值 ×100000，角度 ×60000）。`arcTo` 的起止角是「视角」，换算为椭圆参数角 `atan2(wR·sin a, hR·cos a)` 后按 ≤ 90° 分段写成 `EllipticalArcTo`（终点 + 弧上中点 + 长短轴比），二次贝塞尔升为三次，三次贝塞尔写为 `NURBSTo`。没有预设定义的形状仍用连接线公式或「合并形状 → 联合」取轮廓。自由曲线读取 `Nodes`，按共线性评分判断节点类型属于前段还是后段。翻转在几何中处理，旋转写入 `Angle`（`-rotation`）。`tests/preset_engine.py` 是同一引擎的 Python 版本，`tests/test_vsdx_presets.py` 用它写出全部预设形状的 VSDX，经 LibreOffice 渲染后与 cairosvg 画出的参考图逐格比对（差异超过 4 % 判为失败）。
- 样式：`FillForegnd` / `FillPattern` / `FillForegndTrans`、`FillGradient` 段、`LineWeight`、`LinePattern`（PowerPoint 虚线样式映射）、`BeginArrow` / `EndArrow` 与大小。
- 文字：`Character` / `Paragraph` 段按文字格式逐段写入，正文中用 `<cp IX>` / `<pp IX>` 切换；`TxtPinX` 等单元格定位文本框，内边距与垂直对齐写入 `LeftMargin` / `VerticalAlign` 等。
- 包：`WriteVsdxPackage` 用模板部件（`src/vsdx`）替换 `{SHAPES}` / `{RELS}`，`[Content_Types].xml` 放在首位。

LibreOffice 端对应 `vecstamp_shapes.py`：自行解析 `EnhancedCustomShapeGeometry`（公式、段命令、`SubViewSize`、`TextFrames`、镜像），不依赖界面渲染。

### 背景菜单、颜色面板、进度窗口与反馈

- 「背景」是一个 `box`：颜色预览按钮 + `menu`。Windows 上预览按钮 `vsBgSwatch` 用 `getImage` 返回运行时生成的 BMP（透明为棋盘格、网格为坐标纸、其余为颜色，`LoadPicture` 加载）；Mac 不能在运行时生成图片，改为五个静态图标按钮 `vsBgSw*`，用 `getVisible` 只显示当前背景对应的一个。菜单中是 `toggleButton`（透明 / 白色 / 米色 / 网格）和「自定义颜色」：Windows 上是颜色库 `vsGalColor`（前 60 项是主题颜色——bg1、tx1、bg2、tx2、accent1–6 按 Office 的 HSL 规则生成 5 级深浅，接着 10 种标准色和最多 10 个最近使用的颜色 `RecentColors`，以及「其他颜色」「取色器」按钮），Mac 上是 `vsBgCustomMac`。
- 「其他颜色」「取色器」在一个临时的幻灯片外形状上调用 `CommandBars.ExecuteMso("ObjectFillMoreColorsDialog" / "EyedropperFill")`，读取其填充色后删除。对话框不可用时 `ColorPanel` 打开矢印自己的颜色面板：一个 `#32770` 窗口，色块是 `SS_BITMAP | SS_NOTIFY` 静态控件中的 BMP，点击位置从消息循环（`PeekMessageW`，点击、双击、Enter / Esc / 空格在分发前判断，其余交给 `IsDialogMessageW`）中读出，不需要窗口过程；面板打开时禁用 PowerPoint 主窗口。再不可用时依次回退到 `ChooseColorW` 和 `ColorListPanel`（输入框中的编号列表）。Mac：辅助脚本的 `chooseColor`（macOS 颜色面板；取消返回空串，出错返回 `error`），出错时使用 `ColorListPanel`。
- 进度窗口：Windows 用 `CreateWindowExW` 创建不激活的对话框类窗口（标题、百分比、`msctls_progress32`、当前步骤与已用时间），按 `LOGPIXELSY` 缩放，`GetAsyncKeyState(VK_ESCAPE)` 检测取消；macOS 用辅助脚本的 `progressShow` / `progressUpdate` / `progressAlive` 显示后台对话框（文字中带百分比，按钮文字由 VBA 传入），对话框被点掉即视为取消。
- 反馈：`DoFeedback` 是「是 / 否 / 取消」消息框：邮件（`SetClipboardText` 复制邮箱，可选 `mailto:`，正文附环境信息）或 GitHub Issue（`VS_REPO/issues/new?title=…&body=…`，UTF-8 百分号编码）。仓库地址在 `VS_REPO` 与 `vecstamp_core.REPO` 中定义。

## 界面语言（i18n）

- 四种语言：`zh-CN`（默认）、`zh-TW`、`en`、`ja`。字符串表都是 `{key: {"zh-CN": …, "zh-TW": …, "en": …, "ja": …}}`：
  - `src/i18n/strings.json`：PowerPoint 加载项（含功能区 `rb.*`）；
  - `libreoffice/extension/pythonpath/vecstamp_strings.json`：LibreOffice 扩展（含 `menu.*` / `tb.*` 菜单标题）；
  - `website/i18n.json`：宣传网站，键是 `index.html` 中的简体中文原文；
  - `build/promo_i18n.json`：README 配图。
- 新增文字：在表中加一条（四种语言），繁体可以先留空，运行 `python build/i18n_tool.py fill-tw` 用 OpenCC（s2twp + 台湾用语修正表 `TW_FIXES`）生成初稿后人工检查；`python build/i18n_tool.py check` 检查缺项和占位符。`fill-tw` 只填空白项，人工修改过的繁体不会被覆盖（`--force` 才会全部重新生成）。
- PowerPoint：当前语言保存在设置项 `Lang`，`CurLang()` 读取并缓存；`SetLang` 写入设置并清空缓存，随后 `gRibbon.Invalidate` 让所有 `get*` 回调重新取字。`LoadStrings` 只加载当前语言（一个 `Collection`）。原生窗口的字体按语言选择（`UiFontName`：微软雅黑 / 微软正黑体 / Yu Gothic UI / Segoe UI）。
- LibreOffice：语言是设置项 `lang`，`core.Settings()` 读取时调用 `vecstamp_i18n.set_lang()`；界面代码一律 `T('key', …)`。`Addons.xcu` 的菜单标题由 `make_lo_extension.py` 从字符串表写入 `xml:lang` 为 `zh-CN` / `zh-TW` / `en-US` / `ja` 的四个值，随 LibreOffice 自身的界面语言显示。
- README：`README.md` 是简体中文主文档；`README.zh-TW.md` 由 `python build/make_readme_tw.py` 转换（代码、链接和仓库中的文件名保持不变），`README.en.md` 与 `README.ja.md` 手工维护，修改主文档时请同步。
- 测试：`python3 tests/test_i18n.py` 检查四个字符串表、VBA 与扩展用到的每个键、功能区每个控件的文字，并解释执行 `dist/VecStamp.bas` 中生成的 `LoadStrings`，与 JSON 逐条比对。

## LibreOffice 扩展

- 菜单和工具栏项在 `build/make_lo_extension.py` 的 `MENU` / `TOOLBAR` 列表中定义（标题来自 `vecstamp_strings.json`），URL 形如 `service:com.vluckyzhang.vecstamp.Exporter?export:svg`，由 `vecstamp.py` 的 `trigger()` 分发。
- 进度：`core.export_target(…, progress)` 在每个步骤调用 `progress(本文件进度 0–1, 步骤文字)`，原生 VSDX 逐个形状回调；界面层的 `Progress` 是一个非模态 UNO 对话框（标题、当前文件、进度条、百分比、步骤、已用时间、「取消」），通过 `XToolkitExperimental.processEventsToIdle()` 刷新，`XKeyHandler` 捕获 Esc，取消后在下一个步骤抛出 `core.Cancelled`；同时更新状态栏进度。没有显示环境时只用状态栏。
- 颜色：`UI.pick_color` 先用 `com.sun.star.cui.ColorPicker`，创建失败时使用 `ColorDialog`（UNO 对话框：主题色及深浅、标准色、最近使用的颜色、颜色值输入与预览，单击选色、双击确定）。
- 引擎用 `com.sun.star.drawing.GraphicExportFilter` 导出 ShapeCollection；背景通过临时插入的矩形 / 网格线实现，并在撤销管理器锁定下完成，导出后删除，不污染文档的撤销历史。
- 导出先写到 ASCII 路径的临时文件，再由 Python 移动到目标位置（规避部分环境下 LibreOffice 无法写入非 ASCII 路径的问题），最后按用户 umask 设置权限。
- PDF：先导出 SVG，再放到一个页面大小与图形一致的临时 Draw 文档中导出 PDF。
- 设置：`$XDG_CONFIG_HOME/VecStamp/libreoffice-settings.json`（Windows：`%APPDATA%\VecStamp\…`，macOS：`~/Library/Application Support/VecStamp/…`）。
- `trigger("selftest:<目录>")` 是不弹任何对话框的自检入口，供测试使用。

### 测试

```bash
python3 tests/test_libreoffice.py               # 以普通用户运行（unopkg 拒绝 root）
VECSTAMP_KEEP=1 python3 tests/test_libreoffice.py   # 保留导出的文件以便查看
```

测试会在临时配置目录中安装扩展、启动无界面的 LibreOffice、绘制示例图形，并对 5 种背景 × 7 种格式逐一校验文件头、SVG 根元素和 VSDX 包结构（原生形状模式检查 `Shape` / `Geometry` / 文字，图片模式检查 `ForeignData`）；检查每种格式的进度回调（比例递增、以 1.0 结束、每步都有文字）、进度窗口的百分比与取消、颜色对话框，并在四种语言下分别构建导出设置与「关于」对话框，不会影响你自己的 LibreOffice 配置。

VSDX 预设形状与界面语言：

```bash
python3 tests/test_vsdx_presets.py   # 需要 LibreOffice 与 cairosvg；VECSTAMP_KEEP=1 时保存左右对比图 tests/presets-compare.png
python3 tests/test_i18n.py           # 先运行 build.py（或 build_bas.py）生成 dist/VecStamp.bas
python3 tests/preset_engine.py out.svg   # 画出全部预设形状的对照表
```

XPS → SVG 转换脚本：

```bash
python3 tests/test_xps2svg.py                     # 需要 pwsh（PowerShell 7）或 Windows PowerShell
```

## README 配图与宣传网站

`python build/make_promo.py` 生成 `assets/promo/*.png`（以及网站用的平底版本 `ribbon-web.png`、`linux-web.png`），并按 `build/promo_i18n.json` 把与界面无关的配图（hero、workflow、formats、backgrounds、vector-vs-bitmap、platforms、social-preview）再画成 `*.zh-TW.png`、`*.en.png`、`*.ja.png`，供各语言的 README 和网站分享图使用；`T()` / `tw()` 绘制和测量的每段中文都会查表，缺少翻译时脚本报错退出。其中：

- 除 `linux.png` 外均为脚本绘制的示意图；`ribbon.png` 按 `customUI14.xml` 的布局绘制并标注「示意图」。
- `assets/screenshots/` 是真实截图：在 Xvfb 虚拟显示器（1600×1000）中以中文界面启动装有扩展的 LibreOffice Impress（GTK3 界面），通过 UNO 绘制示例幻灯片，再用 ImageMagick `import` 截取菜单、导出设置和取色器。界面有变化时请重新截取。
- 赞赏二维码原图在 `assets/donate/*-original.*`；裁切区域在 `make_icons.py` 的 `DONATE_CARDS` 中定义。更换二维码后重新运行 `build.py --promo`，并用扫码工具确认生成的卡片能被识别。

`website/` 是纯静态网站：`index.html` 内联全部 CSS / JS，首屏演示用浏览器把同一段内联 SVG 画到 600×360 的 canvas 上作为「截图」，与矢量 SVG 左右对比（缩放通过修改 `viewBox` 实现，保证矢量侧始终清晰）。`make_website.py` 用正则逐个替换文本节点和 `alt` / `aria-label` / `title` / `content` / `data-badge` 属性（`<style>`、`<script>` 与注释保持不动，`id="i18n"` 的 JSON 块按值翻译），写出 `zh-TW/`、`en/`、`ja/` 页面（共享文件路径加上 `../`，按语言调整 CJK 字体顺序），语言菜单与 SEO 块按页面生成。部署方法见 [website/README.md](../website/README.md)。

## GitHub 仓库设置（便于检索）

- **About → Description**：`矢印 VecStamp – free PowerPoint add-in & LibreOffice extension: export shapes to EMF, SVG, PDF, Visio VSDX and high-DPI PNG in one click. 一键导出 PPT 矢量图。`
- **Website**：`http://vluckyzhang.cn/shiyin/`
- **Topics**：`powerpoint`、`powerpoint-addin`、`vba`、`libreoffice`、`libreoffice-extension`、`impress`、`svg`、`emf`、`vector-graphics`、`visio`、`vsdx`、`pdf`、`export`、`scientific-figures`、`academic`、`diagram`、`office-addin`、`i18n`
- **Social preview**：上传 `assets/promo/social-preview.png`（1280×640）。
- Release 标题写成「矢印 VecStamp v1.2.2 – PowerPoint 导出 SVG / EMF / PDF / Visio」，正文摘要 CHANGELOG 并附 `VecStamp.ppam`、`VecStamp-LibreOffice.oxt`。

## 发布清单

1. 更新版本号：`src/VecStamp.bas`（`VS_VERSION`）、`libreoffice/extension/pythonpath/vecstamp_core.py`（`VERSION`）、`vecstamp_vsdx.py` 的默认参数、`installer/windows/VecStamp-Setup.ps1`、`installer/macos/setup.sh`、`installer/macos/VecStamp.applescript`、`installer/linux/setup.sh`、四个 README 的徽章、四个快速开始文件（网站版本号由 `make_website.py` 同步）。
2. 在 `CHANGELOG.md` 中记录变更，更新 `website/index.html` 中的「更新」一节和 `website/i18n.json` 中对应的翻译，更新 README（并重新生成 `README.zh-TW.md`，同步英文与日文版）。
3. `python build/build.py --promo`，并运行 `python3 build/i18n_tool.py check`、`python3 tests/test_i18n.py`、`python3 tests/test_vsdx_presets.py`、`python3 tests/test_libreoffice.py` 与 `python3 tests/test_xps2svg.py`。
4. 检查安装器与脚本：`pwsh` 解析 `VecStamp-Setup.ps1` 和 `src/helpers/xps2svg.ps1`，`shellcheck installer/*/*.sh 安装或卸载-*.sh 安装或卸载-macOS.command`。
5. 在 Windows PowerPoint 中用安装器 `4` 生成 `dist/VecStamp.ppam`，在 PowerPoint（以及可能的话 Mac 与 Visio）中实测一遍。
6. 打包发布：上传项目压缩包，并单独附上 `VecStamp.ppam` 与 `VecStamp-LibreOffice.oxt`；需要时把压缩包放到网站的 `download/` 并修改 `website/config.js`。
