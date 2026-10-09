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
│   └── vsdx/                         # VSDX 模板部件（PowerPoint 与 LibreOffice 共用）
├── libreoffice/extension/            # LibreOffice 扩展源码（Linux / 任意系统的 LibreOffice）
│   ├── vecstamp.py                   # UNO 组件 com.vluckyzhang.vecstamp.Exporter（XJobExecutor）
│   ├── pythonpath/vecstamp_core.py   # 导出引擎：格式、背景、设置，不含界面
│   ├── pythonpath/vecstamp_ui.py     # 对话框、文件选择、取色器、消息框
│   ├── pythonpath/vecstamp_vsdx.py   # VSDX 写出：形状 / 几何 / 文字 XML、打包
│   ├── pythonpath/vecstamp_shapes.py # LibreOffice 形状 → Visio 原生形状（自定义形状公式求值）
│   ├── Addons.xcu                    # 菜单与工具栏（make_lo_extension.py 生成）
│   ├── description.xml, META-INF/, description/, icons/
├── build/
│   ├── build.py                      # 一键构建（见下）
│   ├── make_icons.py                 # Logo、横幅、功能区与工具栏图标（纯 SVG 绘制）
│   ├── make_promo.py                 # README 配图（assets/promo，另生成网站用的 *-web.png）
│   ├── make_website.py               # 复制网站图片与二维码、同步网站版本号
│   ├── build_bas.py                  # 生成纯 ASCII 的 dist/VecStamp.bas，并嵌入 VSDX 模板
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
│   ├── test_libreoffice.py           # LibreOffice 扩展端到端测试
│   ├── test_xps2svg.py               # XPS → SVG 转换脚本测试（需要 pwsh）
│   ├── make_xps_fixture.py           # 生成 tests/xps/sample.xps
│   └── xps/                          # sample.xps 与参考渲染 ref-1.png
├── website/                          # 静态宣传网站（index.html + config.js + assets）
├── assets/                           # Logo、横幅、promo 配图、Linux 实机截图、赞赏二维码原图
├── docs/DEVELOPMENT.md
├── 安装或卸载-Windows.bat / 安装或卸载-macOS.command / 安装或卸载-Linux.sh
├── README.md, CHANGELOG.md, 快速开始.txt, LICENSE, THIRD_PARTY_NOTICES.md
```

## 构建

```bash
pip install cairosvg python-pptx pillow     # cairosvg 缺失时跳过图标，沿用仓库中的 PNG
python build/build.py                       # 加 --promo 同时重绘 README 与网站配图
```

`build.py` 依次执行：

1. `make_icons.py`：Logo、横幅、功能区图标、赞赏二维码卡片（由 `assets/donate/*-original.*` 裁切）、LibreOffice 工具栏图标；
2. `build_bas.py src/VecStamp.bas dist/VecStamp.bas`：中文字符串编码为 `U("十六进制")`，把 `src/vsdx` 嵌入 `VsdxPart` / `VsdxMap` 生成块，把 `src/helpers/xps2svg.ps1` 嵌入 `HelperScript1..N` / `HelperScript`（每个过程约 200 行，远低于 VBA 单个过程 64 KB 的上限）；
3. `check_vba.py`：分别检查源码和生成的模块；
4. `check_ribbon.py`：每个回调在 VBA 中都存在、控件 id 唯一、图标文件齐全；
5. `make_template.py`：把功能区 XML 和图标写入 `dist/VecStamp-template.pptm`；
6. `make_lo_extension.py`：生成 `Addons.xcu`，复制 VSDX 模板，写入版本号，打包 `.oxt`；
7. `make_promo.py`（仅 `--promo`）：需要 Noto Sans CJK SC 和 DejaVu 字体；
8. `make_website.py`（仅 `--promo`）：把配图、二维码、图标复制到 `website/assets`，同步 `website/config.js` 与 `index.html` 中的版本号。

### 生成加载项

`VecStamp.ppam` 只能由 PowerPoint 生成：

- **Windows**：运行安装器选择 `4`（会征得同意后临时开启「信任对 VBA 工程对象模型的访问」，完成后恢复）。
  手动方式：用 PowerPoint 打开 `dist/VecStamp-template.pptm` → Alt+F11 → 「文件 → 导入文件」选 `dist/VecStamp.bas` → 「调试 → 编译 VBAProject」→ 关闭 VBE → 「另存为 → PowerPoint 加载项 (*.ppam)」，保存为 `dist/VecStamp.ppam`。
- **macOS**：同上，在「工具 → 宏 → Visual Basic 编辑器」中导入模块，另存为「PowerPoint 加载项 (.ppam)」。

## 开发约定（VBA）

- **只改 `src/VecStamp.bas`**。VBE 按系统 ANSI 代码页导入 `.bas`，所以发布用的 `dist/VecStamp.bas` 必须是纯 ASCII；`build_bas.py` 会把中文字符串转换为 `U("…")`。注释请写英文。
- 所有 PowerPoint / Visio 对象后期绑定（`As Object`），枚举一律用数值常量（例如 `ppShapeFormatSVG = 6`、`ppSaveAsPDF = 32`）。
- Windows API 只能出现在 `#If Mac Then … #Else … #End If` 的 Windows 分支中，并使用 `PtrSafe` / `LongPtr`。macOS 上需要沙盒外权限的操作通过 `AppleScriptTask "VecStamp.scpt", "<handler>", "<参数>"` 完成，处理程序写在 `installer/macos/VecStamp.applescript`。
- 错误处理：在活动的错误处理程序中不要再写 `On Error Resume Next`；需要清理时使用 `On Error GoTo Failed` → `Failed:` → `Resume Cleanup` 模式，确保临时幻灯片、临时演示文稿和剪贴板状态被恢复。
- 功能区回调统一以 `VS_` 开头；新增控件时同时修改 `customUI14.xml` 与 `VecStamp.bas`，再运行 `python build/check_ribbon.py`。
- 设置在 Windows 上保存在注册表 `HKCU\Software\VecStamp`（`WScript.Shell.RegRead/RegWrite`，卸载时可一并删除），Mac 上使用 `GetSetting` / `SaveSetting`；读取时一律做范围限制（例如 DPI 1 – 3000、网格 2 – 200 磅、边距 0 – 500 磅）。
- `check_vba.py` 能发现未声明的变量、参数个数不符和 `Exit Sub/Function` 错配，但**不能代替 VBE 中的「调试 → 编译」**，提交前请在真实的 PowerPoint 中编译一次。
- 中文字符串经 `U("…")` 解码，所以**局部变量不要命名为 `u`**（VBA 不区分大小写，会遮蔽 `U()` 函数）。
- 长任务用 `ProgressBegin` / `ProgressStep` / `ProgressEnd` 包裹，并在循环中检查 `ProgressCancelled()`；入口宏用 `gBusy` 防止重入。进度窗口打开时用 `AskUser` 代替 `MsgBox`，它会先隐藏进度窗口。

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

`XpsToSvg` 把图形粘贴到隐藏母版与背景的临时幻灯片上，用 `PrintOptions.Ranges` 只选这一页，`ExportAsFixedFormat` 导出 XPS（打印意图），然后把内嵌的 `xps2svg.ps1` 写到临时文件夹（`xps2svg-<版本>.ps1`），用

```
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -EncodedCommand …
& ([scriptblock]::Create([IO.File]::ReadAllText('<脚本>', [Text.Encoding]::UTF8))) -Xps … -Out … -Frame 'l,t,w,h' -Slide 'w,h'
```

运行（`-ExecutionPolicy Bypass` 只作用于这一个进程，不修改系统设置）。脚本兼容 Windows PowerShell 5.1 与 PowerShell 7，只用 .NET 自带的 `System.IO.Compression`：

- 读取 FixedDocumentSequence → FixedDocument → 第一页，解析 `Canvas` / `Path` / `Glyphs`、`RenderTransform`、`Clip`、`Opacity`，以及 `StaticResource` 引用的资源字典；
- 画刷：纯色、线性 / 径向渐变（含 `MappingMode` 与 `SpreadMethod`）、`ImageBrush`（转为 `<pattern>`）；
- 字体：按 XPS 规范用文件名中的 GUID 反序字节异或前 32 字节还原混淆字体，读取 `name` / `OS/2` / `cmap` / `hmtx` 得到族名与字形，`Glyphs` 转为 `<text>`（`UnicodeString` + `Indices` 中的字距），字体以 `@font-face` 内嵌（`-NoEmbedFonts` 可关闭）；
- 去掉整页白色背景，按「图形框 ∪ 内容包围盒」裁剪，输出无 BOM 的 UTF-8；成功返回 0。

修改脚本后运行 `python3 tests/test_xps2svg.py`；样例由 `tests/make_xps_fixture.py` 生成。

### VSDX 原生形状

`ExportVsdx` → `VxShape` 递归遍历：组合写成 `Type='Group'`（子形状使用组合内的局部坐标），其余由 `VxNative` 尝试转换，失败时 `VxPicture` 导出 EMF 作为 `ForeignData` 形状。

- 几何：`VxPreset` 按 OOXML `presetShapeDefinitions` 的公式计算常用预设形状（`Adjustments` 以比例给出）；连接线按调整值个数选择 `bentConnector2–5` / `curvedConnector2–3`；自由曲线读取 `Nodes`（`SegmentType` 为曲线的控制点，1 或 3 个节点一段，回到起点时开始新子路径），贝塞尔段写为 `NURBSTo`；其他形状把未旋转副本与自身做「合并形状 → 联合」取轮廓。翻转在几何中处理，旋转写入 `Angle`（`-rotation`）。
- 样式：`FillForegnd` / `FillPattern` / `FillForegndTrans`、`FillGradient` 段、`LineWeight`、`LinePattern`（PowerPoint 虚线样式映射）、`BeginArrow` / `EndArrow` 与大小。
- 文字：`Character` / `Paragraph` 段按文字格式逐段写入，正文中用 `<cp IX>` / `<pp IX>` 切换；`TxtPinX` 等单元格定位文本框，内边距与垂直对齐写入 `LeftMargin` / `VerticalAlign` 等。
- 包：`WriteVsdxPackage` 用模板部件（`src/vsdx`）替换 `{SHAPES}` / `{RELS}`，`[Content_Types].xml` 放在首位。

LibreOffice 端对应 `vecstamp_shapes.py`：自行解析 `EnhancedCustomShapeGeometry`（公式、段命令、`SubViewSize`、`TextFrames`、镜像），不依赖界面渲染。

### 功能区颜色库、进度窗口与反馈

- 「背景颜色」库：前 60 项是主题颜色（`ThemeColorScheme` 的 bg1、tx1、bg2、tx2、accent1–6，按 Office 的 HSL 规则生成 5 级深浅），接着 10 种标准色和最多 10 个最近使用的颜色（设置项 `RecentColors`）。色块是写到临时文件夹的 BMP，用 `LoadPicture` 加载（仅 Windows）。「其他颜色」「取色器」在一个临时的幻灯片外形状上调用 `CommandBars.ExecuteMso("ObjectFillMoreColorsDialog" / "EyedropperFill")`，读取其填充色后删除。
- 进度窗口：Windows 用 `CreateWindowExW` 创建不激活的对话框类窗口（Static + `msctls_progress32`），按 `LOGPIXELSY` 缩放，`GetAsyncKeyState(VK_ESCAPE)` 检测取消；macOS 用辅助脚本的 `progressShow` / `progressHide` 显示后台 `display dialog`。
- 反馈：`DoFeedback` 是「是 / 否 / 取消」消息框：邮件（`SetClipboardText` 复制邮箱，可选 `mailto:`，正文附环境信息）或 GitHub Issue（`VS_REPO/issues/new?title=…&body=…`，UTF-8 百分号编码）。仓库地址在 `VS_REPO` 与 `vecstamp_core.REPO` 中定义。

## LibreOffice 扩展

- 菜单和工具栏项在 `build/make_lo_extension.py` 的 `MENU` / `TOOLBAR` 列表中定义，URL 形如 `service:com.vluckyzhang.vecstamp.Exporter?export:svg`，由 `vecstamp.py` 的 `trigger()` 分发。
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

测试会在临时配置目录中安装扩展、启动无界面的 LibreOffice、绘制示例图形，并对 5 种背景 × 7 种格式逐一校验文件头、SVG 根元素和 VSDX 包结构（原生形状模式检查 `Shape` / `Geometry` / 文字，图片模式检查 `ForeignData`），同时构建「关于」与选择对话框，不会影响你自己的 LibreOffice 配置。

XPS → SVG 转换脚本：

```bash
python3 tests/test_xps2svg.py                     # 需要 pwsh（PowerShell 7）或 Windows PowerShell
```

## README 配图与宣传网站

`python build/make_promo.py` 生成 `assets/promo/*.png`（以及网站用的平底版本 `ribbon-web.png`、`linux-web.png`）。其中：

- 除 `linux.png` 外均为脚本绘制的示意图；`ribbon.png` 按 `customUI14.xml` 的布局绘制并标注「示意图」。
- `assets/screenshots/` 是真实截图：在 Xvfb 虚拟显示器（1600×1000）中以中文界面启动装有扩展的 LibreOffice Impress（GTK3 界面），通过 UNO 绘制示例幻灯片，再用 ImageMagick `import` 截取菜单、导出设置和取色器。界面有变化时请重新截取。
- 赞赏二维码原图在 `assets/donate/*-original.*`；裁切区域在 `make_icons.py` 的 `DONATE_CARDS` 中定义。更换二维码后重新运行 `build.py --promo`，并用扫码工具确认生成的卡片能被识别。

`website/` 是纯静态网站：`index.html` 内联全部 CSS / JS，首屏演示用浏览器把同一段内联 SVG 画到 600×360 的 canvas 上作为「截图」，与矢量 SVG 左右对比（缩放通过修改 `viewBox` 实现，保证矢量侧始终清晰）。部署方法见 [website/README.md](../website/README.md)。

## 发布清单

1. 更新版本号：`src/VecStamp.bas`（`VS_VERSION`）、`libreoffice/extension/pythonpath/vecstamp_core.py`（`VERSION`）、`installer/windows/VecStamp-Setup.ps1`、`installer/macos/setup.sh`、`installer/macos/VecStamp.applescript`、`installer/linux/setup.sh`、README 徽章、`快速开始.txt`（网站版本号由 `make_website.py` 同步）。
2. 在 `CHANGELOG.md` 中记录变更，并更新 `website/index.html` 中的「更新」一节。
3. `python build/build.py --promo`，并运行 `python3 tests/test_libreoffice.py` 与 `python3 tests/test_xps2svg.py`。
4. 检查安装器与脚本：`pwsh` 解析 `VecStamp-Setup.ps1` 和 `src/helpers/xps2svg.ps1`，`shellcheck installer/*/*.sh 安装或卸载-*.sh 安装或卸载-macOS.command`。
5. 在 Windows PowerPoint 中用安装器 `4` 生成 `dist/VecStamp.ppam`，在 PowerPoint（以及可能的话 Mac 与 Visio）中实测一遍。
6. 打包发布：上传项目压缩包，并单独附上 `VecStamp.ppam` 与 `VecStamp-LibreOffice.oxt`；需要时把压缩包放到网站的 `download/` 并修改 `website/config.js`。
