# 第三方声明 / Third-party notices

矢印 VecStamp 本身以 MIT 协议发布（见 [LICENSE](LICENSE)）。以下内容来自或参考了第三方项目。

## ECMA-376 preset shape definitions

`src/presets/presetShapeDefinitions.xml.gz` 是 ECMA-376《Office Open XML File Formats》第 1 部分附带的
预设形状定义文件 `presetShapeDefinitions.xml`（与 ISO/IEC 29500 相同；本副本取自 LibreOffice 源码仓库
`oox/source/drawingml/customshapes/`，内容未修改，仅压缩存放）。`build/make_presets.py` 把其中的几何公式
编译成加载项内的紧凑格式（函数 `PresetSource`），用于生成 VSDX 中的 Visio 原生形状。

`src/presets/presetShapeDefinitions.xml.gz` is the `presetShapeDefinitions.xml` file that accompanies
Part 1 of ECMA-376 *Office Open XML File Formats* (identical to ISO/IEC 29500; this copy was taken unchanged
from the LibreOffice source tree, `oox/source/drawingml/customshapes/`, and is stored compressed).
`build/make_presets.py` compiles its geometry formulas into a compact form inside the add-in.

Copyright © Ecma International. ECMA-376 is published free of charge by Ecma International
(<https://ecma-international.org/publications-and-standards/standards/ecma-376/>); its copyright notice
allows the standard and derivative works that help implement it to be copied and distributed, provided the
Ecma copyright notice is kept. The definitions are provided "as is", without warranty of any kind.

## OpenCC (build tool only)

`build/i18n_tool.py` 和 `build/make_readme_tw.py` 在开发时可以调用 OpenCC（Apache License 2.0，
<https://github.com/BYVoid/OpenCC>，Python 版 opencc-python-reimplemented）生成繁体中文初稿。
OpenCC 不随加载项或扩展分发。

OpenCC (Apache License 2.0) is optionally used at development time to draft the Traditional Chinese
texts; it is not distributed with the add-in or the extension.

## vsdx (Python package)

`src/vsdx/document.xml` 中根样式「No Style」的默认单元格取值，参考了 vsdx 软件包
（<https://github.com/dave-howard/vsdx>，版本 0.6.1）所附、由 Microsoft Visio 保存的空白绘图模板。
VecStamp 没有包含 vsdx 的任何代码。

The default cell values of the root "No Style" style sheet in `src/vsdx/document.xml`
follow the blank drawing template (saved by Microsoft Visio) that ships with the vsdx
package. No vsdx code is included in VecStamp.

```
BSD 3-Clause License

Copyright (c) 2020, Dave Howard
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this
   list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its
   contributors may be used to endorse or promote products derived from
   this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
```
