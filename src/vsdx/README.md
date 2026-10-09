# VSDX template parts

Building blocks of the `.vsdx` file VecStamp writes when Visio is not available:
one page whose only shape is the exported drawing embedded as an EMF picture
(`Type='Foreign'`, `ForeignType='EnhMetaFile'`). In Visio you can right-click the
picture and choose Group > Ungroup to turn it into editable Visio shapes.

Placeholders such as `{PW}` (page width, inches) are filled in at export time.
`map.txt` says where each file goes inside the package. The same parts are used by
the PowerPoint add-in (embedded into the VBA module by `build/build_bas.py`) and by
the LibreOffice extension (`libreoffice/extension/pythonpath/vecstamp_vsdx.py`).

`document.xml` holds only the root style sheet ("No Style"); its default cell values
follow the blank drawing (saved by Microsoft Visio) bundled with the BSD-3-Clause licensed
[vsdx](https://github.com/dave-howard/vsdx) package - see `THIRD_PARTY_NOTICES.md`.
