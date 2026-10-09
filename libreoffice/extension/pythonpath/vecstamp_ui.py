"""VecStamp (Shi Yin) for LibreOffice - dialogs, message boxes and menu actions.

Copyright (c) 2026 vluckyzhang <vluckyzhang@gmail.com>. MIT License.
"""
import os
import platform
import time
import traceback
import urllib.parse

import uno
import unohelper
from com.sun.star.awt import XActionListener, XItemListener
from com.sun.star.datatransfer import DataFlavor, XTransferable

import vecstamp_core as core

TITLE = '矢印 VecStamp'
ICONS = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'icons')
FILESAVE_AUTOEXTENSION = 10          # com.sun.star.ui.dialogs.TemplateDescription
MB_OK, MB_YES_NO, MB_YES_NO_CANCEL = 1, 3, 4   # com.sun.star.awt.MessageBoxButtons
RES_OK, RES_YES, RES_NO, RES_CANCEL = 1, 2, 3, 0


class UI:
    def __init__(self, ctx, doc=None):
        self.ctx = ctx
        self.smgr = ctx.ServiceManager
        self.desktop = self.smgr.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
        self.doc = doc

    # ------------------------------------------------------------ basics
    def parent(self):
        try:
            return self.doc.getCurrentController().getFrame().getContainerWindow()
        except Exception:
            return self.desktop.getCurrentFrame().getContainerWindow()

    def msg(self, text, kind='infobox', buttons=MB_OK):
        tk = self.smgr.createInstanceWithContext('com.sun.star.awt.Toolkit', self.ctx)
        box_type = uno.Enum('com.sun.star.awt.MessageBoxType', kind.upper())
        box = tk.createMessageBox(self.parent(), box_type, buttons, TITLE, text)
        try:
            return box.execute()
        finally:
            box.dispose()

    def info(self, text):
        self.msg(text, 'infobox')

    def warn(self, text):
        self.msg(text, 'warningbox')

    def ask(self, text):
        return self.msg(text, 'querybox', MB_YES_NO) == RES_YES

    def shell_open(self, target):
        """Opens a folder in the file manager or a mailto: link."""
        sx = self.smgr.createInstanceWithContext('com.sun.star.system.SystemShellExecute', self.ctx)
        if '://' in target or target.startswith('mailto:'):
            sx.execute(target, '', 1)                    # URIS_ONLY
        else:
            sx.execute(core.to_url(target), '', 1)

    def progress(self, text, total):
        """LibreOffice's own progress bar in the status bar (or None)."""
        try:
            ind = self.doc.getCurrentController().getFrame().createStatusIndicator()
            ind.start(text, max(int(total), 1))
            return ind
        except Exception:
            return None

    def copy_text(self, text):
        try:
            cb = self.smgr.createInstanceWithContext('com.sun.star.datatransfer.clipboard.SystemClipboard', self.ctx)
            cb.setContents(_TextTransferable(text), None)
            return True
        except Exception:
            return False

    # ------------------------------------------------------------ pickers
    def save_path(self, folder, base, ext):
        fp = self.smgr.createInstanceWithArgumentsAndContext(
            'com.sun.star.ui.dialogs.FilePicker', (FILESAVE_AUTOEXTENSION,), self.ctx)
        fp.setTitle(TITLE + ' - 导出选中图形')
        names = {}
        for e, label in core.FORMATS:
            name = '%s (*.%s)' % (label, e)
            names[name] = e
            fp.appendFilter(name, '*.' + e)
            if e == ext:
                fp.setCurrentFilter(name)
        try:
            fp.setDisplayDirectory(core.to_url(folder))
        except Exception:
            pass
        fp.setDefaultName('%s.%s' % (base, ext))
        if fp.execute() != RES_OK:
            return None, None
        files = fp.getSelectedFiles() if hasattr(fp, 'getSelectedFiles') else fp.getFiles()
        path = uno.fileUrlToSystemPath(files[0])
        typed = os.path.splitext(path)[1].lower().lstrip('.')
        if typed in core.FMT_INDEX:
            return path, typed
        ext = names.get(fp.getCurrentFilter(), ext)
        return path + '.' + ext, ext

    def pick_folder(self, folder, title):
        fp = self.smgr.createInstanceWithContext('com.sun.star.ui.dialogs.FolderPicker', self.ctx)
        fp.setTitle(title)
        try:
            fp.setDisplayDirectory(core.to_url(folder))
        except Exception:
            pass
        if fp.execute() != RES_OK:
            return ''
        return uno.fileUrlToSystemPath(fp.getDirectory())

    def pick_color(self, rgb):
        """System / LibreOffice colour picker. Returns [r, g, b] or None."""
        try:
            cp = self.smgr.createInstanceWithArgumentsAndContext('com.sun.star.cui.ColorPicker',
                                                                 (self.parent(),), self.ctx)
            cp.setPropertyValues((core.pv('Color', core.rgb_int(rgb)),))
            if cp.execute() != RES_OK:
                return None
            for p in cp.getPropertyValues():
                if p.Name == 'Color':
                    c = int(p.Value) & 0xFFFFFF
                    return [(c >> 16) & 255, (c >> 8) & 255, c & 255]
        except Exception:
            pass
        return None

    # ------------------------------------------------------------ actions
    def document(self):
        doc = self.doc or self.desktop.getCurrentComponent()
        if doc is None or not hasattr(doc, 'getDrawPages'):
            self.info('请在 LibreOffice Impress 或 Draw 中使用矢印。')
            return None
        self.doc = doc
        return doc

    def confirm_big_png(self, w, h):
        return self.ask('这张 PNG 约为 %d × %d 像素（约 %d 百万像素），可能很慢或内存不足。仍然继续吗？'
                        % (w, h, round(w * h / 1e6)))

    def export(self, ext=None):
        doc = self.document()
        if doc is None:
            return
        st = core.Settings()
        ext = ext or st.ext
        target = core.selection_target(doc)
        if target is None:
            target = core.current_page_target(doc)
            if target is None:
                self.info('当前页面上没有可导出的图形。')
                return
            if not self.ask('没有选中任何图形。\n是否导出当前页面上的全部图形？'):
                return
        base = core.doc_base_name(doc)
        if target.page_index:
            base += '_第%d页' % target.page_index
        auto = st['naming'] == 1
        separate = st['mode'] == 1 and len(target.shapes) > 1
        files = []
        ind = None
        try:
            if separate:
                folder = core.default_folder(doc, st) if auto else \
                    self.pick_folder(core.initial_folder(doc, st), '选择导出文件夹（每个图形一个文件）')
                if not folder:
                    return
                ind = self.progress('矢印：正在导出 %d 个 %s 文件…' % (len(target.shapes), ext.upper()), len(target.shapes))
                for i, shp in enumerate(target.shapes, 1):
                    name = core.safe_name(shp.getPropertyValue('Name') or 'shape%d' % i)
                    path = core.unique_path(folder, '%s_%s' % (base, name), ext)
                    if ind:
                        ind.setValue(i - 1)
                    core.export_target(self.ctx, doc, core.Target(target.page, [shp], target.page_index),
                                       ext, path, st, self.confirm_big_png)
                    files.append(path)
            else:
                if auto:
                    path = core.unique_path(core.default_folder(doc, st), base + time.strftime('_%Y%m%d_%H%M%S'), ext)
                else:
                    path, ext = self.save_path(core.initial_folder(doc, st), base, ext)
                    if not path:
                        return
                ind = self.progress('矢印：正在导出 %s…' % ext.upper(), 2)
                if ind:
                    ind.setValue(1)
                core.export_target(self.ctx, doc, target, ext, path, st, self.confirm_big_png)
                files.append(path)
        finally:
            if ind:
                ind.end()
        self.report(st, files, auto or separate)

    def batch(self):
        doc = self.document()
        if doc is None:
            return
        st = core.Settings()
        n = doc.getDrawPages().getCount()
        if not self.ask('将把全部 %d 页逐页导出为 %s（每页一个文件，跳过空白页）。继续吗？' % (n, st.ext.upper())):
            return
        folder = core.default_folder(doc, st) if st['naming'] == 1 else \
            self.pick_folder(core.initial_folder(doc, st), '选择逐页导出的文件夹')
        if not folder:
            return
        ind = self.progress('矢印：逐页导出 %s…' % st.ext.upper(), n)

        def step(i, total):
            if ind:
                ind.setText('矢印：逐页导出 %s（第 %d / %d 页）' % (st.ext.upper(), i, total))
                ind.setValue(i - 1)
        try:
            files, skipped = core.export_pages(self.ctx, doc, folder, st.ext, st, confirm_big=self.confirm_big_png,
                                               progress=step)
        finally:
            if ind:
                ind.end()
        if not files:
            self.info('没有可导出的图形。')
            return
        if skipped:
            self.info('已跳过 %d 张空白页。' % skipped)
        self.report(st, files, True)

    def picture(self):
        doc = self.document()
        if doc is None:
            return
        target = core.selection_target(doc)
        if target is None:
            self.info('请先选中图形。')
            return
        core.picture_copy(self.ctx, doc, target)

    def open_folder(self):
        st = core.Settings()
        for d in (st['last_folder'], st['folder']):
            if d and os.path.isdir(d):
                self.shell_open(d)
                return
        self.info('还没有导出过文件。导出一次后，这里会打开最近使用的文件夹。')

    def report(self, st, files, show_list):
        if not files:
            return
        st['last_folder'] = os.path.dirname(files[0])
        st.save()
        if st['open_folder']:
            self.shell_open(os.path.dirname(files[0]))
            return
        if show_list:
            lines = files[:8] + (['……'] if len(files) > 8 else [])
            self.info('已导出 %d 个文件：\n\n%s' % (len(files), '\n'.join(lines)))

    def about(self):
        AboutDialog(self).run()

    def feedback(self):
        choice = ChoiceDialog(self, '反馈', '请选择反馈方式：',
                              [('mail', '发送邮件给作者（%s）' % core.EMAIL),
                               ('issue', '在 GitHub 提交 Issue（需要 GitHub 账号，问题会公开显示）')]).run()
        env = 'VecStamp: %s (LibreOffice)\nOS: %s %s' % (core.VERSION, platform.system(), platform.release())
        if choice == 'mail':
            copied = self.copy_text(core.EMAIL)
            if self.ask(('作者邮箱已复制到剪贴板：' if copied else '作者邮箱：') + core.EMAIL +
                        '\n\n你可以在任何邮箱中新建邮件、粘贴收件人后发送。\n\n要现在打开电脑上的默认邮件程序吗？'):
                q = urllib.parse.urlencode({'subject': '矢印 VecStamp %s 反馈' % core.VERSION,
                                            'body': '\n\n----\n' + env}, quote_via=urllib.parse.quote)
                try:
                    self.shell_open('mailto:%s?%s' % (core.EMAIL, q))
                except Exception:
                    self.info('无法打开邮件程序，请直接发邮件到：' + core.EMAIL)
        elif choice == 'issue':
            body = '**问题描述 / 建议：**\n\n\n**复现步骤：**\n1. \n\n**运行环境：**\n```\n%s\n```\n' % env
            q = urllib.parse.urlencode({'title': '[反馈] ', 'body': body}, quote_via=urllib.parse.quote)
            try:
                self.shell_open('%s/issues/new?%s' % (core.REPO, q))
            except Exception:
                self.info('无法打开浏览器，请手动访问：\n%s/issues' % core.REPO)

    def home(self):
        try:
            self.shell_open(core.REPO)
        except Exception:
            self.info('项目主页：' + core.REPO)

    def help(self):
        self.info('矢印 VecStamp 使用说明\n\n'
                  '① 选中图形，在「矢印」菜单或工具栏中选择导出格式；没有选中时可以导出整页。\n'
                  '② 逐页导出：每页一个文件，跳过空白页。\n'
                  '③ EMF 图片副本：把所选图形复制成矢量图片放在旁边。\n'
                  '④ 导出设置：默认格式、保存方式、PNG 分辨率（最高 3000 DPI）、背景（透明 / 白色 / 米色 / 网格 / 自定义颜色）、'
                  '透明度、网格间距、边距和 VSDX 模式。\n\n'
                  'PDF 的页面与图形同样大小；VSDX 中的图形是 Visio 原生形状，文字可直接编辑。导出进度显示在窗口底部的状态栏。')

    def handle_error(self, exc):
        if isinstance(exc, core.Cancelled):
            return
        self.msg('操作失败：%s\n\n%s' % (exc, traceback.format_exc(limit=3)), 'errorbox')

    # ------------------------------------------------------------ settings dialog
    def settings_dialog(self, execute=True):
        return SettingsDialog(self).run(execute)


class _Listener(unohelper.Base, XActionListener, XItemListener):
    def __init__(self, fn):
        self.fn = fn

    def actionPerformed(self, ev):
        self.fn()

    def itemStateChanged(self, ev):
        self.fn()

    def disposing(self, ev):
        pass


class SettingsDialog:
    W = 260

    def __init__(self, ui):
        self.ui = ui
        self.st = core.Settings()
        smgr, ctx = ui.smgr, ui.ctx
        self.model = smgr.createInstanceWithContext('com.sun.star.awt.UnoControlDialogModel', ctx)
        self.model.Title = TITLE + ' - 导出设置'
        self.model.Width = self.W
        self.model.Height = 262
        self.y = 8
        self._row_list('format', '默认格式', [label for _, label in core.FORMATS], self.st['format'])
        self._row_list('mode', '多个图形', ['合并为一个文件', '每个图形一个文件'], self.st['mode'])
        self._row_list('naming', '保存方式', ['每次选择保存位置', '自动保存到默认文件夹'], self.st['naming'])
        self._label('默认文件夹')
        self._add('folder', 'Edit', 92, self.y, 112, 12, Text=self.st['folder'] or '（演示文稿所在文件夹或桌面）', ReadOnly=True)
        self._add('browse', 'Button', 208, self.y, 44, 12, Label='浏览…')
        self.y += 20
        self._row_list('dpi_index', 'PNG 分辨率', ['150 DPI', '300 DPI', '600 DPI', '自定义'], self.st['dpi_index'], width=70)
        self._add('dpi_custom', 'NumericField', 166, self.y - 20, 50, 12, Value=self.st['dpi_custom'],
                  ValueMin=1, ValueMax=core.DPI_MAX, DecimalAccuracy=0, Spin=True, StrictFormat=True)
        self._add('dpi_unit', 'FixedText', 220, self.y - 18, 30, 10, Label='DPI')
        self._row_list('bg_type', '背景', core.BACKGROUNDS, self.st['bg_type'])
        self._label('自定义颜色')
        self._add('bg_color', 'Edit', 92, self.y, 70, 12, Text=core.color_text(self.st['bg_color']))
        self._add('pick', 'Button', 166, self.y, 44, 12, Label='选色…')
        self._add('swatch', 'FixedText', 214, self.y, 38, 12, Label='', Border=1,
                  BackgroundColor=core.rgb_int(self.st['bg_color']))
        self.y += 20
        self._row_num('bg_alpha', '背景透明度 (%)', self.st['bg_alpha'], 0, 100, 0)
        self._row_num('grid_step', '网格间距 (磅)', self.st['grid_step'], 2, 200, 1)
        self._row_num('margin', '边距 (磅)', self.st['margin'], 0, 500, 1)
        self._add('open_folder', 'CheckBox', 8, self.y, 200, 12, Label='导出完成后打开文件夹',
                  State=1 if self.st['open_folder'] else 0)
        self.y += 16
        self._add('vsdx_native', 'CheckBox', 8, self.y, 244, 12, Label='VSDX 使用 Visio 原生形状（文字可编辑）',
                  State=1 if self.st['vsdx_native'] else 0)
        self.y += 22
        self._add('reset', 'Button', 8, self.y, 60, 14, Label='恢复默认')
        self._add('cancel', 'Button', self.W - 116, self.y, 52, 14, Label='取消', PushButtonType=2)
        self._add('ok', 'Button', self.W - 60, self.y, 52, 14, Label='确定', PushButtonType=1, DefaultButton=True)
        self.model.Height = self.y + 22

        self.dlg = smgr.createInstanceWithContext('com.sun.star.awt.UnoControlDialog', ctx)
        self.dlg.setModel(self.model)
        self.dlg.getControl('browse').addActionListener(_Listener(self.on_browse))
        self.dlg.getControl('pick').addActionListener(_Listener(self.on_pick))
        self.dlg.getControl('reset').addActionListener(_Listener(self.on_reset))
        self.dlg.getControl('dpi_index').addItemListener(_Listener(self.update_enabled))
        self.dlg.getControl('bg_type').addItemListener(_Listener(self.update_enabled))
        self.update_enabled()

    # layout helpers
    def _add(self, name, kind, x, y, w, h, **kw):
        m = self.model.createInstance('com.sun.star.awt.UnoControl%sModel' % kind)
        m.Name = name
        m.PositionX, m.PositionY, m.Width, m.Height = x, y, w, h
        for k, v in kw.items():
            setattr(m, k, v)
        self.model.insertByName(name, m)
        return m

    def _label(self, text):
        self._add('lbl_%d' % self.y, 'FixedText', 8, self.y + 2, 82, 10, Label=text)

    def _row_list(self, name, label, items, sel, width=160):
        self._label(label)
        self._add(name, 'ListBox', 92, self.y, width, 12, Dropdown=True,
                  StringItemList=tuple(items), SelectedItems=(int(sel),))
        self.y += 20

    def _row_num(self, name, label, value, lo, hi, decimals):
        self._label(label)
        self._add(name, 'NumericField', 92, self.y, 60, 12, Value=float(value), ValueMin=float(lo),
                  ValueMax=float(hi), DecimalAccuracy=decimals, Spin=True, StrictFormat=True)
        self.y += 20

    def _sel(self, name):
        items = self.model.getByName(name).SelectedItems
        return int(items[0]) if items else 0

    # events
    def update_enabled(self):
        self.model.getByName('dpi_custom').Enabled = (self._sel('dpi_index') == 3)
        bg = self._sel('bg_type')
        self.model.getByName('grid_step').Enabled = (bg == core.BG_GRID)
        self.model.getByName('bg_alpha').Enabled = (bg != core.BG_NONE)

    def on_browse(self):
        d = self.ui.pick_folder(self.st['folder'] or core.desktop_folder(), '选择“自动保存”使用的默认文件夹')
        if d:
            self.st['folder'] = d
            self.model.getByName('folder').Text = d

    def on_pick(self):
        cur = core.parse_color(self.model.getByName('bg_color').Text) or self.st['bg_color']
        c = self.ui.pick_color(cur)
        if c:
            self._set_color(c)
            self.model.getByName('bg_type').SelectedItems = (core.BG_CUSTOM,)
            self.update_enabled()

    def _set_color(self, c):
        self.model.getByName('bg_color').Text = core.color_text(c)
        self.model.getByName('swatch').BackgroundColor = core.rgb_int(c)

    def on_reset(self):
        self.st.reset()
        self.st['folder'] = ''
        for name in ('format', 'mode', 'naming', 'dpi_index', 'bg_type'):
            self.model.getByName(name).SelectedItems = (int(self.st[name]),)
        for name in ('dpi_custom', 'bg_alpha', 'grid_step', 'margin'):
            self.model.getByName(name).Value = float(self.st[name])
        self._set_color(self.st['bg_color'])
        self.model.getByName('open_folder').State = 0
        self.model.getByName('vsdx_native').State = 1
        self.model.getByName('folder').Text = '（演示文稿所在文件夹或桌面）'
        self.update_enabled()

    def collect(self):
        st = self.st
        for name in ('format', 'mode', 'naming', 'dpi_index', 'bg_type'):
            st[name] = self._sel(name)
        st['dpi_custom'] = int(self.model.getByName('dpi_custom').Value)
        st['bg_alpha'] = int(self.model.getByName('bg_alpha').Value)
        st['grid_step'] = float(self.model.getByName('grid_step').Value)
        st['margin'] = float(self.model.getByName('margin').Value)
        st['open_folder'] = self.model.getByName('open_folder').State == 1
        st['vsdx_native'] = self.model.getByName('vsdx_native').State == 1
        c = core.parse_color(self.model.getByName('bg_color').Text)
        if c is None:
            self.ui.warn('颜色格式示例：255,248,235 或 #FFF8EB。自定义颜色未修改。')
        else:
            st['bg_color'] = c
        st.normalize()
        return st

    def run(self, execute=True):
        if not execute:                      # used by the self-test: build only
            st = self.collect()
            self.dlg.dispose()
            return st
        try:
            tk = self.ui.smgr.createInstanceWithContext('com.sun.star.awt.Toolkit', self.ui.ctx)
            self.dlg.createPeer(tk, self.ui.parent())
            if self.dlg.execute() == RES_OK:
                self.collect().save()
                return True
            return False
        finally:
            self.dlg.dispose()


class _TextTransferable(unohelper.Base, XTransferable):
    def __init__(self, text):
        self.text = text
        f = DataFlavor()
        f.MimeType = 'text/plain;charset=utf-16'
        f.HumanPresentableName = 'Unicode-Text'
        f.DataType = uno.getTypeByName('string')
        self.flavor = f

    def getTransferData(self, flavor):
        return self.text

    def getTransferDataFlavors(self):
        return (self.flavor,)

    def isDataFlavorSupported(self, flavor):
        return flavor.MimeType == self.flavor.MimeType


class _Dialog:
    """Small helper around a programmatically built UNO dialog."""

    def __init__(self, ui, title, width):
        self.ui = ui
        self.model = ui.smgr.createInstanceWithContext('com.sun.star.awt.UnoControlDialogModel', ui.ctx)
        self.model.Title = title
        self.model.Width = width
        self.model.Height = 100
        self.dlg = None

    def add(self, name, kind, x, y, w, h, **kw):
        m = self.model.createInstance('com.sun.star.awt.UnoControl%sModel' % kind)
        m.Name = name
        m.PositionX, m.PositionY, m.Width, m.Height = x, y, w, h
        for k, v in kw.items():
            setattr(m, k, v)
        self.model.insertByName(name, m)
        return m

    def create(self):
        self.dlg = self.ui.smgr.createInstanceWithContext('com.sun.star.awt.UnoControlDialog', self.ui.ctx)
        self.dlg.setModel(self.model)
        tk = self.ui.smgr.createInstanceWithContext('com.sun.star.awt.Toolkit', self.ui.ctx)
        self.dlg.createPeer(tk, self.ui.parent())
        return self.dlg


class ChoiceDialog(_Dialog):
    """A question with one button per choice. run() returns the chosen key or None."""

    def __init__(self, ui, title, text, choices):
        super().__init__(ui, TITLE + ' - ' + title, 250)
        self.choices = choices
        self.result = None
        self.add('text', 'FixedText', 10, 8, 230, 12, Label=text)
        y = 26
        for key, label in choices:
            self.add('c_' + key, 'Button', 10, y, 230, 18, Label=label, Align=0)
            y += 22
        self.add('cancel', 'Button', 180, y + 4, 60, 14, Label='取消', PushButtonType=2)
        self.model.Height = y + 24

    def run(self):
        dlg = self.create()
        try:
            for key, _ in self.choices:
                dlg.getControl('c_' + key).addActionListener(_Listener(lambda k=key: self._pick(k)))
            dlg.execute()
            return self.result
        finally:
            dlg.dispose()

    def _pick(self, key):
        self.result = key
        self.dlg.endExecute()


class AboutDialog(_Dialog):
    """About VecStamp: version, licence, the free / open-source promise and the donation QR codes."""

    def __init__(self, ui):
        super().__init__(ui, '关于 ' + TITLE, 300)
        url = lambda name: uno.systemPathToFileUrl(os.path.join(ICONS, name))
        self.add('logo', 'ImageControl', 10, 10, 32, 32, ImageURL=url('logo_64.png'), ScaleImage=True, Border=0)
        self.add('name', 'FixedText', 50, 12, 240, 14, Label='矢印 VecStamp  v%s（LibreOffice 版）' % core.VERSION,
                 FontHeight=13, FontWeight=150)
        self.add('tag', 'FixedText', 50, 28, 240, 10, Label='选中即印 —— 图形一键导出 EMF / SVG / PDF / VSDX / PNG')
        self.add('info', 'FixedText', 10, 48, 280, 40, MultiLine=True,
                 Label='作者：%s    邮箱：%s\n主页：%s\n开源协议：MIT License    Copyright (c) 2026 %s'
                       % (core.AUTHOR, core.EMAIL, core.REPO, core.AUTHOR))
        self.add('free', 'FixedText', 10, 90, 280, 32, MultiLine=True,
                 Label='矢印承诺永久开源、永久免费：没有收费功能，没有广告，不联网收集任何信息，全部源码在 GitHub 公开。'
                       '如果它为你节省了时间，欢迎扫码请作者喝杯咖啡（完全自愿，不影响任何功能）。')
        self.add('qr1', 'ImageControl', 40, 126, 100, 126, ImageURL=url('donate_alipay.png'), ScaleImage=True,
                 ScaleMode=1, Border=0)
        self.add('qr2', 'ImageControl', 160, 126, 100, 126, ImageURL=url('donate_wechat.png'), ScaleImage=True,
                 ScaleMode=1, Border=0)
        self.add('home', 'Button', 10, 262, 70, 14, Label='项目主页')
        self.add('fb', 'Button', 86, 262, 70, 14, Label='反馈…')
        self.add('ok', 'Button', 230, 262, 60, 14, Label='确定', PushButtonType=1, DefaultButton=True)
        self.model.Height = 284

    def run(self, execute=True):
        dlg = self.create() if execute else None
        if not execute:
            return True
        try:
            dlg.getControl('home').addActionListener(_Listener(self.ui.home))
            dlg.getControl('fb').addActionListener(_Listener(self._feedback))
            dlg.execute()
            return True
        finally:
            dlg.dispose()

    def _feedback(self):
        self.dlg.endExecute()
        self.ui.feedback()
