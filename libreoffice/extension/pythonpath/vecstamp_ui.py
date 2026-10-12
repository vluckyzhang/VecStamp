"""VecStamp (Shi Yin) for LibreOffice - dialogs, message boxes and menu actions.

Copyright (c) 2026 vluckyzhang <vluckyzhang@gmail.com>. MIT License.
"""
import colorsys
import os
import platform
import time
import traceback
import urllib.parse

import uno
import unohelper
from com.sun.star.awt import XActionListener, XItemListener, XKeyHandler, XMouseListener, XTextListener, \
    XTopWindowListener
from com.sun.star.datatransfer import DataFlavor, XTransferable

import vecstamp_core as core
import vecstamp_i18n as i18n
from vecstamp_i18n import T

ICONS = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'icons')
FILESAVE_AUTOEXTENSION = 10          # com.sun.star.ui.dialogs.TemplateDescription
MB_OK, MB_YES_NO, MB_YES_NO_CANCEL = 1, 3, 4   # com.sun.star.awt.MessageBoxButtons
RES_OK, RES_YES, RES_NO, RES_CANCEL = 1, 2, 3, 0
KEY_ESCAPE = 1281                    # com.sun.star.awt.Key.ESCAPE
GRAY = 0x666666


def title():
    return T('app.title')


def icon_url(name):
    return uno.systemPathToFileUrl(os.path.join(ICONS, name))


class UI:
    def __init__(self, ctx, doc=None):
        self.ctx = ctx
        self.smgr = ctx.ServiceManager
        self.desktop = self.smgr.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
        self.doc = doc
        core.Settings()                      # applies the interface language

    # ------------------------------------------------------------ basics
    def parent(self):
        try:
            return self.doc.getCurrentController().getFrame().getContainerWindow()
        except Exception:
            return self.desktop.getCurrentFrame().getContainerWindow()

    def toolkit(self):
        return self.smgr.createInstanceWithContext('com.sun.star.awt.Toolkit', self.ctx)

    def msg(self, text, kind='infobox', buttons=MB_OK):
        box_type = uno.Enum('com.sun.star.awt.MessageBoxType', kind.upper())
        box = self.toolkit().createMessageBox(self.parent(), box_type, buttons, title(), text)
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

    def status_indicator(self, text):
        """LibreOffice's own progress bar in the status bar (0..100), or None."""
        try:
            ind = (self.doc or self.desktop.getCurrentComponent()).getCurrentController().getFrame() \
                .createStatusIndicator()
            ind.start(text, 100)
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
        fp.setTitle(title() + ' - ' + T('dlg.saveTitle'))
        names = {}
        for e in core.FORMATS:
            name = '%s (*.%s)' % (core.format_label(e), e)
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

    def pick_folder(self, folder, caption):
        fp = self.smgr.createInstanceWithContext('com.sun.star.ui.dialogs.FolderPicker', self.ctx)
        fp.setTitle(caption)
        try:
            fp.setDisplayDirectory(core.to_url(folder))
        except Exception:
            pass
        if fp.execute() != RES_OK:
            return ''
        return uno.fileUrlToSystemPath(fp.getDirectory())

    def pick_color(self, rgb, recent=()):
        """LibreOffice's colour picker; VecStamp's own colour dialog when it is not available.

        Returns [r, g, b] or None (cancelled)."""
        try:
            cp = self.smgr.createInstanceWithArgumentsAndContext('com.sun.star.cui.ColorPicker',
                                                                 (self.parent(),), self.ctx)
        except Exception:
            cp = None
        if cp is not None:
            try:
                cp.setPropertyValues((core.pv('Color', core.rgb_int(rgb)),))
                if cp.execute() != RES_OK:
                    return None
                for p in cp.getPropertyValues():
                    if p.Name == 'Color':
                        return core.int_rgb(p.Value)
                return None
            except Exception:
                pass
        return ColorDialog(self, rgb, recent).run()

    # ------------------------------------------------------------ actions
    def document(self):
        doc = self.doc or self.desktop.getCurrentComponent()
        if doc is None or not hasattr(doc, 'getDrawPages'):
            self.info(T('msg.useImpress'))
            return None
        self.doc = doc
        return doc

    def confirm_big_png(self, w, h):
        return self.ask(T('msg.bigPng', w, h, round(w * h / 1e6)))

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
                self.info(T('msg.noShapesPage'))
                return
            if not self.ask(T('msg.askWholePage')):
                return
        base = core.doc_base_name(doc)
        if target.page_index:
            base += '_' + T('file.page', target.page_index)
        auto = st['naming'] == 1
        separate = st['mode'] == 1 and len(target.shapes) > 1
        files = []
        try:
            if separate:
                folder = core.default_folder(doc, st) if auto else \
                    self.pick_folder(core.initial_folder(doc, st), T('dlg.pickFolderEach'))
                if not folder:
                    return
                n = len(target.shapes)
                with Progress(self, T('prog.exportN', n, ext.upper()), n) as prog:
                    for i, shp in enumerate(target.shapes, 1):
                        name = core.safe_name(shp.getPropertyValue('Name') or 'shape%d' % i)
                        prog.item(i, T('prog.item', i, n, name))
                        path = core.unique_path(folder, '%s_%s' % (base, name), ext)
                        core.export_target(self.ctx, doc, core.Target(target.page, [shp], target.page_index),
                                           ext, path, st, self.confirm_big_png, prog)
                        files.append(path)
            else:
                if auto:
                    path = core.unique_path(core.default_folder(doc, st), base + time.strftime('_%Y%m%d_%H%M%S'), ext)
                else:
                    path, ext = self.save_path(core.initial_folder(doc, st), base, ext)
                    if not path:
                        return
                with Progress(self, T('prog.export1', ext.upper()), 1) as prog:
                    prog.item(1, os.path.basename(path))
                    core.export_target(self.ctx, doc, target, ext, path, st, self.confirm_big_png, prog)
                files.append(path)
        except core.Cancelled as exc:
            self.cancelled(exc, st, files)
            return
        self.report(st, files, auto or separate)

    def batch(self):
        doc = self.document()
        if doc is None:
            return
        st = core.Settings()
        n = doc.getDrawPages().getCount()
        if not self.ask(T('msg.askBatch', n, st.ext.upper())):
            return
        folder = core.default_folder(doc, st) if st['naming'] == 1 else \
            self.pick_folder(core.initial_folder(doc, st), T('dlg.pickFolderBatch'))
        if not folder:
            return
        done = []                            # files written so far, also when cancelled
        try:
            with Progress(self, T('prog.batch', st.ext.upper()), n) as prog:
                files, skipped = core.export_pages(
                    self.ctx, doc, folder, st.ext, st, confirm_big=self.confirm_big_png,
                    on_page=lambda i, total: prog.item(i, T('prog.page', i, total)), progress=prog, files=done)
        except core.Cancelled as exc:
            self.cancelled(exc, st, done)
            return
        if not files:
            self.info(T('msg.nothing'))
            return
        if skipped:
            self.info(T('msg.skipped', skipped))
        self.report(st, files, True)

    def cancelled(self, exc, st, files):
        if exc.quiet and not files:
            return
        if files:
            st['last_folder'] = os.path.dirname(files[0])
            st.save()
            lines = files[:8] + (['……'] if len(files) > 8 else [])
            self.info(T('msg.cancelledPartial', len(files)) + '\n\n' + '\n'.join(lines))
        else:
            self.info(T('msg.cancelled'))

    def picture(self):
        doc = self.document()
        if doc is None:
            return
        target = core.selection_target(doc)
        if target is None:
            self.info(T('msg.selectFirst'))
            return
        core.picture_copy(self.ctx, doc, target)

    def open_folder(self):
        st = core.Settings()
        for d in (st['last_folder'], st['folder']):
            if d and os.path.isdir(d):
                self.shell_open(d)
                return
        self.info(T('msg.noFolderYet'))

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
            self.info(T('msg.exported', len(files)) + '\n\n' + '\n'.join(lines))

    def about(self):
        AboutDialog(self).run()

    def language(self):
        """Choose the interface language; it applies at once to every following dialog and message."""
        choices = [(code, name + ('  ✓' if code == i18n.lang() else ''))
                   for code, name in zip(i18n.LANGS, i18n.NAMES)]
        code = ChoiceDialog(self, T('lang.title'), T('lang.prompt'), choices).run()
        if code:
            st = core.Settings()
            st['lang'] = code
            st.save()
            self.info(T('lang.done', i18n.NAMES[i18n.lang_index(code)]))

    def feedback(self):
        choice = ChoiceDialog(self, T('fb.title'), T('fb.prompt'),
                              [('mail', T('fb.mail', core.EMAIL)), ('issue', T('fb.issue'))]).run()
        env = 'VecStamp: %s (LibreOffice)\nOS: %s %s\nLanguage: %s' % (core.VERSION, platform.system(),
                                                                      platform.release(), i18n.lang())
        if choice == 'mail':
            copied = self.copy_text(core.EMAIL)
            if self.ask(T('fb.mailCopied' if copied else 'fb.mailPlain', core.EMAIL)):
                q = urllib.parse.urlencode({'subject': T('fb.subject', core.VERSION),
                                            'body': '\n\n----\n' + env}, quote_via=urllib.parse.quote)
                try:
                    self.shell_open('mailto:%s?%s' % (core.EMAIL, q))
                except Exception:
                    self.info(T('fb.noMailApp', core.EMAIL))
        elif choice == 'issue':
            q = urllib.parse.urlencode({'title': T('fb.issueTitle'), 'body': T('fb.issueBody', env)},
                                       quote_via=urllib.parse.quote)
            try:
                self.shell_open('%s/issues/new?%s' % (core.REPO, q))
            except Exception:
                self.info(T('fb.noBrowser', core.REPO + '/issues'))

    def home(self):
        try:
            self.shell_open(core.REPO)
        except Exception:
            self.info(T('msg.home', core.REPO))

    def help(self):
        self.info(T('help.lo', core.DPI_MAX))

    def handle_error(self, exc):
        if isinstance(exc, core.Cancelled):
            return
        self.msg(T('msg.failed', exc, traceback.format_exc(limit=3)), 'errorbox')

    # ------------------------------------------------------------ settings dialog
    def settings_dialog(self, execute=True):
        return SettingsDialog(self).run(execute)


# ---------------------------------------------------------------------------- listeners
class _Listener(unohelper.Base, XActionListener, XItemListener, XTextListener):
    def __init__(self, fn):
        self.fn = fn

    def actionPerformed(self, ev):
        self.fn()

    def itemStateChanged(self, ev):
        self.fn()

    def textChanged(self, ev):
        self.fn()

    def disposing(self, ev):
        pass


class _Click(unohelper.Base, XMouseListener):
    """Mouse clicks on a control: fn(double_click)."""

    def __init__(self, fn):
        self.fn = fn

    def mousePressed(self, ev):
        if ev.Buttons == 1:
            self.fn(ev.ClickCount >= 2)

    def mouseReleased(self, ev):
        pass

    def mouseEntered(self, ev):
        pass

    def mouseExited(self, ev):
        pass

    def disposing(self, ev):
        pass


class _EscKey(unohelper.Base, XKeyHandler):
    """Application-wide Esc while the progress window is open."""

    def __init__(self, fn):
        self.fn = fn

    def keyPressed(self, ev):
        if ev.KeyCode == KEY_ESCAPE:
            self.fn()
            return True
        return False

    def keyReleased(self, ev):
        return ev.KeyCode == KEY_ESCAPE

    def disposing(self, ev):
        pass


class _Closing(unohelper.Base, XTopWindowListener):
    def __init__(self, fn):
        self.fn = fn

    def windowClosing(self, ev):
        self.fn()

    def windowOpened(self, ev):
        pass

    def windowClosed(self, ev):
        pass

    def windowMinimized(self, ev):
        pass

    def windowNormalized(self, ev):
        pass

    def windowActivated(self, ev):
        pass

    def windowDeactivated(self, ev):
        pass

    def disposing(self, ev):
        pass


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


# ---------------------------------------------------------------------------- dialogs
class _Dialog:
    """Small helper around a programmatically built UNO dialog."""

    def __init__(self, ui, caption, width):
        self.ui = ui
        self.model = ui.smgr.createInstanceWithContext('com.sun.star.awt.UnoControlDialogModel', ui.ctx)
        self.model.Title = caption
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
        self.dlg.createPeer(self.ui.toolkit(), self.ui.parent())
        return self.dlg

    def get(self, name):
        return self.model.getByName(name)


class Progress:
    """Export progress: a small window with the task, the current file or page, the current step,
    a bar with the percentage, the elapsed time and a Cancel button (Esc also cancels); LibreOffice's
    status bar shows the same. Without a window (e.g. headless) only the status bar is used.

    item(i, label) starts item i of n; calling the object with (fraction of the item, step text)
    updates the display and raises core.Cancelled once the user has cancelled."""

    def __init__(self, ui, heading, items=1):
        self.ui, self.heading, self.items = ui, heading, max(int(items), 1)
        self.index, self.label, self.detail = 0, '', ''
        self.cancelled = False
        self.t0 = self.last = time.time()
        self.win = self.tk = self.keys = None
        self.ind = ui.status_indicator(title() + T('sep') + heading)
        try:
            self._open()
        except Exception:
            self._close_window()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.end()
        return False

    def _open(self):
        d = _Dialog(self.ui, title(), 260)
        d.add('head', 'FixedText', 10, 8, 240, 11, Label=self.heading, FontWeight=150)
        d.add('item', 'FixedText', 10, 22, 240, 10, Label='', NoLabel=True)
        d.add('bar', 'ProgressBar', 10, 36, 200, 10, ProgressValueMin=0, ProgressValueMax=1000, ProgressValue=0)
        d.add('pct', 'FixedText', 214, 36, 36, 10, Label='0%', Align=2, FontWeight=150, NoLabel=True)
        d.add('detail', 'FixedText', 10, 50, 240, 10, Label='', NoLabel=True)
        d.add('foot', 'FixedText', 10, 68, 180, 10, Label=T('prog.hint'), TextColor=GRAY, NoLabel=True)
        d.add('cancel', 'Button', 196, 64, 54, 14, Label=T('btn.cancel'))
        d.model.Height = 86
        dlg = d.create()
        dlg.getControl('cancel').addActionListener(_Listener(self.cancel))
        try:
            dlg.addTopWindowListener(_Closing(self.cancel))
        except Exception:
            pass
        dlg.setVisible(True)
        self.win = d
        self.tk = self.ui.toolkit()
        try:
            self.keys = _EscKey(self.cancel)
            self.tk.addKeyHandler(self.keys)
        except Exception:
            self.keys = None
        self._pump()

    def _pump(self):
        if self.tk is None:
            return
        for method in ('processEventsToIdle', 'reschedule'):
            try:
                getattr(self.tk, method)()
                return
            except Exception:
                pass

    def cancel(self):
        if not self.cancelled:
            self.cancelled = True
            self.detail = T('prog.cancelling')
            self._show(force=True)

    def item(self, i, label):
        self.index, self.label = max(int(i) - 1, 0), label
        self(0.0, '')

    def __call__(self, frac, detail=''):
        frac = min(max(float(frac), 0.0), 1.0)
        self.frac = (self.index + frac) / self.items
        changed = detail != self.detail
        if not self.cancelled:
            self.detail = detail
        self._show(force=changed or frac >= 1.0)
        if self.cancelled:
            raise core.Cancelled()

    def percent(self):
        return int(getattr(self, 'frac', 0.0) * 100 + 1e-9)

    def _show(self, force=False):
        now = time.time()
        if not force and now - self.last < 0.1:
            return
        self.last = now
        pct = self.percent()
        elapsed = int(now - self.t0)
        if self.win is not None:
            try:
                self.win.get('item').Label = self.label
                self.win.get('bar').ProgressValue = int(getattr(self, 'frac', 0.0) * 1000)
                self.win.get('pct').Label = '%d%%' % pct
                self.win.get('detail').Label = self.detail
                self.win.get('foot').Label = T('prog.elapsed', elapsed) + '  ·  ' + T('prog.hint')
            except Exception:
                pass
        if self.ind is not None:
            try:
                parts = [p for p in (self.label, self.detail) if p]
                self.ind.setText('%s%s%s — %s  %d%%' % (title(), T('sep'), self.heading, T('sep').join(parts), pct))
                self.ind.setValue(pct)
            except Exception:
                pass
        self._pump()

    def _close_window(self):
        if self.keys is not None and self.tk is not None:
            try:
                self.tk.removeKeyHandler(self.keys)
            except Exception:
                pass
        self.keys = None
        if self.win is not None and self.win.dlg is not None:
            try:
                self.win.dlg.setVisible(False)
                self.win.dlg.dispose()
            except Exception:
                pass
        self.win = None

    def end(self):
        self._close_window()
        if self.ind is not None:
            try:
                self.ind.end()
            except Exception:
                pass
            self.ind = None


def palette_rows():
    """The 6 rows of the theme part of the colour palette (base colours, then tints and shades)."""
    rows = [[core.int_rgb(c) for c in core.THEME]]
    variants = []
    for c in core.THEME:
        r, g, b = [x / 255.0 for x in core.int_rgb(c)]
        h, l, s = colorsys.rgb_to_hls(r, g, b)
        if l > 0.8:
            mods = [('d', 0.05), ('d', 0.15), ('d', 0.25), ('d', 0.35), ('d', 0.5)]
        elif l < 0.2:
            mods = [('l', 0.5), ('l', 0.35), ('l', 0.25), ('l', 0.15), ('l', 0.05)]
        else:
            mods = [('l', 0.8), ('l', 0.6), ('l', 0.4), ('d', 0.25), ('d', 0.5)]
        col = []
        for kind, t in mods:
            l2 = l + (1 - l) * t if kind == 'l' else l * (1 - t)
            rr, gg, bb = colorsys.hls_to_rgb(h, l2, s)
            col.append([int(round(v * 255)) for v in (rr, gg, bb)])
        variants.append(col)
    for k in range(5):
        rows.append([variants[i][k] for i in range(len(core.THEME))])
    return rows


class ColorDialog(_Dialog):
    """VecStamp's own colour dialog (used when LibreOffice's colour picker is not available):
    theme colours with tints and shades, standard colours, recent colours, and a colour value
    field (R,G,B or #RRGGBB) with a preview. run() returns [r, g, b] or None."""

    SW, SH, DX, DY = 14, 10, 18, 12

    def __init__(self, ui, rgb, recent=()):
        super().__init__(ui, title() + ' - ' + T('col.title'), 196)
        self.color = list(rgb)
        self.result = None
        self.swatches = {}
        y = 6
        self.add('lt', 'FixedText', 8, y, 180, 10, Label=T('col.theme'))
        y += 12
        for r, row in enumerate(palette_rows()):
            for c, rgbv in enumerate(row):
                self._swatch('t%d_%d' % (r, c), 8 + c * self.DX, y + r * self.DY + (2 if r else 0), rgbv)
        y += 6 * self.DY + 6
        self.add('ls', 'FixedText', 8, y, 180, 10, Label=T('col.standard'))
        y += 12
        for c, v in enumerate(core.STANDARD):
            self._swatch('s%d' % c, 8 + c * self.DX, y, core.int_rgb(v))
        y += self.DY + 4
        recent = [list(x) for x in recent][:10]
        if recent:
            self.add('lr', 'FixedText', 8, y, 180, 10, Label=T('col.recent'))
            y += 12
            for c, v in enumerate(recent):
                self._swatch('r%d' % c, 8 + c * self.DX, y, v)
            y += self.DY + 4
        self.add('lv', 'FixedText', 8, y + 2, 50, 10, Label=T('col.value'))
        self.add('value', 'Edit', 60, y, 74, 12, Text=core.color_text(self.color))
        self.add('preview', 'FixedText', 140, y, 48, 12, Label='', Border=1, BackgroundColor=core.rgb_int(self.color))
        y += 15
        self.add('hint', 'FixedText', 8, y, 180, 20, Label=T('col.hint'), TextColor=GRAY, MultiLine=True)
        y += 24
        self.add('cancel', 'Button', 196 - 8 - 52, y, 52, 14, Label=T('btn.cancel'), PushButtonType=2)
        self.add('ok', 'Button', 196 - 8 - 52 - 56, y, 52, 14, Label=T('btn.ok'), DefaultButton=True)
        self.model.Height = y + 22

    def _swatch(self, name, x, y, rgb):
        self.swatches[name] = list(rgb)
        self.add(name, 'FixedText', x, y, self.SW, self.SH, Label='', Border=1, BackgroundColor=core.rgb_int(rgb),
                 HelpText='#%02X%02X%02X  (%d, %d, %d)' % (tuple(rgb) * 2))

    def _set(self, rgb):
        self.color = list(rgb)
        self.get('value').Text = core.color_text(rgb)
        self.get('preview').BackgroundColor = core.rgb_int(rgb)

    def _typed(self):
        c = core.parse_color(self.get('value').Text)
        if c is not None:
            self.color = c
            self.get('preview').BackgroundColor = core.rgb_int(c)

    def _pick(self, name, double):
        self._set(self.swatches[name])
        if double:
            self._ok()

    def _ok(self):
        c = core.parse_color(self.get('value').Text)
        if c is None:
            self.ui.warn(T('col.bad'))
            return
        self.result = c
        self.dlg.endExecute()

    def run(self, execute=True):
        if not execute:
            return None
        dlg = self.create()
        try:
            for name in self.swatches:
                dlg.getControl(name).addMouseListener(_Click(lambda dbl, n=name: self._pick(n, dbl)))
            dlg.getControl('value').addTextListener(_Listener(self._typed))
            dlg.getControl('ok').addActionListener(_Listener(self._ok))
            dlg.execute()
            return self.result
        finally:
            dlg.dispose()


class SettingsDialog(_Dialog):
    W = 290

    def __init__(self, ui, settings=None):
        st = settings if settings is not None else core.Settings()     # (loading applies the saved language)
        super().__init__(ui, title() + ' - ' + T('set.title'), self.W)
        self.st = st
        self.y = 8
        self._row_list('lang', T('set.lang'), list(i18n.NAMES), i18n.lang_index(self.st['lang']))
        self._row_list('format', T('set.format'), [core.format_label(e) for e in core.FORMATS], self.st['format'])
        self._row_list('mode', T('set.mode'), [T('mode.merge'), T('mode.split')], self.st['mode'])
        self._row_list('naming', T('set.naming'), [T('naming.ask'), T('naming.auto')], self.st['naming'])
        self._label(T('set.folder'))
        self.add('folder', 'Edit', 110, self.y, 122, 12, Text=self.st['folder'] or T('set.folderNone'), ReadOnly=True)
        self.add('browse', 'Button', 236, self.y, 46, 12, Label=T('btn.browse'))
        self.y += 20
        self._row_list('dpi_index', T('set.dpi'), ['150 DPI', '300 DPI', '600 DPI', T('dpi.custom')],
                       self.st['dpi_index'], width=70)
        self.add('dpi_custom', 'NumericField', 184, self.y - 20, 50, 12, Value=self.st['dpi_custom'],
                 ValueMin=1, ValueMax=core.DPI_MAX, DecimalAccuracy=0, Spin=True, StrictFormat=True)
        self.add('dpi_unit', 'FixedText', 238, self.y - 18, 30, 10, Label='DPI')
        # background: colour preview square, label, list (the last entry opens the colour picker), change button
        self.add('swatch', 'ImageControl', 8, self.y + 1, 10, 10, Border=1, ScaleImage=True)
        self.add('lbl_bg', 'FixedText', 22, self.y + 2, 86, 10, Label=T('set.bg'))
        self.add('bg_type', 'ListBox', 110, self.y, 122, 12, Dropdown=True,
                 StringItemList=tuple(core.backgrounds()[:4] + [T('bg.customMenu')]),
                 SelectedItems=(int(self.st['bg_type']),))
        self.add('pick', 'Button', 236, self.y, 46, 12, Label=T('btn.color'), HelpText=T('set.pickTip'))
        self.y += 20
        self.bg_color = list(self.st['bg_color'])
        self.bg_last = int(self.st['bg_type'])
        self._row_num('bg_alpha', T('set.alpha'), self.st['bg_alpha'], 0, 100, 0)
        self._row_num('grid_step', T('set.grid'), self.st['grid_step'], 2, 200, 1)
        self._row_num('margin', T('set.margin'), self.st['margin'], 0, 500, 1)
        self.add('open_folder', 'CheckBox', 8, self.y, 274, 12, Label=T('set.openFolder'),
                 State=1 if self.st['open_folder'] else 0)
        self.y += 16
        self.add('vsdx_native', 'CheckBox', 8, self.y, 274, 12, Label=T('set.vsdxNative'),
                 State=1 if self.st['vsdx_native'] else 0)
        self.y += 22
        self.add('reset', 'Button', 8, self.y, 64, 14, Label=T('btn.reset'))
        self.add('cancel', 'Button', self.W - 116, self.y, 52, 14, Label=T('btn.cancel'), PushButtonType=2)
        self.add('ok', 'Button', self.W - 60, self.y, 52, 14, Label=T('btn.ok'), PushButtonType=1, DefaultButton=True)
        self.model.Height = self.y + 22
        self.update_swatch()

    # layout helpers
    def _label(self, text):
        self.add('lbl_%d' % self.y, 'FixedText', 8, self.y + 2, 100, 10, Label=text)

    def _row_list(self, name, label, items, sel, width=170):
        self._label(label)
        self.add(name, 'ListBox', 110, self.y, width, 12, Dropdown=True,
                 StringItemList=tuple(items), SelectedItems=(int(sel),))
        self.y += 20

    def _row_num(self, name, label, value, lo, hi, decimals):
        self._label(label)
        self.add(name, 'NumericField', 110, self.y, 60, 12, Value=float(value), ValueMin=float(lo),
                 ValueMax=float(hi), DecimalAccuracy=decimals, Spin=True, StrictFormat=True)
        self.y += 20

    def _sel(self, name):
        items = self.get(name).SelectedItems
        return int(items[0]) if items else 0

    # events
    def update_enabled(self):
        self.get('dpi_custom').Enabled = (self._sel('dpi_index') == 3)
        bg = self._sel('bg_type')
        self.get('grid_step').Enabled = (bg == core.BG_GRID)
        self.get('bg_alpha').Enabled = (bg != core.BG_NONE)
        self.update_swatch()

    def update_swatch(self):
        """The square in front of 'Background' previews the chosen background."""
        bg = self._sel('bg_type')
        m = self.get('swatch')
        if bg in (core.BG_NONE, core.BG_GRID):
            m.ImageURL = icon_url('bg_none.png' if bg == core.BG_NONE else 'bg_grid.png')
        else:
            m.ImageURL = ''
            m.BackgroundColor = core.rgb_int(self.bg_color if bg == core.BG_CUSTOM else
                                             ([255, 255, 255] if bg == core.BG_WHITE else core.BEIGE))
        m.HelpText = T('set.bgTip', core.backgrounds()[bg], core.color_text(self.bg_color))

    def on_bg(self):
        if self._sel('bg_type') == core.BG_CUSTOM and self.bg_last != core.BG_CUSTOM:
            if not self.on_pick():
                self.get('bg_type').SelectedItems = (self.bg_last,)
        self.bg_last = self._sel('bg_type')
        self.update_enabled()

    def on_browse(self):
        d = self.ui.pick_folder(self.st['folder'] or core.desktop_folder(), T('set.pickDefaultFolder'))
        if d:
            self.st['folder'] = d
            self.get('folder').Text = d

    def on_pick(self):
        c = self.ui.pick_color(self.bg_color, self.st['recent_colors'])
        if not c:
            return False
        self.bg_color = list(c)
        self.st.add_recent(c)
        self.get('bg_type').SelectedItems = (core.BG_CUSTOM,)
        self.bg_last = core.BG_CUSTOM
        self.update_enabled()
        return True

    def on_reset(self):
        self.st.reset()
        self.st['folder'] = ''
        for name in ('format', 'mode', 'naming', 'dpi_index', 'bg_type'):
            self.get(name).SelectedItems = (int(self.st[name]),)
        for name in ('dpi_custom', 'bg_alpha', 'grid_step', 'margin'):
            self.get(name).Value = float(self.st[name])
        self.bg_color = list(self.st['bg_color'])
        self.bg_last = int(self.st['bg_type'])
        self.get('open_folder').State = 0
        self.get('vsdx_native').State = 1
        self.get('folder').Text = T('set.folderNone')
        self.update_enabled()

    def collect(self):
        st = self.st
        for name in ('format', 'mode', 'naming', 'dpi_index', 'bg_type'):
            st[name] = self._sel(name)
        st['lang'] = i18n.LANGS[self._sel('lang')]
        st['dpi_custom'] = int(self.get('dpi_custom').Value)
        st['bg_alpha'] = int(self.get('bg_alpha').Value)
        st['grid_step'] = float(self.get('grid_step').Value)
        st['margin'] = float(self.get('margin').Value)
        st['open_folder'] = self.get('open_folder').State == 1
        st['vsdx_native'] = self.get('vsdx_native').State == 1
        st['bg_color'] = list(self.bg_color)
        st.normalize()
        return st

    def run(self, execute=True):
        if not execute:                      # used by the self-test: build only
            st = self.collect()
            return st
        dlg = self.create()
        try:
            dlg.getControl('browse').addActionListener(_Listener(self.on_browse))
            dlg.getControl('pick').addActionListener(_Listener(self.on_pick))
            dlg.getControl('reset').addActionListener(_Listener(self.on_reset))
            dlg.getControl('dpi_index').addItemListener(_Listener(self.update_enabled))
            dlg.getControl('bg_type').addItemListener(_Listener(self.on_bg))
            self.update_enabled()
            if dlg.execute() == RES_OK:
                self.collect().save()
                return True
            return False
        finally:
            dlg.dispose()


class ChoiceDialog(_Dialog):
    """A question with one button per choice. run() returns the chosen key or None."""

    def __init__(self, ui, caption, text, choices):
        super().__init__(ui, title() + ' - ' + caption, 250)
        self.choices = choices
        self.result = None
        self.add('text', 'FixedText', 10, 8, 230, 12, Label=text)
        y = 26
        for key, label in choices:
            self.add('c_' + key, 'Button', 10, y, 230, 18, Label=label, Align=0)
            y += 22
        self.add('cancel', 'Button', 180, y + 4, 60, 14, Label=T('btn.cancel'), PushButtonType=2)
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
    """About VecStamp: version, licence, the free / open-source promise, the donation QR codes
    and the interface language."""

    def __init__(self, ui):
        super().__init__(ui, T('about.title', title()), 320)
        self.add('logo', 'ImageControl', 10, 10, 32, 32, ImageURL=icon_url('logo_64.png'), ScaleImage=True, Border=0)
        self.add('name', 'FixedText', 50, 12, 260, 14, Label=T('about.name', core.VERSION), FontHeight=13, FontWeight=150)
        self.add('tag', 'FixedText', 50, 27, 260, 19, Label=T('about.tag'), MultiLine=True)
        self.add('info', 'FixedText', 10, 48, 300, 40, MultiLine=True,
                 Label=T('about.info', core.AUTHOR, core.EMAIL, core.REPO))
        self.add('free', 'FixedText', 10, 88, 300, 34, MultiLine=True, Label=T('about.free'))
        self.add('qr1', 'ImageControl', 50, 126, 100, 126, ImageURL=icon_url('donate_alipay.png'), ScaleImage=True,
                 ScaleMode=1, Border=0)
        self.add('qr2', 'ImageControl', 170, 126, 100, 126, ImageURL=icon_url('donate_wechat.png'), ScaleImage=True,
                 ScaleMode=1, Border=0)
        self.add('home', 'Button', 10, 262, 76, 14, Label=T('about.home'))
        self.add('fb', 'Button', 90, 262, 68, 14, Label=T('about.feedback'))
        self.add('lang', 'Button', 162, 262, 92, 14, Label=T('about.langButton', i18n.NAMES[i18n.lang_index()]),
                 HelpText=T('lang.title'))
        self.add('ok', 'Button', 258, 262, 52, 14, Label=T('btn.ok'), PushButtonType=1, DefaultButton=True)
        self.model.Height = 284

    def run(self, execute=True):
        if not execute:
            return True
        dlg = self.create()
        try:
            dlg.getControl('home').addActionListener(_Listener(self.ui.home))
            dlg.getControl('fb').addActionListener(_Listener(self._feedback))
            dlg.getControl('lang').addActionListener(_Listener(self._language))
            dlg.execute()
            return True
        finally:
            dlg.dispose()

    def _feedback(self):
        self.dlg.endExecute()
        self.ui.feedback()

    def _language(self):
        self.dlg.endExecute()
        self.ui.language()
