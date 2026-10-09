"""VecStamp (Shi Yin) for LibreOffice Impress / Draw - UNO component.

Menu and toolbar entries (Addons.xcu) call
    service:com.vluckyzhang.vecstamp.Exporter?<action>
where <action> is export:<ext> | export:default | batch | picture | settings |
openfolder | about | donate | feedback | home | help.

Copyright (c) 2026 vluckyzhang <vluckyzhang@gmail.com>. MIT License.
"""
import json
import os

import unohelper
from com.sun.star.task import XJobExecutor

import vecstamp_core as core
import vecstamp_ui

IMPLEMENTATION_NAME = 'com.vluckyzhang.vecstamp.Exporter'


class VecStampJob(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx

    def trigger(self, args):
        action, _, param = (args or '').partition(':')
        ui = vecstamp_ui.UI(self.ctx)
        try:
            if action == 'export':
                ui.export(None if param in ('', 'default') else param)
            elif action == 'batch':
                ui.batch()
            elif action == 'picture':
                ui.picture()
            elif action == 'settings':
                ui.settings_dialog()
            elif action == 'openfolder':
                ui.open_folder()
            elif action in ('about', 'donate'):
                ui.about()
            elif action == 'feedback':
                ui.feedback()
            elif action == 'home':
                ui.home()
            elif action == 'help':
                ui.help()
            elif action == 'selftest':
                self.selftest(param)
        except Exception as exc:          # never let an exception escape into LibreOffice
            if action == 'selftest':
                raise
            ui.handle_error(exc)

    def selftest(self, out_dir):
        """Non-interactive check used by the project's tests: exports the selection
        (or the current page) of the active document in every format to out_dir."""
        desktop = self.ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', self.ctx)
        doc = desktop.getCurrentComponent()
        if doc is None:                   # hidden documents are never "current"
            docs = desktop.getComponents().createEnumeration()
            while doc is None and docs.hasMoreElements():
                d = docs.nextElement()
                if hasattr(d, 'getDrawPages'):
                    doc = d
        target = core.selection_target(doc) or core.current_page_target(doc)
        st = core.Settings(os.path.join(out_dir, 'settings.json'))
        report = {'version': core.VERSION, 'files': {}}
        for ext, _ in core.FORMATS:
            path = os.path.join(out_dir, 'selftest.' + ext)
            core.export_target(self.ctx, doc, target, ext, path, st)
            report['files'][ext] = os.path.getsize(path)
        ui = vecstamp_ui.UI(self.ctx, doc)
        report['dialog'] = sorted(ui.settings_dialog(execute=False).keys())
        about = vecstamp_ui.AboutDialog(ui)
        report['about'] = sorted(about.model.getElementNames())
        report['choice'] = sorted(vecstamp_ui.ChoiceDialog(ui, 't', 'x', [('a', 'A'), ('b', 'B')]).model.getElementNames())
        with open(os.path.join(out_dir, 'selftest.json'), 'w', encoding='utf-8') as f:
            json.dump(report, f, ensure_ascii=False, indent=1)


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(VecStampJob, IMPLEMENTATION_NAME, ('com.sun.star.task.Job',))
