"""VecStamp (Shi Yin) for LibreOffice - interface languages.

Simplified Chinese (default), Traditional Chinese, English and Japanese. The strings live in
vecstamp_strings.json next to this file: {key: {"zh-CN": ..., "zh-TW": ..., "en": ..., "ja": ...}}.
Placeholders are {0}, {1}, ... The language is a VecStamp setting and takes effect at once
(the next dialog or message uses it); LibreOffice's own menus follow LibreOffice's UI language.

Copyright (c) 2026 vluckyzhang <vluckyzhang@gmail.com>. MIT License.
"""
import json
import os

LANGS = ('zh-CN', 'zh-TW', 'en', 'ja')
NAMES = ('简体中文', '繁體中文', 'English', '日本語')
DEFAULT = 'zh-CN'

_HERE = os.path.dirname(os.path.abspath(__file__))
_table = None
_lang = DEFAULT


def table():
    global _table
    if _table is None:
        try:
            with open(os.path.join(_HERE, 'vecstamp_strings.json'), encoding='utf-8') as f:
                _table = json.load(f)
        except (OSError, ValueError):
            _table = {}
    return _table


def set_lang(lang):
    global _lang
    _lang = lang if lang in LANGS else DEFAULT


def lang():
    return _lang


def lang_index(lang_code=None):
    code = lang_code or _lang
    return LANGS.index(code) if code in LANGS else 0


def T(key, *args):
    """The text of key in the current language, with {0}, {1}, ... replaced by args."""
    entry = table().get(key)
    if not entry:
        s = key
    else:
        s = entry.get(_lang) or entry.get(DEFAULT) or key
    for i, a in enumerate(args):
        s = s.replace('{%d}' % i, str(a))
    return s
