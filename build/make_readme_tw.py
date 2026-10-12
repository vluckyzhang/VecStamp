"""README.zh-TW.md from README.md (Simplified -> Traditional Chinese, Taiwan wording).

    pip install opencc-python-reimplemented
    python build/make_readme_tw.py

Code spans, URLs, file names and image / link targets stay untouched; the same Taiwan
fix-ups as the interface strings are applied (build/i18n_tool.py). Review the result
before committing - README.zh-TW.md is a normal file that can be edited by hand afterwards.
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import i18n_tool  # noqa: E402

SWITCH_CN = ('<b>简体中文</b> | <a href="README.zh-TW.md">繁體中文</a> | <a href="README.en.md">English</a> | '
             '<a href="README.ja.md">日本語</a>')
SWITCH_TW = ('<a href="README.md">简体中文</a> | <b>繁體中文</b> | <a href="README.en.md">English</a> | '
             '<a href="README.ja.md">日本語</a>')
PROTECT = [
    r'安装或卸载-[\w.]+|快速开始\.txt',  # file names in the repository
    r'`[^`\n]+`',                       # code spans
    r'(?:src|href)="(?!#)[^"]*"',       # image / link targets (anchors are converted like the headings)
    r'\]\((?!#)[^)]*\)',                # markdown link targets
    r'https?://[^\s<>")]+',             # bare URLs
    r'<[^>\n]*@[^>\n]*>',               # e-mail autolinks
]


def convert(text, cc):
    keep = []

    def stash(m):
        keep.append(m.group(0))
        return '\uE000%d\uE001' % (len(keep) - 1)
    for pat in PROTECT:
        text = re.sub(pat, stash, text, flags=re.S)
    text = i18n_tool.to_tw(text, cc)
    while '\uE000' in text:                          # placeholders can be nested
        text = re.sub('\uE000(\\d+)\uE001', lambda m: keep[int(m.group(1))], text)
    return text


def main():
    from opencc import OpenCC
    src = open(os.path.join(ROOT, 'README.md'), encoding='utf-8').read()
    if SWITCH_CN not in src:
        raise SystemExit('language switcher not found in README.md')
    out = convert(src.replace(SWITCH_CN, ''), OpenCC('s2twp')).replace('', SWITCH_TW)
    out = out.replace('http://vluckyzhang.cn/shiyin/"', 'http://vluckyzhang.cn/shiyin/zh-TW/"')
    # README images drawn in Traditional Chinese (build/make_promo.py)
    out = re.sub(r'assets/promo/(hero|workflow|formats|backgrounds|vector-vs-bitmap|platforms)\.png',
                 r'assets/promo/\1.zh-TW.png', out)
    with open(os.path.join(ROOT, 'README.zh-TW.md'), 'w', encoding='utf-8', newline='\n') as f:
        f.write(out)
    print('wrote README.zh-TW.md')


if __name__ == '__main__':
    main()
