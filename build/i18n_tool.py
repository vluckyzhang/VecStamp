"""Maintenance tool for VecStamp's string tables (src/i18n/strings.json for PowerPoint,
libreoffice/extension/pythonpath/vecstamp_strings.json for LibreOffice).

Each table is {key: {"zh-CN": ..., "zh-TW": ..., "en": ..., "ja": ...}}; placeholders are {0}, {1}, ...
(website/i18n.json is keyed by the Simplified Chinese text of the page; an empty en / ja text is allowed
there when the words move to another fragment of the same sentence.)

    python build/i18n_tool.py check                 # every key has all four languages and the same placeholders
    python build/i18n_tool.py fill-tw [--force]     # fill missing zh-TW from zh-CN with OpenCC (s2twp) + fix-ups
                                                    #   (pip install opencc-python-reimplemented)

fill-tw only writes entries whose zh-TW is missing or empty (--force: all of them), so hand-made
Traditional Chinese wording is kept.
"""
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TABLES = [os.path.join(ROOT, 'src', 'i18n', 'strings.json'),
          os.path.join(ROOT, 'libreoffice', 'extension', 'pythonpath', 'vecstamp_strings.json'),
          os.path.join(ROOT, 'website', 'i18n.json'),
          os.path.join(ROOT, 'build', 'promo_i18n.json')]
FILE_NAMES = re.compile(r'安装或卸载-[\w.]+|快速开始\.txt')     # names of files in the repository stay as they are
LANGS = ('zh-CN', 'zh-TW', 'en', 'ja')
# Taiwan wording OpenCC does not choose by itself
TW_FIXES = [('自定義', '自訂'), ('選中', '選取'), ('生成', '產生'), ('創建', '建立'), ('鏈接', '連結'),
            ('菜單', '選單'), ('文本', '文字'), ('窗口', '視窗'), ('質量', '品質'),
            ('卸載', '解除安裝'), ('信息', '資訊'), ('默認', '預設'), ('文件夾', '資料夾'), ('剪貼板', '剪貼簿'),
            ('打印', '列印'), ('屏幕', '螢幕'), ('激活', '啟用'), ('適配', '相容'), ('彈出', '跳出'),
            ('位圖', '點陣圖'), ('矢量', '向量'), ('幻燈片', '投影片'), ('反饋', '意見回饋'), ('郵箱', '信箱'),
            ('訪達', 'Finder'), ('資源管理器', '檔案總管'), ('佔位符', '預留位置'), ('讚賞', '贊助'), ('網格', '格線'),
            ('當前', '目前'), ('單擊', '按一下'), ('雙擊', '按兩下'), ('工具欄', '工具列'), ('新建', '新增'),
            ('邊距', '邊界'), ('賬號', '帳號'), ('個性色', '輔色'), ('主頁', '首頁'), ('開源協議', '開源授權'),
            ('MIT 協議', 'MIT 授權'), ('後臺', '背景'), ('進位制', '進位'), ('二維碼', 'QR Code'), ('重啟', '重新啟動'),
            ('“', '「'), ('”', '」'), ('‘', '『'), ('’', '』')]
PLACEHOLDER = re.compile(r'\{\d+\}')


def load(path):
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def save(path, table):
    with open(path, 'w', encoding='utf-8', newline='\n') as f:
        json.dump(table, f, ensure_ascii=False, indent=1)
        f.write('\n')


def to_tw(text, cc):
    names = FILE_NAMES.findall(text)
    t = cc.convert(FILE_NAMES.sub('\uE010', text))
    for name in names:
        t = t.replace('\uE010', name, 1)
    for a, b in TW_FIXES:
        t = t.replace(a, b)
    t = re.sub(r'([\u4e00-\u9fff])(Finder)', r'\1 \2', t)      # spaces around words that replaced Chinese ones
    return re.sub(r'(Finder)([\u4e00-\u9fff])', r'\1 \2', t)


def check():
    bad = 0
    for path in TABLES:
        if not os.path.exists(path):
            continue
        table = load(path)
        for key, entry in table.items():
            missing = [l for l in LANGS if not isinstance(entry.get(l), str) or (not entry[l] and l in ('zh-CN', 'zh-TW'))]
            if missing:
                print(f'{os.path.relpath(path, ROOT)}: {key}: missing {", ".join(missing)}')
                bad += 1
                continue
            ref = sorted(PLACEHOLDER.findall(entry['zh-CN']))
            for lang in LANGS[1:]:
                if sorted(PLACEHOLDER.findall(entry[lang])) != ref:
                    print(f'{os.path.relpath(path, ROOT)}: {key}: placeholders differ in {lang}')
                    bad += 1
        print(f'{os.path.relpath(path, ROOT)}: {len(table)} strings x {len(LANGS)} languages')
    return bad


def fill_tw(force=False):
    from opencc import OpenCC
    cc = OpenCC('s2twp')
    for path in TABLES:
        if not os.path.exists(path):
            continue
        table = load(path)
        n = 0
        for entry in table.values():
            if force or not entry.get('zh-TW'):
                entry['zh-TW'] = to_tw(entry['zh-CN'], cc)
                n += 1
        for key in table:                       # keep the language order inside every entry
            table[key] = {lang: table[key].get(lang, '') for lang in LANGS}
        save(path, table)
        print(f'{os.path.relpath(path, ROOT)}: {n} zh-TW strings filled')


if __name__ == '__main__':
    cmd = sys.argv[1] if len(sys.argv) > 1 else 'check'
    if cmd == 'fill-tw':
        fill_tw('--force' in sys.argv)
    else:
        sys.exit(1 if check() else 0)
