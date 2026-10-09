"""Static checks for the VBA module (run on any OS, no Office needed).

    python build/check_vba.py src/VecStamp.bas

For each target platform (Windows 64-bit, Windows 32-bit, macOS) it
1. resolves #If / #Else / #End If the way the VBA compiler would,
2. checks Option Explicit (every identifier used in a procedure is declared),
3. checks Exit Sub / Exit Function match the procedure kind,
4. checks every called module procedure exists on that platform and gets
   at least its required number of arguments,
5. writes build/out/VecStamp.<platform>.bas for an optional grammar check.
It is a safety net, not a compiler: always do a final Debug > Compile in VBE.
"""
import os
import re
import sys

KEYWORDS = set('''and as boolean byref byval call case const dim do each else elseif end exit explicit false for
function goto if in is let long loop mod new next not nothing on option or private public resume select set
single static step string sub then to true type until variant wend while with integer double error ptrsafe
declare lib longptr object attribute vb_name me xor like typeof redim preserve currency date byte optional
paramarray friend property get event'''.split())
BUILTINS = set('''msgbox len left right mid trim lcase ucase replace instr instrrev cstr clng cdbl val str format now
timer doevents createobject environ chrw rnd randomize int isnumeric typename strptr lenb array iif vbcrlf vbnullchar
vbtab vbcr vblf vbinformation vbexclamation vbquestion vbyesno vbyesnocancel vbyes vbno vbcancel vbobjecterror err
application inputbox rgb collection vbmsgboxresult dir mkdir kill getsetting savesetting deletesetting
applescripttask vbdirectory varptr ascb midb instrb leftb rightb chrb chr asc lbound ubound split hex round
freefile open output print close filecopy binary access read write put lof vbtextcompare vbbinarycompare join abs
sgn fix cint csng clngptr isarray isempty isnull lcase ascw cos sin tan sqr atn vbdefaultbutton2 vbokonly
activepresentation loadpicture'''.split())


def preprocess(lines, defines):
    out, stack = [], []          # stack of (active_now, any_branch_taken, parent_active)
    for line in lines:
        t = line.strip()
        m = re.match(r'#If\s+(.*?)\s+Then$', t, re.I)
        if m:
            parent = all(s[0] for s in stack)
            val = evaluate(m.group(1), defines)
            stack.append([parent and val, val, parent])
            out.append('')
            continue
        m = re.match(r'#ElseIf\s+(.*?)\s+Then$', t, re.I)
        if m:
            top = stack[-1]
            val = evaluate(m.group(1), defines) and not top[1]
            top[0] = top[2] and val
            top[1] = top[1] or val
            out.append('')
            continue
        if re.match(r'#Else$', t, re.I):
            top = stack[-1]
            top[0] = top[2] and not top[1]
            top[1] = True
            out.append('')
            continue
        if re.match(r'#End If$', t, re.I):
            stack.pop()
            out.append('')
            continue
        out.append(line if all(s[0] for s in stack) else '')
    assert not stack, 'unbalanced #If'
    return out


def evaluate(expr, defines):
    expr = re.sub(r'\bNot\b', ' not ', expr)
    expr = re.sub(r'\bAnd\b', ' and ', expr)
    expr = re.sub(r'\bOr\b', ' or ', expr)
    return bool(eval(expr, {}, defines))


def strip(line):
    out, i, n = [], 0, len(line)
    while i < n:
        c = line[i]
        if c == '"':
            j = i + 1
            while j < n:
                if line[j] == '"':
                    if j + 1 < n and line[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            out.append('""')
            i = j + 1
            continue
        if c == "'":
            break
        out.append(c)
        i += 1
    return ''.join(out)


def logical_lines(lines):
    res, buf, start = [], '', None
    for k, l in enumerate(lines, 1):
        s = strip(l)
        if start is None:
            start = k
        if s.rstrip().endswith(' _'):
            buf += s.rstrip()[:-1] + ' '
            continue
        res.append((start, buf + s))
        buf, start = '', None
    return res


def check(lines, label):
    errors = []
    code = logical_lines(lines)
    procs = {}
    module_names = set()
    for _, l in code:
        m = re.match(r'\s*(?:Private|Public)?\s*(Sub|Function)\s+(\w+)\s*\((.*)\)', l, re.I)
        if m and not re.match(r'\s*(?:Private|Public)\s+Declare', l, re.I):
            params = [p for p in m.group(3).split(',') if p.strip()]
            required = sum(1 for p in params if 'optional' not in p.lower() and 'paramarray' not in p.lower())
            total = 10 ** 6 if any('paramarray' in p.lower() for p in params) else len(params)
            procs[m.group(2).lower()] = (m.group(1).lower(), required, total)
        m = re.match(r'\s*(?:Private|Public)\s+Declare\s+(?:PtrSafe\s+)?(?:Function|Sub)\s+(\w+)', l, re.I)
        if m:
            module_names.add(m.group(1).lower())
        m = re.match(r'\s*(?:Private|Public)\s+(?:Const\s+)?(\w+)\s*(?:\([^)]*\))?\s+As\b', l, re.I)
        if m:
            module_names.add(m.group(1).lower())
            if not re.match(r'\s*(?:Private|Public)\s+Const\b', l, re.I):
                rest = re.sub(r'^\s*(?:Private|Public)\s+', '', l)
                for part in split_args(rest):                      # Private a As X, b() As Y
                    w = re.findall(r'\w+', part)
                    if w:
                        module_names.add(w[0].lower())
        m = re.match(r'\s*Private Type (\w+)', l, re.I)
        if m:
            module_names.add(m.group(1).lower())
    module_names |= set(procs)

    cur, kind, decl = None, None, set()
    for ln, l in code:
        m = re.match(r'\s*(?:Private|Public)\s+(Sub|Function)\s+(\w+)\s*\((.*)\)', l, re.I)
        if m:
            kind, cur = m.group(1).lower(), m.group(2)
            decl = set()
            for p in m.group(3).split(','):
                w = re.findall(r'\w+', re.sub(r'\b(ByVal|ByRef|Optional|ParamArray)\b', '', p, flags=re.I))
                if w:
                    decl.add(w[0].lower())
            continue
        if re.match(r'\s*End (Sub|Function)\b', l, re.I):
            cur = None
            continue
        if cur is None:
            continue
        for m in re.finditer(r'\b(?:Dim|Static|ReDim)\s+(?:Preserve\s+)?(.*)', l, re.I):
            for part in m.group(1).split(','):
                w = re.findall(r'\w+', part)
                if w:
                    decl.add(w[0].lower())
        for ex in re.findall(r'\bExit (Sub|Function)\b', l, re.I):
            if ex.lower() != kind:
                errors.append(f'{label} line {ln}: Exit {ex} inside a {kind} ({cur})')
        if re.match(r'^\s*\w+:\s*$', l):
            continue
        for m in re.finditer(r'(?<![\w$.&#])([A-Za-z_]\w*)\$?', l):
            name = m.group(1).lower()
            if name in KEYWORDS or name in BUILTINS or name in decl or name in module_names:
                continue
            if re.search(r'GoTo\s+' + re.escape(m.group(1)) + r'\b', l, re.I) or \
               re.search(r'Resume\s+' + re.escape(m.group(1)) + r'\b', l, re.I):
                continue
            if name.startswith('h') and re.search(r'&' + re.escape(m.group(1)), l):
                continue
            errors.append(f'{label} line {ln}: undeclared "{m.group(1)}" in {cur}')
        # calls to module procedures: argument count
        for stmt in statements(l):
            for name, (pkind, req, total) in procs.items():
                for m in re.finditer(r'(?<![\w.])' + name + r'\b', stmt, re.I):
                    if re.match(r'\s*(?:Private|Public)', stmt, re.I) or name == cur.lower():
                        continue
                    rest = stmt[m.end():]
                    if rest.lstrip().startswith('('):
                        k = stmt.index('(', m.end())
                        inner = balanced(stmt, k)
                        n = len(split_args(inner)) if inner.strip() else 0
                    elif stmt.strip().lower().startswith(name) and m.start() == len(stmt) - len(stmt.lstrip()):
                        n = len(split_args(rest)) if rest.strip() else 0   # statement call
                    elif re.match(r'\s*=', rest):
                        continue
                    else:
                        n = 0
                    if n < req or n > total:
                        errors.append(f'{label} line {ln}: {name} called with {n} args (expects {req}..{total})')
    return errors


def statements(l):
    out = []
    for seg in l.split(':'):
        for part in re.split(r'\bThen\b|\bElse\b', seg, flags=re.I):
            if part.strip():
                out.append(part)
    return out


def balanced(s, k):
    depth = 0
    for j in range(k, len(s)):
        if s[j] == '(':
            depth += 1
        elif s[j] == ')':
            depth -= 1
            if depth == 0:
                return s[k + 1:j]
    return s[k + 1:]


def split_args(s):
    depth, cur, out = 0, '', []
    for c in s:
        if c == '(':
            depth += 1
        elif c == ')':
            depth -= 1
        if c == ',' and depth == 0:
            out.append(cur)
            cur = ''
        else:
            cur += c
    if cur.strip():
        out.append(cur)
    return out


def main(path):
    lines = open(path, encoding='utf-8-sig').read().splitlines()
    outdir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'out')
    os.makedirs(outdir, exist_ok=True)
    total = []
    for label, defines in [('win64', {'Mac': False, 'VBA7': True, 'Win64': True}),
                           ('win32', {'Mac': False, 'VBA7': True, 'Win64': False}),
                           ('mac', {'Mac': True, 'VBA7': True, 'Win64': False})]:
        pre = preprocess(lines, defines)
        open(os.path.join(outdir, f'VecStamp.{label}.bas'), 'w', encoding='utf-8', newline='\r\n').write('\n'.join(pre))
        errs = check(pre, label)
        total += errs
        print(f'{label}: {len(errs)} issue(s)')
    for e in total:
        print('  ', e)
    sys.exit(1 if total else 0)


if __name__ == '__main__':
    main(sys.argv[1])
