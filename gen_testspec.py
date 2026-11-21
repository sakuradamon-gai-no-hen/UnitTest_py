#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
C++ソースから単体試験仕様書(Excel)を自動生成する。

使い方:
    python gen_testspec.py 試験対象.cpp            # このスクリプトと同じ場所に出力
    python gen_testspec.py ソースのディレクトリ      # 配下の .cpp/.cc/.cxx をまとめて1ファイルに出力
    python gen_testspec.py 試験対象.cpp -o 出力先     # 出力先を変更
    python gen_testspec.py 試験対象.cpp --mode mcc   # 複数条件網羅(全組合せ)で作成

カバレッジ基準:
    c2  (既定) 条件網羅。各判定の「個々の条件」が真・偽を最低1回ずつ取り、
               かつ判定全体も真・偽を取る(C1も同時に満たす)最小ケース集合を作成。
    mcc        複数条件網羅。短絡評価を考慮した全組合せ。

制限(正規表現+簡易構文解析のため):
    - マクロ展開/#if による切替は考慮しない(前処理行は無視)
    - ラムダ式の内部分岐は解析しない
    - 入力値は「各条件を真/偽にする」という形で出力(具体値は人手で記入)
依存: openpyxl (pip install openpyxl)
"""
import argparse
import bisect
import itertools
import re
import sys
from pathlib import Path
from typing import List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

MAX_ATOMS = 12
IDENT = re.compile(r"[A-Za-z0-9_]")
SRC_EXT = {".cpp", ".cc", ".cxx", ".c++"}
HEADERS = ["cppファイル名", "クラス::メソッド名", "前提条件", "入力値",
           "期待結果(分岐で実行される行)", "結果", "備考"]


# ---------------------------------------------------------------- 前処理
def read_text(p: Path) -> str:
    raw = p.read_bytes()
    for enc in ("utf-8-sig", "cp932", "euc-jp"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            pass
    return raw.decode("utf-8", errors="replace")


def clean(src: str) -> Tuple[str, str]:
    """(構造解析用, 表示用) を返す。位置はsrcと完全一致。
    構造解析用: コメント/文字列/前処理行を空白化。表示用: コメント/前処理行のみ空白化。"""
    s, d = list(src), list(src)
    n, i = len(src), 0

    def blank(buf, a, b):
        for k in range(a, b):
            if buf[k] not in "\r\n":
                buf[k] = " "

    while i < n:
        c, two = src[i], src[i:i + 2]
        if two == "//":
            j = src.find("\n", i)
            j = n if j < 0 else j
            blank(s, i, j); blank(d, i, j); i = j
        elif two == "/*":
            j = src.find("*/", i + 2)
            j = n if j < 0 else j + 2
            blank(s, i, j); blank(d, i, j); i = j
        elif c in "\"'":
            j = i + 1
            while j < n and src[j] != c and src[j] != "\n":
                if src[j] == "\\":
                    j += 1
                j += 1
            blank(s, i + 1, min(j, n)); i = j + 1
        elif c == "#" and src[src.rfind("\n", 0, i) + 1:i].strip() == "":
            j = i
            while True:
                e = src.find("\n", j)
                if e < 0:
                    e = n
                    break
                if src[e - 1] == "\\" or src[max(e - 2, 0):e] == "\\\r":
                    j = e + 1
                    continue
                break
            blank(s, i, e); blank(d, i, e); i = e
        else:
            i += 1
    return "".join(s), "".join(d)


def match(t: str, i: int) -> int:
    o = t[i]
    cl = {"(": ")", "[": "]", "{": "}"}[o]
    depth = 0
    for k in range(i, len(t)):
        if t[k] == o:
            depth += 1
        elif t[k] == cl:
            depth -= 1
            if depth == 0:
                return k
    return len(t) - 1


def split_spans(s: str, a: int, b: int, sep: str) -> List[Tuple[int, int]]:
    spans, depth, last, i = [], 0, a, a
    while i < b:
        c = s[i]
        if c in "([{":
            depth += 1
        elif c in ")]}":
            depth -= 1
        elif depth == 0 and s.startswith(sep, i):
            spans.append((last, i))
            i += len(sep)
            last = i
            continue
        i += 1
    spans.append((last, b))
    return spans


def has_logic(s: str, a: int, b: int) -> bool:
    return len(split_spans(s, a, b, "||")) > 1 or len(split_spans(s, a, b, "&&")) > 1


# ---------------------------------------------------------------- 条件式 → 原子条件
class CondBuilder:
    def __init__(self, s: str, d: str):
        self.s, self.d, self.atoms = s, d, []

    def build(self, a: int, b: int) -> str:
        s = self.s
        while True:
            while a < b and s[a].isspace():
                a += 1
            while b > a and s[b - 1].isspace():
                b -= 1
            if b - a >= 2 and s[a] == "(" and match(s, a) == b - 1:
                a += 1
                b -= 1
                continue
            break
        if a >= b:
            return "True"
        for sep, op in (("||", " or "), ("&&", " and ")):
            spans = split_spans(s, a, b, sep)
            if len(spans) > 1:
                return "(" + op.join(self.build(x, y) for x, y in spans) + ")"
        if s[a] == "!":
            a2 = a + 1
            while a2 < b and s[a2].isspace():
                a2 += 1
            if a2 < b and s[a2] == "(" and match(s, a2) == b - 1 and has_logic(s, a2 + 1, b - 1):
                return "(not " + self.build(a2 + 1, b - 1) + ")"
        self.atoms.append(" ".join(self.d[a:b].split()))
        return "_a(%d)" % (len(self.atoms) - 1)


def make_cases(expr: str, atoms: List[str], mode: str):
    """[(判定結果, {原子条件index: 値(評価されたもののみ)})] を返す"""
    code = compile(expr, "<cond>", "eval")
    seen = {}
    for bits in itertools.product([True, False], repeat=len(atoms)):
        trace = {}

        def _a(i, bits=bits, trace=trace):
            trace[i] = bits[i]
            return bits[i]

        out = bool(eval(code, {"_a": _a, "__builtins__": {}}))
        key = tuple(sorted(trace.items()))
        if key not in seen:
            seen[key] = (out, trace)
    items = list(seen.values())
    if mode == "mcc":
        return items
    # C2: 全原子条件の真/偽 + 判定の真/偽 を貪欲法で被覆
    def cover(it):
        return {(i, v) for i, v in it[1].items()} | {("D", it[0])}

    universe = set().union(*(cover(it) for it in items))
    covered, chosen = set(), []
    while covered != universe:
        best = max(items, key=lambda it: len(cover(it) - covered))
        chosen.append(best)
        covered |= cover(best)
    return chosen


# ---------------------------------------------------------------- 構文木
class Node:
    def __init__(self, kind, line, text="", **kw):
        self.kind, self.line, self.text = kind, line, text
        self.body: list = kw.get("body", [])
        self.els: Optional[list] = kw.get("els")
        self.cond = kw.get("cond")            # (expr, atoms) or None
        self.cond_text = kw.get("cond_text", "")
        self.cases = kw.get("cases", [])      # switch: (label, line, children, falls)
        self.catches = kw.get("catches", [])  # try: (decl, line, children)
        self.ternaries = kw.get("ternaries", [])  # [(line, cond_text, (expr, atoms))]
        self.do = kw.get("do", False)


class Parser:
    def __init__(self, src: str):
        self.src = src
        self.t, self.d = clean(src)
        self.n = len(src)
        self.starts = [0] + [m.end() for m in re.finditer(r"\n", src)]

    # --- helpers
    def line(self, pos): return bisect.bisect_right(self.starts, pos)
    def disp(self, a, b): return " ".join(self.d[a:b].split())
    def ws(self, i):
        while i < self.n and self.t[i].isspace():
            i += 1
        return i

    def kw(self, i, w):
        return (self.t.startswith(w, i)
                and not (i + len(w) < self.n and IDENT.match(self.t[i + len(w)])))

    def mkcond(self, a, b):
        cb = CondBuilder(self.t, self.d)
        expr = cb.build(a, b)
        if len(cb.atoms) > MAX_ATOMS:
            return (("_a(0)", [self.disp(a, b)]))
        return (expr, cb.atoms)

    # --- 文
    def parse_list(self, a, b):
        nodes, i = [], a
        while True:
            i = self.ws(i)
            if i >= b:
                break
            node, i2 = self.stmt(i, b)
            i = max(i2, i + 1)
            if node:
                nodes.append(node)
        return nodes

    def stmt(self, i, b):
        i = self.ws(i)
        if i >= b:
            return None, b
        c = self.t[i]
        if c == ";":
            return None, i + 1
        if c == "{":
            j = match(self.t, i)
            return Node("block", self.line(i), body=self.parse_list(i + 1, j)), j + 1
        for w in ("if", "while", "for", "switch", "do", "try"):
            if self.kw(i, w):
                return getattr(self, "p_" + w)(i, b)
        return self.simple(i, b)

    def branch(self, i, b):
        node, j = self.stmt(i, b)
        if node is None:
            return [], j
        return (node.body if node.kind == "block" else [node]), j

    def paren(self, i):
        i = self.ws(i)
        return i, match(self.t, i)

    def p_if(self, i, b):
        ln = self.line(i)
        i = self.ws(i + 2)
        if self.kw(i, "constexpr"):
            i += 9
        o, j = self.paren(i)
        spans = split_spans(self.t, o + 1, j, ";")     # if (init; cond)
        ca, cb = spans[-1]
        then, k = self.branch(j + 1, b)
        els, k2 = None, self.ws(k)
        if self.kw(k2, "else"):
            els, k = self.branch(k2 + 4, b)
        text = "if (%s)" % self.disp(ca, cb)
        return Node("if", ln, text, body=then, els=els, cond=self.mkcond(ca, cb)), k

    def p_while(self, i, b):
        ln = self.line(i)
        o, j = self.paren(i + 5)
        body, k = self.branch(j + 1, b)
        text = "while (%s)" % self.disp(o + 1, j)
        return Node("loop", ln, text, body=body, cond=self.mkcond(o + 1, j)), k

    def p_do(self, i, b):
        ln = self.line(i)
        body, k = self.branch(i + 2, b)
        k = self.ws(k)
        if self.kw(k, "while"):
            o, j = self.paren(k + 5)
            text = "do { } while (%s)" % self.disp(o + 1, j)
            return Node("loop", ln, text, body=body, cond=self.mkcond(o + 1, j), do=True), j + 1
        return Node("block", ln, body=body), k

    def p_for(self, i, b):
        ln = self.line(i)
        o, j = self.paren(i + 3)
        text = "for (%s)" % self.disp(o + 1, j)
        body, k = self.branch(j + 1, b)
        spans = split_spans(self.t, o + 1, j, ";")
        cond = None
        if len(spans) == 3:
            a, c = spans[1]
            if self.t[a:c].strip():
                cond = self.mkcond(a, c)
        else:  # 範囲for
            a, c = spans[-1]
            cond = ("_a(0)", ["範囲 (%s) に未処理の要素が存在する" % self.disp(a, c)])
        return Node("loop", ln, text, body=body, cond=cond), k

    def p_switch(self, i, b):
        ln = self.line(i)
        o, j = self.paren(i + 6)
        expr = self.disp(o + 1, j)
        k = self.ws(j + 1)
        cases = []
        if k < self.n and self.t[k] == "{":
            e = match(self.t, k)
            p = self.ws(k + 1)
            while p < e:
                if self.kw(p, "case") or self.kw(p, "default"):
                    q = p
                    while q < e and not (self.t[q] == ":" and self.t[q + 1:q + 2] != ":"
                                         and self.t[q - 1:q] != ":"):
                        q += 1
                    label = self.disp(p, q)
                    lline = self.line(p)
                    p = self.ws(q + 1)
                    children = []
                    while p < e and not (self.kw(p, "case") or self.kw(p, "default")):
                        nd, p2 = self.stmt(p, e)
                        p = self.ws(max(p2, p + 1))
                        if nd:
                            children.append(nd)
                    cases.append([label, lline, children])
                else:
                    p = self.ws(p + 1)
            k = e + 1
        return Node("switch", ln, "switch (%s)" % expr, cond_text=expr, cases=cases), k

    def p_try(self, i, b):
        ln = self.line(i)
        body, k = self.branch(i + 3, b)
        catches = []
        while True:
            k2 = self.ws(k)
            if not self.kw(k2, "catch"):
                break
            o, j = self.paren(k2 + 5)
            ch, k = self.branch(j + 1, b)
            catches.append((self.disp(o + 1, j), self.line(k2), ch))
        return Node("try", ln, "try", body=body, catches=catches), k

    def simple(self, i, b):
        t, j, depth = self.t, i, 0
        while j < b:
            c = t[j]
            if c in "([{":
                depth += 1
            elif c in ")]}":
                depth -= 1
            elif c == ";" and depth == 0:
                break
            j += 1
        end = min(j + 1, b)
        node = Node("stmt", self.line(i), self.disp(i, end))
        for q in range(i, j):
            if t[q] == "?":
                a = self.tern_start(q, i)
                ct = re.sub(r"^\s*return\b", "", t[a:q])
                a += len(t[a:q]) - len(ct)
                node.ternaries.append((self.line(q), self.disp(a, q), self.mkcond(a, q)))
        return node, end

    def tern_start(self, q, lo):
        t, depth, k = self.t, 0, q - 1
        while k >= lo:
            c = t[k]
            if c in ")]}":
                depth += 1
            elif c in "([{":
                if depth == 0:
                    return k + 1
                depth -= 1
            elif depth == 0:
                if c in ",;":
                    return k + 1
                if c == "=" and t[k - 1:k] not in ("=", "!", "<", ">") and t[k + 1:k + 2] != "=":
                    return k + 1
                if c == ":" and t[k - 1:k] != ":" and t[k + 1:k + 2] != ":":
                    return k + 1
            k -= 1
        return lo


# ---------------------------------------------------------------- 関数の抽出
class Func:
    def __init__(self, name, sig, a, b):
        self.name, self.sig, self.a, self.b = name, sig, a, b


NAME_RE = re.compile(r"((?:[A-Za-z_]\w*\s*::\s*)*~?[A-Za-z_]\w*|operator\s*[^\s(]+)\s*\(")
KEYWORDS = {"if", "for", "while", "switch", "catch", "return", "sizeof"}


def find_functions(p: Parser) -> List[Func]:
    t, d = p.t, p.d
    funcs, stack, i, n, hs = [], [], 0, len(t), 0
    while i < n:
        c = t[i]
        if c == ";":
            hs = i + 1
        elif c == "}":
            if stack:
                stack.pop()
            hs = i + 1
        elif c == "{":
            h = t[hs:i]
            kind, name = "other", ""
            m = NAME_RE.search(h) if "(" in h else None
            if m and m.group(1) not in KEYWORDS and "=" not in h[:m.start()]:
                kind, name = "func", re.sub(r"\s+", "", m.group(1))
            else:
                h2 = re.sub(r"template\s*<[^>]*>", "", h)
                mc = re.search(r"\b(?:class|struct)\s+([A-Za-z_]\w*)", h2)
                if mc and "(" not in h2:
                    kind, name = "class", mc.group(1)
                elif re.search(r"\bnamespace\b", h2):
                    kind = "ns"
            if kind == "func":
                op = h.index("(", m.end() - 1)
                close = match(h, op)
                tail = h[close + 1:]
                prev = t[:i].rstrip()[-1:]
                if re.search(r"(?<!:):(?!:)", tail) and (IDENT.match(prev or " ") or prev == ">"):
                    i = match(t, i) + 1          # コンストラクタ初期化子の {} は読み飛ばす
                    continue
                scope = [x[1] for x in stack if x[0] == "class"]
                full = name if "::" in name or not scope else "::".join(scope + [name])
                sig = " ".join(d[hs:i].split())
                sig = re.sub(r"^(?:(?:public|private|protected)\s*:\s*)+", "", sig)
                j = match(t, i)
                funcs.append(Func(full, sig, i + 1, j))
                i, hs = j + 1, j + 1
                continue
            stack.append((kind, name))
            hs = i + 1
        i += 1
    return funcs


# ---------------------------------------------------------------- 行生成
def terminates(nodes) -> bool:
    for nd in nodes:
        if nd.kind == "stmt" and re.match(r"(return|throw|break|continue|goto)\b", nd.text):
            return True
        if nd.kind == "block" and terminates(nd.body):
            return True
    return False


def flat(nodes):
    for nd in nodes:
        if nd.kind == "block":
            yield from flat(nd.body)
        else:
            yield nd


def branch_lines(nodes) -> List[str]:
    out = []
    for nd in flat(nodes):
        out.append("L%d: %s" % (nd.line, nd.text))
        if nd.kind == "try":
            out.extend(branch_lines(nd.body))
    return out


TF = {True: "真(true)", False: "偽(false)"}


class Gen:
    def __init__(self, fname, func: Func, mode):
        self.fname, self.func, self.mode = fname, func, mode
        self.rows = []

    def add(self, path, inp, exp, note):
        self.rows.append([self.fname, self.func.name,
                          "\n".join(path) if path else "特になし",
                          "\n".join(inp), "\n".join(exp), "", note])

    def decision(self, line, text, cond, path, lines_true, lines_false, note_kind):
        expr, atoms = cond
        label = "C2(条件網羅)" if self.mode == "c2" else "複数条件網羅"
        for out, trace in make_cases(expr, atoms, self.mode):
            inp = ["%s : %s" % (a, TF[trace[i]]) for i, a in enumerate(atoms) if i in trace]
            skipped = [a for i, a in enumerate(atoms) if i not in trace]
            note = "L%d %s %s / 判定結果=%s" % (line, note_kind, label, "真" if out else "偽")
            if skipped:
                note += "\n短絡評価で未評価: " + " , ".join(skipped)
            exp = ["L%d: %s → %s" % (line, text, "真" if out else "偽")]
            exp += lines_true if out else lines_false
            self.add(path, inp, exp, note)

    def walk(self, nodes, path):
        path = list(path)
        for nd in nodes:
            k = nd.kind
            if k == "block":
                self.walk(nd.body, path)
            elif k == "stmt":
                for ln, ctext, cond in nd.ternaries:
                    one = ["L%d: %s" % (nd.line, nd.text)]
                    self.decision(ln, "%s ? A : B" % ctext, cond, path,
                                  one + ["(三項演算子: A側を評価)"], one + ["(三項演算子: B側を評価)"],
                                  "三項演算子")
            elif k == "if":
                lt = branch_lines(nd.body) or ["(処理なし)"]
                lf = branch_lines(nd.els) if nd.els is not None else ["else節なし → 後続の処理へ"]
                self.decision(nd.line, nd.text, nd.cond, path, lt, lf or ["(処理なし)"], "if")
                self.walk(nd.body, path + ["L%d %s が真" % (nd.line, nd.text)])
                if nd.els is not None:
                    self.walk(nd.els, path + ["L%d %s が偽" % (nd.line, nd.text)])
                elif terminates(nd.body):
                    path = path + ["L%d %s が偽 (真の場合は return等で抜ける)" % (nd.line, nd.text)]
            elif k == "loop":
                if nd.cond is not None:
                    bl = branch_lines(nd.body) or ["(処理なし)"]
                    if nd.do:
                        lt, lf = ["(ループ継続) 本体を再実行"] + bl, ["ループ終了 → 後続の処理へ"]
                    else:
                        lt, lf = ["ループ本体を実行"] + bl, ["ループ本体を実行せず → 後続の処理へ"]
                    self.decision(nd.line, nd.text, nd.cond, path, lt, lf, "ループ")
                self.walk(nd.body, path + ["L%d %s が真(ループ内)" % (nd.line, nd.text)])
            elif k == "switch":
                has_default = False
                for idx, (label, ln, ch) in enumerate(nd.cases):
                    has_default |= label.startswith("default")
                    exp = ["L%d: %s" % (ln, label + ":")] + (branch_lines(ch) or ["(処理なし)"])
                    note = "L%d switch 分岐網羅(各case)" % nd.line
                    if idx < len(nd.cases) - 1 and not terminates(ch):
                        exp.append("※break等なし: 次のcaseへフォールスルー")
                        note += "\nフォールスルー有り"
                    inp = (["%s == %s" % (nd.cond_text, label[5:].strip())]
                           if label.startswith("case")
                           else ["%s : 上記case以外の値(default)" % nd.cond_text])
                    self.add(path, inp, exp, note)
                    self.walk(ch, path + ["L%d %s : %s" % (nd.line, nd.text, label)])
                if not has_default:
                    self.add(path, ["%s がいずれのcaseにも一致しない" % nd.cond_text],
                             ["L%d: %s → 該当caseなし(defaultなし)" % (nd.line, nd.text),
                              "switchを抜けて後続の処理へ"], "L%d switch defaultなし" % nd.line)
            elif k == "try":
                self.walk(nd.body, path)
                for decl, ln, ch in nd.catches:
                    self.add(path, ["try内で例外 (%s) が送出される" % decl],
                             ["L%d: catch (%s)" % (ln, decl)] + (branch_lines(ch) or ["(処理なし)"]),
                             "L%d catch" % ln)
                    self.walk(ch, path + ["例外 (%s) 発生" % decl])

    def run(self, nodes):
        self.walk(nodes, [])
        if not self.rows:
            self.add([], ["(分岐なし) 引数は代表値で可"],
                     branch_lines(nodes) or ["(処理なし)"], "分岐なし: 1ケースで全行を実行")
        self.rows[0][6] += ("\n" if self.rows[0][6] else "") + "シグネチャ: " + self.func.sig
        return self.rows


# ---------------------------------------------------------------- Excel出力
def write_excel(rows, out: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "単体試験仕様書"
    ws.append(HEADERS)
    for r in rows:
        ws.append(r)
    side = Side(style="thin", color="999999")
    border = Border(left=side, right=side, top=side, bottom=side)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="305496")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.border = border
    for col, w in zip("ABCDEFG", (20, 28, 40, 45, 55, 10, 40)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    if len(rows):
        dv = DataValidation(type="list", formula1='"OK,NG,未実施"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add("F2:F%d" % (len(rows) + 1))
    wb.save(out)


def main():
    ap = argparse.ArgumentParser(description="C++から単体試験仕様書(Excel)を生成")
    ap.add_argument("target", help="試験対象の .cpp ファイル or ディレクトリ")
    ap.add_argument("-o", "--outdir", help="出力先(既定: このスクリプトのある場所)")
    ap.add_argument("--mode", choices=["c2", "mcc"], default="c2")
    args = ap.parse_args()

    target = Path(args.target)
    if not target.exists():
        sys.exit("見つかりません: %s" % target)
    files = (sorted(p for p in target.rglob("*") if p.suffix.lower() in SRC_EXT)
             if target.is_dir() else [target])
    if not files:
        sys.exit("対象ファイルがありません")

    rows = []
    for f in files:
        parser = Parser(read_text(f))
        funcs = find_functions(parser)
        if not funcs:
            print("[WARN] 関数定義が見つかりません: %s" % f)
        for fn in funcs:
            rows += Gen(f.name, fn, args.mode).run(parser.parse_list(fn.a, fn.b))

    outdir = Path(args.outdir) if args.outdir else Path(__file__).resolve().parent
    outdir.mkdir(parents=True, exist_ok=True)
    stem = target.name if target.is_dir() else target.stem
    out = outdir / ("%s_単体試験仕様書.xlsx" % stem)
    write_excel(rows, out)
    print("出力しました: %s (%d ケース)" % (out, len(rows)))


if __name__ == "__main__":
    main()
