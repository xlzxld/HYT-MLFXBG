#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""bracket_lint.py - 语言感知的括号平衡检查器（零依赖，单文件，Python 3.8+）

为什么需要它
------------
真正的痛点不是"发现括号不平衡"，而是"找不到那个没闭合的括号是在哪一行开的"。
编译器通常只丢一句 "unexpected EOF while parsing"，
于是人就开始肉眼翻几百行 —— 这就是"修半天甚至修不好"的来源。

本工具做的事
------------
1. 精确报告第一个不平衡的位置（行:列），并画出插入符
2. 报告文件结束时仍然"开着"的括号栈 —— 每个都带行:列
3. 直接给出"在文件末尾按顺序补上 } )"的建议
4. 抓出混进代码的全角括号 （）［］｛｝【】 与弯引号 "" '' —— 中文输入法经典事故
   （注意：字符串和注释里的全角括号是合法的，不会被误报）
5. 语言感知，不会把字符串/注释/模板串/正则里的括号误判：
   Python 三引号、JS 模板串与正则字面量、Go 反引号原串、
   Rust 生命周期与 raw string、C++ raw string、Rust 嵌套块注释……

定位说明：它是"定位器"，不是编译器的替代品。
对 Python / Go / Rust 这类有极快语法检查的语言，先跑语法检查；
它只给一句 EOF 报错时，再用本工具定位到具体行列。

分发形态：本文件带 tsc-managed 标记，随 tsc install / sync 与
structure_guard.py 一起复制到项目 .agents/（结构门禁落盘件，git 钩子与
CI 使用）；技能目录里的是正本。

用法
----
    python bracket_lint.py src/**/*.py
    python bracket_lint.py --lang js - < app.js
    python bracket_lint.py --json broken.go
    python bracket_lint.py --depth deep.ts
    python bracket_lint.py --selftest

退出码：0 全部通过，1 存在不平衡或混入字符，2 用法错误。
"""

# tsc-managed —— 落盘件：随 tsc install / sync 复制到项目 .agents/，由技能托管更新；删除本行即视为项目接管。
from __future__ import annotations

import argparse
import json
import os
import sys

__version__ = "1.2.1"

OPEN = {"(": ")", "[": "]", "{": "}"}
CLOSE = {v: k for k, v in OPEN.items()}

# 全角括号：几乎必然是中文输入法事故，直接归一化后再判平衡
FULLWIDTH_BRACKET = {
    "\uff08": "(",   # （
    "\uff09": ")",   # ）
    "\uff3b": "[",   # ［
    "\uff3d": "]",   # ］
    "\uff5b": "{",   # ｛
    "\uff5d": "}",   # ｝
    "\u3010": "[",   # 【
    "\u3011": "]",   # 】
}
# 其余全角标点：不参与括号匹配，只报警
FULLWIDTH_PUNCT = {
    "\uff0c": ",",   # ，
    "\u3001": ",",   # 、
    "\uff1a": ":",   # ：
    "\uff1b": ";",   # ；
    "\uff01": "!",   # ！
    "\uff1f": "?",   # ？
    "\u3002": ".",   # 。
    "\uff0b": "+",   # ＋
    "\uff0d": "-",   # －
    "\uff0a": "*",   # ＊
    "\uff0f": "/",   # ／
    "\uff1d": "=",   # ＝
    "\u3000": " ",   # 全角空格
    "\uff5c": "|",   # ｜
    "\uff06": "&",   # ＆
    "\uff05": "%",   # ％
    "\uff04": "$",   # ＄
    "\uff03": "#",   # ＃
    "\uff20": "@",   # ＠
    "\uff3e": "^",   # ＾
    "\uff5e": "~",   # ～
    "\uff1c": "<",   # ＜
    "\uff1e": ">",   # ＞
}
SMART_QUOTES = {
    "\u201c": '"', "\u201d": '"',   # “ ”
    "\u2018": "'", "\u2019": "'",   # ‘ ’
    "\uff02": '"',                  # ＂
}

# 字符串规格: (开, 闭, 支持转义, 可跨行, 未闭合是否报警)
_S_PY = [
    ('"""', '"""', True, True, True),
    ("'''", "'''", True, True, True),
    ('"', '"', True, False, True),
    ("'", "'", True, False, True),
]
_S_C = [
    ('"', '"', True, False, True),
    ("'", "'", True, False, True),
]
_S_JSON = [('"', '"', True, False, True)]
_S_JS = [
    ('"', '"', True, False, True),
    ("'", "'", True, False, True),
    ("`", "`", True, True, True),
]
_S_GO = [
    ('"', '"', True, False, True),
    ("'", "'", True, False, True),
    ("`", "`", False, True, False),
]
_S_SH = [
    ('"', '"', True, False, True),
    ("'", "'", False, False, False),
    ("`", "`", True, False, False),
]
_S_SOFT = [
    ('"', '"', True, False, False),
    ("'", "'", False, False, False),
]
_S_SQL = [
    ("'", "'", False, True, False),
    ('"', '"', False, True, False),
]
_S_LUA = [
    ('"', '"', True, False, True),
    ("'", "'", True, False, True),
    ("[[", "]]", False, True, False),
]
_S_PS1 = [
    ('"', '"', True, True, True),
    ("'", "'", False, True, True),
]

PROFILES = {
    "py": dict(
        line=["#"],
        block=[('"""', '"""'), ("'''", "'''")],
        strings=_S_PY,
    ),
    "js": dict(
        line=["//"],
        block=[("/*", "*/")],
        strings=_S_JS,
        regex=True,
        template_interp=True,
    ),
    "go": dict(
        line=["//"],
        block=[("/*", "*/")],
        strings=_S_GO,
    ),
    "rust": dict(
        line=["//"],
        block=[("/*", "*/")],
        nest_block=True,
        strings=[('"', '"', True, False, True)],
        rust_raw=True,
        rust_char=True,
    ),
    "c": dict(
        line=["//"],
        block=[("/*", "*/")],
        strings=_S_C,
        cpp_raw=True,
    ),
    "json": dict(line=[], block=[], strings=_S_JSON),
    # yaml / plain 是"文本为主"的格式：裸标量可含中文与圆括号（如 shell case 的
    # `*.py)`），所以 () 不参与平衡（pairs 只留 [{），全角标点检测关闭
    # （fw=False）——文案里的 （）：， 是合法内容，不是输入法事故。
    "yaml": dict(line=["#"], block=[], strings=_S_SOFT, pairs="[{", fw=False),
    "toml": dict(line=["#"], block=[], strings=_S_SOFT, pairs="[{"),
    "sh": dict(line=["#"], block=[], strings=_S_SH),
    "sql": dict(line=["--"], block=[("/*", "*/")], strings=_S_SQL),
    "lua": dict(line=["--"], block=[("--[[", "]]")], strings=_S_LUA),
    "rb": dict(line=["#"], block=[("=begin", "=end")], strings=_S_C),
    "php": dict(
        line=["//", "#"],
        block=[("/*", "*/")],
        strings=[('"', '"', True, False, True), ("'", "'", True, False, True), ("`", "`", True, False, True)],
    ),
    "ps1": dict(
        line=["#"],
        block=[("<#", "#>")],
        strings=_S_PS1,
        # PowerShell 的转义符是反引号 `，不是反斜杠（收官体检 A-04①：
        # 按 \ 转义会把合法的 `" 误报成未闭合串，把非法的 \" 放行）
        esc="`",
    ),
    "plain": dict(
        line=["//", "#", "--"],
        block=[("/*", "*/")],
        strings=[('"', '"', True, False, False), ("'", "'", True, True, False), ("`", "`", True, True, False)],
        pairs="[{",
        fw=False,
    ),
}

EXT_MAP = {
    ".py": "py", ".pyw": "py", ".pyi": "py",
    ".js": "js", ".mjs": "js", ".cjs": "js", ".jsx": "js",
    ".ts": "js", ".tsx": "js", ".mts": "js", ".cts": "js",
    ".vue": "js", ".svelte": "js",
    ".go": "go",
    ".rs": "rust",
    ".c": "c", ".h": "c", ".cc": "c", ".cpp": "c", ".cxx": "c", ".hpp": "c",
    ".hh": "c", ".java": "c", ".cs": "c", ".kt": "c", ".kts": "c",
    ".swift": "c", ".scala": "c", ".dart": "c", ".m": "c", ".mm": "c",
    ".json": "json", ".jsonc": "json",
    ".yaml": "yaml", ".yml": "yaml",
    ".toml": "toml",
    ".sh": "sh", ".bash": "sh", ".zsh": "sh",
    ".sql": "sql",
    ".lua": "lua",
    ".rb": "rb",
    ".php": "php",
    ".ps1": "ps1", ".psm1": "ps1", ".psd1": "ps1",
}

REGEX_PREV_CHARS = set("(,=:[!&|?{};+-*%~^<>")
REGEX_PREV_WORDS = {
    "return", "typeof", "instanceof", "in", "of", "new", "delete", "void",
    "do", "else", "yield", "await", "case", "throw", "default",
}

USE_COLOR = False


def _c(code: str, text: str) -> str:
    return f"\033[{code}m{text}\033[0m" if USE_COLOR else text


class Scanner:
    def __init__(self, src: str, profile: dict, path: str, max_issues: int = 5):
        self.src = src
        self.p = profile
        self.path = path
        self.max_issues = max_issues
        self.n = len(src)
        self.lines = src.split("\n")
        self.i = 0
        self.line = 1
        self.ls = 0
        self.stack = []          # (char, line, col)
        self.issues = []         # dict(kind, line, col, msg)
        self.depth = []          # (结束的行号, 该行结束时的嵌套深度)
        self._seen = set()
        self._capped = False
        # profile 级开关：pairs=参与平衡的括号集（默认全部）；fw=全角/弯引号检测
        # （代码类语言开，yaml/plain 等文本格式关——裸标量里的全角标点是内容不是事故）
        self.pairs = set(self.p.get("pairs", "([{"))
        self.fw = bool(self.p.get("fw", True))

    # ---------- 基础工具 ----------

    def _col(self, idx: int | None = None) -> int:
        return (self.i if idx is None else idx) - self.ls + 1

    def _adv(self, new_i: int) -> None:
        if new_i <= self.i:
            return
        seg = self.src[self.i:new_i]
        k = seg.count("\n")
        if k:
            self.line += k
            self.ls = self.i + seg.rfind("\n") + 1
        self.i = new_i

    def _text(self, ln: int) -> str:
        if 1 <= ln <= len(self.lines):
            return self.lines[ln - 1].rstrip("\r")
        return ""

    def _report(self, kind: str, msg: str, line: int | None = None, col: int | None = None) -> None:
        line = self.line if line is None else line
        col = self._col() if col is None else col
        key = (kind, line, col)
        if key in self._seen:
            return
        self._seen.add(key)
        if len(self.issues) >= self.max_issues:
            self._capped = True
            return
        self.issues.append(dict(kind=kind, line=line, col=col, msg=msg))

    # ---------- 跳过类 ----------

    def _try_line_comment(self) -> bool:
        for tok in self.p["line"]:
            if self.src.startswith(tok, self.i):
                j = self.src.find("\n", self.i)
                self._adv(self.n if j < 0 else j)
                return True
        return False

    def _try_block_comment(self) -> bool:
        for op, cl in self.p["block"]:
            if not self.src.startswith(op, self.i):
                continue
            start_line, start_col = self.line, self._col()
            if self.p.get("nest_block"):
                depth, k = 1, self.i + len(op)
                while k < self.n and depth:
                    if self.src.startswith(op, k):
                        depth += 1
                        k += len(op)
                    elif self.src.startswith(cl, k):
                        depth -= 1
                        k += len(cl)
                    else:
                        k += 1
                if depth:
                    self._report("unterminated", f"块注释 {op} 从这里开始，一直没有 {cl}", start_line, start_col)
                self._adv(min(k, self.n))
            else:
                k = self.src.find(cl, self.i + len(op))
                if k < 0:
                    self._report("unterminated", f"块注释 {op} 从这里开始，一直没有 {cl}", start_line, start_col)
                    self._adv(self.n)
                else:
                    self._adv(k + len(cl))
            return True
        return False

    def _try_string(self) -> bool:
        esc_char = self.p.get("esc", "\\")  # 转义符按语言 profile（PS 是 `，其余默认 \）
        for op, cl, esc, multi, report in self.p["strings"]:
            if not self.src.startswith(op, self.i):
                continue
            start_line, start_col = self.line, self._col()
            k = self.i + len(op)
            closed = False
            while k < self.n:
                ch = self.src[k]
                if esc and ch == esc_char:
                    k += 2
                    continue
                if self.src.startswith(cl, k):
                    k += len(cl)
                    closed = True
                    break
                if not multi and ch == "\n":
                    break
                k += 1
            if closed:
                self._adv(min(k, self.n))
            elif multi:
                if report:
                    self._report("unterminated", f"字符串 {op} 从这里开始，一直没闭合", start_line, start_col)
                self._adv(self.n)
            else:
                if report:
                    self._report("unterminated", f"字符串 {op} 从这里开始，到行尾都没闭合", start_line, start_col)
                self._adv(min(k, self.n))
            return True
        return False

    def _try_rust_raw(self) -> bool:
        """r"..." / r#"..."# / br"..." / br##"..."##"""
        src, i, n = self.src, self.i, self.n
        j = i
        if src.startswith("b", j):
            j += 1
        if not src.startswith("r", j):
            return False
        j += 1
        hashes = 0
        while j < n and src[j] == "#":
            hashes += 1
            j += 1
        if j >= n or src[j] != '"':
            return False
        closer = '"' + "#" * hashes
        k = src.find(closer, j + 1)
        self._adv(n if k < 0 else k + len(closer))
        return True

    def _try_cpp_raw(self) -> bool:
        """R"delim( ... )delim" """
        src, i = self.src, self.i
        for pre in ("u8R\"", 'uR"', 'UR"', 'LR"', 'R"'):
            if not src.startswith(pre, i):
                continue
            k = i + len(pre)
            p = src.find("(", k)
            if p < 0 or p - k > 16:
                continue
            delim = src[k:p]
            closer = ")" + delim + '"'
            e = src.find(closer, p + 1)
            self._adv(self.n if e < 0 else e + len(closer))
            return True
        return False

    def _try_rust_char(self) -> bool:
        """'a' / '\\n' / '\\'' 是字符字面量；'a / 'static 是生命周期"""
        if self.src[self.i] != "'":
            return False
        k = self.i + 1
        if k < self.n and self.src[k] == "\\":
            k += 2
        else:
            k += 1
        if k < self.n and self.src[k] == "'":
            self._adv(k + 1)
            return True
        # 生命周期：只吃掉这个引号，不当作字符串
        self._adv(self.i + 1)
        return True

    def _try_js_regex(self) -> bool:
        """在 '/' 位置判断它是不是正则字面量"""
        src, i, n = self.src, self.i, self.n
        j = i - 1
        while j >= 0 and src[j] in " \t":
            j -= 1
        if j >= 0:
            if src[j].isalnum() or src[j] in "_$":
                k = j
                while k >= 0 and (src[k].isalnum() or src[k] in "_$"):
                    k -= 1
                if src[k + 1:j + 1] not in REGEX_PREV_WORDS:
                    return False
            elif src[j] not in REGEX_PREV_CHARS:
                return False
        k = i + 1
        in_class = False
        while k < n:
            ch = src[k]
            if ch == "\\":
                k += 2
                continue
            if ch == "\n":
                return False
            if in_class:
                if ch == "]":
                    in_class = False
            elif ch == "[":
                in_class = True
            elif ch == "/":
                k += 1
                break
            k += 1
        else:
            return False
        while k < n and src[k].isalpha():
            k += 1
        self._adv(k)
        return True

    # ---------- 模板串插值（JS `${...}` 内是完整表达式，可嵌套串与模板） ----------

    def _try_regex_at(self, k: int):
        """在插值内部 k 位置按正则字面量跳过；不是正则返回 None。

        复用主扫描的 `_try_js_regex`（同一套"前一个有效字符 / 关键字"判据），
        临时借用 self.i 并原样还原——`_adv` 只动 i/line/ls，且正则不可能跨行。
        """
        save = (self.i, self.line, self.ls)
        self.i = k
        try:
            return self.i if self._try_js_regex() else None
        finally:
            self.i, self.line, self.ls = save

    def _scan_interp(self, k: int, line: int, ls: int):
        """从 `${` 的 `{` 之后扫到配对 `}`。返回 (其后位置, 行号, 行首偏移,
        括号事件表)——插值是真实 JS 表达式，其中的 ( ) [ ] { } 参与平衡，
        事件带回主扫描按正常括号处理；嵌套模板的文本部分不产生事件。

        插值内必须按 JS 词法逐层吃：注释、正则字面量、字符串都要先跳过。
        否则 `${ s.replace(/'/g, "") }` 里正则内的那个 `'` 会被当成字符串开头，
        一路吃到 EOF 并顺手吞掉插值的闭 `}`——合法代码被判成"字符串没闭合 +
        括号没闭合"（A-08 实测：完全合法的 .js 被误报 3 处）。
        """
        src, n = self.src, self.n
        d = 1
        events = []
        while k < n and d:
            ch = src[k]
            if ch == "{":
                d += 1
                events.append(("{", line, k - ls + 1))
                k += 1
            elif ch == "}":
                d -= 1
                if d:
                    events.append(("}", line, k - ls + 1))
                k += 1
            elif ch in "\"'":
                q = ch
                k += 1
                # JS 单/双引号串不允许裸换行，但 `\` + 换行是合法的行继续。
                # 撞到裸换行即视为未闭合：停在那里交给外层处理，绝不再吃到文件尾。
                while k < n and src[k] != q and src[k] != "\n":
                    if src[k] == "\\":
                        if src[k + 1:k + 2] == "\n":
                            line += 1
                            ls = k + 2
                        k += 2
                    else:
                        k += 1
                if k < n and src[k] == q:
                    k += 1
            elif ch == "/":
                if src.startswith("//", k):
                    e = src.find("\n", k)
                    k = n if e < 0 else e
                elif src.startswith("/*", k):
                    e = src.find("*/", k + 2)
                    if e < 0:
                        k = n
                    else:
                        newlines = src.count("\n", k, e)
                        if newlines:
                            line += newlines
                            ls = src.rfind("\n", k, e) + 1
                        k = e + 2
                elif self.p.get("regex"):
                    end = self._try_regex_at(k)
                    k = end if end is not None else k + 1
                else:
                    k += 1
            elif ch == "`":
                k, line, ls = self._scan_nested_template(k, line, ls)
            elif ch in "([":
                events.append((ch, line, k - ls + 1))
                k += 1
            elif ch in ")]":
                events.append((ch, line, k - ls + 1))
                k += 1
            else:
                if ch == "\n":
                    line += 1
                    ls = k + 1
                k += 1
        return k, line, ls, events

    def _scan_nested_template(self, k: int, line: int, ls: int):
        """跳过一个嵌套模板串（其内部还可再含 `${...}` 插值）。返回
        (其后位置, 行号, 行首偏移)——文本与嵌套插值整体作为原子跳过。"""
        src, n = self.src, self.n
        k += 1
        while k < n:
            ch = src[k]
            if ch == "\\":
                k += 2
            elif ch == "`":
                k += 1
                break
            elif src.startswith("${", k):
                k, line, ls, _ = self._scan_interp(k + 2, line, ls)
            else:
                if ch == "\n":
                    line += 1
                    ls = k + 1
                k += 1
        return k, line, ls

    def _try_js_template(self) -> bool:
        """JS 模板串：`` `...${expr}...` ``——文本部分跳到闭 backtick，插值整体
        递归扫描并把其中的括号事件回放到主平衡栈。没有这层，
        `` `${a.map(x=>`<i>${x}</i>`)}（` `` 这类嵌套会把串边界认错。"""
        if self.src[self.i] != "`":
            return False
        start_line, start_col = self.line, self._col()
        src, n = self.src, self.n
        k, line, ls = self.i + 1, self.line, self.ls
        events = []
        while k < n:
            ch = src[k]
            if ch == "\\":
                k += 2
                continue
            if ch == "`":
                k += 1
                break
            if src.startswith("${", k):
                k, line, ls, ev = self._scan_interp(k + 2, line, ls)
                events.extend(ev)
                continue
            if ch == "\n":
                line += 1
                ls = k + 1
            k += 1
        else:
            self._report("unterminated", "字符串 ` 从这里开始，一直没闭合", start_line, start_col)
            self._adv(n)
            for c, ln, cl in events:
                self._replay_bracket(c, ln, cl)
            return True
        self._adv(k)
        for c, ln, cl in events:
            self._replay_bracket(c, ln, cl)
        return True

    def _replay_bracket(self, c: str, ln: int, cl: int) -> None:
        """把插值内捕获的括号事件按主循环同款规则记入平衡栈。"""
        if c in OPEN:
            self.stack.append((c, ln, cl))
        elif c in CLOSE:
            if not self.stack:
                self._report("extra", f"多出来的闭合符 {c}，前面没有对应的 {CLOSE[c]}", ln, cl)
            else:
                op, ol, oc = self.stack[-1]
                if OPEN[op] != c:
                    self.stack.pop()
                    self._report(
                        "mismatch",
                        f"这里是 {c}，但第 {ol} 行第 {oc} 列的 {op} 还没闭合（{op} 应该配 {OPEN[op]}）",
                        ln, cl,
                    )
                else:
                    self.stack.pop()

    # ---------- 主循环 ----------

    def run(self) -> None:
        src, n = self.src, self.n
        p = self.p
        while self.i < n:
            raw = src[self.i]

            if raw == "\n":
                self.depth.append((self.line, len(self.stack)))
                self.i += 1
                self.line += 1
                self.ls = self.i
                continue
            if raw in " \t\r\f\v":
                self.i += 1
                continue

            if self._try_line_comment():
                continue
            if self._try_block_comment():
                continue
            if p.get("rust_raw") and self._try_rust_raw():
                continue
            if p.get("cpp_raw") and self._try_cpp_raw():
                continue
            if p.get("rust_char") and raw == "'" and self._try_rust_char():
                continue
            if p.get("template_interp") and raw == "`" and self._try_js_template():
                continue
            if self._try_string():
                continue
            if p.get("regex") and raw == "/" and self._try_js_regex():
                continue

            if self.fw and (raw in FULLWIDTH_BRACKET or raw in FULLWIDTH_PUNCT
                            or raw in SMART_QUOTES):
                if raw in FULLWIDTH_BRACKET:
                    c = FULLWIDTH_BRACKET[raw]
                    self._report("fullwidth", f"全角 {raw} 混进了代码，应写成半角 {c}")
                elif raw in FULLWIDTH_PUNCT:
                    self._report("fullwidth", f"全角 {raw} 混进了代码，应写成半角 {FULLWIDTH_PUNCT[raw]}")
                    self.i += 1
                    continue
                else:
                    self._report("smartquote", f"弯引号 {raw} 混进了代码，应写成半角 {SMART_QUOTES[raw]}")
                    self.i += 1
                    continue
            else:
                c = raw

            if c in OPEN and c in self.pairs:
                self.stack.append((c, self.line, self._col()))
            elif c in CLOSE and CLOSE[c] in self.pairs:
                if not self.stack:
                    self._report("extra", f"多出来的闭合符 {c}，前面没有对应的 {CLOSE[c]}")
                else:
                    op, ol, oc = self.stack[-1]
                    if OPEN[op] != c:
                        self.stack.pop()
                        self._report(
                            "mismatch",
                            f"这里是 {c}，但第 {ol} 行第 {oc} 列的 {op} 还没闭合（{op} 应该配 {OPEN[op]}）",
                        )
                    else:
                        self.stack.pop()
            self.i += 1

        if self.stack:
            self.depth.append((self.line, len(self.stack)))

    # ---------- 结果 ----------

    @property
    def ok(self) -> bool:
        return not self.issues and not self.stack

    def result(self) -> dict:
        tail = [
            dict(char=ch, line=ln, col=co, closing=OPEN[ch])
            for ch, ln, co in self.stack
        ]
        return dict(
            path=self.path,
            ok=self.ok,
            issues=self.issues,
            open_stack=tail,
            suggestion=" ".join(t["closing"] for t in reversed(tail)),
            depth_trace=[dict(line=l, depth=d) for l, d in self.depth],
        )

    def render(self, show_depth: bool = False) -> str:
        out = []
        head = f"{self.path}"
        if self.ok:
            out.append(_c("32", f"[ok] {head}"))
        else:
            n_bad = len(self.issues) + len(self.stack)
            out.append(_c("31", f"[!!] {head}  —— 发现 {n_bad} 处问题"))

        for iss in self.issues:
            ln, co = iss["line"], iss["col"]
            out.append("")
            out.append(f"     第 {ln} 行 第 {co} 列  {_c('31', iss['msg'])}")
            txt = self._text(ln)
            if txt:
                out.append(f"   {ln:>5} | {txt}")
                out.append(f"   {'':>5} | {_c('31', ' ' * (co - 1) + '^')}")

        if self.stack:
            out.append("")
            out.append(f"     文件结束时还有 {len(self.stack)} 个括号没闭合：")
            for t in self.result()["open_stack"]:
                txt = self._text(t["line"])
                out.append(
                    f"     第 {t['line']} 行 第 {t['col']} 列  "
                    f"{_c('33', t['char'])}  ->  少了 {t['closing']}"
                )
                if txt:
                    out.append(f"   {t['line']:>5} | {txt}")
                    out.append(f"   {'':>5} | {_c('33', ' ' * (t['col'] - 1) + '^')}")
            out.append("")
            out.append(
                _c("36", f"     建议：在文件末尾按顺序补上  {self.result()['suggestion']}")
            )

        if show_depth and self.depth:
            out.append("")
            out.append("     嵌套深度轨迹（行 -> 该行结束时的深度）：")
            row = []
            for ln, d in self.depth:
                row.append(f"{ln}:{d}")
            for k in range(0, len(row), 12):
                out.append("       " + "  ".join(row[k:k + 12]))

        if self._capped:
            out.append("")
            out.append(_c("90", f"     （超过 {self.max_issues} 处，后续已省略；先修第一处）"))
        return "\n".join(out)


def detect_lang(path: str, src: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext in EXT_MAP:
        return EXT_MAP[ext]
    first = src.split("\n", 1)[0]
    if first.startswith("#!"):
        for key, prof in (
            ("python", "py"), ("node", "js"), ("deno", "js"), ("bash", "sh"),
            ("sh", "sh"), ("zsh", "sh"), ("ruby", "rb"), ("php", "php"), ("lua", "lua"),
            ("pwsh", "ps1"), ("powershell", "ps1"),
        ):
            if key in first:
                return prof
    return "plain"


def read_text(path: str) -> str:
    if path == "-":
        data = sys.stdin.buffer.read()
    else:
        with open(path, "rb") as fh:
            data = fh.read()
    for enc in ("utf-8-sig", "gbk", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", "replace")


# --------------------------------------------------------------------------
# 自检
# --------------------------------------------------------------------------

_SELFTEST = [
    ("py 正常", "py", "def f(a):\n    return [1, 2]\n", True),
    ("py 少闭合", "py", "def f(a):\n    return [1, 2\n", False),
    ("py 类型不符", "py", "x = (1]\n", False),
    ("py 多余闭合", "py", "x = 1)\n", False),
    ("py 全角括号", "py", "x = \uff081 + 2\uff09\n", False),
    ("py 字符串里的全角（合法）", "py", "s = \"\uff08\u4e2d\u6587\uff09\"\n", True),
    ("py 注释里的括号（合法）", "py", "# \u8fd9\u662f ( \u6ce8\u91ca\nx = 1\n", True),
    ("py 三引号串", "py", "x = \"\"\"\n) ] (\n\"\"\"\n", True),
    ("py 三引号未闭合", "py", "x = \"\"\"\nabc\n", False),
    ("py 单引号未闭合", "py", "x = 'abc\n", False),
    ("js 模板串+正则", "js", "const r = /[a-z)]/g;\nconst s = `a${1 + 2}b`;\nconsole.log(r, s);\n", True),
    ("js 少右花括号", "js", "function f() {\n  return 1;\n", False),
    ("js 除号不误判为正则", "js", "const a = 10 / 2;\nconst b = a / 5;\n", True),
    ("js 正则里的括号", "js", "if (x) { y = z.replace(/\\)/g, ''); }\n", True),
    # A-08（2026-09-23）：插值内必须先按 JS 词法吃掉正则/注释，并给引号串加换行边界
    ("js 插值内正则含引号（合法）", "js",
     "const s = `${items.map(s => s.replace(/'/g, \"\")).join(\",\")}`;\n", True),
    ("js 插值内注释含引号（合法）", "js", "const t = `${ /* don't */ x }`;\n", True),
    ("js 插值内除号（合法）", "js", "const u = `${a / 2}`;\n", True),
    # 插值内"引号没配对"刻意只在 L1 停住、不报错：若在这里报，合法 JSX 文案里的撇号
    # （如 `${<Foo>don't</Foo>}`）就会被误伤。真语法问题归 L2（ast / node --check）。
    ("js 插值内串跨行（交 L2 判语法）", "js", "const t = `${ \"abc\n}`;\n", True),
    ("go 反引号原串", "go", "s := `a)b(c`\n", True),
    ("go 少右花括号", "go", "func f() {\n\tx := []int{1, 2\n", False),
    ("rust 生命周期与字符", "rust", "fn f<'a>(x: &'a str) -> char { 'x' }\n", True),
    ("rust raw string", "rust", "let s = r#\"a)b\"#;\n", True),
    ("rust 嵌套块注释", "rust", "/* a /* b */ c */\nfn f() -> i32 { 1 }\n", True),
    ("rust 真错", "rust", "fn main() {\n    let v = vec![1, 2;\n}\n", False),
    ("c 块注释未闭合", "c", "int main() {\n/* oops\n}\n", False),
    ("cpp raw string", "c", "auto s = R\"(a)b\";\nint x = 1;\n", True),
    ("弯引号", "py", "x = \u201ca\u201d\n", False),
    ("全角分号", "js", "let x = 1\uff1b\n", False),
    ("json 正常", "json", '{"a": [1, 2], "b": {"c": 3}}\n', True),
    ("json 缺右括号", "json", '{"a": [1, 2}\n', False),
    ("ps1 正常", "ps1", "function Get-Test { Write-Output 'OK' }\n", True),
    ("ps1 块注释与括号失衡", "ps1", "<# comment #> function Get-Test { Write-Output [1, 2\n", False),
]


def selftest() -> int:
    bad = 0
    for name, lang, src, want_ok in _SELFTEST:
        sc = Scanner(src, PROFILES[lang], f"<{lang}>")
        sc.run()
        got = sc.ok
        flag = "PASS" if got == want_ok else "FAIL"
        if got != want_ok:
            bad += 1
        detail = ""
        if got != want_ok:
            detail = "  " + "; ".join(
                f"{i['line']}:{i['col']} {i['kind']}" for i in sc.issues
            )
            if sc.stack:
                detail += "  unclosed=" + "".join(t[0] for t in sc.stack)
        print(f"  {flag}  {name:<26} 期望={'通过' if want_ok else '报错'} 实得={'通过' if got else '报错'}{detail}")
    total = len(_SELFTEST)
    print(f"\n  {total - bad}/{total} 通过")
    return 0 if bad == 0 else 1


def _init_stdout():
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


def main(argv=None) -> int:
    global USE_COLOR
    _init_stdout()
    ap = argparse.ArgumentParser(
        prog="bracket_lint",
        description="语言感知的括号平衡检查器：精确定位未闭合括号的起始行列。",
    )
    ap.add_argument("paths", nargs="*", help="要检查的文件；用 - 表示从标准输入读")
    ap.add_argument("--lang", default="auto", help="强制语言 profile（默认按扩展名自动判断）")
    ap.add_argument("--json", action="store_true", help="输出 JSON，便于程序/agent 消费")
    ap.add_argument("--quiet", action="store_true", help="只输出有问题的文件")
    ap.add_argument("--depth", action="store_true", help="额外打印每行结束时的嵌套深度")
    ap.add_argument("--max", type=int, default=5, help="每个文件最多报告几处（默认 5）")
    ap.add_argument("--color", choices=["auto", "always", "never"], default="auto")
    ap.add_argument("--selftest", action="store_true", help="跑内置用例，验证本工具自身是否可信")
    ap.add_argument("--version", action="version", version=f"bracket_lint {__version__}")
    args = ap.parse_args(argv)

    if args.color == "always":
        USE_COLOR = True
    elif args.color == "auto":
        USE_COLOR = sys.stdout.isatty() and os.environ.get("NO_COLOR") is None

    if args.selftest:
        return selftest()

    if not args.paths:
        ap.print_help()
        return 2

    if args.lang != "auto":
        if args.lang not in PROFILES:
            print(f"未知语言 profile: {args.lang}（可选：{', '.join(sorted(PROFILES))}）", file=sys.stderr)
            return 2

    results = []
    for path in args.paths:
        try:
            src = read_text(path)
        except OSError as exc:
            print(f"读不了 {path}: {exc}", file=sys.stderr)
            return 2
        label = "<stdin>" if path == "-" else path
        lang = args.lang if args.lang != "auto" else detect_lang(label, src)
        sc = Scanner(src, PROFILES[lang], label, max_issues=args.max)
        sc.run()
        res = sc.result()
        res["lang"] = lang
        res["_scanner"] = sc
        results.append(res)

    if args.json:
        payload = []
        for r in results:
            r2 = {k: v for k, v in r.items() if k != "_scanner"}
            payload.append(r2)
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        bad_files = [r for r in results if not r["ok"]]
        for r in results:
            if r["ok"] and args.quiet:
                continue
            print(r["_scanner"].render(show_depth=args.depth))
            print()
        if not args.quiet:
            n_ok = len(results) - len(bad_files)
            print(f"  {n_ok} 个通过，{len(bad_files)} 个有问题")

    return 1 if any(not r["ok"] for r in results) else 0


if __name__ == "__main__":
    _init_stdout()
    sys.exit(main())
