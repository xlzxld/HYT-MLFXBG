#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""structure_guard.py - AI 代码结构完整性统一门禁入口（零依赖，单文件，Python 3.8+）

为什么需要它
------------
bracket_lint 解决"括号在哪一行失衡"，但 HYT-CAD 坑 #75 证明：括号总数平衡的
代码照样可能是坏的（setq 少一个右括号后 else 分支被吞进参数表，平衡检查全部
漏过）。本脚本把五层检测编排到一个入口，供四道闸共用：

  L0 编码层   二进制探测（NUL 字节跳过）、文档扩展名跳过、编码解码链
              （utf-8-sig -> gbk -> latin-1，复用 bracket_lint.read_text）
  L1 平衡层   语言感知括号栈（复用 bracket_lint.Scanner，含全角/弯引号检测）
  L2 结构层   权威语法检查（按扩展名分派，工具缺失自动降级并标注，不假绿）：
                .py   ast.parse（缩进错乱 TabError/IndentationError 在此拦截）
                .json json.loads；.toml tomllib（Python 3.11+，缺失降级）
                .js/.mjs/.cjs  node --check（.ts/.jsx 不送 node，防误报）
                .sh   bash -n；.rb ruby -c；.go gofmt -e
                .ps1  PowerShell Parser::ParseFile（Windows 自带）
  L3 形态层   调用形态元数与顶层不变量（LISP 家族，坑 #75/#73 泛化）：
                setq 参数须偶数 / if 2~3 段 / foreach >=3 段 / defun 须顶层
  L4 断言层   项目专属回归断言：**仅当显式传 `--asserts <路径>` 时加载**，
              约定暴露 run(paths) -> [(path, line, msg), ...]（check_audit_fixes 模式）。
              不再从项目根自动加载 guard_asserts.py——仓库内容不构成执行授权：
              任何 PR 都不该能让开发者本机的提交钩子执行任意代码（收官体检 A-01）

四道闸共用本入口
----------------
  闸1 AI 编辑后   agent hook（--from-hook：stdin JSON 取文件路径，有问题退出 2）
  闸2 git 提交时  install_hook.py 生成的 pre-commit（--staged 或文件表，
                  调用项目内落盘件 .agents/structure_guard.py，可移植）
  闸3 CI          全仓扫描（gate.yml 的结构门禁步骤，读项目内落盘件）
  闸4 交付前      AGENTS.md §2 静态检查行 / tsc verify

分发形态（tsc-managed）
-----------------------
  本文件带 tsc-managed 标记，随 tsc install / sync 作为执法包落盘件复制到
  项目 .agents/structure_guard.py（连同 bracket_lint.py）。技能目录里的是
  正本；项目里的落盘件供 git 钩子与 CI 使用（它们只认项目自己的文件）。
  上游正本更新后，落盘件由 tsc-managed 托管机制自动跟进；项目删掉标记行
  即视为接管，技能不再覆盖。

退出码契约（与 tsc.py 对齐）
---------------------------
  0  全部通过
  1  代码结构问题（闸 2/3 据此拦截）
  2  工具自身故障（钩子告警放行，不冒充代码问题；CI 层 fail-closed 兜底）
  3  配置非法（如 --lang 给了不认识的值）

行内豁免（误报逃生口）
---------------------
  在问题所在行写注释 guard:skip 豁免该行全部问题；guard:skip=code1,code2
  只豁免指定问题码。豁免会在输出中计数显示，不静默吞掉。

用法
----
    python structure_guard.py src/app.py                 # 查指定文件
    python structure_guard.py scripts/                   # 查目录（递归，跳过 .git 等）
    python structure_guard.py --staged                   # 查 git 暂存文件
    python structure_guard.py --json src/                # JSON 输出（AI 消费）
    echo '{"tool_input":{"file_path":"a.py"}}' | python structure_guard.py --from-hook
    python structure_guard.py --selftest

设计要点
--------
  - 权威工具缺失 = "降级并标注"，绝不因环境缺工具而报通过（降级写进结果）。
  - **严格模式（--strict）**：文件截断 / 超大文件降级 / 未知语言 / 工具缺失中的
    任何一项都会让退出码非 0（默认 2）——CI / pre-commit / 聚合门禁都应带此开关，
    "没真正查完"永远不等于"通过"。`--allow-missing-tools` 显式放行工具缺席这一类
    降级（明确配置允许），其余降级类别照拦。本地人工模式（不带 --strict）只给
    warning，但汇总行会如实标注"检查不完整"，绝不输出"完整检查通过"。
  - 闸1（--from-hook）fail-open：stdin 解析失败、工具故障一律放行（闸2/3 兜底）；
    只有确凿的结构问题才退出 2 拦截，防止钩子自身变成事故源。文件数超过上限时
    同样放行但**明确告警**（静默截断 = 假绿，闸2/3 兜底）。
  - 子进程判定以**退出码**为准；stderr 仅作诊断信息（有些合法命令会往 stderr 打
    warning，"stderr 有内容"本身不是失败信号）。
  - 超大文件（>2MB）只跑 L0/L1：深层解析对巨型生成物的耗时与误报都不划算；
    该降级在严格模式下必须拦下。
"""

# tsc-managed —— 落盘件：随 tsc install / sync 复制到项目 .agents/，由技能托管更新；删除本行即视为项目接管。
from __future__ import annotations

import argparse
import ast
import json
import os
import re
import shutil
import subprocess
import sys

__version__ = "1.3.0"

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import bracket_lint  # noqa: E402  （与本脚本同目录，作为 L1 引擎）

MAX_FILES = 1000          # 一次扫描的文件数上限，超出截断并告警
MAX_HOOK_FILES = 20       # 闸1 单次最多检查的文件数
MAX_MSG = 200             # 单条消息最长字符
BIG_FILE = 2 * 1024 * 1024  # 超过此字节数只跑 L0/L1（深层解析对巨型生成物不划算）

# 这些扩展名是文档/数据/标记模板，不做结构检查（中文散文与 HTML 文案里的
# 全角标点是合法内容；.vue/.html 等标记文件的引号语义由框架解析，平衡检查必误报）
DOC_EXTS = {".md", ".txt", ".rst", ".adoc", ".log", ".csv", ".ini", ".cfg",
            ".vue", ".svelte", ".html", ".htm", ".astro"}
# 递归扫描目录时跳过
SKIP_DIRS = {".git", ".hg", ".svn", "__pycache__", "node_modules",
             ".venv", "venv", ".tox", "dist", "build", ".idea", ".vscode"}

# bracket_lint 未覆盖的语言 profile 与扩展名。
# lisp 字符串允许字面跨行（AutoLISP/Elisp 语义，HYT-CAD wx_runner.lsp 实证：
# strcat " 换行续写是合法源码，multi=False 会把它报成 unterminated 误报）
LISP_PROFILE = dict(line=[";"], block=[], strings=[('"', '"', True, True, True)])
EXTRA_PROFILES = {"lisp": LISP_PROFILE}
EXTRA_EXT = {".lsp": "lisp", ".lisp": "lisp", ".el": "lisp", ".cl": "lisp"}

USE_COLOR = False


def _c(code: str, text: str) -> str:
    return f"\033[{code}m{text}\033[0m" if USE_COLOR else text


# --------------------------------------------------------------------------
# L3：LISP 家族形态检查（坑 #75 / 坑 #73 泛化，移植自 HYT-CAD check_sexpr /
# check_defun_depth，语义保持一致）
# --------------------------------------------------------------------------

class _LispError(Exception):
    pass


def _lisp_tokenize(src):
    """token: (kind, value, line, col)；kind in '(' ')' 'atom' 'str'。
    字符串可跨行（AutoLISP 语义），token 记起始行列，行号随扫描推进不漂移。"""
    toks = []
    i, n = 0, len(src)
    line, ls = 1, 0
    in_s = in_c = False
    while i < n:
        ch = src[i]
        if ch == "\n":
            line += 1
            ls = i + 1
            in_c = False
            i += 1
            continue
        if in_c:
            i += 1
            continue
        if in_s:
            if ch == "\\":
                i += 2
                continue
            if ch == '"':
                in_s = False
            i += 1
            continue
        if ch == ";":
            in_c = True
            i += 1
            continue
        if ch == '"':
            start_line, start_col = line, i - ls + 1
            j = i + 1
            while j < n:
                if src[j] == "\\":
                    j += 2
                    continue
                if src[j] == '"':
                    break
                if src[j] == "\n":
                    line += 1
                    ls = j + 1
                j += 1
            toks.append(("str", src[i:min(j + 1, n)], start_line, start_col))
            i = min(j + 1, n)
            continue
        if ch in "()":
            toks.append((ch, None, line, i - ls + 1))
            i += 1
            continue
        if ch.isspace():
            i += 1
            continue
        j = i
        while j < n and not src[j].isspace() and src[j] not in '();"':
            j += 1
        toks.append(("atom", src[i:j], line, i - ls + 1))
        i = j
    return toks


def _lisp_parse(toks):
    """token -> 节点树；节点 = ('atom'|'str', text, line, col) 或 list。
    引用糖 'X 与 X 视为同一参数（粘合），避免 (setq x '(1 2)) 误报奇数。"""
    pos = [0]

    def rd():
        if pos[0] >= len(toks):
            raise _LispError("unexpected end of input")
        kind, val, ln, cl = toks[pos[0]]
        pos[0] += 1
        if kind in ("atom", "str"):
            if kind == "atom" and val == "'" and pos[0] < len(toks):
                return rd()
            return (kind, val, ln, cl)
        if kind == "(":
            lst = []
            while pos[0] < len(toks):
                if toks[pos[0]][0] == ")":
                    pos[0] += 1
                    return lst
                lst.append(rd())
            raise _LispError("unbalanced ( at line %d" % ln)
        raise _LispError("unexpected ) at line %d" % ln)

    out = []
    while pos[0] < len(toks):
        out.append(rd())
    return out


def _lisp_head(node):
    """(head_atom, line, col, arg_count)；表头不是原子时不做元数判断。"""
    if isinstance(node, list) and node and isinstance(node[0], tuple):
        head, _, ln, cl = node[0]
        if head == "atom":
            return head and node[0][1].lower(), ln, cl, len(node) - 1
    return None, 0, 0, 0


def _lisp_walk(node, issues):
    if isinstance(node, tuple):
        return
    head, ln, cl, args = _lisp_head(node)
    if head == "setq" and args % 2 != 0:
        issues.append(("L3", ln, cl, "setq-odd-args",
                       "setq 参数个数为奇数(%d)——缺右括号或符号/值不配对" % args))
    elif head == "if" and args not in (2, 3):
        issues.append(("L3", ln, cl, "if-arity",
                       "if 参数个数为 %d（应为 2 或 3）" % args))
    elif head == "foreach" and args < 3:
        issues.append(("L3", ln, cl, "foreach-arity",
                       "foreach 参数个数为 %d（应 >=3）" % args))
    for x in node:
        _lisp_walk(x, issues)


def _lisp_defun_top(src):
    """每个 (defun 所在行"开始时"的括号深度必须为 0（坑 #73）。
    行"结束"时深度 +1 属正常（defun 跨行）；开始时 >0 = 嵌套在别的表达式里。"""
    bad = []
    bal, in_s, in_c, esc, ln = 0, False, False, False, 0
    for line in src.split("\n"):
        ln += 1
        start = bal
        for ch in line:
            if in_c:
                continue
            if in_s:
                if esc:
                    esc = False
                elif ch == "\\":
                    esc = True
                elif ch == '"':
                    in_s = False
                continue
            if ch == ";":
                in_c = True
            elif ch == '"':
                in_s = True
            elif ch == "(":
                bal += 1
            elif ch == ")":
                bal -= 1
        in_c = False
        esc = False
        s = line.strip()
        if s.startswith("(defun") and start != 0 and not s.startswith("(defun *error*"):
            bad.append((ln, start))
    return bad


def check_lisp_forms(src):
    """返回 issue 列表；括号/引号结构坏到无法解析时交给 L1 报告，此处静默跳过。"""
    issues = []
    try:
        tree = _lisp_parse(_lisp_tokenize(src))
    except _LispError:
        return issues
    for form in tree:
        _lisp_walk(form, issues)
    for ln, depth in _lisp_defun_top(src):
        issues.append(("L3", ln, 1, "nested-defun",
                       "defun 定义在第 %d 层嵌套内，加载时不会成为全局函数（须在顶层）" % depth))
    return issues


# --------------------------------------------------------------------------
# L2：权威语法检查（按扩展名分派；返回 (issue|None, tool_error|None, degraded|None)）
# --------------------------------------------------------------------------

def _run_cmd(cmd, timeout=30):
    """统一子进程执行。返回 (rc, stdout, stderr, tool_err)。"""
    try:
        p = subprocess.run(cmd, capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=timeout)
        return p.returncode, p.stdout or "", p.stderr or "", None
    except FileNotFoundError:
        return None, "", "", "%s 不可用" % cmd[0]
    except subprocess.TimeoutExpired:
        return None, "", "", "%s 执行超时(%ds)" % (cmd[0], timeout)
    except OSError as e:
        return None, "", "", "%s 执行失败: %s" % (cmd[0], e)


def _mk_issue(layer, line, col, code, msg):
    return {"layer": layer, "line": line or 1, "col": col or 1,
            "code": code, "msg": (msg or "").strip()[:MAX_MSG]}


def check_python(path, src):
    try:
        ast.parse(src, filename=path)
        return None, None, None
    except SyntaxError as e:
        # IndentationError / TabError 是 SyntaxError 子类：缩进错乱在此拦截
        return _mk_issue("L2", e.lineno, e.offset, "syntax-error",
                         e.msg or "语法错误"), None, None
    except ValueError as e:
        return _mk_issue("L2", 1, 1, "syntax-error", "无法解析: %s" % e), None, None
    except Exception as e:  # 解析器自身异常：走 tool_error 通道，不算代码问题
        return None, "Python ast 解析异常: %r" % (e,), None


def check_json(path, src):
    try:
        json.loads(src)
        return None, None, None
    except json.JSONDecodeError as e:
        line = src.count("\n", 0, e.pos) + 1
        col = e.pos - (src.rfind("\n", 0, e.pos) + 1) + 1
        return _mk_issue("L2", line, col, "json-error", e.msg), None, None
    except Exception as e:
        return None, "JSON 解析异常: %r" % (e,), None


def check_toml(path, src):
    try:
        import tomllib
    except ImportError:
        return None, None, "Python <3.11 无 tomllib，.toml 仅做平衡检查"
    try:
        tomllib.loads(src)
        return None, None, None
    except tomllib.TOMLDecodeError as e:
        return _mk_issue("L2", 1, 1, "toml-error", str(e)), None, None
    except Exception as e:
        return None, "TOML 解析异常: %r" % (e,), None


def check_yaml(path, src):
    try:
        import yaml  # PyYAML 非标准库，缺席时降级为平衡检查
    except ImportError:
        return None, None, "无 PyYAML，.yaml 仅做平衡检查"
    try:
        yaml.safe_load(src)
        return None, None, None
    except yaml.YAMLError as e:
        mark = getattr(e, "problem_mark", None)
        line = mark.line + 1 if mark else 1
        col = mark.column + 1 if mark else 1
        return _mk_issue("L2", line, col, "yaml-error",
                         str(getattr(e, "problem", None) or e)), None, None
    except Exception as e:
        return None, "YAML 解析异常: %r" % (e,), None


def _subprocess_check(path, cmd, parse_stderr, toolname):
    """通用外部命令检查：以**退出码**为唯一失败判据，stderr 仅作诊断。

    合法命令也可能往 stderr 打 warning（旧版"stderr 非空即失败"会把它们误判成
    语法错误——假红与假绿同样不可接受）。退出码非 0 才是问题，此时解析
    stderr / stdout 提取行列信息。"""
    rc, out, err, terr = _run_cmd(cmd)
    if terr:
        return None, None, "%s 不可用，跳过 %s 深检" % (terr, toolname)
    if rc == 0:
        return None, None, None
    line, col, msg = parse_stderr(err or out or ("退出码 %s" % rc))
    return _mk_issue("L2", line, col, "syntax-error", "%s: %s" % (toolname, msg)), None, None


def check_node(path, src):
    node = shutil.which("node")
    if not node:
        return None, None, "node 不可用，JS 仅做平衡检查"

    def parse(err):
        first = next((l for l in err.splitlines() if l.strip()), "")
        m = re.search(r":(\d+)$", first)
        mline = re.search(r"SyntaxError: (.*)", err)
        return (int(m.group(1)) if m else 1), 1, (mline.group(1) if mline else first[:MAX_MSG])

    return _subprocess_check(path, [node, "--check", path], parse, "node --check")


def check_bash(path, src):
    bash = shutil.which("bash")
    if not bash:
        return None, None, "bash 不可用，shell 仅做平衡检查"

    def parse(err):
        first = next((l for l in err.splitlines() if l.strip()), "")
        m = re.search(r"line (\d+)", first)
        return (int(m.group(1)) if m else 1), 1, first[:MAX_MSG]

    issue, terr, deg = _subprocess_check(path, [bash, "-n", path], parse, "bash -n")
    # shellcheck（有则增强）：bash -n 只查语法，error 级的语义结构错误靠它补
    sc = shutil.which("shellcheck")
    if sc is not None:
        rc, out, err, terr2 = _run_cmd([sc, "-S", "error", "-f", "gcc", path])
        if terr2 is None and rc != 0:
            first = next((l for l in out.splitlines() if l.strip()), "")
            m = re.search(r":(\d+):(\d+):\s*(.*)$", first)
            sc_issue = _mk_issue("L2", int(m.group(1)) if m else 1,
                                 int(m.group(2)) if m else 1, "syntax-error",
                                 "shellcheck: %s" % (m.group(3) if m else first[:MAX_MSG]))
            return sc_issue, terr, deg
    return issue, terr, deg


def check_ruby(path, src):
    ruby = shutil.which("ruby")
    if not ruby:
        return None, None, "ruby 不可用，仅做平衡检查"

    def parse(err):
        first = next((l for l in err.splitlines() if l.strip()), "")
        m = re.search(r":(\d+):\s*(.*)$", first)
        return (int(m.group(1)) if m else 1), 1, (m.group(2) if m else first)[:MAX_MSG]

    return _subprocess_check(path, [ruby, "-c", path], parse, "ruby -c")


def check_gofmt(path, src):
    gofmt = shutil.which("gofmt")
    if not gofmt:
        return None, None, "gofmt 不可用，Go 仅做平衡检查"

    def parse(err):
        first = next((l for l in err.splitlines() if l.strip()), "")
        m = re.search(r":(\d+):(\d+):\s*(.*)$", first)
        if m:
            return int(m.group(1)), int(m.group(2)), m.group(3)
        return 1, 1, first[:MAX_MSG]

    # -e 报告全部错误；-l 只列文件名不影响判定；stdout（格式化输出）不参与判定
    return _subprocess_check(path, [gofmt, "-e", "-l", path], parse, "gofmt")


_PS_SCRIPT = (
    "$errs=$null; $toks=$null; "
    "[System.Management.Automation.Language.Parser]::ParseFile("
    "'{path}', [ref]$toks, [ref]$errs) | Out-Null; "
    "if ($errs -and $errs.Count -gt 0) {{ "
    "$errs | ForEach-Object {{ Write-Output ('{{0}}:{{1}}: {{2}}' -f "
    "$_.Extent.StartLineNumber, $_.Extent.StartColumnNumber, $_.Message) }}; exit 1 }}"
)


def check_powershell(path, src):
    exe = shutil.which("pwsh") or shutil.which("powershell")
    if not exe:
        return None, None, "PowerShell 不可用，.ps1 仅做平衡检查"
    script = _PS_SCRIPT.format(path=path.replace("'", "''"))
    rc, out, err, terr = _run_cmd([exe, "-NoProfile", "-NonInteractive", "-Command", script])
    if terr:
        return None, None, "PowerShell 不可用，.ps1 仅做平衡检查"
    first = next((l for l in out.splitlines() if l.strip()), "")
    m = re.match(r"^(\d+):(\d+):\s*(.+)$", first)
    if rc != 0:  # 退出码为准：rc==0 时 stdout/stderr 的内容只是诊断信息
        msg = m.group(3) if m else (first or "退出码 %s" % rc)
        line = int(m.group(1)) if m else 1
        col = int(m.group(2)) if m else 1
        return _mk_issue("L2", line, col, "ps-error", "PowerShell: %s" % msg), None, None
    return None, None, None


# 需要真实文件的外部命令检查器（自检模式下跳过，进程内检查器照常运行）
for _fn in (check_node, check_bash, check_ruby, check_gofmt, check_powershell):
    _fn.needs_real_file = True

# L2 分派按扩展名（.ts/.jsx 含非 JS 语法，不送 node --check，防误报）
L2_BY_EXT = {
    ".py": check_python,
    ".pyw": check_python, ".pyi": check_python,
    ".json": check_json,
    ".toml": check_toml,
    ".yaml": check_yaml, ".yml": check_yaml,
    ".js": check_node, ".mjs": check_node, ".cjs": check_node,
    ".sh": check_bash, ".bash": check_bash, ".zsh": check_bash,
    ".rb": check_ruby,
    ".go": check_gofmt,
    ".ps1": check_powershell, ".psm1": check_powershell, ".psd1": check_powershell,
}

# L2 语言兜底（用于无扩展名文件/标准输入/显式 --lang 覆盖）
L2_BY_LANG = {
    "py": check_python,
    "json": check_json,
    "toml": check_toml,
    "yaml": check_yaml,
    "js": check_node,
    "sh": check_bash,
    "rb": check_ruby,
    "go": check_gofmt,
    "ps1": check_powershell,
}

SUPPORTED_EXTS = set(bracket_lint.EXT_MAP) | set(EXTRA_EXT)


# --------------------------------------------------------------------------
# 单文件检测主管线
# --------------------------------------------------------------------------

def detect_lang(label, src):
    ext = os.path.splitext(label)[1].lower()
    if ext in EXTRA_EXT:
        return EXTRA_EXT[ext]
    return bracket_lint.detect_lang(label, src)


def _comment_span(line, lang):
    """返回该行**真正注释**的起始列号（1-based）；没有则返回 None。

    单行词法：按语言 profile 的字符串规则跳过字符串内容，注释符出现在字符串外
    才算注释起点。guard:skip 只有落在注释区内才可能生效——写在字符串字面量里的
    `guard:skip` 是数据，不是豁免（否则一行 `s = "guard:skip"` 就能把该行的问题
    全部静默吞掉）。只做单行分析：问题行按定义是代码行，多行串/块注释内部不会
    产生 L1~L3 的机械发现。
    """
    prof = bracket_lint.PROFILES.get(lang) or EXTRA_PROFILES.get(lang) \
        or bracket_lint.PROFILES["plain"]
    opens = sorted({s[0] for s in prof.get("strings", [])}, key=len, reverse=True)
    comment_tokens = sorted(
        set(prof.get("line", [])) | {b[0] for b in prof.get("block", [])},
        key=len, reverse=True)
    close_of = {s[0]: (s[1], s[2]) for s in prof.get("strings", [])}
    i, n = 0, len(line)
    while i < n:
        for tok in comment_tokens:
            if line.startswith(tok, i):
                return i + 1
        for tok in opens:
            if line.startswith(tok, i):
                close, escapes = close_of[tok]
                j = i + len(tok)
                while j < n:
                    if escapes and line[j] == "\\":
                        j += 2
                        continue
                    if line.startswith(close, j):
                        break
                    j += 1
                if j >= n:
                    return None  # 该行结尾仍在字符串里（多行串）：本行没有注释
                i = j + len(close)
                break
        else:
            i += 1
    return None


def _exempt_filter(src, issues, lang):
    """行内豁免：问题行的**注释里**含 guard:skip（或 guard:skip=code1,code2）才豁免。

    返回 (保留的问题, 豁免计数)。标记必须出现在真实注释上下文中——字符串字面量
    里的 guard:skip 不生效（见 _comment_span）。
    """
    lines = src.split("\n")
    kept, dropped = [], 0
    for iss in issues:
        ln = iss.get("line")
        text = lines[ln - 1] if isinstance(ln, int) and 1 <= ln <= len(lines) else ""
        col = _comment_span(text, lang)
        if col is not None:
            m = re.search(r"guard:skip(?:=(\S+))?", text[col - 1:])
            if m:
                codes = set(m.group(1).split(",")) if m.group(1) else None
                if codes is None or iss.get("code") in codes:
                    dropped += 1
                    continue
        kept.append(iss)
    return kept, dropped


def check_source(label, src, lang=None, max_issues=5, real_file=True):
    """对一段源码跑 L0~L3。real_file=False 时跳过需真实文件的外部命令（自检用）。"""
    res = {"path": label, "lang": None, "ok": True, "skipped": None,
           "issues": [], "degraded": [], "tool_error": None, "suggestion": "",
           "exempted": 0}

    lang_explicit = lang is not None
    if lang is None:
        ext = os.path.splitext(label)[1].lower()
        if ext in DOC_EXTS:
            res["skipped"] = "docs"
            return res
        lang = detect_lang(label, src)
    res["lang"] = lang

    big = len(src.encode("utf-8", "replace")) > BIG_FILE
    if big:
        res["degraded"].append("超大文件(>2MB)，仅做 L0/L1 平衡检查")

    prof = bracket_lint.PROFILES.get(lang) or EXTRA_PROFILES.get(lang)
    if prof is None:
        prof = bracket_lint.PROFILES["plain"]
        res["degraded"].append("未知语言 %s，按 plain 处理" % lang)

    # L1 平衡层（含全角/弯引号，bracket_lint 引擎）
    sc = None
    try:
        sc = bracket_lint.Scanner(src, prof, label, max_issues=max_issues)
        sc.run()
    except Exception as e:
        res["tool_error"] = "bracket_lint 扫描异常: %r" % (e,)
    if sc is not None:
        for i in sc.issues:
            res["issues"].append({"layer": "L1", "line": i["line"], "col": i["col"],
                                  "code": i["kind"], "msg": i["msg"]})
        r = sc.result()
        for t in r["open_stack"]:
            res["issues"].append({"layer": "L1", "line": t["line"], "col": t["col"],
                                  "code": "unclosed",
                                  "msg": "%s 未闭合（缺 %s）" % (t["char"], t["closing"])})
        if r["open_stack"]:
            res["suggestion"] = r["suggestion"]

    # L2 结构层
    ext = os.path.splitext(label)[1].lower()
    checker = L2_BY_EXT.get(ext)
    if not checker and (not ext or lang_explicit) and lang:
        checker = L2_BY_LANG.get(lang)
    if checker and not big:
        is_real = real_file and label != "-" and os.path.exists(label)
        if not is_real and getattr(checker, "needs_real_file", False):
            res["degraded"].append("自检/管道模式：跳过外部命令深检")
        else:
            try:
                issue, terr, deg = checker(label, src)
            except Exception as e:
                issue, terr, deg = None, "%s 检查异常: %r" % (getattr(checker, "__name__", "checker"), e), None
            if issue:
                res["issues"].append(issue)
            if terr and not res["tool_error"]:
                res["tool_error"] = terr
            if deg:
                res["degraded"].append(deg)

    # L3 形态层
    if lang in EXTRA_PROFILES and not big:
        try:
            for layer, ln, cl, code, msg in check_lisp_forms(src):
                res["issues"].append({"layer": layer, "line": ln, "col": cl,
                                      "code": code, "msg": msg})
        except Exception as e:
            if not res["tool_error"]:
                res["tool_error"] = "L3 形态检查异常: %r" % (e,)

    # 行内豁免（guard:skip）：最后统一过滤，只作用于 L1~L3 的机械发现；
    # 标记必须在该行的真实注释里，字符串字面量里的不算
    res["issues"], res["exempted"] = _exempt_filter(src, res["issues"], lang)

    res["issues"] = res["issues"][:50]
    res["ok"] = not res["issues"]
    return res


def read_source(path):
    """读文件并做 L0 二进制探测。返回 (src|None, skip_reason|None, tool_err|None)。"""
    if path == "-":
        data = sys.stdin.buffer.read()
    else:
        try:
            with open(path, "rb") as fh:
                data = fh.read()
        except OSError as e:
            return None, None, "读不了 %s: %s" % (path, e)
    if b"\x00" in data[:8192]:
        return None, "binary", None
    return _decode(data), None, None


def _decode(data: bytes) -> str:
    for enc in ("utf-8-sig", "gbk", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", "replace")


# --------------------------------------------------------------------------
# L4：项目专属断言（显式 --asserts <路径>，约定 run(paths) -> [(path, line, msg)...]；
# 默认不加载——仓库内容不构成执行授权，见 load_asserts）
# --------------------------------------------------------------------------

def load_asserts(path):
    """加载显式指定的断言模块（--asserts）。返回 (run 函数|None, 错误|None)。

    安全边界（收官体检 A-01）：只加载调用方**显式指定路径**的断言文件——
    仓库根目录的 guard_asserts.py 绝不自动加载。否则任何 PR 新增该文件，
    就能在开发者 `git commit` 时（pre-commit / git hook 以仓库根为 cwd 调本工具）
    在其本机执行任意代码。显式传参 = 运维者本人的授权，仓库内容不算。
    """
    if not path:
        return None, None
    if not os.path.isfile(path):
        return None, "断言文件不存在: %s" % path
    import importlib.util
    spec = importlib.util.spec_from_file_location("guard_asserts_%d" % os.getpid(), path)
    if spec is None or spec.loader is None:
        return None, "断言文件无法加载: %s" % path
    mod = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(mod)
    except Exception as e:
        return None, "断言文件加载失败: %r" % (e,)
    fn = getattr(mod, "run", None)
    if not callable(fn):
        return None, "断言文件缺 run(paths) 函数"
    return fn, None


def run_asserts(fn, paths):
    """返回 (issues, tool_error)。条目宽容解析：dict 或 (path, line, msg[, col])。"""
    try:
        raw = fn(list(paths))
    except Exception as e:
        return [], "guard_asserts.run 执行失败: %r" % (e,)
    issues = []
    for item in raw or []:
        if isinstance(item, dict):
            issues.append(_mk_issue("L4", item.get("line"), item.get("col"),
                                    item.get("code", "assert"), item.get("msg", "")))
            issues[-1]["path"] = item.get("path", "")
        elif isinstance(item, (tuple, list)) and len(item) >= 3:
            issues.append(_mk_issue("L4", item[1], 1, "assert", str(item[2])))
            issues[-1]["path"] = str(item[0])
        else:
            return issues, "guard_asserts.run 返回了不认识的条目: %r" % (item,)
    return issues, None


# --------------------------------------------------------------------------
# 渲染与输出
# --------------------------------------------------------------------------

def render(res, lines=None):
    out = []
    if res["skipped"]:
        out.append(_c("90", "[--] %s  跳过（%s）" % (res["path"], res["skipped"])))
        return "\n".join(out)
    if res["ok"]:
        out.append(_c("32", "[ok] %s  (lang=%s)" % (res["path"], res["lang"])))
    else:
        out.append(_c("31", "[!!] %s  —— %d 处结构问题 (lang=%s)"
                      % (res["path"], len(res["issues"]), res["lang"])))
    for iss in res["issues"]:
        ln, co = iss["line"], iss["col"]
        out.append("     %s 第 %d 行 第 %d 列  %s"
                   % (_c("36", iss["layer"]), ln, co,
                      _c("31", "%s: %s" % (iss["code"], iss["msg"]))))
        txt = lines[ln - 1].rstrip("\r") if lines and 1 <= ln <= len(lines) else ""
        if txt:
            out.append("   %5d | %s" % (ln, txt))
            out.append("   %5s | %s" % ("", _c("31", " " * (co - 1) + "^")))
    if res["suggestion"]:
        out.append(_c("33", "     建议：在文件末尾按顺序补上  %s" % res["suggestion"]))
    for d in res["degraded"]:
        out.append(_c("90", "     降级：%s" % d))
    if res.get("exempted"):
        out.append(_c("90", "     已按 guard:skip 豁免 %d 处（标记见对应行）"
                      % res["exempted"]))
    if res["tool_error"]:
        out.append(_c("33", "     工具故障：%s" % res["tool_error"]))
    return "\n".join(out)


def summarize(results, truncated=0):
    n_bad = sum(1 for r in results if not r["ok"] and not r["skipped"])
    n_ok = sum(1 for r in results if r["ok"] and not r["skipped"])
    n_skip = sum(1 for r in results if r["skipped"])
    n_deg = sum(len(r["degraded"]) for r in results)
    return {"files": len(results), "ok": n_ok, "bad": n_bad, "skipped": n_skip,
            "degraded": n_deg, "truncated": truncated,
            "tool_errors": [r["tool_error"] for r in results if r["tool_error"]]}


def degraded_kind(msg):
    """降级条目分类：严格模式下哪些拦、哪些可被 --allow-missing-tools 显式放行。"""
    if ("不可用" in msg or "无 PyYAML" in msg or "无 tomllib" in msg
            or "自检/管道模式" in msg):
        return "missing-tool"   # 环境缺工具：可用 --allow-missing-tools 显式放行
    if "超大文件" in msg:
        return "big-file"       # 深检被跳过：严格模式下必须拦
    if "未知语言" in msg:
        return "unknown-lang"   # 根本没有对应检查器：严格模式下必须拦
    return "other"              # 未知降级一律从严：严格模式下必须拦


def strict_blockers(results, truncated, allow_missing_tools):
    """严格模式下导致"未真正查完"的问题清单（非空 => 退出码非 0）。"""
    blockers = []
    if truncated:
        blockers.append("文件数超过上限 %d，已截断 %d 个未检查" % (MAX_FILES, truncated))
    seen = set()
    for r in results:
        for d in r["degraded"]:
            kind = degraded_kind(d)
            if kind == "missing-tool" and allow_missing_tools:
                continue
            key = (kind, d)
            if key not in seen:
                seen.add(key)
                blockers.append(d)
    return blockers


def exit_code_for(results, strict=False, allow_missing_tools=False, truncated=0):
    """1=代码问题（优先）；2=工具故障 / 严格模式下检查不完整；0=通过。

    严格模式（CI / pre-commit / 聚合门禁）：截断、超大文件降级、未知语言、
    工具缺失（未显式放行）都意味着"没有真正完成检查"，一律 2——没有真正
    完成检查，就不能显示全绿。"""
    if any(r["issues"] for r in results):
        return 1
    if strict and strict_blockers(results, truncated, allow_missing_tools):
        return 2
    if any(r["tool_error"] for r in results):
        return 2
    return 0


# --------------------------------------------------------------------------
# 文件收集
# --------------------------------------------------------------------------

def is_supported_file(path: str) -> bool:
    """是否为受支持的代码文件（含带 shebang 的无后缀脚本）。"""
    ext = os.path.splitext(path)[1].lower()
    if ext in DOC_EXTS:
        return False
    if ext in SUPPORTED_EXTS:
        return True
    if not ext and os.path.isfile(path):
        try:
            with open(path, "rb") as fh:
                return fh.read(2) == b"#!"
        except OSError:
            return False
    return False


def expand_paths(paths):
    """文件与目录 -> 待检文件表（目录递归，跳过 SKIP_DIRS 与不支持扩展名）。"""
    files = []
    for p in paths:
        if os.path.isdir(p):
            for root, dirs, names in os.walk(p):
                dirs[:] = sorted(d for d in dirs if d not in SKIP_DIRS)
                for name in sorted(names):
                    full = os.path.join(root, name)
                    if is_supported_file(full):
                        files.append(full)
        else:
            files.append(p)
    seen, uniq = set(), []
    for f in files:
        if f not in seen:
            seen.add(f)
            uniq.append(f)
    return uniq[:MAX_FILES], max(0, len(files) - MAX_FILES)


def staged_files():
    """git 暂存的代码文件（-z 按分隔符切，兼容空格/中文文件名）。"""
    git = shutil.which("git")
    if not git:
        return None, "git 不可用，--staged 无法获取暂存文件"
    rc, out, err, terr = _run_cmd(
        [git, "diff", "--cached", "--name-only", "--diff-filter=ACM", "-z"])
    if terr:
        return None, terr
    if rc != 0:
        return None, "git diff 失败(退出码 %d): %s" % (rc, (err or "").strip()[:MAX_MSG])
    return [f for f in out.split("\x00") if f.strip()], None


# --------------------------------------------------------------------------
# 闸1：agent hook 模式
# --------------------------------------------------------------------------

_HOOK_PATH_KEYS = {"file_path", "path", "notebook_path", "filePath"}


def _extract_hook_paths(obj, out, path_key=False):
    """从 hook payload 里抽文件路径。path_key=True 表示当前值挂在路径键下——
    此时列表里的字符串也是路径（`file_path: [a, b]` 与 `file_path: a` 等价）。"""
    if isinstance(obj, dict):
        for k, v in obj.items():
            _extract_hook_paths(v, out, path_key=k in _HOOK_PATH_KEYS)
    elif isinstance(obj, list):
        for v in obj:
            _extract_hook_paths(v, out, path_key=path_key)
    elif isinstance(obj, str) and obj and path_key:
        out.append(obj)


def hook_main():
    """宿主 hook 入口（"AI 改完文件立刻查一遍"）：stdin JSON -> 检查被编辑文件。

    退出码约定：0=通过，2=拦截（报告走 stderr）。任何支持
    "工具调用后跑一条命令" 的宿主都可以把本模式接上去，请在宿主侧配置
    `<技能根>/scripts/structure_guard.py --from-hook`，本技能不绑定任何平台。

    fail-open 原则：stdin 不是 JSON、路径拿不到、工具自身故障，一律 0 放行
    （闸 2/3 兜底），钩子自身绝不能变成打断工作的故障源。"""
    _init_stdout()
    try:
        if sys.stdin.isatty():
            return 0
        payload = json.load(sys.stdin)
    except Exception as e:
        print("structure_guard --from-hook: stdin 非 JSON(%s)，放行" % e, file=sys.stderr)
        return 0
    found = []
    _extract_hook_paths(payload, found)
    paths = []
    for p in found:
        if p not in paths and os.path.isfile(p) and not p.endswith(".pyc"):
            paths.append(p)
        if len(paths) >= MAX_HOOK_FILES:
            break
    if len(found) > MAX_HOOK_FILES:
        # fail-open 契约下钩子不拦，但绝不静默截断：闸 2/3 会对全量兜底
        print("structure_guard --from-hook: 输入文件 %d 个，超过上限 %d，"
              "本次只检查前 %d 个（其余由 pre-commit / CI 全量兜底）"
              % (len(found), MAX_HOOK_FILES, MAX_HOOK_FILES), file=sys.stderr)
    if not paths:
        return 0
    results = []
    for p in paths:
        src, skip, terr = read_source(p)
        if terr:
            results.append({"path": p, "lang": None, "ok": True, "skipped": None,
                            "issues": [], "degraded": [], "tool_error": terr,
                            "suggestion": "", "exempted": 0})
            continue
        if skip:
            continue
        results.append(check_source(p, src))
    bad = [r for r in results if r["issues"]]
    for r in results:
        if r["tool_error"]:
            print("structure_guard 工具故障（不拦截）：%s" % r["tool_error"],
                  file=sys.stderr)
    if bad:
        for r in bad:
            lines = None
            try:
                with open(r["path"], "rb") as fh:
                    lines = _decode(fh.read()).split("\n")
            except OSError:
                pass
            print(render(r, lines), file=sys.stderr)
        print("structure_guard: %d 个文件结构有问题，已拦截（按上面行列修复）"
              % len(bad), file=sys.stderr)
        return 2
    return 0


# --------------------------------------------------------------------------
# 自检
# --------------------------------------------------------------------------

_SELFTEST = [
    # (名称, 伪路径, 源码, 期望 ok)
    ("py 正常", "x.py", "def f(a):\n    return [1, 2]\n", True),
    ("py 括号失衡(L1)", "x.py", "def f(a):\n    return [1, 2\n", False),
    ("py 语法错误(L2)", "x.py", "def f(:\n    pass\n", False),
    ("py 缩进错乱(L2)", "x.py", "x = 1\n  y = 2\n", False),
    ("py Tab空格混用(L2)", "x.py", "def f():\n\treturn 1\n    x = 2\n", False),
    ("py 全角括号(L1)", "x.py", "x = （1 + 2）\n", False),
    ("py 字符串内全角（合法）", "x.py", 's = "（中文）"\n', True),
    ("json 正常", "x.json", '{"a": [1, 2]}', True),
    ("json 断裂", "x.json", '{"a": [1, 2}', False),
    ("toml 正常", "x.toml", "[a]\nb = 1\n", True),
    ("toml 断裂", "x.toml", "[a\nb = 1\n", False),
    ("lisp 平衡", "x.lsp", "(defun f (x)\n  (setq y (* x 2))\n  (list y))\n", True),
    ("lisp 少闭合(L1)", "x.lsp", "(defun f (x)\n  (setq y x)\n", False),
    ("lisp setq 奇数参数(L3)", "x.lsp", "(defun f (x)\n  (setq y x 2)\n  (list y))\n", False),
    ("lisp if 四段(L3)", "x.lsp", "(defun f (x)\n  (if x 1 2 3))\n", False),
    ("lisp foreach 少段(L3)", "x.lsp", "(defun f (l)\n  (foreach e l))\n", False),
    ("lisp 嵌套 defun(L3)", "x.lsp",
     "(defun outer (x)\n  (defun inner (y) y)\n  (inner x))\n", False),
    ("lisp 引用表不误报", "x.lsp", "(defun f ()\n  (setq pts '((1 2) (3 4)))\n  (length pts))\n", True),
    ("lisp 字符串含括号（合法）", "x.lsp", '(princ "(hello)")\n', True),
    ("md 文档跳过", "x.md", "# 标题（全角括号合法）\n", True),
    ("豁免：同行 guard:skip", "x.py", "x = (1  # guard:skip\n", True),
    ("豁免：字符串里的 guard:skip 不生效", "x.py", 'x = ("guard:skip"\n', False),
    ("豁免：字符串里的 guard:skip 不生效(引号后注释)", "x.py",
     's = "guard:skip"\nx = (1\n', False),
    ("豁免：指定问题码命中", "x.lsp",
     "(defun f ()\n  (setq y x 2)  ; guard:skip=setq-odd-args\n  (list y))\n", True),
    ("豁免：指定问题码不命中仍报", "x.py", "x = (1  # guard:skip=fullwidth\n", False),
    ("yaml 断裂(L1 兜底+深检)", "x.yaml", "a: [1, 2\nb: 3\n", False),
    ("超大文件只跑 L1", "big.py", "# " + "a" * (2 * 1024 * 1024) + "\nx = (1\n", False),
    # 以下四类来自下游工程实测误报的回归（2026-09-20，v3.3.1 修复）：
    ("js 嵌套模板插值不误报", "x.js",
     'const t = `${a.map((p) => `<o v="${p.id}">${esc(p.name)}（${p.k}）</o>`).join("")}`;\n', True),
    ("js 插值内失衡仍报", "x.js", "const x = `${f(}`;\n", False),
    # 以下三类来自 A-08 实测（2026-09-23）：插值内的正则/注释/换行边界。
    # 修复前那个 `'`（正则里的引号）被当成字符串开头，一路吃到 EOF 吞掉闭 `}`，
    # 于是完全合法的文件被判成"字符串没闭合 + 括号没闭合"。
    ("js 插值内正则含引号不误报", "x.js",
     'const s = `${items.map(s => s.replace(/\'/g, "")).join(",")}`;\n', True),
    ("js 插值内注释含引号不误报", "x.js", "const t = `${ /* don't */ x }`;\n", True),
    ("js 插值内除号不误判正则", "x.js", "const u = `${a / 2}`;\n", True),
    # 插值内"引号没配对"L1 刻意只停住不报（合法 JSX 文案里的撇号会被误伤），
    # 真语法错误归 L2；此处固定"L1 放行"这一层口径，防止有人回头改成误报。
    ("js 插值内串跨行（L1 不报，交 L2）", "x.js", 'const t = `${ "abc\n}`;\n', True),
    ("lisp 跨行字符串不误报", "x.lsp",
     '(defun f ()\n  (princ (strcat "\n【精雕】已将 " (itoa 1) " 个")))\n', True),
    ("yaml 内嵌 shell case 不误报", "x.yaml",
     "run: |\n  case \"$f\" in *.py) echo $f ;; esac\n", True),
    ("yaml 裸标量全角不误报", "x.yaml", "title: 我的（备用）方案\n", True),
    ("vue 标记文件跳过", "x.vue", "<div>（中文文案）{{ msg }}</div>\n", True),
    ("无后缀脚本 shebang 语法错误(L2)", "my_script", "#!/usr/bin/env python3\nx = 1 + * 2\n", False),
    ("ps1 正常", "x.ps1", "function Get-ProcessName { Get-Process | Select-Object -First 1 }\n", True),
    ("ps1 括号失衡(L1)", "x.ps1", "function Get-ProcessName { Get-Process | Select-Object -First 1\n", False),
    ("ps1 反引号转义引号不误报(A-04)", "x.ps1", '$m = "before `" after"\nWrite-Output $m\n', True),
    ("ps1 反斜杠不转义引号(A-04)", "x.ps1", '$m = "before \\" after"\n', False),
]

# 伪路径下 L1 的 plain profile 未闭合串不报警——md 用例依赖 skipped 分支，不受影响
# 超大文件用例同时验证"跳过 L2"：x = (1 是语法错误，但大文件模式不跑 ast，
# 只报 L1 的 unclosed——期望仍为不通过，但 issue 全部来自 L1。


def selftest() -> int:
    bad = 0
    for name, label, src, want_ok in _SELFTEST:
        res = check_source(label, src, real_file=False)
        got = res["ok"]
        flag = "PASS" if got == want_ok else "FAIL"
        if got != want_ok:
            bad += 1
        detail = ""
        if got != want_ok:
            detail = "  " + "; ".join(
                "%s %d:%d %s" % (i["layer"], i["line"], i["col"], i["code"])
                for i in res["issues"]) + ("  tool=%s" % res["tool_error"] if res["tool_error"] else "")
        print("  %s  %-24s 期望=%s 实得=%s%s"
              % (flag, name, "通过" if want_ok else "报错", "通过" if got else "报错", detail))
    total = len(_SELFTEST)
    print("\n  %d/%d 通过" % (total - bad, total))
    return 0 if bad == 0 else 1


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def main(argv=None) -> int:
    global USE_COLOR
    _init_stdout()
    ap = argparse.ArgumentParser(
        prog="structure_guard",
        description="AI 代码结构完整性统一门禁：L0 编码 / L1 平衡 / L2 结构 / L3 形态 / L4 断言。",
    )
    ap.add_argument("paths", nargs="*", help="要检查的文件或目录")
    ap.add_argument("--staged", action="store_true", help="检查 git 暂存的代码文件")
    ap.add_argument("--from-hook", action="store_true",
                    help="agent hook 模式：从 stdin JSON 取文件路径；有问题退出 2")
    ap.add_argument("--asserts", default=None, metavar="PATH",
                    help="显式加载断言层（L4）：run(paths) -> [(path, line, msg)...]。"
                         "默认不加载——仓库内容不构成执行授权，绝不自动执行项目根的 guard_asserts.py")
    ap.add_argument("--lang", default=None, help="强制语言 profile（默认按扩展名判断）")
    ap.add_argument("--json", action="store_true", help="输出 JSON，便于程序/agent 消费")
    ap.add_argument("--quiet", action="store_true", help="只输出有问题的文件")
    ap.add_argument("--strict", action="store_true",
                    help="严格模式：截断 / 降级 / 未知语言导致非零退出（CI、pre-commit、"
                         "聚合门禁必带——没真正查完就不算通过）")
    ap.add_argument("--allow-missing-tools", action="store_true",
                    help="配合 --strict：显式放行'工具缺席'类降级（如 runner 无 node / "
                         "PyYAML）；截断、超大文件、未知语言仍会拦")
    ap.add_argument("--max", type=int, default=5, help="L1 每文件最多报告几处（默认 5）")
    ap.add_argument("--color", choices=["auto", "always", "never"], default="auto")
    ap.add_argument("--selftest", action="store_true", help="跑内置用例，验证本工具自身可信")
    ap.add_argument("--version", action="version", version=f"structure_guard {__version__}")
    args = ap.parse_args(argv)

    if args.color == "always":
        USE_COLOR = True
        bracket_lint.USE_COLOR = True
    elif args.color == "auto":
        USE_COLOR = sys.stdout.isatty() and os.environ.get("NO_COLOR") is None
        bracket_lint.USE_COLOR = USE_COLOR

    if args.selftest:
        return selftest()

    if args.from_hook:
        return hook_main()

    if args.lang is not None and args.lang not in bracket_lint.PROFILES \
            and args.lang not in EXTRA_PROFILES:
        print("未知语言 profile: %s（可选：%s, %s）"
              % (args.lang, ", ".join(sorted(bracket_lint.PROFILES)),
                 ", ".join(sorted(EXTRA_PROFILES))), file=sys.stderr)
        return 3

    if args.paths:
        files, truncated = expand_paths(args.paths)
    elif args.staged:
        files, terr = staged_files()
        if terr:
            print("structure_guard 工具故障：%s" % terr, file=sys.stderr)
            return 2
        files = [f for f in files if is_supported_file(f)]
        truncated = max(0, len(files) - MAX_FILES)
        if truncated:
            files = files[:MAX_FILES]  # 暂存文件同样不静默截断：严格模式据此拦下
    else:
        ap.print_help()
        return 3

    if not files:
        if not args.quiet:
            print("structure_guard: 没有可检查的代码文件")
        return 0
    if truncated:
        print("structure_guard: 文件数超过上限 %d，已截断 %d 个未检查（建议缩小范围或分批）"
              % (MAX_FILES, truncated), file=sys.stderr)

    # L4 断言层（可选）：只有显式 --asserts 才加载，默认零执行
    asserts_fn, terr = load_asserts(args.asserts)
    if terr:
        print("structure_guard 工具故障：%s" % terr, file=sys.stderr)
        return 2

    results = []
    for p in files:
        src, skip, terr = read_source(p)
        if terr:
            results.append({"path": p, "lang": None, "ok": True, "skipped": None,
                            "issues": [], "degraded": [], "tool_error": terr,
                            "suggestion": "", "exempted": 0})
            continue
        if skip:
            results.append({"path": p, "lang": None, "ok": True, "skipped": skip,
                            "issues": [], "degraded": [], "tool_error": None,
                            "suggestion": "", "exempted": 0})
            continue
        res = check_source(p, src, lang=args.lang, max_issues=args.max)
        results.append(res)
        if not args.json and (not res["ok"] or not args.quiet):
            print(render(res, src.split("\n")))
            print()

    if asserts_fn is not None:
        # 一次性把全部路径交给断言层（断言可能做跨文件统计），再按路径归并回各文件
        all_paths = [r["path"] for r in results if not r["skipped"]]
        a_issues, a_err = run_asserts(asserts_fn, all_paths)
        if a_err:
            print("structure_guard 工具故障：%s" % a_err, file=sys.stderr)
            return 2
        by_key = {}
        for i in a_issues:
            by_key.setdefault(str(i.get("path") or ""), []).append(i)
        for r in results:
            extra = by_key.get(r["path"], []) + by_key.get(os.path.basename(r["path"]), [])
            if extra:
                r["issues"].extend(extra)
                r["ok"] = False
        matched = {r["path"] for r in results} | {os.path.basename(r["path"]) for r in results}
        leftover = [i for lst in by_key.values() for i in lst if i.get("path") not in matched]
        if leftover:
            for i in leftover:
                print(render({"path": i["path"], "lang": None, "ok": False,
                              "skipped": None, "issues": [i], "degraded": [],
                              "tool_error": None, "suggestion": ""}))
                print()
            results.append({"path": "<guard_asserts>", "lang": None, "ok": False,
                            "skipped": None, "issues": leftover, "degraded": [],
                            "tool_error": None, "suggestion": ""})

    if args.json:
        print(json.dumps({"summary": summarize(results, truncated), "results": results},
                         ensure_ascii=False, indent=2))
    elif not args.quiet:
        s = summarize(results, truncated)
        print("  %d 个通过，%d 个有问题，%d 个跳过；降级 %d 项%s"
              % (s["ok"], s["bad"], s["skipped"], s["degraded"],
                 ("，截断 %d 个未检查" % s["truncated"]) if s["truncated"] else ""))
        if s["degraded"] or s["truncated"]:
            print("  注意：本次检查不完整（存在降级/截断），不能视为完整检查通过"
                  + ("；--strict 模式下会以此拦下" if not args.strict else ""))

    if args.strict:
        blockers = strict_blockers(results, truncated, args.allow_missing_tools)
        if blockers:
            print("structure_guard 严格模式：检查不完整，按未通过处理（退出码 2）：",
                  file=sys.stderr)
            for b in blockers:
                print("  - %s" % b, file=sys.stderr)

    return exit_code_for(results, strict=args.strict,
                         allow_missing_tools=args.allow_missing_tools,
                         truncated=truncated)


def _init_stdout():
    """把标准输出/错误强制成 UTF-8。

    本检查器所有输出都是中文。Windows 上控制台编码不一定是能编码中文的代码页：
    英文版 runner / 英文 Windows 是 **cp1252**，`print("未闭合")` 会直接抛
    `UnicodeEncodeError: 'charmap' codec can't encode ...`，进程以 rc=1 崩掉——
    看起来像"检查器坏了"，实际只是打印不出来。

    必须在**每个入口**里调用，不能只放在 `if __name__ == "__main__"`：宿主 hook、
    单元测试都是 in-process 调用 `main()`，那条路径压根不走 `__main__` 分支
    （CI 在 Windows 上实测撞到过：3 个用例因此报错）。
    """
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


if __name__ == "__main__":
    _init_stdout()
    sys.exit(main())
