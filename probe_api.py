"""Moldflow Synergy COM API 探针 (stdlib + 可选 pywin32)。

目的: 回答"我们现在用到的 Synergy API 是否完整/正确"。
  --static  从官方 synapi.chm (Moldflow 安装自带, Doxygen 生成) 解包并解析出
            全量 API 目录 -> temp/api_probe/catalog.json + catalog.md
  --audit   扫描本仓库 VBS 实际调用, 与目录对照, 输出 未用/未文档化 清单
            -> temp/api_probe/audit.json + audit.md
  --live    连接正在运行的 Moldflow Synergy 实例 (pywin32), 对目录中每个成员做
            IDispatch 名称解析 (GetIDsOfNames, 只查询不执行), 核对"文档有但实机
            没有"与"实机有但文档没有" -> temp/api_probe/live.json + live.md
            --invoke 额外调用零参数只读 getter 并记录返回值 (默认关闭)

用法:
  python probe_api.py                     # = --static --audit (无需 Moldflow)
  python probe_api.py --static --live
  python probe_api.py --live --invoke --start
  python probe_api.py --live --invoke --cross-probe --temp-project --start
  退出码: 0=成功, 1=环境问题(找不到 CHM/实例), 2=解析或运行失败
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DEFAULT_OUT = ROOT / "temp" / "api_probe"
CHM_NAME = "synapi.chm"
PROGID = "Synergy.Synergy"

# 仓库内 COM 对象变量 -> 文档类名 映射 (变量名取自 AutoReport.vbs / list_plots.vbs,
# 类名取自 synapi.chm 目录; 其余对象靠全局成员名匹配兜底)
VBS_OBJECT_MAP = {
    "Synergy": "Synergy",
    "StudyDoc": "StudyDoc",
    "PlotManager": "PlotMgr",
    "PlotObj": "Plot",
    "pActive": "Plot",
    "P": "Plot",
    "MP": "MaterialPlot",
    "Viewer": "Viewer",
    "LayerManager": "LayerManager",
    "LM": "LayerManager",
    "DiagnosisManager": "DiagnosisManager",
    "MeshSummary": "MeshSummary",
    "PropEd": "PropertyEditor",
    "Prop": "Property",
    "FieldValsRaw": "DoubleArray",
    "PartCADNames": "StringArray",
    "Proj": "Project",
    "ProjTmp": "Project",
}

# 实机 --invoke 的名字黑名单兜底 (写/改/耗时/GUI 弹窗), 命中一律跳过。
# "print" 必须保留: Viewer.Print 会弹打印/保存 PDF 对话框 (2026-09-16 实测)。
INVOKE_BLOCK = (
    "set",
    "create",
    "add",
    "remove",
    "delete",
    "save",
    "close",
    "quit",
    "open",
    "import",
    "export",
    "print",
    "dialog",
    "prompt",
    "pick",
    "select",
    "browse",
    "analyze",
    "mesh",
    "undo",
    "redo",
    "silence",
    "regenerate",
    "show",
    "hide",
    "fit",
    "rotate",
    "apply",
    "mark",
    "warpquery",
    "duplicate",
    "generate",
    "transform",
    "run",
    "start",
    "send",
    "play",
    "compact",
)

EXTRACT_WAIT_SEC = 90


class ProbeError(RuntimeError):
    """探针可预期的失败 (环境/解析), 汇报给用户而非堆栈。"""


def log(msg: str) -> None:
    print(msg, flush=True)


# ---------------------------------------------------------------------------
# 1. 定位 CHM
# ---------------------------------------------------------------------------


def registry_install_root() -> Path | None:
    """从 ProgID 注册信息推出 Synergy 安装目录 (不硬编码 GUID/路径)。"""
    try:
        import winreg
    except ImportError:
        return None
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, PROGID + r"\CLSID") as key:
            clsid = str(winreg.QueryValueEx(key, None)[0])
        with winreg.OpenKey(
            winreg.HKEY_CLASSES_ROOT, rf"CLSID\{clsid}\LocalServer32"
        ) as key:
            exe = str(winreg.QueryValueEx(key, None)[0]).strip('"')
    except OSError:
        return None
    exe_path = Path(exe)
    if not exe_path.is_file():
        return None
    return exe_path.parent.parent


def find_chm_candidates(explicit: str | None) -> list[Path]:
    found: list[Path] = []

    def add(path: Path | None) -> None:
        if path is not None and path.is_file() and path not in found:
            found.append(path)

    if explicit:
        add(Path(explicit).resolve())
        if not found:
            raise ProbeError(f"--chm 指定的文件不存在: {explicit}")
        return found

    root = registry_install_root()
    if root is not None:
        add(root / "help" / CHM_NAME)

    for env_name in ("ProgramFiles", "ProgramFiles(x86)"):
        base = os.environ.get(env_name)
        if not base:
            continue
        base_path = Path(base)
        if not base_path.is_dir():
            continue
        for pattern in ("Autodesk/Moldflow Synergy*/help/" + CHM_NAME,):
            for hit in sorted(base_path.glob(pattern), reverse=True):
                add(hit)
    return found


# ---------------------------------------------------------------------------
# 2. 解包 CHM
# ---------------------------------------------------------------------------


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def extract_chm(chm: Path, dest: Path, force: bool = False) -> Path:
    """把 CHM 解包到 dest。实测: 32 位 hh.exe 可用, 64 位静默失败。"""
    stamp_path = dest / ".extract_stamp.json"
    digest = sha256_file(chm)
    if not force and stamp_path.is_file():
        try:
            stamp = json.loads(stamp_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            stamp = {}
        if (
            stamp.get("chm_sha256") == digest
            and stamp.get("files", 0) > 0
            and (dest / "classes.html").is_file()
        ):
            log(f"[static] 复用已解包目录: {dest}")
            return dest

    dest.mkdir(parents=True, exist_ok=True)
    for old in dest.glob("*"):
        if old.is_file():
            old.unlink()

    # 实测: hh.exe 对含空格/受保护目录的 CHM 路径静默失败 (Program Files 下 0 输出),
    # 先复制成无空格本地路径再解包即可 (2026-09-16 本机实测)。
    local_copy = dest / "_source.chm"
    local_copy.write_bytes(chm.read_bytes())

    system_root = os.environ.get("SystemRoot", r"C:\Windows")
    hh_candidates = [
        Path(system_root) / "SysWOW64" / "hh.exe",
        Path(system_root) / "hh.exe",
    ]
    launched = []
    for hh in hh_candidates:
        if not hh.is_file():
            continue
        subprocess.run(
            [str(hh), "-decompile", str(dest), str(local_copy)],
            check=False,
            timeout=120,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        launched.append(str(hh))
        deadline = time.monotonic() + EXTRACT_WAIT_SEC
        while time.monotonic() < deadline:
            if (dest / "classes.html").is_file() and (dest / "index.hhc").is_file():
                break
            time.sleep(1)
        if (dest / "classes.html").is_file():
            break

    files = [p for p in dest.rglob("*") if p.is_file()]
    if not (dest / "classes.html").is_file() or not files:
        raise ProbeError(
            f"CHM 解包失败: {chm} -> {dest} (已尝试 {launched})。"
            " 可手动执行: %SystemRoot%\\SysWOW64\\hh.exe -decompile <目录> <chm>"
        )
    stamp_path.write_text(
        json.dumps({"chm_sha256": digest, "files": len(files)}, indent=2),
        encoding="utf-8",
    )
    log(f"[static] CHM 已解包: {len(files)} 个文件 -> {dest}")
    return dest


# ---------------------------------------------------------------------------
# 3. 解析 API 目录
# ---------------------------------------------------------------------------

_TAG_RE = re.compile(r"<[^>]+>")
_WS_RE = re.compile(r"\s+")
_CLASS_LINK_RE = re.compile(r'<a class="el" href="(class[^"#]+)\.html">([^<]+)</a>')
_MEM_ROW_RE = re.compile(r'<tr class="memitem:[^"]*">(.*?)</tr>', re.S)
_HEADING_RE = re.compile(r'<tr class="heading">(.*?)</tr>', re.S)
_ROW_LEFT_RE = re.compile(r'<td class="memItemLeft"[^>]*>(.*?)</td>', re.S)
_ROW_RIGHT_RE = re.compile(r'<td class="memItemRight"[^>]*>(.*?)</td>', re.S)
_NAME_LINK_RE = re.compile(r'<a class="el" href="[^"]*">([^<]+)</a>')
_SECTION_NAME_RE = re.compile(r'name="([^"]+)"')
_PAREN_RE = re.compile(r"\(([^()]*)\)")


def _plain(fragment: str) -> str:
    text = _TAG_RE.sub(" ", fragment)
    text = (
        text.replace("&#160;", " ")
        .replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
    )
    return _WS_RE.sub(" ", text).strip()


def parse_params(text: str) -> list[dict]:
    """把 "(long aIndex, Object aOpts)" 拆成 [{type,name}]。"""
    match = _PAREN_RE.search(text)
    if not match:
        return []
    inner = match.group(1).strip()
    if not inner:
        return []
    params = []
    for chunk in inner.split(","):
        parts = chunk.strip().split()
        if not parts:
            continue
        if len(parts) == 1:
            params.append({"type": parts[0], "name": ""})
        else:
            params.append({"type": " ".join(parts[:-1]), "name": parts[-1]})
    return params


def parse_class_html(html_text: str, fallback_name: str = "") -> dict:
    """解析 Doxygen class 页: 名称/简介/方法/属性/全量成员名。"""
    name_match = re.search(
        r'<div class="title">([^<]+) Class Reference</div>', html_text
    )
    class_name = name_match.group(1).strip() if name_match else fallback_name

    brief = ""
    brief_match = re.search(
        r'<div class="contents">\s*<p>(.*?)<a href=', html_text, re.S
    )
    if brief_match:
        brief = _plain(brief_match.group(1))

    # 成员简述在紧随其后的 <tr class="memdesc:ID"> 行的 mdescRight 单元格
    descs: dict[str, str] = {}
    for match in re.finditer(
        r'<tr class="memdesc:([^"]*)">(.*?)</tr>', html_text, re.S
    ):
        desc_match = re.search(
            r'<td class="mdescRight">(.*?)</td>', match.group(2), re.S
        )
        if desc_match:
            descs[match.group(1)] = _plain(desc_match.group(1))

    methods: list[dict] = []
    attributes: list[dict] = []
    section = ""
    tokens = []
    for match in re.finditer(
        r'<tr class="(heading|memitem):?([^"]*)">(.*?)</tr>',
        html_text,
        re.S,
    ):
        if match.group(1) == "heading":
            tokens.append(("heading", "", match.group(3)))
        else:
            tokens.append(("row", match.group(2), match.group(3)))

    for kind, row_id, payload in tokens:
        if kind == "heading":
            section = _plain(payload)
            continue
        left_match = _ROW_LEFT_RE.search(payload)
        right_match = _ROW_RIGHT_RE.search(payload)
        right_html = right_match.group(1) if right_match else ""
        link_match = _NAME_LINK_RE.search(right_html)
        if not link_match:
            continue
        member_name = link_match.group(1).strip()
        left_text = _plain(left_match.group(1)) if left_match else ""
        brief_text = descs.get(row_id, "")

        if "Member Functions" in section:
            params = parse_params(_plain(right_html[link_match.end() :]))
            methods.append(
                {
                    "name": member_name,
                    "returns": left_text,
                    "params": params,
                    "brief": brief_text,
                }
            )
        elif "Attributes" in section:
            attributes.append(
                {
                    "name": member_name,
                    "type": left_text,
                    "brief": brief_text,
                }
            )
    return {
        "name": class_name,
        "brief": brief,
        "methods": methods,
        "attributes": attributes,
    }


def parse_all_members(html_text: str) -> list[str]:
    """从 classX-members.html 取全量成员名 (含继承)。"""
    names = []
    for match in re.finditer(
        r'<td class="entry"><a class="el" href="[^"]*">([^<]+)</a>', html_text
    ):
        candidate = match.group(1).strip()
        if candidate and candidate not in names:
            names.append(candidate)
    return names


def build_catalog(extract_dir: Path, chm: Path) -> dict:
    classes_html = (extract_dir / "classes.html").read_text(
        encoding="utf-8", errors="replace"
    )
    entries = []
    seen = set()
    for match in _CLASS_LINK_RE.finditer(classes_html):
        href, visible = match.group(1), match.group(2).strip()
        if href in seen:
            continue
        seen.add(href)
        entries.append((href, visible))

    classes = []
    for href, visible in entries:
        page = extract_dir / f"{href}.html"
        if not page.is_file():
            continue
        info = parse_class_html(
            page.read_text(encoding="utf-8", errors="replace"), fallback_name=visible
        )
        members_page = extract_dir / f"{href}-members.html"
        all_members = []
        if members_page.is_file():
            all_members = parse_all_members(
                members_page.read_text(encoding="utf-8", errors="replace")
            )
        info["all_members"] = all_members
        classes.append(info)

    catalog = {
        "schema": 1,
        "source": {
            "chm": str(chm),
            "sha256": sha256_file(chm),
            "classes": len(classes),
        },
        "classes": classes,
    }
    return catalog


def write_catalog_md(catalog: dict, out_path: Path) -> None:
    lines = [
        "# Moldflow Synergy API 全量目录",
        "",
        f"- 来源: `{catalog['source']['chm']}`",
        f"- sha256: `{catalog['source']['sha256']}`",
        f"- 类数量: {catalog['source']['classes']}",
        "",
    ]
    for cls in catalog["classes"]:
        lines.append(f"## {cls['name']}")
        if cls["brief"]:
            lines.append("")
            lines.append(f"{cls['brief']}")
        if cls["attributes"]:
            lines.append("")
            lines.append("属性:")
            for attr in cls["attributes"]:
                lines.append(
                    f"- `{attr['type']} {attr['name']}` {attr['brief']}".rstrip()
                )
        if cls["methods"]:
            lines.append("")
            lines.append("方法:")
            for method in cls["methods"]:
                params = ", ".join(
                    f"{p['type']} {p['name']}".strip() for p in method["params"]
                )
                ret = method["returns"] or "void"
                lines.append(f"- `{ret} {method['name']}({params})`")
        lines.append("")
    out_path.write_text("\n".join(lines), encoding="utf-8")


# ---------------------------------------------------------------------------
# 4. 仓库用法对照
# ---------------------------------------------------------------------------

_OBJ_MEMBER_CALL_RE = re.compile(r"([A-Za-z_]\w*)\s*\.\s*([A-Za-z_]\w*)\s*\(")
_OBJ_MEMBER_READ_RE = re.compile(r"([A-Za-z_]\w*)\s*\.\s*([A-Za-z_]\w*)\b")


def strip_vbs_noise(source: str) -> str:
    """去掉 VBS 注释与字符串字面量, 避免 "Synergy.Synergy" 之类被当成成员调用。"""
    out: list[str] = []
    in_string = False
    index = 0
    length = len(source)
    while index < length:
        char = source[index]
        if in_string:
            if char == '"':
                if index + 1 < length and source[index + 1] == '"':
                    index += 2
                    continue
                in_string = False
            index += 1
            continue
        if char == '"':
            in_string = True
            out.append(" ")
            index += 1
            continue
        if char == "'":
            while index < length and source[index] not in "\r\n":
                index += 1
            continue
        out.append(char)
        index += 1
    return "".join(out)


def extract_vbs_usages(source: str) -> dict:
    """返回 {"<Class>": {"calls": [...], "reads": [...]}, "_vars": {...}}。"""
    usages: dict[str, dict] = {}
    for var, member in _OBJ_MEMBER_CALL_RE.findall(strip_vbs_noise(source)):
        cls = VBS_OBJECT_MAP.get(var)
        if cls is None:
            continue
        bucket = usages.setdefault(cls, {"calls": [], "reads": []})
        if member not in bucket["calls"]:
            bucket["calls"].append(member)
    for var, member in _OBJ_MEMBER_READ_RE.findall(strip_vbs_noise(source)):
        cls = VBS_OBJECT_MAP.get(var)
        if cls is None:
            continue
        bucket = usages.setdefault(cls, {"calls": [], "reads": []})
        if member not in bucket["reads"] and member not in bucket["calls"]:
            bucket["reads"].append(member)
    return usages


def build_audit(catalog: dict, sources: dict[str, str]) -> dict:
    by_class = {cls["name"]: cls for cls in catalog["classes"]}
    documented_members: dict[str, set[str]] = {}
    for cls in catalog["classes"]:
        names = {m["name"] for m in cls["methods"]}
        names |= {a["name"] for a in cls["attributes"]}
        names |= set(cls.get("all_members", []))
        documented_members[cls["name"]] = names

    def is_documented_anywhere(member: str) -> bool:
        return any(member in names for names in documented_members.values())

    result = {"classes": {}, "unknown_members": {}, "sources": list(sources)}
    for path, source in sources.items():
        file_usages = extract_vbs_usages(source)
        for cls, bucket in file_usages.items():
            entry = result["classes"].setdefault(
                cls, {"used": [], "unused": [], "undocumented": [], "files": []}
            )
            if path not in entry["files"]:
                entry["files"].append(path)
            for member in bucket["calls"] + bucket["reads"]:
                if member not in entry["used"]:
                    entry["used"].append(member)
                if not is_documented_anywhere(member):
                    if member not in entry["undocumented"]:
                        entry["undocumented"].append(member)
                    result["unknown_members"].setdefault(member, []).append(
                        f"{path}:{cls}"
                    )
    for cls, entry in result["classes"].items():
        doc_names = documented_members.get(cls, set())
        entry["used"].sort()
        entry["undocumented"].sort()
        entry["unused"] = sorted(doc_names - set(entry["used"]))
    return result


def write_audit_md(audit: dict, catalog: dict, out_path: Path) -> None:
    lines = [
        "# 仓库用法 vs 官方 API 目录 对照",
        "",
        f"- 扫描文件: {', '.join(audit['sources'])}",
        "",
    ]
    if audit["unknown_members"]:
        lines.append("## 用到了但官方目录查无此名 (重点疑点)")
        lines.append("")
        for member, places in sorted(audit["unknown_members"].items()):
            lines.append(f"- `{member}` <- {', '.join(places)}")
        lines.append("")
    for cls, entry in sorted(audit["classes"].items()):
        doc_cls = next((c for c in catalog["classes"] if c["name"] == cls), None)
        lines.append(f"## {cls}")
        lines.append("")
        lines.append(
            f"- 已用 {len(entry['used'])} 项: {', '.join('`' + m + '`' for m in entry['used'])}"
        )
        if entry["undocumented"]:
            lines.append(
                f"- 官方目录未收录: {', '.join('`' + m + '`' for m in entry['undocumented'])}"
            )
        unused = entry["unused"]
        if doc_cls and unused:
            lines.append(f"- 未用的官方成员 {len(unused)} 项:")
            lines.append("")
            for name in unused:
                detail = ""
                method = next(
                    (m for m in doc_cls["methods"] if m["name"] == name), None
                )
                if method:
                    params = ", ".join(
                        f"{p['type']} {p['name']}".strip() for p in method["params"]
                    )
                    detail = f"`{method['returns'] or 'void'} {name}({params})`"
                    if method["brief"]:
                        detail += f" - {method['brief']}"
                else:
                    attr = next(
                        (a for a in doc_cls["attributes"] if a["name"] == name), None
                    )
                    if attr:
                        detail = f"`{attr['type']} {name}` {attr['brief']}".rstrip()
                lines.append(f"  - {detail or name}")
        lines.append("")
    out_path.write_text("\n".join(lines), encoding="utf-8")


# ---------------------------------------------------------------------------
# 5. 实机探测 (pywin32)
# ---------------------------------------------------------------------------


def connect_synergy(start: bool = False):
    """返回 (synergy_dispatch, 连接方式说明)。"""
    try:
        import pythoncom  # noqa: F401
        import win32com.client
    except ImportError as exc:
        raise ProbeError(
            "缺少 pywin32 (实机探测需要, 静态/对照模式不需要): pip install pywin32"
        ) from exc

    sa_instance = os.environ.get("SAInstance", "")
    if sa_instance and sa_instance != "%SAInstance%":
        try:
            getter = win32com.client.GetObject(sa_instance)
            synergy = getter.GetSASynergy()
            if synergy is not None:
                return synergy, f"SAInstance -> {sa_instance}"
        except Exception as exc:  # noqa: BLE001 - 汇报真实 COM 错误
            log(f"[live] SAInstance 连接失败: {exc!r}")
    try:
        return win32com.client.GetObject(Class=PROGID), f"GetObject({PROGID})"
    except Exception as exc:  # noqa: BLE001
        log(f"[live] GetObject 未命中已运行实例: {exc!r}")
    if not start:
        raise ProbeError(
            "未找到正在运行的 Moldflow Synergy 实例。请先打开 Moldflow, 或加 --start 由探针拉起。"
        )
    log("[live] 正在创建 Synergy 实例 (会弹出 Moldflow GUI, 请稍候)...")
    synergy = win32com.client.Dispatch(PROGID)
    return synergy, f"CreateObject({PROGID})"


def _oleof(dispatch):
    return getattr(dispatch, "_oleobj_", dispatch)


def resolve_member(dispatch, name: str) -> int | None:
    """GetIDsOfNames: 只查询名称是否存在, 不执行。返回 dispid 或 None。"""
    try:
        return _oleof(dispatch).GetIDsOfNames(name)
    except Exception:  # noqa: BLE001 - 名称不存在属正常结果
        return None


def invoke_zero_arg(dispatch, name: str):
    """按 PROPERTYGET|METHOD 调用零参成员。返回 (ok, value_or_error)。"""
    import pythoncom

    try:
        dispid = _oleof(dispatch).GetIDsOfNames(name)
    except Exception as exc:  # noqa: BLE001
        return False, f"GetIDsOfNames: {exc}"
    # LCID 取 0: pywin32 晚绑定自身即用 0 (win32com/client/dynamic.py LCID = 0x0)
    try:
        value = _oleof(dispatch).Invoke(
            dispid,
            0,
            pythoncom.DISPATCH_PROPERTYGET | pythoncom.DISPATCH_METHOD,
            True,
        )
        return True, value
    except Exception as exc:  # noqa: BLE001 - 记录真实 COM 报错
        return False, f"Invoke: {exc!r}"


def call_with_args(dispatch, name: str, *args) -> tuple[bool, object]:
    """先按 METHOD 再按 PROPERTYGET 调带参成员 (仅探针内部定向使用)。"""
    import pythoncom

    errors = []
    for flag in (pythoncom.DISPATCH_METHOD, pythoncom.DISPATCH_PROPERTYGET):
        try:
            dispid = _oleof(dispatch).GetIDsOfNames(name)
            value = _oleof(dispatch).Invoke(dispid, 0, flag, True, *args)
            return True, value
        except Exception as exc:  # noqa: BLE001 - 记录真实 COM 报错
            errors.append(exc)
    return False, errors[-1]


def is_blocked(name: str) -> bool:
    lowered = name.lower()
    return any(word in lowered for word in INVOKE_BLOCK)


def is_invokable_read(cls: dict | None, member: str) -> bool:
    """--invoke 白名单语义: 只读属性, 或零参 Get/Is/Has 方法。

    2026-09-16 实测教训: 零参 void 成员里藏着 `Viewer.Print` (会弹"保存 PDF"
    对话框卡住实机), 必须按白名单语义挑选, 不能只看"零参且不在黑名单"。
    """
    if cls is None:
        return False
    for attribute in cls["attributes"]:
        if attribute["name"] == member:
            return True
    for method in cls["methods"]:
        if method["name"] == member:
            if method["params"]:
                return False
            return bool(re.match(r"^(Get|Is|Has)[A-Z_]", member))
    return False


def brief_value(value, limit: int = 120) -> str:
    if isinstance(value, str):
        return value[:limit]
    if isinstance(value, (int, float, bool)) or value is None:
        return repr(value)
    try:
        text = repr(value)
    except Exception:  # noqa: BLE001
        text = f"<{type(value).__name__}>"
    return text[:limit]


def render_value(value, by_name: dict | None = None) -> str:
    """对象值渲染: 数组类对象试着展开元素, 其余给 <类名>/<对象>。"""
    descriptor = describe_object(value)
    if not descriptor["is_object"]:
        return brief_value(value)
    class_name = descriptor.get("type_name", "")
    if not class_name and by_name:
        class_name = guess_class_of(value, by_name)
    array_text = ""
    ok, size = call_with_args(value, "Size")
    if ok and isinstance(size, int) and 0 < size <= 8:
        items = []
        for index in range(size):
            ok_item, item = call_with_args(value, "Val", index)
            if not ok_item:
                items = []
                break
            items.append(brief_value(item, limit=40))
        if items:
            array_text = " [" + "|".join(items) + "]"
    if array_text:
        return f"<{class_name or '数组'}>{array_text}"
    return f"<{class_name}>" if class_name else "<对象>"


def describe_object(value) -> dict:
    """判断返回值是不是 COM 对象, 并尽量取类型名。"""
    handle = getattr(value, "_oleobj_", None)
    if handle is None and hasattr(value, "GetIDsOfNames"):
        handle = value
    if handle is None:
        return {"is_object": False}
    info = {"is_object": True, "type_name": ""}
    try:
        type_info = handle.GetTypeInfo()
        if type_info:
            doc = type_info.GetDocumentation(-1)
            info["type_name"] = doc[0] if doc else ""
    except Exception:  # noqa: BLE001 - 不支持 ITypeInfo 属常见情况
        pass
    return info


# 对象成员 -> 文档类名 的已知归属 (来源: synapi.chm 各管理器的职责描述 + 仓库实际用法)
OBJECT_HINTS = {
    "PlotManager": "PlotMgr",
    "JobManager": "JobMgr",
    "GetFirstPlot": "Plot",
    "GetNextPlot": "Plot",
    "FindPlotByName": "Plot",
    "CreateMaterialPlot": "MaterialPlot",
    "CreateUserPlot": "UserPlot",
    "GetMeshSummary": "MeshSummary",
    "GetMeshSummary2": "MeshSummary",
    "GetPartCadNames": "StringArray",
    "CreateStringArrayAuto": "StringArray",
    "CreateDoubleArray": "DoubleArray",
    "CreateIntegerArray": "IntegerArray",
    "CreateVectorArray": "VectorArray",
    "CreateVector": "Vector",
    "FieldValues": "DoubleArray",
}


def guess_class_of(dispatch, by_name: dict, sample: int = 24) -> str:
    """无类型库时按"能解析多少文档成员"给实机对象打分, 推断它属于哪个类。"""
    best_name = ""
    best_score = 0
    for class_name, cls in by_name.items():
        names = [m["name"] for m in cls["methods"][:sample]]
        names += [a["name"] for a in cls["attributes"][:sample]]
        if not names:
            continue
        score = sum(1 for name in names if resolve_member(dispatch, name) is not None)
        if score > best_score:
            best_name, best_score = class_name, score
    if best_score >= 3:
        return best_name
    return ""


def probe_object_graph(
    synergy,
    catalog: dict,
    invoke: bool,
    max_depth: int,
    max_objects: int,
    extra_names: dict[str, list[str]] | None = None,
    cross_probe: bool = False,
) -> dict:
    by_name = {cls["name"]: cls for cls in catalog["classes"]}
    extra_names = extra_names or {}
    all_documented: list[str] = []
    for cls in catalog["classes"]:
        for member in cls["methods"]:
            if member["name"] not in all_documented:
                all_documented.append(member["name"])
        for attr in cls["attributes"]:
            if attr["name"] not in all_documented:
                all_documented.append(attr["name"])

    def members_of(class_name: str) -> list[str]:
        cls = by_name.get(class_name)
        if cls is None:
            return []
        names = [m["name"] for m in cls["methods"]]
        names += [a["name"] for a in cls["attributes"]]
        return names

    def is_safe_read_member(class_name: str, member: str) -> bool:
        return is_invokable_read(by_name.get(class_name), member)

    results: dict[str, dict] = {}
    visited: set[int] = set()

    def walk(dispatch, class_name: str, depth: int, path: str) -> None:
        if class_name is None or depth > max_depth or len(results) >= max_objects:
            return
        key = id(_oleof(dispatch))
        if key in visited:
            return
        visited.add(key)

        doc_names = members_of(class_name)
        entry = {
            "path": path,
            "class": class_name,
            "documented": len(doc_names),
            "resolved": [],
            "missing": [],
            "undocumented_used": {},
            "cross_hits": [],
            "extra_values": {},
            "children": [],
        }
        if cross_probe:
            own = set(doc_names)
            entry["cross_hits"] = sorted(
                name
                for name in all_documented
                if name not in own and resolve_member(dispatch, name) is not None
            )
        for extra in extra_names.get(class_name, []):
            if extra in doc_names:
                continue
            has = resolve_member(dispatch, extra) is not None
            entry["undocumented_used"][extra] = "存在" if has else "不存在"
        resolved_children: list[tuple[str, str, object]] = []
        for member in doc_names:
            if resolve_member(dispatch, member) is None:
                entry["missing"].append(member)
                continue
            entry["resolved"].append(member)
            if (
                invoke
                and is_safe_read_member(class_name, member)
                and not is_blocked(member)
            ):
                ok, value = invoke_zero_arg(dispatch, member)
                if not ok:
                    entry["extra_values"][member] = f"ERROR {value}"
                    continue
                descriptor = describe_object(value)
                if not descriptor["is_object"]:
                    entry["extra_values"][member] = brief_value(value)
                    continue
                child_class = descriptor.get("type_name", "")
                if child_class not in by_name:
                    hinted = OBJECT_HINTS.get(member, "")
                    if hinted in by_name:
                        child_class = hinted
                    else:
                        child_class = guess_class_of(value, by_name)
                if child_class:
                    resolved_children.append((member, child_class, value))
                    entry["extra_values"][member] = f"<{child_class}>"
                else:
                    entry["extra_values"][member] = "<object 未知类>"
        results[f"{class_name}@{path}"] = entry
        for member, child_class, value in resolved_children:
            child_path = f"{path}.{member}"
            entry["children"].append(child_path)
            walk(value, child_class, depth + 1, child_path)

    walk(synergy, "Synergy", 0, "Synergy")
    return results


def write_live_md(live: dict, out_path: Path) -> None:
    lines = [
        "# 实机 API 探针结果",
        "",
        f"- 连接方式: {live['connection']['method']}",
        f"- SAInstance 环境变量: {live['connection']['sa_instance'] or '(未设置)'}",
        f"- 是否调用只读 getter: {live['invoke']}",
        "",
    ]
    getter_probe = live["connection"].get("sa_instance_probe")
    if getter_probe:
        lines.append("## SAInstance getter 探测")
        lines.append("")
        for member, status in getter_probe.items():
            lines.append(f"- `{member}`: {status}")
        lines.append("")
    targeted = live.get("targeted_probes") or {}
    if targeted:
        lines.append("## 审计疑点实机验证")
        lines.append("")
        for member, status in targeted.items():
            lines.append(f"- `{member}`: {status}")
        lines.append("")
    for key, entry in live["objects"].items():
        lines.append(f"## {key}")
        lines.append("")
        lines.append(
            f"- 文档成员 {entry['documented']} 项, 实机可解析 {len(entry['resolved'])} 项"
        )
        if entry["missing"]:
            lines.append(
                f"- 文档有/实机无: {', '.join('`' + m + '`' for m in entry['missing'])}"
            )
        for member, status in entry.get("undocumented_used", {}).items():
            lines.append(f"- 仓库在用/官方目录无 `{member}`: 实机 {status}")
        if entry.get("cross_hits"):
            lines.append(
                f"- 实机还响应这些官方成员名 (不属本类文档, 共 {len(entry['cross_hits'])} 项):"
                f" {', '.join('`' + m + '`' for m in entry['cross_hits'])}"
            )
        if entry["extra_values"]:
            lines.append("- 只读取值:")
            lines.append("")
            for member, value in entry["extra_values"].items():
                lines.append(f"  - `{member}` = {value}")
        lines.append("")
    out_path.write_text("\n".join(lines), encoding="utf-8")


# ---------------------------------------------------------------------------
# 5.5 研究清点 (只读): 图层 / 各类型实体 / 选择内容 / CAD 扫描 / 网格包围盒
# ---------------------------------------------------------------------------

# 官方文档 (classPredicateManager.html) 明确的标签谓词前缀, 不含 CAD 面/体
STUDY_TYPE_PREFIXES = ("N", "B", "T", "TE", "C", "S", "R", "STL", "NBC", "SBC", "LCS")
# 迭代器能直接枚举的网格类型 (计数精确, 带上限防止大模型卡死)
ITERATOR_TYPES = {
    "N": ("GetFirstNode", "GetNextNode"),
    "T": ("GetFirstTri", "GetNextTri"),
    "TE": ("GetFirstTet", "GetNextTet"),
    "B": ("GetFirstBeam", "GetNextBeam"),
    "C": ("GetFirstCurve", "GetNextCurve"),
}
# 无迭代器的类型: 逐号单实体探测 (官方前缀表 + CAD 面/体)
PROBE_PREFIXES = ("S", "R", "STL", "NBC", "SBC", "LCS", "BD", "F")


def read_size(dispatch) -> int:
    ok, value = invoke_zero_arg(dispatch, "Size")
    if ok and isinstance(value, int):
        return value
    return -1


def entity_meta(study_doc, layer_mgr, entity) -> dict:
    meta = {"id": None, "layer": ""}
    ok, value = call_with_args(study_doc, "GetEntityID", entity)
    if ok and isinstance(value, int):
        meta["id"] = value
    ok_layer, layer = call_with_args(study_doc, "GetEntityLayer", entity)
    if ok_layer and describe_object(layer)["is_object"] and layer_mgr is not None:
        ok_name, name = call_with_args(layer_mgr, "GetName", layer)
        if ok_name and isinstance(name, str):
            meta["layer"] = name
    return meta


def count_by_iterator(
    study_doc, layer_mgr, first_name, next_name, limit, prefix, progress
) -> tuple[int, list]:
    """用 GetFirstX/GetNextX 迭代计数 (上限 limit), 返回 (计数, 前 3 个样本)。"""
    count = 0
    samples = []
    ok_first, item = invoke_zero_arg(study_doc, first_name)
    while ok_first and describe_object(item)["is_object"] and count < limit:
        count += 1
        if len(samples) < 3:
            samples.append(entity_meta(study_doc, layer_mgr, item))
        ok_next, item = call_with_args(study_doc, next_name, item)
        if not ok_next or not describe_object(item)["is_object"]:
            break
    progress(f"  {prefix}: {count} 个" + (" (已达上限)" if count >= limit else ""))
    return count, samples


def extend_bbox(bbox, xyz) -> list:
    if bbox is None:
        return [xyz[0], xyz[1], xyz[2], xyz[0], xyz[1], xyz[2]]
    for index, value in enumerate(xyz):
        bbox[index] = min(bbox[index], value)
        bbox[index + 3] = max(bbox[index + 3], value)
    return bbox


def run_study_inventory(
    out_dir: Path,
    start: bool,
    cad_scan_limit: int = 300,
    iter_limit: int = 3000,
    node_limit: int = 200,
) -> dict:
    """只读清点。2026-09-16 教训: 全量标签谓词选择 (N1:N99999999) 会让 Moldflow
    长时间无响应, 改为"迭代器计数 + 有界单实体探测", 且每步写进度文件。"""
    out_dir.mkdir(parents=True, exist_ok=True)
    progress_path = out_dir / "inventory_progress.log"

    def progress(msg: str) -> None:
        log(f"[inventory] {msg}")
        with open(progress_path, "a", encoding="utf-8") as handle:
            handle.write(f"{time.strftime('%H:%M:%S')} {msg}\n")

    progress_path.write_text("", encoding="utf-8")
    progress("连接 Synergy...")
    synergy, method = connect_synergy(start=start)
    progress(f"已连接: {method}")
    report: dict = {
        "connection": {
            "method": method,
            "sa_instance": os.environ.get("SAInstance", ""),
        }
    }
    ok_units, units = invoke_zero_arg(synergy, "GetUnits")
    report["units"] = units if ok_units else ""

    ok_study, study = invoke_zero_arg(synergy, "StudyDoc")
    if not ok_study or not describe_object(study)["is_object"]:
        raise ProbeError("当前没有打开的研究, 无法清点")
    study_info = {}
    for attr in (
        "StudyName",
        "MeshType",
        "AnalysisSequence",
        "AnalysisSequenceDescription",
        "NumberOfAnalyses",
        "MoldingProcess",
    ):
        ok, value = invoke_zero_arg(study, attr)
        study_info[attr] = value if ok else "<读取失败>"
    report["study"] = study_info

    ok_layer_mgr, layer_mgr = invoke_zero_arg(synergy, "LayerManager")
    if not ok_layer_mgr or not describe_object(layer_mgr)["is_object"]:
        layer_mgr = None

    layers = []
    if layer_mgr is not None:
        ok_first, layer = invoke_zero_arg(layer_mgr, "GetFirst")
        guard = 0
        while ok_first and describe_object(layer)["is_object"] and guard < 200:
            guard += 1
            ok_name, name = call_with_args(layer_mgr, "GetName", layer)
            visible = {}
            for prefix in STUDY_TYPE_PREFIXES:
                ok_v, value = call_with_args(layer_mgr, "GetTypeVisible", layer, prefix)
                visible[prefix] = bool(value) if ok_v else None
            layers.append({"name": name if ok_name else "", "visible": visible})
            ok_next, nxt = call_with_args(layer_mgr, "GetNext", layer)
            if not ok_next or not describe_object(nxt)["is_object"]:
                break
            layer = nxt
    report["layers"] = layers

    progress("枚举网格实体 (N/T/TE/B/C, 迭代器)...")
    type_report = {}
    for prefix, (first_name, next_name) in ITERATOR_TYPES.items():
        count, samples = count_by_iterator(
            study, layer_mgr, first_name, next_name, iter_limit, prefix, progress
        )
        type_report[prefix] = {
            "count": count,
            "capped": count >= iter_limit,
            "sample": samples,
        }
    report["types"] = type_report

    progress(f"探测其余前缀 (S/R/STL/NBC/SBC/LCS/BD/F, 逐号 1..{cad_scan_limit})...")
    probe_report = {}
    ok_lst, lst_probe = call_with_args(study, "CreateEntityList")
    if ok_lst:
        for prefix in PROBE_PREFIXES:
            found = []
            hits = 0
            for index in range(1, cad_scan_limit + 1):
                call_with_args(lst_probe, "SelectFromString", f"{prefix}{index}")
                if read_size(lst_probe) > 0:
                    hits += 1
                    if len(found) < 3:
                        ok_ent, entity = call_with_args(lst_probe, "Entity", 0)
                        if ok_ent:
                            found.append(entity_meta(study, layer_mgr, entity))
            probe_report[prefix] = {
                "ids_found_in_first": cad_scan_limit,
                "hits": hits,
                "sample": found,
            }
            progress(f"  {prefix}: 命中 {hits} 个")
    report["probe_types"] = probe_report

    ok_lst3, lst_probe = call_with_args(study, "CreateEntityList")
    if ok_lst3:
        call_with_args(lst_probe, "SelectFromString", "N1")
        size_after_n1 = read_size(lst_probe)
        call_with_args(lst_probe, "SelectFromString", "T1")
        size_after_t1 = read_size(lst_probe)
        call_with_args(lst_probe, "SelectFromString", "")
        report["experiments"] = {
            "SelectFromString_after_N1": size_after_n1,
            "SelectFromString_after_T1": size_after_t1,
            "SelectFromString_empty": read_size(lst_probe),
            "append_mode": size_after_t1 > size_after_n1,
        }

    ok_sel, sel = invoke_zero_arg(study, "Selection")
    if ok_sel and describe_object(sel)["is_object"]:
        size = read_size(sel)
        ok_str, text = invoke_zero_arg(sel, "ConvertToString")
        samples = []
        for index in range(min(3, max(0, size))):
            ok_ent, entity = call_with_args(sel, "Entity", index)
            if ok_ent:
                samples.append(entity_meta(study, layer_mgr, entity))
        report["selection"] = {
            "size": size,
            "string": text if ok_str else "",
            "sample": samples,
        }

    progress(f"采样节点包围盒 (前 {iter_limit} 个节点)...")
    node_count = 0
    bbox = None
    ok_node, node = invoke_zero_arg(study, "GetFirstNode")
    while ok_node and describe_object(node)["is_object"] and node_count < node_limit:
        node_count += 1
        ok_c, coord = call_with_args(study, "GetNodeCoord", node)
        if ok_c and describe_object(coord)["is_object"]:
            values = []
            for axis in ("X", "Y", "Z"):
                ok_a, value = invoke_zero_arg(coord, axis)
                values.append(
                    value if ok_a and isinstance(value, (int, float)) else None
                )
            if all(item is not None for item in values):
                bbox = extend_bbox(bbox, values)
        if node_count % 500 == 0:
            progress(f"  节点 {node_count}...")
        ok_n, node = call_with_args(study, "GetNextNode", node)
        if not ok_n or not describe_object(node)["is_object"]:
            break
    report["mesh"] = {"nodes_sampled": node_count, "bbox": bbox}
    progress(f"节点采样完成: {node_count} 个, 包围盒 {bbox}")

    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "study_inventory.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    write_study_inventory_md(report, out_dir / "study_inventory.md")
    log(f"[inventory] 清点完成 -> {out_dir / 'study_inventory.md'}")
    return report


def write_study_inventory_md(report: dict, out_path: Path) -> None:
    lines = ["# 研究只读清点", ""]
    lines.append(f"- 连接: {report['connection']['method']} | 单位: {report['units']}")
    study = report.get("study", {})
    lines.append(
        f"- 研究: {study.get('StudyName')} | 网格类型: {study.get('MeshType')}"
    )
    lines.append(
        f"- 分析序列: {study.get('AnalysisSequence')} ({study.get('AnalysisSequenceDescription')})"
    )
    lines.append("")
    lines.append("## 图层")
    lines.append("")
    for layer in report.get("layers", []):
        shown = [k for k, v in (layer.get("visible") or {}).items() if v]
        lines.append(f"- `{layer.get('name')}` 可见类型: {', '.join(shown) or '(无)'}")
    lines.append("")
    lines.append("## 网格实体计数 (迭代器, 上限见 capped)")
    lines.append("")
    for prefix, entry in (report.get("types") or {}).items():
        cap = " (已达上限, 实际更多)" if entry.get("capped") else ""
        lines.append(
            f"- {prefix}: {entry.get('count')}{cap} 样本: {entry.get('sample')}"
        )
    lines.append("")
    lines.append("## 其余类型探测 (SelectFromString 逐号)")
    lines.append("")
    for prefix, entry in (report.get("probe_types") or {}).items():
        lines.append(
            f"- {prefix}: 前 {entry.get('ids_found_in_first')} 号命中 {entry.get('hits')} 个"
            f" 样本: {entry.get('sample')}"
        )
    lines.append("")
    lines.append("## 选择状态")
    lines.append("")
    lines.append(f"- {report.get('selection')}")
    lines.append("")
    lines.append("## 实验 (SelectFromString 追加/替换语义)")
    lines.append("")
    lines.append(f"- {report.get('experiments')}")
    lines.append("")
    lines.append("## 网格包围盒 (采样 20000 节点)")
    lines.append("")
    mesh = report.get("mesh") or {}
    lines.append(
        f"- 采样节点数: {mesh.get('nodes_sampled')} | 包围盒: {mesh.get('bbox')}"
    )
    lines.append("")
    out_path.write_text("\n".join(lines), encoding="utf-8")


# ---------------------------------------------------------------------------
# 6. CLI
# ---------------------------------------------------------------------------


def run_static(chm: Path | None, out_dir: Path, force: bool) -> dict:
    candidates = find_chm_candidates(str(chm) if chm else None)
    if not candidates:
        raise ProbeError(
            "未找到官方 API 文档 synapi.chm。请用 --chm 指定路径, "
            "或确认本机已安装 Moldflow Synergy (含 help\\synapi.chm)。"
        )
    target = candidates[0]
    if len(candidates) > 1:
        log(f"[static] 发现 {len(candidates)} 份 CHM, 使用最新的: {target}")
    extract_dir = out_dir / "extract" / f"{target.stem}_{sha256_file(target)[:8]}"
    extract_chm(target, extract_dir, force=force)
    catalog = build_catalog(extract_dir, target)
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "catalog.json").write_text(
        json.dumps(catalog, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    write_catalog_md(catalog, out_dir / "catalog.md")
    total_members = sum(
        len(cls["methods"]) + len(cls["attributes"]) for cls in catalog["classes"]
    )
    log(
        f"[static] 目录已导出: {len(catalog['classes'])} 个类 / {total_members} 个成员"
        f" -> {out_dir / 'catalog.md'}"
    )
    return catalog


def run_audit(catalog: dict, out_dir: Path) -> dict:
    sources = {}
    for name in ("AutoReport.vbs", "list_plots.vbs"):
        path = ROOT / name
        if path.is_file():
            sources[name] = path.read_bytes().decode("gbk", errors="replace")
    if not sources:
        raise ProbeError("未找到 AutoReport.vbs / list_plots.vbs, 无法做用法对照")
    audit = build_audit(catalog, sources)
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "audit.json").write_text(
        json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    write_audit_md(audit, catalog, out_dir / "audit.md")
    log(
        f"[audit] 对照完成: 疑点成员 {len(audit['unknown_members'])} 个"
        f" -> {out_dir / 'audit.md'}"
    )
    return audit


def probe_sa_instance_getter(extra_names: dict[str, list[str]]) -> dict:
    """探测 %SAInstance% 指向的 getter 对象 (仓库靠 GetSASynergy 拿 Synergy)。"""
    sa_instance = os.environ.get("SAInstance", "")
    result: dict[str, str] = {}
    if not sa_instance or sa_instance == "%SAInstance%":
        return result
    try:
        import win32com.client
    except ImportError:
        return result
    try:
        getter = win32com.client.GetObject(sa_instance)
    except Exception as exc:  # noqa: BLE001 - 汇报真实 COM 错误
        result["<GetObject>"] = f"失败: {exc!r}"
        return result
    candidates = {"GetSASynergy"}
    for names in extra_names.values():
        candidates.update(names)
    for name in sorted(candidates):
        result[name] = "存在" if resolve_member(getter, name) is not None else "不存在"
    return result


def make_temp_project(synergy, scratch_dir: Path) -> dict:
    """新建临时工程+空研究, 让 StudyDoc/PlotMgr/Viewer 等对研究敏感的对象可探测。

    仅写 scratch_dir 下的临时目录, 不碰用户工程。
    幂等: 已打开的工程若就在 scratch_dir 下则直接复用 (不重复 NewProject,
    避免 Moldflow 弹"已存在使用该名称的工程"对话框); 新建时工程名/研究名
    带时间戳保证唯一。
    """
    scratch_dir.mkdir(parents=True, exist_ok=True)
    report: dict[str, str] = {}
    ok, value = call_with_args(synergy, "Silence", True)
    report["Silence(True)"] = "OK" if ok else f"失败: {value!r}"
    ok_cur, current = invoke_zero_arg(synergy, "Project")
    if ok_cur and describe_object(current)["is_object"]:
        ok_path, current_path = invoke_zero_arg(current, "Path")
        if (
            ok_path
            and isinstance(current_path, str)
            and str(scratch_dir).lower() in current_path.lower().replace("/", "\\")
        ):
            ok_name, current_name = invoke_zero_arg(current, "Name")
            report["复用已打开的临时工程"] = (
                f"{current_name if ok_name else '?'} @ {current_path}"
            )
            ok_first, first_study = invoke_zero_arg(current, "GetFirstStudyName")
            if ok_first and isinstance(first_study, str) and first_study:
                report["复用已有研究"] = first_study
            else:
                study_name = "Probe_" + time.strftime("%Y%m%d_%H%M%S")
                ok, value = call_with_args(current, "NewStudy", study_name)
                report["NewStudy"] = "OK" if ok else f"失败: {value!r}"
            return report

    # 工程名带时间戳: 同名工程会弹"已存在使用该名称的工程"对话框卡住实机
    # (2026-09-16 实测教训), 且不得依赖 Silence() 抑制该对话框。
    project_name = "ApiProbe_" + time.strftime("%Y%m%d_%H%M%S")
    ok, value = call_with_args(synergy, "NewProject", project_name, str(scratch_dir))
    report["NewProject"] = f"{project_name} OK" if ok else f"失败: {value!r}"
    if not ok:
        return report
    ok_proj, project = invoke_zero_arg(synergy, "Project")
    if not ok_proj or not describe_object(project)["is_object"]:
        report["Project"] = "为空"
        return report
    study_name = "Probe_" + time.strftime("%Y%m%d_%H%M%S")
    ok, value = call_with_args(project, "NewStudy", study_name)
    report["NewStudy"] = f"{study_name} OK" if ok else f"失败: {value!r}"
    return report


def probe_property_undocumented(synergy, extras: list[str], invoke: bool) -> dict:
    """实机验证审计疑点: Property.GetFieldDescription 是否存在/可调用。

    取材料库对象: PropertyEditor.GetFirstProperty(21000 -> 20030), 与
    AutoReport.vbs 5.0 节枚举材料库的取值顺序一致 (来源: 仓库代码)。
    """
    report: dict[str, str] = {}
    if not extras:
        return report
    ok, editor = invoke_zero_arg(synergy, "PropertyEditor")
    if not ok:
        report["PropertyEditor"] = f"获取失败: {editor!r}"
        return report
    if describe_object(editor)["is_object"] is False:
        report["PropertyEditor"] = "为空 (无材料库?)"
        return report
    prop = None
    for db_code in (21000, 20030):
        ok, candidate = call_with_args(editor, "GetFirstProperty", db_code)
        if ok and candidate is not None:
            report["GetFirstProperty"] = f"OK (db={db_code})"
            prop = candidate
            break
        report["GetFirstProperty"] = f"db={db_code} 无对象: {candidate!r}"
    if prop is None:
        return report
    for name in extras:
        has = resolve_member(prop, name) is not None
        report[f"Property.{name}"] = "存在" if has else "不存在"
        if not (has and invoke):
            continue
        ok_field, field_id = call_with_args(prop, "GetFirstField")
        if ok_field and isinstance(field_id, int) and field_id > 0:
            ok, value = call_with_args(prop, name, field_id)
            report[f"Property.{name}({field_id})"] = (
                f"OK -> {brief_value(value)}" if ok else f"调用失败: {value!r}"
            )
        else:
            report["GetFirstField"] = f"字段枚举失败: {field_id!r}"

    # 官方文档给出的替代通道: FieldDescription / FieldUnits 属性。
    # 实测其是否随 FieldValues(字段ID) 调用而切换 -> 决定能否替代 GetFieldDescription。
    if invoke:
        ok, field_id = call_with_args(prop, "GetFirstField")
        if ok and isinstance(field_id, int) and field_id > 0:
            for label in ("FieldDescription", "FieldUnits"):
                ok, value = invoke_zero_arg(prop, label)
                report[f"Property.{label}"] = (
                    f"OK -> {brief_value(value)}" if ok else f"无参读取失败: {value!r}"
                )
            for probe_round in range(3):
                ok_desc, desc = call_with_args(prop, "FieldDescription", field_id)
                ok_units, units = call_with_args(prop, "FieldUnits", field_id)
                report[f"FieldDescription(字段{field_id})"] = (
                    f"OK -> {brief_value(desc)!r}" if ok_desc else f"调用失败: {desc!r}"
                )
                report[f"FieldUnits(字段{field_id})"] = (
                    f"OK -> {render_value(units)}"
                    if ok_units
                    else f"调用失败: {units!r}"
                )
                ok, next_id = call_with_args(prop, "GetNextField", field_id)
                if not (ok and isinstance(next_id, int)) or next_id <= 0:
                    break
                field_id = next_id
    return report


def run_live(
    catalog: dict,
    out_dir: Path,
    invoke: bool,
    start: bool,
    depth: int,
    max_objects: int,
    temp_project: bool = False,
    cross_probe: bool = False,
) -> dict:
    extra_names: dict[str, list[str]] = {}
    audit_path = out_dir / "audit.json"
    if audit_path.is_file():
        try:
            audit = json.loads(audit_path.read_text(encoding="utf-8"))
            for cls, entry in audit.get("classes", {}).items():
                if entry.get("undocumented"):
                    extra_names[cls] = list(entry["undocumented"])
        except (OSError, ValueError):
            log("[live] audit.json 读取失败, 跳过疑点成员实测")
    synergy, method = connect_synergy(start=start)
    log(f"[live] 已连接: {method}")
    getter_probe = probe_sa_instance_getter(extra_names)
    temp_project_report = {}
    if temp_project:
        log("[live] 正在新建临时工程/空研究 (scratch 目录, 不碰用户数据)...")
        temp_project_report = make_temp_project(synergy, out_dir / "scratch")
        for key, value in temp_project_report.items():
            log(f"[live]   {key}: {value}")
    objects = probe_object_graph(
        synergy, catalog, invoke, depth, max_objects, extra_names, cross_probe
    )
    targeted = probe_property_undocumented(
        synergy, extra_names.get("Property", []), invoke
    )
    for key, value in targeted.items():
        log(f"[live] 疑点实测 {key}: {value}")
    live = {
        "schema": 1,
        "connection": {
            "method": method,
            "sa_instance": os.environ.get("SAInstance", ""),
            "sa_instance_probe": getter_probe,
            "temp_project": temp_project_report,
        },
        "targeted_probes": targeted,
        "invoke": invoke,
        "objects": objects,
    }
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "live.json").write_text(
        json.dumps(live, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    write_live_md(live, out_dir / "live.md")
    missing_total = sum(len(o["missing"]) for o in objects.values())
    log(
        f"[live] 探测完成: {len(objects)} 个对象, 文档有/实机无 {missing_total} 项"
        f" -> {out_dir / 'live.md'}"
    )
    return live


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Moldflow Synergy COM API 探针 (静态目录 / 用法对照 / 实机探测)"
    )
    parser.add_argument("--static", action="store_true", help="解析官方 CHM 出全量目录")
    parser.add_argument("--audit", action="store_true", help="仓库用法 vs 目录对照")
    parser.add_argument(
        "--live", action="store_true", help="连接运行中的 Moldflow 实机探测"
    )
    parser.add_argument(
        "--all", action="store_true", help="等价 --static --audit --live"
    )
    parser.add_argument(
        "--invoke", action="store_true", help="live 时额外调用零参只读 getter"
    )
    parser.add_argument(
        "--start", action="store_true", help="live 时允许由探针拉起 Moldflow"
    )
    parser.add_argument(
        "--temp-project",
        action="store_true",
        help="live 时新建临时工程+空研究, 解锁 StudyDoc/PlotMgr/Viewer 等对象",
    )
    parser.add_argument(
        "--cross-probe",
        action="store_true",
        help="live 时用全量官方成员名交叉探测每个实机对象 (发现文档外的可用成员)",
    )
    parser.add_argument(
        "--study-inventory",
        action="store_true",
        help="只读清点当前打开的研究 (图层/各类型实体/选择/CAD 扫描/包围盒)",
    )
    parser.add_argument(
        "--depth", type=int, default=3, help="live 对象图递归深度 (默认 3)"
    )
    parser.add_argument(
        "--max-objects", type=int, default=40, help="live 最多探测对象数"
    )
    parser.add_argument("--chm", help="显式指定 synapi.chm 路径")
    parser.add_argument("--out", help=f"输出目录 (默认 {DEFAULT_OUT})")
    parser.add_argument("--force-extract", action="store_true", help="强制重新解包 CHM")
    args = parser.parse_args(argv)

    out_dir = Path(args.out).resolve() if args.out else DEFAULT_OUT
    if args.study_inventory:
        run_study_inventory(out_dir, args.start)
        return 0
    do_static = args.static or args.all or not (args.audit or args.live)
    do_audit = args.audit or args.all or not (args.static or args.live)
    do_live = args.live or args.all

    catalog = None
    if do_static or do_audit or do_live:
        catalog_path = out_dir / "catalog.json"
        if not do_static and not args.force_extract and catalog_path.is_file():
            try:
                catalog = json.loads(catalog_path.read_text(encoding="utf-8"))
                log(f"[static] 复用已有目录: {catalog_path}")
            except (OSError, ValueError):
                catalog = None
        if catalog is None:
            catalog = run_static(
                Path(args.chm) if args.chm else None, out_dir, args.force_extract
            )
    if do_audit and catalog is not None:
        run_audit(catalog, out_dir)
    if do_live and catalog is not None:
        run_live(
            catalog,
            out_dir,
            args.invoke,
            args.start,
            args.depth,
            args.max_objects,
            args.temp_project,
            args.cross_probe,
        )
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except ProbeError as err:
        log(f"[ERROR] {err}")
        sys.exit(1)
    except Exception as err:  # noqa: BLE001 - 兜底汇报, 不掩盖
        log(f"[FATAL] {type(err).__name__}: {err}")
        sys.exit(2)
