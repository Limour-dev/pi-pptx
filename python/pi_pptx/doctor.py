"""Environment self-check: versions, soffice, font substitution, theme fonts."""
from __future__ import annotations

import shutil
import subprocess
import sys

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml

from .utils import CACHE_ROOT, KitError, emit, load

def theme_fonts(prs) -> dict:
    """Read major/minor latin + East-Asian typefaces from the first theme."""
    try:
        theme_part = prs.slide_masters[0].part.part_related_by(RT.THEME)
        root = parse_xml(theme_part.blob)
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        scheme = root.find(".//a:fontScheme", ns)
        if scheme is None:
            return {}
        out = {}
        for kind in ("major", "minor"):
            node = scheme.find(f"a:{kind}Font", ns)
            if node is None:
                continue
            entry = {}
            for tag in ("latin", "ea", "cs"):
                el = node.find(f"a:{tag}", ns)
                if el is not None and el.get("typeface"):
                    entry[tag] = el.get("typeface")
            scripts = {}
            for el in node.findall("a:font", ns):
                script = el.get("script")
                if script in ("Hans", "Hant", "Jpan", "Hang"):
                    scripts[script] = el.get("typeface")
            if scripts:
                entry["scripts"] = scripts
            out[kind] = entry
        return out
    except Exception:
        return {}


CJK_PROBES = ["微软雅黑", "思源黑体 CN", "宋体", "Noto Sans CJK SC"]
_FONT_CACHE: dict | None = None


def _font_substitutions(ttl: float = 3600.0) -> dict:
    """Requested CJK face -> what fontconfig/LibreOffice will actually use.

    Cached on disk: every render process would otherwise pay four fc-match
    calls, and the answer only changes when fonts are installed.
    """
    global _FONT_CACHE
    if _FONT_CACHE is not None:
        return _FONT_CACHE
    cache_file = CACHE_ROOT / "fonts.json"
    try:
        import json
        import time

        if cache_file.exists() and time.time() - cache_file.stat().st_mtime < ttl:
            _FONT_CACHE = json.loads(cache_file.read_text(encoding="utf-8"))
            return _FONT_CACHE
    except Exception:
        pass
    fonts = {}
    for font in CJK_PROBES:
        try:
            result = subprocess.run(["fc-match", font], capture_output=True, text=True, timeout=10)
            fonts[font] = result.stdout.strip() or "unknown"
        except Exception:
            fonts[font] = "fc-match unavailable"
    _FONT_CACHE = fonts
    try:
        import json

        CACHE_ROOT.mkdir(parents=True, exist_ok=True)
        cache_file.write_text(json.dumps(fonts, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass
    return fonts


def doctor_payload(path: str | None = None) -> dict:
    """Collect versions, soffice, font substitution and theme fonts."""
    info = {
        "python": sys.version.split()[0],
        "cache_dir": str(CACHE_ROOT),
    }
    try:
        import pptx

        info["python_pptx"] = pptx.__version__
    except Exception as exc:
        info["python_pptx"] = f"ERROR: {exc}"
    try:
        import pymupdf

        info["pymupdf"] = pymupdf.__version__
    except Exception as exc:
        info["pymupdf"] = f"ERROR: {exc}"
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    info["soffice"] = soffice
    if soffice:
        try:
            result = subprocess.run([soffice, "--version"], capture_output=True, text=True, timeout=30)
            info["soffice_version"] = result.stdout.strip().splitlines()[0] if result.stdout.strip() else ""
        except Exception as exc:
            info["soffice_version"] = f"ERROR: {exc}"
    fonts = _font_substitutions()
    info["fonts"] = fonts

    if path:
        try:
            prs, _path = load(path)
            info["deck_theme_fonts"] = theme_fonts(prs)
        except KitError as exc:
            info["deck_theme_fonts"] = f"ERROR: {exc}"
    info["notes"] = [
        "ppt_add_text(font=) writes a:latin + a:ea + a:cs, so CJK uses the same face.",
        "LibreOffice substitutes missing fonts: rendered widths can differ ~5%; leave 5-10% margin.",
        "Run this first: knowing the substitution up front avoids re-fitting every line later.",
    ]
    return info


def cmd_doctor(args):
    emit(doctor_payload(getattr(args, "pptx", None)))
