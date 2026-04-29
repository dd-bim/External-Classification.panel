# -*- coding: utf-8 -*-
"""Shared text and encoding helpers for external classification scripts."""

import re

text_type = type(u"")

_unicode_char = globals().get("unichr", chr)


_IFC_X1_RE = re.compile(r"\\X\\([0-9A-Fa-f]{2})\\")
_IFC_X2_RE = re.compile(r"\\X2\\([0-9A-Fa-f]+)\\X0\\")


def _decode_ifc_escape_block(block):
    if not block:
        return u""

    if len(block) % 4 == 0 and len(block) >= 4:
        chars = []
        for idx in range(0, len(block), 4):
            part = block[idx:idx + 4]
            try:
                chars.append(_unicode_char(int(part, 16)))
            except Exception:
                chars.append(u"?")
        return u"".join(chars)

    if len(block) % 2 == 0 and len(block) >= 2:
        chars = []
        for idx in range(0, len(block), 2):
            part = block[idx:idx + 2]
            try:
                chars.append(chr(int(part, 16)).decode("cp1252"))
            except Exception:
                try:
                    chars.append(chr(int(part, 16)).decode("latin-1"))
                except Exception:
                    chars.append(u"?")
        return u"".join(chars)

    return u""


def safe_unicode(value):
    try:
        if value is None:
            return u""
        try:
            if isinstance(value, text_type):
                return value
        except Exception:
            pass

        if hasattr(value, "ToString"):
            try:
                value = value.ToString()
            except Exception:
                pass

        if isinstance(value, (bytes, bytearray)):
            for encoding in ("utf-8", "cp1252", "latin-1"):
                try:
                    return value.decode(encoding)
                except Exception:
                    pass
            return u""

        try:
            return u"{}".format(value)
        except Exception:
            try:
                return text_type(value)
            except Exception:
                return u""
    except Exception:
        try:
            return u"{}".format(repr(value))
        except Exception:
            return u""


def safe_query_value(value):
    txt = safe_unicode(value)
    try:
        return txt.encode("utf-8")
    except Exception:
        return txt


def decode_escaped_text(value):
    txt = safe_unicode(value)
    if not txt:
        return txt

    try:
        if "\\x" in txt or "\\u" in txt or "\\U" in txt:
            try:
                decoded = txt.encode("utf-8").decode("unicode_escape")
            except Exception:
                decoded = str(txt).decode("unicode_escape")
            if decoded:
                txt = safe_unicode(decoded)
    except Exception:
        pass

    try:
        if "\\X2\\" in txt or "\\X\\" in txt:
            txt = _IFC_X2_RE.sub(lambda m: _decode_ifc_escape_block(m.group(1)), txt)
            txt = _IFC_X1_RE.sub(lambda m: _decode_ifc_escape_block(m.group(1)), txt)
    except Exception:
        pass

    return txt


def ifc_escape_text(value):
    txt = decode_escaped_text(value)
    out = []
    for ch in txt:
        try:
            code = ord(ch)
        except Exception:
            code = 0

        if ch == "'":
            out.append("''")
        elif 32 <= code <= 126 and ch != "\\":
            out.append(ch)
        else:
            out.append("\\X2\\{:04X}\\X0\\".format(code & 0xFFFF))

    return u"".join(out)


def ifc_literal_ascii(value):
    txt = ifc_escape_text(value).strip()
    if not txt:
        return "$"
    return "'{}'".format(txt)
