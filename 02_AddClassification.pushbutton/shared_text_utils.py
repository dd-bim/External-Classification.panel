# -*- coding: utf-8 -*-
"""Shared text and encoding helpers for external classification scripts."""

text_type = type(u"")


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
