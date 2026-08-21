# -*- coding: utf-8 -*-
"""从有道词典拉取英/美音标。"""
from __future__ import annotations

import json
import re
import time
import urllib.parse
import urllib.request
from typing import Optional


UA = "Mozilla/5.0 (compatible; twelve-vocab/1.0)"


def _clean_phone(raw: str) -> str:
    s = str(raw or "").strip()
    if not s or s in {"—", "-", "无"}:
        return ""
    # 只去掉最外层包裹的 [] / () / //，保留音标内部的 (r) 之类可选发音
    if len(s) >= 2:
        pairs = {("[", "]"), ("(", ")"), ("/", "/")}
        for a, b in pairs:
            if s.startswith(a) and s.endswith(b):
                s = s[1:-1].strip()
                break
    s = re.sub(r"\s+", " ", s)
    return s


def format_youdao_ipa(uk: str, us: str, phone: str = "") -> str:
    uk = _clean_phone(uk)
    us = _clean_phone(us)
    phone = _clean_phone(phone)
    if uk and us:
        if uk == us:
            return f"/{uk}/"
        return f"英 /{uk}/  美 /{us}/"
    if uk:
        return f"/{uk}/"
    if us:
        return f"/{us}/"
    if phone:
        return f"/{phone}/"
    return ""


def fetch_youdao_phones(word: str, timeout: float = 12.0) -> tuple[str, str, str]:
    """返回 (uk, us, phone)。失败返回空字符串。"""
    q = str(word or "").strip()
    if not q:
        return "", "", ""
    url = "https://dict.youdao.com/jsonapi?" + urllib.parse.urlencode({"q": q, "le": "en"})
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = json.loads(resp.read().decode("utf-8", "replace"))

    uk = us = phone = ""

    def take(obj: dict) -> None:
        nonlocal uk, us, phone
        if not isinstance(obj, dict):
            return
        if not uk:
            uk = str(obj.get("ukphone") or obj.get("uk-phonetic") or "")
        if not us:
            us = str(obj.get("usphone") or obj.get("us-phonetic") or "")
        if not phone:
            phone = str(obj.get("phone") or obj.get("phonetic") or "")

    simple = data.get("simple") or {}
    for item in simple.get("word") or []:
        take(item)

    ec = data.get("ec") or {}
    for item in ec.get("word") or []:
        take(item)

    return _clean_phone(uk), _clean_phone(us), _clean_phone(phone)


def fetch_youdao_ipa(word: str, timeout: float = 12.0) -> str:
    try:
        uk, us, phone = fetch_youdao_phones(word, timeout=timeout)
        return format_youdao_ipa(uk, us, phone)
    except Exception:
        return ""


def fetch_youdao_ipa_cached(
    word: str,
    cache: dict[str, str],
    sleep_s: float = 0.08,
) -> str:
    key = str(word or "").strip().lower()
    if not key:
        return ""
    if key in cache:
        return cache[key]
    ipa = fetch_youdao_ipa(word)
    cache[key] = ipa
    if sleep_s > 0:
        time.sleep(sleep_s)
    return ipa
