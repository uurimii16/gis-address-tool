# -*- coding: utf-8 -*-
"""Streamlit 기본 부팅 화면(static/index.html)의 'Streamlit' 제목·파비콘을
앱 제목으로 덮어써서, 페이지 열 때 잠깐 뜨는 'Streamlit' 마크 깜빡임을 없앤다.

Streamlit 서버는 시작 시 index.html을 한 번 읽어 캐시하므로, 반드시
streamlit 실행 '전'에 돌려야 한다. Cloudtype 의 Pre start Command 에
`python patch_index.py` 로 등록해 사용한다. 실패해도 앱 동작엔 영향 없음.
"""
import os
import re

try:
    import streamlit
    path = os.path.join(os.path.dirname(streamlit.__file__), "static", "index.html")
    html = open(path, encoding="utf-8").read()
    favicon = ("data:image/svg+xml,"
               "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'>"
               "<text y='.9em' font-size='90'>%F0%9F%8C%90</text></svg>")
    new = re.sub(r"<title>.*?</title>", "<title>GIS 주소 변환기</title>", html, flags=re.S)
    new = re.sub(r'(<link[^>]*rel="[^"]*icon"[^>]*href=")[^"]*(")',
                 r"\1" + favicon + r"\2", new)
    if new != html:
        open(path, "w", encoding="utf-8").write(new)
    print("[patch_index] index.html debranded:", new != html)
except Exception as e:  # 패치 실패해도 앱은 정상 실행되도록
    print("[patch_index] skipped:", e)
