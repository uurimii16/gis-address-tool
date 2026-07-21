# -*- coding: utf-8 -*-
"""Streamlit 기본 부팅 화면(static/index.html)의 'Streamlit' 제목·파비콘을
앱 제목으로 덮어써서, 페이지 열 때 잠깐 뜨는 'Streamlit' 마크 깜빡임을 없앤다.

◆ 왜 빌드 단계여야 하나(실측으로 확인한 제약, 2026-07-21):
  1) Streamlit 서버는 시작 시 index.html을 한 번 읽어 캐시 → 실행 '후' 파일을
     고쳐도 반영 안 됨(반드시 실행 '전'에 패치).
  2) 실행 단계 컨테이너는 site-packages에 쓰기 권한 없음 → pre-start에서 실행 시
     [Errno 13] Permission denied. (빌드 단계는 root라 쓰기 가능)
  3) Cloudtype 빌드의 install 단계에는 requirements.txt만 있고 소스 파일
     (patch_index.py)은 아직 복사 전 → `python patch_index.py`는
     [Errno 2] No such file or directory 로 빌드 실패.

◆ 그래서 실제 운영은 이 파일이 아니라, Cloudtype 서비스 설정
  → 더 많은 옵션 → Install command 에 아래 '인라인' 명령을 넣어 해결했다
  (파일 의존 X, 빌드=root 권한, 이미지에 구워짐):

    pip install -r requirements.txt && python -c "import streamlit,os,re;p=os.path.join(os.path.dirname(streamlit.__file__),'static','index.html');h=open(p,encoding='utf-8').read();q=chr(39);g='data:image/svg+xml,<svg xmlns='+q+'http://www.w3.org/2000/svg'+q+' viewBox='+q+'0 0 64 64'+q+'><text y='+q+'52'+q+' font-size='+q+'56'+q+'>%F0%9F%8C%90</text></svg>';open(p,'w',encoding='utf-8').write(re.sub(r'<title>.*?</title>','<title>GIS 주소 변환기</title>',h,flags=re.S).replace('./favicon.png',g))"

  이 파일(patch_index.py)은 같은 로직의 참고용 스크립트다. Pre start Command 는 비워 둔다.
실패해도 앱 동작엔 영향 없음(try/except).
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
