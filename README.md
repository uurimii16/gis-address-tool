# 🌐 GIS 주소 변환기 (GIS Address Tool)

주소 목록을 **PNU · 좌표 · QGIS 지도 레이어**로 한 번에 변환해 주는 웹 도구입니다.
도시계획·GIS 실무에서 반복되는 "주소 → 필지/좌표 변환" 작업을 엑셀·CSV 업로드만으로 자동 처리합니다.

**▶ 바로 사용하기:** https://gis-address-tool-3gahdq8hlmkpfpyfg2dz4q.streamlit.app/

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://gis-address-tool-3gahdq8hlmkpfpyfg2dz4q.streamlit.app/)
![Python](https://img.shields.io/badge/Python-3.11-blue)
![Streamlit](https://img.shields.io/badge/Streamlit-app-red)
![VWorld](https://img.shields.io/badge/API-VWorld-brightgreen)

---

## ✨ 주요 기능

| 기능 | 설명 | 결과물 |
|---|---|---|
| **① 주소 → PNU** | 주소를 19자리 필지고유번호(PNU)로 변환. 정제주소·본번·부번 포함, 공시지가(원/㎡)·기준연월 선택 조회 | 엑셀(.xlsx) |
| **② 주소 → 좌표** | 주소를 지도 좌표로 변환 (위경도 EPSG:4326 / 중부원점TM EPSG:5186) | 엑셀(.xlsx) + 지도 미리보기 |
| **③ QGIS 레이어** | 주소를 QGIS에서 바로 열리는 지도 레이어(GeoJSON)로 생성 — 포인트 / 실제 필지 경계 | .geojson |

### 편의 기능
- 📄 **엑셀(.xlsx)·CSV 업로드** — 주소 열 **자동 인식** + 미리보기로 확인·수정
- 🧩 **다양한 주소 형태 지원** — "한 칸에 전체주소" / "시도·시군구·읍면동·리 분리" / "본번·부번 분리"
- 🔤 **CSV 인코딩 자동 감지** (UTF-8 / CP949 / EUC-KR) — 공공데이터·SGIS CSV 대응
- 📊 변환 결과 요약(성공/실패/건너뜀) + 세션별 변환 기록

---

## 🖥 사용 방법

1. **VWorld 인증키 발급** (무료) — [vworld.kr](https://www.vworld.kr) → 오픈API → 인증키 발급
   - 활용 API에 **2D 데이터 API** 체크, 사이트 URL은 `http://localhost`
2. 앱에서 인증키 입력 → **기능 선택**(①/②/③)
3. 주소가 담긴 **엑셀·CSV 업로드** → 자동 인식된 주소 열 확인
4. **변환 시작** → 결과 확인 후 내려받기 (③은 .geojson을 QGIS 창에 끌어다 놓으면 바로 표시)

> 🔒 입력한 인증키와 파일은 변환에만 사용되며 서버에 저장되지 않습니다. (키는 각자 발급·입력)

---

## 🛠 기술 스택

- **Python** · **Streamlit** — 웹 UI
- **VWorld 오픈API** — 지오코딩(getcoord) / 필지 데이터(GetFeature)
- **openpyxl · pandas** — 엑셀/표 처리
- **GeoJSON** — QGIS 연동 지도 레이어

---

## ⏰ 무중단 유지 (GitHub Actions 자동화)

무료 Streamlit Community Cloud는 트래픽이 없으면 앱이 절전 상태로 잠듭니다.
이를 막기 위해 **GitHub Actions가 6시간마다 앱을 자동 방문**하고, 잠들어 있으면 깨우기까지 수행합니다.

- `.github/workflows/keep-awake.yml` — cron `0 */6 * * *` (6시간마다) 스케줄 실행
- `.github/keep_awake.py` — Playwright로 앱 접속·깨우기 자동화
- 앱 배포에는 영향 0 (playwright는 워크플로 안에서만 설치)

---

## 📂 저장소 구성

```
gis-address-tool/
├─ streamlit_app.py            # 메인 앱
├─ requirements.txt            # 앱 의존성 (streamlit, requests, openpyxl, pandas)
├─ config.toml                 # 설정
└─ .github/
   ├─ keep_awake.py            # 절전 방지 스크립트 (Playwright)
   └─ workflows/keep-awake.yml # 6시간마다 자동 실행
```

---

## 📝 라이선스 / 참고

- 개인 실무·포트폴리오 용도로 제작한 도구입니다.
- 지오코딩·필지 데이터는 **국토교통부 VWorld** 오픈API를 사용합니다. (사용자가 각자 무료 인증키 발급)
