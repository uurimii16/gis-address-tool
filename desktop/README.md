# GIS 주소 변환기 — 데스크톱 (Windows)

엑셀/CSV 주소 목록을 **PNU · 좌표 · QGIS 레이어(GeoJSON)** 로 일괄 변환하는 데스크톱 앱.
[웹 버전](../streamlit_app.py)과 같은 병렬 엔진을 쓰지만, **로컬 실행이라 연결 끊김·타임아웃이 없어 대용량(수만 행)** 작업에 적합합니다. (VWorld 는 한국 IP에서만 정상 동작 — 웹은 해외 클라우드 IP가 차단되지만, 이 데스크톱 앱은 사용자 PC(한국 IP)에서 직접 호출하므로 안정적)

## 기능
- ① 주소 → PNU (정제주소·본번·부번, 공시지가 옵션)
- ② 주소 → 좌표 (위경도 EPSG:4326 / 중부원점TM EPSG:5186)
- ③ QGIS 레이어 (포인트·필지 경계 GeoJSON)

출력은 **원본 열을 그대로 두고 결과 열을 오른쪽에 덧붙인** `<원본이름>_결과.xlsx`. PNU 등 긴 숫자는 텍스트로 저장돼 엑셀에서 지수표기(1.15E+18)로 깨지지 않습니다.

## 처음 설정 (아주 간단)
1. `주소PNU변환기.exe` 실행.
2. 앱 위쪽 **🔑 VWorld 인증키** 칸에 본인 키를 붙여넣고 **[확인·저장]** 클릭.
   - 키가 없으면 옆의 **vworld.kr 열기** 버튼으로 무료 발급 (활용 API "2D 데이터 API" 체크, 사이트 URL `http://localhost`).
3. 끝. 키는 자동으로 `config.txt` 에 저장돼, **다음에 켜면 그대로 기억**됩니다. 파일을 직접 편집할 필요 없어요.

> 각자 자기 키를 앱에 한 번 넣으면 됩니다. `config.txt` 를 손으로 만들 필요가 없어요(고급 사용자는 편집해도 됨 — 아래 `config.txt.example` 참고).

## 배포 (다른 사람에게 줄 때)
- **`주소PNU변환기.exe` 파일 하나만** 주면 됩니다. 받은 사람은 실행 후 자기 인증키를 칸에 넣으면 끝.
- ⚠️ **내 키가 든 `config.txt` 는 함께 주지 마세요** (남이 내 사용량을 씀). exe 만 주면 상대는 자기 키를 넣습니다.
- exe 를 처음 실행하면 같은 폴더에 `config.txt` 가 자동 생성됩니다(쓰기 가능한 위치 — 바탕화면/다운로드 폴더 권장, Program Files 같은 곳은 피하기).

## 소스로 실행
```bash
pip install -r requirements.txt
python app.py
```

## exe 빌드 (PyInstaller)
```bash
pip install -r requirements.txt
python -m PyInstaller --noconfirm --onefile --windowed --icon=icon.ico ^
  --add-data "icon.ico;." --collect-all tkinterdnd2 --distpath dist --name 주소PNU변환기 app.py
```
- 빌드 결과: `dist/주소PNU변환기.exe`. 실행 시 같은 폴더에 `config.txt` 가 있어야 합니다(없으면 자동 생성 후 키 입력 안내).
- 기존 exe 가 **실행 중이면 잠겨서 빌드 실패** → 창을 닫거나 `--distpath` 를 다른 폴더로.

## 대용량 팁 (통신실패가 많다면)
- `config.txt` 의 `WORKERS` 값을 줄이세요(예: 4 → 3). VWorld 가 동시 연결이 많으면 끊습니다.
- 공시지가 옵션은 주소당 조회가 2배라 대용량에선 느리고 일일 한도에 빨리 닿습니다.
- 그래도 통신실패가 남으면, 잠시 뒤 실패분만 추려 다시 돌리면 대부분 채워집니다.
