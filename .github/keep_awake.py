"""무료 Streamlit 앱이 잠들지 않도록 주기적으로 방문/깨우는 스크립트."""
import os
from playwright.sync_api import sync_playwright

URL = os.environ.get("STREAMLIT_URL", "").strip()
if not URL:
    raise SystemExit("STREAMLIT_URL 환경변수를 설정하세요.")

with sync_playwright() as pw:
    browser = pw.chromium.launch()
    page = browser.new_page()
    print(f"방문: {URL}")
    page.goto(URL, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(4000)
    # 앱이 자고 있으면 "Yes, get this app back up!" 버튼이 보인다 -> 클릭
    try:
        btn = page.get_by_text("get this app back up", exact=False)
        if btn.count() > 0:
            print("앱이 잠들어 있어 깨우기 버튼 클릭")
            btn.first.click()
            page.wait_for_timeout(30000)  # 스핀업 대기
        else:
            print("이미 깨어 있음")
    except Exception as e:
        print("깨우기 버튼 없음/이미 활성:", e)
    page.wait_for_timeout(5000)
    print("완료. 제목:", page.title())
    browser.close()
