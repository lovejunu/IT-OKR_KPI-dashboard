import asyncio, re, tempfile, os
from pathlib import Path
from PIL import Image
from playwright.async_api import async_playwright

HTML_PATH = Path(__file__).parent / "index.html"
OUT_GIF   = Path(r"C:\Users\NKIA1\Desktop\dashboard_demo.gif")

# 핸드폰 크기 (iPhone 14 기준)
PHONE_W, PHONE_H = 390, 844

def hide_polestar(html: str) -> str:
    """Polestar 단어를 적절한 대체어로 교체"""
    replacements = [
        ("Polestar EMS·ITG·WSS", "EMS·ITG·WSS"),
        ("Polestar EMS·ITG",     "EMS·ITG"),
        ("Polestar AI·AIOps",    "AI·AIOps"),
        ("Polestar AIOps",       "AIOps"),
        ("Polestar R&D",         "R&D"),
        ("Polestar 10",          "플랫폼"),
        ("Polestar 플랫폼",      "플랫폼"),
        ("EMS·Polestar 고객",    "EMS·AIOps 고객"),
        ("EMS → Polestar",       "EMS → AIOps"),
        ("Polestar 운영 고객",   "AIOps 운영 고객"),
        ("Polestar 운영 중인 고객", "AIOps 운영 중인 고객"),
        ("Polestar 운영 환경",   "운영 환경"),
        ("Polestar SLA",         "플랫폼 SLA"),
        ("Polestar는",           "플랫폼은"),
        ("Polestar의",           "플랫폼의"),
        ("Polestar 계약 전환율", "계약 전환율"),
        ("Polestar 신규 구축",   "신규 구축"),
        ("Polestar 모니터링",    "플랫폼 모니터링"),
        ("Polestar 자동화 엔진", "자동화 엔진"),
        ("Polestar 자동화 콘솔", "자동화 콘솔"),
        ("Polestar 자체 모니터링", "자체 모니터링"),
        ("Polestar 이벤트 로그", "이벤트 로그"),
        ("Polestar 릴리즈 노트", "릴리즈 노트"),
        ("Polestar 영업 담당",   "영업 담당"),
        ("Polestar 가용성",      "플랫폼 가용성"),
        ("Polestar EMS·ITG 구축", "EMS·ITG 구축"),
        ("Polestar EMS·ITG 제품", "EMS·ITG 제품"),
        ("Polestar 에이전트",    "에이전트"),
        ("Polestar ",            "플랫폼 "),
        ("Polestar",             "플랫폼"),
    ]
    for old, new in replacements:
        html = html.replace(old, new)
    return html

async def capture_frames(page, scroll_height, num_frames=80):
    frames = []
    # 탭 클릭 (영업 → 연구소 → 사업수행 순서로 스크롤)
    tabs = ["영업", "연구소", "사업수행"]
    frames_per_tab = num_frames // len(tabs)

    for tab_name in tabs:
        # 탭 클릭
        await page.locator(f'.tab:has-text("{tab_name}")').click()
        await page.wait_for_timeout(400)

        # 현재 탭의 페이지 높이
        tab_height = await page.evaluate("document.body.scrollHeight")
        scroll_range = max(tab_height - PHONE_H, 0)

        steps = frames_per_tab
        for i in range(steps):
            scroll_y = int(scroll_range * i / max(steps - 1, 1))
            await page.evaluate(f"window.scrollTo(0, {scroll_y})")
            await page.wait_for_timeout(50)
            shot = await page.screenshot(
                clip={"x": 0, "y": 0, "width": PHONE_W, "height": PHONE_H}
            )
            frames.append(Image.open(__import__("io").BytesIO(shot)).convert("RGBA"))

    return frames

async def main():
    html_text = HTML_PATH.read_text(encoding="utf-8")
    clean_html = hide_polestar(html_text)

    # 임시 HTML 파일 생성
    tmp = tempfile.NamedTemporaryFile(
        suffix=".html", delete=False, mode="w", encoding="utf-8"
    )
    tmp.write(clean_html)
    tmp.close()
    tmp_path = Path(tmp.name)

    print(f"임시 HTML: {tmp_path}")
    print("브라우저 실행 중...")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        ctx = await browser.new_context(
            viewport={"width": PHONE_W, "height": PHONE_H},
            device_scale_factor=2,
            is_mobile=True,
            user_agent=(
                "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) "
                "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1"
            ),
        )
        page = await ctx.new_page()
        await page.goto(tmp_path.as_uri())
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(800)

        scroll_height = await page.evaluate("document.body.scrollHeight")
        print(f"페이지 높이: {scroll_height}px | 캡처 시작...")

        frames = await capture_frames(page, scroll_height, num_frames=90)
        await browser.close()

    tmp_path.unlink()

    print(f"프레임 수: {len(frames)} | GIF 저장 중...")

    # RGBA → RGB → P 팔레트 변환 (GIF용)
    rgb_frames = [f.convert("RGB") for f in frames]
    p_frames = [f.quantize(colors=256) for f in rgb_frames]

    p_frames[0].save(
        OUT_GIF,
        save_all=True,
        append_images=p_frames[1:],
        optimize=False,
        duration=80,   # ms per frame
        loop=0,
    )
    size_mb = OUT_GIF.stat().st_size / 1024 / 1024
    print(f"완료! {OUT_GIF} ({size_mb:.1f} MB)")

asyncio.run(main())
