import asyncio, tempfile, io
from pathlib import Path
from PIL import Image
from playwright.async_api import async_playwright

HTML_PATH = Path(__file__).parent / "index.html"
OUT_GIF   = Path(r"C:\Users\NKIA1\Desktop\dashboard_demo.gif")

PHONE_W, PHONE_H = 390, 844
FRAME_MS = 150   # 프레임당 지속 시간(ms) — 속도 조절

# ─── Polestar 제거 ────────────────────────────────────────
def hide_polestar(html: str) -> str:
    pairs = [
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
    for old, new in pairs:
        html = html.replace(old, new)
    return html

# ─── 엔키아 헤더 흐리게 CSS 주입 ─────────────────────────
BLUR_CSS = """
<style id="blur-nkia">
  .logo-mark { opacity: 0.15 !important; }
  h1 { opacity: 0.15 !important; }
</style>
"""

def inject_blur(html: str) -> str:
    return html.replace("</head>", BLUR_CSS + "</head>", 1)

# ─── 스크린샷 헬퍼 ────────────────────────────────────────
async def shot(page) -> Image.Image:
    data = await page.screenshot(
        clip={"x": 0, "y": 0, "width": PHONE_W, "height": PHONE_H}
    )
    return Image.open(io.BytesIO(data)).convert("RGB")

async def add_frames(frames, page, count, wait_ms=60):
    """현재 화면을 count번 캡처해 frames에 추가"""
    for _ in range(count):
        frames.append(await shot(page))
        if wait_ms:
            await page.wait_for_timeout(wait_ms)

async def scroll_to(page, frames, target_y, steps=12, hold=2):
    """부드럽게 스크롤하며 캡처"""
    cur = await page.evaluate("window.scrollY")
    for i in range(1, steps + 1):
        y = int(cur + (target_y - cur) * i / steps)
        await page.evaluate(f"window.scrollTo(0,{y})")
        await page.wait_for_timeout(40)
        frames.append(await shot(page))
    await add_frames(frames, page, hold)

# ─── 메인 시퀀스 ─────────────────────────────────────────
async def main():
    html = HTML_PATH.read_text(encoding="utf-8")
    html = hide_polestar(html)
    html = inject_blur(html)

    tmp = tempfile.NamedTemporaryFile(
        suffix=".html", delete=False, mode="w", encoding="utf-8"
    )
    tmp.write(html)
    tmp.close()
    tmp_path = Path(tmp.name)
    print(f"임시 HTML: {tmp_path}")

    frames: list[Image.Image] = []

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        ctx = await browser.new_context(
            viewport={"width": PHONE_W, "height": PHONE_H},
            device_scale_factor=2,
            is_mobile=True,
        )
        page = await ctx.new_page()
        await page.goto(tmp_path.as_uri())
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(600)

        # ── 1. 영업 탭 — 헤더/상단 정지 (2초) ───────────────
        print("1. 상단 헤더 표시")
        await page.evaluate("window.scrollTo(0,0)")
        await add_frames(frames, page, 14)   # ~2.1s

        # ── 2. KPI 테이블까지 스크롤 ─────────────────────────
        print("2. KPI 테이블로 스크롤")
        # KPI 섹션 위치 파악
        kpi_y = await page.evaluate(
            "document.querySelector('.card-tbl-head') ? "
            "document.querySelector('.card-tbl-head').getBoundingClientRect().top + window.scrollY - 10 : 1200"
        )
        await scroll_to(page, frames, kpi_y, steps=16, hold=3)

        # ── 3. 데모 프리셋 버튼 클릭 ──────────────────────────
        print("3. 데모 프리셋 버튼 클릭")
        demo_btn = page.locator(".demo-btn").first
        await demo_btn.scroll_into_view_if_needed()
        await add_frames(frames, page, 3)    # 버튼 보이는 정지

        await demo_btn.click()
        await page.wait_for_timeout(300)

        # ── 4. 달성률 색상 변화 — 여러 프레임 ───────────────
        print("4. 달성률 색상 변화 표시")
        await add_frames(frames, page, 18, wait_ms=80)   # ~2.7s

        # ── 5. KPI 평균 달성률 요약 보이도록 스크롤 ──────────
        ach_y = await page.evaluate(
            "document.querySelector('.kpi-ach-summary') ? "
            "document.querySelector('.kpi-ach-summary').getBoundingClientRect().top + window.scrollY - 80 : 0"
        )
        if ach_y > 0:
            await scroll_to(page, frames, ach_y, steps=10, hold=4)

        # ── 6. KPI 행 클릭 → 측정공식 모달 팝업 ────────────
        print("5. 측정공식 모달 열기")
        first_kpi_row = page.locator("table.kpi-tbl tbody tr").first
        await first_kpi_row.scroll_into_view_if_needed()
        await add_frames(frames, page, 2)
        await first_kpi_row.click()
        await page.wait_for_timeout(400)

        modal = page.locator("#kpi-modal")
        if await modal.is_visible():
            await add_frames(frames, page, 20, wait_ms=100)  # 모달 3s

            # 모달 닫기
            close_btn = page.locator("#modal-close")
            await close_btn.click()
            await page.wait_for_timeout(300)
            await add_frames(frames, page, 4)

        # ── 7. 연구소 탭 ─────────────────────────────────────
        print("6. 연구소 탭 전환")
        await page.locator('.tab:has-text("연구소")').click()
        await page.wait_for_timeout(500)
        await page.evaluate("window.scrollTo(0,0)")
        await add_frames(frames, page, 5)

        lab_h = await page.evaluate("document.body.scrollHeight")
        await scroll_to(page, frames, lab_h * 0.4, steps=14, hold=2)
        await scroll_to(page, frames, lab_h * 0.8, steps=14, hold=2)

        # ── 8. 사업수행 탭 ───────────────────────────────────
        print("7. 사업수행 탭 전환")
        await page.locator('.tab:has-text("사업수행")').click()
        await page.wait_for_timeout(500)
        await page.evaluate("window.scrollTo(0,0)")
        await add_frames(frames, page, 5)

        biz_h = await page.evaluate("document.body.scrollHeight")
        await scroll_to(page, frames, biz_h * 0.5, steps=16, hold=2)
        await scroll_to(page, frames, biz_h * 0.9, steps=14, hold=3)

        await browser.close()

    tmp_path.unlink()
    print(f"총 프레임: {len(frames)} | GIF 생성 중...")

    # RGB → 팔레트 변환
    p_frames = [f.quantize(colors=256) for f in frames]

    p_frames[0].save(
        OUT_GIF,
        save_all=True,
        append_images=p_frames[1:],
        optimize=False,
        duration=FRAME_MS,
        loop=0,
    )
    mb = OUT_GIF.stat().st_size / 1024 / 1024
    print(f"완료! {OUT_GIF}  ({len(frames)} frames · {FRAME_MS}ms · {mb:.1f} MB)")

asyncio.run(main())
