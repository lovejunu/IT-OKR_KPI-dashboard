import subprocess
import re
from bs4 import BeautifulSoup

SCRATCH = r"C:\Users\NKIA1\AppData\Local\Temp\claude\C--Users-NKIA1-hipo-2-IT-OKR-KPI-dashboard\88d62775-e5d1-47a2-99bc-f95356b62c7a\scratchpad"

# (logNo, date, category) category: 'product' or 'order'
POSTS = [
    ("224284295218", "2026.5.13", "product"),
    ("224153398189", "2026.1.20", "product"),
    ("224151722071", "2026.1.19", "product"),
    ("224146473146", "2026.1.14", "product"),
    ("223846036659", "2025.4.25", "product"),
    ("223428398863", "2024.4.26", "product"),
    ("222955293719", "2022.12.14", "product"),
    ("222721007568", "2022.5.4", "product"),
    ("222649362934", "2022.2.16", "product"),
    ("221922866110", "2020.4.22", "product"),
    ("222023441718", "2020.7.7", "product"),
    ("221545603464_SKIP", "", ""),  # placeholder removed below
    ("221533073579", "2019.5.9", "product"),
    ("221436157054", "2019.1.7", "product"),
    ("221371649256", "2018.10.5", "product"),
    ("224135994788", "2026.1.6", "order"),
    ("224074594881", "2025.11.13", "order"),
    ("224010129760", "2025.9.16", "order"),
    ("223327274525", "2024.1.19", "order"),
    ("222978905384", "2023.1.9", "order"),
    ("222431131631", "2021.7.14", "order"),
    ("221374557783", "2018.10.10", "order"),
    ("221373516123", "2018.10.8", "order"),  # placeholder, will fix mapping below
]

# clean, correct final list
POSTS = [
    ("224284295218", "product"),  # Polestar 10 GS인증 1등급
    ("224153398189", "product"),  # WSS 조달 혁신제품 지정
    ("224151722071", "product"),  # Polestar 10 신제품 출시
    ("224146473146", "product"),  # Polestar ITSM GS인증
    ("223846036659", "product"),  # 스마트건설챌린지 혁신상(WSS)
    ("223428398863", "product"),  # 2023 유망 SaaS 개발육성 선정
    ("222955293719", "product"),  # 글로벌 SaaS 육성 우수과제 선정
    ("222721007568", "product"),  # 조달청 우수제품 지정
    ("222649362934", "product"),  # POLESTAR Automation v3 GS인증
    ("221922866110", "product"),  # POLESTAR ITG v8 GS인증
    ("222023441718", "product"),  # 아이오션 KC인증
    ("221533073579", "product"),  # 아이오션 GS 1등급
    ("221436157054", "product"),  # AIOTION 홍보 동영상 공개
    ("221371649256", "product"),  # POLESTAR XEUS 신SW상품대상
    ("224135994788", "order"),   # 서브게이트 WSS 일본 총대리점 계약
    ("224074594881", "order"),   # 부산항만공사 공동연구 성과 공개
    ("224010129760", "order"),   # 한화건설 WSS 도입
    ("223327274525", "order"),   # BPA-엔키아 와이어로프 상시진단 사업 착수
    ("222978905384", "order"),   # 엔키아-오케스트로 클라우드 AI옵스 사업 추진
    ("222431131631", "order"),   # 나라장터 POLESTAR ITG v8 3자단가 등록
]

out_path = SCRATCH + r"\blog_post_texts.txt"
with open(out_path, "w", encoding="utf-8") as out:
    for logNo, cat in POSTS:
        html_path = SCRATCH + rf"\post_{logNo}.html"
        subprocess.run([
            "curl", "-s", "-A", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            f"https://blog.naver.com/PostView.naver?blogId=nkia_official&logNo={logNo}",
            "-o", html_path, "--max-time", "15",
        ], check=False)
        with open(html_path, encoding="utf-8", errors="ignore") as f:
            html = f.read()
        soup = BeautifulSoup(html, "html.parser")
        title_tag = soup.find("title")
        title = title_tag.text.replace(" : 네이버 블로그", "").strip() if title_tag else ""
        paras = soup.select(".se-text-paragraph, .se-main-container p, .post_ct p")
        texts = []
        for p in paras:
            t = p.get_text(" ", strip=True)
            if t:
                texts.append(t)
        body = "\n".join(texts)
        date_tag = soup.select_one(".se_publishDate, .date, .blog2_container .se_date")
        out.write(f"===== [{cat}] logNo={logNo} =====\n")
        out.write(f"TITLE: {title}\n")
        out.write(f"URL: https://blog.naver.com/nkia_official/{logNo}\n")
        out.write("BODY:\n")
        out.write(body[:1500])
        out.write("\n\n")

print("done ->", out_path)
