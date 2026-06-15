"""KPI 실적 입력 + 달성률 자동계산 기능 추가 패치"""
import re, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

PATH = r'C:\Users\NKIA1\hipo-2\IT OKR_KPI-dashboard\index.html'
with open(PATH, encoding='utf-8') as f:
    html = f.read()

# ─── 1. CSS 추가 (</style> 바로 앞) ──────────────────────────────────────
NEW_CSS = """
/* ── KPI Actual Input ── */
.kpi-act-input{background:var(--bg);border:1px solid var(--line);border-radius:5px;
  color:var(--txt);font-size:12px;font-weight:700;width:68px;padding:3px 5px;
  text-align:center;transition:border-color .15s}
.kpi-act-input:focus{outline:none;border-color:var(--accent)}
.kpi-act-input::placeholder{color:var(--muted);font-weight:400;font-size:10px}
.unit-lbl{font-size:10px;color:var(--muted);margin-left:2px}
.ach-chip{display:inline-block;min-width:48px;padding:2px 7px;border-radius:4px;
  font-size:11px;font-weight:800;text-align:center;letter-spacing:-.3px}
.ach-green{background:rgba(34,197,94,.18);color:#4ade80;border:1px solid rgba(34,197,94,.35)}
.ach-yellow{background:rgba(234,179,8,.15);color:#fbbf24;border:1px solid rgba(234,179,8,.3)}
.ach-red{background:rgba(239,68,68,.16);color:#f87171;border:1px solid rgba(239,68,68,.3)}
.ach-empty{background:rgba(255,255,255,.04);color:var(--muted);border:1px solid rgba(255,255,255,.1);font-weight:400;font-size:10px}
.demo-btn{background:linear-gradient(135deg,#7c3aed,#2563eb);border:none;color:#fff;
  padding:5px 14px;border-radius:14px;font-size:11px;font-weight:700;cursor:pointer;
  transition:opacity .15s;white-space:nowrap;flex-shrink:0}
.demo-btn:hover{opacity:.82}
.reset-btn{background:rgba(255,255,255,.06);border:1px solid var(--line);color:var(--muted);
  padding:5px 10px;border-radius:14px;font-size:11px;cursor:pointer;transition:.15s;flex-shrink:0}
.reset-btn:hover{border-color:var(--accent);color:var(--txt)}
.kpi-tbl-actions{display:flex;align-items:center;gap:6px;padding:8px 14px 0;flex-wrap:wrap}
.kpi-ach-summary{display:flex;align-items:center;gap:10px;padding:10px 14px;
  border-top:1px solid var(--line);background:rgba(255,255,255,.018);flex-wrap:wrap}
.kpi-ach-lbl{font-size:10px;color:var(--muted);font-weight:700;white-space:nowrap}
.kpi-ach-val{font-size:18px;font-weight:800;color:var(--accent);min-width:48px}
.kpi-ach-bar{flex:1;min-width:80px;height:5px;background:var(--line);border-radius:3px;overflow:hidden}
.kpi-ach-bar-fill{height:100%;background:linear-gradient(90deg,var(--accent),var(--teal));
  border-radius:3px;transition:width .5s ease;width:0%}
.kpi-ach-sub{font-size:10px;color:var(--muted)}
th.col-act,td.col-act{min-width:80px}
th.col-ach,td.col-ach{min-width:62px}
"""
html = html.replace('</style>', NEW_CSS + '</style>', 1)

# ─── 2. DATA KPI에 tgtNum/tgtDir/unit/demoVal 추가 ───────────────────────
# 각 KPI의 howto 뒤에 새 속성 추가. howto는 unique하므로 안전.

kpi_additions = {
    # 영업
    'CRM 파이프라인 리포트 → "Qualified" 스테이지 이상 필터 → 주간 누적 집계. 영업 주간 리뷰에서 관리자와 확인':
        (30, 'hi', '건', 23),
    'CRM Opportunity 스테이지 히스토리 리포트 + 나라장터 제출 영수증 병행 집계':
        (15, 'hi', '건', 11),
    'CRM "POC In Progress" 스테이지 + 착수 합의서 파일 첨부 여부로 월간 집계':
        (5, 'hi', '건', 4),
    'CRM 파이프라인 리포트(Qualified 이상 스테이지) 총액 ÷ 당분기 목표액 | 월초 업데이트':
        (3, 'hi', '×', 2.4),
    '나라장터 입찰 시스템 참여 이력 월간 집계 | CRM에 입찰 번호 기록 병행':
        (25, 'hi', '건', 18),
    '분기 마지막 달 5문항 설문 이메일 발송 | 응답률 60%+ 시 유효 집계':
        (4.0, 'hi', '점', 4.1),
    'CRM Account Type: Existing Customer + Deal Type: Expansion 필터 | 분기 집계':
        (5, 'hi', '건', 4),
    '분기 마감 후 CRM Closed Won ÷ 전체 Proposal Submitted | Closed Lost 원인 분류 병행':
        (25, 'hi', '%', 21),
    'CRM POC Completed → Closed Won 전환 추적 | POC 완료 날짜 기준 90일 윈도우':
        (60, 'hi', '%', 62),
    'CRM Opportunity 생성일·Close Date 차이 자동 계산 | 분기별 평균 추이 트래킹':
        (120, 'lo', '일', 105),
    '재무시스템·CRM 연계 계약서 기준 매출 집계 | 분기 중간 50%·75% 달성 확인':
        (100, 'hi', '%', 87),
    '계약 관리 시스템에서 만료일 기준 3개월 전 파이프라인 생성 → 갱신 완료 집계':
        (90, 'hi', '%', 92),
    'CRM Deal Type: Expansion/Upsell 매출 합산 ÷ 전체 신규 계약 매출 | 분기 마감 집계':
        (30, 'hi', '%', 28),
    '나라장터 낙찰 기록 + 직접 계약 공공기관 계약서 기준 | 분기 2건 페이스 유지 확인':
        (8, 'hi', '건', 5),
    # 연구소
    'Jira 스프린트 리포트 → Completed vs Committed Story Points | 2주 스프린트 종료 시 자동 생성':
        (85, 'hi', '%', 82),
    'CI/CD 파이프라인 배포 로그 (Jenkins·GitLab CI) 주간 집계 | 핫픽스·긴급 패치 포함':
        (1, 'hi', '회/주', 1.2),
    'GitHub/GitLab API로 PR 이벤트 타임스탬프 추출 → 평균 계산 | 주간 리포트 자동화 권장':
        (24, 'lo', 'h', 18),
    'Polestar 자동화 콘솔에서 신규 등록 룰 태그 기준 집계 | 분기 릴리즈 노트에 명시':
        (3, 'hi', '건', 3),
    '사내 Wiki·GitHub Docs·블로그 게시 후 PM 승인 기준 | Confluence에서 작성일 기준 집계':
        (2, 'hi', '건', 2),
    'Polestar 자체 모니터링 또는 외부 체크 도구 | 월간 SLA 리포트 발행':
        (99.9, 'hi', '%', 99.95),
    'Jira Bug 이슈 + Priority: P1·P2 필터 + 배포 이후 레이블 | 배포 릴리즈별 추적 권장':
        (3, 'lo', '건', 2),
    'Jira 인시던트 이슈 생성 시각 ~ 해결 처리 시각 차이 | 분기 평균 전분기 대비 %':
        (10, 'hi', '%', 13),
    'CI/CD 배포 태그(Rollback·Hotfix) 기준 집계 ÷ 전체 배포 건수 | 배포 로그 자동 집계':
        (5, 'lo', '%', 3.2),
    '레이블된 운영 이벤트 데이터셋 기준 분기 1회 평가 | 알람 발생 시 T/P·F/P 태깅 필요':
        (85, 'hi', '%', 83),
    '레이블 데이터셋에서 실제 이상 이벤트 중 탐지된 비율 | Precision과 함께 F1 Score 보조 추적':
        (80, 'hi', '%', 78),
    'Ask Lucida 응답 하단 피드백 클릭율 추적 | 분기 응답 샘플 품질 리뷰 병행':
        (80, 'hi', '%', 79),
    'Polestar 이벤트 로그 기준 알람 건수 월별 추이 | 기준선(Q1 평균)과 당분기 비교':
        (30, 'hi', '%', 34),
    '분기 릴리즈 계획서(Epic/Story) 기준 완료 태그 집계 | 범위 변경 건은 제외':
        (90, 'hi', '%', 88),
    'Polestar 릴리즈 노트에서 "Cloud Integration" 카테고리 건수 | 반기 목표 달성 현황':
        (2, 'hi', '건', 1),
    # 사업수행
    '고객사별 월간 헬스체크 캘린더 + 완료 체크리스트 | 티켓 시스템에 헬스체크 태스크 자동 생성':
        (100, 'hi', '%', 95),
    '티켓 시스템 자동 타임스탬프 (Jira Service Desk·Freshdesk 등) | P1 실시간 모니터링 필요':
        (30, 'lo', '분', 22),
    '구축 완료 후 3개월에 영업 담당자가 전화/이메일로 의향 확인 | CRM에 결과 기록':
        (95, 'hi', '%', 93),
    '업무 시간 기록(팀원 주간 업무 로그)에서 자동화 vs 수동 비율 집계 | 분기 초 기준선 설정':
        (20, 'hi', '%', 18),
    'Polestar 모니터링 가용성 리포트 → 고객사별 월간 SLA 달성 여부 자동 산출 | 월간 발행':
        (99.5, 'hi', '%', 99.7),
    '티켓 P1 태그 + 발생~Resolved 타임스탬프 자동 집계 | 분기 평균 + 최악값 추적':
        (2, 'lo', 'h', 1.5),
    '티켓 P2 태그 기준 자동 집계 | P1·P2 Priority Matrix 팀 내 공유 필요':
        (8, 'lo', 'h', 6.5),
    '티켓 원인 분류(RCA 태그) 필수화 → 동일 태그 2회 이상 건수 / 전체 건수 | 분기 집계':
        (20, 'lo', '%', 15),
    '프로젝트 관리 도구에서 예정 Go-Live일 vs 실제 Go-Live일 비교 | 프로젝트 종료 시 기록':
        (90, 'hi', '%', 88),
    'Go-Live 날짜 태그된 프로젝트 티켓에서 30일 이내 P1 이슈 건수 집계 | 프로젝트별 리뷰':
        (0, 'zero', '건', 0),
    'Go-Live 후 1주일 내 의사결정자·실무자 각 1명 이상 5항목 설문 발송 | 영업 담당자 연계':
        (4.2, 'hi', '점', 4.3),
    '티켓 시스템 분기 리포트 → Status: Resolved·Closed 건수 ÷ 전체 수신 건수 | 주간 모니터링':
        (98, 'hi', '%', 97),
    '티켓 해결 처리 시 자동 이메일 설문 발송(Freshdesk·Jira) | 응답률 50%+ 시 유효':
        (4.2, 'hi', '점', 4.2),
    '"Escalated to Dev" 태그 필수화 → 분기 집계 | 원인 분류(버그·기능부재·지식부족) 구분':
        (10, 'lo', '%', 8),
    '주간 업무 일지에서 Toil 카테고리 시간 기록 필수 | 분기 평균 집계 → 자동화 우선순위 도출':
        (25, 'lo', '%', 22),
    'Ansible Tower/AWX 실행 로그 + 수동 패치 작업 일지 병행 집계 | 분기 패치 관리 리뷰':
        (70, 'hi', '%', 65),
    '인시던트 포스트모템 RCA 결과에서 원인 분류: "Configuration Error" 건수 집계':
        (0, 'zero', '건', 0),
}

for howto_text, (tgtNum, tgtDir, unit, demoVal) in kpi_additions.items():
    old = f"howto:'{howto_text}'}}"
    new = f"howto:'{howto_text}',tgtNum:{tgtNum},tgtDir:'{tgtDir}',unit:'{unit}',demoVal:{demoVal}}}"
    if old in html:
        html = html.replace(old, new, 1)
    else:
        print(f'[WARN] Not found: {howto_text[:60]}...')

# ─── 3. JS 상태 + 헬퍼 함수 추가 (FILTER 바로 뒤) ───────────────────────
NEW_STATE_FUNCS = """
const KPI_ACTUAL = {영업:{},연구소:{},사업수행:{}};

const DEMO_KPI = {
  영업:[23,11,4,2.4,18,4.1,4,21,62,105,87,92,28,5],
  연구소:[82,1.2,18,3,2,99.95,2,13,3.2,83,78,79,34,88,1],
  사업수행:[95,22,93,18,99.7,1.5,6.5,15,88,0,4.3,97,4.2,8,22,65,0]
};
const DEMO_PROGRESS = {
  영업:{O1:75,O2:88,O3:85,O4:63},
  연구소:{O1:95,O2:91,O3:90,O4:55},
  사업수행:{O1:96,O2:91,O3:98,O4:85}
};

function calcAch(kpi, val){
  const v = parseFloat(val);
  if(val===''||val===null||val===undefined||isNaN(v)) return null;
  if(kpi.tgtDir==='zero') return v===0 ? 100 : Math.max(0,Math.round(100-v*25));
  if(kpi.tgtDir==='lo'){
    if(v<=0) return null;
    return Math.min(200,Math.round((kpi.tgtNum/v)*100));
  }
  return Math.min(200,Math.round((v/kpi.tgtNum)*100));
}
function achChip(ach){
  if(ach===null) return '<span class="ach-chip ach-empty">미입력</span>';
  const cls = ach>=100?'ach-green':ach>=80?'ach-yellow':'ach-red';
  return `<span class="ach-chip ${cls}">${ach}%</span>`;
}
function avgKpiAch(key){
  const kpis = DATA[key].kpis;
  const vals = kpis.map((k,i)=>calcAch(k, KPI_ACTUAL[key][i])).filter(v=>v!==null);
  if(!vals.length) return null;
  return Math.round(vals.reduce((a,b)=>a+b,0)/vals.length);
}
"""
html = html.replace(
    "const FILTER = {영업:'all',연구소:'all',사업수행:'all'};",
    "const FILTER = {영업:'all',연구소:'all',사업수행:'all'};\n" + NEW_STATE_FUNCS
)

# ─── 4. renderTab: KPI 테이블 헤더·행 수정 + 데모 버튼 + 요약바 ──────────
OLD_TABLE_HEAD = """  const kpiRows = d.kpis.map((k,i)=>{
    const hidden = (activeFilter !== 'all' && k.kt !== activeFilter) ? ' kpi-row-hidden' : '';
    return `<tr class="kpi-row-clickable${k.tp==='lead'?' r-lead':''}${hidden}" data-idx="${i}" data-kt="${k.kt}">
      <td>${k.nm}<span class="kpi-hint">▸</span></td>
      <td>${badge(TP_CLASS[k.tp],TP_LABEL[k.tp])}</td>
      <td>${badge(KT_CLASS[k.kt],KT_LABEL[k.kt])}</td>
      <td><strong>${k.tgt}</strong></td>
      <td><span class="cycle">${k.cy}</span></td>
    </tr>`;
  }).join('');"""

NEW_TABLE_HEAD = """  const kpiRows = d.kpis.map((k,i)=>{
    const hidden = (activeFilter !== 'all' && k.kt !== activeFilter) ? ' kpi-row-hidden' : '';
    const actVal = KPI_ACTUAL[key][i]??'';
    const ach = calcAch(k, actVal);
    return `<tr class="kpi-row-clickable${k.tp==='lead'?' r-lead':''}${hidden}" data-idx="${i}" data-kt="${k.kt}">
      <td>${k.nm}<span class="kpi-hint">▸</span></td>
      <td>${badge(TP_CLASS[k.tp],TP_LABEL[k.tp])}</td>
      <td>${badge(KT_CLASS[k.kt],KT_LABEL[k.kt])}</td>
      <td><strong>${k.tgt}</strong></td>
      <td class="col-act" onclick="event.stopPropagation()">
        <input class="kpi-act-input" type="number" step="any"
          placeholder="입력" value="${actVal}" data-kpi-idx="${i}">
        <span class="unit-lbl">${k.unit??''}</span>
      </td>
      <td class="col-ach" id="ach-${key}-${i}">${achChip(ach)}</td>
      <td><span class="cycle">${k.cy}</span></td>
    </tr>`;
  }).join('');"""

html = html.replace(OLD_TABLE_HEAD, NEW_TABLE_HEAD)

OLD_KPI_SECTION = """  <div class="sec-label">KPI 종합 지표표</div>
  <div class="card-tbl" id="kpi-section">
    <div class="card-tbl-head">📊 Leading · Lagging KPI 통합 요약 <span class="hint-txt">· 행 클릭 시 상세</span></div>
    <div class="kpi-filter-bar">${filterBtns}</div>
    <div class="tbl-wrap">
      <table>
        <thead><tr>
          <th>KPI 항목</th><th>구분</th><th>유형</th><th>목표치</th><th>측정 주기</th>
        </tr></thead>
        <tbody id="kpi-tbody">${kpiRows}</tbody>
      </table>
    </div>
    <div class="legend">"""

NEW_KPI_SECTION = """  <div class="sec-label">KPI 실적 입력 &amp; 달성률</div>
  <div class="card-tbl" id="kpi-section">
    <div class="card-tbl-head" style="display:flex;align-items:center;gap:6px;flex-wrap:wrap">
      📊 KPI 실적 입력 &amp; 달성률 자동계산
      <span class="hint-txt">· 실적 입력 시 달성률 자동 계산 · 행 클릭 시 상세</span>
      <button class="demo-btn" id="demo-btn">🎲 데모 프리셋</button>
      <button class="reset-btn" id="reset-btn">↺ 초기화</button>
    </div>
    <div class="kpi-filter-bar">${filterBtns}</div>
    <div class="tbl-wrap">
      <table>
        <thead><tr>
          <th>KPI 항목</th><th>구분</th><th>유형</th><th>목표치</th>
          <th class="col-act">실적 입력</th><th class="col-ach">달성률</th><th>측정 주기</th>
        </tr></thead>
        <tbody id="kpi-tbody">${kpiRows}</tbody>
      </table>
    </div>
    <div class="kpi-ach-summary" id="kpi-ach-summary">
      <span class="kpi-ach-lbl">KPI 평균 달성률</span>
      <span class="kpi-ach-val" id="kpi-ach-val">—</span>
      <div class="kpi-ach-bar"><div class="kpi-ach-bar-fill" id="kpi-ach-fill"></div></div>
      <span class="kpi-ach-sub" id="kpi-ach-sub">실적을 입력하면 달성률이 자동 계산됩니다</span>
    </div>
    <div class="legend">"""

html = html.replace(OLD_KPI_SECTION, NEW_KPI_SECTION)

# ─── 5. attachHandlers: KPI 입력 이벤트 + 데모/초기화 버튼 추가 ──────────
OLD_FILTER_HANDLER = """  // [F5] Filter buttons
  document.querySelectorAll('.kpi-filter-btn').forEach(btn=>{"""

NEW_KPI_HANDLERS = """  // [KPI] 실적 입력 → 달성률 계산
  function updateKpiAchSummary(){
    const avg = avgKpiAch(key);
    const valEl = document.getElementById('kpi-ach-val');
    const fillEl = document.getElementById('kpi-ach-fill');
    const subEl = document.getElementById('kpi-ach-sub');
    if(valEl){
      if(avg===null){
        valEl.textContent='—';
        if(fillEl) fillEl.style.width='0%';
        if(subEl) subEl.textContent='실적을 입력하면 달성률이 자동 계산됩니다';
      } else {
        valEl.textContent=avg+'%';
        valEl.style.color = avg>=100?'#4ade80':avg>=80?'#fbbf24':'#f87171';
        if(fillEl) fillEl.style.width=Math.min(100,avg)+'%';
        const entered = DATA[key].kpis.filter((_,i)=>KPI_ACTUAL[key][i]!==undefined&&KPI_ACTUAL[key][i]!=='').length;
        if(subEl) subEl.textContent=`${entered}개 항목 입력됨 · 미입력 항목 제외 평균`;
      }
    }
  }
  document.querySelectorAll('.kpi-act-input').forEach(el=>{
    el.addEventListener('input', e=>{
      e.stopPropagation();
      const idx = parseInt(el.dataset.kpiIdx);
      KPI_ACTUAL[key][idx] = el.value;
      const kpi = DATA[key].kpis[idx];
      const ach = calcAch(kpi, el.value);
      const cell = document.getElementById(`ach-${key}-${idx}`);
      if(cell) cell.innerHTML = achChip(ach);
      updateKpiAchSummary();
    });
    el.addEventListener('click', e=>e.stopPropagation());
    el.addEventListener('keydown', e=>e.stopPropagation());
  });
  updateKpiAchSummary();

  // [DEMO] 데모 프리셋
  document.getElementById('demo-btn')?.addEventListener('click', ()=>{
    const vals = DEMO_KPI[key];
    vals.forEach((v,i)=>{
      KPI_ACTUAL[key][i] = v;
    });
    Object.assign(PROGRESS[key], DEMO_PROGRESS[key]);
    setTab(key);
  });
  // [RESET] 초기화
  document.getElementById('reset-btn')?.addEventListener('click', ()=>{
    KPI_ACTUAL[key] = {};
    PROGRESS[key] = {O1:0,O2:0,O3:0,O4:0};
    setTab(key);
  });

  // [F5] Filter buttons
  document.querySelectorAll('.kpi-filter-btn').forEach(btn=>{"""

html = html.replace(OLD_FILTER_HANDLER, NEW_KPI_HANDLERS)

# ─── 저장 ────────────────────────────────────────────────────────────────
with open(PATH, 'w', encoding='utf-8') as f:
    f.write(html)

print('패치 완료. index.html 저장됨.')
