with open('index.html', encoding='utf-8') as f:
    h = f.read()

checks = [
    ('CSS ach-chip', '.ach-chip{' in h),
    ('CSS demo-btn', '.demo-btn{' in h),
    ('CSS kpi-ach-summary', '.kpi-ach-summary{' in h),
    ('KPI_ACTUAL state', 'const KPI_ACTUAL' in h),
    ('DEMO_KPI data', 'const DEMO_KPI' in h),
    ('calcAch function', 'function calcAch' in h),
    ('achChip function', 'function achChip' in h),
    ('tgtNum:30 in data', 'tgtNum:30,' in h),
    ('tgtDir hi', "tgtDir:'hi'" in h),
    ('tgtDir lo', "tgtDir:'lo'" in h),
    ('tgtDir zero', "tgtDir:'zero'" in h),
    ('demoVal:23', 'demoVal:23}' in h),
    ('kpi-act-input class', 'kpi-act-input' in h),
    ('col-act th', 'col-act' in h),
    ('달성률 th', '달성률' in h),
    ('demo-btn id', 'id="demo-btn"' in h),
    ('reset-btn id', 'id="reset-btn"' in h),
    ('kpi-ach-summary id', 'id="kpi-ach-summary"' in h),
    ('updateKpiAchSummary fn', 'updateKpiAchSummary' in h),
    ('DEMO_PROGRESS const', 'const DEMO_PROGRESS' in h),
]
all_ok = True
for name, ok in checks:
    status = 'OK' if ok else 'FAIL'
    if not ok: all_ok = False
    print(f'[{status}] {name}')
print()
print('ALL PASS' if all_ok else 'SOME FAILED - check above')

# Count how many KPIs got tgtNum
cnt = h.count('tgtNum:')
print(f'tgtNum count: {cnt} (expect 46)')
