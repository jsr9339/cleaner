# 상담원장_2025.xlsx 정리
# 결정 기록: my-wiki/decisions/d-007-상담원장-정리-다섯-결정.md
#   d-007 결정 1: 종결 = 완료, 대기 = 보류 (CS기획팀 답)
#   d-007 결정 2: 처리시간 음수는 지우지 않고 값 그대로, 맨 끝 _표시 칸에 이유
#   d-007 결정 3: 만족도 6~10점은 지우지 않고 _표시 칸에 이유. 평균 만족도에서는 표시한 줄을 뺀다
#   d-007 결정 4: 칸 이름은 원본 그대로 (끝에 _표시 칸만 더한다)
#   d-007 결정 5: 고객번호 앞뒤 공백 제거
# 그 밖의 처리(테스트·중복 제거, 초→분, 표기 통일 등)는 상담원장 «작성기준» 시트와 앞서 정한 대로다.
#
# 원본은 읽기만 한다. 쓰는 곳: 정리본/상담원장_정리.csv · 정리본/지운줄.csv · 정리본/변경요약.md
import sys
from pathlib import Path
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent          # cleaner/
SRC = ROOT.parent / 'messy' / '상담원장_2025.xlsx'
OUT_DIR = ROOT / '정리본'
OUT_CSV = OUT_DIR / '상담원장_정리.csv'
OUT_RM = OUT_DIR / '지운줄.csv'
OUT_MD = OUT_DIR / '변경요약.md'

MARK_NEG = '처리시간이 음수 — 틀린 값이지만 진짜 값을 몰라 지우지 않고 그대로 둠'
MARK_SAT = '만족도 6~10점 — NPS를 잘못 넣은 것으로 보이나 확인 못 함. 지우지 않고 두며 평균 만족도에서 뺌'
CHANNEL = {'유선': '전화', '전화상담': '전화', 'TEL': '전화', 'chatbot': '챗봇', 'Chatbot': '챗봇', '챗 봇': '챗봇',
           'APP': '앱', 'app': '앱', '모바일앱': '앱', 'e-mail': '이메일', 'E-mail': '이메일', '메일': '이메일'}
STATUS = {'처리완료': '완료', '종결': '완료', '처리중': '진행중', '대기': '보류'}      # 종결·대기: d-007 결정 1
REASK = {'y': 'Y', '예': 'Y', 'TRUE': 'Y', 'n': 'N', '아니오': 'N', 'FALSE': 'N'}

# ── 읽기 (원본은 읽기만) ──────────────────────────────────────────────
raw = pd.read_excel(SRC, sheet_name=0, dtype=str)
raw.insert(0, '원본행번호', raw.index + 2)                 # 엑셀 행 번호 (머리글 = 1행)
cols = [c for c in raw.columns if c != '원본행번호']
N0 = len(raw)

# 앞뒤 공백 제거 (d-007 결정 5 포함) — 줄 수를 세려고 바뀐 칸을 센다
d = raw.copy()
strip_cells = {}
for c in cols:
    s = d[c]
    t = s.where(s.isna(), s.str.strip())
    strip_cells[c] = int((s.fillna('') != t.fillna('')).sum())
    d[c] = t

# ── 1. 테스트 접수 제거 ───────────────────────────────────────────────
is_test = d.상담원ID == 'AG999'
removed = pd.Series('', index=d.index)
removed[is_test] = '테스트 접수(상담원ID AG999)'
d = d[~is_test].copy()

# ── 2. 표기 통일 (날짜·채널·팀·지역·월정액) ─────────────────────────────
before = d.copy()
d['접수일'] = pd.to_datetime(d['접수일'].str.replace(r'[/.]', '-', regex=True)).dt.strftime('%Y-%m-%d')
d['접수채널'] = d.접수채널.replace(CHANNEL)
d['상담원팀'] = d.상담원팀.str.replace(' ', '', regex=False)
d['지역'] = d.지역.replace({'경기도': '경기', '서울시': '서울', '서울특별시': '서울'})
d['월정액'] = pd.to_numeric(d.월정액.str.replace(r'[,원]', '', regex=True))
status_before = d.처리상태.copy()
d['처리상태'] = d.처리상태.replace(STATUS)

# ── 3. 같은 접수번호는 최종수정일이 늦은 줄만 남긴다 ───────────────────────
d = d.sort_values(['접수번호', '최종수정일'], kind='stable')
drop = d.duplicated('접수번호', keep='last')
kept_dt = d[~drop].set_index('접수번호').최종수정일
for i in d.index[drop]:
    same = d.at[i, '최종수정일'] == kept_dt[d.at[i, '접수번호']]
    removed[i] = ('중복 추출(같은 접수번호·같은 최종수정일, 날짜 표기만 다를 수 있음)' if same
                  else '상태 변경 — 같은 접수번호의 더 오래된 줄(최종수정일이 빠름)')
d = d[~drop].copy()
before = before.loc[d.index]; status_before = status_before.loc[d.index]

# ── 4. 상담시간: 접수일 7월 1일 이후 전화는 초 단위 → 분 (값의 크기가 아니라 날짜+채널로 가른다) ──
t0 = pd.to_numeric(d.상담시간)
sec = (d.접수일 >= '2025-07-01') & (d.접수채널 == '전화')
d['상담시간'] = np.where(sec, t0 / 60, t0).round(2)

# ── 5. 재문의: 빈칸은 «기록 없음» 그대로 ─────────────────────────────────
d['재문의'] = d.재문의.replace(REASK)

# ── 6. 만족도: 0 · «-» · «미응답» 은 빈칸, «n점» 은 숫자. 6~10점은 지우지 않는다(결정 3) ──
sat0 = pd.to_numeric(d.만족도.str.replace('점', '', regex=False), errors='coerce')
sat_blank = d.만족도.isin(['0', '-', '미응답'])
d['만족도'] = sat0.where(sat0 >= 1)
sat_hi = d.만족도 > 5

# ── 7. 처리시간 음수는 지우지 않는다(결정 2) ───────────────────────────────
neg = pd.to_numeric(d.처리시간) < 0

# ── 8. _표시 칸 (맨 끝). 칸 이름은 원본 그대로(결정 4) ─────────────────────
mark = pd.Series('', index=d.index)
mark[neg] = MARK_NEG
mark[sat_hi] = np.where(neg[sat_hi], MARK_NEG + ' / ' + MARK_SAT, MARK_SAT)
d['_표시'] = mark.where(mark != '')
d = d.drop(columns=['원본행번호']).sort_values('접수번호')
assert list(d.columns[:-1]) == cols and d.columns[-1] == '_표시'
d.to_csv(OUT_CSV, index=False, encoding='utf-8-sig')

# ── 지운 줄 (원본 값 그대로 + 지운이유) ───────────────────────────────────
rm = raw[removed != ''].copy()
rm['지운이유'] = removed[removed != '']
rm.to_csv(OUT_RM, index=False, encoding='utf-8-sig')

# ── 다섯 숫자와 고객 수 ───────────────────────────────────────────────────
sat_ok = d.만족도.where(d.만족도.between(1, 5))                 # 표시한 줄(6~10점)은 뺀다
asked = d.재문의.dropna()
five = {
    '상담 건수': len(d),
    '평균 상담시간(분)': round(d.상담시간.mean(), 2),
    '재문의율(%)': round((asked == 'Y').mean() * 100, 2),
    '평균 만족도': round(sat_ok.mean(), 2),
    '완료율(%)': round((d.처리상태 == '완료').mean() * 100, 2),
}
n_cust = d.고객번호.nunique()
print(five, '고객 수', n_cust)

# ── 변경요약.md ───────────────────────────────────────────────────────────
def n(x): return f'{int(x):,}'
vc = removed[removed != ''].value_counts()
rt = status_before.value_counts()
changed = lambda a, b: int((a.fillna('').astype(str) != b.fillna('').astype(str)).sum())
cust_strip = strip_cells['고객번호']
# 고객번호 앞뒤 공백: 남은 줄 기준
cust_strip_kept = int((raw.loc[d.index, '고객번호'] != raw.loc[d.index, '고객번호'].str.strip()).sum())
ch_kept = int((before.접수채널 != d.접수채널).sum())
dt_kept = int((before.접수일.str.match(r'^\d{4}-\d{2}-\d{2}$') == False).sum())
fee_kept = int(before.월정액.astype(str).str.contains('원').sum())
reg_kept = int(before.지역.isin(['경기도', '서울시', '서울특별시']).sum())
team_kept = int(before.상담원팀.str.contains(' ').sum())
reask_kept = int(before.재문의.isin(REASK).sum())
sat_pt = int(before.만족도.str.contains('점', na=False).sum())
sat_cnt = lambda v: int((before.만족도 == v).sum())

L = []
L += ['# 상담원장 정리 — 변경 요약', '',
      '원본: `messy/상담원장_2025.xlsx` (읽기만 했다) · 스크립트: `정리/정리_상담원장.py` · 결정 기록: d-007', '',
      '## 줄 수', '',
      '| 단계 | 줄 수 |', '|---|---:|',
      f'| 원본 | {n(N0)} |',
      f'| − 테스트 접수 (상담원ID AG999) | {n(vc.get("테스트 접수(상담원ID AG999)", 0))} |',
      f'| − 중복 추출 (같은 접수번호·같은 최종수정일) | {n(vc.get("중복 추출(같은 접수번호·같은 최종수정일, 날짜 표기만 다를 수 있음)", 0))} |',
      f'| − 상태 변경으로 생긴 옛 줄 (같은 접수번호, 최종수정일이 더 빠른 줄) | {n(vc.get("상태 변경 — 같은 접수번호의 더 오래된 줄(최종수정일이 빠름)", 0))} |',
      f'| 정리본 | {n(len(d))} |', '',
      f'지운 줄은 모두 `정리본/지운줄.csv` 에 있다(맨 앞에 원본행번호, 맨 끝에 지운이유). 지운 줄 합계 {n((removed != "").sum())}줄.', '',
      '## 상담시간 — 7월 이후 전화 줄은 초에서 분으로 바꿨다', '',
      f'- 접수일이 2025-07-01 이후인 **전화** 접수 {n(sec.sum())}줄은 원본에 초 단위로 적혀 있어 60으로 나눠 분으로 바꿨다(소수 둘째 자리까지). 작성기준 시트에 적힌 사실이다.',
      f'- 같은 시기의 챗봇·앱·이메일 줄과 7월 전 전화 줄은 분 단위라 그대로 두었다.', '',
      '## 바꾼 값 (원본 대비, 남은 줄 기준)', '',
      '| 칸 | 무엇을 | 줄 수 |', '|---|---|---:|',
      f'| 상담시간 | 7월 1일 이후 전화: 초 → 분 | {n(sec.sum())} |',
      f'| 처리상태 | 종결 → 완료 (CS기획팀 답, 결정 1) | {n(rt.get("종결", 0))} |',
      f'| 처리상태 | 처리완료 → 완료 | {n(rt.get("처리완료", 0))} |',
      f'| 처리상태 | 대기 → 보류 (CS기획팀 답, 결정 1) | {n(rt.get("대기", 0))} |',
      f'| 처리상태 | 처리중 → 진행중 | {n(rt.get("처리중", 0))} |',
      f'| 만족도 | 0 · «-» · «미응답» → 빈칸 | {n(sat_blank.sum())} |',
      f'| 만족도 | «n점» → 숫자 | {n(sat_pt)} |',
      f'| 만족도 | 6~10점은 값을 그대로 두고 `_표시` 에 이유 (결정 3) | {n(sat_hi.sum())} |',
      f'| 처리시간 | 음수는 값을 그대로 두고 `_표시` 에 이유 (결정 2) | {n(neg.sum())} |',
      f'| 고객번호 | 앞뒤 공백 제거 (결정 5) | {n(cust_strip_kept)} |',
      f'| 접수채널 | 17가지 표기 → 전화·챗봇·앱·이메일 | {n(ch_kept)} |',
      f'| 재문의 | y · 예 · TRUE · n · 아니오 · FALSE → Y/N | {n(reask_kept)} |',
      f'| 접수일 | 표기를 YYYY-MM-DD 로 통일 | {n(dt_kept)} |',
      f'| 월정액 | «원»·쉼표 제거 | {n(fee_kept)} |',
      f'| 지역 | 경기도 · 서울시 · 서울특별시 → 경기 · 서울 | {n(reg_kept)} |',
      f'| 상담원팀 | 공백 제거 («요금 팀» → «요금팀») | {n(team_kept)} |',
      f'| _표시 | 칸 하나를 맨 끝에 더함. 채워진 줄: 음수 {n(neg.sum())} + 만족도 6~10점 {n(sat_hi.sum())} | {n((mark != "").sum())} |',
      '', '앞뒤 공백을 지우면서 바뀐 칸은 모든 칸을 합쳐 지운 줄까지 포함해 '
      f'{n(sum(strip_cells.values()))}칸이다(그중 고객번호 {n(cust_strip)}칸).', '',
      '## 바꾸지 않은 것', '',
      f'- 칸 이름: 원본 {len(cols)}칸의 이름과 순서 그대로 (결정 4). `_표시` 는 끝에 더한 칸이다.',
      '- 재문의 빈칸: 3월 이전 접수분은 «기록 없음» 이라 N으로 바꾸지 않았다.',
      f'- 처리시간 음수 {n(neg.sum())}줄과 만족도 6~10점 {n(sat_hi.sum())}줄은 지우지 않았다.', '']
OUT_MD.write_text('\n'.join(L), encoding='utf-8')
