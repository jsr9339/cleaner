# 상담원장_메모없음.xlsx 정리 — 작성기준 시트가 없는 파일. 규칙은 모두 데이터를 보고 정했다(근거는 정리본/메모없음_변경요약.md).
# 원본은 읽기만 한다. 쓰는 곳: 정리본/메모없음_정리.csv · 정리본/메모없음_지운줄.csv · 정리본/메모없음_변경요약.md
from pathlib import Path
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT.parent / 'messy' / '상담원장_메모없음.xlsx'
OUT = ROOT / '정리본'
CHANNEL = {'유선': '전화', '전화상담': '전화', 'TEL': '전화', 'chatbot': '챗봇', 'Chatbot': '챗봇', '챗 봇': '챗봇',
           'APP': '앱', 'app': '앱', '모바일앱': '앱', 'e-mail': '이메일', 'E-mail': '이메일', '메일': '이메일'}
STATUS = {'처리완료': '완료', '종결': '완료', '처리중': '진행중', '대기': '보류'}
REASK = {'y': 'Y', '예': 'Y', 'TRUE': 'Y', 'n': 'N', '아니오': 'N', 'FALSE': 'N'}
MARK_NEG = '처리시간이 음수 — 틀린 값이지만 진짜 값을 몰라 지우지 않고 그대로 둠'
MARK_SAT = '만족도 6~10점 — 1~5점 척도를 벗어남. 지우지 않고 두며 평균 만족도에서 뺌'

raw = pd.read_excel(SRC, sheet_name=0, dtype=str)
raw.insert(0, '원본행번호', raw.index + 2)
cols = [c for c in raw.columns if c != '원본행번호']
d = raw.copy()
for c in cols:
    d[c] = d[c].where(d[c].isna(), d[c].str.strip())

removed = pd.Series('', index=d.index)
test = d.상담원ID == 'AG999'
removed[test] = '테스트 접수(상담원ID AG999)'
d = d[~test].copy()

d['접수일'] = pd.to_datetime(d.접수일.str.replace(r'[/.]', '-', regex=True)).dt.strftime('%Y-%m-%d')
d['접수채널'] = d.접수채널.replace(CHANNEL)
d['상담원팀'] = d.상담원팀.str.replace(' ', '', regex=False)
d['지역'] = d.지역.replace({'경기도': '경기', '서울시': '서울', '서울특별시': '서울'})
d['월정액'] = pd.to_numeric(d.월정액.str.replace(r'[,원]', '', regex=True))
d['처리상태'] = d.처리상태.replace(STATUS)
d['최종수정일'] = pd.to_datetime(d.최종수정일.str.replace(r'[/.]', '-', regex=True)).dt.strftime('%Y-%m-%d')

d = d.sort_values(['접수번호', '최종수정일'], kind='stable')
drop = d.duplicated('접수번호', keep='last')
kept = d[~drop].set_index('접수번호').최종수정일
for i in d.index[drop]:
    same = d.at[i, '최종수정일'] == kept[d.at[i, '접수번호']]
    removed[i] = '중복 추출(같은 접수번호·같은 최종수정일)' if same else '같은 접수번호의 더 오래된 줄(최종수정일이 빠름)'
d = d[~drop].copy()

# 상담시간: 값의 모양이 아니라 «접수일 7월 1일 이후 + 전화» 로 가른다
t0 = pd.to_numeric(d.상담시간)
sec = (d.접수일 >= '2025-07-01') & (d.접수채널 == '전화')
d['상담시간'] = np.where(sec, t0 / 60, t0).round(2)

d['재문의'] = d.재문의.replace(REASK)          # 빈칸은 «기록 없음» 그대로
sat0 = pd.to_numeric(d.만족도.str.replace('점', '', regex=False), errors='coerce')
sat_blank = d.만족도.isin(['0', '-', '미응답'])
d['만족도'] = sat0.where(sat0 >= 1)
sat_hi = d.만족도 > 5
neg = pd.to_numeric(d.처리시간) < 0
mark = pd.Series('', index=d.index)
mark[neg] = MARK_NEG
mark[sat_hi] = np.where(neg[sat_hi], MARK_NEG + ' / ' + MARK_SAT, MARK_SAT)
d['_표시'] = mark.where(mark != '')
d = d.drop(columns=['원본행번호']).sort_values('접수번호')
d.to_csv(OUT / '메모없음_정리.csv', index=False, encoding='utf-8-sig')
rm = raw[removed != ''].copy()
rm['지운이유'] = removed[removed != '']
rm.to_csv(OUT / '메모없음_지운줄.csv', index=False, encoding='utf-8-sig')

sat_ok = d.만족도.where(d.만족도.between(1, 5))
asked = d.재문의.dropna()
print('줄', len(raw), '→', len(d), removed[removed != ''].value_counts().to_dict())
print('초→분', int(sec.sum()), '만족도 비움', int(sat_blank.sum()), '6~10점', int(sat_hi.sum()), '음수', int(neg.sum()))
print({'상담 건수': len(d), '평균 상담시간(분)': round(d.상담시간.mean(), 2),
       '재문의율(%)': round((asked == 'Y').mean() * 100, 2), '평균 만족도': round(sat_ok.mean(), 2),
       '완료율(%)': round((d.처리상태 == '완료').mean() * 100, 2)}, '고객 수', d.고객번호.nunique())
print('상담시간 60 이상 비전화/7월전:', int(((d.상담시간 >= 60) & ~sec).sum()), '최대', d.상담시간.max())
