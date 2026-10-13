#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Gmail 정리 도구(gmail_to_sheets.gs) 가이드북 생성.

2026-10-13 사용자 요청: 구글 문서 'gmail정리 참고' 의 설명과 캡처 3장을 바탕으로 가이드북을 만들 것.
서식은 공용 docx_kit (2026-10-13 판: Noto Sans KR, Heading 스타일 → 구글 문서 목차).

실행: python3 build_guidebook.py   → docs/Gmail정리_가이드북.docx
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docx_kit import Doc, S  # noqa: E402

OUT = os.path.join(HERE, 'Gmail정리_가이드북.docx')
IMG = lambda n: os.path.join(HERE, 'img', n)  # noqa: E731
B = {'bold': True}
ACC = {'bold': True, 'color': 'accent'}
RED = {'bold': True, 'color': 'danger'}

d = Doc()
d.title('Gmail 정리 도구 가이드북',
        subtitle='받은편지함 메일을 구글 시트로 모으고 AI 로 요약하기 | 2026-10-13 · 작성 Claude Code')

# ── 0 ────────────────────────────────────────────────────────────────
d.h1('0. 한눈에 보기')
d.p('받은편지함에 쌓인 메일을 구글 스프레드시트 ', ('mails', B), ' 시트로 옮겨 적고, 긴급도를 자동으로 분류하고, '
    '구글 시트의 AI() 함수로 내용을 요약하는 도구입니다. 시트로 옮긴 메일은 받은편지함에서 보관(archive)되어 '
    '받은편지함이 비워집니다.')
d.code('Gmail 받은편지함 (별표 없는 메일, 최대 50개)',
       '   ↓  Email Tools > Get Emails (Popup)',
       'mails 시트에 한 줄씩 기록  →  F열 긴급도 자동 분류  →  G열 =AI() 요약 수식',
       '   ↓',
       'Gmail 에서는 보관(archive) 처리  →  받은편지함에서 사라짐',
       caption='한 번 실행할 때 일어나는 일')
d.table(['하고 싶은 것', '메뉴', '장'], [
    ['메일을 시트로 모으기', 'Get Emails (Popup)', '2'],
    ['AI 요약을 값으로 굳히기', 'Freeze AI() Summaries to Values', '4'],
    ['색·서식이 흐트러졌을 때 다시 칠하기', 'Refresh Formatting', '5'],
    ['시트를 비우고 새로 시작', 'Clear Emails Contents', '6'],
    ['분류 키워드·색 바꾸기', '(Apps Script 편집기)', '7'],
], widths=[6.5, 7.5, 3])
d.callout('계속 받은편지함에 두고 챙겨야 할 메일은 Gmail 에서 별표(★)를 붙여 두세요. '
          '별표가 있는 메일은 모으지도, 보관하지도 않습니다.', kind='note', head='팁')

# ── 1 ────────────────────────────────────────────────────────────────
d.h1('1. 메뉴 사용법 (Email Tools)')
d.p('스프레드시트를 열면 상단 메뉴줄 맨 오른쪽에 ', ('Email Tools', B), ' 메뉴가 생깁니다. '
    '실행하고 싶은 기능을 고르기만 하면 됩니다.')
d.image(IMG('01_menu.png'), width_cm=11, caption='그림 1. Email Tools 메뉴')
d.table(['메뉴', '하는 일', '실행되는 함수'], [
    ['Get Emails (Popup)', '기간을 입력받아 받은편지함 메일을 시트로 모으고, 분류·요약 수식·서식까지 한 번에 처리', 'fetchUnreadEmails'],
    ['Freeze AI() Summaries to Values', 'G열의 =AI() 수식 중 요약이 끝난 칸을 글자(값)로 바꿈', 'freezeAiSummaries'],
    ['Refresh Formatting (Thread 색상 / Urgency 서식)', '전체 행에 줄무늬·내 메일 표시·긴급도 색을 다시 칠함', 'refreshFormatting'],
    ['Clear Emails Contents', '머리글(1행)만 남기고 A~K열 내용과 서식을 지움', 'clearSpecificFormats'],
], widths=[5.2, 8.3, 3.5], mono_cols=(2,))
d.callout('메뉴가 안 보이면 스프레드시트를 새로고침(F5) 하세요. 메뉴는 시트를 열 때(onOpen) 만들어집니다. '
          '처음 실행할 때는 Gmail·시트 접근 권한을 허용하는 창이 뜹니다.', kind='warn', head='메뉴가 없을 때')

# ── 2 ────────────────────────────────────────────────────────────────
d.h1('2. Get Emails (Popup) — 메일 모으기')
d.h2('조회 기간 입력')
d.p('메뉴를 고르면 조회 기간을 묻는 창이 뜹니다. 기본 기간(어제 ~ 내일)이 안내되어 있어서, ',
    ('빈칸 그대로 확인을 누르면 기본 기간으로 조회', ACC), '합니다.')
d.image(IMG('02_popup.png'), width_cm=12, caption='그림 2. 조회 기간 입력 창')
d.table(['입력 예', '조회되는 메일'], [
    ['(빈칸)', '기본값: 어제 0시 ~ 오늘 끝 (내일 0시 전)'],
    ['2026/10/01 ~ 2026/10/10', '10월 1일 0시 ~ 10월 9일 끝'],
    ['2026/10/01, 2026/10/10', '위와 같음 (쉼표로 구분해도 됨)'],
    ['어제 ~ 내일  /  오늘, 내일', '단어도 됩니다. 어제·오늘·내일 은 실행한 날 기준 날짜로 바뀝니다'],
], widths=[6, 11], mono_cols=(0,))
d.callout('종료일은 "그날 0시 전까지" 입니다 (Gmail 의 before: 규칙). 10월 10일 메일까지 포함하려면 종료일을 10월 11일로 쓰세요.',
          kind='warn', head='종료일 주의')

d.h2('어떤 메일을 모으나')
d.bullet(['받은편지함', ('(inbox)', {'color': 'muted'}), ' 에 있는 메일'])
d.bullet([('별표(★)가 없는', B), ' 메일. 받은편지함에 남겨 두고 후속 조치(follow up)할 메일은 별표를 붙이면 제외됩니다.'])
d.bullet([('한 번에 최대 50개', B), ' 대화(스레드). 많이 모으면 Apps Script 실행 시간 제한으로 오류가 날 수 있어 50개로 제한했습니다. '
          '더 있으면 메뉴를 여러 번 실행해 누적해서 모으세요 (이미 모은 메일은 받은편지함에서 빠졌으므로 겹치지 않습니다).'])
d.bullet(['대화 하나에서는 ', ('가장 최근 메일 1통', B), '만 기록합니다. 본문은 앞부분 400자까지.'])
d.callout('모은 메일은 Gmail 에서 보관(archive) 처리되어 받은편지함에서 사라집니다. 지워지는 것은 아닙니다. '
          'Gmail 의 "전체보관함" 에서 언제든 다시 볼 수 있습니다.', kind='note', head='메일은 지워지지 않는다')

d.h2('실행하면 일어나는 일 (순서대로)')
d.steps([
    '지난번 실행 때 넣은 G열 =AI() 수식 중 요약이 끝난 것을 값으로 굳힙니다 (4장).',
    '받은편지함에서 조건에 맞는 메일을 찾아 시트 맨 아래에 이어 적습니다.',
    'F열에 긴급도(🔴 긴급 / 🟡 요청/확인 / ⚪ 단순공지 / 일반)를 자동으로 적습니다 (7장).',
    '모은 메일을 Gmail 에서 보관합니다.',
    '전체 데이터를 Thread ID → 날짜 순으로 정렬해 같은 대화끼리 모읍니다.',
    'G열이 빈 행에 =AI() 요약 수식을 넣습니다.',
    '줄무늬 · 내 메일 표시 · 긴급도 색을 칠하고 완료 알림을 띄웁니다.',
])

# ── 3 ────────────────────────────────────────────────────────────────
d.h1('3. 시트 열 구성 (mails 시트)')
d.table(['열', '이름', '내용'], [
    ['A', 'Date', '받은 일시 (yyyy-mm-dd hh:mm:ss)'],
    ['B', 'Sender', '보낸 사람'],
    ['C', 'Recipient', '받는 사람 (To). 여기에 내 주소가 있으면 행 전체가 파란 굵은 기울임 글씨'],
    ['D', 'Labels', 'Gmail 라벨 (쉼표로 구분)'],
    ['E', 'Subject', '제목'],
    ['F', 'Urgency', '긴급도 자동 분류 결과 (제목 + 본문 + 라벨에서 키워드 검사)'],
    ['G', 'Summary', '=AI() 요약. 요약이 끝나면 다음 실행 때 값으로 굳어짐'],
    ['H', 'Snippet', '본문 앞부분 400자'],
    ['I', 'Note', '비워 둔 칸. 메모용으로 자유롭게 쓰세요'],
    ['J', 'Thread ID', 'Gmail 대화 고유번호. 정렬과 줄무늬의 기준'],
    ['K', 'Full Content', '"Subject: 제목 / Snippet: 본문" — AI 요약에 넘기는 글'],
], widths=[1.2, 3, 12.8])

# ── 4 ────────────────────────────────────────────────────────────────
d.h1('4. G열 AI 요약 쓰는 법')
d.p('G열(Summary)은 구글 스프레드시트의 ', ('AI() 함수', B), '로 K열(제목 + 본문 일부)을 요약합니다. '
    '스크립트는 수식만 넣고 끝나며, 실제 요약은 시트에서 따로 진행됩니다.')
d.image(IMG('03_ai_fill.png'), width_cm=13, caption='그림 3. =AI() 수식이 들어간 G열과 화면 아래 "생성 및 채우기" 버튼')
d.h2('요약을 실제로 만드는 방법')
d.steps([
    'Get Emails 실행이 끝나면 G열에 =AI(K29 & "…") 같은 수식이 들어가 있습니다.',
    '화면 아래쪽에 뜨는 [생성 및 채우기] 버튼을 누릅니다 (그림 3 아래). 오른쪽 숫자는 생성할 AI 셀 개수입니다.',
    '잠시 기다리면 각 칸에 요약 글이 채워집니다.',
])
d.callout('스크립트가 수식을 넣어도 [생성 및 채우기] 를 누르기 전에는 요약이 만들어지지 않습니다. '
          '버튼이 안 보이면 G열의 AI 수식 칸을 하나 클릭해 보세요.', kind='warn', head='버튼은 직접 눌러야 한다')
d.h2('요약을 값으로 굳히기')
d.p('AI() 수식은 시트를 열 때마다 다시 계산될 수 있어서, 요약이 끝난 칸은 글자(값)로 바꿔 두는 것이 안전합니다.')
d.bullet([('자동', B), ': Get Emails 를 실행할 때마다 지난번 수식 중 요약이 끝난 칸을 먼저 값으로 굳힙니다.'])
d.bullet([('수동', B), ': 바로 굳히고 싶으면 메뉴 Freeze AI() Summaries to Values 를 실행합니다.'])
d.bullet(['아직 생성 중이거나 오류(#ERROR! 등)인 칸은 건너뛰고 수식으로 남겨 둡니다. 완료 알림에 건수가 나옵니다.'])
d.callout('요약 문구(프롬프트)는 Apps Script 의 SUMMARY_PROMPT 변수에서 바꿉니다 (7장). '
          '이미 들어간 수식은 바뀌지 않고, 다음에 새로 들어가는 수식부터 적용됩니다.', kind='note', head='요약 방식 바꾸기')

# ── 5 ────────────────────────────────────────────────────────────────
d.h1('5. 시트 서식 규칙')
d.table(['보이는 모양', '뜻', '적용 범위'], [
    ['회색 / 흰색 줄무늬', '같은 Thread ID(같은 대화)끼리 같은 색. 대화가 바뀔 때마다 번갈아 칠함', 'A~K열'],
    ['파란 · 굵은 · 기울임 글씨', '받는 사람(C열)에 내 주소(MY_ADDRESS)가 있는 메일', 'A~K열'],
    ['연한 빨강 / 노랑 / 파랑 배경', '긴급 / 요청·확인 / 단순공지 (조건부 서식)', 'G열만'],
], widths=[5, 8.5, 3.5])
d.p('긴급도 색은 줄무늬·내 메일 표시 위에 마지막으로 덧입혀서, G열에서는 긴급도 색이 보이고 나머지 열은 줄무늬가 유지됩니다.')
d.callout('행을 손으로 지우거나 정렬해서 색이 어긋났다면 Refresh Formatting 메뉴를 실행하세요. '
          '데이터는 건드리지 않고 서식만 처음부터 다시 칠합니다.', kind='good', head='색이 흐트러졌을 때')

# ── 6 ────────────────────────────────────────────────────────────────
d.h1('6. Clear Emails Contents — 시트 비우기')
d.p('1행 머리글만 남기고 A~K열의 내용과 서식을 모두 지웁니다. 새 기간으로 처음부터 다시 정리할 때 씁니다.')
d.callout('지운 내용은 되돌릴 수 없습니다 (스프레드시트의 "버전 기록" 으로만 복구). '
          'Gmail 메일은 그대로이지만 이미 보관된 메일은 받은편지함에 없으므로 Get Emails 로 다시 모이지 않습니다.',
          kind='danger', head='되돌릴 수 없음')

# ── 7 ────────────────────────────────────────────────────────────────
d.h1('7. 내 마음대로 바꾸는 설정')
d.p('스프레드시트 메뉴 ', ('확장 프로그램 → Apps Script', B), ' 를 열면 코드 맨 위 "[공통 설정]" 부분에 바꿀 수 있는 값이 모여 있습니다. '
    '고친 뒤에는 꼭 저장(Ctrl+S)하고, 서식 관련 값을 바꿨다면 Refresh Formatting 을 한 번 실행하세요.')
d.table(['변수', '바꾸면 달라지는 것', '바꾼 뒤'], [
    ['URGENCY_RULES', 'F열 긴급도 분류 키워드·이름·G열 배경색', '다음 Get Emails 부터 (+ Refresh)'],
    ['URGENCY_BACKGROUNDS', '규칙에 배경색을 안 적었을 때 쓰는 기본 배경색', 'Refresh Formatting'],
    ['URGENCY_FONTS', 'G열 긴급도별 글자색 · 굵게', 'Refresh Formatting'],
    ['THREAD_BAND_COLOR', '대화별 줄무늬 색', 'Refresh Formatting'],
    ['MY_ADDRESS', '"내 메일" 로 볼 주소', 'Refresh Formatting'],
    ['SUMMARY_PROMPT', 'AI 요약 지시 문구', '다음 Get Emails 부터'],
], widths=[4.5, 7.5, 5], mono_cols=(0,))

d.h2('URGENCY_RULES — 긴급도 분류 규칙')
d.code('var URGENCY_RULES = [',
       '  { pattern: /URGENT|긴급|AOG|결함|CANCELLED|중요/i, label: "🔴 긴급",      background: "#f4cccc" },',
       '  { pattern: /요청|의뢰|문의|지연|Recovery/i,        label: "🟡 요청/확인", background: "#fff2cc" },',
       '  { pattern: /공지|부고|결혼|축하|보고/i,             label: "⚪ 단순공지",  background: "#cfe2f3" }',
       '];',
       'var URGENCY_DEFAULT = "일반";',
       caption='현재 설정')
d.table(['부분', '뜻', '고치는 법'], [
    ['pattern', '찾을 키워드. 제목 + 본문 + 라벨을 합친 글에서 검사', '/ 와 / 사이에 키워드를 | (또는) 로 이어 씀'],
    ['/i', '대소문자 무시. urgent, URGENT, Urgent 모두 잡음', '빼면 대소문자를 구분'],
    ['label', 'F열에 적힐 이름', '자유롭게. 단 URGENCY_BACKGROUNDS · URGENCY_FONTS 의 이름과 맞출 것'],
    ['background', 'G열 배경색', '#rrggbb 색 코드'],
    ['URGENCY_DEFAULT', '어느 규칙에도 안 걸릴 때 이름', '예: "일반"'],
], widths=[3.2, 7, 6.8], mono_cols=(0,))
d.callout('위에서부터 차례로 검사해서 처음 걸린 규칙 하나만 적용됩니다. "긴급 공지" 라는 제목은 긴급 규칙이 위에 있어서 🔴 긴급이 됩니다. '
          '중요한 규칙일수록 위에 두세요.', kind='warn', head='순서가 중요하다')
d.h3('자주 하는 수정 예')
d.table(['하고 싶은 것', '이렇게 고친다'], [
    ['"마감" 도 긴급으로', '/URGENT|긴급|AOG|결함|CANCELLED|중요|마감/i'],
    ['"TOOL 이송", "tool이송" 모두 잡기 (띄어쓰기 상관없이)', '/TOOL\\s*이송/i   (\\s* = 공백 0개 이상)'],
    ['단어 사이에 공백이 꼭 있을 때만', '/TOOL\\s+이송/i   (\\s+ = 공백 1개 이상)'],
    ['새 등급 추가', '{ pattern: /교육|훈련/i, label: "🟢 교육", background: "#d9ead3" }, 을 원하는 순서에 한 줄 추가'],
], widths=[7, 10], mono_cols=(1,))

d.h2('URGENCY_BACKGROUNDS — 기본 배경색 (안전장치)')
d.code('var URGENCY_BACKGROUNDS = { "긴급": "#f4cccc", "요청/확인": "#fff2cc", "단순공지": "#cfe2f3" };')
d.p('URGENCY_RULES 를 고치다가 background 를 빠뜨려도 G열 색이 사라지지 않도록, 라벨에 들어 있는 이름("긴급" 등)으로 '
    '여기서 색을 찾아 씁니다. 규칙에 background 가 있으면 그쪽이 우선입니다. 새 등급을 만들었다면 여기에도 한 줄 추가해 두면 좋습니다.')

d.h2('URGENCY_FONTS — G열 글자색 · 굵게')
d.code('var URGENCY_FONTS = {',
       '  "긴급":      { color: null, bold: true },',
       '  "요청/확인": { color: null, bold: null },',
       '  "단순공지":  { color: null, bold: null }',
       '};')
d.table(['값', 'color (글자색)', 'bold (굵게)'], [
    ['null', '바꾸지 않음 → 내 메일이면 파란 글씨 그대로', '바꾸지 않음 → 원래 굵기 유지'],
    ['"#990000" / true', '그 색으로 바꿈 (내 메일이라도 G열은 이 색)', 'true = 굵게'],
    ['false', '—', '굵게 해제'],
], widths=[3.5, 7.5, 6], mono_cols=(0,))
d.p('예) 긴급 요약을 빨간 굵은 글씨로: ', ('"긴급": { color: "#990000", bold: true }', {'color': 'accent'}))

d.h2('THREAD_BAND_COLOR — 줄무늬 색')
d.code('var THREAD_BAND_COLOR = "#d9d9d9";   // 회색 ↔ 흰색 번갈아')
d.p('같은 대화(Thread ID)끼리 같은 색으로 묶어 보여 줍니다. 더 연하게 하려면 "#efefef", 연한 파랑은 "#e8f0fe" 처럼 바꿉니다.')

d.h2('그 밖의 값')
d.table(['변수', '현재 값', '설명'], [
    ['MY_ADDRESS', 'thlim@koreanair.com', '받는 사람에 이 주소가 있으면 파란 굵은 기울임 글씨'],
    ['SUMMARY_PROMPT', '"는 메일의 … 요약해 줘"', '=AI(K행 & 이 문구) 로 붙는 요약 지시'],
    ['num_of_maxMails_collected', '50', '한 번에 모을 대화 수 (fetchUnreadEmails 함수 안). 늘리면 시간 초과 위험'],
    ['len_of_mailContents', '400', '본문을 몇 글자까지 적을지 (함수 안)'],
    ['num_of_mails_in_thread', '1', '대화 하나에서 몇 통을 적을지 (함수 안, 최근 것부터)'],
], widths=[4.8, 4.7, 7.5], mono_cols=(0, 1))

# ── 8 ────────────────────────────────────────────────────────────────
d.h1('8. 자주 겪는 문제')
d.table(['증상', '원인', '해결'], [
    ['Email Tools 메뉴가 없음', '시트를 연 뒤 스크립트를 저장함', '시트 새로고침'],
    ['"처리할 메일이 없습니다"', '기간 안에 별표 없는 받은편지함 메일이 없음 / 이미 보관됨', '기간 확인, 종료일은 다음 날로'],
    ['G열이 수식 그대로', '[생성 및 채우기] 를 안 누름', '4장 순서대로'],
    ['G열에 #ERROR!', 'AI 생성 실패 또는 사용량 제한', '잠시 뒤 다시 생성, 굳히기는 자동으로 건너뜀'],
    ['실행 중 시간 초과 오류', '메일이 너무 많음', '기간을 나눠 여러 번 실행'],
    ['색이 어긋남', '행을 손으로 지우거나 정렬함', 'Refresh Formatting'],
    ['긴급도가 엉뚱함', '키워드가 다른 규칙에 먼저 걸림', 'URGENCY_RULES 순서·키워드 조정'],
])

d.save(OUT)
print('저장:', OUT)
