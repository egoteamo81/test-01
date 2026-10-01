#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx_kit.py — 워드(.docx) 문서를 '가독성 좋게' 찍어내는 공용 도구

────────────────────────────────────────────────────────────────────────
■ 이 파일이 왜 있나
────────────────────────────────────────────────────────────────────────
가이드북·로드맵 같은 문서를 구글 문서 API 로 만들다 보니 두 가지가 걸렸다.

  1. 구글 문서 API 는 본문 크기 한계가 있다. 16KB 근처에서
     "Resource has been exhausted" 가 나고, 한 번에 다 넣어야 해서
     (update 는 제목·위치만 바꾼다) 글을 줄이는 수밖에 없었다.
  2. 색·음영·표 스타일을 세밀하게 먹이기가 번거롭다.

워드 파일은 **우리 디스크에 그냥 쓰는 파일**이라 두 제약이 다 없다.
크기 한계가 없고, 서식도 마음대로다. 만들어서 사용자가 구글 드라이브에
올리면 구글 문서로 열린다.

────────────────────────────────────────────────────────────────────────
■ 비유로 이해하는 워드 파일의 구조
────────────────────────────────────────────────────────────────────────
.docx 는 사실 **그냥 zip 압축 파일**이다. 확장자를 .zip 으로 바꿔서
풀어보면 XML 문서 몇 개가 들어 있다. 워드는 그 XML 을 읽어서 그린다.

  .docx  ─┬─ word/document.xml   ← 본문. 글자와 서식이 전부 여기
          ├─ word/styles.xml     ← '제목1은 이렇게 생겼다' 같은 서식 정의
          └─ ...

python-docx 라이브러리는 그 XML 을 대신 써 주는 타자기다. 다만 **타자기가
모든 키를 갖고 있지는 않다.** 음영(배경색)이나 동아시아 폰트처럼 자주 쓰는
기능인데 python-docx 에 함수가 없는 것들이 있어서, 그럴 때는 XML 을 직접
꽂아 넣어야 한다. 아래 `_shade()`, `_font()` 같은 함수가 그 역할이다.

────────────────────────────────────────────────────────────────────────
■ 반드시 알아야 할 함정: 한글 폰트는 따로 지정해야 한다
────────────────────────────────────────────────────────────────────────
워드는 글자를 **세 부류로 나눠서** 각각 다른 폰트를 쓴다.

  w:ascii    영문·숫자 (A-Z, 0-9)
  w:hAnsi    서유럽 문자
  w:eastAsia 한중일 문자  ←★ 한글이 여기

그런데 python-docx 의 `run.font.name = '맑은 고딕'` 은 **앞의 둘만** 설정한다.
2026-09-22 실측으로 확인했다. 생성된 XML 이 이랬다.

    <w:rFonts w:ascii="맑은 고딕" w:hAnsi="맑은 고딕"/>      ← eastAsia 없음

이 상태면 영문은 맑은 고딕인데 한글은 워드 기본 동아시아 폰트로 떨어진다.
보는 PC 설정에 따라 엉뚱한 폰트가 되거나 글자 폭이 달라진다. 로그에도
안 나오고 에러도 안 나서, 열어보기 전까지 모른다.

→ 이 파일의 `_font()` 가 `w:eastAsia` 를 **항상 같이** 박는다.
  폰트를 직접 만지려거든 반드시 이 함수를 거칠 것.

────────────────────────────────────────────────────────────────────────
■ 쓰는 법 (가장 짧은 예)
────────────────────────────────────────────────────────────────────────
    from docx_kit import Doc, S

    d = Doc('제목', '부제목')
    d.h1('1장')
    d.p('본문이다.')
    d.callout('이건 꼭 기억할 것', kind='warn')
    d.table(['항목', '값'], [['CPU 비중', '1.3%']])
    d.save('/path/문서.docx')

────────────────────────────────────────────────────────────────────────
■ 이 파일을 고칠 때
────────────────────────────────────────────────────────────────────────
색이나 기호를 바꾸고 싶으면 아래 PALETTE / S 만 고치면 문서 전체에 반영된다.
개별 호출부에 색을 흩뿌리지 말 것. 나중에 통일이 안 된다.
"""

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Cm


# ══════════════════════════════════════════════════════════════════════
#  1. 폰트
# ══════════════════════════════════════════════════════════════════════

# 한국어 윈도우 워드에 기본 설치되어 있는 폰트만 쓴다.
# 없는 폰트를 지정하면 워드가 멋대로 대체하는데, 그때 자간과 줄높이가
# 틀어져서 표가 깨진다. '맑은 고딕' 은 윈도우 비스타 이후 기본 탑재라
# 사실상 어디서나 있다. 맥 워드에도 들어 있다.
FONT_TEXT = '맑은 고딕'

# 코드·경로·수치처럼 세로줄을 맞춰야 하는 것에 쓰는 고정폭 폰트.
# 'D2Coding' 이나 'Consolas' 가 후보인데, D2Coding 은 설치를 따로 해야 해서
# 윈도우 기본인 Consolas 로 간다. Consolas 에는 한글이 없지만, 한글은
# 어차피 w:eastAsia 로 맑은 고딕이 받으므로 섞여도 문제없다.
FONT_MONO = 'Consolas'


# ══════════════════════════════════════════════════════════════════════
#  2. 색
# ══════════════════════════════════════════════════════════════════════
#
# 색을 고를 때 지킨 기준
#   - 배경색은 아주 연하게(명도 90% 이상). 글자가 검은색이어야 읽힌다.
#   - 흑백 인쇄해도 구분되도록 색마다 명도를 다르게 뒀다.
#   - 형광색은 피했다. 화면에서 눈이 아프고 인쇄하면 회색 덩어리가 된다.
#
# 값은 RGB 16진수 문자열(샵 없이). 워드 XML 이 그 형식을 요구한다.

PALETTE = {
    # ── 글자색 ──
    'ink':        '1A1A1A',   # 본문. 순수 검정(000000)보다 눈이 덜 피로하다
    'muted':      '6B7280',   # 부연 설명, 각주
    'navy':       '1E3A5F',   # 제목
    'accent':     '0F5C8C',   # 강조 글자, 링크
    'danger':     'B32020',   # 경고 글자
    'good':       '15803D',   # 확인됨/정상

    # ── 배경(음영)색 ──
    'bg_head':    '1E3A5F',   # 표 머리글 (진한 남색 + 흰 글자)
    'bg_zebra':   'F2F5F8',   # 표 줄무늬. 아주 연해야 글자를 안 가린다
    'bg_note':    'EAF2F8',   # 참고 상자 (연한 파랑)
    'bg_warn':    'FDF1E7',   # 주의 상자 (연한 주황)
    'bg_danger':  'FBEAEA',   # 경고 상자 (연한 빨강)
    'bg_good':    'EAF6EE',   # 확인 상자 (연한 초록)
    'bg_code':    'F4F4F5',   # 코드 블록 (연한 회색)

    # ── 선 ──
    'line':       'C9D2DC',
}


# ══════════════════════════════════════════════════════════════════════
#  3. 기호
# ══════════════════════════════════════════════════════════════════════
#
#  ★ 이모지(✅ ⚠️ 🔴 …)를 쓰지 않는 이유 ★
#
#  이모지는 유니코드 상위 영역이라 '맑은 고딕' 에 글자가 없다. 워드는
#  Segoe UI Emoji 로 대체하는데, PC 에 따라 흑백으로 나오거나 네모(두부)로
#  나온다. 구글 드라이브에 올려 구글 문서로 변환하면 또 달라진다.
#
#  아래 기호들은 전부 KS X 1001(한국 표준 완성형)에 포함된 것들이라
#  한글이 보이는 환경이면 무조건 같이 보인다. 안전하다.

class S:
    """문서에서 쓰는 기호 모음. 직접 타이핑하지 말고 여기를 거칠 것."""
    OK      = '●'    # 확인됨, 완료
    NO      = '○'    # 미확인, 예정
    WARN    = '※'    # 주의. (⚠ 대신 쓴다)
    ARROW   = '→'    # 흐름, 변환
    STEP    = '▶'    # 절차의 각 단계
    STAR    = '★'    # 중요
    DOT     = '·'    # 나열
    BULLET  = '■'    # 소제목 앞
    SUB     = '－'    # 하위 항목 (전각 대시. 하이픈보다 글머리처럼 보인다)
    UP      = '↑'
    DOWN    = '↓'


# ══════════════════════════════════════════════════════════════════════
#  4. python-docx 에 없어서 XML 을 직접 꽂는 부분
# ══════════════════════════════════════════════════════════════════════
#
# 여기부터는 라이브러리가 대신 해 주지 않는 것들이다.
# 워드 XML 을 손으로 조립한다. 건드릴 일은 거의 없다.

def _hyperlink(par, run, url):
    """이미 만든 글자(run)를 외부 링크로 감싼다. 밑줄을 긋고, 워드·구글 문서 어디서 열어도 클릭된다.
    python-docx 에 하이퍼링크 API 가 없어서 w:hyperlink 요소를 직접 만든다."""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    r_id = par.part.relate_to(url, RT.HYPERLINK, is_external=True)
    h = OxmlElement('w:hyperlink')
    h.set(qn('r:id'), r_id)
    run.font.underline = True
    par._p.append(h)
    h.append(run._r)                     # add_run 이 붙인 자리에서 hyperlink 안으로 옮긴다


def _font(run, name=FONT_TEXT, size=None, bold=None, color=None, italic=None):
    """글자 하나(run)에 폰트를 먹인다. **한글 폰트 함정을 막는 유일한 통로.**

    python-docx 의 run.font.name 은 w:eastAsia 를 안 넣어서 한글에 폰트가
    안 걸린다(파일 상단 설명 참고). 그래서 rFonts 를 직접 만들어 세 속성을
    전부 채운다.
    """
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(PALETTE.get(color, color))

    # rPr(run properties)은 '이 글자의 서식표'다. 없으면 만든다.
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), name)      # 영문·숫자
    rFonts.set(qn('w:hAnsi'), name)      # 서유럽
    rFonts.set(qn('w:eastAsia'), name)   # ★ 한글. 이 줄이 핵심이다
    return run


def _shade(element, fill):
    """문단이나 표 칸에 배경색(음영)을 칠한다.

    python-docx 에 음영 API 가 아예 없다. <w:shd> 태그를 pPr(문단 서식표)이나
    tcPr(칸 서식표)에 직접 넣는다.

    val="clear" 는 '무늬 없이 단색으로 채움'이라는 뜻이다.
    """
    pr = element.get_or_add_pPr() if hasattr(element, 'get_or_add_pPr') else element
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), PALETTE.get(fill, fill))
    pr.append(shd)


def _borders(paragraph, color='line', size=6, sides=('left',), space=8):
    """문단에 테두리 선을 긋는다.

    상자(callout)의 왼쪽 굵은 세로줄이 이걸로 만들어진다. 배경색만으로는
    밋밋해서, 왼쪽에 색 막대를 세우면 눈이 한 번에 찾는다.

    size 는 1/8 포인트 단위다. 6이면 0.75pt.
    """
    pPr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    for side in sides:
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), str(size))
        el.set(qn('w:space'), str(space))
        el.set(qn('w:color'), PALETTE.get(color, color))
        pbdr.append(el)
    pPr.append(pbdr)


def _spacing(paragraph, before=None, after=None, line=None):
    """문단 위아래 여백과 줄간격.

    한글은 라틴 문자보다 글자가 빽빽해서 줄간격을 넉넉히 줘야 읽힌다.
    영문 문서 감각으로 1.0 을 주면 답답하다. 1.45 정도가 적당하다.
    """
    pf = paragraph.paragraph_format
    if before is not None:
        pf.space_before = Pt(before)
    if after is not None:
        pf.space_after = Pt(after)
    if line is not None:
        pf.line_spacing = line


def _cell_bg(cell, fill):
    """표 칸 배경색."""
    _shade(cell._tc.get_or_add_tcPr(), fill)


def _no_autofit(table):
    """표 너비를 내용에 따라 멋대로 늘리지 않게 고정한다.

    autofit 을 켜 두면 긴 URL 하나 때문에 칸 하나가 페이지를 다 먹는다.
    """
    table.autofit = False
    tblPr = table._tbl.tblPr
    layout = OxmlElement('w:tblLayout')
    layout.set(qn('w:type'), 'fixed')
    tblPr.append(layout)


# ══════════════════════════════════════════════════════════════════════
#  5. 실제로 쓰는 문서 클래스
# ══════════════════════════════════════════════════════════════════════

class Doc:
    """워드 문서 한 개.

    비유하자면 **문서라는 두루마리에 도장을 순서대로 찍는 것**이다.
    h1() 을 부르면 제목 도장, p() 를 부르면 본문 도장, table() 을 부르면
    표 도장이 그 자리에 찍힌다. 찍은 순서가 곧 문서 순서다.
    마지막에 save() 로 두루마리를 묶는다.
    """

    def __init__(self, title=None, subtitle=None, margin_cm=2.0):
        self.d = Document()
        self._setup_page(margin_cm)
        self._setup_base_style()
        if title:
            self.title(title, subtitle)

    # ── 문서 전체 설정 ──────────────────────────────────────────────

    def _setup_page(self, margin_cm):
        """용지 여백. 기본값(2.54cm)은 한글 문서엔 좀 넓어서 조금 줄인다."""
        for s in self.d.sections:
            s.left_margin = s.right_margin = Cm(margin_cm)
            s.top_margin = s.bottom_margin = Cm(margin_cm)

    def _setup_base_style(self):
        """'Normal' 스타일(모든 문단의 부모)에 기본 폰트를 심는다.

        여기서 한 번 심어 두면 내가 깜빡하고 _font() 를 안 불러도
        한글이 두부가 되지 않는다. 안전망이다.
        """
        st = self.d.styles['Normal']
        st.font.name = FONT_TEXT
        st.font.size = Pt(10.5)
        st.font.color.rgb = RGBColor.from_string(PALETTE['ink'])
        rPr = st.element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        rFonts.set(qn('w:eastAsia'), FONT_TEXT)   # ★ 여기도 잊지 말 것
        st.paragraph_format.line_spacing = 1.45
        st.paragraph_format.space_after = Pt(6)
        self._setup_heading_styles()

    # 제목 스타일: (크기 pt, 색) — 아래 title/h1/h2/h3 의 모양과 똑같이 맞춘다.
    _HEADING_STYLES = {'Title': (22, 'navy'), 'Heading 1': (14.5, 'navy'),
                       'Heading 2': (12, 'accent'), 'Heading 3': (10.8, 'ink')}

    def _setup_heading_styles(self):
        """제목 문단이 쓰는 워드 내장 스타일(Title, Heading 1~3)을 우리 모양으로 다시 정의한다.

        왜 내장 스타일을 쓰나: 글자 서식만 입힌 일반 문단은 워드·구글 문서가 '제목'으로 알아보지
        못한다. Heading 스타일이어야 워드 탐색 창과 구글 문서의 '개요(목차)'에 잡힌다
        (2026-09-28 사용자 요청: 구글 드라이브에서 열 때 목차가 바로 생기도록).

        왜 다시 정의하나: python-docx 기본 템플릿의 Heading 스타일은 테마 글꼴(Calibri Light)과
        테마 색(파랑)을 속성으로 갖고 있다. 워드는 같은 요소 안에서 테마 속성을 일반 값보다
        우선하므로, 지우지 않으면 한글 폰트가 테마 글꼴로 바뀌거나 색이 파랗게 나올 수 있다.
        그래서 테마 속성을 지우고 FONT_TEXT·PALETTE 로 채운다. 제목 문단은 run 에도 _font() 로
        같은 서식을 직접 입히므로 이건 이중 안전장치다."""
        for name, (size, color) in self._HEADING_STYLES.items():
            st = self.d.styles[name]
            st.font.size = Pt(size)
            st.font.bold = True
            st.font.italic = False
            rPr = st.element.get_or_add_rPr()
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.append(rFonts)
            for attr in ('w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme', 'w:cstheme'):
                rFonts.attrib.pop(qn(attr), None)
            for attr in ('w:ascii', 'w:hAnsi', 'w:eastAsia'):
                rFonts.set(qn(attr), FONT_TEXT)
            col = rPr.find(qn('w:color'))
            if col is not None:
                rPr.remove(col)                       # themeColor 가 딸린 기존 색 요소를 통째로 제거
            st.font.color.rgb = RGBColor.from_string(PALETTE[color])

    # ── 제목 ────────────────────────────────────────────────────────

    def title(self, text, subtitle=None):
        """문서 맨 위 큰 제목. 아래에 굵은 가로줄을 긋는다. (워드 'Title' 스타일)"""
        p = self.d.add_paragraph(style='Title')
        _font(p.add_run(text), size=22, bold=True, color='navy')
        _spacing(p, after=2, line=1.2)
        _borders(p, color='navy', size=18, sides=('bottom',), space=6)
        if subtitle:
            sp = self.d.add_paragraph()
            _font(sp.add_run(subtitle), size=10, color='muted')
            _spacing(sp, before=4, after=16)
        return self

    def h1(self, text):
        """장 제목. 배경 음영 + 왼쪽 막대로 확실히 구분한다. (워드 'Heading 1' → 목차 1단계)"""
        p = self.d.add_paragraph(style='Heading 1')
        # 막대와 글자 사이 간격은 공백 문자가 아니라 테두리 여백(space)으로 준다. 공백으로 띄우면
        # 워드 탐색 창·구글 문서 목차에 제목이 '  1. …' 처럼 앞 공백째로 나온다.
        _font(p.add_run(text), size=14.5, bold=True, color='navy')
        _spacing(p, before=20, after=8, line=1.3)
        _shade(p._p, 'bg_note')
        _borders(p, color='navy', size=24, sides=('left',), space=12)
        return self

    def h2(self, text):
        """절 제목. 아래 얇은 선. (워드 'Heading 2' → 목차 2단계)"""
        p = self.d.add_paragraph(style='Heading 2')
        _font(p.add_run(f'{S.BULLET} {text}'), size=12, bold=True, color='accent')
        _spacing(p, before=14, after=5, line=1.3)
        _borders(p, color='line', size=4, sides=('bottom',), space=3)
        return self

    def h3(self, text):
        """소제목. (워드 'Heading 3' → 목차 3단계)"""
        p = self.d.add_paragraph(style='Heading 3')
        _font(p.add_run(text), size=10.8, bold=True, color='ink')
        _spacing(p, before=10, after=3)
        return self

    # ── 본문 ────────────────────────────────────────────────────────

    def p(self, *parts, size=10.5, align=None):
        """본문 문단.

        인자를 여러 개 주면 이어 붙이는데, 튜플로 주면 그 조각만 서식을
        다르게 할 수 있다.

            d.p('보통 글자 ', ('굵고 빨간 글자', {'bold': True, 'color': 'danger'}))
            d.p('자세한 값은 ', ('예시 시트', {'link': 'https://…'}), ' 참고')   # 클릭되는 링크
        """
        par = self.d.add_paragraph()
        for part in parts:
            if isinstance(part, tuple):
                text, opt = part
                opt = dict(opt)                       # 호출한 쪽 dict 를 건드리지 않게
                url = opt.pop('link', None)
                run = par.add_run(text)
                _font(run, size=opt.pop('size', size), **({'color': 'accent'} if url else {}) | opt)
                if url:
                    _hyperlink(par, run, url)
            else:
                _font(par.add_run(str(part)), size=size)
        if align == 'center':
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return self

    def bullet(self, text, level=0, mark=None):
        """글머리 기호 문단.

        워드 기본 목록 스타일을 안 쓰고 직접 기호를 찍는 이유:
        기본 목록은 구글 문서로 변환할 때 들여쓰기가 자주 깨진다.
        그냥 기호 + 들여쓰기가 어디서 열어도 똑같이 나온다.
        """
        mark = mark or (S.DOT if level == 0 else S.SUB)
        par = self.d.add_paragraph()
        par.paragraph_format.left_indent = Cm(0.5 + level * 0.6)
        par.paragraph_format.first_line_indent = Cm(-0.42)
        _font(par.add_run(f'{mark} '), size=10.5, color='accent' if level == 0 else 'muted')
        _font(par.add_run(text), size=10.5)
        _spacing(par, after=3)
        return self

    def steps(self, items):
        """번호 붙은 절차. 번호를 동그라미 숫자로 찍는다."""
        circled = '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮'
        for i, t in enumerate(items):
            par = self.d.add_paragraph()
            par.paragraph_format.left_indent = Cm(0.5)
            par.paragraph_format.first_line_indent = Cm(-0.5)
            mark = circled[i] if i < len(circled) else f'{i + 1}.'
            _font(par.add_run(f'{mark} '), size=11, bold=True, color='accent')
            _font(par.add_run(t), size=10.5)
            _spacing(par, after=3)
        return self

    # ── 강조 상자 ───────────────────────────────────────────────────

    _KIND = {
        'note':   ('bg_note',   'accent', S.STAR),
        'warn':   ('bg_warn',   'danger', S.WARN),
        'danger': ('bg_danger', 'danger', S.WARN),
        'good':   ('bg_good',   'good',   S.OK),
    }

    def callout(self, *lines, kind='note', head=None):
        """눈에 띄는 상자. 배경 음영 + 왼쪽 색 막대.

        kind
          note   파랑  — 알아두면 좋은 것
          warn   주황  — 실수하기 쉬운 것
          danger 빨강  — 하면 안 되는 것
          good   초록  — 확인됐다 / 이렇게 하면 된다
        """
        bg, col, mark = self._KIND[kind]
        for i, line in enumerate(lines):
            par = self.d.add_paragraph()
            par.paragraph_format.left_indent = Cm(0.25)
            par.paragraph_format.right_indent = Cm(0.1)
            if i == 0:
                _font(par.add_run(f'{mark} '), size=10.5, bold=True, color=col)
                if head:
                    _font(par.add_run(f'{head}  '), size=10.5, bold=True, color=col)
            else:
                _font(par.add_run('   '), size=10.5)
            _font(par.add_run(line), size=10.3)
            _shade(par._p, bg)
            _spacing(par,
                     before=8 if i == 0 else 0,
                     after=8 if i == len(lines) - 1 else 0,
                     line=1.4)
            _borders(par, color=col, size=20, sides=('left',), space=6)
        return self

    def code(self, *lines, caption=None):
        """명령어나 코드. 고정폭 폰트 + 회색 배경.

        세로줄이 맞아야 읽히는 것들(경로, 명령, 수치 정렬)에만 쓴다.
        """
        if caption:
            cp = self.d.add_paragraph()
            _font(cp.add_run(caption), size=9, bold=True, color='muted')
            _spacing(cp, before=8, after=1)
        for i, line in enumerate(lines):
            par = self.d.add_paragraph()
            par.paragraph_format.left_indent = Cm(0.3)
            _font(par.add_run(line or ' '), name=FONT_MONO, size=9.5)
            _shade(par._p, 'bg_code')
            _spacing(par,
                     before=6 if i == 0 and not caption else 0,
                     after=6 if i == len(lines) - 1 else 0,
                     line=1.25)
            _borders(par, color='line', size=12, sides=('left',), space=6)
        return self

    # ── 표 ──────────────────────────────────────────────────────────

    def table(self, header, rows, widths=None, mono_cols=()):
        """표.

        header     머리글 리스트
        rows       각 행. 리스트의 리스트
        widths     각 열 너비(cm). 안 주면 균등 분배
        mono_cols  고정폭 폰트로 찍을 열 번호들 (수치나 경로 열)

        줄무늬(zebra)를 넣는 이유: 열이 4개를 넘어가면 눈이 행을 놓친다.
        연한 배경을 한 줄 걸러 깔면 시선이 가로로 안 샌다.
        """
        t = self.d.add_table(rows=1, cols=len(header))
        t.style = 'Table Grid'
        t.alignment = WD_TABLE_ALIGNMENT.LEFT
        _no_autofit(t)

        if widths:
            for r in t.rows:
                for c, w in zip(r.cells, widths):
                    c.width = Cm(w)

        # 머리글: 진한 남색 배경 + 흰 글자
        for cell, text in zip(t.rows[0].cells, header):
            cell.text = ''
            par = cell.paragraphs[0]
            _font(par.add_run(str(text)), size=10, bold=True, color='FFFFFF')
            _spacing(par, before=3, after=3, line=1.25)
            _cell_bg(cell, 'bg_head')

        # 본문 행
        for ri, row in enumerate(rows):
            cells = t.add_row().cells
            if widths:
                for c, w in zip(cells, widths):
                    c.width = Cm(w)
            for ci, (cell, val) in enumerate(zip(cells, row)):
                cell.text = ''
                par = cell.paragraphs[0]
                _font(par.add_run(str(val)),
                      name=FONT_MONO if ci in mono_cols else FONT_TEXT,
                      size=9.8)
                _spacing(par, before=2, after=2, line=1.3)
                if ri % 2 == 1:                 # 한 줄 걸러 줄무늬
                    _cell_bg(cell, 'bg_zebra')
        self.space(6)
        return self

    # ── 그림 ────────────────────────────────────────────────────────

    def image(self, path, width_cm=16.0, caption=None):
        """그림 한 장을 가운데 정렬로 넣는다. 캡션은 아래에 작은 회색 글자로.

        폭 16cm 는 A4 에서 좌우 여백 2cm 를 뺀 본문 폭(17cm)보다 살짝 작은 값이다.
        꽉 채우면 워드가 다음 쪽으로 밀어내면서 앞 쪽에 큰 빈칸이 생기기 쉽다.
        """
        par = self.d.add_paragraph()
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.add_run().add_picture(path, width=Cm(width_cm))
        _spacing(par, before=8, after=2 if caption else 10)
        if caption:
            cp = self.d.add_paragraph()
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _font(cp.add_run(caption), size=9, color='muted')
            _spacing(cp, before=0, after=12)
        return self

    # ── 기타 ────────────────────────────────────────────────────────

    def space(self, pt=10):
        """빈 줄. 문단 사이를 띄울 때."""
        p = self.d.add_paragraph()
        _font(p.add_run(''), size=1)
        _spacing(p, before=0, after=pt, line=1.0)
        return self

    def rule(self):
        """가로 구분선."""
        p = self.d.add_paragraph()
        _font(p.add_run(''), size=1)
        _spacing(p, before=6, after=10, line=1.0)
        _borders(p, color='line', size=6, sides=('bottom',), space=1)
        return self

    def page_break(self):
        from docx.enum.text import WD_BREAK
        self.d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        return self

    def save(self, path):
        self.d.save(path)
        # 프로젝트의 docs/ 아래에 저장한 문서는 구글 드라이브의 최신본도 갈아 끼운다.
        # 규칙과 이유는 drive_publish.py 머리말. 실패해도 저장은 끝난 뒤라 예외를 안 낸다.
        # DRIVE_PUBLISH=0 이면 끈다. (2026-10-01)
        try:
            import os
            import importlib.util as _u
            _spec = _u.spec_from_file_location(
                'drive_publish', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'drive_publish.py'))
            _m = _u.module_from_spec(_spec)
            _spec.loader.exec_module(_m)
            _m.maybe_publish(path)
        except Exception as e:                       # noqa: BLE001
            print(f'[drive] ※ 업로드 단계 오류(저장은 됨): {e}')
        return path


# ══════════════════════════════════════════════════════════════════════
#  6. 자기 검사
# ══════════════════════════════════════════════════════════════════════
#
#  이 파일을 직접 실행하면 견본 문서를 만들고, 한글 폰트가 제대로
#  박혔는지 XML 을 뒤져서 확인한다.
#
#      /root/myenv/bin/python3 docx_kit.py /tmp/견본.docx

if __name__ == '__main__':
    import sys, zipfile, re

    out = sys.argv[1] if len(sys.argv) > 1 else '/tmp/docx_kit_견본.docx'

    d = Doc('docx_kit 견본', '색·음영·기호가 제대로 나오는지 확인하는 문서')
    d.h1('1. 글자와 상자')
    d.p('본문은 맑은 고딕 10.5pt, 줄간격 1.45 다. ',
        ('중요한 곳은 이렇게 강조한다.', {'bold': True, 'color': 'danger'}))
    d.callout('알아두면 좋은 것은 파란 상자에 넣는다.', kind='note', head='참고')
    d.callout('실수하기 쉬운 것은 주황 상자다.', kind='warn', head='주의')
    d.callout('하면 안 되는 것은 빨간 상자.', kind='danger', head='금지')
    d.callout('확인된 사실은 초록 상자.', kind='good', head='확인')
    d.h2('기호 점검')
    d.p(' '.join([S.OK, S.NO, S.WARN, S.ARROW, S.STEP, S.STAR,
                  S.DOT, S.BULLET, S.SUB, S.UP, S.DOWN]))
    d.steps(['첫째 단계', '둘째 단계', '셋째 단계'])
    d.h1('2. 표와 코드')
    d.table(['항목', '값', '비고'],
            [['CPU 비중', '1.3%', '나머지는 API 대기'],
             ['크롬 콜드 스타트', '120초', '재사용하면 5초'],
             ['GrabCut 1000x750', '57초', '쓰지 않는다']],
            widths=[5, 3, 8], mono_cols=[1])
    d.code('cd /root/my_project/scripts/onlineSeller/DetailPageCreator',
           './sc_extract.sh _sample_01 --merge-only', caption='명령 예시')
    d.save(out)

    # ── 검산: 한글 폰트가 실제로 박혔는가 ──
    xml = zipfile.ZipFile(out).read('word/document.xml').decode()
    fonts = re.findall(r'<w:rFonts[^/]*/>', xml)
    missing = [f for f in fonts if 'eastAsia' not in f]
    print(f'저장: {out}')
    print(f'rFonts 태그 {len(fonts)}개 중 eastAsia 누락 {len(missing)}개')
    if missing:
        print('  ' + S.WARN + ' 누락 예시:', missing[:3])
        sys.exit(1)
    print(f'음영(w:shd) {xml.count("<w:shd")}개, 테두리(w:pBdr) {xml.count("<w:pBdr")}개')
    heads = {v: xml.count(f'<w:pStyle w:val="{v}"/>') for v in ('Title', 'Heading1', 'Heading2', 'Heading3')}
    print(f'제목 스타일 문단: {heads}')
    if not all(heads[k] for k in ('Title', 'Heading1', 'Heading2')):
        print('  ' + S.WARN + ' 제목 스타일이 적용되지 않은 제목이 있음')
        sys.exit(1)
    styles = zipfile.ZipFile(out).read('word/styles.xml').decode()
    if 'asciiTheme' in styles.split('w:styleId="Heading1"')[1].split('</w:style>')[0]:
        print('  ' + S.WARN + ' Heading 1 스타일에 테마 글꼴이 남아 있음')
        sys.exit(1)
    print(S.OK + ' 통과')
