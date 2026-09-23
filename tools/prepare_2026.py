# -*- coding: utf-8 -*-
"""2026 제주교육마루 — 원고(docx/hwpx) → build.py용 마크다운 소스 트리 생성.

  python prepare_2026.py
  → C:\\Users\\user\\newsletter-jeju\\_src2026\\
"""
import argparse
import os
import re
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from docx2md import convert as docx_convert     # noqa: E402
from hwpx2md import convert as hwpx_convert     # noqa: E402
import refstyle                                 # noqa: E402

# 원고 원본이 모여 있는 폴더 — 구글 드라이브의 '2026 제주지회 뉴스레터'.
# 기기마다 드라이브 마운트 위치가 달라 차례로 찾는다. --src 나 NEWSLETTER_SRC
# 환경변수로 덮어쓸 수 있다.
DRIVE_FOLDER = "2026 제주지회 뉴스레터"


def _find_drive():
    env = os.environ.get("NEWSLETTER_SRC")
    if env:
        return Path(env)
    home = Path.home()
    cands = [
        Path("G:/내 드라이브") / DRIVE_FOLDER,                    # Windows
        home / "Library/CloudStorage" / "GoogleDrive-yangbyul2@gmail.com"
             / "내 드라이브" / DRIVE_FOLDER,                       # macOS
        home / "Google Drive/내 드라이브" / DRIVE_FOLDER,
        home / "Downloads",
    ]
    for c in cands:
        if c.exists():
            return c
    return home / "Downloads"


DL = _find_drive()
REPO = Path(__file__).resolve().parent.parent
OUT = REPO / "_src2026"
SUB = OUT / "한국교육학회 제주지회 뉴스레터"

# --------------------------------------------------------------------------
# 원고 목록
#   title_override: 원고에 제목이 비어 있거나 템플릿 문구인 경우 편집자가 지정.
#                   ★ 표시는 PI 확인 필요.
# --------------------------------------------------------------------------
DOCX = [
    # 2026-09-22 PI: 대학원생 필자 넷의 소속 표기를 하나로 맞춘다 —
    #   「제주대학교 대학원 교육학과 박사과정」. 원고에는 '제주대학교 / 교육학과',
    #   '제주대학교 대학원 교육학과', '제주대학교 교육학과', '제주대학교'로 네 가지가
    #   섞여 있었다. 카드 목록(config_2026.py의 card_author)도 같이 맞췄다.
    dict(file="뉴스레터 원고_아라캠퍼스_홍지환.docx", slug="제주교육소식_홍지환",
         title_override="성읍마을에서 만난 작은 주인공들",   # ★ 임시 제목
         needs_title=True,
         affil_override="제주대학교 대학원 교육학과", role_override="박사과정"),
    dict(file="뉴스레터 원고_김지원.docx", slug="제주교육소식_김지원",
         affil_override="제주대학교 대학원 교육학과", role_override="박사과정"),
    dict(file="뉴스레터 원고 (전새미).docx", slug="제주교육소식_전새미",
         title_from_body=True,
         affil_override="제주대학교 대학원 교육학과", role_override="박사과정"),
    dict(file="뉴스레터 원고_홍지오.docx", slug="제주교육소식_홍지오"),
    # 2026-09-22 PI 결정: 자격(1급 언어재활사)은 **글 안의 필자 카드에만** 싣고,
    #   첫 화면 목록에는 다른 대학원생과 같이 '박사과정'으로만 보인다.
    #   바깥 표기는 config_2026.py 의 card_author 가 맡는다 — 그쪽이 우선한다.
    #   (글이 화용언어를 다루므로 자격이 본문에서는 글의 무게를 받쳐 준다.)
    dict(file="석진아_뉴스레터 원고_템플릿.docx", slug="제주교육소식_석진아",
         affil_override="제주대학교 대학원 교육학과",
         role_override="박사과정 · 1급 언어재활사"),
    # 사진 아래 캡션이 맨 문장으로만 적혀 있어 본문과 구분되지 않았다(PI 2026-09-22).
    # [사진설명= …] 꼴로 맞추면 다른 원고와 같은 캡션 꼴로 그려진다.
    dict(file="조천_마을탐방_뉴스레터_원고_양유정 (4).docx", slug="활동소개_양유정",
         title_from_body=True,
         captions=[
             "너븐숭이에서 고영철 선생님의 설명을 듣는 탐방 참가자들",
             "조천만세운동 기념탑 앞에서 항일운동의 역사를 되새기는 탐방 참가자들",
             "조천의 역사 현장에서 함께한 한국교육학회 제주지회 회원들",
         ]),
    # 이 글은 사무국장으로서 쓴 행사 기록이다. 같은 절의 양유정(제주지회 총무)과
    # 같은 꼴로 맞춘다 — 원고의 '제주영송학교 / 교사'는 싣지 않는다.
    # (PI 2026-09-21, 병기했다가 되돌림)
    dict(file="뉴스레터 원고_템플릿 _창립기념행사_황현철.docx", slug="활동소개_황현철",
         affil_override="한국교육학회 제주지회",
         role_override="사무국장"),
    dict(file="뉴스레터 원고_강한호.docx", slug="제주교육소식_강한호",
         photo_crop=(35, 0, 145, 110),    # 세로 사진 — 얼굴이 원 중앙에 오게
         # 원고에는 '교사'로만 적혀 있으나 이번 호 회원 동정에 박사학위 취득이
         # 함께 실리므로 직함에 박사를 덧붙였다. (PI 확인 2026-09-11)
         role_override="교사 · 교육학박사"),
    # 2026-09-22 수합. 필자가 html-to-docx 로 만들어 보내 제목·소제목이 모두
    # 굵은 글씨 한 줄로만 왔다 — 제목은 본문 첫 줄에서 올리고(title_from_body),
    # 소제목 다섯은 bold_headings 로 다른 원고와 같은 소제목이 되게 한다.
    # 이 서식에는 필자 표(필자/소속/직함)가 없어 이름·소속·직함을 지정한다.
    # 직함은 이번 호 회원 소식과 맞춰 '부교수'로 적는다(2026.3.1. 승진).
    # 소속은 학부까지 밝힌다 — 시론 두 편을 「제주대학교 초등교육학부 + 직함」으로
    # 맞췄다(PI 2026-09-22). 김민호 시론도 같다.
    # ★ 사진은 이번 호 회원 소식에 쓴 제주대신문 기사 사진을 같이 쓴다 —
    #   한 호에 같은 사진이 두 번 나온다. 본인 사진을 받으면 교체.
    dict(file="칼럼_통합교육과_신경다양성_시론_연준모(260922).docx",
         slug="시론_연준모",
         title_from_body=True,
         bold_headings=True,
         name_override="연준모",
         affil_override="제주대학교 초등교육학부",
         role_override="부교수",
         photo="연준모 교수.jpg", photo_crop=(0, 0, 150, 150)),
]

HWPX = [
    dict(file="뉴스레터 2026 주론-0906.hwpx", slug="주론_이인회",
         title="지나온 60년, 새롭게 열어갈 제주 교육",
         name="이인회", affil="한국교육학회 제주지회", role="회장",
         photo="이인회 교수.jpg", photo_crop=(5, 0, 195, 190),
         drop_first=2),
    dict(file="한국교육학회 제주지회 창립 60주년 축사-교육감(0825).hwpx",
         slug="인사말_고의숙",
         title="한국교육학회 제주지회 창립 60주년을 향한 축하말씀",
         name="고의숙", affil="제주특별자치도교육청", role="교육감",
         # webp 원본이 해상도가 높아 그쪽을 쓴다
         photo="고의숙 교육감.webp", photo_crop=(8, 0, 988, 980),
         drop_first=1),
    # 2026-09-21 수합. 원고에 필자 사진은 오지 않았다.
    dict(file="김민호_몰로카이 주민의 선택.hwpx", slug="시론_김민호",
         title="하와이 몰로카이 주민의 선택과 ‘지속가능한 생활양식’ 교육",
         name="김민호", affil="제주대학교 초등교육학부", role="명예교수",
         drop_first=2,
         sub_headings=True,     # 가./나./다. 와 1)/2)/3) 을 소제목으로 올린다
         # 이번 원고에는 필자 사진이 오지 않았다. 2025년 재창간호의 필자 글
         # 「강단을 떠나, 교육을 품다」 저자 카드에 실렸던 사진을 다시 쓴다 —
         # 필자가 직접 보내 이미 지면에 실린 사진이다. ★ 재사용 확인 권장,
         # 최근 사진을 받으면 교체.
         photo_path="2025/images/99ec5b63_image.png",
         photo_crop=(0, 0, 250, 250),
         # 원고의 절 번호가 '시작하며 / 2 / 3 / 3 / 4. 나오며'로 3이 두 번이다.
         # 1~5로 이어지도록 고쳤다. ★ 필자 확인 필요.
         heading_fix=[
             ("4. 나오며", "5. 나오며"),
             ("3. 지역사회 기반 몰로카이의 교육 프로그램",
              "4. 지역사회 기반 몰로카이의 교육 프로그램"),
             ("시작하며", "1. 시작하며"),
         ],
         # 본문 사진 2장이 빈 표 한 칸에 캡션만 담겨 있었다. 표를 걷어내고
         # 사진과 캡션을 이 호의 사진 표기 방식으로 다시 놓는다.
         # 캡션 문구는 필자가 쓴 것을 그대로 옮겼다(영문 포함). ★ 필자 확인 필요.
         table_replace="""![image](images/김민호_01.jpg)

[사진설명= Meli Watanuki, 사진 출처=Kevin Fujii/Civil Beat(2025), Civil Beat(2026)에서 재인용]

![image](images/김민호_02.jpg)

[사진설명= Maui County Council member Keani Rawlins-Fernandez holds the gate while a National Park Service ranger tries to open it Thursday, 사진 출처=Courtesy: Walter Ritte(2026), Civil Beat(2026)에서 재인용]"""),
]

# 신간 안내에 싣는 책 표지. 인물 사진과 달리 정사각으로 자르지 않고 그대로 옮긴다.
# 출처는 알라딘 상품 이미지(2026-09-21 내려받음). ★ 출판사 제공본을 받으면 교체.
BOOK_COVERS = [
    "책표지_태극도와서명도.jpg",
    "책표지_놀이와유아교육.jpg",
    "책표지_오늘의교육내일의교육정책.jpg",
    # 2026-09-22 PI 전달. 학지사 신간 안내 리플릿(2026-09-16) 표지면을 잘라 썼다.
    # 앞의 세 권과 달리 판매처 상품 이미지가 아니라 출판사 인쇄용 원본이다.
    "책표지_학습장애및학습지원대상학생교육.jpg",
]

# 회원 소식에 넣을 인물 사진 — 얼굴이 원 중앙에 오도록 자른다.
NEWS_PHOTOS = [
    ("강동호", "강동호 교수.jpg", (18, 0, 132, 114)),
    # 머리가 원 아래쪽에 치우쳐 있어 크롭 창을 내렸다. 눈높이가 위에서 약 36%로
    # 오는 자리다. (PI 2026-09-21)
    ("김대영", "김대영 교수.jpg", (23, 48, 153, 178)),
    # 제주대신문 「마음을 다루는 심리학의 매력」(슬기로운 교수생활, 2024. 3.)의
    # 본문 사진. 원본이 150×183으로 작다 — 저자 카드가 104px 원형이라 겨우 맞는다.
    # ★ 본인에게 받은 사진이 생기면 교체. (PI 지시 2026-09-21)
    ("연준모", "연준모 교수.jpg", (0, 0, 150, 150)),
]

# --------------------------------------------------------------------------
# PI 확인을 거친 참고문헌 확정본 — 자동 정리를 거치지 않고 그대로 쓴다.
# (주론은 원본 hwpx의 필드가 깨져 학술지 정보가 겹쳐 있었다. 2026-09-11 PI 확인)
# --------------------------------------------------------------------------
REF_OVERRIDE = {
    "주론_이인회": """
김봉종(1993). 제주지회의 성장. 한국교육학회 편. 교육탐구의 세월: 한국교육학회 40년사(pp. 413-416). 서울: 교육과학사.
김일방(2021). ‘제주이해교육’의 실태 분석 및 미래 발전방안. 교육과학연구, 23(2), 39-63.
송성대, 김정숙, 진관훈, 강만익, 정지훈(2023). 제주문화의 원류 해민정신. 제주: 각.
양진건(1991). 제주교육행정사. 제주: 경신인쇄사.
유홍준(2012). 나의 문화유산답사기: 돌하르방 어디 감수광. 경기: 창비.
이인회(2023). 제주교육학 정립의 필요성에 대한 탐색적 연구: 학교 현장 전문가 심층면담을 중심으로. 교육과학연구, 25(4), 1-26.
제주교육학연구회(2022). 지역교육학으로서 제주교육학은 가능한가? 2022 제주교육학 3차 포럼 자료집, 57.
한국교육학회(1973). 한국교육학회 20년사: 1953.4.4.~1973.4.4. 서울: 대광인쇄공사.
현용준(1986). 제주도 무속연구. 서울: 집문당.
""",
}

# --------------------------------------------------------------------------
# 오탈자 교정 — 명백한 오기만 고친다.
#   '고칠 것이 확실한 것'의 기준: 같은 원고 안에 바른 표기가 이미 있거나
#   (화와이/하와이, 오하우/오아후, 수늘음/수눌음), 고유명사의 철자가 정해져
#   있는 것(Molokai, Smithsonian, Television). 판단이 필요한 것은 고치지 않고
#   PI·필자 확인 목록으로 넘긴다.
# --------------------------------------------------------------------------
TYPO_FIX = {
    "시론_김민호": [
        # 국문
        ("화와이", "하와이"),
        ("오하우", "오아후"),
        ("인프라카", "인프라가"),
        ("정잭", "정책"),
        ("호화시러운", "호화스러운"),
        ("반대할 수 밌는", "반대할 수 있는"),
        ("풀뿔리", "풀뿌리"),
        ("통합 자치 제제", "통합 자치 체제"),
        ("수늘음", "수눌음"),
        # 영문 고유명사·철자
        ("Nwews", "News"),
        ("Molikai", "Molokai"),
        ("Molokia", "Molokai"),
        ("Foiundation", "Foundation"),
        ("Televion", "Television"),
        ("Smithonian", "Smithsonian"),
        ("respod", "respond"),
        ("immesion", "immersion"),
        ("comming soon", "coming soon"),
        ("os the least developed", "is the least developed"),
        # 본문 인용 표기
        ("Kilohana elementary school,2026", "Kilohana elementary school, 2026"),
        ("Hawaii Public Radio. 2020", "Hawaii Public Radio, 2020"),
        # 참고문헌: 주소의 이음줄 두 개가 붙임표(—)로 자동 변환돼 링크가 죽어
        # 있었다. 실제 주소는 하이픈 두 개다 (2026-09-21 접속 확인, 200).
        ("culture—699256ina-based", "culture--699256ina-based"),
        # 참고문헌: 두 항목이 줄바꿈 없이 붙어 있었다
        ("2026.9.14.인출Pacific American",
         "2026.9.14. 인출\nPacific American"),
        # 참고문헌: 괄호 연도가 없어 항목으로 인식되지 않던 것
        ("Maui District Television. Facebook July 29, 2026.",
         "Maui District Television(2026). Facebook, July 29, 2026."),
    ],
}

# 빼기로 한 항목 — 항목에 이 문구가 들어 있으면 제외한다.
# (전새미: 7월 6일 기사와 같은 사안인데 출처 URL이 비어 있어 중복으로 판단. PI 확인)
REF_DROP = {
    "제주교육소식_전새미": ["제주대-조지아공대 글로벌 런케이션 캡스톤 프로그램 성료"],
}

# --------------------------------------------------------------------------
# 원고가 아닌 고정 지면 — 지회가 직접 채우는 페이지.
# 구성원 소개는 2025년 재창간호에서 그대로 가져왔다(PI 지시 2026-09-11).
# ★ 구성원 수·기준 시점은 2026년 기준으로 갱신이 필요하다.
# --------------------------------------------------------------------------
STATIC = [
    # 회원가입·회비 박스는 회비 납부자 명단으로 옮겼다(PI 2026-09-21).
    # 구성원 수 줄은 '임원진'과 같은 크기(h2)로 둔다. ★ 인원은 PI가 파악해 갱신.
    # 2026-09-22 PI: **임원진과 위원회를 나눠 적는다** (안암교육학회 뉴스레터
    #   창간호 14쪽 '조직 및 연락처'를 참고). 전에는 뉴스레터위원회 위원장만
    #   임원진 목록에 섞여 있었고 나머지 위원회는 지면에 없었다.
    #   상설/신설은 구분하지 않는다 — 독자에게 필요한 것은 '어떤 위원회가 있고
    #   누가 이끄는가'이지 언제 생겼는가가 아니다.
    #   근거: 활동 소개(황현철) 본문 — 2026.1.23. 임시총회에서 3개 소위원회
    #   (교육공동체·상담교육·영유아교육)를 신설해 기존 네트워크위원회·사무국과
    #   함께 가기로 결정. 주론(이인회)에도 '교육공동체위원회(위원장: 강은주 교장)'.
    #   ★ 강은주 교장의 소속 학교명은 어느 원고에도 없어 비워 두었다.
    # 2026-09-22 PI: 묶음(임원진/위원회/사무국) 사이에 가로줄(---)을 넣는다.
    #   h2(묶음)와 h3(직책)의 글자 크기 차이가 0.13rem뿐이고 위 여백도 같아
    #   영역이 갈리지 않았다. CSS로 h2를 키우면 원고 41개 소제목이 전부 바뀌므로,
    #   이 지면의 내용 쪽에 구분선을 둔다 — 다른 지면에 영향이 없다.
    # 2026-09-22 PI: 회장 칸도 이름 한 줄만 둔다 — 임기·전화·이메일·지회 주소를
    #   뺐다. 다른 임원 칸이 모두 한 줄이라 회장만 다섯 줄이면 목록이 고르지 않고,
    #   한 줄일 때 오히려 눈에 잘 들어온다.
    #   임기 삭제는 9/21 '회원 소식에서 임기·임용기간을 뺀다' 원칙과도 맞는다.
    #   지회 주소는 바닥글 판권의 '주소'와 같은 값이라 중복이었다.
    #   ★ 다만 회장 전화(064-754-2163)·이메일(tomlee@jejunu.ac.kr)은 지면에서
    #     이 자리에만 있었다 — 빼면 지회장 직통 연락처가 사라진다. 바닥글에는
    #     뉴스레터 문의(양은별)와 후원 계좌만 남는다. PI 확인 요망.
    # 2026-09-22 PI: 사무국(사무국장·총무)도 임원진에서 빼내 위원회 **아래**
    #   독립 묶음으로 둔다 (안암교육학회 뉴스레터와 같은 짜임).
    #   → 임원진(회장·부회장·상임이사·감사) / 위원회 / 사무국 세 묶음이 된다.
    # 2026-09-23 임원 회람 피드백 + PI 결정: **수석부회장을 회장 아래 별도 칸으로**
    #   둔다. 연준모 교수를 부회장 칸 맨 위에 두고 괄호로 「수석부회장」을 병기하는
    #   안을 먼저 냈으나, 그러면 다섯 줄 중 한 줄만 괄호가 길어져 목록이 고르지
    #   않고 직책이 소속 안에 묻힌다. 회장·감사처럼 직책 하나에 칸 하나가
    #   이 지면의 짜임이다. 부회장 칸의 나머지 네 분 차례는 그대로.
    #   상임이사 칸의 연준모 교수는 겸직이므로 그대로 둔다.
    dict(slug="활동소개_구성원소개", title="구성원 소개", body="""
## 제주지회 구성원 : 70명 (2025년 2월 기준)

## 임원진

### 회장

<aside>
이인회 교수 (제주대학교 아라캠퍼스)
</aside>

### 수석부회장

<aside>
연준모 교수 (제주대학교 사라캠퍼스)
</aside>

### 부회장

<aside>
김영희 교수 (제주관광대학교)
김은정 교수 (제주국제대학교)
최보영 교수 (제주대학교 아라캠퍼스)
박한샘 교수 (제주한라대학교)
</aside>

### 상임이사

<aside>
연준모 교수 (제주대학교 사라캠퍼스)
</aside>

### 감    사

<aside>
하진의 교수 (제주국제대학교)
</aside>

---

## 위원회

### 뉴스레터위원회

<aside>
위원장 양은별 교수 (제주대학교 사라캠퍼스)
</aside>

### 네트워크위원회

<aside>
위원장 연준모 교수 (제주대학교 사라캠퍼스)
</aside>

### 교육공동체위원회

<aside>
위원장 강은주 교장
</aside>

### 상담교육위원회

<aside>
위원장 박한샘 교수 (제주한라대학교)
</aside>

### 영유아교육위원회

<aside>
위원장 김은정 교수 (제주국제대학교)
</aside>

---

## 사무국

### 사무국장

<aside>
황현철 (제주대학교)
</aside>

### 총무

<aside>
양유정 (제주대학교)
</aside>
"""),

    # 박사학위 취득 — 제주대학교 대학원 교육학과, 2026년 학위수여자.
    # 출처: 제주대학교 dCollection(학과별 학위논문, 대학원 교육학과) 2026년 4건.
    # 요약은 공개된 초록을 바탕으로 2~3문장으로 쓴다 (PI 2026-09-21).
    # 처음에는 초록을 길게 옮겼는데 초록이 없는 항목과 길이 차이가 너무 컸다.
    # 초록이 공개되지 않은 2건(강윤심·양순미)은 작년 호와 같이 제목·지도교수·
    # 학위수여 시기만 싣는다(내용 지어내지 않음). 2026-09-21 재확인: 양순미
    # 논문은 RISS에 아직 등재되지 않았고 제주대 리포지터리는 접속되지 않는다.
    # 항목 하나가 카드(aside) 하나 — 2025년 호의 신간 안내와 같은 짜임.
    # 요약이 있는 것과 없는 것이 섞여도 카드 경계로 구분된다.
    dict(slug="회원동정_박사학위", title="박사학위 취득", body="""
2026년 제주대학교 대학원 교육학과에서 박사학위를 취득하신 분들입니다. 축하드립니다.

<aside>
강윤심 박사
지도교수 최보영
2026. 2.
---
학교 밖 청소년의 회복탄력성 증진을 위한 현실치료 기반 집단상담 프로그램 개발 및 효과
(Development and Effectiveness of a Reality Therapy-Based Group Counseling Program to Enhance Resilience in Out-of-School Youth)
</aside>

<aside>
강한호 박사
지도교수 이인회
2026. 8.
---
지역사회 기반 대안학교 설립·운영 사례연구 : ‘광명YMCA 볍씨학교’를 중심으로
(A Case Study of the Establishment and Operation of a Community-Based Alternative School: Focusing on Gwangmyeong YMCA Byeopssi School)
지역사회 기반 대안학교인 광명YMCA 볍씨학교가 설립된 때부터 현재에 이르기까지의 전개 과정을 사례연구로 추적한다. 학교와 지역사회의 관계가 자원을 주고받는 차원을 넘어, 공동의 교육적 가치를 중심으로 서로의 역할과 의미를 다시 세우는 ‘공존적 관계’로 형성되어 왔음을 밝힌다. 지역사회 기반 대안학교의 지속가능성은 변화하는 조건 속에서 관계를 다시 짜는 역량에 달려 있다고 본다.
</aside>

<aside>
양순미 박사
지도교수 김성봉
2026. 8.
---
대학상담센터 상담자의 심리적 소진과 회복 경험에 관한 근거이론적 접근
</aside>

<aside>
이승록 박사
지도교수 박정환
2026. 8.
---
상담심리학과 교육공학 전문가의 융합적 실천 과정
(The Convergent Practice Process of Counseling Psychology Professionals and Educational Technology)
상담심리학과 교육공학을 한 실천 안에서 함께 수행해 온 전문가 15명을 면담하여, 두 전문성의 경계에서 겪는 위기와 정체성의 재구성 과정을 근거이론으로 규명한다. 그 한가운데에는 ‘전문성 경계 위기’가 놓여 있었고, 양쪽 어디에도 온전히 속하지 못하는 ‘경계인’의 자리를 약점이 아니라 강점으로 다시 읽을 때 ‘통합자’로의 전환이 일어났다. 융합을 지식의 결합이 아니라 전문가의 존재 방식이 바뀌는 과정으로 재개념화하고, 새로운 탐구 영역으로 ‘상담공학’의 가능성을 제안한다.
</aside>
"""),

    # 회원 소식 — 발령 내용은 PI가 전달한 인사 공고문에 근거한다(2026-09-11).
    # 문안 원칙 (PI 2026-09-21):
    #   · **임기·임용기간은 적지 않는다.** 공고문에 있어도 지면에는 싣지 않는다.
    #   · 확인되지 않은 것은 쓰지 않는다 — 강동호 교수가 대외협력홍보실장을
    #     내려놓았는지 알 수 없어 그 문장을 뺐다. "비서실장을 맡으셨다"까지만.
    #   · '겸해 맡으셨습니다'는 어색하다 → '맡으셨습니다'.
    # 2026-09-21 PI가 인사 공고문 원문을 전달. 김대영 교수(교육대학원 교육학과)는
    #   ① 2026.09.01. 부교수 → **교수 승진**
    #   ② 2026.03.30.~2028.03.29. 교육혁신처장·원격교육원장·교양교육원장·
    #      교육성과관리센터장 네 자리 겸보
    #   ③ 2026.05.14. 교수학습지원센터장을 하나 더 (후임자 정해질 때까지)
    #   앞서 적혀 있던 5/14 건은 틀린 것이 아니라 나중에 더해진 자리였다.
    #   승진이 가장 큰 소식이라 앞세우고 보직은 시간 순으로 묶었다.
    #   카드 머리의 직함도 '부교수 → 교수'로 고쳤다.
    #   강동호 교수는 비서실장이 맞아 그대로 둔다(PI 재확인).
    # 2026-09-22 PI: 회원 소식 카드의 소속 깊이는 **학부/학과까지**로 맞춘다.
    #   연준모 교수를 전공까지(「…초등교육학전공」) 적었다가, 다른 두 카드와
    #   깊이가 어긋나 「제주대학교 교육대학 초등교육학부」로 되돌렸다.
    #   앞머리도 '제주대학교'로 통일했다(PI 2026-09-22) — 제주교육소식·시론이
    #   모두 '제주대학교'를 쓰고 있어 회원 소식만 줄여 쓸 이유가 없다.
    # 2026-09-21 PI 전달: 연준모 교수(교육대학 초등교육학부) 조교수 → 부교수 승진,
    #   2026.03.01. 발령, 임용기간 2036.02.29.까지.
    #   사진은 PI 지시로 제주대신문 기사(idxno=117173, 2024. 3.)의 본문 사진을 썼다.
    #   ★ 본인에게 받은 사진이 생기면 교체. 출처는 2026-작업기록.md §2에 적었다.
    # 2026-09-22 PI 결정: 카드 순서는 **직위 순** — 교수(김대영) · 부교수(연준모) ·
    #   조교수(강동호). 새 소식을 넣을 때도 이 순서를 지킨다.
    dict(slug="회원동정_회원소식", title="회원 소식", body="""
<aside>
![김대영 프로필](images/김대영_photo.jpg)
---
**김대영 교수 (제주대학교 교육대학원 교육학과)**
교수로 승진하셨습니다
김대영 교수님께서 2026년 9월 1일자로 **교수**로 승진하셨습니다. 진심으로 축하드립니다.
올해는 보직도 여럿 맡으셨습니다. 3월 30일자로 **교육혁신처장**을 맡으시며 **원격교육원장 · 교양교육원장 · 교육성과관리센터장**을 함께 겸하게 되셨고, 5월 14일에는 **교수학습지원센터장**까지 맡아 후임자가 정해질 때까지 센터를 이끄십니다.
대학의 교육과정과 수업, 교육성과를 두루 살피는 자리인 만큼, 앞으로 보여주실 활약을 기대합니다.
</aside>

<aside>
![연준모 프로필](images/연준모_photo.jpg)
---
**연준모 부교수 (제주대학교 교육대학 초등교육학부)**
부교수로 승진하셨습니다
연준모 교수님께서 2026년 3월 1일자로 **부교수**로 승진하셨습니다. 진심으로 축하드립니다.
제주지회에서는 부회장과 상임이사를 맡아 지회 일에 힘을 보태고 계십니다.
</aside>

<aside>
![강동호 프로필](images/강동호_photo.jpg)
---
**강동호 조교수 (제주대학교 교육대학원 교육학과)**
비서실장을 맡으셨습니다
강동호 교수님께서 2026년 7월 29일자로 제주대학교 **비서실장**을 맡으셨습니다.
연구와 강의에 더해 학내 일까지 이어 맡으시는 교수님께 감사와 응원을 함께 전합니다.
</aside>
"""),

    # 회원 신간 안내 — 2025년 재창간호와 같은 지면 이름·짜임
    #   (제목 / 저자·출판사·출간일 / 소개). 회원 이름을 굵게 하여 찾기 쉽게.
    # 2026-09-21 PI 확인 사항:
    #   · 서명석(제주대 교육대학) 저서는 검색으로 확인. 정가는 싣지 않는다(PI 2026-09-21).
    #     동일인 근거 — 교수진 소개의
    #     세부전공 '교육철학 및 교육과정철학', 책인숲의 성리학·동양고전 연속 저작,
    #     제주대 리포지터리 「퇴계 『성학십도』의 수양치료적 의의」가 일치한다.
    #   · 김은정(제주국제대)은 PI가 동일인임을 확인. 2011년 초판이고
    #     2026년 개정판이 나올 예정이라 '출간 예정'으로 적는다. ★ 나오면 출간일을 채운다.
    #   · 이인회 저서는 2025년 재창간호에 [eBook]판으로 이미 소개됐다.
    #     PI 판단으로 이번 호에 다시 싣는다(2026-09-21). 종이책 초판은 2021-10-30,
    #     전자책은 2025-04-30이다. PI가 "올해 개정판이 나오는 것 같다"고 하여
    #     김은정 저서와 같은 형식으로 '2026년 개정판 출간 예정'을 덧붙였다.
    #     ★ PI도 확실하지 않다고 했다. 이인회 회장이 회람 대상이니 직접 확인받을 것.
    # 2026-09-22 PI가 학지사 신간 안내 리플릿(2026-09-16)을 전달. 연준모 교수가
    #   공저자로 참여한 한국학습장애학회 편 신간이다. 출간일은 리플릿에도 적혀
    #   있지 않아 PI 전언대로 '출간 예정'으로 적는다. ★ 출간일을 받으면 채운다.
    #   정가(27,000원)는 리플릿에 있으나 서명석 저서와 같이 싣지 않는다
    #   (PI 2026-09-21 정가 표기 삭제 방침).
    # 2026-09-22 PI: 신간을 다 모았다 → 「서지사항을 알려주시면 이 자리에 함께
    #   싣습니다」 자리 표시를 뺐다. 연구비 수주·회비 납부자 명단의 같은 표시는
    #   아직 받는 중이라 그대로 둔다.
    # 2026-09-22 PI 지적: 책마다 저자/출판사/출간일 배열이 달랐다. **세 줄로 맞춘다.**
    #     1줄  **제목** (부제가 있으면 ` — 부제`)
    #     2줄  저자 / 출판사      ← 회원 이름만 굵게, 출판사는 반드시 뒤
    #     3줄  출간일            ← 굵게 하지 않는다 (굵은 글씨는 회원 이름 표시용)
    #   새 책을 넣을 때도 이 순서를 지킨다.
    # 2026-09-22 PI 논의: 카드 **차례**는 「이미 출간된 책 먼저, 나머지는 저자
    #   가나다순」 — 서명석(2026.2 출간) · 김은정 · 연준모 · 이인회.
    #   연차 순은 쓰지 않는다: 독자가 규칙을 볼 수 없고, 연차를 확인할 근거가
    #   없으며, 네 권 중 셋이 공저라 누구의 연차인지도 모호하고, 회원 동정란에서
    #   서열로 읽힌다. 제주교육소식도 저자 가나다순이라 지면 원칙이 하나로 묶인다.
    dict(slug="회원동정_신간안내", title="회원 신간 안내", body="""
회원들의 새 저서를 소개합니다.

<aside>
![태극도와 서명도 표지](images/책표지_태극도와서명도.jpg)
**태극도와 서명도** — 성학십도 2
**서명석** 저 / 책인숲
2026년 2월 출간
---
퇴계 이황의 『성학십도』를 풀어 읽는 연속 작업의 두 번째 권이다. 『성학십도』 제1도 태극도(太極圖)와 제2도 서명도(西銘圖)를 다룬다.
</aside>

<aside>
![놀이와 유아교육 표지](images/책표지_놀이와유아교육.jpg)
**놀이와 유아교육**
신은수 · **김은정** · 유영의 · 박현경 · 백경순 저 / 학지사
2011년 초판 · 2026년 개정판 출간 예정
---
놀이 이론, 놀이와 유아 발달, 놀이의 교육적 적용을 다룬다. 이론 연구와 현장 연구를 함께 엮어 유아교육에서 놀이가 갖는 의미를 짚는다.
</aside>

<aside>
![학습장애 및 학습지원 대상 학생 교육 표지](images/책표지_학습장애및학습지원대상학생교육.jpg)
**학습장애 및 학습지원 대상 학생 교육**
한국학습장애학회 편, 강은영 · 김동일 · 김소희 · 김우리 · 김제린 · 김희은 · 서선진 · 손승현 · 신미경 · 신재현 · **연준모** · 옥민욱 · 유은미 · 이대식 · 이예다나 · 정평강 · 지은 · 최승숙 공저 / 학지사
2026년 출간 예정
---
학습의 어려움을 가진 학생을 어떻게 이해하고 교실 안에서 어떻게 지원할 것인지를 다룬다. 읽기와 쓰기, 수학, 사회, 과학에서 사회성 및 행동 지원까지 학습장애와 학습지원 대상 학생을 위한 교육적 지원 방안을 영역별로 짚는다.
</aside>

<aside>
![오늘의 교육 내일의 교육정책 표지](images/책표지_오늘의교육내일의교육정책.jpg)
**오늘의 교육 내일의 교육정책**
박수정 · 김용 · 엄문영 · **이인회** · 이희숙 · 차성현 · 한은정 저 / 학지사
2021년 초판 · 전자책 2025년 4월 · 2026년 개정판 출간 예정
---
주요 교육정책 이슈를 소개하고 함께 생각할 주제를 던지는 책이다. 학교제도와 교육과정, 교육재정, 학교교육 체제, 교육자치와 참여의 네 부분으로 나누어 오늘의 교육을 짚고 내일의 정책 방향을 묻는다.
</aside>
"""),

    # 연구비 수주 — 과제명·기간·금액·발주기관은 공개 검색으로 확인되지 않는다.
    # 해당 교수님들께 직접 받아 아래 표의 빈칸을 채운다.
    # 2026-09-22 PI가 연준모 교수분을 전달: 과제명·연구기간·발주기관.
    # 2026-09-22 PI 결정: **총 연구비 열을 뺀다.** 본인이 금액 공개를 부담스러워
    #   하셨다. 빈칸으로 남겨두면 나머지 두 분께도 요구처럼 보인다.
    # 2026-09-22 PI 결정: 「확인되는 대로 채웁니다」 안내 문구도 뺀다 — 지면에
    #   편집 사정을 적지 않는다.
    # 2026-09-22 PI 결정: **고전·박정환 교수 줄을 뺀다.** 두 분의 수주 사실은
    #   본인이 알려주신 것이 아니라 전해 들은 것이다. 본인 확인 없이 회원 동정란에
    #   이름을 올릴 근거가 못 된다. 회신이 오면 그때 줄을 다시 넣는다.
    #   (PI가 고전 교수께 연락드렸고 회신 대기 중 — 내일 단톡방에서도 뵙는다.)
    # 2026-09-23 고전 교수 회신(임원 회람 피드백):
    #   ① 꼭지 이름을 **「연구 프로젝트 수행」**으로. '수주'는 돈을 받은 사실을
    #      가리키지만 이 표가 알리려는 것은 회원이 어떤 연구를 하고 있는가다.
    #   ② **총 연구비 대신 「연구진」**을 싣는다. 금액은 밝히지 않는 편이 좋다는
    #      뜻을 다시 확인하셨다(본인 과제 1,800만 원도 싣지 않는다).
    #   ③ 본인 과제를 직접 알려주셨다 — 9/22에 뺐던 줄을 근거를 갖춰 되살린다.
    #   ★ 연준모 교수 줄은 역할(연구책임/공동연구)을 확인하지 못해 이름만 둔다.
    #      확인되면 고전 교수 줄과 같은 꼴로 맞춘다.
    #   ★ 발주기관은 공식 명칭 「제주특별자치도의회」로 적었다(회신은 '도의회').
    #   ★ 연구진 칸의 규칙: **연구책임자에만 직함을 붙이고 공동연구자는 이름만**
    #      쓴다. 공동연구자 중에는 회원이 아닌 분도 있어 직함을 일부만 붙이면
    #      표가 들쭉날쭉해진다. 줄 차례는 연구 시작일 순.
    # 2026-09-23 PI 결정: 김민호 교수 과제를 **지금 회람본에 넣는다.** 9/28
    #   제안서 발표 뒤 수의계약 예정이라 아직 계약 전이지만, 고전 교수(공동연구자)가
    #   연구진까지 확정해 직접 알려주신 정보다.
    #   ★★ 9/29~30에 계약 성사를 확인하고 발간본에 유지할지 결정한다. 불발이면
    #      이 줄을 빼고 다시 빌드한다. 확인 없이 10/1 발간으로 넘기지 말 것.
    dict(slug="회원동정_연구비수주", title="연구 프로젝트 수행", body="""
| 연구과제명 | 연구진 | 연구기간 | 발주기관 |
|---|---|---|---|
| 제주 발달장애인 가족의 돌봄 현황 및 지원 방안 탐색 | 연준모 교수 | 2026년 3월 ~ 10월 | 제주여성가족연구원 |
| 제주특별자치도 교육의원·교육위원회제 20년 성과 분석 | 고전 교수(연구책임), 김성기, 백혜선, 장우천 | 2026년 7월 ~ 11월 | 제주특별자치도의회 |
| 제주형 자율학교 종합 진단 및 재구조화 | 김민호 교수(연구책임), 고전, 강은주, 곽유진, 문현식, 홍지환 | 2026년 10월 ~ 2027년 2월 | 제주특별자치도교육청 |
"""),

    # 명단을 받으면 아래 '명단' 아래 줄만 채우면 된다.
    # 구성원 소개에 있던 회원가입·회비 박스를 여기로 옮겼다(PI 2026-09-21).
    # 계좌가 두 상자에 겹치므로 기존 '납부 문의'와 한 상자로 합쳤다.
    dict(slug="활동소개_회비납부자", title="회비 납부자 명단", body="""
지회 운영에 함께해 주신 회원 여러분께 감사드립니다.

<aside>
**회원가입**
언제든지 가능 ([회원가입 바로가기](https://naver.me/5zXVy1er))
**가입비 겸 연회비**
30,000원 (농협, 302-2028-2520-51, 이인회)
**납부 문의**
사무국 황현철 (제주대학교)
</aside>

## 명단

<aside>
명단을 정리하는 대로 이 자리에 싣습니다.
</aside>
"""),
]

# 미제출 원고 — 자리(placeholder)만 만들어 레이아웃에서 보이게 한다
PENDING = [
    dict(slug="제주교육소식_김경주", title="지역사회 교육활동의 성과와 제언",
         name="김경주", affil="", role="박사",
         note="추석 전 제출 예정"),
    # 임원진이 한 편씩 더 쓰기로 했으나 아직 도착하지 않았다. 회람본에 자리를
    # 미리 만들어 두면 본인이 보고 보내주실 수 있다. (PI 2026-09-21)
    # 학과는 PI가 확인해 주었다 — 사회복지학과 (2026-09-22).
    dict(slug="제주교육소식_박한샘", title="한국상담심리학회 제주지회 창립",
         name="박한샘", affil="제주한라대학교 사회복지학과", role="교수",
         note="집필을 요청드렸습니다"),
    dict(slug="활동소개_이인회", title="제주교육학 제5차 공동학술대회",
         name="이인회", affil="한국교육학회 제주지회", role="회장",
         note="집필을 요청드렸습니다"),
]


# --------------------------------------------------------------------------
def find_src(filename):
    """원고 폴더에서 파일을 찾는다.

    드라이브는 00_발간사 / 01_주론 / 03_제주교육소식 … 으로 나뉘어 있어
    하위 폴더까지 뒤진다. 내려받을 때 붙는 ' (1)' 꼬리표도 떼고 찾아본다.
    """
    p = DL / filename
    if p.exists():
        return p
    hits = sorted(DL.rglob(filename))
    if hits:
        return hits[0]
    stem, dot, ext = filename.rpartition(".")
    bare = re.sub(r"(\s*\(\d+\))+$", "", stem).strip()
    if bare and bare != stem:
        hits = sorted(DL.rglob(bare + dot + ext))
        if hits:
            return hits[0]
    # 꼬리표가 사양이 아니라 '파일 쪽'에 붙어 있는 경우 — 같은 파일을 두 번
    # 내려받으면 드라이브에 '…(260922) (1).docx' 로 남는다. 사양에 그 꼬리표를
    # 적어두면 다음에 깨끗한 이름으로 받았을 때 못 찾으므로, 여기서 떼고 견준다.
    hits = [q for q in sorted(DL.rglob("*" + dot + ext))
            if re.sub(r"(\s*\(\d+\))+$", "", q.stem).strip() == bare]
    if hits:
        return hits[0]
    return None


def parse_meta(md):
    meta = {}
    for key in ("필자", "소속", "직함"):
        m = re.search(rf"^\|\s*{key}\s*\|\s*(.+?)\s*\|", md, re.M)
        if m:
            meta[key] = m.group(1).strip()
    return meta


# 템플릿이 두 종류다.
#   (가) "# 실제 제목"            — 제목을 H1에 바로 쓴 원고
#   (나) "## 제목" 다음 줄에 제목 — 템플릿의 항목 이름을 그대로 둔 원고
TITLE_HEADING = re.compile(r"^#{1,3}\s*제목\s*$", re.M)


def parse_title(md):
    """(제목, 본문 시작 위치)를 돌려준다."""
    m = TITLE_HEADING.search(md)
    if m:
        tail = md[m.end():]
        for line in tail.split("\n"):
            if line.strip():
                cut = m.end() + tail.index(line) + len(line)
                return line.strip().strip("*").strip(), cut
    m = re.search(r"^#\s+(.*)$", md, re.M)
    if m:
        return m.group(1).strip(), m.end()
    return "", 0


def split_body(md):
    _, cut = parse_title(md)
    body = md[cut:] if cut else md
    body = re.sub(r"^\|.*\|\s*$\n?", "", body, flags=re.M)
    body = re.sub(r"^\|-+\|.*$\n?", "", body, flags=re.M)

    prof = re.search(r"^#{1,3}\s*프로필\s*사진.*$", body, re.M)
    profile_img = None
    if prof:
        tail = body[prof.end():]
        im = re.search(r"!\[[^\]]*\]\(([^)]+)\)", tail)
        if im:
            profile_img = im.group(1)
        body = body[:prof.start()]

    # 템플릿 잔여 머리말·구분선 제거
    body = re.sub(r"^[─—-]{5,}\s*$", "", body, flags=re.M)
    body = re.sub(r"^##\s*본문\s*$", "", body, flags=re.M)
    return body.strip(), profile_img


# 원고 템플릿(`뉴스레터 원고_템플릿.docx`)의 안내 문구. 지우지 않고 그 위에 덧써
# 보내신 원고가 있어 본문에 그대로 딸려 온다 — 전새미 원고의 사진 아래에
# 「여기에 본문을 자유롭게 작성해주세요」가 남아 있었다(PI 지적 2026-09-22).
# 줄이 붙어 있든 나뉘어 있든 잡히도록 공백을 지우고 견준다.
TEMPLATE_RESIDUE = [
    "여기에 본문을 자유롭게 작성해주세요.",
    "이미지를 넣고 싶은 위치에 직접 삽입하시면 됩니다.(삽입 → 그림 → 이 디바이스에서)",
    "본문에 이미지를 넣고 싶은 위치에 직접 삽입해주세요.",
    "참고문헌이 있는 경우 아래에 APA 형식으로 작성해주세요.",
    "예) 홍길동(2024). 논문제목. 학술지명, 1(2), 1-10.",
    "아래에 프로필 사진을 삽입해주세요. (정면 얼굴, 가로세로 비율 1:1 권장)(삽입 → 그림 → 이 디바이스에서)",
    "[여기에 프로필 사진 삽입]",
    "아래 양식에 맞춰 내용을 작성해주세요.",
    "프로필 사진은 문서 맨 아래 지정된 위치에 삽입해주세요.",
    "작성 완료 후 .docx 파일 그대로 메일로 회신해주세요.",
    "분량: 1,000자(A4 1장) 이상",
    "참고문헌은 APA 형식을 권장합니다.",
    "※ 아래 점선 이후부터 작성해주세요. 이 안내문은 삭제하셔도 됩니다.",
    "작성 안내",
]
_RESIDUE_KEYS = {re.sub(r"\s+", "", t) for t in TEMPLATE_RESIDUE}


def drop_template_residue(body):
    """필자가 지우지 않고 보낸 템플릿 안내 문구를 걷어낸다.

    한 줄이 통째로 안내 문구일 때만 지운다 — 본문 안에 같은 말이 섞여 있는
    경우까지 건드리면 필자의 글을 깎는다.
    """
    kept = []
    for line in body.split("\n"):
        key = re.sub(r"\s+", "", line.strip().strip("*").strip())
        if key and key in _RESIDUE_KEYS:
            continue
        kept.append(line)
    return re.sub(r"\n{3,}", "\n\n", "\n".join(kept))


def lift_title_from_body(body):
    """본문 첫 실질 줄을 제목으로 올린다 (부제가 뒤따르면 함께)."""
    lines = [l for l in body.split("\n")]
    idx = next((i for i, l in enumerate(lines) if l.strip()), None)
    if idx is None:
        return "", "", body
    title = lines[idx].strip().strip("*").strip()
    sub = ""
    j = idx + 1
    while j < len(lines) and not lines[j].strip():
        j += 1
    if j < len(lines) and lines[j].strip().startswith(("―", "-", "—")):
        sub = lines[j].strip()
        rest = lines[j + 1:]
    else:
        rest = lines[idx + 1:]
    return title, sub, "\n".join(rest).strip()


# '1. 제주지회의 3가지 전환점' 처럼 번호가 붙은 소제목.
# 그대로 두면 마크다운이 번호 목록으로 읽어 각 항목이 따로 <ol>이 되고,
# 화면에는 1. 1. 1. 로 나온다. 소제목으로 올려 번호를 글자 그대로 살린다.
NUM_HEADING = re.compile(r"^(\d{1,2}\.\s+\S.{1,38})$", re.M)


def promote_numbered_headings(body):
    def repl(m):
        line = m.group(1).strip()
        # 문장이면 소제목이 아니다
        if line.endswith(("다.", "요.", "까?", "다", "음.")):
            return m.group(0)
        return f"## {line}"
    return NUM_HEADING.sub(repl, body)


# '■ 정책보다 사람이 먼저다' 처럼 네모 기호로 소제목을 표시한 원고가 있다.
# 그대로 두면 본문 문단으로 흘러 절의 층이 보이지 않고, 기호도 글꼴에 따라
# 깨져 보인다. 소제목으로 올리면 다른 원고와 같은 초록 굵은 글씨가 된다.
# (PI 2026-09-21)
BULLET_HEADING = re.compile(r"^[\u25a0\u25aa\u25c6\u25cf]\s*(\S.{0,60})$", re.M)


def promote_bullet_headings(body):
    return BULLET_HEADING.sub(lambda m: "## " + m.group(1).strip(), body)


# 소제목을 굵은 글씨 한 줄로만 표시한 원고가 있다 (연준모 시론 — 필자가
# html-to-docx 로 만들어 보내 제목·소제목이 모두 **굵게** 로만 왔다).
# 그대로 두면 본문 문단으로 흘러 절의 층이 보이지 않는다. 문장이 아닌 짧은
# 줄만 소제목으로 올려 다른 원고와 같은 초록 굵은 글씨가 되게 한다.
BOLD_HEADING = re.compile(r"^\*\*(\S.{0,40})\*\*$", re.M)


def promote_bold_headings(body):
    def repl(m):
        line = m.group(1).strip()
        if line.endswith(("다.", "요.", "까?", "음.", ".")):
            return m.group(0)
        return f"## {line}"
    return BOLD_HEADING.sub(repl, body)


# '가. 초기 저항과 토지 접근권 투쟁' / '1) 미 연방정부로부터의 자율성' 처럼
# 한글 순서 기호·괄호 번호가 붙은 하위 소제목. 그대로 두면 본문 문단으로 흘러
# 절의 층이 보이지 않는다.
GANADA_HEADING = re.compile(r"^([가나다라마바사아자차카타파하]\.\s+\S.{1,40})$", re.M)
PAREN_HEADING = re.compile(r"^(\d\)\s+\S.{1,40})$", re.M)


def promote_sub_headings(body):
    def repl(level):
        def inner(m):
            line = m.group(1).strip()
            if line.endswith(("다.", "요.", "까?", "음.")):
                return m.group(0)
            return f"{level} {line}"
        return inner
    body = GANADA_HEADING.sub(repl("###"), body)
    return PAREN_HEADING.sub(repl("####"), body)


def wrap_captions(body, captions):
    """사진 아래에 맨 문장으로만 적힌 캡션을 [사진설명= …] 꼴로 감싼다.

    원고마다 캡션 표기가 제각각이라 지면에서 캡션인지 본문인지 구분되지 않았다.
    표기를 하나로 맞추면 build.py가 캡션 꼴(작은 회색 가운데 글씨)로 그린다.
    (PI 2026-09-22)
    """
    for line in captions or []:
        old_line = line.strip()
        assert body.count(old_line) == 1, f"캡션 줄을 찾지 못했거나 여럿이다: {old_line[:30]}"
        body = body.replace(old_line, f"[사진설명= {old_line}]")
    return body


def apply_fixes(body, pairs):
    """오탈자·절 번호 교정. (틀린 것, 바른 것) 순서대로 적용한다."""
    for wrong, right in pairs or []:
        body = body.replace(wrong, right)
    return body


def replace_table(body, block):
    """본문의 표를 주어진 덩어리로 갈아끼운다.

    사진 캡션만 담긴 빈 표를 사진+캡션으로 바꿀 때 쓴다.
    """
    if not block:
        return body
    lines = body.split("\n")
    start = next((i for i, l in enumerate(lines) if l.startswith("|")), None)
    if start is None:
        return body
    end = start
    while end < len(lines) and lines[end].startswith("|"):
        end += 1
    return "\n".join(lines[:start] + block.split("\n") + lines[end:])


# 참고문헌 한 항목이 여러 줄로 끊겨 있는 원고가 있다.
#   Hawaii State Department of Education(2026). Adult education.
#   https://hawaiipublicschools.org/...에서
#   2026.9.14. 인출
# refstyle은 줄 단위로 항목을 가르므로, 이어지는 줄(URL·인출일)이 참고문헌이
# 아닌 부스러기로 버려진다. 넘기기 전에 한 줄로 이어 붙인다.
REF_ENTRY_START = re.compile(r"^[A-Z가-힣][^()\n]{0,70}\(\s*\d{4}[a-z]?\s*\)")


def join_reference_lines(body):
    m = refstyle.HEADING.search(body)
    if not m:
        return body
    head, tail = body[:m.end()], body[m.end():]
    merged = []
    for line in (l.strip() for l in tail.split("\n")):
        if not line:
            continue
        if merged and not REF_ENTRY_START.match(line):
            merged[-1] = f"{merged[-1]} {line}"
        else:
            merged.append(line)
    return head + "\n" + "\n".join(merged) + "\n"


def add_photo(filename, slug, box, src=None):
    """인물 사진을 잘라 소스 트리에 넣고 상대경로를 돌려준다.

    기본 출처는 드라이브의 뉴스레터 폴더다. src 를 주면 그 파일을 쓴다 —
    지난 호에 필자가 직접 보내 실린 사진을 다시 쓸 때.
    """
    from PIL import Image
    src = Path(src) if src else find_src(filename)
    if not src or not src.exists():
        return None
    img_dir = SUB / "images"
    img_dir.mkdir(parents=True, exist_ok=True)
    im = Image.open(src).convert("RGB")
    if box:
        im = im.crop(box)
    side = min(im.size)
    im = im.crop((0, 0, side, side))
    name = f"{slug}_photo.jpg"
    im.save(img_dir / name, quality=92)
    return f"images/{name}"


def crop_square(path, box):
    """프로필 사진을 지정 영역으로 잘라 정사각형으로 만든다.

    세로로 긴 사진은 원형 틀에서 얼굴이 위로 몰린다. 얼굴을 감싸는
    정사각형으로 미리 잘라두면 원 중앙에 얼굴이 온다.
    """
    from PIL import Image
    im = Image.open(path)
    mode = im.mode
    im = im.crop(box)
    side = min(im.size)
    im = im.crop((0, 0, side, side))
    im.convert(mode).save(path)


# hwp·워드에서 넘어오는 방점(U+302E/302F)은 대부분 글꼴에 없어 ○ 로 보인다.
# 구분 기호로 쓰인 것이므로 가운뎃점으로 바꾼다.
TONE_MARKS = {"〮": "·", "〯": "·"}


def clean_text(s):
    for bad, good in TONE_MARKS.items():
        s = s.replace(bad, good)
    return re.sub(r"\s*·\s*", " · ", s).strip() if "·" in s else s.strip()


def author_block(name, affil, role, profile_img=None):
    name, affil, role = clean_text(name), clean_text(affil), clean_text(role)
    if not name:
        return ""
    p = ["<aside>"]
    if profile_img:
        p += [f"![{name} 프로필]({profile_img})", ""]
    p += [f"**{name}**", ""]
    if affil:
        p += [affil, ""]
    if role:
        p += [role, ""]
    p += ["</aside>", ""]
    return "\n".join(p)


def write_md(slug, title, author, body, subtitle=""):
    parts = [f"# {title}", ""]
    if subtitle:
        parts += [f"*{subtitle}*", ""]
    if author:
        parts.append(author)
    parts += [body, ""]
    (SUB / f"{slug}.md").write_text("\n".join(parts), encoding="utf-8")


# --------------------------------------------------------------------------
def main():
    global DL
    ap = argparse.ArgumentParser(description="2026 제주교육마루 원고 → 마크다운 소스 트리")
    ap.add_argument("--src", default=str(DL),
                    help="원고 원본(docx/hwpx)이 모여 있는 폴더")
    ap.add_argument("--only", default=None,
                    help="이 슬러그 하나만 다시 만든다 (나머지 파일은 그대로 둔다). "
                         "원고가 한 편 늦게 도착했을 때 쓴다 — 전체 재생성은 "
                         "원본 docx가 모두 있어야 하므로 나머지를 지워버린다.")
    ns = ap.parse_args()
    DL = Path(ns.src)
    only = ns.only

    if only:
        if not SUB.exists():
            sys.exit(f"--only 는 기존 소스 트리가 있어야 한다: {SUB}")
    else:
        if OUT.exists():
            shutil.rmtree(OUT)
    SUB.mkdir(parents=True, exist_ok=True)

    def skip(slug):
        return bool(only) and slug != only

    rows = []
    ref_notes = []

    # ---- docx 원고 ----
    for spec in DOCX:
        if skip(spec["slug"]):
            continue
        src = find_src(spec["file"])
        if not src:
            rows.append((spec["slug"], "없음", 0, 0, ""))
            continue
        raw, images = docx_convert(src, SUB, spec["slug"].split("_")[-1])
        meta = parse_meta(raw)
        body, prof = split_body(raw)
        title, _ = parse_title(raw)
        subtitle = ""

        if spec.get("title_from_body"):
            t, subtitle, body = lift_title_from_body(body)
            if t:
                title = t
        if spec.get("title_override"):
            title = spec["title_override"]
        if not title or "제목을 적어" in title:
            title = spec.get("title_override") or "(제목 미정)"

        body = drop_template_residue(body)
        body = wrap_captions(body, spec.get("captions"))
        body = promote_bullet_headings(body)
        if spec.get("bold_headings"):
            body = promote_bold_headings(body)

        if spec["slug"] in REF_OVERRIDE:
            body = refstyle.replace_section(body, REF_OVERRIDE[spec["slug"]])
            ref = {"count": 0, "dropped": [], "dupes": [], "removed": [],
                   "override": True}
        else:
            body, ref = refstyle.normalize_section(
                body, drop=REF_DROP.get(spec["slug"]))
        if ref:
            ref_notes.append((spec["slug"], ref))

        if prof and spec.get("photo_crop"):
            crop_square(SUB / prof, spec["photo_crop"])
        elif spec.get("photo"):
            # 원고에 프로필 사진이 딸려 오지 않은 경우 드라이브의 사진을 쓴다.
            prof = add_photo(spec["photo"], spec["slug"], spec.get("photo_crop"))

        author = author_block(spec.get("name_override") or meta.get("필자", ""),
                              spec.get("affil_override") or meta.get("소속", ""),
                              spec.get("role_override") or meta.get("직함", ""),
                              prof)
        write_md(spec["slug"], title, author, body, subtitle)
        rows.append((spec["slug"], "수합", len(body), len(images), title))

    # ---- hwpx 원고 ----
    for spec in HWPX:
        if skip(spec["slug"]):
            continue
        src = find_src(spec["file"])
        if not src:
            rows.append((spec["slug"], "없음", 0, 0, ""))
            continue
        raw, images = hwpx_convert(src, SUB, spec["slug"].split("_")[-1])
        lines = raw.split("\n")
        # 제목·필자 줄 제거
        kept, dropped = [], 0
        for l in lines:
            if dropped < spec["drop_first"] and l.strip():
                dropped += 1
                continue
            kept.append(l)
        body = drop_template_residue("\n".join(kept).strip())
        body = apply_fixes(body, TYPO_FIX.get(spec["slug"]))
        body = apply_fixes(body, spec.get("heading_fix"))
        body = replace_table(body, spec.get("table_replace"))
        body = promote_numbered_headings(body)
        if spec.get("sub_headings"):
            body = promote_sub_headings(body)
        body = join_reference_lines(body)
        if spec["slug"] in REF_OVERRIDE:
            body = refstyle.replace_section(body, REF_OVERRIDE[spec["slug"]])
            ref = {"count": 0, "dropped": [], "dupes": [], "removed": [],
                   "override": True}
        else:
            body, ref = refstyle.normalize_section(
                body, drop=REF_DROP.get(spec["slug"]))
        if ref:
            ref_notes.append((spec["slug"], ref))
        prof = None
        if spec.get("photo") or spec.get("photo_path"):
            prof = add_photo(spec.get("photo"), spec["slug"],
                             spec.get("photo_crop"),
                             src=REPO / spec["photo_path"]
                             if spec.get("photo_path") else None)
        author = author_block(spec["name"], spec["affil"], spec["role"], prof)
        write_md(spec["slug"], spec["title"], author, body)
        rows.append((spec["slug"], "수합", len(body), len(images), spec["title"]))

    # ---- 신간 안내 책 표지 ----
    if not skip("회원동정_신간안내"):
        (SUB / "images").mkdir(parents=True, exist_ok=True)
        for name in BOOK_COVERS:
            c = find_src(name)
            if c:
                shutil.copy2(c, SUB / "images" / name)
            else:
                print(f"  ! 표지 없음: {name}")

    # ---- 회원 소식 인물 사진 ----
    for slug, fname, box in NEWS_PHOTOS:
        if skip(slug):
            continue
        if not add_photo(fname, slug, box):
            print(f"  ! 사진 없음: {fname}")

    # ---- 고정 지면 (구성원 소개·회비 납부자 명단) ----
    for spec in STATIC:
        if skip(spec["slug"]):
            continue
        write_md(spec["slug"], spec["title"], "", spec["body"].strip())
        rows.append((spec["slug"], "고정지면", len(spec["body"]), 0,
                     spec["title"]))

    # ---- 미제출 자리 ----
    for spec in PENDING:
        if skip(spec["slug"]):
            continue
        author = author_block(spec["name"], spec["affil"], spec["role"])
        body = (f"<aside>\n원고 준비 중입니다. ({spec['note']})\n</aside>\n")
        write_md(spec["slug"], spec["title"], author, body)
        rows.append((spec["slug"], "준비중", 0, 0, spec["title"]))

    if not only:
        # ---- 발간사 (루트 MD) ----
        # (--only 로 한 편만 고칠 때는 아래 고정 산출물을 다시 쓰지 않는다)
        # 회장 명의 초안. 작년 재창간호의 어투를 따랐다. 최종 문안은 회장 확인 필요.
        (OUT / "발간사.md").write_text("""# 발간의 글

    존경하는 회원 여러분,

    한국교육학회 제주지회 회장 이인회입니다.

    지난해 재창간호를 내놓은 데 이어 『제주교육마루』 제2호를 펴냅니다. 한 해에 한 번 나오는 지면이지만, 그 사이에 제주 교육에는 적지 않은 변화가 있었습니다. 새 교육감이 취임하였고, 우리 지회는 창립 60주년을 눈앞에 두고 있습니다.

    이번 호에서는 지난 60년을 함께 돌아보았습니다. 1967년 창립 이래 제주지회가 지나온 세 번의 전환점을 짚고, 앞으로 우리가 어디로 가야 할지를 주론에 담았습니다. 회원 여러분께서 보내주신 글에는 연구실과 교실, 마을과 학교 현장에서 교육을 고민해 온 시간이 고스란히 담겨 있습니다. 저마다 자리는 다르지만 같은 물음을 품고 있다는 것을 이 지면에서 확인하게 됩니다.

    발간 형태도 한 걸음 나아갔습니다. 지난해 노션(Notion)으로 옮겨왔던 뉴스레터를 올해는 독립된 웹페이지로 다시 지었습니다. 주소 하나로 언제든 찾아볼 수 있고, 해마다 쌓인 기록이 그대로 남습니다. 작은 변화지만 지회의 이야기를 오래 간직하려는 뜻입니다.

    바쁘신 가운데 원고를 보내주신 필자 여러분, 축하의 말씀을 보내주신 고의숙 교육감님, 그리고 한 호를 엮어내기까지 애써주신 뉴스레터위원회에 깊이 감사드립니다.

    이 지면이 소식을 전하는 데 그치지 않고 제주 교육의 내일을 함께 궁리하는 자리가 되기를 바랍니다. 회원 여러분의 많은 관심과 참여를 부탁드립니다.

    감사합니다.

    한국교육학회 제주지회 회장  이인회
    """, encoding="utf-8")

        # ---- 공통 이미지 ----
        out_img = REPO / "2026" / "images"
        out_img.mkdir(parents=True, exist_ok=True)
        for src_name, dst_name in [("배너.jpg", "banner.jpg"),
                                   ("로고.png", "logo.png")]:
            s = find_src(src_name)      # 배너_이미지/ 하위에 있다
            if s:
                shutil.copy2(s, out_img / dst_name)
        # 캘리그래피는 드라이브 원본에 투명 격자가 박혀 있어 2025년 확정본을 쓴다
        for name in ["calligraphy-white.png", "calligraphy-white-v2.png"]:
            s = REPO / "2025" / "images" / name
            if s.exists():
                shutil.copy2(s, out_img / name)
        for name in ["favicon.ico"]:
            s = REPO / name
            if s.exists():
                shutil.copy2(s, REPO / "2026" / name)
        s = REPO / "2025" / "images" / "apple-touch-icon.png"
        if s.exists():
            shutil.copy2(s, out_img / "apple-touch-icon.png")

    print(f"{'파일':<24}{'상태':<8}{'본문자수':>8}{'이미지':>6}  제목")
    print("-" * 96)
    for slug, st, n, ni, title in rows:
        print(f"{slug:<24}{st:<8}{n:>8}{ni:>6}  {title[:42]}")
    if ref_notes:
        print("\n참고문헌 형식 통일")
        for slug, r in ref_notes:
            if r.get("override"):
                print(f"  {slug:<22} PI 확정본 적용")
                continue
            line = f"  {slug:<22} {r['count']}항목"
            if r["dupes"]:
                line += f" · 중복 {len(r['dupes'])}건 제거"
            if r["dropped"]:
                line += f" · 참고문헌 아닌 줄 {len(r['dropped'])}건 제외"
            if r.get("removed"):
                line += f" · 빼기로 한 항목 {len(r['removed'])}건 제외"
            print(line)
            for d in r["dropped"]:
                print(f"      제외: {d[:60]}")

    print(f"\n소스: {OUT}")


if __name__ == "__main__":
    main()
