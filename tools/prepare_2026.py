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

# 원고 원본이 모여 있는 폴더. 구글 드라이브 '2026 제주지회 뉴스레터'에서
# 내려받은 docx/hwpx를 한곳에 모아두고 --src 로 지정한다.
DL = Path(os.environ.get("NEWSLETTER_SRC", Path.home() / "Downloads"))
REPO = Path(__file__).resolve().parent.parent
OUT = REPO / "_src2026"
SUB = OUT / "한국교육학회 제주지회 뉴스레터"

# --------------------------------------------------------------------------
# 원고 목록
#   title_override: 원고에 제목이 비어 있거나 템플릿 문구인 경우 편집자가 지정.
#                   ★ 표시는 PI 확인 필요.
# --------------------------------------------------------------------------
DOCX = [
    dict(file="뉴스레터 원고_아라캠퍼스_홍지환.docx", slug="제주교육소식_홍지환",
         title_override="성읍마을에서 만난 작은 주인공들",   # ★ 임시 제목
         needs_title=True),
    dict(file="뉴스레터 원고_김지원.docx", slug="제주교육소식_김지원"),
    dict(file="뉴스레터 원고 (전새미).docx", slug="제주교육소식_전새미",
         title_from_body=True),
    dict(file="뉴스레터 원고_홍지오.docx", slug="제주교육소식_홍지오"),
    dict(file="석진아_뉴스레터 원고_템플릿 (1).docx", slug="제주교육소식_석진아"),
    dict(file="조천_마을탐방_뉴스레터_원고_양유정 (4) (1).docx", slug="활동소개_양유정",
         title_from_body=True),
    dict(file="뉴스레터 원고_템플릿 _창립기념행사_황현철.docx", slug="활동소개_황현철"),
    dict(file="뉴스레터 원고_강한호.docx", slug="제주교육소식_강한호",
         photo_crop=(35, 0, 145, 110),    # 세로 사진 — 얼굴이 원 중앙에 오게
         # 원고에는 '교사'로만 적혀 있으나 이번 호 회원 동정에 박사학위 취득이
         # 함께 실리므로 직함에 박사를 덧붙였다. (PI 확인 2026-09-11)
         role_override="교사 · 교육학박사"),
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
]

# 회원 소식에 넣을 인물 사진 — 얼굴이 원 중앙에 오도록 자른다.
NEWS_PHOTOS = [
    ("강동호", "강동호 교수.jpg", (18, 0, 132, 114)),
    ("김대영", "김대영 교수.jpg", (23, 23, 153, 153)),
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
    dict(slug="활동소개_구성원소개", title="구성원 소개", body="""
제주지회 구성원 : 70명 (2025년 2월 기준)

<aside>
**회원가입**
언제든지 가능 ([회원가입 바로가기](https://naver.me/5zXVy1er))
**가입비 겸 연회비**
30,000 (농협, 302-2028-2520-51, 이인회)
</aside>

## 임원진

### 회장

<aside>
이인회 교수 (제주대학교 아라캠퍼스)
- 임기(2025.1.1.-2026.12.31.)
- 전화번호: 064-754-2163
- 이메일: tomlee@jejunu.ac.kr
- 지회 주소: 제주대학교 아라캠퍼스 사범대학 2호관 1312호 (이인회 교수 연구실)
</aside>

### 부회장

<aside>
김영희 교수 (제주관광대학교)
김은정 교수 (제주국제대학교)
연준모 교수 (제주대학교 사라캠퍼스)
최보영 교수 (제주대학교 아라캠퍼스)
박한샘 교수 (제주한라대학교)
</aside>

### 상임이사

<aside>
연준모 교수 (제주대학교 사라캠퍼스)
</aside>

### 뉴스레터위원회 위원장

<aside>
양은별 교수 (제주대학교 사라캠퍼스)
</aside>

### 감    사

<aside>
하진의 교수 (제주국제대학교)
</aside>

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
    # 요약은 공개된 초록을 바탕으로 작성했다. 초록이 공개되지 않은 2건은
    # 작년 호와 같이 제목·지도교수·학위수여 시기만 싣는다(내용 지어내지 않음).
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
이 연구는 지역사회 기반 대안학교인 광명YMCA 볍씨학교가 설립된 때부터 현재에 이르기까지의 전개 과정을 사례연구로 추적한다. 학교는 공동 의사결정 구조와 지역사회 기반 시민참여형 교육과정을 갖추며 출발하였고, 지역사회 상호작용 및 확장기(2007~2017)를 거치면서 학부모와 학생이 학교 운영과 공동체 생활, 자립의 경험을 통해 교육의 주체로 성장하였다. 연구는 학교와 지역사회의 관계가 자원을 활용하거나 일방적으로 지원하는 차원을 넘어, 공동의 교육적 가치를 중심으로 서로의 역할과 의미를 다시 세우는 ‘공존적 관계’로 형성되어 왔음을 밝힌다. 관계 재구성기(2018~현재)에는 학생 수 감소와 재정·제도의 한계로 기존 관계망이 도전에 직면하였으나, 학교는 그 속에서 지역사회 기반 대안교육의 정당성을 다시 물으며 관계와 운영 방식을 새롭게 짜고 있었다. 이는 지역사회 기반 대안학교의 지속가능성이 변화하는 조건 속에서 관계를 재구성하는 역량에 의해 형성됨을 보여준다.
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
이 연구는 상담심리학과 교육공학을 한 실천 안에서 함께 수행해 온 전문가들이 두 전문성의 경계에서 겪는 위기와 정체성의 재구성 과정을 근거이론으로 규명한다. 융합적 실천 전문가 15명을 면담한 결과, 그 한가운데에는 ‘전문성 경계 위기’가 놓여 있었다. 이 위기는 능력의 부족이 아니라 충분한 숙련 끝에 자기 전문성의 구조적 한계를 자각할 때 비로소 나타났으며, 위기를 학습으로 밀고 간 동력은 외적 보상이 아니라 과거에 충분히 돕지 못한 대상에 대한 윤리적 책임감이었다. 융합적 실천은 자각기·결단기·학습기·통합기·결말기의 다섯 단계를 거치며, 양쪽 어디에도 온전히 속하지 못하는 ‘경계인’의 자리를 약점이 아니라 양쪽을 함께 조망하는 강점으로 다시 읽을 때 ‘통합자’로의 전환이 일어났다. 연구는 융합을 지식의 양적 결합이 아니라 전문가의 존재 방식이 바뀌는 과정으로 재개념화하고, 두 논리가 한 실천 안에서 부딪히고 재구성되는 현상을 다루는 새로운 탐구 영역으로 ‘상담공학’의 가능성을 제안한다.
</aside>
"""),

    # 회원 소식 — 발령 내용은 PI가 전달한 인사 공고문에 근거한다(2026-09-11).
    dict(slug="회원동정_회원소식", title="회원 소식", body="""
<aside>
![강동호 프로필](images/강동호_photo.jpg)
---
**강동호 조교수 (제주대 교육대학원 교육학과)**
비서실장을 맡으셨습니다
강동호 교수님께서 2026년 7월 29일자로 제주대학교 **비서실장**을 겸해 맡으셨습니다. 임기는 2028년 3월 29일까지입니다.
그동안 맡아 오신 대외협력홍보실장 소임은 이번에 내려놓으셨습니다.
연구와 강의에 더해 학내 일까지 이어 맡으시는 교수님께 감사와 응원을 함께 전합니다.
</aside>

<aside>
![김대영 프로필](images/김대영_photo.jpg)
---
**김대영 부교수 (제주대 교육대학원 교육학과)**
교수학습지원센터를 이끄십니다
김대영 교수님께서 2026년 5월 14일자로 제주대학교 **교수학습지원센터장**을 겸해 맡으셨습니다. 후임자가 정해질 때까지 센터를 이끄십니다.
학생들의 배움과 교수님들의 수업을 함께 살피는 자리인 만큼, 앞으로 보여주실 활약을 기대합니다.
</aside>
"""),

    # 신간 안내 — 연준모 교수 저서는 온라인 서점(교보·알라딘·예스24)에서
    # 최근 발간본이 확인되지 않았다. 서지사항을 받으면 카드 하나를 추가한다.
    dict(slug="회원동정_신간안내", title="신간 안내", body="""
<aside>
회원들의 새 저서를 소개합니다. 서지사항을 알려주시면 이 자리에 싣습니다.
</aside>
"""),

    # 연구비 수주 — 과제명·기간·금액·발주기관은 공개 검색으로 확인되지 않는다.
    # 해당 교수님들께 직접 받아 아래 표의 빈칸을 채운다.
    dict(slug="회원동정_연구비수주", title="연구비 수주", body="""
| 연구과제명 | 연구책임자 | 연구기간 | 총 연구비 | 발주기관 |
|---|---|---|---|---|
|  | 고 전 교수 |  |  |  |
|  | 연준모 교수 |  |  |  |
|  | 박정환 교수 |  |  |  |

<aside>
과제명·연구기간·연구비·발주기관은 확인되는 대로 채웁니다.
</aside>
"""),

    # 명단을 받으면 아래 '명단' 아래 줄만 채우면 된다.
    dict(slug="활동소개_회비납부자", title="회비 납부자 명단", body="""
지회 운영에 함께해 주신 회원 여러분께 감사드립니다.

<aside>
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
    dict(slug="시론_김민호", title="제주지회 60주년 소회와 전망",
         name="김민호", affil="제주대학교", role="명예교수",
         note="9월 12일(금) 제출 예정"),
    dict(slug="시론_연준모", title="시론",
         name="연준모", affil="제주대학교", role="교수",
         note="제출 기한 확인 중"),
    dict(slug="제주교육소식_김경주", title="지역사회 교육활동의 성과와 제언",
         name="김경주", affil="", role="박사",
         note="추석 전 제출 예정"),
]


# --------------------------------------------------------------------------
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


# 구글 드라이브의 뉴스레터 폴더 — 인물 사진이 여기 모여 있다.
PHOTO_DIR = Path(os.environ.get(
    "NEWSLETTER_PHOTOS", r"G:\내 드라이브\2026 제주지회 뉴스레터"))


def add_photo(filename, slug, box):
    """드라이브의 인물 사진을 잘라 소스 트리에 넣고 상대경로를 돌려준다."""
    from PIL import Image
    src = PHOTO_DIR / filename
    if not src.exists():
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
    DL = Path(ap.parse_args().src)

    if OUT.exists():
        shutil.rmtree(OUT)
    SUB.mkdir(parents=True, exist_ok=True)

    rows = []
    ref_notes = []

    # ---- docx 원고 ----
    for spec in DOCX:
        src = DL / spec["file"]
        if not src.exists():
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

        author = author_block(meta.get("필자", ""), meta.get("소속", ""),
                              spec.get("role_override") or meta.get("직함", ""),
                              prof)
        write_md(spec["slug"], title, author, body, subtitle)
        rows.append((spec["slug"], "수합", len(body), len(images), title))

    # ---- hwpx 원고 ----
    for spec in HWPX:
        src = DL / spec["file"]
        if not src.exists():
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
        body = promote_numbered_headings("\n".join(kept).strip())
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
        if spec.get("photo"):
            prof = add_photo(spec["photo"], spec["slug"], spec.get("photo_crop"))
        author = author_block(spec["name"], spec["affil"], spec["role"], prof)
        write_md(spec["slug"], spec["title"], author, body)
        rows.append((spec["slug"], "수합", len(body), len(images), spec["title"]))

    # ---- 회원 소식 인물 사진 ----
    for slug, fname, box in NEWS_PHOTOS:
        if not add_photo(fname, slug, box):
            print(f"  ! 사진 없음: {fname}")

    # ---- 고정 지면 (구성원 소개·회비 납부자 명단) ----
    for spec in STATIC:
        write_md(spec["slug"], spec["title"], "", spec["body"].strip())
        rows.append((spec["slug"], "고정지면", len(spec["body"]), 0,
                     spec["title"]))

    # ---- 미제출 자리 ----
    for spec in PENDING:
        author = author_block(spec["name"], spec["affil"], spec["role"])
        body = (f"<aside>\n원고 준비 중입니다. ({spec['note']})\n</aside>\n")
        write_md(spec["slug"], spec["title"], author, body)
        rows.append((spec["slug"], "준비중", 0, 0, spec["title"]))

    # ---- 발간사 (루트 MD) ----
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
        s = DL / src_name
        if s.exists():
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
