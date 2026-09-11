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
         photo_crop=(35, 0, 145, 110)),   # 세로 사진 — 얼굴이 원 중앙에 오게
]

HWPX = [
    dict(file="뉴스레터 2026 주론-0906.hwpx", slug="주론_이인회",
         title="지나온 60년, 새롭게 열어갈 제주 교육",
         name="이인회", affil="한국교육학회 제주지회", role="회장",
         drop_first=2),
    dict(file="한국교육학회 제주지회 창립 60주년 축사-교육감(0825).hwpx",
         slug="인사말_고의숙",
         title="한국교육학회 제주지회 창립 60주년을 향한 축하말씀",
         name="고의숙", affil="제주특별자치도교육청", role="교육감",
         drop_first=1),
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
    dict(slug="회원동정", title="회원 동정", name="", affil="", role="",
         note="연구비 수주(고전·연준모·박정환), 저서 발간(연준모), "
              "보직(강동호), 학위 취득 — 조사 예정"),
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
                              meta.get("직함", ""), prof)
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
        body = "\n".join(kept).strip()
        if spec["slug"] in REF_OVERRIDE:
            body = refstyle.replace_section(body, REF_OVERRIDE[spec["slug"]])
            ref = {"count": 0, "dropped": [], "dupes": [], "removed": [],
                   "override": True}
        else:
            body, ref = refstyle.normalize_section(
                body, drop=REF_DROP.get(spec["slug"]))
        if ref:
            ref_notes.append((spec["slug"], ref))
        author = author_block(spec["name"], spec["affil"], spec["role"])
        write_md(spec["slug"], spec["title"], author, body)
        rows.append((spec["slug"], "수합", len(body), len(images), spec["title"]))

    # ---- 미제출 자리 ----
    for spec in PENDING:
        author = author_block(spec["name"], spec["affil"], spec["role"])
        body = (f"<aside>\n원고 준비 중입니다. ({spec['note']})\n</aside>\n")
        write_md(spec["slug"], spec["title"], author, body)
        rows.append((spec["slug"], "준비중", 0, 0, spec["title"]))

    # ---- 발간사 (루트 MD) ----
    (OUT / "발간사.md").write_text(
        "# 발간의 글\n\n"
        "<aside>\n회장 명의 발간사는 준비 중입니다.\n</aside>\n\n"
        "2026 제주교육마루(제2호)를 펴냅니다. "
        "한국교육학회 제주지회는 창립 60주년을 앞두고 있습니다.\n",
        encoding="utf-8")

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
