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
    dict(slug="제주교육소식_강한호", title="박사학위 취득 소감",
         name="강한호", affil="제주대학교", role="박사",
         note="수합 완료로 기록돼 있으나 드라이브에 파일 없음 — 확인 필요"),
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


def parse_title(md):
    m = re.search(r"^#\s+(.*)$", md, re.M)
    return m.group(1).strip() if m else ""


def split_body(md):
    m = re.search(r"^#\s+.*$", md, re.M)
    body = md[m.end():] if m else md
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


def author_block(name, affil, role, profile_img=None):
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

    # ---- docx 원고 ----
    for spec in DOCX:
        src = DL / spec["file"]
        if not src.exists():
            rows.append((spec["slug"], "없음", 0, 0, ""))
            continue
        raw, images = docx_convert(src, SUB, spec["slug"].split("_")[-1])
        meta = parse_meta(raw)
        body, prof = split_body(raw)
        title = parse_title(raw)
        subtitle = ""

        if spec.get("title_from_body"):
            t, subtitle, body = lift_title_from_body(body)
            if t:
                title = t
        if spec.get("title_override"):
            title = spec["title_override"]
        if not title or "제목을 적어" in title:
            title = spec.get("title_override") or "(제목 미정)"

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
    print(f"\n소스: {OUT}")


if __name__ == "__main__":
    main()
