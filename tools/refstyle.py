# -*- coding: utf-8 -*-
"""참고문헌 형식 통일.

원고마다 참고문헌 표기가 제각각이라(국문 붙여쓰기·APA 국문·APA 영문·각주식)
하나로 맞춘다. 목표 형식은 국내 교육학 관행을 따른다.

  국문  저자(연도). 제목. 출처, 권(호), 쪽.
  영문  APA 그대로 — Adams, C. (2002). Title. Journal, 43(8), 973-987.

내용은 바꾸지 않는다. 구두점·간격·순서만 손대고, 참고문헌이 아닌 부스러기와
완전히 같은 중복 항목만 걸러낸다. 판단이 필요한 것은 report에 담아 사람이 본다.
"""
import re

HEADING = re.compile(r"^\s*(?:#{1,3}\s*)?<?\s*참고\s*문헌\s*>?\s*$", re.M)
HANGUL = re.compile(r"[가-힣]")
YEAR_IN_PAREN = re.compile(r"\(\s*\d{4}")


def _is_reference(line):
    """참고문헌 항목인가 — 연도가 담긴 괄호가 있어야 한다."""
    return bool(YEAR_IN_PAREN.search(line)) or bool(
        re.search(r",\s*\d{4}\s*\)", line))


def split_jammed(line):
    """한 줄에 여러 항목이 붙어 있으면 나눈다.

    '...1-25. 제주대학교, & 제주특별자치도. (2025). ...' 처럼
    마침표 뒤에 새 '저자 (연도)'가 시작되는 지점에서 자른다.
    """
    # 영문 이름의 이니셜('Levinson, S. C. (1983)')에서 잘리지 않도록
    #   · 앞이 한 글자 이니셜이면 경계로 보지 않는다
    #   · 뒤따르는 저자는 한글 2자 이상 또는 영문 단어여야 한다
    boundary = re.compile(
        r"(?<![A-Z]\.)(?<=\.)\s+"
        r"(?=(?:[가-힣]{2,}|[A-Z][a-z]+)[^()]{0,60}?\.?\s*\(\s*\d{4})")
    parts = [p.strip() for p in boundary.split(line) if p.strip()]
    return parts or [line]


def normalize_entry(text):
    """한 항목의 구두점·간격을 목표 형식으로 맞춘다."""
    s = " ".join(text.split())

    # 각주식 → 저자-연도식
    #   교육부, 『제목』 (서울: 교육부, 2022), 1108쪽.
    # → 교육부(2022). 『제목』. 서울: 교육부, 1108쪽.
    m = re.match(
        r"^(?P<author>[^,]+),\s*(?P<title>[『「\"].+?[』」\"])\s*"
        r"\(\s*(?P<place>[^:()]+):\s*(?P<pub>[^,()]+),\s*(?P<year>\d{4})\s*\)"
        r"(?P<rest>.*)$", s)
    if m:
        rest = m.group("rest").strip().lstrip(",").strip()
        s = (f'{m.group("author").strip()}({m.group("year")}). '
             f'{m.group("title").strip()}. '
             f'{m.group("place").strip()}: {m.group("pub").strip()}')
        s = s + (f", {rest}" if rest else ".")

    # 괄호 안 날짜 표기 통일:  (2026, 7월 6일) → (2026. 7. 6.)
    def fix_date(mm):
        y, mo, d = mm.group(1), mm.group(2), mm.group(3)
        return f"({y}. {mo}." + (f" {d}.)" if d else ")")
    s = re.sub(r"\(\s*(\d{4})\s*,\s*(\d{1,2})월(?:\s*(\d{1,2})일)?\s*\)",
               fix_date, s)

    # 한글 저자: '저자. (연도)' / '저자 (연도)' → '저자(연도)'
    m = re.match(r"^(?P<a>.+?)\s*\.?\s*\((?P<y>[^()]*\d{4}[^()]*)\)", s)
    if m and HANGUL.search(m.group("a")):
        head = m.group("a").rstrip(" .")
        s = f"{head}({m.group('y')})" + s[m.end():]

    # 'https://....../gp에서 2026.9.14. 인출' — 주소 뒤에 조사가 바로 붙으면
    # 주소가 어디서 끝나는지 보이지 않고, 독자가 주소만 떼어 쓸 수도 없다.
    # 한 칸 띄운다. (이미 띄어져 있으면 걸리지 않는다)
    s = re.sub(r"(https?://\S+?)에서", r"\1 에서", s)

    # 빈 URL/출처 자리 정리:  '성료. .' → '성료.'
    s = re.sub(r"\.\s+\.(\s|$)", r".\1", s)
    s = re.sub(r"\s{2,}", " ", s).strip()
    if not s.endswith((".", "쪽.", ")")):
        s += "."
    return s


def sort_key(entry):
    """국문 먼저(가나다), 영문 뒤(알파벳).

    같은 저자의 여러 글은 날짜순으로 놓이도록 숫자를 자리맞춤해 비교한다.
    ('2024. 12.'가 '2024. 6.'보다 앞서는 문자열 비교를 막는다.)
    """
    first = entry.strip()[:1]
    padded = re.sub(r"\d+", lambda m: m.group().zfill(4), entry)
    return (0 if HANGUL.match(first) else 1, padded)


def replace_section(body, entries):
    """참고문헌 절을 주어진 항목들로 통째로 갈아끼운다.

    PI가 확인해 준 확정본을 넣을 때 쓴다. 자동 정리를 거치지 않는다.
    """
    m = HEADING.search(body)
    before = body[:m.start()].rstrip() if m else body.rstrip()
    items = [l.strip() for l in entries.strip().split("\n") if l.strip()]
    return before + "\n\n## 참고문헌\n\n" + "\n\n".join(items) + "\n"


def normalize_section(body, drop=None):
    """본문에서 참고문헌 절을 찾아 통일한다.

    drop: 이 문구가 들어 있는 항목은 빼라고 지정된 것들(PI 확인 사항).
    (새 본문, 리포트) 를 돌려준다. 참고문헌이 없으면 원본 그대로.
    """
    drop = drop or []
    m = HEADING.search(body)
    if not m:
        return body, None

    before = body[:m.start()].rstrip()
    tail = body[m.end():]

    raw_lines = [l.strip() for l in tail.split("\n") if l.strip()]

    entries, dropped = [], []
    for line in raw_lines:
        for part in split_jammed(line):
            if _is_reference(part):
                entries.append(normalize_entry(part))
            else:
                dropped.append(part)

    # 완전히 같은 항목만 제거 (앞선 것 유지)
    seen, unique, dupes = set(), [], []
    for e in entries:
        if e in seen:
            dupes.append(e)
            continue
        seen.add(e)
        unique.append(e)

    # 빼기로 지정된 항목 제거
    removed = [e for e in unique if any(d in e for d in drop)]
    unique = [e for e in unique if e not in removed]

    unique.sort(key=sort_key)

    new_body = before + "\n\n## 참고문헌\n\n" + "\n\n".join(unique) + "\n"
    report = {"count": len(unique), "dropped": dropped,
              "dupes": dupes, "removed": removed}
    return new_body, report
