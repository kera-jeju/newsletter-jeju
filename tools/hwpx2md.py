# -*- coding: utf-8 -*-
"""hwpx → 마크다운 + 이미지 추출.

hwpx는 zip이고 Contents/section*.xml에 본문이 들어 있다.
표(tbl)는 마크다운 표로, BinData의 이미지는 파일로 뽑는다.
"""
import re
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET


def _local(tag):
    return tag.rsplit('}', 1)[-1]


def _para_text(p_el):
    """문단 하나의 텍스트.

    표(tbl) 하위는 타고 들어가지 않는다. 표를 감싼 문단이 셀 텍스트까지
    끌어와 본문에 한 줄로 쏟아지는 것을 막는다(표는 _table_to_md가 따로 만든다).
    """
    buf = []

    def walk(el):
        for ch in el:
            name = _local(ch.tag)
            if name == 'tbl':
                continue
            if name == 't' and ch.text:
                buf.append(ch.text)
            walk(ch)

    walk(p_el)
    return "".join(buf).strip()


def _is_inside_table(p_el, table_paras):
    return id(p_el) in table_paras


def _quote_para_ids(header_xml):
    """양쪽(왼쪽·오른쪽)을 모두 들여쓴 문단모양 id 집합.

    한글에서 직접인용을 표시하는 관행적인 서식이다 — 본문보다 양쪽을 좁혀
    블록으로 앉힌다. 한쪽만 들여쓴 것(참고문헌의 내어쓰기 등)은 인용이 아니므로
    양쪽이 모두 0보다 클 때만 센다. 단위는 HWPUNIT(1/7200인치).
    """
    ids = set()
    root = ET.fromstring(header_xml)
    for el in root.iter():
        if _local(el.tag) != 'paraPr':
            continue
        left = right = 0
        for m in el.iter():
            if _local(m.tag) != 'margin':
                continue
            for mm in m:
                name = _local(mm.tag)
                try:
                    v = int(mm.get('value', '0'))
                except ValueError:
                    continue
                if name == 'left':
                    left = max(left, v)
                elif name == 'right':
                    right = max(right, v)
        if left > 0 and right > 0:
            ids.add(el.get('id'))
    return ids


def _collect_table_paras(root):
    """표 안에 들어 있는 문단 id 집합 (본문 순회 시 중복 방지)."""
    inside = set()
    for tbl in root.iter():
        if _local(tbl.tag) != 'tbl':
            continue
        for p in tbl.iter():
            if _local(p.tag) == 'p':
                inside.add(id(p))
    return inside


def _table_to_md(tbl):
    rows = []
    for tr in tbl:
        if _local(tr.tag) != 'tr':
            continue
        cells = []
        for tc in tr:
            if _local(tc.tag) != 'tc':
                continue
            texts = []
            for p in tc.iter():
                if _local(p.tag) == 'p':
                    t = _para_text(p)
                    if t:
                        texts.append(t)
            cells.append(" ".join(texts).replace("|", "/"))
        if cells:
            rows.append(cells)
    if not rows:
        return ""
    ncol = max(len(r) for r in rows)
    out = []
    head = rows[0] + [""] * (ncol - len(rows[0]))
    out.append("| " + " | ".join(head) + " |")
    out.append("|" + "---|" * ncol)
    for r in rows[1:]:
        r = r + [""] * (ncol - len(r))
        out.append("| " + " | ".join(r) + " |")
    return "\n".join(out)


def convert(hwpx_path, out_dir, slug):
    hwpx_path = Path(hwpx_path)
    out_dir = Path(out_dir)
    img_dir = out_dir / "images"
    img_dir.mkdir(parents=True, exist_ok=True)

    z = zipfile.ZipFile(str(hwpx_path))

    # 이미지 추출
    images = []
    for n in z.namelist():
        if n.startswith("BinData/"):
            ext = Path(n).suffix.lower() or ".png"
            fname = f"{slug}_{len(images)+1:02d}{ext}"
            (img_dir / fname).write_bytes(z.read(n))
            images.append(f"images/{fname}")

    # 필자가 양쪽 들여쓰기로 표시한 직접인용을 찾아 둔다 (아래에서 > 로 옮긴다)
    try:
        quote_ids = _quote_para_ids(z.read("Contents/header.xml").decode("utf-8"))
    except KeyError:
        quote_ids = set()

    sections = sorted(n for n in z.namelist()
                      if re.match(r"Contents/section\d+\.xml$", n))
    lines = []
    for sname in sections:
        root = ET.fromstring(z.read(sname).decode("utf-8"))
        in_table = _collect_table_paras(root)

        # 본문 순서대로: p와 tbl을 훑되, 표 안 문단은 건너뛴다
        for el in root.iter():
            tag = _local(el.tag)
            if tag == 'tbl':
                md = _table_to_md(el)
                if md:
                    lines.append(md)
                    lines.append("")
            elif tag == 'p' and not _is_inside_table(el, in_table):
                t = _para_text(el)
                if t and el.get('paraPrIDRef') in quote_ids:
                    # 앞뒤에 빈 줄을 두어야 본문 문단에 붙지 않는다
                    lines.append("")
                    lines.append("> " + t)
                    lines.append("")
                else:
                    lines.append(t if t else "")

    md = "\n".join(lines)
    md = re.sub(r"\n{3,}", "\n\n", md).strip() + "\n"
    return md, images
