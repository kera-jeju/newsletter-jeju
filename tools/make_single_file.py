# -*- coding: utf-8 -*-
"""검토본 만들기 — 이미지를 안에 품은 HTML 한 파일.

임원진 회람용. 파일 하나만 보내면 되고, 인터넷에 아무것도 올라가지 않는다.
받는 사람은 첨부를 브라우저로 열기만 하면 실제 지면 그대로 본다.

  python tools/make_single_file.py            # 2026 → 2026-검토본.html
  python tools/make_single_file.py --max-px 800 --quality 72   # 더 줄이기
"""
import argparse
import base64
import io
import re
from pathlib import Path

from PIL import Image

REPO = Path(__file__).resolve().parent.parent


def shrink(path, max_px, quality):
    """이미지를 줄여 (mime, bytes)로 돌려준다. 투명도가 있으면 PNG로 유지."""
    im = Image.open(path)
    has_alpha = im.mode in ("RGBA", "LA") or (
        im.mode == "P" and "transparency" in im.info)

    w, h = im.size
    if max(w, h) > max_px:
        scale = max_px / max(w, h)
        im = im.resize((max(1, int(w * scale)), max(1, int(h * scale))),
                       Image.LANCZOS)

    buf = io.BytesIO()
    if has_alpha:
        im.convert("RGBA").save(buf, format="PNG", optimize=True)
        return "image/png", buf.getvalue()
    im.convert("RGB").save(buf, format="JPEG", quality=quality, optimize=True)
    return "image/jpeg", buf.getvalue()


def main():
    ap = argparse.ArgumentParser(description="검토본 단일 HTML 만들기")
    ap.add_argument("--src", default="2026", help="빌드된 호 폴더")
    ap.add_argument("--out", default=None, help="출력 파일")
    ap.add_argument("--max-px", type=int, default=1000, help="이미지 긴 변 최대")
    ap.add_argument("--quality", type=int, default=78, help="JPEG 품질")
    args = ap.parse_args()

    src_dir = REPO / args.src
    html = (src_dir / "index.html").read_text(encoding="utf-8")
    out = Path(args.out) if args.out else REPO / f"{args.src}-검토본.html"

    cache, stats = {}, {"embedded": 0, "missing": 0, "before": 0, "after": 0}

    def to_data_uri(rel):
        if rel in cache:
            return cache[rel]
        path = src_dir / rel
        if not path.exists():
            stats["missing"] += 1
            cache[rel] = rel
            return rel
        stats["before"] += path.stat().st_size
        mime, data = shrink(path, args.max_px, args.quality)
        stats["after"] += len(data)
        stats["embedded"] += 1
        uri = f"data:{mime};base64,{base64.b64encode(data).decode('ascii')}"
        cache[rel] = uri
        return uri

    # src="images/..." 와 href="favicon.ico" 를 모두 품는다
    def repl(m):
        attr, quote, rel = m.group(1), m.group(2), m.group(3)
        return f'{attr}={quote}{to_data_uri(rel)}{quote}'

    html = re.sub(r'(src|href)=(["\'])((?:images/)[^"\']+)\2', repl, html)
    # 파비콘은 회람본에 필요 없다 — 깨진 링크로 남지 않게 지운다
    html = re.sub(r'<link[^>]+rel="(?:icon|apple-touch-icon)"[^>]*>\s*', '', html)

    # 검토본임을 알리는 띠
    banner = (
        '<div style="background:#8a1c1c;color:#fff;font-family:sans-serif;'
        'font-size:0.92rem;padding:0.7rem 1rem;text-align:center;'
        'letter-spacing:0.01em;">'
        '검토본입니다 · 아직 발간되지 않았습니다 · 외부 공유를 삼가 주십시오'
        '</div>')
    html = re.sub(r'(<body[^>]*>)', r'\1' + banner, html, count=1)

    out.write_text(html, encoding="utf-8")
    size = out.stat().st_size / 1024 / 1024
    print(f"이미지 {stats['embedded']}개 삽입 "
          f"({stats['before']/1024/1024:.1f}MB → {stats['after']/1024/1024:.1f}MB)")
    if stats["missing"]:
        print(f"  ! 찾지 못한 이미지 {stats['missing']}개")
    print(f"검토본: {out}  ({size:.1f} MB)")


if __name__ == "__main__":
    main()
