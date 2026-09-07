"""生成済み pptx を読み取り、各スライドをプレビュー画像として描き出す。

この環境には LibreOffice が無いため、目視QA用に pptx の中身
（図形の座標・塗り・テキストと書式）をそのまま再現して描画する。
PowerPoint そのものの描画ではないが、はみ出し・重なり・余白の確認には足りる。
"""

import os
import sys
import zipfile

from PIL import Image, ImageDraw, ImageFilter, ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

EMU_IN = 914400.0
PPI = 120
FONT_DIR = ".fonts"
OUT_DIR = "build/preview"  # main() で第2引数があれば差し替える

_font_cache = {}


def font(size_pt, bold=False):
    key = (round(size_pt * 2), bold)
    if key not in _font_cache:
        name = "NotoSansCJKjp-Bold.otf" if bold else "NotoSansCJKjp-Regular.otf"
        _font_cache[key] = ImageFont.truetype(
            os.path.join(FONT_DIR, name), max(1, int(round(size_pt * PPI / 72.0)))
        )
    return _font_cache[key]


def px(emu):
    return emu / EMU_IN * PPI


def rgb(color):
    try:
        return tuple(int(str(color.rgb)[i : i + 2], 16) for i in (0, 2, 4))
    except Exception:
        return None


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def slide_bg_image(slide):
    bg = slide._element.find("p:cSld/p:bg", NS)
    if bg is None:
        return None
    blip = bg.find(".//a:blip", NS)
    if blip is None:
        return None
    rid = blip.get("{%s}embed" % NS["r"])
    try:
        return slide.part.related_part(rid).blob
    except Exception:
        return None


def wrap(text, fnt, max_w):
    """CJK は文字単位、ラテンは単語単位で折り返す。"""
    lines = []
    for para in text.split("\n"):
        if not para:
            lines.append("")
            continue
        cur = ""
        i = 0
        while i < len(para):
            ch = para[i]
            chunk = ch
            if ch.isascii() and (ch.isalnum() or ch in "@.-_:/"):
                j = i
                while j < len(para) and para[j].isascii() and (para[j].isalnum() or para[j] in "@.-_:/"):
                    j += 1
                chunk = para[i:j]
            if cur and fnt.getlength(cur + chunk) > max_w:
                lines.append(cur)
                cur = chunk.lstrip() if chunk.strip() else ""
            else:
                cur += chunk
            i += len(chunk)
        lines.append(cur)
    return lines


def draw_text_frame(img, draw, sh, overflow):
    tf = sh.text_frame
    body = sh._element.find(".//a:bodyPr", NS)
    ins = {"l": 91440, "t": 45720, "r": 91440, "b": 45720}
    if body is not None:
        for k, attr in (("l", "lIns"), ("t", "tIns"), ("r", "rIns"), ("b", "bIns")):
            v = body.get(attr)
            if v is not None:
                ins[k] = int(v)
    anchor = body.get("anchor") if body is not None else None

    x0 = px(sh.left + ins["l"])
    y0 = px(sh.top + ins["t"])
    box_w = px(sh.width - ins["l"] - ins["r"])
    box_h = px(sh.height - ins["t"] - ins["b"])

    blocks = []
    for para in tf.paragraphs:
        pPr = para._p.find("a:pPr", NS)
        runs = [r for r in para.runs if r.text]
        if not runs:
            blocks.append(None)
            continue
        r0 = runs[0]
        size = r0.font.size.pt if r0.font.size else 18
        bold = bool(r0.font.bold)
        color = rgb(r0.font.color) or (0, 0, 0)
        fnt = font(size, bold)

        line_h = size * 1.36
        space_after = 0
        bullet = None
        mar_l = 0
        align = para.alignment
        if pPr is not None:
            sp = pPr.find("a:lnSpc/a:spcPts", NS)
            if sp is not None:
                line_h = int(sp.get("val")) / 100.0
            sa = pPr.find("a:spcAft/a:spcPts", NS)
            if sa is not None:
                space_after = int(sa.get("val")) / 100.0
            bu = pPr.find("a:buChar", NS)
            if bu is not None:
                bullet = bu.get("char")
            if pPr.get("marL"):
                mar_l = px(int(pPr.get("marL")))

        text = "".join(r.text for r in runs)
        avail = box_w - mar_l
        lines = wrap(text, fnt, avail)
        blocks.append(
            {
                "lines": lines,
                "fnt": fnt,
                "color": color,
                "lh": line_h * PPI / 72.0,
                "sa": space_after * PPI / 72.0,
                "bullet": bullet,
                "marL": mar_l,
                "align": str(align) if align else None,
                "size": size,
            }
        )

    total = 0
    for b in blocks:
        if b is None:
            continue
        total += len(b["lines"]) * b["lh"] + b["sa"]

    y = y0
    if anchor == "ctr":
        y = y0 + max(0, (box_h - total) / 2)
    elif anchor == "b":
        y = y0 + max(0, box_h - total)

    if total > box_h + 3:
        overflow.append(
            (sh, total / PPI * 72.0, box_h / PPI * 72.0, tf.text[:40])
        )

    for b in blocks:
        if b is None:
            continue
        for i, line in enumerate(b["lines"]):
            tx = x0 + b["marL"]
            if b["align"] and "CENTER" in b["align"]:
                tx = x0 + (box_w - b["fnt"].getlength(line)) / 2
            elif b["align"] and "RIGHT" in b["align"]:
                tx = x0 + box_w - b["fnt"].getlength(line)
            ty = y + (b["lh"] - b["fnt"].size) / 2
            if b["bullet"] and i == 0:
                draw.text(
                    (x0, ty), b["bullet"], font=b["fnt"], fill=b["color"], anchor="la"
                )
            draw.text((tx, ty), line, font=b["fnt"], fill=b["color"], anchor="la")
            y += b["lh"]
        y += b["sa"]


def round_radius(sh, w, h):
    gd = sh._element.find(".//a:prstGeom/a:avLst/a:gd", NS)
    if gd is None:
        return 0
    val = gd.get("fmla", "")
    if "val" in val:
        try:
            adj = int(val.split()[-1])
        except ValueError:
            return 0
        return adj / 100000.0 * min(w, h)
    return 0


def main(path):
    os.makedirs(OUT_DIR, exist_ok=True)
    prs = Presentation(path)
    W = int(px(prs.slide_width))
    H = int(px(prs.slide_height))
    all_overflow = []

    for idx, slide in enumerate(prs.slides, 1):
        img = Image.new("RGB", (W, H), (255, 255, 255))
        blob = slide_bg_image(slide)
        if blob:
            import io

            bgimg = Image.open(io.BytesIO(blob)).convert("RGB").resize((W, H))
            img.paste(bgimg, (0, 0))

        shadow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        sdraw = ImageDraw.Draw(shadow)
        draw = ImageDraw.Draw(img)
        overflow = []

        for sh in slide.shapes:
            if sh.left is None:
                continue
            x, y = px(sh.left), px(sh.top)
            w, h = px(sh.width), px(sh.height)
            box = [x, y, x + w, y + h]

            if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                fill = None
                try:
                    if sh.fill.type is not None and sh.fill.type == 1:
                        fill = rgb(sh.fill.fore_color)
                except Exception:
                    pass
                geom = sh._element.find(".//a:prstGeom", NS)
                prst = geom.get("prst") if geom is not None else "rect"
                has_shadow = sh._element.find(".//a:outerShdw", NS) is not None
                if fill:
                    if prst == "ellipse":
                        if has_shadow:
                            sdraw.ellipse(
                                [box[0] + 2, box[1] + 3, box[2] + 2, box[3] + 3],
                                fill=(150, 110, 120, 60),
                            )
                        draw.ellipse(box, fill=fill)
                    else:
                        r = round_radius(sh, w, h)
                        if has_shadow:
                            sdraw.rounded_rectangle(
                                [box[0] + 2, box[1] + 4, box[2] + 2, box[3] + 4],
                                radius=r,
                                fill=(150, 110, 120, 55),
                            )
                        draw.rounded_rectangle(box, radius=r, fill=fill)

            if sh.has_text_frame and sh.text_frame.text.strip():
                pass

        img_s = Image.alpha_composite(
            img.convert("RGBA"), shadow.filter(ImageFilter.GaussianBlur(7))
        )
        img = img_s.convert("RGB")
        draw = ImageDraw.Draw(img)

        # 影を合成し直したので、図形とテキストをこの順で描き直す
        for sh in slide.shapes:
            if sh.left is None:
                continue
            x, y = px(sh.left), px(sh.top)
            w, h = px(sh.width), px(sh.height)
            box = [x, y, x + w, y + h]
            if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                fill = None
                try:
                    if sh.fill.type == 1:
                        fill = rgb(sh.fill.fore_color)
                except Exception:
                    pass
                geom = sh._element.find(".//a:prstGeom", NS)
                prst = geom.get("prst") if geom is not None else "rect"
                if fill:
                    if prst == "ellipse":
                        draw.ellipse(box, fill=fill)
                    else:
                        draw.rounded_rectangle(box, radius=round_radius(sh, w, h), fill=fill)
            if sh.has_text_frame and sh.text_frame.text.strip():
                draw_text_frame(img, draw, sh, overflow)

        out = f"{OUT_DIR}/slide-{idx:02d}.png"
        img.save(out)
        for o in overflow:
            all_overflow.append((idx,) + o[1:])
        print(out)

    if all_overflow:
        print("\n[はみ出しの可能性]")
        for idx, need, have, t in all_overflow:
            print(f" - S{idx}: 必要{need:.1f}pt > 枠{have:.1f}pt :: {t}…")
    else:
        print("\nテキストのはみ出しは検出されませんでした")


if __name__ == "__main__":
    if len(sys.argv) > 2:
        OUT_DIR = sys.argv[2]
    main(sys.argv[1])
