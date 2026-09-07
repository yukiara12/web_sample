// 配色トークンと共通パーツ。DECK_PALETTE=pink|blue で切り替える。

const PALETTES = {
  pink: {
    label: "ピンク",
    ink: "4A3C40", // 見出し・本文
    inkSoft: "806E73", // 補足・キャプション
    accent: "B85F7A", // 主アクセント（文字に使える濃さ）
    accentSoft: "E0A3B2", // 図形の装飾
    accentTint: "FCF1F4", // カード面
    accentTintDeep: "F7E3E9", // 一段濃いカード面
    alt: "5F8069", // 副アクセント（文字に使える濃さ）
    altSoft: "9BB6A2",
    altTint: "EFF4F0",
    pageNum: "B7A5AA",
    shadow: "C99AA6",
  },
  blue: {
    label: "ブルー",
    ink: "3A4450",
    inkSoft: "6E7B87",
    accent: "2F6B92",
    accentSoft: "9CC0D8",
    accentTint: "EFF6FB",
    accentTintDeep: "DEEAF4",
    alt: "3B7480",
    altSoft: "8DB8C2",
    altTint: "ECF4F6",
    pageNum: "A9B4BE",
    shadow: "9AAFC0",
  },
};

const PALETTE = process.env.DECK_PALETTE || "pink";
if (!PALETTES[PALETTE]) {
  throw new Error(`unknown palette: ${PALETTE}`);
}

const C = Object.assign({ white: "FFFFFF" }, PALETTES[PALETTE]);

const FONT = "Meiryo";

const T = {
  deckTitle: 40,
  title: 27,
  lead: 15,
  cardHead: 16,
  body: 13,
  small: 12,
  caption: 10.5,
  numeral: 30,
};

const PAGE = { w: 13.333, h: 7.5, m: 0.72 };

// pptxgenjs はオプションを内部で書き換えるため、毎回新しいオブジェクトを返す
const softShadow = (opacity = 0.22) => ({
  type: "outer",
  color: C.shadow,
  blur: 14,
  offset: 2,
  angle: 90,
  opacity,
});

function bg(slide, kind = "body") {
  slide.background = { path: `assets/${PALETTE}/bg-${kind}.png` };
}

function card(pptx, slide, o) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: o.x,
    y: o.y,
    w: o.w,
    h: o.h,
    fill: { color: o.fill || C.white },
    line: o.line ? { color: o.line, width: 1 } : { type: "none" },
    rectRadius: o.radius === undefined ? 0.14 : o.radius,
    shadow: o.shadow === false ? { type: "none" } : softShadow(o.shadowOpacity),
  });
}

// 見出し脇に置く小さな円。中の記号はフォント依存を避けて幾何記号のみ使う
function badge(pptx, slide, o) {
  slide.addShape(pptx.ShapeType.ellipse, {
    x: o.x,
    y: o.y,
    w: o.d,
    h: o.d,
    fill: { color: o.fill || C.accentSoft },
    line: { type: "none" },
    shadow: { type: "none" },
  });
  if (o.text !== undefined) {
    slide.addText(o.text, {
      x: o.x,
      y: o.y,
      w: o.d,
      h: o.d,
      align: "center",
      valign: "middle",
      margin: 0,
      fontFace: FONT,
      fontSize: o.size || 13,
      bold: true,
      color: o.color || C.white,
    });
  }
}

function slideTitle(slide, text, lead) {
  slide.addText(text, {
    x: PAGE.m,
    y: 0.52,
    w: PAGE.w - PAGE.m * 2,
    h: 0.62,
    margin: 0,
    fontFace: FONT,
    fontSize: T.title,
    bold: true,
    color: C.ink,
    charSpacing: 0.6,
    valign: "middle",
  });
  if (lead) {
    slide.addText(lead, {
      x: PAGE.m,
      y: 1.18,
      w: PAGE.w - PAGE.m * 2,
      h: 0.34,
      margin: 0,
      fontFace: FONT,
      fontSize: T.lead,
      color: C.inkSoft,
      valign: "middle",
    });
  }
}

function bullets(items, opts = {}) {
  return items.map((t, i) => ({
    text: t,
    options: {
      bullet: { code: "30FB" },
      breakLine: i < items.length - 1,
      paraSpaceAfter: opts.gap === undefined ? 7 : opts.gap,
      fontSize: opts.size || T.body,
      color: opts.color || C.ink,
      fontFace: FONT,
    },
  }));
}

function footnote(slide, text, y) {
  slide.addText(text, {
    x: PAGE.m,
    y: y || 6.82,
    w: PAGE.w - PAGE.m * 2,
    h: 0.32,
    margin: 0,
    fontFace: FONT,
    fontSize: T.caption,
    color: C.inkSoft,
    valign: "middle",
  });
}

function pageNumber(slide, n, total) {
  slide.addText(`${n} / ${total}`, {
    x: PAGE.w - 1.5,
    y: 6.96,
    w: 0.9,
    h: 0.3,
    margin: 0,
    align: "right",
    fontFace: FONT,
    fontSize: 9.5,
    color: C.pageNum,
    valign: "middle",
  });
}

module.exports = {
  C,
  PALETTE,
  FONT,
  T,
  PAGE,
  softShadow,
  bg,
  card,
  badge,
  slideTitle,
  bullets,
  footnote,
  pageNumber,
};
