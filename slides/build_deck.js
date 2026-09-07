// 研究説明スライド（同意説明文書 02_同意説明文書_260710.docx に基づく）
const pptxgen = require("pptxgenjs");
const {
  C,
  PALETTE,
  FONT,
  T,
  PAGE,
  bg,
  card,
  badge,
  slideTitle,
  bullets,
  footnote,
  pageNumber,
} = require("./theme");

const TOTAL = 12;
const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "横浜市立大学医学部 公衆衛生学教室";
pptx.title = "研究の説明と協力のお願い";

const txt = (slide, text, o) =>
  slide.addText(text, Object.assign({ margin: 0, fontFace: FONT, valign: "top" }, o));

// ---------------------------------------------------------------- 1 表紙
{
  const s = pptx.addSlide();
  bg(s, "title");

  txt(s, "この研究の対象者の皆様へ", {
    x: 1.35,
    y: 1.5,
    w: 8,
    h: 0.36,
    fontSize: 14.5,
    bold: true,
    color: C.accent,
    charSpacing: 1.4,
    valign: "middle",
  });
  txt(s, "研究の説明と協力のお願い", {
    x: 1.35,
    y: 1.98,
    w: 10.8,
    h: 0.95,
    fontSize: T.deckTitle,
    bold: true,
    color: C.ink,
    charSpacing: 1.2,
    valign: "middle",
  });
  txt(s, "企業の構造的・文化的要因と働く男女の健康の関連に関する疫学研究", {
    x: 1.35,
    y: 3.05,
    w: 10.8,
    h: 0.4,
    fontSize: 15.5,
    color: C.inkSoft,
    valign: "middle",
  });

  card(pptx, s, { x: 1.35, y: 4.05, w: 10.6, h: 1.75, fill: C.white });
  const col = [
    ["研究代表機関", "横浜市立大学医学部\n公衆衛生学教室", 1.85, 3.1],
    ["研究責任者", "荒川 裕貴", 5.15, 2.2],
    ["連携機関", "横浜市 健康福祉局健康推進課\n協会けんぽ神奈川支部", 7.55, 3.9],
  ];
  col.forEach(([label, value, x, w]) => {
    txt(s, label, {
      x,
      y: 4.4,
      w,
      h: 0.28,
      fontSize: 11,
      bold: true,
      color: C.accent,
      charSpacing: 0.8,
    });
    txt(s, value, { x, y: 4.75, w, h: 0.85, fontSize: 13, color: C.ink, lineSpacing: 19 });
  });
}

// ---------------------------------------------------------------- 2 背景
{
  const s = pptx.addSlide();
  bg(s, "body");
  slideTitle(s, "なぜ、この研究を行うのか", "働く人の健康と、その人が働く会社の環境との関係を明らかにします");

  txt(
    s,
    "働く男女のメンタルヘルスや、女性の月経困難症・更年期障害などの健康課題、孤独感などの社会的な課題は、仕事や生活に大きな影響を与えることが知られています。",
    { x: PAGE.m, y: 1.95, w: 5.5, h: 1.5, fontSize: T.body, color: C.ink, lineSpacing: 23 },
  );
  txt(
    s,
    "一方で、勤務先の制度・体制・職場の風土といった要因が、これらの健康にどのような影響を与えるかは、まだ十分に解明されていません。",
    { x: PAGE.m, y: 3.35, w: 5.5, h: 1.3, fontSize: T.body, color: C.ink, lineSpacing: 23 },
  );

  card(pptx, s, { x: 6.6, y: 1.85, w: 6.02, h: 2.25, fill: C.accentTint });
  badge(pptx, s, { x: 7.0, y: 2.2, d: 0.34, fill: C.accentSoft, text: "" });
  txt(s, "見過ごされがちな健康課題", {
    x: 7.5,
    y: 2.19,
    w: 4.7,
    h: 0.36,
    fontSize: T.cardHead,
    bold: true,
    color: C.ink,
    valign: "middle",
  });
  txt(
    s,
    bullets(["月経困難症・更年期症状などの体の不調", "うつなどの心の不調", "職場での孤独感・孤立感"]),
    { x: 7.0, y: 2.72, w: 5.25, h: 1.15, fontSize: T.body, color: C.ink, lineSpacing: 20 },
  );

  card(pptx, s, { x: 6.6, y: 4.35, w: 6.02, h: 2.25, fill: C.altTint });
  badge(pptx, s, { x: 7.0, y: 4.7, d: 0.34, fill: C.altSoft, text: "" });
  txt(s, "まだ分かっていないこと", {
    x: 7.5,
    y: 4.69,
    w: 4.7,
    h: 0.36,
    fontSize: T.cardHead,
    bold: true,
    color: C.ink,
    valign: "middle",
  });
  txt(
    s,
    "会社の制度や体制、職場の雰囲気といった「働く環境」が、一人ひとりの健康とどう結びついているのか。その関係はまだ科学的に十分示されていません。",
    { x: 7.0, y: 5.22, w: 5.25, h: 1.2, fontSize: T.body, color: C.ink, lineSpacing: 21 },
  );

  pageNumber(s, 2, TOTAL);
}

// ---------------------------------------------------------------- 3 目的
{
  const s = pptx.addSlide();
  bg(s, "accent");
  slideTitle(s, "研究の目的と意義", "データを用いて、働く環境と健康のつながりを明らかにすることを目指します");

  const items = [
    ["1", "関連を明らかにする", "就労先企業の機能的・構造的・文化的な要因が、心・体・社会的な健康にどう影響するかをデータで解明します。", C.accent, C.accentTint],
    ["2", "根拠をつくる", "企業や保険者が健康づくりの取り組みを検討・改善する際に使える、科学的な根拠を示します。", C.alt, C.altTint],
    ["3", "健康づくりに活かす", "働く女性、そして働く男女すべての健康づくりに役立つ、具体的な手がかりにつなげます。", C.accent, C.accentTint],
  ];
  const w = 3.698;
  items.forEach(([n, head, body, accent, fill], i) => {
    const x = PAGE.m + i * (w + 0.4);
    card(pptx, s, { x, y: 2.0, w, h: 3.4, fill });
    badge(pptx, s, { x: x + 0.42, y: 2.42, d: 0.68, fill: accent, text: n, size: 19 });
    txt(s, head, {
      x: x + 0.42,
      y: 3.32,
      w: w - 0.84,
      h: 0.4,
      fontSize: T.cardHead,
      bold: true,
      color: C.ink,
      valign: "middle",
    });
    txt(s, body, {
      x: x + 0.42,
      y: 3.85,
      w: w - 0.84,
      h: 1.4,
      fontSize: T.body,
      color: C.ink,
      lineSpacing: 21,
    });
  });

  footnote(s, "研究成果は、企業・保険者が働く人の健康づくりを進めるための判断材料として活用されることが期待されます。", 5.75);
  pageNumber(s, 3, TOTAL);
}

// ---------------------------------------------------------------- 4 研究の概要
{
  const s = pptx.addSlide();
  bg(s, "body");
  slideTitle(s, "研究の概要", "本研究は倫理審査委員会の承認と、研究機関の長の許可を得て実施します");

  const cells = [
    ["研究の名称", "企業の構造的・文化的要因と働く女性の健康の関連に関する疫学研究"],
    ["実施体制", "横浜市立大学を研究代表機関とし、横浜市（健康福祉局健康推進課）および協会けんぽ神奈川支部と連携して実施します。"],
    ["倫理審査と許可", "横浜市立大学「人を対象とする生命科学・医学系研究倫理委員会」の審査・承認を受け、研究機関の長の許可を得ています。"],
    ["研究期間", "研究期間：実施許可日 〜 2031年3月31日\n予定登録期間：実施許可日 〜 2028年3月31日"],
  ];
  const cw = 5.746;
  cells.forEach(([label, value], i) => {
    const x = PAGE.m + (i % 2) * (cw + 0.4);
    const y = 1.95 + Math.floor(i / 2) * 2.4;
    card(pptx, s, { x, y, w: cw, h: 2.15, fill: C.white });
    badge(pptx, s, { x: x + 0.42, y: y + 0.42, d: 0.3, fill: i % 2 ? C.altSoft : C.accentSoft, text: "" });
    txt(s, label, {
      x: x + 0.86,
      y: y + 0.4,
      w: cw - 1.3,
      h: 0.34,
      fontSize: T.cardHead,
      bold: true,
      color: C.ink,
      valign: "middle",
    });
    txt(s, value, {
      x: x + 0.42,
      y: y + 0.95,
      w: cw - 0.84,
      h: 1.0,
      fontSize: T.body,
      color: C.ink,
      lineSpacing: 21,
    });
  });

  pageNumber(s, 4, TOTAL);
}

// ---------------------------------------------------------------- 5 対象となる方
{
  const s = pptx.addSlide();
  bg(s, "body");
  slideTitle(s, "対象となる方", "事業所と、そこで働く方の双方にご協力をお願いしています");

  const cw = 5.746;
  const groups = [
    [
      "事業所",
      C.accentTint,
      C.accent,
      ["協会けんぽ神奈川支部に加入する事業所", "説明会等を通じて参加を希望した、日本に所在地を持つ事業所"],
      "ご回答は、人事・労務・健康づくりのご担当者にお願いします。",
    ],
    [
      "働く方（労働者）",
      C.altTint,
      C.alt,
      ["日本に所在地を持つ事業所で働いている方", "回答時点で18歳以上74歳以下の方", "本調査への参加に同意いただける方"],
      "性別を問わず、どなたでもご参加いただけます。",
    ],
  ];
  groups.forEach(([head, fill, accent, list, note], i) => {
    const x = PAGE.m + i * (cw + 0.4);
    card(pptx, s, { x, y: 1.9, w: cw, h: 4.3, fill });
    badge(pptx, s, { x: x + 0.45, y: 2.3, d: 0.62, fill: accent, text: String(i + 1), size: 17 });
    txt(s, head, {
      x: x + 1.2,
      y: 2.32,
      w: cw - 1.6,
      h: 0.58,
      fontSize: 19,
      bold: true,
      color: C.ink,
      valign: "middle",
    });
    txt(s, bullets(list, { gap: 10 }), {
      x: x + 0.45,
      y: 3.2,
      w: cw - 0.9,
      h: 1.9,
      fontSize: T.body,
      color: C.ink,
      lineSpacing: 21,
    });
    txt(s, note, {
      x: x + 0.45,
      y: 5.35,
      w: cw - 0.9,
      h: 0.6,
      fontSize: T.caption,
      color: C.inkSoft,
      lineSpacing: 17,
    });
  });

  pageNumber(s, 5, TOTAL);
}

// ---------------------------------------------------------------- 6 調査の流れ
{
  const s = pptx.addSlide();
  bg(s, "accent");
  slideTitle(s, "調査の方法と流れ", "オンライン質問票（ウェブアンケート）にご回答いただきます");

  const steps = [
    ["ご案内", "研究用ホームページで、本説明事項をご確認いただきます。"],
    ["同意", "内容にご納得いただけたら、画面上で同意の手続きを行います。"],
    ["ご回答", "オンライン質問票にご回答いただきます。所要時間は約20分です。"],
    ["結果のお返し", "全体の集計結果とご自身・自社の回答との比較をお返しします。"],
  ];
  const w = 2.748;
  steps.forEach(([head, body], i) => {
    const x = PAGE.m + i * (w + 0.3);
    card(pptx, s, { x, y: 2.05, w, h: 2.75, fill: C.white });
    badge(pptx, s, { x: x + 0.4, y: 2.42, d: 0.6, fill: i === 2 ? C.accent : C.accentSoft, text: String(i + 1), size: 16 });
    txt(s, head, {
      x: x + 0.4,
      y: 3.16,
      w: w - 0.8,
      h: 0.36,
      fontSize: 15,
      bold: true,
      color: C.ink,
      valign: "middle",
    });
    txt(s, body, {
      x: x + 0.4,
      y: 3.6,
      w: w - 0.8,
      h: 1.0,
      fontSize: 12,
      color: C.ink,
      lineSpacing: 19,
    });
  });

  card(pptx, s, { x: PAGE.m, y: 5.15, w: PAGE.w - PAGE.m * 2, h: 1.15, fill: C.altTint });
  badge(pptx, s, { x: PAGE.m + 0.42, y: 5.53, d: 0.38, fill: C.altSoft, text: "" });
  txt(
    s,
    "血液などの生体試料は扱いません。ご回答いただいた内容は、研究対象者識別コードを付けた研究用データとして取り扱います。",
    {
      x: PAGE.m + 1.0,
      y: 5.42,
      w: PAGE.w - PAGE.m * 2 - 1.5,
      h: 0.6,
      fontSize: T.body,
      color: C.ink,
      valign: "middle",
      lineSpacing: 21,
    },
  );

  footnote(s, "サーバー構築や情報セキュリティ管理などの業務を外部に委託する場合がありますが、個人情報保護の規定に従い厳重に管理します。", 6.5);
  pageNumber(s, 6, TOTAL);
}

// ---------------------------------------------------------------- 7 質問の内容
{
  const s = pptx.addSlide();
  bg(s, "body");
  slideTitle(s, "おうかがいする内容", "所要時間の目安は約20分です（設問数により変動します）");

  const cw = 5.746;
  const cols = [
    [
      "事業所のご担当者へ",
      C.accentTint,
      C.accent,
      ["業種・従業員数などの基本情報", "健康づくりの体制", "休暇制度の内容と利用状況", "職員交流の取り組み", "性差に配慮した取り組み"],
    ],
    [
      "働く方へ",
      C.altTint,
      C.alt,
      ["年齢・性別・勤務形態などの背景", "生活習慣", "メンタルヘルス", "孤独感、労働生産性", "女性の方には月経困難症・更年期症状に関する質問"],
    ],
  ];
  cols.forEach(([head, fill, accent, list], i) => {
    const x = PAGE.m + i * (cw + 0.4);
    card(pptx, s, { x, y: 1.9, w: cw, h: 3.85, fill });
    badge(pptx, s, { x: x + 0.45, y: 2.3, d: 0.34, fill: accent, text: "" });
    txt(s, head, {
      x: x + 0.95,
      y: 2.28,
      w: cw - 1.4,
      h: 0.4,
      fontSize: 17,
      bold: true,
      color: C.ink,
      valign: "middle",
    });
    txt(s, bullets(list, { gap: 9 }), {
      x: x + 0.45,
      y: 2.9,
      w: cw - 0.9,
      h: 2.6,
      fontSize: T.body,
      color: C.ink,
      lineSpacing: 21,
    });
  });

  footnote(s, "答えたくない質問には、回答いただかなくて構いません。ご回答は自由意思に基づくものです。", 6.05);
  pageNumber(s, 7, TOTAL);
}

// ---------------------------------------------------------------- 8 利益と負担
{
  const s = pptx.addSlide();
  bg(s, "body");
  slideTitle(s, "ご参加による利益と負担", "予測される利益と、負担・リスクの両方をお伝えします");

  const cw = 5.746;
  card(pptx, s, { x: PAGE.m, y: 1.9, w: cw, h: 4.05, fill: C.altTint });
  badge(pptx, s, { x: PAGE.m + 0.45, y: 2.28, d: 0.34, fill: C.altSoft, text: "" });
  txt(s, "期待される利益", {
    x: PAGE.m + 0.95,
    y: 2.26,
    w: cw - 1.4,
    h: 0.4,
    fontSize: 17,
    bold: true,
    color: C.ink,
    valign: "middle",
  });
  txt(
    s,
    bullets(
      [
        "回答を通じて、ご自身の健康状態を振り返る機会になります",
        "全体の集計値とご自身・自社の回答との比較をお返しします",
        "健康づくりに役立つ情報提供動画のURLをご案内します",
        "研究成果が、企業や保険者の健康づくりの改善につながる可能性があります",
      ],
      { gap: 10 },
    ),
    { x: PAGE.m + 0.45, y: 2.88, w: cw - 0.9, h: 2.9, fontSize: T.body, color: C.ink, lineSpacing: 21 },
  );

  const x2 = PAGE.m + cw + 0.4;
  card(pptx, s, { x: x2, y: 1.9, w: cw, h: 4.05, fill: C.accentTint });
  badge(pptx, s, { x: x2 + 0.45, y: 2.28, d: 0.34, fill: C.accentSoft, text: "" });
  txt(s, "予測される負担・リスク", {
    x: x2 + 0.95,
    y: 2.26,
    w: cw - 1.4,
    h: 0.4,
    fontSize: 17,
    bold: true,
    color: C.ink,
    valign: "middle",
  });
  txt(
    s,
    bullets(
      [
        "ご回答に20分ほどのお時間をいただきます",
        "月経・更年期・メンタルヘルスに関する質問を含むため、心理的な負担を感じる可能性があります",
        "個人情報の漏えいに対しては、識別コードの付与、対応表の厳重な管理、アクセス権限の限定などの措置を講じます",
      ],
      { gap: 10 },
    ),
    { x: x2 + 0.45, y: 2.88, w: cw - 0.9, h: 2.9, fontSize: T.body, color: C.ink, lineSpacing: 21 },
  );

  card(pptx, s, { x: PAGE.m, y: 6.05, w: PAGE.w - PAGE.m * 2, h: 0.85, fill: C.white });
  txt(
    s,
    "本研究はオンライン質問票による観察研究で、身体的な侵襲はありません。ご参加による費用の負担も生じません。",
    {
      x: PAGE.m + 0.45,
      y: 6.05,
      w: PAGE.w - PAGE.m * 2 - 0.9,
      h: 0.85,
      fontSize: T.body,
      color: C.ink,
      valign: "middle",
    },
  );

  pageNumber(s, 8, TOTAL);
}

// ---------------------------------------------------------------- 9 個人情報
{
  const s = pptx.addSlide();
  bg(s, "accent");
  slideTitle(s, "個人情報の取り扱い", "個人が特定されない形で、厳重に管理します");

  card(pptx, s, { x: PAGE.m, y: 1.9, w: PAGE.w - PAGE.m * 2, h: 3.55, fill: C.white });
  const rows = [
    ["識別コードで管理します", "ご回答には研究対象者識別コードを付けて管理し、氏名など個人を直接特定できる情報は研究用データと分けて保管します。"],
    ["対応表は外に出しません", "識別コードと個人情報を対応させる表は作成しますが、作成した研究機関の外へ提供することはありません。"],
    ["第三者へ提供しません", "個人を識別できる情報が、事業者・協会けんぽ・横浜市など、研究者以外の方に提供されることはありません。"],
  ];
  rows.forEach(([head, body], i) => {
    const y = 2.25 + i * 1.08;
    badge(pptx, s, { x: PAGE.m + 0.5, y: y + 0.06, d: 0.44, fill: i === 1 ? C.altSoft : C.accentSoft, text: "" });
    txt(s, head, {
      x: PAGE.m + 1.15,
      y,
      w: 3.6,
      h: 0.4,
      fontSize: 15,
      bold: true,
      color: C.ink,
      valign: "middle",
    });
    txt(s, body, {
      x: PAGE.m + 4.85,
      y: y - 0.03,
      w: PAGE.w - PAGE.m * 2 - 5.35,
      h: 0.9,
      fontSize: T.body,
      color: C.ink,
      lineSpacing: 21,
    });
  });

  card(pptx, s, { x: PAGE.m, y: 5.68, w: PAGE.w - PAGE.m * 2, h: 0.9, fill: C.accentTint });
  txt(
    s,
    "ウェブシステムへのアクセス権限は必要最小限の者に限定し、パスワード管理などの安全管理措置を講じています。",
    {
      x: PAGE.m + 0.45,
      y: 5.68,
      w: PAGE.w - PAGE.m * 2 - 0.9,
      h: 0.9,
      fontSize: T.body,
      color: C.ink,
      valign: "middle",
    },
  );

  pageNumber(s, 9, TOTAL);
}

// ---------------------------------------------------------------- 10 参加は自由
{
  const s = pptx.addSlide();
  bg(s, "body");
  slideTitle(s, "ご参加は自由意思です");

  card(pptx, s, { x: PAGE.m, y: 1.45, w: PAGE.w - PAGE.m * 2, h: 1.25, fill: C.accentTintDeep });
  txt(s, "参加を断っても、途中でやめても、不利益を受けることはありません。", {
    x: PAGE.m + 0.55,
    y: 1.45,
    w: PAGE.w - PAGE.m * 2 - 1.1,
    h: 1.25,
    fontSize: 21,
    bold: true,
    color: C.ink,
    valign: "middle",
    charSpacing: 0.5,
  });

  const w = 3.698;
  const items = [
    ["同意はいつでも撤回できます", "研究参加に同意した後でも、いつでも同意を撤回（参加をやめる）ことができます。", C.accentSoft],
    ["不利益はありません", "同意しない場合も、同意後に撤回した場合も、そのことで不利益を受けることはありません。", C.altSoft],
    ["撤回のご連絡先", "撤回をご希望の場合は、最後のページに記載のお問い合わせ先までご連絡ください。", C.accentSoft],
  ];
  items.forEach(([head, body, accent], i) => {
    const x = PAGE.m + i * (w + 0.4);
    card(pptx, s, { x, y: 3.15, w, h: 2.85, fill: C.white });
    badge(pptx, s, { x: x + 0.42, y: 3.55, d: 0.38, fill: accent, text: "" });
    txt(s, head, {
      x: x + 0.42,
      y: 4.14,
      w: w - 0.84,
      h: 0.38,
      fontSize: 15,
      bold: true,
      color: C.ink,
      valign: "middle",
    });
    txt(s, body, {
      x: x + 0.42,
      y: 4.62,
      w: w - 0.84,
      h: 1.15,
      fontSize: 12.5,
      color: C.ink,
      lineSpacing: 20,
    });
  });

  footnote(s, "ただし、撤回のお申し出の時点ですでに研究結果が公表されているなど、データから除外できない場合があります。", 6.28);
  pageNumber(s, 10, TOTAL);
}

// ---------------------------------------------------------------- 11 保管・二次利用・資金
{
  const s = pptx.addSlide();
  bg(s, "body");
  slideTitle(s, "情報の保管・二次利用と研究資金", "研究終了後の取り扱いについてもお知らせします");

  const cw = 5.746;
  card(pptx, s, { x: PAGE.m, y: 1.9, w: cw, h: 2.3, fill: C.accentTint });
  badge(pptx, s, { x: PAGE.m + 0.45, y: 2.26, d: 0.34, fill: C.accentSoft, text: "" });
  txt(s, "保管と廃棄", {
    x: PAGE.m + 0.95,
    y: 2.24,
    w: cw - 1.4,
    h: 0.38,
    fontSize: T.cardHead,
    bold: true,
    color: C.ink,
    valign: "middle",
  });
  txt(
    s,
    "研究終了の報告日から5年、または最終公表の報告日から3年のいずれか遅い日まで保管します。紙媒体はシュレッダーで、電子データは復元できない方法で消去します。",
    { x: PAGE.m + 0.45, y: 2.78, w: cw - 0.9, h: 1.25, fontSize: 12.5, color: C.ink, lineSpacing: 20 },
  );

  const x2 = PAGE.m + cw + 0.4;
  card(pptx, s, { x: x2, y: 1.9, w: cw, h: 2.3, fill: C.altTint });
  badge(pptx, s, { x: x2 + 0.45, y: 2.26, d: 0.34, fill: C.altSoft, text: "" });
  txt(s, "二次利用について", {
    x: x2 + 0.95,
    y: 2.24,
    w: cw - 1.4,
    h: 0.38,
    fontSize: T.cardHead,
    bold: true,
    color: C.ink,
    valign: "middle",
  });
  txt(
    s,
    "将来、他の研究に利用したり研究機関へ提供する場合があります。その際は新たに研究計画書を作成し、倫理審査委員会の承認と研究機関の長の許可を得たうえで、適切な手続を行います。",
    { x: x2 + 0.45, y: 2.78, w: cw - 0.9, h: 1.25, fontSize: 12.5, color: C.ink, lineSpacing: 20 },
  );

  card(pptx, s, { x: PAGE.m, y: 4.45, w: PAGE.w - PAGE.m * 2, h: 2.1, fill: C.white });
  badge(pptx, s, { x: PAGE.m + 0.45, y: 4.81, d: 0.34, fill: C.accentSoft, text: "" });
  txt(s, "研究資金・利益相反・情報公開", {
    x: PAGE.m + 0.95,
    y: 4.79,
    w: 6.0,
    h: 0.38,
    fontSize: T.cardHead,
    bold: true,
    color: C.ink,
    valign: "middle",
  });
  txt(
    s,
    bullets(
      [
        "日本学術振興会 科学研究費助成事業（若手研究）および横浜市受託研究の研究費により実施します",
        "研究成果に影響を及ぼすような利害関係（利益相反）はありません",
        "成果は学会発表や論文等で公表しますが、個人が特定される情報は公表しません",
      ],
      { gap: 6, size: 12.5 },
    ),
    { x: PAGE.m + 0.45, y: 5.3, w: PAGE.w - PAGE.m * 2 - 0.9, h: 1.15, fontSize: 12.5, color: C.ink, lineSpacing: 20 },
  );

  pageNumber(s, 11, TOTAL);
}

// ---------------------------------------------------------------- 12 手続きと連絡先
{
  const s = pptx.addSlide();
  bg(s, "closing");
  slideTitle(s, "ご参加の手続きとお問い合わせ先", "内容をよくご理解いただいたうえで、ご自身の意思でご判断ください");

  card(pptx, s, { x: PAGE.m, y: 1.9, w: 6.9, h: 4.4, fill: C.white });
  txt(s, "ご参加いただける場合の手順", {
    x: PAGE.m + 0.5,
    y: 2.22,
    w: 5.9,
    h: 0.4,
    fontSize: T.cardHead,
    bold: true,
    color: C.ink,
    valign: "middle",
  });
  const steps = [
    "元の画面に戻り、「上の項目を十分理解したうえで、研究への協力に同意します」にチェックを入れます。",
    "画面を進み、必要事項を入力します。\n・事業所担当者の方：氏名、役職、メールアドレス\n・働く方：生年月日、メールアドレス",
    "オンライン質問票にご回答・送信いただくと、本研究への参加に同意いただいたものとみなします（電磁的方法による同意）。",
  ];
  steps.forEach((body, i) => {
    const y = 2.82 + i * 1.12;
    badge(pptx, s, { x: PAGE.m + 0.5, y, d: 0.42, fill: C.accentSoft, text: String(i + 1), size: 13 });
    txt(s, body, {
      x: PAGE.m + 1.15,
      y: y - 0.04,
      w: 5.3,
      h: 1.0,
      fontSize: 12,
      color: C.ink,
      lineSpacing: 19,
    });
  });

  const x2 = PAGE.m + 7.3;
  card(pptx, s, { x: x2, y: 1.9, w: PAGE.w - PAGE.m - x2, h: 4.4, fill: C.accentTint });
  txt(s, "お問い合わせ先", {
    x: x2 + 0.5,
    y: 2.22,
    w: 3.5,
    h: 0.4,
    fontSize: T.cardHead,
    bold: true,
    color: C.ink,
    valign: "middle",
  });
  const contact = [
    ["機関名", "横浜市立大学医学部\n公衆衛生学教室"],
    ["電話", "045-787-2610"],
    ["メール", "arakawa.yuk.tq@yokohama-cu.ac.jp"],
    ["研究責任者", "荒川 裕貴"],
  ];
  let cy = 2.85;
  contact.forEach(([label, value]) => {
    const lines = value.split("\n").length;
    txt(s, label, {
      x: x2 + 0.5,
      y: cy,
      w: 3.6,
      h: 0.26,
      fontSize: 10.5,
      bold: true,
      color: C.accent,
      charSpacing: 0.6,
    });
    txt(s, value, {
      x: x2 + 0.5,
      y: cy + 0.28,
      w: 3.9,
      h: 0.28 * lines + 0.06,
      fontSize: 12.5,
      color: C.ink,
      lineSpacing: 18,
    });
    cy += 0.42 + 0.28 * lines;
  });

  footnote(s, "ご同意のあとも、研究用ホームページ上で本説明事項をいつでもご覧いただけます。", 6.55);
  pageNumber(s, 12, TOTAL);
}

const FILE_NAME = {
  pink: "研究の説明と協力のお願い.pptx",
  blue: "研究の説明と協力のお願い（事業所向け）.pptx",
}[PALETTE];

pptx.writeFile({ fileName: `build/${FILE_NAME}` }).then((f) => {
  console.log("written:", f);
});
