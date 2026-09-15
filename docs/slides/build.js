const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";           // 13.333 x 7.5
pres.author = "米倉まな";
pres.title = "経絡に触れ、こころに耳を傾ける";
pres.subject = "APNET 講演 60分 / 鍼灸師のための傾聴トレーニング";

// ── ブランドトークン（オンラインサロン「ここちめいど」ロゴ由来）────────
// ロゴ実測：深緑 #1D8A53 / ミント #34D39B / ピンク #F15C9E
// 小さい文字でもコントラストが足りるよう、深緑とピンクは一段濃く調整している。
const CREAM = "F0FAF5", PAPER = "FFFFFF", INK = "123C2A";
const SHU = "17784A";    // 主アクセント＝深緑（どのサイズでも使える）
const MORI = "C92A78";   // 副アクセント＝ピンク（濃いめ。小さい文字でも可）
const PINKB = "F15C9E";  // ロゴのピンクそのまま。濃色スライドの見せ場だけ
const MINT = "34D39B";   // ロゴのミント。面で使い、文字は INK を乗せる
const MUTED = "4F6F60", MUTED_D = "8FB8A2", ONDARK = "E8F5EE";
const MIN = "游明朝", GO = "游ゴシック";     // 見出し=明朝 / 本文=ゴシック

const W = 13.333, H = 7.5, M = 0.8, CW = W - M * 2;
let n = 0;
const decks = [];

// スライドは一旦記録しておき、最後に縦方向を自動リフローしてから出力する。
// （13.3×7.5 のキャンバスに対して下部が空きすぎるのを防ぐ）
function S(dark) {
  n++;
  const rec = { dark, num: n, items: [], notes: null };
  decks.push(rec);
  return {
    addText: (t, o) => rec.items.push({ k: "text", t, o }),
    addShape: (st, o) => rec.items.push({ k: "shape", st, o }),
    addChart: (ct, d, o) => rec.items.push({ k: "chart", ct, d, o }),
    addNotes: (t) => { rec.notes = t; },
  };
}
const note = (s, t) => s.addNotes(t);

// 見出し（明朝・左寄せ）
function title(s, txt, dark, y) {
  s.addText(txt, {
    x: M, y: y === undefined ? 0.62 : y, w: CW, h: 1.05, isTextBox: true, margin: 0,
    fontFace: MIN, fontSize: 32, bold: true, color: dark ? ONDARK : INK,
    align: "left", valign: "middle", lineSpacing: 42,
  });
}
// 見出し上のラベル（章の位置を示す）
function eyebrow(s, txt, dark) {
  s.addText(txt, {
    x: M, y: 0.34, w: CW, h: 0.26, isTextBox: true, margin: 0,
    fontFace: GO, fontSize: 11, bold: true, charSpacing: 2,
    color: dark ? MUTED_D : SHU, align: "left",
  });
}
// 円モチーフ（番号・記号）
function circle(s, label, x, y, d, fill, txtColor, size) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill } });
  s.addText(label, {
    x, y, w: d, h: d, isTextBox: true, margin: 0, align: "center", valign: "middle",
    fontFace: MIN, fontSize: size || 20, bold: true, color: txtColor || PAPER,
  });
}
// カード
function card(s, x, y, w, h, fill) {
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.1, fill: { color: fill || PAPER },
    shadow: { type: "outer", angle: 90, blur: 10, offset: 0.04, color: "7FA795", opacity: 0.25 },
  });
}
function body(s, txt, x, y, w, h, opt) {
  s.addText(txt, Object.assign({
    x, y, w, h, isTextBox: true, margin: 0, fontFace: GO, fontSize: 15,
    color: INK, align: "left", valign: "top", lineSpacing: 26,
  }, opt || {}));
}

/* ══════════════════════════════════════════════════════════
   A. オープニング
   ══════════════════════════════════════════════════════════ */

// 1 タイトル
{
  const s = S(true);
  s.addShape(pres.ShapeType.ellipse, { x: 9.5, y: 1.1, w: 5.3, h: 5.3, fill: { color: "1B4E36" } });
  s.addText("APNET 講演", { x: M, y: 1.9, w: 8, h: 0.3, isTextBox: true, margin: 0, fontFace: GO, fontSize: 12, bold: true, charSpacing: 3, color: MUTED_D });
  s.addText("経絡に触れ、\nこころに耳を傾ける", {
    x: M, y: 2.4, w: 8.6, h: 2.0, isTextBox: true, margin: 0,
    fontFace: MIN, fontSize: 44, bold: true, color: ONDARK, lineSpacing: 62,
  });
  s.addText("鍼灸師のための傾聴トレーニング", { x: M, y: 4.5, w: 8.6, h: 0.4, isTextBox: true, margin: 0, fontFace: GO, fontSize: 17, color: MUTED_D });
  s.addText("米倉まな（よねくら まな）\n認定鍼灸師／はりきゅう処ここちめいど 院長\nオンラインサロン ここちめいど 主宰", {
    x: M, y: 5.25, w: 8.6, h: 1.1, isTextBox: true, margin: 0, fontFace: GO, fontSize: 12, color: MUTED_D, lineSpacing: 20,
  });
  note(s, "【0:00-0:20】\n開始前アナウンスで「紙とペンをご用意ください」を流してもらう。\nカメラオフでOKと最初に明言する。\n自分のカメラは必ずONにする。");
}

// 2 問いかけ
{
  const s = S(true);
  s.addText("最後に患者さんの話を、\n最後まで遮らずに聴いたのは\nいつですか？", {
    x: M, y: 1.9, w: CW, h: 3.2, isTextBox: true, margin: 0,
    fontFace: MIN, fontSize: 40, bold: true, color: ONDARK, lineSpacing: 66, align: "left",
  });
  s.addText("いまは答えなくて大丈夫です。思い出しながら、聞いてください。", {
    x: M, y: 5.5, w: CW, h: 0.4, isTextBox: true, margin: 0, fontFace: GO, fontSize: 14, color: MUTED_D,
  });
  note(s, "【0:20-1:20】最重要スライド。\nゆっくり読み上げて、5秒黙る。\nここで主語が「私」から「あなた」に移る。刺さらないと最後のCTAが効かない。");
}

// 3 11秒
{
  const s = S(true);
  s.addText("11", { x: 0.6, y: 1.5, w: 5.2, h: 3.2, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 190, bold: true, color: PINKB, align: "center", valign: "middle" });
  s.addText("秒", { x: 5.4, y: 3.3, w: 1.0, h: 0.8, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 30, bold: true, color: ONDARK, valign: "middle" });
  s.addText("医師が、患者の話を遮るまでの時間", { x: 6.8, y: 2.35, w: 5.8, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 22, bold: true, color: ONDARK, lineSpacing: 34 });
  body(s, "話し始めてから遮られるまでの中央値。\n先行研究では平均18秒と報告されている。", 6.8, 3.05, 5.8, 1.0, { color: MUTED_D, fontSize: 14 });
  s.addText("Singh Ospina et al., 2019 ／ Beckman & Frankel, 1984", {
    x: 6.8, y: 4.2, w: 5.8, h: 0.6, isTextBox: true, margin: 0, fontFace: GO, fontSize: 9.5, color: MUTED_D, lineSpacing: 15,
  });
  s.addText("では、私たち鍼灸師は？　――　私たちには 40分 あります。", {
    x: M, y: 5.9, w: CW, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 21, bold: true, color: ONDARK,
  });
  note(s, "【1:20-3:00】\n「これは医師を対象にした研究です」と必ず断る。鍼灸師のデータではない。\n最後の一行で、問題提起から一転して『鍼灸師は聴ける職種』という誇りを立てる。");
}

/* ══════════════════════════════════════════════════════════
   B. 自己紹介
   ══════════════════════════════════════════════════════════ */

// 3b 私は、聴いてもらう側でした
{
  const s = S(false);
  eyebrow(s, "はじめに", false);
  title(s, "私は、聴いてもらう側でした");
  card(s, M, 2.1, CW * 0.54, 2.5);
  s.addText("うつ病・パニック障害・双極性障害", {
    x: M + 0.5, y: 2.45, w: CW * 0.54 - 1.0, h: 0.4, isTextBox: true, margin: 0,
    fontFace: MIN, fontSize: 18, bold: true, color: INK,
  });
  body(s, "最重症期は、1日30錠。\n患者だった期間は、8年。\n\nいまは16年間、寛解。投薬もありません。",
    M + 0.5, 2.95, CW * 0.54 - 1.0, 1.4, { fontSize: 15, lineSpacing: 28 });
  s.addText("あのとき私がいちばん欲しかったのは、\n正しい助言では、ありませんでした。", {
    x: M + CW * 0.54 + 0.6, y: 2.45, w: CW * 0.46 - 0.6, h: 1.2, isTextBox: true, margin: 0,
    fontFace: MIN, fontSize: 20, bold: true, color: INK, lineSpacing: 34,
  });
  s.addText("ただ、話を最後まで\n聴いてもらうことでした。", {
    x: M + CW * 0.54 + 0.6, y: 3.75, w: CW * 0.46 - 0.6, h: 1.0, isTextBox: true, margin: 0,
    fontFace: MIN, fontSize: 22, bold: true, color: MORI, lineSpacing: 36,
  });
  note(s, "【3:00-4:30】\n★非常に個人的な開示です。話すかどうかは、ご自身で決めてください。\n　削除しても講演は成立します。ただ、この講演では最も強い一枚になり得ます。\n\n淡々と言う。同情を求めない。事実だけ置いて、最後の一行に重心をかける。\n★ここが、41枚目のワーク④『最後にちゃんと聴いてもらえたのはいつ？』への伏線になります。");
}

// 4 私は聴けていませんでした
{
  const s = S(false);
  eyebrow(s, "はじめに", false);
  title(s, "私も、聴けていませんでした");
  card(s, M, 2.05, CW, 2.9);
  body(s, "開業してしばらく、私は「主訴を聞き出す」ことに必死でした。\n患者さんの話を、情報として処理していたんです。\n\nいま思い出しても苦しいのは、涙をこぼした方に、\n次の質問をかぶせてしまったこと。",
    M + 0.5, 2.45, CW - 1.0, 2.2, { fontSize: 16, lineSpacing: 29 });
  s.addText("いまの私の話ではありません。これから話すことは、全部そこから始まっています。", {
    x: M, y: 5.2, w: CW, h: 0.4, isTextBox: true, margin: 0, fontFace: GO, fontSize: 14, color: MUTED,
  });
  note(s, "【3:00-5:00】\n自己開示。ここが弱いと、後半すべてが『上手い人の話』になって他人事になる。\nエピソードは1つだけ。具体的に、短く。\n※ここに当時の写真を1枚入れてもよい。");
}

// 5 転機
{
  const s = S(false);
  eyebrow(s, "はじめに", false);
  title(s, "変わったのは、症状ではありませんでした");
  const items = [["通院", "続くようになった"], ["紹介", "増えた"], ["自分", "消耗しなくなった"]];
  items.forEach((it, i) => {
    const x = M + i * (CW / 3);
    card(s, x, 2.15, CW / 3 - 0.35, 1.85);
    s.addText(it[0], { x: x + 0.4, y: 2.45, w: CW / 3 - 1.1, h: 0.45, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 24, bold: true, color: SHU });
    s.addText(it[1], { x: x + 0.4, y: 3.0, w: CW / 3 - 1.1, h: 0.7, isTextBox: true, margin: 0, fontFace: GO, fontSize: 15, color: INK, lineSpacing: 24 });
  });
  const st = [["11年目", "鍼灸院 開業"], ["20,000人+", "のべ臨床患者数"], ["修士", "武蔵野大学大学院 人間学"]];
  st.forEach((v, i) => {
    const x = M + i * (CW / 3);
    s.addText(v[0], { x, y: 4.5, w: CW / 3 - 0.35, h: 0.6, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 30, bold: true, color: INK });
    s.addText(v[1], { x, y: 5.12, w: CW / 3 - 0.35, h: 0.35, isTextBox: true, margin: 0, fontFace: GO, fontSize: 12, color: MUTED });
  });
  note(s, "【5:00-6:30】\n傾聴で何が変わったかを、症状の改善ではなく『関係・継続・自分の消耗度』で語る。\n数字は信頼の担保として置くだけ。誇らない。");
}

// 6 3事業
{
  const s = S(false);
  eyebrow(s, "はじめに", false);
  title(s, "いま、やっていること");
  const biz = [
    ["01", "臨床", "はりきゅう処ここちめいど", "三重県四日市市／2015年開院\n2026年夏 町田院オープン予定", SHU],
    ["02", "学び", "オンラインサロン ここちめいど", "傾聴を学ぶ鍼灸師コミュニティ\n2020年4月開始", MORI],
    ["03", "育成", "ここちはり", "感情・心・身体・技術の4軸\n2025年11月開始", SHU],
  ];
  biz.forEach((b, i) => {
    const x = M + i * (CW / 3);
    const w = CW / 3 - 0.35;
    card(s, x, 2.0, w, 2.6);
    circle(s, b[0], x + 0.35, 2.28, 0.6, b[4], PAPER, 15);
    s.addText(b[1], { x: x + 1.1, y: 2.35, w: w - 1.4, h: 0.45, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 21, bold: true, color: INK, valign: "middle" });
    s.addText(b[2], { x: x + 0.35, y: 3.08, w: w - 0.7, h: 0.6, isTextBox: true, margin: 0, fontFace: GO, fontSize: 13, bold: true, color: INK, lineSpacing: 21 });
    s.addText(b[3], { x: x + 0.35, y: 3.72, w: w - 0.7, h: 0.75, isTextBox: true, margin: 0, fontFace: GO, fontSize: 11.5, color: MUTED, lineSpacing: 19 });
  });
  body(s, "研究：全日本鍼灸学会「鍼灸院におけるうつと不安症状を有する患者の実態調査」第3報・第4報\n2026.06 日本精神神経学会 鍼灸シンポジウム シンポジスト",
    M, 4.9, CW, 0.75, { fontSize: 12, color: MUTED, lineSpacing: 21 });
  note(s, "【6:30-8:00】20秒で流す。\n「詳しくは最後に」と言って進む＝最後まで残る理由をつくる。\n実績の羅列にならないよう、早口にならず淡々と。");
}

/* ══════════════════════════════════════════════════════════
   C. なぜ鍼灸師に傾聴か
   ══════════════════════════════════════════════════════════ */

// 7 医師からの評価
{
  const s = S(false);
  eyebrow(s, "鍼灸師は、どう見られているか", false);
  title(s, "医師の先生方から、こう言われます");
  const v = [
    ["「患者さんと、\n　時間がある」", SHU],
    ["「身体全体を、\n　見てくれる」", MORI],
    ["「東洋医学の、\n　専門家だから」", SHU],
  ];
  v.forEach((it, i) => {
    const x = M + i * (CW / 3);
    const w = CW / 3 - 0.35;
    card(s, x, 2.1, w, 2.0);
    circle(s, "●", x + 0.35, 2.42, 0.3, it[1], PAPER, 9);
    s.addText(it[0], {
      x: x + 0.35, y: 2.95, w: w - 0.7, h: 0.95, isTextBox: true, margin: 0,
      fontFace: MIN, fontSize: 20, bold: true, color: INK, lineSpacing: 32,
    });
  });
  s.addText("これが、外から見た私たちです。期待されている、ということでもあります。", {
    x: M, y: 4.55, w: CW, h: 0.5, isTextBox: true, margin: 0, fontFace: GO, fontSize: 16, color: INK,
  });
  note(s, "【8:00-9:15】\n★『40分あります』（3枚目）を受けて、「実際、医師の先生方からもこう言われます」と繋ぐ。\n★これは調査データではなく、ご自身が言われてきた実感として話すこと。\n　「私が言われるのは」「よく言われるのは」という言い方で。データとして提示しない。\n誇らしい気持ちで読み上げてよい。次のスライドの落差が効きます。");
}

// 8 自問（濃色）
{
  const s = S(true);
  eyebrow(s, "鍼灸師は、どう見られているか", true);
  s.addText("では、できているでしょうか。", {
    x: M, y: 1.55, w: CW, h: 0.7, isTextBox: true, margin: 0, fontFace: GO, fontSize: 18, color: MUTED_D,
  });
  const q = ["東洋医学、できていますか？", "対話、できていますか？"];
  q.forEach((t, i) => {
    const y = 2.5 + i * 1.15;
    circle(s, "？", M, y + 0.02, 0.62, i === 0 ? "2C6448" : MINT, i === 0 ? ONDARK : INK, 20);
    s.addText(t, {
      x: M + 1.0, y, w: CW - 1.0, h: 0.68, isTextBox: true, margin: 0, valign: "middle",
      fontFace: MIN, fontSize: 32, bold: true, color: ONDARK,
    });
  });
  s.addText("私は、胸を張って「はい」とは言えませんでした。", {
    x: M, y: 5.1, w: CW, h: 0.55, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 20, bold: true, color: MINT,
  });
  note(s, "【9:15-10:15】\n★講演でいちばん空気が変わる場所。ここは急がない。\n2つの問いを読んだあと、必ず3秒黙る。会場（画面の向こう）に考えさせる。\n最後の一行は、自分のこととして静かに言う。責める調子にしない。\n\n※ここで『私も聴けていませんでした』（4枚目）を思い出してもらえると効く。");
}

// 9 今日のテーマ
{
  const s = S(false);
  eyebrow(s, "今日のテーマ", false);
  title(s, "今日は、対話の話をします");
  card(s, M, 2.15, CW, 2.3);
  s.addText("対　話", {
    x: M + 0.7, y: 2.7, w: 3.0, h: 0.9, isTextBox: true, margin: 0, valign: "middle",
    fontFace: MIN, fontSize: 44, bold: true, color: INK,
  });
  s.addText("→", {
    x: M + 3.8, y: 2.7, w: 0.8, h: 0.9, isTextBox: true, margin: 0, align: "center", valign: "middle",
    fontFace: GO, fontSize: 24, color: MUTED,
  });
  s.addText("そのための、ひとつの技法", {
    x: M + 4.8, y: 2.38, w: 6.0, h: 0.38, isTextBox: true, margin: 0, fontFace: GO, fontSize: 14, color: MUTED,
  });
  s.addText("傾　聴", {
    x: M + 4.8, y: 2.7, w: 6.0, h: 0.9, isTextBox: true, margin: 0, valign: "middle",
    fontFace: MIN, fontSize: 40, bold: true, color: SHU,
  });
  s.addText("東洋医学の話は、今日はしません。対話だけに、絞ります。", {
    x: M + 0.7, y: 3.8, w: CW - 1.4, h: 0.45, isTextBox: true, margin: 0, fontFace: GO, fontSize: 15, color: MUTED,
  });
  s.addText("対話は、才能ではありません。学べる技法です。", {
    x: M, y: 4.95, w: CW, h: 0.55, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 22, bold: true, color: INK,
  });
  note(s, "【10:15-10:45】\n★テーマ宣言。ここで聴衆に今日の地図を渡す。\n『東洋医学の話はしません』と絞ることで、話が締まる。欲張らない。\n★『技法です』と言い切ることが、後半の『訓練』『環境』の話への伏線になる。");
}

// 10 東洋医学は、情報がすべて
{
  const s = S(false);
  eyebrow(s, "なぜ、対話が要るのか", false);
  title(s, "「証」は、聞かなければ立ちません");
  s.addText("四診", { x: M, y: 2.1, w: 2.0, h: 0.35, isTextBox: true, margin: 0, fontFace: GO, fontSize: 12, bold: true, charSpacing: 2, color: MUTED });
  const shin = [["望", "見る"], ["聞", "聞く・嗅ぐ"], ["問", "尋ねる"], ["切", "触れる"]];
  shin.forEach((it, i) => {
    const x = M + i * 1.2;
    const on = it[0] === "問";
    circle(s, it[0], x, 2.55, 1.0, on ? SHU : "CDEBDC", on ? PAPER : INK, 28);
    s.addText(it[1], {
      x: x - 0.1, y: 3.68, w: 1.2, h: 0.35, isTextBox: true, margin: 0, align: "center",
      fontFace: GO, fontSize: 11, bold: on, color: on ? SHU : MUTED,
    });
  });
  card(s, M + 5.2, 2.2, CW - 5.2, 2.5);
  s.addText("主訴の、その裏にあること", {
    x: M + 5.7, y: 2.5, w: CW - 6.2, h: 0.4, isTextBox: true, margin: 0,
    fontFace: MIN, fontSize: 18, bold: true, color: INK,
  });
  s.addText("睡眠　／　食欲　／　便通　／　月経\n情志　／　生活リズム　／　仕事と家族", {
    x: M + 5.7, y: 3.0, w: CW - 6.2, h: 0.85, isTextBox: true, margin: 0,
    fontFace: GO, fontSize: 15, color: INK, lineSpacing: 28,
  });
  s.addText("どれも、聞かなければ出てきません。", {
    x: M + 5.7, y: 3.95, w: CW - 6.2, h: 0.4, isTextBox: true, margin: 0,
    fontFace: MIN, fontSize: 16, bold: true, color: SHU,
  });
  s.addText("東洋医学ができていない、のではなく ――　聞けていないだけ、かもしれません。", {
    x: M, y: 4.95, w: CW, h: 0.55, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 21, bold: true, color: INK,
  });
  note(s, "【10:45-12:15】\n★8枚目の「東洋医学、できていますか？」への答えを、ここで返す。\n四診のうち「問」だけを朱にしているのがポイント。望・聞・切は自分で取れるが、\n問診だけは患者さんに話してもらわないと1ミリも進まない。\n\n★最後の一行は、責めずに、自分の実感として。ここで会場が静かになります。");
}

// 11 慢性疼痛診療でも同じ
{
  const s = S(false);
  eyebrow(s, "なぜ、対話が要るのか", false);
  title(s, "現代の慢性疼痛診療でも、同じです");
  const pts = [
    ["患者教育", "医療者と患者の信頼関係を築く／痛みへの捉え方を是正する"],
    ["目標設定", "QOLにつながる目標を、患者と協働で設定する"],
    ["協働的意思決定", "患者の主体性と参加を引き出す（shared decision-making）"],
    ["ペーシングの管理", "「0か100か」という患者さんの思考パターンをつかむ"],
    ["アドヒアランスの管理", "続けられない理由を、一緒に分析する"],
  ];
  pts.forEach((it, i) => {
    const y = 2.15 + i * 0.76;
    card(s, M, y, CW, 0.64);
    circle(s, "●", M + 0.3, y + 0.2, 0.24, i === 2 ? MORI : SHU, PAPER, 7);
    s.addText(it[0], {
      x: M + 0.8, y, w: 3.1, h: 0.64, isTextBox: true, margin: 0, valign: "middle",
      fontFace: MIN, fontSize: 16, bold: true, color: INK,
    });
    s.addText(it[1], {
      x: M + 4.0, y, w: CW - 4.4, h: 0.64, isTextBox: true, margin: 0, valign: "middle",
      fontFace: GO, fontSize: 13, color: MUTED,
    });
  });
  s.addText("慢性疼痛診療研修会「痛みの治療」（日本いたみ財団）より要約", {
    x: M, y: 6.0, w: CW, h: 0.35, isTextBox: true, margin: 0, fontFace: GO, fontSize: 10.5, color: "4F6F60",
  });
  s.addText("5つとも、患者さんから話を聴かなければ、始まりません。", {
    x: M, y: 6.45, w: CW, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 22, bold: true, color: SHU,
  });
  note(s, "【12:15-14:00】\n★東洋医学（前スライド）と現代の疼痛診療が、同じ結論に着地することを見せる。\n「古い/新しい」の話ではなく、どちらも入口は対話だ、と。\n\n★出典は必ず明示すること。原典のスライドをそのまま複写せず、要約で提示しています。\n　もし図表をそのまま使う場合は、事前に許諾を取ってください。\n\n読み上げは5つ全部やらない。①と③だけ読んで、「全部そうです」と締めるとテンポが保てます。");
}

// 12 地域の鍼灸師に求められる4つの役割
{
  const s = S(false);
  eyebrow(s, "なぜ、対話が要るのか", false);
  title(s, "地域の鍼灸師に、求められている4つの役割");
  const f = [
    ["1", "レッドフラッグのアセスメント", "器質的疾患・神経疾患・精神疾患の可能性を念頭に問診し、\n専門医への早期紹介を判断する", SHU],
    ["2", "話を聴き、困りごとを整理する", "症状そのものより「症状によって できなくなったこと」に\n焦点を当てる", MORI],
    ["3", "医療・福祉等に繋げる", "精神科、訪問看護、ハローワーク、地域包括支援センター。\n多様な資源への橋渡し", SHU],
    ["4", "服薬アドヒアランスの向上", "薬への不安・不満を傾聴して言語化。医師・薬剤師と\n共有できる形に整理する", MORI],
  ];
  f.forEach((it, i) => {
    const x = M + (i % 2) * (CW / 2 + 0.15);
    const y = 2.1 + Math.floor(i / 2) * 1.55;
    const w = CW / 2 - 0.15;
    card(s, x, y, w, 1.35);
    circle(s, it[0], x + 0.3, y + 0.25, 0.5, it[3], PAPER, 15);
    s.addText(it[1], {
      x: x + 0.95, y: y + 0.22, w: w - 1.25, h: 0.42, isTextBox: true, margin: 0, valign: "middle",
      fontFace: MIN, fontSize: 16, bold: true, color: INK,
    });
    s.addText(it[2], {
      x: x + 0.32, y: y + 0.72, w: w - 0.64, h: 0.55, isTextBox: true, margin: 0,
      fontFace: GO, fontSize: 11.5, color: MUTED, lineSpacing: 19,
    });
  });
  s.addText("4つとも、入口は同じです。――　聴くこと。", {
    x: M, y: 5.35, w: CW, h: 0.55, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 24, bold: true, color: INK,
  });
  s.addText("米倉まな「地域医療における鍼灸師と精神科医との連携の実際」（2026.06）より", {
    x: M, y: 5.95, w: CW, h: 0.35, isTextBox: true, margin: 0, fontFace: GO, fontSize: 10.5, color: "4F6F60",
  });
  note(s, "【13:30-15:00】\n★ご自身の学会発表からの引用なので、堂々と出せます。\n★②と④は傾聴そのもの。①も、聴かなければ異変に気づけない。③も、\n　困りごとを聞き出せなければ どの資源に繋ぐか決められない。\n　→「4つとも入口は聴くこと」という締めに説得力が出ます。\n\n読み上げは①と②だけ。③④は「こういうこともします」と流す。");
}

// 13 レッドフラッグの事例
{
  const s = S(false);
  eyebrow(s, "なぜ、対話が要るのか", false);
  title(s, "聴いていたから、気づけました");
  const talk = [
    [0, "最近、気分が落ち込んでいて、眠れません"],
    [1, "落ち込みと、眠れない。日中の眠気はありますか？"],
    [0, "眠気はないんですが、時々ぼーっとして、誰かいるように感じたり、虫が見えるような気がして"],
    [1, "（あれっ）手が震えたり、お薬が効きすぎることはありませんか？"],
    [0, "あります。うつのお薬を飲むと、逆効果に感じるんです"],
  ];
  talk.forEach((t, i) => {
    const mine = t[0] === 1;
    const w = CW * 0.76;
    const x = mine ? M + CW - w : M;
    const y = 2.05 + i * 0.72;
    if (mine) {
      s.addShape(pres.ShapeType.roundRect, { x, y, w, h: 0.6, rectRadius: 0.1, fill: { color: INK } });
    } else {
      card(s, x, y, w, 0.6);
    }
    s.addText(t[1], {
      x: x + 0.35, y, w: w - 0.7, h: 0.6, isTextBox: true, margin: 0, valign: "middle",
      fontFace: GO, fontSize: 13, color: mine ? ONDARK : INK,
    });
  });
  s.addText("神経内科にご高診を依頼　→　レビー小体型認知症でした。", {
    x: M, y: 5.8, w: CW, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 22, bold: true, color: SHU,
  });
  s.addText("抑うつを訴えて来院され、ご高診後に判明した診断名（当院事例）：パーキンソン病／レビー小体型認知症／アルツハイマー病／脳腫瘍／甲状腺機能亢進症・低下症／更年期障害／月経前不快気分障害／産後うつ／糖尿病／微小血管狭心症　など", {
    x: M, y: 6.4, w: CW, h: 0.55, isTextBox: true, margin: 0, fontFace: GO, fontSize: 10.5, color: MUTED, lineSpacing: 17,
  });
  note(s, "【15:00-17:00】\n★この講演でいちばん強い実例。ここを丁寧に。\n会話を実際に演じる（患者役と自分役で声色を変える）と一気に引き込めます。\n\n★言いたいのは「傾聴＝優しさ」ではなく「傾聴＝安全性」。\n　聴いていなければ、幻視も手の震えも出てこなかった。\n　うつの薬を出し続けられていたかもしれない、という重さ。\n\n下段の診断名リストは読み上げない。目で見せるだけで十分効きます。");
}

// 8 あるある（投票）
{
  const s = S(false);
  eyebrow(s, "なぜ、鍼灸師に傾聴か", false);
  title(s, "でも、こうなっていませんか");
  const a = ["問診票を見ながら「他には？」と聞いている", "話が長い患者さんに、内心あせっている", "沈黙が怖くて、つい説明を始めてしまう"];
  a.forEach((t, i) => {
    const y = 2.15 + i * 1.05;
    card(s, M, y, CW, 0.85);
    circle(s, String(i + 1), M + 0.3, y + 0.16, 0.53, INK, PAPER, 16);
    s.addText(t, { x: M + 1.05, y, w: CW - 1.4, h: 0.85, isTextBox: true, margin: 0, fontFace: GO, fontSize: 17, color: INK, valign: "middle" });
  });
  s.addText("▶ Zoom投票：当てはまるもの全部（複数選択可）", {
    x: M, y: 5.5, w: CW, h: 0.4, isTextBox: true, margin: 0, fontFace: GO, fontSize: 13, bold: true, color: MORI,
  });
  note(s, "【14:00-15:30】自覚のスイッチ。\n★Zoom投票をここで使う（カメラオフ100名の参加感をつくる）。\n投票を締めて結果を画面共有し、「8割の方が③ですね」と読み上げる。\nその数字が次のスライドに直結する。");
}

// 9 構造の問題
{
  const s = S(false);
  eyebrow(s, "なぜ、鍼灸師に傾聴か", false);
  title(s, "「聴けない」のは、性格ではありません");
  const c = [["時間の圧", "次の患者さんが待っている"], ["経営の圧", "1人あたりの単価と回転"], ["職業的責任感", "「治さなければ」という思い"]];
  c.forEach((it, i) => {
    const x = M + i * (CW / 3);
    card(s, x, 2.1, CW / 3 - 0.35, 1.7);
    s.addText(it[0], { x: x + 0.35, y: 2.4, w: CW / 3 - 1.05, h: 0.45, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 20, bold: true, color: SHU });
    s.addText(it[1], { x: x + 0.35, y: 2.92, w: CW / 3 - 1.05, h: 0.65, isTextBox: true, margin: 0, fontFace: GO, fontSize: 13, color: MUTED, lineSpacing: 21 });
  });
  s.addText("構造の問題です。", { x: M, y: 4.2, w: CW, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 24, bold: true, color: INK });
  s.addText("悪いのは、あなたではありません。習っていないだけです。", {
    x: M, y: 4.85, w: CW, h: 0.6, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 28, bold: true, color: MORI,
  });
  note(s, "【15:30-16:15】\n★ここで責めると問合せは来ない。必ず逃げ道を用意する。\n「習っていないだけ」＝『習えば変わる』への伏線。");
}

// データ（実態調査）
{
  const s = S(false);
  eyebrow(s, "なぜ、対話が要るのか", false);
  title(s, "鍼灸院に、どんな方が来ているか");
  const cols = [
    ["診断名", ["うつ病　28%", "うつ状態　20%", "不眠症・睡眠障害　15%", "自律神経失調症　15%"], SHU],
    ["病　歴", ["10年以上　25%", "5〜10年　22%", "", "慢性的な経過の方が多い"], MORI],
    ["来院の背景", ["精神科で治療中の方", "他科から受診を勧められた方", "受診を迷っている・未受診の方", ""], SHU],
  ];
  cols.forEach((c, i) => {
    const x = M + i * (CW / 3);
    const w = CW / 3 - 0.35;
    card(s, x, 2.1, w, 2.75);
    s.addText(c[0], {
      x: x + 0.35, y: 2.4, w: w - 0.7, h: 0.42, isTextBox: true, margin: 0,
      fontFace: MIN, fontSize: 19, bold: true, color: c[2],
    });
    s.addText(c[1].filter(Boolean).join("\n"), {
      x: x + 0.35, y: 2.95, w: w - 0.7, h: 1.7, isTextBox: true, margin: 0,
      fontFace: GO, fontSize: 13.5, color: INK, lineSpacing: 26,
    });
  });
  s.addText("受診を迷っている方が、鍼灸院には来ます。", {
    x: M, y: 5.1, w: CW, h: 0.6, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 26, bold: true, color: INK,
  });
  s.addText("米倉まな・松浦悠人ら「鍼灸院におけるうつと不安症状を有する患者の実態調査（第4報）電子システムを用いた集積」第74回 全日本鍼灸学会学術大会", {
    x: M, y: 5.8, w: CW, h: 0.45, isTextBox: true, margin: 0, fontFace: GO, fontSize: 10.5, color: "4F6F60", lineSpacing: 17,
  });
  note(s, "【19:15-20:45】\n★ご自身の第4報のデータ。数字はPDFの記載どおりです。発表前に最新値をご確認ください。\n\n★締めの一行がこのスライドの主役。\n　「受診を迷っている方が来る」＝だからレッドフラッグを拾う耳が要る、\n　という13枚目への回収になります。\n数字は読み上げず、「うつ病が3割弱、10年以上の方が4分の1」くらいで流す。");
}

// 12 スルーライン（濃色）
{
  const s = S(true);
  s.addShape(pres.ShapeType.ellipse, { x: 4.7, y: 0.6, w: 6.3, h: 6.3, fill: { color: "18452F" } });
  s.addText("経絡に触れ、\nこころに耳を傾ける。", {
    x: 1.5, y: 2.4, w: 10.3, h: 1.9, isTextBox: true, margin: 0, align: "center",
    fontFace: MIN, fontSize: 40, bold: true, color: ONDARK, lineSpacing: 66,
  });
  s.addText("これこそが、鍼灸のチカラだと思っています。", {
    x: 1.5, y: 4.6, w: 10.3, h: 0.5, isTextBox: true, margin: 0, align: "center",
    fontFace: GO, fontSize: 16, color: MUTED_D,
  });
  note(s, "【18:40-19:15】\nこの講演を貫く一行。ここで一度置いて、最後（53枚目）で同じ言葉に戻る。\n読んだあと一拍おいてから次へ。急がない。");
}

/* ══════════════════════════════════════════════════════════
   D. 傾聴の正体
   ══════════════════════════════════════════════════════════ */

// 13 3つの誤解
{
  const s = S(false);
  eyebrow(s, "傾聴の正体", false);
  title(s, "傾聴の、3つの誤解");
  const c = ["黙って聞くこと", "共感して、同意すること", "アドバイスを我慢すること"];
  c.forEach((t, i) => {
    const x = M + i * (CW / 3);
    card(s, x, 2.2, CW / 3 - 0.35, 1.9);
    circle(s, "✕", x + (CW / 3 - 0.35) / 2 - 0.28, 2.5, 0.56, SHU, PAPER, 18);
    s.addText(t, { x: x + 0.3, y: 3.25, w: CW / 3 - 0.95, h: 0.7, isTextBox: true, margin: 0, align: "center", fontFace: GO, fontSize: 15, bold: true, color: INK, lineSpacing: 24 });
  });
  s.addText("全部、ちがいます。", { x: M, y: 4.6, w: CW, h: 0.7, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 32, bold: true, color: INK });
  note(s, "【19:15-20:40】\n言い切る。ここで『え、違うの？』という顔をさせたい。\n特に③は多くの人が『我慢すること』だと思っている。");
}

// 14 定義
{
  const s = S(false);
  eyebrow(s, "傾聴の正体", false);
  title(s, "鍼灸師の言葉に、翻訳すると");
  const r = [
    ["自己一致", "術者が整っていないと、鍼はぶれる"],
    ["無条件の肯定的関心", "証を立てる前に、まず在るものに触れる"],
    ["共感的理解", "相手の世界の「内側」から聴く"],
  ];
  r.forEach((it, i) => {
    const y = 2.15 + i * 1.05;
    card(s, M, y, CW, 0.88);
    s.addText(it[0], { x: M + 0.4, y, w: 3.4, h: 0.88, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 19, bold: true, color: SHU, valign: "middle" });
    s.addText("→", { x: M + 3.9, y, w: 0.5, h: 0.88, isTextBox: true, margin: 0, fontFace: GO, fontSize: 15, color: MUTED, valign: "middle" });
    s.addText(it[1], { x: M + 4.5, y, w: CW - 4.9, h: 0.88, isTextBox: true, margin: 0, fontFace: GO, fontSize: 16, color: INK, valign: "middle" });
  });
  s.addText("カール・ロジャーズが示した3条件（1957）", {
    x: M, y: 5.45, w: CW, h: 0.4, isTextBox: true, margin: 0, fontFace: GO, fontSize: 11.5, color: MUTED,
  });
  note(s, "【20:40-22:30】\nロジャーズをそのまま説明すると眠くなる。必ず鍼灸の言葉に翻訳して話す。\n『証を立てる前に』のくだりが、この講演でいちばん鍼灸師に刺さる言い回し。");
}

// 15-17 三原則
const principles = [
  ["原則 1", "沈黙を待つ", "置鍼の待ち時間、あなたは待てますよね。あれと同じです。", "相手が言い終わってから、心のなかで 1・2・3。\nたった3秒。ここで多くの人が、自分の言葉をかぶせています。", SHU],
  ["原則 2", "解釈を置く", "頭に浮かんだ「証」を、いったん脇に置く。", "早すぎる解釈は、患者さんの話を止めます。\n診立てはあとからでも間に合う。まず、最後まで。", MORI],
  ["原則 3", "自分を整える", "疲れている日は、聴けません。", "傾聴は知識ではなく「状態」です。\nだからこそ、一度学べば終わりではなく、訓練なのです。", SHU],
];
principles.forEach((p, i) => {
  const s = S(false);
  eyebrow(s, "傾聴の正体", false);
  circle(s, String(i + 1), M, 0.62, 0.85, p[4], PAPER, 30);
  s.addText(p[0], { x: M + 1.15, y: 0.62, w: 3, h: 0.35, isTextBox: true, margin: 0, fontFace: GO, fontSize: 12, bold: true, charSpacing: 2, color: MUTED });
  s.addText(p[1], { x: M + 1.15, y: 0.98, w: 8, h: 0.62, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 34, bold: true, color: INK });
  card(s, M, 2.3, CW, 1.35);
  s.addText("「" + p[2] + "」", { x: M + 0.5, y: 2.3, w: CW - 1.0, h: 1.35, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 22, bold: true, color: p[4], valign: "middle", lineSpacing: 34 });
  body(s, p[3], M, 4.0, CW, 1.3, { fontSize: 16, lineSpacing: 30 });
  note(s, ["【22:30-24:00】沈黙は『気を待つ』のと同じ、と重ねると鍼灸師には一発で伝わる。",
    "【46:30-48:00】その場で1回、声に出して実演してみせるとよい。\n\n★実例を足すなら：「薬が怖い」と言われたとき、\n　『どんなところが怖いですか？』『どうしてそう思うんですか？』と返すだけ。\n　それで出てきた不安を、主治医に相談できる形に整理して渡す。\n　（服薬アドヒアランスの支援＝反復と問い返しだけで成立します）",
    "【25:30-26:50】ここで『訓練』という言葉を初めて出す。後半のサロンの話への伏線。"][i]);
});

// 18 つなぎ
{
  const s = S(true);
  s.addText("ここまでは、知識です。", { x: M, y: 2.5, w: CW, h: 0.8, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 34, bold: true, color: MUTED_D });
  s.addText("知識では、1ミリも変わりません。", { x: M, y: 3.4, w: CW, h: 0.9, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 40, bold: true, color: ONDARK });
  s.addText("では、やってみましょう。", { x: M, y: 4.6, w: CW, h: 0.6, isTextBox: true, margin: 0, fontFace: GO, fontSize: 20, color: MINT });
  note(s, "【26:50-27:20】\nワークへの切り替え。ここで声のトーンを一段上げる。");
}

/* ══════════════════════════════════════════════════════════
   E. 体験ワーク（オンライン・カメラオフ版）
   ══════════════════════════════════════════════════════════ */

// 19 リフレーム
{
  const s = S(false);
  eyebrow(s, "体験ワーク", false);
  title(s, "今日は、顔が見えないほうが好都合です");
  card(s, M, 2.1, CW, 2.55);
  body(s, "表情という情報がない分、声そのものを聴くしかない。\nこれは、電話相談とまったく同じ条件です。\n\nカメラは、オフのままで大丈夫です。",
    M + 0.6, 2.45, CW - 1.2, 1.9, { fontSize: 18, lineSpacing: 32 });
  s.addText("使うもの：紙とペン ／ チャット ／ 耳", {
    x: M, y: 5.0, w: CW, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 19, bold: true, color: MORI,
  });
  note(s, "【27:20-28:00】\n★制約を『わざとそうした』に変える。ここで場の空気が一段変わる。\n『カメラはオフのままで大丈夫』は必ず言う。言うほど参加率が上がる。");
}

// 20 ワーク①説明
{
  const s = S(false);
  eyebrow(s, "体験ワーク ①", false);
  title(s, "思い出せますか");
  card(s, M, 2.05, CW, 3.05);
  const steps = ["紙とペンを用意してください", "直近に施術した患者さんを、1人 思い浮かべてください", "その方が言った「言葉そのもの」を、3つ書き出してください"];
  steps.forEach((t, i) => {
    const y = 2.3 + i * 0.72;
    circle(s, String(i + 1), M + 0.45, y + 0.06, 0.45, i === 2 ? SHU : INK, PAPER, 14);
    s.addText(t, { x: M + 1.15, y, w: CW - 1.8, h: 0.58, isTextBox: true, margin: 0, fontFace: GO, fontSize: 17, color: INK, valign: "middle" });
  });
  s.addText("要約ではなく、口にした言葉のまま。", {
    x: M + 1.15, y: 4.35, w: CW - 1.8, h: 0.45, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 17, bold: true, color: SHU,
  });
  s.addText("時間は 90秒 です。", { x: M, y: 5.35, w: CW, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 22, bold: true, color: INK });
  note(s, "【28:00-28:40】\n★カメラオフでも100%全員が当事者になる、唯一のワーク。絶対に削らない。\n『紙とペン』は開始前アナウンスでも流してもらうこと。\n録画で見ている方へ：「ここで一時停止してください」と一言添える。");
}

// 21 90秒タイマー
{
  const s = S(true);
  s.addShape(pres.ShapeType.ellipse, { x: 4.87, y: 1.35, w: 3.6, h: 3.6, fill: { color: INK }, line: { color: MINT, width: 3 } });
  s.addText("90", { x: 4.87, y: 1.35, w: 3.6, h: 3.6, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: MIN, fontSize: 96, bold: true, color: ONDARK });
  s.addText("秒", { x: 8.5, y: 3.4, w: 0.8, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 22, bold: true, color: MUTED_D });
  s.addText("患者さんが言った「言葉そのもの」を、3つ", { x: 1.5, y: 5.4, w: 10.3, h: 0.5, isTextBox: true, margin: 0, align: "center", fontFace: GO, fontSize: 16, color: MUTED_D });
  note(s, "【28:40-30:10】\nここは黙る。BGMを小さくかけてもよい。\nタイマーはスマホか腕時計で計る（PowerPointは自動カウントダウンできません）。\n途中で話しかけない。60秒過ぎたら次のスライドへ。");
}

// 22 のこり30秒
{
  const s = S(true);
  s.addShape(pres.ShapeType.ellipse, { x: 4.87, y: 1.35, w: 3.6, h: 3.6, fill: { color: INK }, line: { color: MUTED_D, width: 3 } });
  s.addText("30", { x: 4.87, y: 1.35, w: 3.6, h: 3.6, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: MIN, fontSize: 96, bold: true, color: MUTED_D });
  s.addText("のこり", { x: 3.6, y: 3.4, w: 1.2, h: 0.5, isTextBox: true, margin: 0, align: "right", fontFace: MIN, fontSize: 20, bold: true, color: MUTED_D });
  s.addText("秒", { x: 8.5, y: 3.4, w: 0.8, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 20, bold: true, color: MUTED_D });
  note(s, "【30:10-30:40】\n無言でこのスライドに切り替えるだけ。声はかけない。");
}

// 23 種明かし
{
  const s = S(false);
  eyebrow(s, "体験ワーク ①", false);
  title(s, "症状は、書けたはずです");
  s.addText("でも「言葉」は、書けましたか？", {
    x: M, y: 1.85, w: CW, h: 0.85, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 36, bold: true, color: SHU,
  });
  card(s, M, 2.95, CW, 1.55);
  body(s, "私たちは、聴いたつもりで「情報」に変換しています。\n人の言葉としては、残っていないんです。",
    M + 0.6, 3.35, CW - 1.2, 0.9, { fontSize: 18, lineSpacing: 32 });
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 4.75, w: CW, h: 0.95, rectRadius: 0.1, fill: { color: MINT } });
  s.addText("▶ チャットへ　「何個 書けましたか？」　数字だけ送ってください", {
    x: M + 0.5, y: 4.75, w: CW - 1.0, h: 0.95, isTextBox: true, margin: 0, valign: "middle",
    fontFace: GO, fontSize: 17, bold: true, color: INK,
  });
  note(s, "【30:40-32:30】\n★数字1文字なら100名でも一気に流れる。\n「0」「1」が並ぶ光景そのものがメッセージ。『ほら、みなさん同じです』と拾う。\n★サロンメンバー1〜2名に、最初に投げる役を頼んでおくこと（口火を切る役）。");
}

// 24 ワーク②説明
{
  const s = S(false);
  eyebrow(s, "体験ワーク ②", false);
  title(s, "10秒、黙ってみます");
  card(s, M, 2.1, CW, 2.35);
  body(s, "今から10秒、私は何も話しません。\nみなさんも、何もしないでください。\n\nスマホも、見ないでください。",
    M + 0.6, 2.45, CW - 1.2, 1.75, { fontSize: 18, lineSpacing: 32 });
  s.addText("では、いきます。", { x: M, y: 4.8, w: CW, h: 0.6, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 26, bold: true, color: SHU });
  note(s, "【32:30-33:00】\n次の10枚が1秒ずつのカウントダウンスライドです。\n★事前設定：カウントダウン10枚を選択 →「画面切り替え」→「自動的に切り替え 00:01」\n　クリック時のチェックは外す。これで自動で10秒進みます。");
}

// 25-34 カウントダウン 10→1
for (let k = 10; k >= 1; k--) {
  const s = S(true);
  s.addShape(pres.ShapeType.ellipse, { x: 5.17, y: 1.75, w: 3.0, h: 3.0, fill: { color: INK }, line: { color: k <= 3 ? PINKB : "2C6448", width: 2.5 } });
  s.addText(String(k), {
    x: 5.17, y: 1.75, w: 3.0, h: 3.0, isTextBox: true, margin: 0, align: "center", valign: "middle",
    fontFace: MIN, fontSize: 88, bold: true, color: k <= 3 ? PINKB : ONDARK,
  });
  s.addText("沈黙のワーク", { x: 1.5, y: 5.35, w: 10.3, h: 0.4, isTextBox: true, margin: 0, align: "center", fontFace: GO, fontSize: 13, color: MUTED_D });
  if (k === 10) note(s, "★このスライドから10枚が、10秒のカウントダウンです。\n\n【設定方法】\n1. スライド一覧でこの10枚を選択\n2.「画面切り替え」タブ →「自動的に切り替え」に 00:01 と入力\n3.「クリック時」のチェックを外す\n\nこれで自動的に10秒進みます。数字を画面に出すことで、\nアーカイブ視聴者に『配信事故』と誤解されるのを防げます。\n\nこの10秒、絶対に話しかけないこと。");
  else note(s, "無言。（自動切り替え 1秒）");
}

// 35 沈黙のあと
{
  const s = S(false);
  eyebrow(s, "体験ワーク ②", false);
  title(s, "どうでしたか");
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 1.95, w: CW, h: 0.95, rectRadius: 0.1, fill: { color: MINT } });
  s.addText("▶ チャットへ　落ち着かなかった方は「1」を送ってください", {
    x: M + 0.5, y: 1.95, w: CW - 1.0, h: 0.95, isTextBox: true, margin: 0, valign: "middle",
    fontFace: GO, fontSize: 17, bold: true, color: INK,
  });
  card(s, M, 3.25, CW, 2.25);
  body(s, "たった10秒でも、落ち着かないですよね。\n\nこの居心地の悪さが、あなたが\n患者さんの沈黙をつぶしている理由です。",
    M + 0.6, 3.6, CW - 1.2, 1.8, { fontSize: 18, lineSpacing: 32 });
  note(s, "【33:10-34:00】\n『あなたが沈黙をつぶしている理由』——ここは責めるトーンにしない。\n自分もそうだった、という言い方で。");
}

// 36 ワーク③説明
{
  const s = S(false);
  eyebrow(s, "体験ワーク ③", false);
  title(s, "聴き比べてみてください");
  card(s, M, 2.05, CW, 1.6);
  s.addText("患者さんの、同じ一言に対して", { x: M + 0.6, y: 2.3, w: CW - 1.2, h: 0.4, isTextBox: true, margin: 0, fontFace: GO, fontSize: 14, color: MUTED });
  s.addText("「最近、夜眠れなくて……仕事も、まあ、いろいろあって」", {
    x: M + 0.6, y: 2.75, w: CW - 1.2, h: 0.7, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 24, bold: true, color: INK,
  });
  s.addText("2通りの応対を、音声で聴いていただきます。", {
    x: M, y: 4.0, w: CW, h: 0.5, isTextBox: true, margin: 0, fontFace: GO, fontSize: 17, color: INK,
  });
  s.addText("目を閉じて聴いてください。顔が見えないぶん、声に集中できます。", {
    x: M, y: 4.6, w: CW, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 18, bold: true, color: MORI,
  });
  note(s, "【34:00-34:40】\n★音声は事前録音しておく（パターンA・B 各60秒）。生の実演はぶれる。\n相方はサロンメンバーに依頼。音量を揃えること。\n★Zoomの画面共有で「コンピューターの音声を共有」を必ずON。リハで要テスト。");
}

// 37 A/B比較
{
  const s = S(false);
  eyebrow(s, "体験ワーク ③", false);
  title(s, "何が、違いましたか");
  const cols = [
    ["A", "聴けていない応対", "「そうですか、不眠ですね。\n何時ごろお休みですか？\n肩こりはあります？」", "患者さんの話は、そこで終わる", INK],
    ["B", "傾聴の応対", "「……いろいろ、あって」\n\n（そして、黙る）", "患者さんが、自分から続きを話し出す", SHU],
  ];
  cols.forEach((c, i) => {
    const x = M + i * (CW / 2 + 0.15);
    const w = CW / 2 - 0.15;
    card(s, x, 2.0, w, 3.0);
    circle(s, c[0], x + 0.35, 2.25, 0.5, c[4], PAPER, 17);
    s.addText(c[1], { x: x + 1.0, y: 2.25, w: w - 1.3, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 19, bold: true, color: c[4], valign: "middle" });
    s.addText(c[2], { x: x + 0.35, y: 2.95, w: w - 0.7, h: 1.25, isTextBox: true, margin: 0, fontFace: GO, fontSize: 14.5, color: INK, lineSpacing: 25 });
    s.addText(c[3], { x: x + 0.35, y: 4.3, w: w - 0.7, h: 0.5, isTextBox: true, margin: 0, fontFace: GO, fontSize: 13, bold: true, color: MUTED });
  });
  s.addText("Aで取れた情報と、Bで取れた情報。どちらが多かったですか？", {
    x: M, y: 5.3, w: CW, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 21, bold: true, color: INK,
  });
  note(s, "【34:40-38:30】\n2本の音声を流したあと、このスライドで問いかける。\n答えは待たなくていい。次のスライドで言い切る。");
}

// 38 情報量の話（濃色）
{
  const s = S(true);
  s.addText("傾聴は、優しさの話ではありません。", { x: M, y: 2.5, w: CW, h: 0.9, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 34, bold: true, color: MUTED_D });
  s.addText("情報量の話です。", { x: M, y: 3.5, w: CW, h: 1.0, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 48, bold: true, color: ONDARK });
  note(s, "【38:30-39:00】\nこの講演でいちばん強いメッセージ。ゆっくり、間を空けて言う。\n『優しくしましょう』では鍼灸師は動かない。『情報が取れる』なら動く。");
}

// 39 ワーク④ チャット一斉投下
{
  const s = S(false);
  eyebrow(s, "体験ワーク ④", false);
  title(s, "最後に、ひとつだけ");
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 1.95, w: CW, h: 1.35, rectRadius: 0.1, fill: { color: SHU } });
  s.addText("あなたが最後に「ちゃんと聴いてもらえた」のは、\nいつ、誰にですか？", {
    x: M + 0.5, y: 1.95, w: CW - 1.0, h: 1.35, isTextBox: true, margin: 0, valign: "middle",
    fontFace: MIN, fontSize: 22, bold: true, color: "FFFFFF", lineSpacing: 36,
  });
  card(s, M, 3.6, CW, 1.6);
  body(s, "チャットに打ってください。でも、まだ送らないでください。\n私が 3・2・1 と数えたら、全員 同時に送信します。",
    M + 0.6, 3.95, CW - 1.2, 0.95, { fontSize: 18, lineSpacing: 32 });
  s.addText("全部は読めません。3つだけ拾わせてください。", {
    x: M, y: 5.55, w: CW, h: 0.4, isTextBox: true, margin: 0, fontFace: GO, fontSize: 13, color: MUTED,
  });
  note(s, "【39:00-41:30】\n★一斉送信にすることで『自分だけ書く』恥ずかしさが消える。\n★共同ホストに事前依頼：流れてきた投稿を3つコピーしておいてもらう。\n★『全部は読めません、3つだけ』を先に宣言（読まれない不満を防ぐ）。\n拾った言葉に短くコメントする。ここで出た言葉が、この講演で一番強いコピーになる。");
}

// 40 13分（濃色 + QR一瞬）
{
  const s = S(true);
  s.addText("いまのは、全部で 13分です。", { x: M, y: 2.2, w: CW, h: 0.8, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 34, bold: true, color: MUTED_D });
  s.addText("あなたの施術は、40分あります。", { x: M, y: 3.1, w: CW, h: 1.0, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 44, bold: true, color: ONDARK });
  s.addShape(pres.ShapeType.roundRect, { x: 9.6, y: 4.45, w: 2.9, h: 2.35, rectRadius: 0.08, fill: { color: "1B4E36" }, line: { color: MUTED_D, width: 1, dashType: "dash" } });
  s.addText("▶ ここに\n公式LINEのQRを\n貼ってください", { x: 9.7, y: 4.75, w: 2.7, h: 1.1, isTextBox: true, margin: 0, align: "center", fontFace: GO, fontSize: 11, color: MUTED_D, lineSpacing: 18 });
  s.addText("先に撮っておきたい方はどうぞ", { x: 9.6, y: 6.25, w: 2.9, h: 0.35, isTextBox: true, margin: 0, align: "center", fontFace: GO, fontSize: 10, color: MUTED_D });
  note(s, "【41:30-42:00】\n★希望のピーク。ここで聴衆は『やってみたい』状態。\n『あとで出しますが、先に撮っておきたい方はどうぞ』と一言だけ。売り込まない。\n★QR画像を貼り替えてください（点線枠を削除して画像を配置）。");
}

/* ══════════════════════════════════════════════════════════
   F. 臨床への実装
   ══════════════════════════════════════════════════════════ */

// 41 3つの型
{
  const s = S(false);
  eyebrow(s, "明日からの臨床", false);
  title(s, "明日からできる、3つの型");
  const t = [["最初の90秒", "遮らない"], ["反復＋一拍", "相手の言葉のまま返す"], ["締めを閉じない", "開いた質問で終わる"]];
  t.forEach((it, i) => {
    const x = M + i * (CW / 3);
    const w = CW / 3 - 0.35;
    card(s, x, 2.15, w, 2.3);
    circle(s, "型" + (i + 1), x + w / 2 - 0.42, 2.45, 0.84, i === 1 ? MORI : SHU, PAPER, 17);
    s.addText(it[0], { x: x + 0.25, y: 3.45, w: w - 0.5, h: 0.45, isTextBox: true, margin: 0, align: "center", fontFace: MIN, fontSize: 20, bold: true, color: INK });
    s.addText(it[1], { x: x + 0.25, y: 3.9, w: w - 0.5, h: 0.4, isTextBox: true, margin: 0, align: "center", fontFace: GO, fontSize: 13, color: MUTED });
  });
  s.addText("全部やらなくていいです。ひとつだけ選んでください。", {
    x: M, y: 4.9, w: CW, h: 0.5, isTextBox: true, margin: 0, fontFace: GO, fontSize: 16, color: INK,
  });
  note(s, "【42:00-43:00】\n先に3つ見せてから、1枚ずつ展開する（地図を渡してから歩く）。");
}

// 42-44 型①②③
const kata = [
  ["型 1", "最初の90秒は、遮らない", ["問診票を見ない", "ペンを、いったん置く", "「今日はどうされましたか？」のあと、90秒 何も足さない"],
    "90秒です。1日10人でも15分。それだけです。", SHU],
  ["型 2", "反復して、一拍おく", ["相手の言葉を、相手の言葉のまま返す", "「ずっと、しんどかったんですね」", "要約もリフレーミングもしない"],
    "言い換えた瞬間、それは自分の言葉になります。", MORI],
  ["型 3", "施術後を、閉じない", ["✕　「軽くなりましたか？」", "◯　「いま、どんな感じですか？」", "閉じた質問は、情報を閉じる"],
    "最後のひとことで、次回の情報量が決まります。", SHU],
];
kata.forEach((k, i) => {
  const s = S(false);
  eyebrow(s, "明日からの臨床", false);
  circle(s, String(i + 1), M, 0.62, 0.85, k[4], PAPER, 30);
  s.addText(k[0], { x: M + 1.15, y: 0.62, w: 3, h: 0.35, isTextBox: true, margin: 0, fontFace: GO, fontSize: 12, bold: true, charSpacing: 2, color: MUTED });
  s.addText(k[1], { x: M + 1.15, y: 0.98, w: 9.5, h: 0.62, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 32, bold: true, color: INK });
  k[2].forEach((t, j) => {
    const y = 2.2 + j * 0.92;
    card(s, M, y, CW, 0.75);
    circle(s, "●", M + 0.35, y + 0.235, 0.28, k[4], PAPER, 8);
    s.addText(t, { x: M + 0.95, y, w: CW - 1.3, h: 0.75, isTextBox: true, margin: 0, fontFace: GO, fontSize: 16, color: INK, valign: "middle" });
  });
  s.addText(k[3], { x: M, y: 5.15, w: CW, h: 0.55, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 21, bold: true, color: k[4] });
  note(s, ["【43:00-45:00】『90秒』という具体的な数字が、持ち帰りやすさを決める。ここを一番ゆっくり話す。",
    "【45:00-46:30】その場で1回、声に出して実演してみせるとよい。",
    "【46:30-48:00】✕と◯の対比は音読する。耳だけで聴いている人がいる。"][i]);
});

// 症例 1／紹介
{
  const s = S(false);
  eyebrow(s, "症例", false);
  title(s, "標準治療で、難渋していた方");
  card(s, M, 2.1, CW * 0.56, 2.75);
  s.addText("42歳　女性　主婦", {
    x: M + 0.5, y: 2.4, w: CW * 0.56 - 1.0, h: 0.4, isTextBox: true, margin: 0,
    fontFace: MIN, fontSize: 19, bold: true, color: INK,
  });
  body(s, "主訴：全身倦怠感・頭痛・睡眠障害\n\nうつ病の発症から17年。再発をくり返してきた。\n2ヶ月前、任意入院を勧められた。\nけれど子どもは3歳。入院は、難しかった。",
    M + 0.5, 2.9, CW * 0.56 - 1.0, 1.8, { fontSize: 14.5, lineSpacing: 27 });
  const rx = M + CW * 0.56 + 0.55, rw = CW - CW * 0.56 - 0.55;
  const facts = [["初診時", "PHQ-9　13点　／　PHQ-15　20点"], ["東洋医学的病態", "脾陽虚・胃熱・肝鬱気滞"], ["ご本人の目標", "「働くこと」"]];
  facts.forEach((f, i) => {
    const y = 2.2 + i * 0.92;
    s.addText(f[0], { x: rx, y, w: rw, h: 0.3, isTextBox: true, margin: 0, fontFace: GO, fontSize: 11, bold: true, color: MUTED });
    s.addText(f[1], {
      x: rx, y: y + 0.32, w: rw, h: 0.42, isTextBox: true, margin: 0,
      fontFace: MIN, fontSize: i === 2 ? 22 : 17, bold: true, color: i === 2 ? SHU : INK,
    });
  });
  s.addText("入院を避けたい。それが、来院の理由でした。", {
    x: M, y: 5.15, w: CW, h: 0.55, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 23, bold: true, color: INK,
  });
  s.addText("米倉まな・松浦悠人・柴田健一「標準治療で難渋したうつ病患者に鍼灸と傾聴が奏功した一症例」日本うつ病学会", {
    x: M, y: 5.85, w: CW, h: 0.35, isTextBox: true, margin: 0, fontFace: GO, fontSize: 10, color: "4F6F60",
  });
  note(s, "【44:30-45:40】\n★共同演者（松浦悠人先生・柴田健一先生）のお名前と学会名・年をご確認ください。\n\n『目標：働くこと』を必ず読む。これが傾聴で聴き取った治療目標です。\n症状ではなく、その人が何をしたいのか。ここが後半の型②③に繋がります。");
}

// 症例 2／鍼灸でやっていたこと
{
  const s = S(false);
  eyebrow(s, "症例", false);
  title(s, "鍼灸で、やっていたこと");
  const tx = [["温補脾陽", "全身倦怠感・めまい・冷え"], ["清胃熱", "胃のむかつき・便秘"], ["疏肝理気", "頭痛・睡眠障害・イライラ"]];
  tx.forEach((t, i) => {
    const y = 2.1 + i * 0.72;
    card(s, M, y, CW * 0.58, 0.6);
    s.addText(t[0], { x: M + 0.35, y, w: 1.6, h: 0.6, isTextBox: true, margin: 0, valign: "middle", fontFace: MIN, fontSize: 16, bold: true, color: SHU });
    s.addText(t[1], { x: M + 2.1, y, w: CW * 0.58 - 2.4, h: 0.6, isTextBox: true, margin: 0, valign: "middle", fontFace: GO, fontSize: 13, color: INK });
  });
  const rx = M + CW * 0.58 + 0.4, rw = CW - CW * 0.58 - 0.4;
  card(s, rx, 2.1, rw, 2.34);
  s.addText("使った経穴", { x: rx + 0.35, y: 2.35, w: rw - 0.7, h: 0.32, isTextBox: true, margin: 0, fontFace: GO, fontSize: 11, bold: true, color: MUTED });
  s.addText("百会　風池　完骨　肩井\n足三里　三陰交", {
    x: rx + 0.35, y: 2.72, w: rw - 0.7, h: 0.85, isTextBox: true, margin: 0,
    fontFace: MIN, fontSize: 19, bold: true, color: INK, lineSpacing: 32,
  });
  s.addText("＋　腹部の散鍼（鍼による擦過刺激）", {
    x: rx + 0.35, y: 3.65, w: rw - 0.7, h: 0.4, isTextBox: true, margin: 0, fontFace: GO, fontSize: 13, color: MORI,
  });
  body(s, "使用鍼：0.12〜0.20mm ステンレス鍼　／　灸：温筒灸・棒灸\n10診目〜 頭部に電気鍼（1Hz・15分）　／　15診目〜 頭痛にセルフケアの耳灸を指導",
    M, 4.6, CW, 0.8, { fontSize: 12, color: MUTED, lineSpacing: 21 });
  s.addText("鍼灸は、特別なことをしていません。", {
    x: M, y: 5.5, w: CW, h: 0.55, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 24, bold: true, color: INK,
  });
  note(s, "【45:40-46:50】\n★ここを丁寧にやると、鍼灸師の聴衆が『自分にもできる』と思えます。\n　奇をてらった手技は使っていない、と言い切ること。\n\n★締めの一行が重要：差は鍼灸の腕ではなく、上に乗せた傾聴にある、という含み。\n　言い過ぎない。「特別なことはしていません」で止めて、次のスライドへ。");
}

// 症例 3／傾聴の4段階
{
  const s = S(false);
  eyebrow(s, "症例", false);
  title(s, "同じように、聴いていたわけではありません");
  const st = [
    ["2–4診", "受容と共感で、まず信頼関係をつくる", "頭痛薬と頓服が、止まった", SHU],
    ["5–10診", "家族の課題について、情報を整理する", "抑うつを感じにくくなり、出かけるように", MORI],
    ["11–15診", "行動と問題に焦点を当て、できる範囲を明確に", "主治医も親も「回復に驚いている」", SHU],
    ["16–20診", "どう感じているかを聴き、感情の認知を促す", "「次回も鍼灸、楽しみにしている」", MORI],
  ];
  st.forEach((t, i) => {
    const y = 2.1 + i * 0.86;
    card(s, M, y, CW, 0.74);
    s.addText(t[0], { x: M + 0.35, y, w: 1.1, h: 0.74, isTextBox: true, margin: 0, valign: "middle", fontFace: MIN, fontSize: 15, bold: true, color: t[3] });
    s.addText(t[1], { x: M + 1.6, y, w: 4.9, h: 0.74, isTextBox: true, margin: 0, valign: "middle", fontFace: GO, fontSize: 13.5, bold: true, color: INK });
    s.addText("→", { x: M + 6.6, y, w: 0.4, h: 0.74, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: GO, fontSize: 12, color: MUTED });
    s.addText(t[2], { x: M + 7.1, y, w: CW - 7.45, h: 0.74, isTextBox: true, margin: 0, valign: "middle", fontFace: GO, fontSize: 13, color: MUTED });
  });
  s.addText("傾聴にも、順番があります。", {
    x: M, y: 5.75, w: CW, h: 0.6, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 26, bold: true, color: INK,
  });
  note(s, "【46:50-48:20】\n★この講演の核心スライド。いちばん時間をかけてよい場所です。\n\n受容共感 → 情報整理 → 行動と問題の焦点化 → 感情の認知。\nいきなり4段目をやると失敗します。信頼ができる前に感情に触れない。\n\n★「順番がある」＝「学べる」ということ。\n　才能の話ではないという、この講演全体の主張がここで実証されます。\n　57枚目『独学で伸びない理由』への最短の橋になります。");
}

// 症例 4／結果
{
  const s = S(false);
  eyebrow(s, "症例", false);
  title(s, "20診　184日");
  s.addText("PHQ-9（うつ症状）の推移", {
    x: M, y: 2.05, w: CW * 0.6, h: 0.32, isTextBox: true, margin: 0, fontFace: GO, fontSize: 12, bold: true, color: MUTED,
  });
  s.addChart(pres.ChartType.line,
    [{ name: "PHQ-9", labels: ["初診", "5診目", "10診目", "15診目", "20診目"], values: [13, 17, 10, 6, 5] }],
    {
      x: M - 0.1, y: 2.4, w: CW * 0.6, h: 2.6,
      chartColors: [MORI], lineSize: 2.5,
      lineDataSymbol: "circle", lineDataSymbolSize: 8, lineDataSymbolLineColor: PAPER,
      showLegend: false, showTitle: false, showValue: false,
      valAxisMinVal: 0, valAxisMaxVal: 20, valAxisMajorUnit: 5,
      valAxisLabelColor: MUTED, valAxisLabelFontSize: 10, valAxisLabelFontFace: GO,
      catAxisLabelColor: MUTED, catAxisLabelFontSize: 10, catAxisLabelFontFace: GO,
      valGridLine: { color: "D2EBDF", size: 1 }, catGridLine: { style: "none" },
      valAxisLineShow: false, catAxisLineColor: "BCDDCC",
      border: { pt: 0, color: CREAM }, fill: CREAM,
    });
  const rx = M + CW * 0.62, rw = CW - CW * 0.62;
  s.addText("PHQ-9", { x: rx, y: 2.2, w: rw, h: 0.3, isTextBox: true, margin: 0, fontFace: GO, fontSize: 11, bold: true, color: MUTED });
  s.addText("13　→　5", {
    x: rx, y: 2.5, w: rw, h: 0.95, isTextBox: true, margin: 0,
    fontFace: MIN, fontSize: 46, bold: true, color: MORI,
  });
  body(s, "身体症状（PHQ-15）は 20 → 19。\n大きくは変わっていません。\n\n変わったのは、気分と睡眠、\nそして日常生活でした。",
    rx, 3.6, rw, 1.5, { fontSize: 13.5, lineSpacing: 24 });
  s.addText("入院は、回避できました。", {
    x: M, y: 5.35, w: CW, h: 0.6, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 28, bold: true, color: INK,
  });
  s.addText("米倉まな・松浦悠人・柴田健一「標準治療で難渋したうつ病患者に鍼灸と傾聴が奏功した一症例」日本うつ病学会", {
    x: M, y: 6.05, w: CW, h: 0.35, isTextBox: true, margin: 0, fontFace: GO, fontSize: 10, color: "4F6F60",
  });
  note(s, "【48:20-49:30】\n★5診目で いったん上がっている（13→17）ことを隠さない。\n　「最初はむしろ上がりました」と正直に言うほうが、信頼されます。\n　※理由は症例報告では述べられていないので、推測を語らないこと。\n\n★PHQ-15（身体症状）がほぼ横ばいであることも、そのまま伝える。\n　過大に言わないことが、この講演全体の誠実さを担保します。\n\n締めの一行は、静かに。「入院は、回避できました」");
}

// 症例：継続患者に起きた変化
{
  const s = S(false);
  eyebrow(s, "明日からの臨床", false);
  title(s, "続けて来られた方に、起きたこと");
  const ch = [
    ["復職・転職", "鍼灸を続けながら復職、または\nより自分に合った職場へ移った", SHU],
    ["起業・就労形態の変更", "新しい働き方を選び、自分の\nペースで社会参加を続けている", MORI],
    ["日常生活の安定", "就労形態を調整しながら、\n安定した生活を継続している", MORI],
    ["多職種連携による支援", "医師・薬剤師・行政制度を\n組み合わせて社会復帰を目指す", SHU],
  ];
  ch.forEach((it, i) => {
    const x = M + (i % 2) * (CW / 2 + 0.15);
    const y = 2.1 + Math.floor(i / 2) * 1.5;
    const w = CW / 2 - 0.15;
    card(s, x, y, w, 1.3);
    s.addText(it[0], {
      x: x + 0.4, y: y + 0.22, w: w - 0.8, h: 0.42, isTextBox: true, margin: 0,
      fontFace: MIN, fontSize: 18, bold: true, color: it[2],
    });
    s.addText(it[1], {
      x: x + 0.4, y: y + 0.68, w: w - 0.8, h: 0.55, isTextBox: true, margin: 0,
      fontFace: GO, fontSize: 12.5, color: INK, lineSpacing: 21,
    });
  });
  s.addText("鍼灸院の中だけで、支援を完結させない。", {
    x: M, y: 5.25, w: CW, h: 0.6, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 25, bold: true, color: INK,
  });
  s.addText("医師・行政制度・地域の医療資源を、組み合わせて使う。", {
    x: M, y: 5.9, w: CW, h: 0.45, isTextBox: true, margin: 0, fontFace: GO, fontSize: 15, color: MUTED,
  });
  note(s, "【49:30-51:00】\n★「症状が良くなりました」ではなく「生活が動きました」で語る。\n　鍼灸師の聴衆には、こちらのほうが圧倒的に刺さります。\n\n★締めの2行が、次の『私の失敗（ひとりでやると壊れます）』への橋です。\n　抱え込まないこと、外に繋ぐこと。ここを言っておくと、\n　後半のサロン／ここちはりの話が『仲間が要る』という文脈で通ります。\n\n※個別の症例を語る場合は、個人が特定されない形で。掲載許可の確認を。");
}

// 46 私の失敗（濃色）
{
  const s = S(true);
  eyebrow(s, "明日からの臨床", true);
  title(s, "ただ、私は一度 失敗しています", true);
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 2.1, w: CW, h: 2.3, rectRadius: 0.1, fill: { color: "1B4E36" } });
  body(s, "聴きすぎて、自分が消耗した時期があります。\nカウンセラーになろうとして、施術がぶれた時期もありました。\n\n傾聴には、自分を守る技術がセットで必要です。",
    M + 0.6, 2.5, CW - 1.2, 1.6, { fontSize: 18, color: ONDARK, lineSpacing: 34 });
  s.addText("ひとりでやると、壊れます。", {
    x: M, y: 4.75, w: CW, h: 0.7, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 30, bold: true, color: PINKB,
  });
  note(s, "【49:00-50:30】\n★ここが、後半への唯一の橋。飛ばすと最後がただの宣伝になる。\n弱さを見せることで、次のスライドの提案が『売り込み』でなく『招待』になる。\n声を落として、ゆっくり。");
}

/* ══════════════════════════════════════════════════════════
   G. 続けるための場
   ══════════════════════════════════════════════════════════ */

// 47 独学で伸びない
{
  const s = S(false);
  eyebrow(s, "続けるために", false);
  title(s, "傾聴が、独学で伸びない3つの理由");
  const r = [["自分の聴き方は、自分では見えない", "フィードバックしてくれる人が要る"], ["聴きっぱなしだと、消耗する", "受け止めてもらう側が要る"], ["3週間で、元に戻る", "続ける仕組みが要る"]];
  r.forEach((it, i) => {
    const y = 2.1 + i * 1.05;
    card(s, M, y, CW, 0.88);
    circle(s, String(i + 1), M + 0.3, y + 0.17, 0.54, INK, PAPER, 16);
    s.addText(it[0], { x: M + 1.05, y, w: 5.4, h: 0.88, isTextBox: true, margin: 0, fontFace: GO, fontSize: 16, bold: true, color: INK, valign: "middle" });
    s.addText("→　" + it[1], { x: M + 6.6, y, w: CW - 6.9, h: 0.88, isTextBox: true, margin: 0, fontFace: GO, fontSize: 15, color: SHU, valign: "middle" });
  });
  s.addText("才能の話ではありません。環境の話です。", {
    x: M, y: 5.4, w: CW, h: 0.6, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 26, bold: true, color: INK,
  });
  note(s, "【50:30-52:00】\n★『できない自分が悪い』を『環境がなかっただけ』に変える。\nここまで来て初めて、場の提案が自然になる。");
}

// 48 サロン
{
  const s = S(false);
  eyebrow(s, "続けるために", false);
  title(s, "だから、場をつくりました");
  card(s, M, 2.05, CW * 0.52, 2.9);
  s.addText("オンラインサロン\nここちめいど", { x: M + 0.5, y: 2.4, w: CW * 0.52 - 1.0, h: 0.95, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 25, bold: true, color: INK, lineSpacing: 38 });
  body(s, "傾聴を学ぶ、鍼灸師のコミュニティ\n2020年4月から、5年目です。", M + 0.5, 3.5, CW * 0.52 - 1.0, 0.8, { fontSize: 14, color: MUTED, lineSpacing: 24 });
  s.addText("cocochimade.me", { x: M + 0.5, y: 4.4, w: CW * 0.52 - 1.0, h: 0.35, isTextBox: true, margin: 0, fontFace: GO, fontSize: 12, color: MORI });
  const rgt = M + CW * 0.52 + 0.5;
  const rw = CW - CW * 0.52 - 0.5;
  ["学ぶだけでは、終わらせない", "書く・出す・立つ　まで伴走する", "5年間で、学会発表の体制をつくった"].forEach((t, i) => {
    const y = 2.25 + i * 0.95;
    circle(s, "◯", rgt, y, 0.42, i === 1 ? SHU : MORI, PAPER, 12);
    s.addText(t, { x: rgt + 0.6, y: y - 0.03, w: rw - 0.6, h: 0.5, isTextBox: true, margin: 0, fontFace: GO, fontSize: 15.5, bold: true, color: INK, valign: "middle" });
  });
  s.addText("2026年 全日本鍼灸学会\n「鍼灸師コミュニティの学会発表体制構築 ―5年間の分析―」", {
    x: rgt, y: 5.15, w: rw, h: 0.7, isTextBox: true, margin: 0, fontFace: GO, fontSize: 11.5, color: MUTED, lineSpacing: 19,
  });
  note(s, "【52:00-53:30】\n★『学ぶ場』ではなく『伴走する場』と言い切る。ここが差別化。\n※サロンの様子の写真（salon-zoom.png / training-workshop.jpg）を入れてもよい。");
}

// 49 メンバーに起きたこと
{
  const s = S(false);
  eyebrow(s, "続けるために", false);
  title(s, "メンバーに、起きたこと");
  const m = [["学会発表", "はじめての演題を、\n自分の名前で出した"], ["書籍出版", "サロンから\n著者が生まれた"], ["院が変わった", "予約が埋まり、\n紹介が増えた"]];
  m.forEach((it, i) => {
    const x = M + i * (CW / 3);
    const w = CW / 3 - 0.35;
    card(s, x, 2.15, w, 2.35);
    s.addText(it[0], { x: x + 0.35, y: 2.5, w: w - 0.7, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 22, bold: true, color: SHU });
    s.addText(it[1], { x: x + 0.35, y: 3.1, w: w - 0.7, h: 1.1, isTextBox: true, margin: 0, fontFace: GO, fontSize: 14, color: INK, lineSpacing: 25 });
  });
  s.addText("全部、私ではなく、メンバーがやったことです。", {
    x: M, y: 4.95, w: CW, h: 0.6, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 24, bold: true, color: INK,
  });
  note(s, "【53:30-54:30】\n★主語を必ずメンバーにする。自分の実績にしない。\n『あの人にできたなら、私にも』という自己投影が、問合せを生む。\n※許可済みの写真や声を1〜2つ入れると、さらに効きます。");
}

// 50 ここちはり
{
  const s = S(false);
  eyebrow(s, "続けるために", false);
  title(s, "1対1で伴走してほしい方へ");
  card(s, M, 2.05, CW, 2.75);
  s.addText("ここちはり", { x: M + 0.55, y: 2.35, w: 4.0, h: 0.6, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 28, bold: true, color: INK });
  s.addText("施術者育成・コンサル／2025年11月開始", { x: M + 0.55, y: 3.0, w: 5.5, h: 0.4, isTextBox: true, margin: 0, fontFace: GO, fontSize: 13, color: MUTED });
  ["感情", "心", "身体", "技術"].forEach((t, i) => {
    circle(s, t, M + 0.55 + i * 1.15, 3.6, 0.95, i % 2 === 0 ? SHU : MORI, PAPER, 15);
  });
  body(s, "この4つの軸で、あなたの臨床を一緒に見ていきます。\n症例データベースも、公開して運用しています。",
    M + 5.6, 3.55, CW - 6.2, 1.1, { fontSize: 15, lineSpacing: 27 });
  note(s, "【54:30-56:00】\n★ここで初めて『メンター』という言葉を出す。\n「1対1で伴走してほしい、という方には こちらがあります」\n売り込まない。選択肢として置くだけ。");
}

/* ══════════════════════════════════════════════════════════
   H. クロージング
   ══════════════════════════════════════════════════════════ */

// 51 ひとつだけ
{
  const s = S(true);
  s.addText("今日、ひとつだけ持ち帰るなら", { x: M, y: 1.9, w: CW, h: 0.6, isTextBox: true, margin: 0, fontFace: GO, fontSize: 18, color: MUTED_D });
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 2.7, w: CW, h: 1.75, rectRadius: 0.12, fill: { color: PINKB } });
  s.addText("最初の90秒、遮らない。", {
    x: M, y: 2.7, w: CW, h: 1.75, isTextBox: true, margin: 0, align: "center", valign: "middle",
    fontFace: MIN, fontSize: 42, bold: true, color: "FFFFFF",
  });
  s.addText("これだけです。明日の、1人目からできます。", {
    x: M, y: 4.75, w: CW, h: 0.5, isTextBox: true, margin: 0, align: "center", fontFace: GO, fontSize: 18, color: ONDARK,
  });
  note(s, "【56:00-57:00】\n★持ち帰りは必ず1つに絞る。3つ渡すと0個になる。\n数字（90秒）が入っていることが重要。抽象的な標語は持ち帰れない。");
}

// 52 スルーラインで閉じる
{
  const s = S(true);
  s.addShape(pres.ShapeType.ellipse, { x: 4.7, y: 0.6, w: 6.3, h: 6.3, fill: { color: "18452F" } });
  s.addText("経絡に触れ、\nこころに耳を傾ける。", {
    x: 1.5, y: 2.3, w: 10.3, h: 1.9, isTextBox: true, margin: 0, align: "center",
    fontFace: MIN, fontSize: 40, bold: true, color: ONDARK, lineSpacing: 66,
  });
  s.addText("それが、私たちの持っているチカラです。", {
    x: 1.5, y: 4.5, w: 10.3, h: 0.6, isTextBox: true, margin: 0, align: "center",
    fontFace: MIN, fontSize: 22, bold: true, color: PINKB,
  });
  note(s, "【57:00-57:40】\n冒頭（12枚目）と同じ言葉で閉じる。\n読んだあと、3秒 黙る。ここで拍手が来る設計。");
}

// 53 CTA
{
  const s = S(false);
  title(s, "その質問、そのまま送ってください", false, 0.75);
  card(s, M, 1.95, CW * 0.56, 3.6);
  body(s, "「うちの院では、どうしたらいいんだろう」\n\n今日の話で、そう思った方。\nその質問を、そのまま送ってください。\n\n私が、全部読みます。",
    M + 0.55, 2.35, CW * 0.56 - 1.1, 2.6, { fontSize: 18, lineSpacing: 34 });
  const qx = M + CW * 0.56 + 0.55;
  const qw = CW - CW * 0.56 - 0.55;
  s.addShape(pres.ShapeType.roundRect, { x: qx, y: 1.95, w: qw, h: 3.6, rectRadius: 0.1, fill: { color: PAPER }, line: { color: SHU, width: 1.5, dashType: "dash" } });
  s.addText("▶ 公式LINEのQRコードを\nここに大きく貼ってください", {
    x: qx + 0.3, y: 3.2, w: qw - 0.6, h: 0.9, isTextBox: true, margin: 0, align: "center",
    fontFace: GO, fontSize: 13, bold: true, color: MUTED, lineSpacing: 22,
  });
  s.addText("▶ 短縮URLをここに（口頭でも読み上げる）", {
    x: qx + 0.3, y: 4.75, w: qw - 0.6, h: 0.4, isTextBox: true, margin: 0, align: "center",
    fontFace: GO, fontSize: 11, color: MUTED,
  });
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 5.75, w: CW, h: 0.95, rectRadius: 0.1, fill: { color: MINT } });
  s.addText("録画でご覧の方も、同じように送ってください。全部読みます。　／　本日のスライドをお送りします。", {
    x: M + 0.4, y: 5.75, w: CW - 0.8, h: 0.95, isTextBox: true, margin: 0, valign: "middle",
    fontFace: GO, fontSize: 15, bold: true, color: INK,
  });
  note(s, "【57:40-60:00】\n★このスライドを30秒 黙って映す。読み取り待ち。ここを急ぐと激減します。\n★Q&Aの間も、このスライドを出しっぱなしにする。\n★チャットにも同じリンクを投下（事前にメモ帳に用意してコピペ／共同ホストに依頼）。\n　ただしチャットは録画に残らないので、画面表示が本体です。\n★短縮URLは口頭でも読み上げる（音声だけで聴いている人がいます）。\n\n【Q&Aのコツ】\n質問には即答せず、まず「その患者さん、なんて仰ってました？」と聴き返す。\n傾聴の実演が、いちばん強い営業になります。");
}

const TIMINGS = {1: "0:00-0:20", 2: "0:20-1:20", 3: "1:20-2:50", 4: "2:50-4:10", 5: "4:10-5:20", 6: "5:20-6:15", 7: "6:15-7:20", 8: "7:20-8:30", 9: "8:30-9:25", 10: "9:25-9:55", 11: "9:55-11:05", 12: "11:05-12:25", 13: "12:25-13:45", 14: "13:45-15:30", 15: "15:30-16:50", 16: "16:50-17:30", 17: "17:30-18:45", 18: "18:45-19:20", 19: "19:20-20:25", 20: "20:25-21:45", 21: "21:45-22:50", 22: "22:50-23:55", 23: "23:55-24:55", 24: "24:55-25:20", 25: "25:20-26:00", 26: "26:00-27:00", 27: "27:00-28:30", 28: "28:30-29:00", 29: "29:00-30:20", 30: "30:20-30:50", 41: "31:00-31:50", 42: "31:50-32:30", 43: "32:30-36:10", 44: "36:10-36:40", 45: "36:40-39:00", 46: "39:00-39:30", 47: "39:30-40:10", 48: "40:10-41:50", 49: "41:50-43:10", 50: "43:10-44:30", 51: "44:30-45:40", 52: "45:40-46:50", 53: "46:50-48:20", 54: "48:20-49:30", 55: "49:30-50:10", 56: "50:10-51:30", 57: "51:30-52:40", 58: "52:40-53:50", 59: "53:50-54:45", 60: "54:45-55:45", 61: "55:45-56:40", 62: "56:40-57:30", 63: "57:30-60:00"};

const TOP = 0.34, BOTTOM = 6.92;

for (const rec of decks) {
  const ys = rec.items.map((it) => it.o.y);
  const bs = rec.items.map((it) => it.o.y + it.o.h);
  const minT = Math.min(...ys), maxB = Math.max(...bs);
  let map;
  if (minT <= 0.85) {
    // 上端そろえの構成：下端が 6.92" に届くまで縦に伸ばす
    const k = Math.max(1, Math.min((BOTTOM - TOP) / (maxB - TOP), 1.34));
    map = (y, h, isEllipse) => {
      if (isEllipse) { const c = TOP + (y + h / 2 - TOP) * k; return [c - h / 2, h]; }
      return [TOP + (y - TOP) * k, h * k];
    };
  } else {
    // 中央寄せの構成：ブロックごと上下中央に置き直す
    const dy = (H - (maxB - minT)) / 2 - minT;
    map = (y, h) => [y + dy, h];
  }

  const s = pres.addSlide();
  s.background = { color: rec.dark ? INK : CREAM };
  for (const it of rec.items) {
    const isEllipse = it.k === "shape" && it.st === pres.ShapeType.ellipse;
    const [ny, nh] = map(it.o.y, it.o.h, isEllipse);
    const o = Object.assign({}, it.o, { y: ny, h: nh });
    if (it.k === "text") s.addText(it.t, o);
    else if (it.k === "chart") s.addChart(it.ct, it.d, o);
    else s.addShape(it.st, o);
  }
  if (rec.num > 1) s.addText(String(rec.num), {
    x: W - 0.85, y: H - 0.5, w: 0.5, h: 0.3, isTextBox: true, margin: 0,
    align: "right", fontFace: GO, fontSize: 10, color: rec.dark ? "3A6450" : "AFD2C1",
  });
  if (rec.notes) {
    let nt = rec.notes;
    if (TIMINGS[rec.num]) nt = nt.replace(/^【[^】]*】/, "【" + TIMINGS[rec.num] + "】");
    s.addNotes(nt);
  }
}

const OUT = process.argv[2] || "APNET-keicho-60min.pptx";
pres.writeFile({ fileName: OUT }).then(() => console.log("slides:", n, "->", OUT));
