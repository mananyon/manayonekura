const pptxgen = require("pptxgenjs");
const path = require("path");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "米倉まな";
pres.title = "鍼灸師のための傾聴セミナー";
pres.subject = "オンラインサロン ここちめいど";

// ── ブランドトークン（「ここちめいど」ロゴ由来）────────────────
const CREAM = "F0FAF5", PAPER = "FFFFFF", INK = "123C2A";
const GREEN = "17784A";   // 主アクセント
const PINK = "C92A78";    // 副アクセント（小さい文字でも読める濃さ）
const PINKB = "F15C9E";   // ロゴのピンク。濃色スライドの見せ場だけ
const MINT = "34D39B";    // 面で使う。文字は INK
const MUTED = "4F6F60", MUTED_D = "8FB8A2", ONDARK = "E8F5EE";
const MIN = "游明朝", GO = "游ゴシック";
const ASSETS = path.join(__dirname, "assets");

const W = 13.333, H = 7.5, M = 0.8, CW = W - M * 2;
let n = 0;
const decks = [];

function S(dark) {
  n++;
  const rec = { dark, num: n, items: [], notes: null };
  decks.push(rec);
  return {
    addText: (t, o) => rec.items.push({ k: "text", t, o }),
    addShape: (st, o) => rec.items.push({ k: "shape", st, o }),
    addImage: (o) => rec.items.push({ k: "image", o }),
    addChart: (ct, d, o) => rec.items.push({ k: "chart", ct, d, o }),
    addNotes: (t) => { rec.notes = t; },
  };
}
const note = (s, t) => s.addNotes(t);

function title(s, txt, dark, y) {
  s.addText(txt, {
    x: M, y: y === undefined ? 0.62 : y, w: CW, h: 1.05, isTextBox: true, margin: 0,
    fontFace: MIN, fontSize: 32, bold: true, color: dark ? ONDARK : INK,
    align: "left", valign: "middle", lineSpacing: 42,
  });
}
function eyebrow(s, txt, dark) {
  s.addText(txt, {
    x: M, y: 0.34, w: CW, h: 0.26, isTextBox: true, margin: 0,
    fontFace: GO, fontSize: 11, bold: true, charSpacing: 2,
    color: dark ? MUTED_D : GREEN, align: "left",
  });
}
function circle(s, label, x, y, d, fill, txtColor, size) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill } });
  s.addText(label, {
    x, y, w: d, h: d, isTextBox: true, margin: 0, align: "center", valign: "middle",
    fontFace: MIN, fontSize: size || 20, bold: true, color: txtColor || PAPER,
  });
}
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
// 見出し＋一行の締め、を持つ標準レイアウト用
function kicker(s, txt, y, color, size) {
  s.addText(txt, {
    x: M, y, w: CW, h: 0.6, isTextBox: true, margin: 0,
    fontFace: MIN, fontSize: size || 24, bold: true, color: color || INK,
  });
}
function source(s, txt, y) {
  s.addText(txt, { x: M, y, w: CW, h: 0.35, isTextBox: true, margin: 0, fontFace: GO, fontSize: 10, color: MUTED });
}

/* ══════════ A. つかみ ══════════ */

// 1 タイトル
{
  const s = S(true);
  s.addShape(pres.ShapeType.ellipse, { x: 9.3, y: 1.0, w: 5.6, h: 5.6, fill: { color: "18452F" } });
  s.addText("鍼灸師のための", { x: M, y: 2.0, w: 8, h: 0.4, isTextBox: true, margin: 0, fontFace: GO, fontSize: 14, bold: true, charSpacing: 3, color: MUTED_D });
  s.addText("傾聴セミナー", {
    x: M, y: 2.5, w: 8.6, h: 1.3, isTextBox: true, margin: 0,
    fontFace: MIN, fontSize: 54, bold: true, color: ONDARK,
  });
  s.addText("患者さんから、情報が集まるようになる。", {
    x: M, y: 4.0, w: 8.6, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 20, bold: true, color: MINT,
  });
  s.addText("米倉まな（よねくら まな）\nはりきゅう処ここちめいど 院長／オンラインサロン ここちめいど 主宰\nはり師・きゅう師／産業カウンセラー／修士（人間学）", {
    x: M, y: 5.0, w: 8.6, h: 1.1, isTextBox: true, margin: 0, fontFace: GO, fontSize: 12, color: MUTED_D, lineSpacing: 20,
  });
  note(s, "【0:00-0:30】\n開始前に「紙とペンをご用意ください」を案内。\n後半にワークがあることを先に伝えておくと、離脱が減ります。");
}

// 2 こんなことありませんか
{
  const s = S(false);
  eyebrow(s, "はじめに", false);
  title(s, "こんなこと、ありませんか");
  const a = [
    "患者さんへ、どうアドバイスしたらいいか分からない",
    "なぜ鍼灸院に来ているんだろう？という方がいる",
    "患者さんに「良くなった」と言われたい",
    "患者さんの施術後、自分が疲れてしまう",
  ];
  a.forEach((t, i) => {
    const y = 2.1 + i * 0.86;
    card(s, M, y, CW, 0.74);
    circle(s, String(i + 1), M + 0.3, y + 0.17, 0.4, i % 2 ? GREEN : PINK, PAPER, 13);
    s.addText(t, { x: M + 0.95, y, w: CW - 1.3, h: 0.74, isTextBox: true, margin: 0, valign: "middle", fontFace: GO, fontSize: 17, color: INK });
  });
  note(s, "【0:30-2:30】\n4つとも読み上げて、「ひとつでも心当たりがある方？」と手を挙げてもらう\n（オンラインならリアクションかチャットに数字）。\n★ここで全員を当事者にする。急がない。");
}

// 3 答えはあります
{
  const s = S(true);
  s.addText("答えは、あります。", {
    x: M, y: 3.0, w: CW, h: 1.3, isTextBox: true, margin: 0,
    fontFace: MIN, fontSize: 52, bold: true, color: ONDARK,
  });
  s.addText("4つとも、同じひとつの技術で変わります。", {
    x: M, y: 4.5, w: CW, h: 0.5, isTextBox: true, margin: 0, fontFace: GO, fontSize: 18, color: MINT,
  });
  note(s, "【2:30-2:50】\n言い切って、一拍おく。ここでフックが効きます。");
}

// 4 自己紹介
{
  const s = S(false);
  eyebrow(s, "はじめに", false);
  title(s, "私は、聴いてもらう側でした");
  card(s, M, 2.1, CW * 0.54, 2.45);
  s.addText("うつ病・パニック障害・双極性障害", {
    x: M + 0.5, y: 2.42, w: CW * 0.54 - 1.0, h: 0.4, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 18, bold: true, color: INK,
  });
  body(s, "最重症期は、1日30錠。患者だった期間は8年。\nいまは16年間、寛解。投薬もありません。",
    M + 0.5, 2.92, CW * 0.54 - 1.0, 1.2, { fontSize: 14.5, lineSpacing: 27 });
  const rx = M + CW * 0.54 + 0.6, rw = CW * 0.46 - 0.6;
  body(s, "三重県四日市市／東京都町田市\nはりきゅう処ここちめいど 院長\nオンラインサロン ここちめいど 主宰\n精神疾患臨床家グループ ここちはり 代表\n産業カウンセラー／修士（人間学）",
    rx, 2.25, rw, 2.0, { fontSize: 13, lineSpacing: 25, color: MUTED });
  kicker(s, "だから、聴いてもらえないつらさを知っています。", 4.9, PINK, 22);
  note(s, "【2:50-5:00】\n★当事者経験は、話すかどうかご自身で決めてください。削除しても成立します。\n淡々と。同情を求めない。最後の一行に重心を置く。");
}

/* ══════════ B. なぜ傾聴か ══════════ */

// 5 前提
{
  const s = S(true);
  eyebrow(s, "なぜ、傾聴なのか", true);
  s.addText("人は、分かり合えない。", {
    x: M, y: 2.3, w: CW, h: 1.0, isTextBox: true, margin: 0,
    fontFace: MIN, fontSize: 44, bold: true, color: ONDARK,
  });
  s.addText("察するなんて、無理。", {
    x: M, y: 3.5, w: CW, h: 1.0, isTextBox: true, margin: 0,
    fontFace: MIN, fontSize: 44, bold: true, color: MUTED_D,
  });
  s.addText("これが、今日の前提です。", {
    x: M, y: 4.9, w: CW, h: 0.5, isTextBox: true, margin: 0, fontFace: GO, fontSize: 17, color: MINT,
  });
  note(s, "【5:00-6:30】\n★ここを最初に置くと、後の全部が楽になります。\n「分かってあげなきゃ」と思っている人ほど疲れている。\nその前提を外してあげるのが、このスライドの役目。");
}

// 6 だから対話
{
  const s = S(false);
  eyebrow(s, "なぜ、傾聴なのか", false);
  title(s, "分からないから、教えてもらう");
  card(s, M, 2.15, CW, 2.1);
  s.addText("分からない", {
    x: M + 0.7, y: 2.75, w: 3.0, h: 0.8, isTextBox: true, margin: 0, valign: "middle",
    fontFace: MIN, fontSize: 32, bold: true, color: MUTED,
  });
  s.addText("→", { x: M + 3.9, y: 2.75, w: 0.8, h: 0.8, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: GO, fontSize: 24, color: MUTED });
  s.addText("教えてほしい", {
    x: M + 4.9, y: 2.75, w: 3.4, h: 0.8, isTextBox: true, margin: 0, valign: "middle",
    fontFace: MIN, fontSize: 32, bold: true, color: INK,
  });
  s.addText("＝", { x: M + 8.4, y: 2.75, w: 0.6, h: 0.8, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: GO, fontSize: 22, color: MUTED });
  s.addText("対　話", {
    x: M + 9.2, y: 2.75, w: 2.4, h: 0.8, isTextBox: true, margin: 0, valign: "middle",
    fontFace: MIN, fontSize: 34, bold: true, color: GREEN,
  });
  kicker(s, "察するのをやめて、聞く。それだけで臨床は変わります。", 4.7, INK, 23);
  note(s, "【6:30-7:30】\n「分からない」を認めることが出発点。\n分かったふりをした瞬間に、情報は取れなくなります。");
}

// 7 どうすればわかるのか
{
  const s = S(false);
  eyebrow(s, "なぜ、傾聴なのか", false);
  title(s, "患者さんが求めていることは、どうすれば分かる？");
  const no = ["経験でなんとなく察する", "問診票から推測する", "同じ主訴の人と同じだと考える"];
  no.forEach((t, i) => {
    const x = M + i * (CW / 3);
    const w = CW / 3 - 0.35;
    card(s, x, 2.15, w, 1.5);
    circle(s, "✕", x + w / 2 - 0.26, 2.4, 0.52, MUTED, PAPER, 16);
    s.addText(t, { x: x + 0.25, y: 3.05, w: w - 0.5, h: 0.5, isTextBox: true, margin: 0, align: "center", fontFace: GO, fontSize: 13, color: INK });
  });
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 4.1, w: CW, h: 1.1, rectRadius: 0.1, fill: { color: MINT } });
  s.addText("聞くしか、ありません。", {
    x: M, y: 4.1, w: CW, h: 1.1, isTextBox: true, margin: 0, align: "center", valign: "middle",
    fontFace: MIN, fontSize: 32, bold: true, color: INK,
  });
  note(s, "【7:30-8:30】\n3つの✕は、みんなやっていること。だから責めない。\n「私もやっていました」と添えると角が立ちません。");
}

// 8 鍼灸師にできるカウンセリング技法
{
  const s = S(false);
  eyebrow(s, "なぜ、傾聴なのか", false);
  title(s, "鍼灸師にできる、カウンセリング技法");
  card(s, M, 2.2, CW, 2.0);
  s.addText("傾　聴", {
    x: M, y: 2.2, w: CW, h: 2.0, isTextBox: true, margin: 0, align: "center", valign: "middle",
    fontFace: MIN, fontSize: 60, bold: true, color: GREEN,
  });
  const p = [["資格は要りません", "ただし練習は必要です"], ["診断ではありません", "情報を集める技術です"], ["時間は増えません", "使い方が変わるだけです"]];
  p.forEach((t, i) => {
    const x = M + i * (CW / 3);
    s.addText(t[0], { x, y: 4.6, w: CW / 3 - 0.4, h: 0.4, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 17, bold: true, color: INK });
    s.addText(t[1], { x, y: 5.05, w: CW / 3 - 0.4, h: 0.4, isTextBox: true, margin: 0, fontFace: GO, fontSize: 13, color: MUTED });
  });
  note(s, "【8:30-9:30】\n★「時間は増えません」が効きます。忙しい人ほど身構えているので、\n　最初に負担が増えないことを保証する。");
}

// 9 2つの情報が同時に集まる
{
  const s = S(false);
  eyebrow(s, "傾聴のメリット ①", false);
  title(s, "傾聴すると、2種類の情報が同時に集まります");
  const cols = [
    ["東洋医学的な情報", ["睡眠　食欲　便通　月経", "情志　冷え　のぼせ", "生活リズム　仕事と家族"], GREEN],
    ["西洋医学的な情報", ["レッドフラッグの兆候", "服薬状況と、薬への不安", "生活機能・就労の状況"], PINK],
  ];
  cols.forEach((c, i) => {
    const x = M + i * (CW / 2 + 0.15);
    const w = CW / 2 - 0.15;
    card(s, x, 2.15, w, 2.55);
    s.addText(c[0], { x: x + 0.45, y: 2.45, w: w - 0.9, h: 0.45, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 21, bold: true, color: c[2] });
    s.addText(c[1].join("\n"), { x: x + 0.45, y: 3.05, w: w - 0.9, h: 1.4, isTextBox: true, margin: 0, fontFace: GO, fontSize: 15, color: INK, lineSpacing: 30 });
  });
  kicker(s, "一度の問診で、両方が取れます。", 5.15, INK, 26);
  note(s, "【9:30-11:00】\n★このセミナーの一番の売り。\n　「傾聴＝優しさ」ではなく「傾聴＝情報収集」だと最初に位置づける。\n　鍼灸師が動くのは、優しさではなく情報です。");
}

// 10 東洋医学：四診の問
{
  const s = S(false);
  eyebrow(s, "傾聴のメリット ①", false);
  title(s, "「証」は、聞かなければ立ちません");
  s.addText("四診", { x: M, y: 2.1, w: 2.0, h: 0.35, isTextBox: true, margin: 0, fontFace: GO, fontSize: 12, bold: true, charSpacing: 2, color: MUTED });
  [["望", "見る"], ["聞", "聞く・嗅ぐ"], ["問", "尋ねる"], ["切", "触れる"]].forEach((it, i) => {
    const x = M + i * 1.2, on = it[0] === "問";
    circle(s, it[0], x, 2.55, 1.0, on ? GREEN : "CDEBDC", on ? PAPER : INK, 28);
    s.addText(it[1], { x: x - 0.1, y: 3.68, w: 1.2, h: 0.35, isTextBox: true, margin: 0, align: "center", fontFace: GO, fontSize: 11, bold: on, color: on ? GREEN : MUTED });
  });
  card(s, M + 5.2, 2.2, CW - 5.2, 2.4);
  s.addText("望・聞・切は、自分で取れる。", { x: M + 5.7, y: 2.5, w: CW - 6.2, h: 0.4, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 18, bold: true, color: INK });
  s.addText("けれど「問」だけは、\n患者さんに話してもらわないと\n1ミリも進みません。", {
    x: M + 5.7, y: 3.0, w: CW - 6.2, h: 1.2, isTextBox: true, margin: 0, fontFace: GO, fontSize: 15, color: INK, lineSpacing: 28,
  });
  kicker(s, "東洋医学ができていない、のではなく ―― 聞けていないだけ、かもしれません。", 5.1, INK, 21);
  note(s, "【11:00-12:30】\n四診のうち「問」だけ色を変えているのがポイント。\n証が立たないのは腕のせいではなく、情報が足りないだけ、という視点の転換。");
}

// 11 西洋医学：レッドフラッグ
{
  const s = S(false);
  eyebrow(s, "傾聴のメリット ①", false);
  title(s, "聴いていたから、気づけました");
  const talk = [
    [0, "最近、気分が落ち込んでいて、眠れません"],
    [1, "落ち込みと、眠れない。日中の眠気はありますか？"],
    [0, "眠気はないんですが、時々ぼーっとして、誰かいるように感じたり、虫が見えるような気がして"],
    [1, "（あれっ）手が震えたり、お薬が効きすぎることはありませんか？"],
    [0, "あります。うつのお薬を飲むと、逆効果に感じるんです"],
  ];
  talk.forEach((t, i) => {
    const mine = t[0] === 1, w = CW * 0.76;
    const x = mine ? M + CW - w : M, y = 2.05 + i * 0.72;
    if (mine) s.addShape(pres.ShapeType.roundRect, { x, y, w, h: 0.6, rectRadius: 0.1, fill: { color: INK } });
    else card(s, x, y, w, 0.6);
    s.addText(t[1], { x: x + 0.35, y, w: w - 0.7, h: 0.6, isTextBox: true, margin: 0, valign: "middle", fontFace: GO, fontSize: 13, color: mine ? ONDARK : INK });
  });
  kicker(s, "神経内科にご高診を依頼　→　レビー小体型認知症でした。", 5.8, PINK, 22);
  note(s, "【12:30-14:30】\n★このセミナーで一番強い実例。声色を変えて演じると引き込めます。\n傾聴＝優しさ ではなく 傾聴＝安全性。\n聴いていなければ、幻視も手の震えも出てこなかった。");
}

// 12 ご高診後の診断名
{
  const s = S(false);
  eyebrow(s, "傾聴のメリット ①", false);
  title(s, "抑うつを訴えて来られた方の、ご高診後の診断名");
  const g = [
    ["脳神経系", "パーキンソン病／レビー小体型認知症／アルツハイマー病／脳腫瘍"],
    ["内分泌系", "甲状腺機能亢進症／甲状腺機能低下症"],
    ["性ホルモン系", "更年期障害／月経前不快気分障害／産後うつ／男性更年期"],
    ["生活習慣・循環器系", "高血圧／糖尿病／虚血性心疾患／微小血管狭心症　など"],
  ];
  g.forEach((it, i) => {
    const y = 2.15 + i * 0.82;
    card(s, M, y, CW, 0.7);
    s.addText(it[0], { x: M + 0.4, y, w: 2.9, h: 0.7, isTextBox: true, margin: 0, valign: "middle", fontFace: MIN, fontSize: 15, bold: true, color: GREEN });
    s.addText(it[1], { x: M + 3.5, y, w: CW - 3.9, h: 0.7, isTextBox: true, margin: 0, valign: "middle", fontFace: GO, fontSize: 13.5, color: INK });
  });
  kicker(s, "聴かなければ、どれも見つかりません。", 5.7, INK, 24);
  source(s, "当院事例", 6.35);
  note(s, "【14:30-15:30】\n読み上げない。目で見せるだけで十分効きます。\n「うつだと思っていた方が、これだけ違う病気だった」という事実だけ伝える。");
}

// 13 実費の強み
{
  const s = S(false);
  eyebrow(s, "傾聴のメリット ②", false);
  title(s, "情報を集められることが、実費診療の強みです");
  const adv = [["40〜60分", "ふたりきりの時間がある"], ["身体に触れる", "言葉の前に、からだの情報がある"], ["会い続ける", "一度きりではない"], ["制約が少ない", "時間の使い方を自分で決められる"]];
  adv.forEach((it, i) => {
    const x = M + (i % 2) * (CW / 2);
    const y = 2.15 + Math.floor(i / 2) * 1.25;
    circle(s, "●", x + 0.05, y + 0.16, 0.32, i % 2 === 0 ? GREEN : PINK, PAPER, 9);
    s.addText(it[0], { x: x + 0.6, y: y + 0.02, w: CW / 2 - 0.9, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 21, bold: true, color: INK });
    s.addText(it[1], { x: x + 0.6, y: y + 0.55, w: CW / 2 - 0.9, h: 0.45, isTextBox: true, margin: 0, fontFace: GO, fontSize: 14, color: MUTED });
  });
  card(s, M, 4.9, CW, 1.0);
  s.addText("医師が患者の話を遮るまで、中央値 11秒。　私たちには、40分あります。", {
    x: M + 0.5, y: 4.9, w: CW - 1.0, h: 1.0, isTextBox: true, margin: 0, valign: "middle",
    fontFace: MIN, fontSize: 20, bold: true, color: INK,
  });
  source(s, "Singh Ospina et al., 2019 ／ Beckman & Frankel, 1984（いずれも医師を対象とした研究）", 6.05);
  note(s, "【15:30-17:00】\n★11秒は医師対象の研究だと必ず断る。鍼灸師のデータではありません。\n「時間があるのに使えていない」のがもったいない、という話に持っていく。");
}

// 14 医師からの評価
{
  const s = S(false);
  eyebrow(s, "傾聴のメリット ②", false);
  title(s, "医師の先生方から、こう言われます");
  [["「患者さんと、\n　時間がある」", GREEN], ["「身体全体を、\n　見てくれる」", PINK], ["「東洋医学の、\n　専門家だから」", GREEN]].forEach((it, i) => {
    const x = M + i * (CW / 3), w = CW / 3 - 0.35;
    card(s, x, 2.15, w, 2.0);
    circle(s, "●", x + 0.35, 2.45, 0.3, it[1], PAPER, 9);
    s.addText(it[0], { x: x + 0.35, y: 3.0, w: w - 0.7, h: 0.95, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 20, bold: true, color: INK, lineSpacing: 32 });
  });
  kicker(s, "期待されている、ということでもあります。", 4.75, INK, 23);
  note(s, "【17:00-18:00】\n★調査データではなく、ご自身が言われてきた実感として話すこと。\n誇らしい気持ちで読み上げてよい。");
}

/* ══════════ C. 治すだけが全てではない ══════════ */

// 15 鍼灸院にいる時間
{
  const s = S(false);
  eyebrow(s, "抱えない、という考え方", false);
  title(s, "鍼灸院にいる時間は、人生のごく一部です");
  card(s, M, 2.2, CW * 0.44, 2.3);
  s.addText("週に1回・45分", { x: M + 0.5, y: 2.55, w: CW * 0.44 - 1.0, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 24, bold: true, color: GREEN });
  s.addText("院にいる時間", { x: M + 0.5, y: 3.1, w: CW * 0.44 - 1.0, h: 0.4, isTextBox: true, margin: 0, fontFace: GO, fontSize: 13, color: MUTED });
  s.addShape(pres.ShapeType.roundRect, { x: M + CW * 0.46, y: 2.2, w: CW * 0.54, h: 2.3, rectRadius: 0.1, fill: { color: INK } });
  s.addText("残りの 10,035分", { x: M + CW * 0.46 + 0.5, y: 2.55, w: CW * 0.54 - 1.0, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 24, bold: true, color: MINT });
  s.addText("患者さんが、ひとりで生きている時間", { x: M + CW * 0.46 + 0.5, y: 3.1, w: CW * 0.54 - 1.0, h: 0.4, isTextBox: true, margin: 0, fontFace: GO, fontSize: 13, color: MUTED_D });
  s.addText("私たちが触れられるのは、0.4%です。", { x: M + CW * 0.46 + 0.5, y: 3.65, w: CW * 0.54 - 1.0, h: 0.4, isTextBox: true, margin: 0, fontFace: GO, fontSize: 14, color: ONDARK });
  kicker(s, "治すことだけが、全てではありません。", 5.0, INK, 26);
  note(s, "【18:00-19:30】\n1週間＝10,080分。うち45分が院にいる時間。残りは患者さんがひとりで過ごす。\n★ここで「じゃあ残りの時間に、私たちは何ができるのか」と問いを立てる。\n答えが、このあとの『依存先を増やす』です。");
}

// 16 困っているのは症状ではない
{
  const s = S(false);
  eyebrow(s, "抱えない、という考え方", false);
  title(s, "患者さんが困っているのは、症状ではありません");
  card(s, M, 2.15, CW, 1.5);
  s.addText("「うつ病の症状で困っているのではなく、\n　症状があることで、できなくなった何かに困っている」", {
    x: M + 0.6, y: 2.15, w: CW - 1.2, h: 1.5, isTextBox: true, margin: 0, valign: "middle",
    fontFace: MIN, fontSize: 21, bold: true, color: INK, lineSpacing: 36,
  });
  body(s, "・仕事に行けない　・家事ができない　・人と会えなくなった\n・以前できていた趣味が、できなくなった",
    M, 4.0, CW, 0.9, { fontSize: 15, color: MUTED, lineSpacing: 27 });
  kicker(s, "聴くのは、症状より「できなくなったこと」。", 5.15, PINK, 25);
  note(s, "【19:30-21:00】\n★問診の焦点を変える一枚。\n「痛みは10段階でいくつですか」ではなく「痛くて、何ができなくなりましたか」。\nこれだけで、出てくる情報が変わります。");
}

// 17 個人の課題／社会の課題
{
  const s = S(false);
  eyebrow(s, "抱えない、という考え方", false);
  title(s, "困りごとは、2つに分けて整理する");
  const cols = [
    ["個人の課題", "内省のサポートに、傾聴を", ["働く不安", "自信喪失", "職業適性", "家族・友人との関係性"], GREEN],
    ["社会の課題", "社会への働きかけへ", ["偏見", "職場環境（人材不足・受入体制）", "知られていない公的資源", "医療資源"], PINK],
  ];
  cols.forEach((c, i) => {
    const x = M + i * (CW / 2 + 0.15), w = CW / 2 - 0.15;
    card(s, x, 2.15, w, 2.85);
    s.addText(c[0], { x: x + 0.45, y: 2.45, w: w - 0.9, h: 0.45, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 22, bold: true, color: c[3] });
    s.addText(c[1], { x: x + 0.45, y: 2.92, w: w - 0.9, h: 0.35, isTextBox: true, margin: 0, fontFace: GO, fontSize: 12, color: MUTED });
    s.addText(c[2].map((t) => "・" + t).join("\n"), { x: x + 0.45, y: 3.4, w: w - 0.9, h: 1.4, isTextBox: true, margin: 0, fontFace: GO, fontSize: 14, color: INK, lineSpacing: 27 });
  });
  kicker(s, "分けると、抱え込まなくて済みます。", 5.4, INK, 24);
  note(s, "【21:00-22:30】\n★「全部なんとかしてあげたい」が一番危ない。\n分けた瞬間、自分がやることと、人に渡すことが見えます。");
}

// 18 依存先を増やす
{
  const s = S(true);
  eyebrow(s, "抱えない、という考え方", true);
  s.addText("自立とは、", { x: M, y: 2.3, w: CW, h: 0.8, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 34, bold: true, color: MUTED_D });
  s.addText("依存先を増やすこと。", {
    x: M, y: 3.2, w: CW, h: 1.1, isTextBox: true, margin: 0,
    fontFace: MIN, fontSize: 48, bold: true, color: ONDARK,
  });
  s.addText("患者さんを抱えるのではありません。つなぐのです。", {
    x: M, y: 4.6, w: CW, h: 0.5, isTextBox: true, margin: 0, fontFace: GO, fontSize: 18, color: MINT,
  });
  note(s, "【22:30-24:00】\n★このセミナーの背骨になる一行。ゆっくり読む。\n自分ひとりが支えると思うから、距離が近くなりすぎる。\n依存先が1つしかない状態が、いちばん危ない。");
}

// 19 連携先
{
  const s = S(false);
  eyebrow(s, "抱えない、という考え方", false);
  title(s, "連携先は、こんなにあります");
  const groups = [
    ["医療", ["精神科・心療内科", "その他の医療機関", "訪問看護ステーション", "薬剤師", "訪問鍼灸"], GREEN],
    ["生活・福祉", ["地域包括支援センター", "保健所", "社会福祉協議会", "全国健康保険協会", "療育"], PINK],
    ["就労・法律", ["ハローワーク", "地域若者サポートステーション", "社会保険労務士", "法律相談"], GREEN],
  ];
  groups.forEach((g, i) => {
    const x = M + i * (CW / 3), w = CW / 3 - 0.35;
    card(s, x, 2.15, w, 2.9);
    s.addText(g[0], { x: x + 0.4, y: 2.45, w: w - 0.8, h: 0.4, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 19, bold: true, color: g[2] });
    s.addText(g[1].map((t) => "・" + t).join("\n"), { x: x + 0.4, y: 2.95, w: w - 0.8, h: 1.9, isTextBox: true, margin: 0, fontFace: GO, fontSize: 13, color: INK, lineSpacing: 26 });
  });
  kicker(s, "全部知らなくていい。ひとつずつ増やせば大丈夫です。", 5.45, INK, 23);
  note(s, "【24:00-25:30】\n★「こんなに覚えられない」と思わせないこと。\n自分の地域で使ったことがある先を1つずつ増やしていけばいい、と伝える。\n※地域包括支援センターなど、ご自身が実際に連携した先を口頭で足すと生々しくなります。");
}

// 20 事例：傷病手当金
{
  const s = S(false);
  eyebrow(s, "つなぐ ― 事例", false);
  title(s, "聴いたから、つなげました");
  const talk = [
    [0, "病院に行きたくない。休職と言われたら困る"],
    [1, "休職すると困りますよね。〇〇さんにとって、どんなことが困りますか？"],
    [0, "働けなくなったらお金がなくなる。家族もいるから、自分は働かないと"],
    [1, "（加入条件を確認したうえで）それならお休みになった場合に、使える制度があるかもしれません。人事や協会けんぽで伺ってみては？"],
    [0, "そういう制度があるって知らなかった！　それなら休むことも出来るかもしれない"],
  ];
  talk.forEach((t, i) => {
    const mine = t[0] === 1, w = CW * 0.78;
    const x = mine ? M + CW - w : M, y = 2.05 + i * 0.78;
    if (mine) s.addShape(pres.ShapeType.roundRect, { x, y, w, h: 0.66, rectRadius: 0.1, fill: { color: INK } });
    else card(s, x, y, w, 0.66);
    s.addText(t[1], { x: x + 0.35, y, w: w - 0.7, h: 0.66, isTextBox: true, margin: 0, valign: "middle", fontFace: GO, fontSize: 12.5, color: mine ? ONDARK : INK });
  });
  kicker(s, "人事に相談され、通院・投薬を続けながら、仕事を続けられました。", 6.0, PINK, 20);
  note(s, "【25:30-27:30】\n★「どんなことが困りますか？」――これが開かれた質問。\n　ここで“お金”が出てこなければ、傷病手当金の話にはたどり着けません。\n★制度を教えるのが仕事ではありません。相談先を示すところまで。\n　加入条件の確認だけはしてから話すこと。");
}

// 21 事例：薬が怖い
{
  const s = S(false);
  eyebrow(s, "つなぐ ― 事例", false);
  title(s, "「薬が怖い」と言われたら");
  card(s, M, 2.15, CW, 1.2);
  s.addText("「薬が怖い！」　「薬をやめたい！」", {
    x: M, y: 2.15, w: CW, h: 1.2, isTextBox: true, margin: 0, align: "center", valign: "middle",
    fontFace: MIN, fontSize: 28, bold: true, color: INK,
  });
  const q = [["どんなところが怖い？", ["家族・友人からの誤った情報", "ネットからの誤った情報", "障害者になりたくない"], GREEN],
             ["どうしてそう思うの？", ["いつまで飲み続けるのだろう", "そもそも医師に減薬の相談をしていない"], PINK]];
  q.forEach((c, i) => {
    const x = M + i * (CW / 2 + 0.15), w = CW / 2 - 0.15;
    card(s, x, 3.6, w, 1.85);
    s.addText(c[0], { x: x + 0.45, y: 3.85, w: w - 0.9, h: 0.4, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 19, bold: true, color: c[2] });
    s.addText(c[1].map((t) => "・" + t).join("\n"), { x: x + 0.45, y: 4.32, w: w - 0.9, h: 1.0, isTextBox: true, margin: 0, fontFace: GO, fontSize: 13, color: INK, lineSpacing: 24 });
  });
  kicker(s, "言語化を手伝って、主治医に相談するようお伝えしています。", 5.75, INK, 21);
  note(s, "【27:30-29:00】\n★返すのは、たった2つの問いだけ。説得しない。\n　出てきた不安を、主治医に相談できる形に整理して渡す。それが仕事です。");
}

// 22 薬は止めない
{
  const s = S(false);
  eyebrow(s, "つなぐ ― 安全のために", false);
  title(s, "薬は、止めないよう指導します");
  const r = [["離脱症状の可能性がある", "自己判断の中断は危険です"], ["越権行為になる", "薬の判断は医師の領域です"], ["自分を守るため", "何かあったとき、責任を負えません"]];
  r.forEach((it, i) => {
    const y = 2.15 + i * 1.0;
    card(s, M, y, CW, 0.85);
    circle(s, String(i + 1), M + 0.3, y + 0.19, 0.48, PINK, PAPER, 14);
    s.addText(it[0], { x: M + 1.0, y, w: 4.6, h: 0.85, isTextBox: true, margin: 0, valign: "middle", fontFace: MIN, fontSize: 17, bold: true, color: INK });
    s.addText(it[1], { x: M + 5.8, y, w: CW - 6.1, h: 0.85, isTextBox: true, margin: 0, valign: "middle", fontFace: GO, fontSize: 14, color: MUTED });
  });
  kicker(s, "減薬の話が出たら、「主治医に相談してみては」が正解です。", 5.5, INK, 22);
  note(s, "【29:00-30:00】\n★ここは必ず入れてください。傾聴を教えると、必ず踏み越える人が出ます。\n聴くことと、介入することは別。境界線をはっきり引いておく。");
}

/* ══════════ D. 距離感 ══════════ */

// 23 こんな患者さん
{
  const s = S(false);
  eyebrow(s, "距離感の話", false);
  title(s, "こんな患者さん、いませんか");
  card(s, M, 2.2, CW, 1.6);
  s.addText("「今まで、話を聴いてもらえなかった」", {
    x: M, y: 2.2, w: CW, h: 1.6, isTextBox: true, margin: 0, align: "center", valign: "middle",
    fontFace: MIN, fontSize: 34, bold: true, color: INK,
  });
  body(s, "・前の院では、すぐ説明が始まった\n・病院では3分で終わった\n・家族にも、職場にも言えない",
    M, 4.15, CW, 1.1, { fontSize: 15, color: MUTED, lineSpacing: 28 });
  note(s, "【30:00-31:00】\n会場に「いますよね」と確認する。ほぼ全員が頷きます。\nここから距離感の話に入ります。");
}

// 24 長く続く患者さんになる
{
  const s = S(false);
  eyebrow(s, "距離感の話", false);
  title(s, "その方は、長く続く患者さんになります");
  const seq = [["聴いてもらえた", GREEN], ["ここなら話せる", GREEN], ["ここしかない", PINK]];
  seq.forEach((t, i) => {
    const x = M + i * (CW / 3);
    const w = CW / 3 - 0.7;
    card(s, x, 2.2, w, 1.1);
    s.addText(t[0], { x, y: 2.2, w, h: 1.1, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: MIN, fontSize: 19, bold: true, color: t[1] });
    if (i < 2) s.addText("→", { x: x + w + 0.05, y: 2.2, w: 0.6, h: 1.1, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: GO, fontSize: 20, color: MUTED });
  });
  body(s, "売上としては、ありがたい話です。けれど――",
    M, 3.65, CW, 0.5, { fontSize: 15, color: MUTED });
  kicker(s, "「ここしかない」は、患者さんにとって危うい状態です。", 4.4, PINK, 25);
  body(s, "治らないわけではありません。良くなっている方も、たくさんいます。\nただ、依存先がひとつしかないまま、というのが問題なのです。",
    M, 5.2, CW, 1.0, { fontSize: 15, lineSpacing: 28 });
  note(s, "【31:00-32:30】\n★ここは誤解されやすいので丁寧に。\n「長く通ってもらうのが悪い」と言っているのではない。\n依存先がひとつきりになることが危ない、という話です。");
}

// 25 私の失敗
{
  const s = S(true);
  eyebrow(s, "距離感の話", true);
  title(s, "私の、失敗です", true);
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 2.1, w: CW, h: 2.2, rectRadius: 0.1, fill: { color: "1B4E36" } });
  body(s, "最初は「とてもいい！」と言ってくださっていました。\nどんどん頼られて、私も応えていました。\n\nそれが、クレームになりました。",
    M + 0.6, 2.5, CW - 1.2, 1.6, { fontSize: 18, color: ONDARK, lineSpacing: 34 });
  s.addText("患者さんのせいではなく、私の距離感が間違っていたのです。", {
    x: M, y: 4.65, w: CW, h: 0.7, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 26, bold: true, color: PINKB,
  });
  note(s, "【32:30-34:30】\n★このセミナーで一番、聴衆の心が動く場所です。\n　具体的なエピソードを1つだけ、短く。相手を悪く言わない。\n「応えれば応えるほど、期待は上がります。上がりきると、応えられない日が必ず来ます」");
}

// 26 距離が近すぎるサイン
{
  const s = S(false);
  eyebrow(s, "距離感の話", false);
  title(s, "距離が近すぎる、3つのサイン");
  const sg = [["施術後、自分が疲れている", "感情を受け取りすぎています"],
              ["患者さんの課題を、自分ごとにしている", "その課題は、誰のものですか？"],
              ["時間外の連絡に、応じている", "枠が消えると、関係は壊れます"]];
  sg.forEach((it, i) => {
    const y = 2.15 + i * 1.02;
    card(s, M, y, CW, 0.88);
    circle(s, "!", M + 0.3, y + 0.19, 0.5, PINK, PAPER, 15);
    s.addText(it[0], { x: M + 1.05, y, w: 6.0, h: 0.88, isTextBox: true, margin: 0, valign: "middle", fontFace: GO, fontSize: 16, bold: true, color: INK });
    s.addText(it[1], { x: M + 7.2, y, w: CW - 7.5, h: 0.88, isTextBox: true, margin: 0, valign: "middle", fontFace: GO, fontSize: 14, color: MUTED });
  });
  kicker(s, "疲れるのは、優しいからではありません。近すぎるからです。", 5.55, INK, 23);
  note(s, "【34:30-36:00】\n★冒頭のあるある④「施術後に疲れてしまう」の回収です。\n必ず「冒頭の4つめ、覚えていますか」と言ってから見せること。");
}

// 27 次の依存先へつなぐ
{
  const s = S(false);
  eyebrow(s, "距離感の話", false);
  title(s, "依存させ切らない。次の依存先へ、つなぐ。");
  const st = [["受け止める", "まず、ちゃんと聴く。ここを飛ばすと繋がりません"],
              ["枠をつくる", "「今日は10分、じっくり伺います」と先に言う"],
              ["渡していく", "自分以外の依存先を、一緒に探す"]];
  st.forEach((it, i) => {
    const x = M + i * (CW / 3), w = CW / 3 - 0.35;
    card(s, x, 2.15, w, 2.3);
    circle(s, String(i + 1), x + w / 2 - 0.32, 2.45, 0.64, i === 2 ? PINK : GREEN, PAPER, 17);
    s.addText(it[0], { x: x + 0.3, y: 3.25, w: w - 0.6, h: 0.42, isTextBox: true, margin: 0, align: "center", fontFace: MIN, fontSize: 19, bold: true, color: INK });
    s.addText(it[1], { x: x + 0.3, y: 3.75, w: w - 0.6, h: 0.6, isTextBox: true, margin: 0, align: "center", fontFace: GO, fontSize: 12.5, color: MUTED, lineSpacing: 21 });
  });
  kicker(s, "適切な距離とは、冷たさではなく、渡す準備のことです。", 4.85, INK, 24);
  note(s, "【36:00-37:30】\n★「枠をつくる」が一番実務的。枠を先に伝えるのも傾聴です。\n無制限に聴くことが傾聴ではない、と明確に言っておく。");
}

// 28 自己理解
{
  const s = S(false);
  eyebrow(s, "距離感の話", false);
  title(s, "なぜ「良くなった」と言われたいのでしょう");
  const two = [["なぜ、良くなったと\n言われたいのか？", GREEN], ["なぜ、アドバイスを\nしたいのか？", PINK]];
  two.forEach((t, i) => {
    const x = M + i * (CW / 2 + 0.15), w = CW / 2 - 0.15;
    card(s, x, 2.15, w, 1.5);
    s.addText(t[0], { x, y: 2.15, w, h: 1.5, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: MIN, fontSize: 21, bold: true, color: t[1], lineSpacing: 34 });
  });
  kicker(s, "この2つは、根っこが似ています。", 4.0, INK, 24);
  body(s, "どちらも「自分が認められたい」から出ていることがあります。\n傾聴を実践していくと、そこに気づけます。腑に落ちると、ぐっと楽になります。",
    M, 4.75, CW, 1.1, { fontSize: 15.5, lineSpacing: 29 });
  note(s, "【37:30-39:00】\n★責めない。「私もそうでした」から入ること。\n自己理解が進むと、患者さんとの距離が自然に適正化します。\nここがサロンで一番時間をかけている部分でもあります。");
}

/* ══════════ E. 傾聴のやり方 ══════════ */

// 症例 1／紹介
{
  const s = S(false);
  eyebrow(s, "症例 ― 鍼灸と傾聴", false);
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
      fontFace: MIN, fontSize: i === 2 ? 22 : 17, bold: true, color: i === 2 ? GREEN : INK,
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
  eyebrow(s, "症例 ― 鍼灸と傾聴", false);
  title(s, "鍼灸で、やっていたこと");
  const tx = [["温補脾陽", "全身倦怠感・めまい・冷え"], ["清胃熱", "胃のむかつき・便秘"], ["疏肝理気", "頭痛・睡眠障害・イライラ"]];
  tx.forEach((t, i) => {
    const y = 2.1 + i * 0.72;
    card(s, M, y, CW * 0.58, 0.6);
    s.addText(t[0], { x: M + 0.35, y, w: 1.6, h: 0.6, isTextBox: true, margin: 0, valign: "middle", fontFace: MIN, fontSize: 16, bold: true, color: GREEN });
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
    x: rx + 0.35, y: 3.65, w: rw - 0.7, h: 0.4, isTextBox: true, margin: 0, fontFace: GO, fontSize: 13, color: PINK,
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
  eyebrow(s, "症例 ― 鍼灸と傾聴", false);
  title(s, "同じように、聴いていたわけではありません");
  const st = [
    ["2–4診", "受容と共感で、まず信頼関係をつくる", "頭痛薬と頓服が、止まった", GREEN],
    ["5–10診", "家族の課題について、情報を整理する", "抑うつを感じにくくなり、出かけるように", PINK],
    ["11–15診", "行動と問題に焦点を当て、できる範囲を明確に", "主治医も親も「回復に驚いている」", GREEN],
    ["16–20診", "どう感じているかを聴き、感情の認知を促す", "「次回も鍼灸、楽しみにしている」", PINK],
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
  eyebrow(s, "症例 ― 鍼灸と傾聴", false);
  title(s, "20診　184日");
  s.addText("PHQ-9（うつ症状）の推移", {
    x: M, y: 2.05, w: CW * 0.6, h: 0.32, isTextBox: true, margin: 0, fontFace: GO, fontSize: 12, bold: true, color: MUTED,
  });
  s.addChart(pres.ChartType.line,
    [{ name: "PHQ-9", labels: ["初診", "5診目", "10診目", "15診目", "20診目"], values: [13, 17, 10, 6, 5] }],
    {
      x: M - 0.1, y: 2.4, w: CW * 0.6, h: 2.6,
      chartColors: [PINK], lineSize: 2.5,
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
    fontFace: MIN, fontSize: 46, bold: true, color: PINK,
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

// 29 8つの技法
{
  const s = S(false);
  eyebrow(s, "傾聴のやり方", false);
  title(s, "傾聴の、8つの技法");
  const t8 = ["相手の気持ちを汲み取る", "繰り返し", "パラフレーズ", "明確化",
              "沈黙", "相手がもっとも問題と感じている所にスポットを当てる", "相手ができる範囲を明確化する", "相手が主体的に行っている事柄に注目する"];
  t8.forEach((t, i) => {
    const x = M + (i % 2) * (CW / 2 + 0.15);
    const y = 2.15 + Math.floor(i / 2) * 0.82;
    const w = CW / 2 - 0.15;
    card(s, x, y, w, 0.7);
    circle(s, String(i + 1), x + 0.28, y + 0.16, 0.38, i % 2 ? PINK : GREEN, PAPER, 12);
    s.addText(t, { x: x + 0.85, y, w: w - 1.1, h: 0.7, isTextBox: true, margin: 0, valign: "middle", fontFace: GO, fontSize: 13.5, color: INK });
  });
  kicker(s, "今日は、この8つを順に見ていきます。", 5.6, INK, 22);
  note(s, "【39:00-40:00】\n先に全体を見せてから、1つずつ展開する（地図を渡してから歩く）。\n★8つ全部を今日で身につける必要はない、と最初に言っておくこと。");
}

// 30-32 技法の解説
const gi = [
  ["① ②　③", "まず、受け取る", [
    ["① 相手の気持ちを汲み取る", "「つらかったですね」――事実ではなく、感情に返す"],
    ["② 繰り返し", "相手の言葉を、相手の言葉のまま返す。言い換えない"],
    ["③ パラフレーズ", "長い話を、短く言い直して確認する。「つまり〇〇ということですか？」"],
  ]],
  ["④　⑤", "深める、待つ", [
    ["④ 明確化", "あいまいなところを聞き返す。その『しんどい』は、どんな感じですか？"],
    ["⑤ 沈黙", "言い終わってから、心のなかで1・2・3。置鍼の待ち時間と同じです"],
  ]],
  ["⑥ ⑦　⑧", "焦点を当て、返していく", [
    ["⑥ もっとも問題と感じている所にスポットを当てる", "全部は扱えません。本人がいちばん困っている一点へ"],
    ["⑦ 相手ができる範囲を明確化する", "「どこまでならできそうですか？」――主語を相手に戻す"],
    ["⑧ 主体的に行っている事柄に注目する", "すでにやれていることを見つけて、言葉にして返す"],
  ]],
];
gi.forEach((g, gi2) => {
  const s = S(false);
  eyebrow(s, "傾聴のやり方", false);
  s.addText(g[0], { x: M, y: 0.6, w: 3.4, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 20, bold: true, color: gi2 === 1 ? PINK : GREEN });
  s.addText(g[1], { x: M + 3.4, y: 0.6, w: CW - 3.4, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 30, bold: true, color: INK });
  g[2].forEach((it, i) => {
    const y = 2.1 + i * 1.3;
    card(s, M, y, CW, 1.1);
    s.addText(it[0], { x: M + 0.5, y: y + 0.18, w: CW - 1.0, h: 0.42, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 18, bold: true, color: INK });
    s.addText(it[1], { x: M + 0.5, y: y + 0.62, w: CW - 1.0, h: 0.4, isTextBox: true, margin: 0, fontFace: GO, fontSize: 14, color: MUTED });
  });
  note(s, ["【40:00-41:30】①〜③は「受け取る」段階。ここができていないと④以降は全部すべります。\n②の繰り返しは、実際に会場で1回やってみせると伝わります。",
    "【41:30-43:00】⑤の沈黙がいちばん難しい。置鍼で待てるのだから待てます、と重ねる。\n④は問い返しなので、慣れないうちは②③のあとに1つだけ。",
    "【43:00-44:30】⑥⑦⑧は、話を相手に返していく段階。\n★⑧が抜けやすい。「もうやれていること」を言葉にして返すだけで、表情が変わります。"][gi2]);
});

// 36b 感情を、追い越さない
{
  const s = S(false);
  eyebrow(s, "傾聴のやり方", false);
  title(s, "感情を、追い越さない");
  const bad = [
    [0, "こういうことが、あったんです"],
    [1, "ああ、それは辛かったですね。大変でしたね"],
  ];
  bad.forEach((t, i) => {
    const mine = t[0] === 1, w = CW * 0.72;
    const x = mine ? M + CW - w : M, y = 2.1 + i * 0.78;
    if (mine) s.addShape(pres.ShapeType.roundRect, { x, y, w, h: 0.66, rectRadius: 0.1, fill: { color: INK } });
    else card(s, x, y, w, 0.66);
    s.addText(t[1], { x: x + 0.35, y, w: w - 0.7, h: 0.66, isTextBox: true, margin: 0, valign: "middle", fontFace: GO, fontSize: 14, color: mine ? ONDARK : INK });
  });
  s.addText("その方の気持ちが「辛い」ではなかったとき、\n患者さんは「分かってもらえない」になります。", {
    x: M, y: 3.8, w: CW, h: 1.0, isTextBox: true, margin: 0,
    fontFace: MIN, fontSize: 22, bold: true, color: PINK, lineSpacing: 36,
  });
  card(s, M, 5.0, CW, 1.5);
  const good = [["まず、繰り返す", "「そういうことが、あったんですね」"], ["そのあと、聞く", "「そのとき、どんなお気持ちでしたか？」"]];
  good.forEach((g, i) => {
    const y = 5.2 + i * 0.6;
    s.addText(g[0], { x: M + 0.5, y, w: 2.6, h: 0.5, isTextBox: true, margin: 0, valign: "middle", fontFace: MIN, fontSize: 15, bold: true, color: GREEN });
    s.addText(g[1], { x: M + 3.2, y, w: CW - 3.7, h: 0.5, isTextBox: true, margin: 0, valign: "middle", fontFace: GO, fontSize: 15, color: INK });
  });
  s.addText("感情は、こちらが決めるものではありません。", {
    x: M, y: 6.7, w: CW, h: 0.55, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 24, bold: true, color: INK,
  });
  note(s, "【41:50-43:10】\n★技法①「相手の気持ちを汲み取る」の、いちばん大きな落とし穴です。\n　共感しているつもりが、決めつけになっている。\n\n★「辛かったですね」と先に言ってしまうと、\n　もしその方の気持ちが 怒り だったり 安堵 だったりしたとき、\n　もう訂正できなくなります。「違うんだけどな」で終わってしまう。\n\n★先回りは、こちらが早く楽になりたいから出ます。そこも正直に言ってよい。\n　「私は、沈黙が怖くて先に言ってしまっていました」\n\n★繰り返し（技法②）が、追い越しを防ぐ一番の安全装置です。ここで②に戻って結びつける。");
}

// 33 自動思考
{
  const s = S(true);
  eyebrow(s, "傾聴のやり方", true);
  s.addText("自動思考", { x: M, y: 2.2, w: CW, h: 1.1, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 48, bold: true, color: ONDARK });
  s.addText("無意識に浮かんでくる思考のこと。", { x: M, y: 3.4, w: CW, h: 0.5, isTextBox: true, margin: 0, fontFace: GO, fontSize: 18, color: MINT });
  const f = [["無意識的で、瞬間的", "自分の意志で考えたわけではなく、反射のように浮かぶ"],
             ["本人は「絶対的な真実」と捉えがち", "疑うことなく、客観的な事実だと信じ込みやすい"],
             ["感情と身体反応を、直接動かす", "不安・落ち込み、緊張・動悸などが引き起こされる"]];
  f.forEach((it, i) => {
    const y = 4.15 + i * 0.62;
    s.addText("・" + it[0], { x: M, y, w: 5.4, h: 0.5, isTextBox: true, margin: 0, fontFace: GO, fontSize: 14, bold: true, color: ONDARK });
    s.addText(it[1], { x: M + 5.6, y, w: CW - 5.6, h: 0.5, isTextBox: true, margin: 0, fontFace: GO, fontSize: 13, color: MUTED_D });
  });
  note(s, "【44:30-45:30】\n★傾聴で「感情」を扱うときの土台になる知識。\nここを知らずに感情に触ると、相手を否定してしまいます。");
}

// 34 出来事→自動思考→感情
{
  const s = S(false);
  eyebrow(s, "傾聴のやり方", false);
  title(s, "自己認知してもらうための、傾聴");
  const chain = [["出来事", "遅刻した", MUTED], ["自動思考", "「なんで起こしてくれないの」\n「自分はダメだ」\n「あの人はできてるのに」", GREEN], ["感情", "焦り・悲しみ\n情けなさ・怒り", PINK]];
  chain.forEach((c, i) => {
    const x = M + i * (CW / 3);
    const w = CW / 3 - 0.7;
    card(s, x, 2.15, w, 2.1);
    s.addText(c[0], { x: x + 0.3, y: 2.4, w: w - 0.6, h: 0.4, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 18, bold: true, color: c[2] });
    s.addText(c[1], { x: x + 0.3, y: 2.9, w: w - 0.6, h: 1.2, isTextBox: true, margin: 0, fontFace: GO, fontSize: 13, color: INK, lineSpacing: 23 });
    if (i < 2) s.addText("→", { x: x + w + 0.05, y: 2.15, w: 0.6, h: 2.1, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: GO, fontSize: 20, color: MUTED });
  });
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 4.55, w: CW, h: 0.95, rectRadius: 0.1, fill: { color: MINT } });
  s.addText("ここに「それはダメだ」と評価を入れてしまうのが、ジャッジメント。", {
    x: M, y: 4.55, w: CW, h: 0.95, isTextBox: true, margin: 0, align: "center", valign: "middle",
    fontFace: MIN, fontSize: 20, bold: true, color: INK,
  });
  note(s, "【45:30-47:00】\n★遅刻の例は誰でも自分ごとにできます。\n「自分はダメだ」という自動思考に、私たちがさらに「そんなこと考えちゃダメ」と\n乗せてしまうのがジャッジ。まずはそれを外す、という話につなげます。");
}

/* ── ワーク：ピンクの象 ── */

// 35 ワーク①予告
{
  const s = S(false);
  eyebrow(s, "ワーク", false);
  title(s, "ワーク①　イラストを覚えてください");
  card(s, M, 2.3, CW, 1.8);
  s.addText("次の画面に出るイラストを、10秒間よく見て、覚えてください。", {
    x: M, y: 2.3, w: CW, h: 1.8, isTextBox: true, margin: 0, align: "center", valign: "middle",
    fontFace: MIN, fontSize: 24, bold: true, color: INK,
  });
  note(s, "【47:00-47:20】\n★このワークがセミナーの山場です。必ずやってください。\n「覚えてください」としか言わないこと。目的は言わない。");
}

// 36 象（10秒）
{
  const s = S(false);
  s.addImage({ path: ASSETS + "/pink-elephant.png", x: 4.9, y: 0.9, w: 3.5, h: 3.5 });
  s.addText("10秒間、よく見てください", {
    x: M, y: 4.7, w: CW, h: 0.6, isTextBox: true, margin: 0, align: "center",
    fontFace: MIN, fontSize: 26, bold: true, color: INK,
  });
  note(s, "【47:20-47:40】\n10秒きっちり見せて、すぐ次へ。長く見せない。");
}

// 37 どっちの足？
{
  const s = S(false);
  eyebrow(s, "ワーク", false);
  title(s, "ピンクの象が挙げていた足は、向かって右？　左？");
  card(s, M, 2.4, CW, 1.9);
  s.addText("右　or　左", {
    x: M, y: 2.4, w: CW, h: 1.9, isTextBox: true, margin: 0, align: "center", valign: "middle",
    fontFace: MIN, fontSize: 44, bold: true, color: GREEN,
  });
  body(s, "答えは、あとで出します。いまは覚えておいてください。",
    M, 4.7, CW, 0.5, { fontSize: 15, color: MUTED, align: "center" });
  note(s, "【47:40-48:20】\n答え：向かって右（象から見て左）。※本番前にイラストでご確認ください。\n★ここでは正解を言わない。「覚えておいてください」で次へ。\n　思い出そうとさせることが、次のワークの仕込みになります。");
}

// 38 ワーク②
{
  const s = S(true);
  eyebrow(s, "ワーク", true);
  title(s, "ワーク②", true);
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 2.2, w: CW, h: 2.0, rectRadius: 0.12, fill: { color: PINKB } });
  s.addText("3分間、目を閉じて\nピンクの象を思い出さないでください。", {
    x: M, y: 2.2, w: CW, h: 2.0, isTextBox: true, margin: 0, align: "center", valign: "middle",
    fontFace: MIN, fontSize: 30, bold: true, color: "FFFFFF", lineSpacing: 48,
  });
  s.addText("では、目を閉じてください。", { x: M, y: 4.6, w: CW, h: 0.5, isTextBox: true, margin: 0, align: "center", fontFace: GO, fontSize: 18, color: MINT });
  note(s, "【48:20-48:40】\n★ゆっくり読む。「思い出さないでください」を強調する。\n次のスライドで3分計ります。時計かスマホのタイマーを使ってください。");
}

// 39 3分
{
  const s = S(true);
  s.addShape(pres.ShapeType.ellipse, { x: 4.87, y: 1.5, w: 3.6, h: 3.6, fill: { color: INK }, line: { color: MINT, width: 3 } });
  s.addText("3", { x: 4.87, y: 1.5, w: 3.6, h: 3.6, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: MIN, fontSize: 110, bold: true, color: ONDARK });
  s.addText("分", { x: 8.5, y: 3.55, w: 0.8, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 22, bold: true, color: MUTED_D });
  s.addText("目を閉じたまま、ピンクの象を思い出さないでください", { x: 1.5, y: 5.5, w: 10.3, h: 0.5, isTextBox: true, margin: 0, align: "center", fontFace: GO, fontSize: 15, color: MUTED_D });
  note(s, "【48:40-51:40】\n★3分は長いです。途中で話しかけないこと。\nPowerPointは自動カウントできないので、時計かスマホで計測。\n30秒残ったところで「あと30秒です」とだけ声をかけてもよい。");
}

// 40 思い浮かべなかった方は
{
  const s = S(false);
  eyebrow(s, "ワーク", false);
  title(s, "3分の間、ピンクの象を思い浮かべなかった方はいますか？");
  card(s, M, 2.4, CW, 1.7);
  s.addText("……たぶん、いませんよね。", {
    x: M, y: 2.4, w: CW, h: 1.7, isTextBox: true, margin: 0, align: "center", valign: "middle",
    fontFace: MIN, fontSize: 32, bold: true, color: INK,
  });
  body(s, "むしろ、思い出さないようにするほど、はっきり浮かんだはずです。",
    M, 4.5, CW, 0.5, { fontSize: 16, color: MUTED, align: "center" });
  note(s, "【51:40-52:10】\n挙手やチャットで確認。ほぼ全員が「浮かんだ」と答えます。\n笑いが起きるところなので、間をとる。");
}

// 41 それが自動思考
{
  const s = S(true);
  s.addText("勝手に浮かんでくるのが、", { x: M, y: 2.5, w: CW, h: 0.8, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 32, bold: true, color: MUTED_D });
  s.addText("自動思考です。", { x: M, y: 3.4, w: CW, h: 1.1, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 48, bold: true, color: ONDARK });
  note(s, "【52:10-52:50】\n★ここで腑に落ちます。理屈で説明するより、体験させたほうが速い。\n患者さんの「変なことを考えてしまう」も、これと同じだと繋げる。");
}

// 45b 気にしないなんて無理
{
  const s = S(false);
  eyebrow(s, "傾聴のやり方", false);
  title(s, "「気にしないで」と、言っていませんか");
  const words = ["「気にしないでくださいね」", "「考えすぎですよ」", "「大丈夫、すぐ良くなりますよ」"];
  words.forEach((t, i) => {
    const x = M + i * (CW / 3), w = CW / 3 - 0.35;
    card(s, x, 2.1, w, 1.15);
    s.addText(t, { x: x + 0.25, y: 2.1, w: w - 0.5, h: 1.15, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: GO, fontSize: 14.5, color: INK, lineSpacing: 24 });
  });
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 3.55, w: CW, h: 0.95, rectRadius: 0.1, fill: { color: MINT } });
  s.addText("これは「ピンクの象を思い出さないでください」と、同じことをお願いしています。", {
    x: M, y: 3.55, w: CW, h: 0.95, isTextBox: true, margin: 0, align: "center", valign: "middle",
    fontFace: MIN, fontSize: 19, bold: true, color: INK,
  });
  s.addText("気にしないなんて、無理なんです。", {
    x: M, y: 4.9, w: CW, h: 0.9, isTextBox: true, margin: 0,
    fontFace: MIN, fontSize: 38, bold: true, color: PINK,
  });
  body(s, "患者さんが気にしてしまうのは、意志が弱いからではありません。\n私たちが「ピンクの象」を思い出してしまったのと、同じ仕組みです。",
    M, 5.95, CW, 1.0, { fontSize: 15, lineSpacing: 28 });
  note(s, "【50:15-51:15】\n★ワークを臨床につなげる、いちばん大事な一枚です。\n　自分たちが3分間できなかったことを、患者さんには毎日お願いしている——\n　この気づきが、傾聴の必要性を一発で腑に落とします。\n\n★3つの言葉は、誰もが言ったことがあるものです。責めない口調で。\n　「私も言っていました」と添えると、素直に受け取ってもらえます。\n\n★では何と言えばいいのか？ ―― の答えが、次のスライドと『8つの技法』です。\n　ここでは答えを出さず、「無理なんです」で止めて次へ。");
}

// 42 ジャッジ
{
  const s = S(false);
  eyebrow(s, "傾聴のやり方", false);
  title(s, "浮かぶのは「当たり前」です");
  card(s, M, 2.2, CW, 1.9);
  s.addText("浮かんだ感情を「ダメ」と否定すること。\nそれが、ジャッジです。", {
    x: M, y: 2.2, w: CW, h: 1.9, isTextBox: true, margin: 0, align: "center", valign: "middle",
    fontFace: MIN, fontSize: 26, bold: true, color: INK, lineSpacing: 44,
  });
  const q = ["なんでダメなの？", "ほんとにダメなの？"];
  q.forEach((t, i) => {
    const x = M + i * (CW / 2 + 0.15), w = CW / 2 - 0.15;
    s.addShape(pres.ShapeType.roundRect, { x, y: 4.45, w, h: 0.95, rectRadius: 0.1, fill: { color: MINT } });
    s.addText(t, { x, y: 4.45, w, h: 0.95, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: MIN, fontSize: 22, bold: true, color: INK });
  });
  kicker(s, "この2つを、自分に向けられるようになると楽になります。", 5.7, INK, 21);
  note(s, "【52:50-53:50】\n★患者さんに向ける前に、まず自分に向ける。\n「疲れた自分はダメだ」と思っている施術者が本当に多いです。");
}

// 41 3つの型
{
  const s = S(false);
  eyebrow(s, "明日から、やってみる", false);
  title(s, "明日からできる、3つの型");
  const t = [["最初の90秒", "遮らない"], ["反復＋一拍", "相手の言葉のまま返す"], ["締めを閉じない", "開いた質問で終わる"]];
  t.forEach((it, i) => {
    const x = M + i * (CW / 3);
    const w = CW / 3 - 0.35;
    card(s, x, 2.15, w, 2.3);
    circle(s, "型" + (i + 1), x + w / 2 - 0.42, 2.45, 0.84, i === 1 ? PINK : GREEN, PAPER, 17);
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
    "90秒です。1日10人でも15分。それだけです。", GREEN],
  ["型 2", "反復して、一拍おく", ["相手の言葉を、相手の言葉のまま返す", "「ずっと、しんどかったんですね」", "要約もリフレーミングもしない"],
    "言い換えた瞬間、それは自分の言葉になります。", PINK],
  ["型 3", "施術後を、閉じない", ["✕　「軽くなりましたか？」", "◯　「いま、どんな感じですか？」", "閉じた質問は、情報を閉じる"],
    "最後のひとことで、次回の情報量が決まります。", GREEN],
];
kata.forEach((k, i) => {
  const s = S(false);
  eyebrow(s, "明日から、やってみる", false);
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

// 45 まとめ
{
  const s = S(false);
  eyebrow(s, "まとめ", false);
  title(s, "今日、持ち帰っていただきたいこと");
  const sum = [["聴くと、情報が集まる", "東洋医学的にも、西洋医学的にも。それが実費診療の強み"],
               ["抱えない。つなぐ。", "自立とは、依存先を増やすこと"],
               ["距離は、冷たさではない", "渡す準備をしておくこと"]];
  sum.forEach((it, i) => {
    const y = 2.15 + i * 1.15;
    card(s, M, y, CW, 1.0);
    circle(s, String(i + 1), M + 0.35, y + 0.22, 0.56, i === 1 ? PINK : GREEN, PAPER, 16);
    s.addText(it[0], { x: M + 1.15, y: y + 0.12, w: CW - 1.5, h: 0.42, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 20, bold: true, color: INK });
    s.addText(it[1], { x: M + 1.15, y: y + 0.56, w: CW - 1.5, h: 0.35, isTextBox: true, margin: 0, fontFace: GO, fontSize: 13.5, color: MUTED });
  });
  kicker(s, "明日は、最初の90秒を遮らないところから。", 5.85, INK, 24);
  note(s, "【56:10-57:10】\n★持ち帰りは3つまで。最後の一行で行動を1つに絞る。");
}

// 46 サロン
{
  const s = S(false);
  eyebrow(s, "ご案内", false);
  title(s, "オンラインサロン ここちめいど");
  card(s, M, 2.1, CW * 0.52, 2.9);
  s.addText("傾聴を学ぶ、\n鍼灸師のコミュニティ", { x: M + 0.5, y: 2.4, w: CW * 0.52 - 1.0, h: 0.9, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 23, bold: true, color: INK, lineSpacing: 36 });
  body(s, "2020年4月から、6年目。\n傾聴の練習と、症例の検討。\n学会発表まで、一緒にやっています。",
    M + 0.5, 3.45, CW * 0.52 - 1.0, 1.2, { fontSize: 14, color: MUTED, lineSpacing: 26 });
  const rx = M + CW * 0.52 + 0.5, rw = CW - CW * 0.52 - 0.5;
  s.addShape(pres.ShapeType.roundRect, { x: rx, y: 2.1, w: rw, h: 1.3, rectRadius: 0.1, fill: { color: INK } });
  s.addText("いまは、募集していません", { x: rx, y: 2.1, w: rw, h: 1.3, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: MIN, fontSize: 22, bold: true, color: ONDARK });
  s.addShape(pres.ShapeType.roundRect, { x: rx, y: 3.6, w: rw, h: 1.4, rectRadius: 0.1, fill: { color: MINT } });
  s.addText("次の募集は\n2027年7月スタート予定", { x: rx, y: 3.6, w: rw, h: 1.4, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: MIN, fontSize: 22, bold: true, color: INK, lineSpacing: 36 });
  kicker(s, "先に知りたい方は、LINEにご登録ください。", 5.4, INK, 23);
  note(s, "【57:10-58:20】\n★「募集していない」を先に言うのが大事です。売り込みに聞こえなくなります。\n★開始時期は予定です。確定していない旨を一言添えてください。\n　年（2027年7月）が違っていたら直してください。");
}

// 47 CTA
{
  const s = S(false);
  title(s, "まずは、LINEにご登録ください", false, 0.7);
  card(s, M, 1.9, CW * 0.48, 3.5);
  s.addText("登録された方へ", { x: M + 0.5, y: 2.2, w: CW * 0.48 - 1.0, h: 0.35, isTextBox: true, margin: 0, fontFace: GO, fontSize: 12, bold: true, color: MUTED });
  const perks = ["本日のスライドをお送りします", "無料セミナーのご案内", "サロン募集開始のお知らせ（先行）"];
  perks.forEach((t, i) => {
    const y = 2.7 + i * 0.75;
    circle(s, "●", M + 0.55, y + 0.1, 0.28, GREEN, PAPER, 8);
    s.addText(t, { x: M + 1.1, y, w: CW * 0.48 - 1.6, h: 0.5, isTextBox: true, margin: 0, valign: "middle", fontFace: GO, fontSize: 14.5, color: INK });
  });
  const qx = M + CW * 0.48 + 0.5, qw = CW - CW * 0.48 - 0.5;
  s.addShape(pres.ShapeType.roundRect, { x: qx, y: 1.9, w: qw, h: 3.5, rectRadius: 0.1, fill: { color: PAPER }, line: { color: GREEN, width: 1.5, dashType: "dash" } });
  s.addText("▶ 公式LINEのQRコードを\nここに大きく貼ってください", {
    x: qx + 0.3, y: 3.1, w: qw - 0.6, h: 0.9, isTextBox: true, margin: 0, align: "center",
    fontFace: GO, fontSize: 13, bold: true, color: MUTED, lineSpacing: 22,
  });
  s.addText("▶ 短縮URLをここに（口頭でも読み上げる）", {
    x: qx + 0.3, y: 4.6, w: qw - 0.6, h: 0.4, isTextBox: true, margin: 0, align: "center", fontFace: GO, fontSize: 11, color: MUTED,
  });
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 5.65, w: CW, h: 1.05, rectRadius: 0.1, fill: { color: MINT } });
  s.addText("近日、無料セミナーを開催します。日程はLINEでお知らせします。", {
    x: M, y: 5.65, w: CW, h: 1.05, isTextBox: true, margin: 0, align: "center", valign: "middle",
    fontFace: MIN, fontSize: 24, bold: true, color: INK,
  });
  note(s, "【58:20-60:00】\n★このスライドを30秒 黙って映す。読み取り待ち。ここを急ぐと激減します。\n★Q&Aの間も、このスライドを出しっぱなしに。\n★オンラインならチャットにもリンクを投下（事前にメモ帳へ用意してコピペ）。\n　ただし録画にチャットは残らないので、画面表示が本体です。\n\n【Q&Aのコツ】\n質問には即答せず、まず「その患者さん、なんて仰ってました？」と聴き返す。\n傾聴の実演が、いちばん強い案内になります。");
}

// 43 社会再適応評価尺度
{
  const s = S(false);
  eyebrow(s, "付録", false);
  title(s, "過去1年に、自分に起きた出来事");
  const A = [["配偶者の死",100],["離婚",73],["夫婦の別居",65],["刑務所などへの拘留",63],["家族の死",63],["怪我や病気",53],["結婚",50],["解雇",47],["配偶者との和解",45],["退職",45],["家族の健康や行動の変化",44],["妊娠",40],["性的な問題",39],["新しい家族を迎える",39],["仕事場の大きな再調整",39],["経済状態の変化",38],["親友の死",37],["違う仕事への異動",36],["配偶者との口論の数の変化",35],["抵当に入れる（借金をする）",31],["担保やローンの損失",30]];
  const B = [["仕事場の責任の変化",29],["子供が家を出る",29],["親戚とのトラブル",29],["素晴らしい成功を収める",28],["配偶者の就職や離職",26],["学校へ入る／学校を辞める",26],["生活環境の大きな変化",25],["生活習慣の改訂",24],["上司とのトラブル",23],["勤務時間や勤務条件の変化",20],["住居の変化",20],["転校",20],["趣味や娯楽の変化",19],["宗教活動の変化",19],["社会活動の変化",18],["ローンを組む",17],["睡眠習慣の変化",16],["家族団欒回数の変化",15],["食習慣の変化",15],["長期休暇",13],["クリスマス",12],["軽微な法律違反",11]];
  s.addText("合計が何点になるか、数えてみてください。（だいたいで構いません）", {
    x: M, y: 1.75, w: CW, h: 0.4, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 17, bold: true, color: INK,
  });
  [A, B].forEach((col, ci) => {
    const x = M + ci * (CW / 2 + 0.2), w = CW / 2 - 0.2;
    card(s, x, 2.3, w, 4.25);
    s.addText(col.map((r) => r[0]).join("\n"), {
      x: x + 0.35, y: 2.5, w: w - 1.45, h: 3.9, isTextBox: true, margin: 0,
      fontFace: GO, fontSize: 9, color: INK, lineSpacing: 13,
    });
    s.addText(col.map((r) => String(r[1])).join("\n"), {
      x: x + w - 1.0, y: 2.5, w: 0.65, h: 3.9, isTextBox: true, margin: 0, align: "right",
      fontFace: GO, fontSize: 9, bold: true, color: GREEN, lineSpacing: 13,
    });
  });
  source(s, "社会再適応評価尺度（ホームズとレイ）", 6.7);
  note(s, "【53:50-55:30】\n★配布資料としても渡せるスライドです。\n会場では「だいたいで構いません」と言って1分ほど数えてもらう。\n自分の点数を知ると、患者さんの背景を想像できるようになります。");
}

// 44 点数の意味
{
  const s = S(false);
  eyebrow(s, "付録", false);
  title(s, "その点数が、意味すること");
  const lv = [["200〜300点", "50%", "の確率で健康上の問題が起きたとの報告", GREEN],
              ["300点以上", "約80%", "の確率で、なんらかの病気に", PINK]];
  lv.forEach((l, i) => {
    const y = 2.2 + i * 1.5;
    card(s, M, y, CW, 1.25);
    s.addText(l[0], { x: M + 0.5, y, w: 3.0, h: 1.25, isTextBox: true, margin: 0, valign: "middle", fontFace: MIN, fontSize: 21, bold: true, color: INK });
    s.addText(l[1], { x: M + 3.6, y, w: 2.0, h: 1.25, isTextBox: true, margin: 0, valign: "middle", fontFace: MIN, fontSize: 34, bold: true, color: l[3] });
    s.addText(l[2], { x: M + 5.8, y, w: CW - 6.1, h: 1.25, isTextBox: true, margin: 0, valign: "middle", fontFace: GO, fontSize: 14, color: MUTED });
  });
  kicker(s, "患者さんは、この点数を抱えて来院されています。", 5.4, INK, 24);
  note(s, "【55:30-56:10】\n★自分の点数と照らすと、患者さんの背景が想像できるようになります。\n「聞かないと分からない」に戻して締める。");
}

/* ══════════ F. クロージング ══════════ */

/* ══════════ 出力（縦方向の自動リフロー付き） ══════════ */
const TIMINGS = {1: "0:00-0:30", 2: "0:30-2:20", 3: "2:20-2:40", 4: "2:40-4:30", 5: "4:30-5:40", 6: "5:40-6:30", 7: "6:30-7:20", 8: "7:20-8:10", 9: "8:10-9:30", 10: "9:30-10:40", 11: "10:40-12:30", 12: "12:30-13:20", 13: "13:20-14:40", 14: "14:40-15:30", 15: "15:30-16:50", 16: "16:50-18:00", 17: "18:00-19:10", 18: "19:10-20:20", 19: "20:20-21:30", 20: "21:30-23:10", 21: "23:10-24:20", 22: "24:20-25:10", 23: "25:10-26:00", 24: "26:00-27:20", 25: "27:20-29:00", 26: "29:00-30:10", 27: "30:10-31:20", 28: "31:20-32:30", 29: "32:30-33:40", 30: "33:40-34:50", 31: "34:50-36:30", 32: "36:30-37:40", 33: "37:40-38:30", 34: "38:30-39:40", 35: "39:40-40:40", 36: "40:40-41:50", 37: "41:50-43:10", 38: "43:10-43:50", 39: "43:50-45:00", 40: "45:00-45:15", 41: "45:15-45:35", 42: "45:35-46:10", 43: "46:10-46:30", 44: "46:30-49:30", 45: "49:30-50:00", 46: "50:00-50:20", 47: "50:20-51:20", 48: "51:20-52:05", 49: "52:05-52:40", 50: "52:40-53:50", 51: "53:50-54:50", 52: "54:50-55:50", 53: "55:50-56:45", 54: "56:45-57:50", 55: "57:50-60:00", 56: "付録・時間があれば", 57: "付録・時間があれば"};

const TOP = 0.34, BOTTOM = 6.92;
for (const rec of decks) {
  const ys = rec.items.map((it) => it.o.y);
  const bs = rec.items.map((it) => it.o.y + it.o.h);
  const minT = Math.min(...ys), maxB = Math.max(...bs);
  let map;
  if (minT <= 0.85) {
    const k = Math.max(1, Math.min((BOTTOM - TOP) / (maxB - TOP), 1.34));
    map = (y, h, keep) => {
      if (keep) { const c = TOP + (y + h / 2 - TOP) * k; return [c - h / 2, h]; }
      return [TOP + (y - TOP) * k, h * k];
    };
  } else {
    const dy = (H - (maxB - minT)) / 2 - minT;
    map = (y, h) => [y + dy, h];
  }
  const s = pres.addSlide();
  s.background = { color: rec.dark ? INK : CREAM };
  for (const it of rec.items) {
    const keep = (it.k === "shape" && it.st === pres.ShapeType.ellipse) || it.k === "image";
    const [ny, nh] = map(it.o.y, it.o.h, keep);
    const o = Object.assign({}, it.o, { y: ny, h: nh });
    if (it.k === "text") s.addText(it.t, o);
    else if (it.k === "image") s.addImage(o);
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

const OUT = process.argv[2] || "keicho-seminar.pptx";
pres.writeFile({ fileName: OUT }).then(() => console.log("slides:", n, "->", OUT));
