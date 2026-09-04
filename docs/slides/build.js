const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";           // 13.333 x 7.5
pres.author = "米倉まな";
pres.title = "経絡に触れ、こころに耳を傾ける";
pres.subject = "APNET 講演 60分 / 鍼灸師のための傾聴トレーニング";

// ── ブランドトークン（docs/HANDOFF.md より） ────────────────
const CREAM = "ECE3D4", PAPER = "FCFAF4", INK = "2B2620";
const SHU = "9C5446", MORI = "5E7060";
const MUTED = "7A6E60", MUTED_D = "A2937F", ONDARK = "ECE3D4";
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
    shadow: { type: "outer", angle: 90, blur: 10, offset: 0.04, color: "8C7F6E", opacity: 0.25 },
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
  s.addShape(pres.ShapeType.ellipse, { x: 9.5, y: 1.1, w: 5.3, h: 5.3, fill: { color: "35302A" } });
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
  s.addText("11", { x: 0.6, y: 1.5, w: 5.2, h: 3.2, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 190, bold: true, color: SHU, align: "center", valign: "middle" });
  s.addText("秒", { x: 5.4, y: 3.3, w: 1.0, h: 0.8, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 30, bold: true, color: ONDARK, valign: "middle" });
  s.addText("医師が、患者の話を遮るまでの時間", { x: 6.8, y: 2.35, w: 5.8, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 22, bold: true, color: ONDARK, lineSpacing: 34 });
  body(s, "話し始めてから遮られるまでの中央値。\n先行研究では平均18秒と報告されている。", 6.8, 3.05, 5.8, 1.0, { color: MUTED_D, fontSize: 14 });
  s.addText("Singh Ospina et al., 2019 ／ Beckman & Frankel, 1984", {
    x: 6.8, y: 4.2, w: 5.8, h: 0.6, isTextBox: true, margin: 0, fontFace: GO, fontSize: 9.5, color: "6B5F51", lineSpacing: 15,
  });
  s.addText("では、私たち鍼灸師は？　――　私たちには 40分 あります。", {
    x: M, y: 5.9, w: CW, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 21, bold: true, color: ONDARK,
  });
  note(s, "【1:20-3:00】\n「これは医師を対象にした研究です」と必ず断る。鍼灸師のデータではない。\n最後の一行で、問題提起から一転して『鍼灸師は聴ける職種』という誇りを立てる。");
}

/* ══════════════════════════════════════════════════════════
   B. 自己紹介
   ══════════════════════════════════════════════════════════ */

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

// 7 最強の聴き手
{
  const s = S(false);
  eyebrow(s, "なぜ、鍼灸師に傾聴か", false);
  title(s, "鍼灸師は、最強の聴き手になれる");
  const c = [["40〜60分", "ふたりきりの時間がある"], ["身体に触れる", "言葉の前に、からだの情報がある"], ["会い続ける", "一度きりではない"], ["白衣でも医師でもない", "ちょうどいい距離"]];
  c.forEach((it, i) => {
    const x = M + (i % 2) * (CW / 2);
    const y = 2.1 + Math.floor(i / 2) * 1.42;
    circle(s, "●", x + 0.05, y + 0.16, 0.34, i % 2 === 0 ? SHU : MORI, PAPER, 9);
    s.addText(it[0], { x: x + 0.6, y: y + 0.02, w: CW / 2 - 0.9, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 22, bold: true, color: INK });
    s.addText(it[1], { x: x + 0.6, y: y + 0.58, w: CW / 2 - 0.9, h: 0.45, isTextBox: true, margin: 0, fontFace: GO, fontSize: 14, color: MUTED });
  });
  s.addText("こんな条件が揃っている職種は、他にありません。", {
    x: M, y: 5.55, w: CW, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 20, bold: true, color: SHU,
  });
  note(s, "【8:00-9:30】\n誇りを立てるブロック。ここで会場の背筋が伸びる。\n『information』は『情報』に直しても可。");
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
  note(s, "【9:30-11:30】自覚のスイッチ。\n★Zoom投票をここで使う（カメラオフ100名の参加感をつくる）。\n投票を締めて結果を画面共有し、「8割の方が③ですね」と読み上げる。\nその数字が次のスライドに直結する。");
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
    x: M, y: 4.85, w: CW, h: 0.6, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 28, bold: true, color: SHU,
  });
  note(s, "【11:30-13:00】\n★ここで責めると問合せは来ない。必ず逃げ道を用意する。\n「習っていないだけ」＝『習えば変わる』への伏線。");
}

// 10 データ（差し替え枠）
{
  const s = S(false);
  eyebrow(s, "なぜ、鍼灸師に傾聴か", false);
  title(s, "身体の主訴の、その後ろにあるもの");
  card(s, M, 2.05, CW * 0.55, 3.1);
  s.addText("▶ 学会発表の図をここに貼付", {
    x: M + 0.3, y: 3.3, w: CW * 0.55 - 0.6, h: 0.6, isTextBox: true, margin: 0,
    align: "center", fontFace: GO, fontSize: 15, bold: true, color: MUTED,
  });
  s.addText("全日本鍼灸学会 第3報・第4報\n「鍼灸院におけるうつと不安症状を有する患者の実態調査」", {
    x: M + 0.3, y: 3.95, w: CW * 0.55 - 0.6, h: 0.7, isTextBox: true, margin: 0,
    align: "center", fontFace: GO, fontSize: 11.5, color: MUTED, lineSpacing: 19,
  });
  body(s, "肩こりで、腰痛で、不眠で来られる。\nでもその方の生活には、\n必ず心の話があります。\n\n私はそれを、感覚ではなく\nデータで確かめたくて、\n調べ続けてきました。",
    M + CW * 0.55 + 0.55, 2.35, CW * 0.45 - 0.55, 2.8, { fontSize: 16, lineSpacing: 30 });
  note(s, "【13:00-15:30】\n★このスライドは必ずご自身の学会発表の図に差し替えてください。\n数字はこちらでは入れていません（推測を載せない方針）。\n狙い：直感ではなくエビデンスで話す人、というポジション取り。\nこれがメンター依頼の『質』を上げます。");
}

// 11 治療の一部
{
  const s = S(false);
  eyebrow(s, "なぜ、鍼灸師に傾聴か", false);
  title(s, "聴くことは、治療の一部です");
  const c = [["治療同盟", "関係の質そのものが、\n転帰に影響する"], ["安心と自律神経", "『わかってもらえた』は\n身体の状態を変える"], ["語りが変える", "話すことで、\n痛みの意味づけが変わる"]];
  c.forEach((it, i) => {
    const x = M + i * (CW / 3);
    circle(s, "◯", x, 2.15, 0.5, i === 1 ? MORI : SHU, PAPER, 15);
    s.addText(it[0], { x, y: 2.85, w: CW / 3 - 0.4, h: 0.45, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 21, bold: true, color: INK });
    s.addText(it[1], { x, y: 3.4, w: CW / 3 - 0.4, h: 1.0, isTextBox: true, margin: 0, fontFace: GO, fontSize: 14, color: MUTED, lineSpacing: 24 });
  });
  s.addText("傾聴は、施術の前段階ではありません。施術そのものです。", {
    x: M, y: 5.2, w: CW, h: 0.5, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 22, bold: true, color: INK,
  });
  note(s, "【15:30-17:00】\n『傾聴＝優しさ』から『傾聴＝治療技術』へ、聴衆の枠組みを移す。");
}

// 12 スルーライン（濃色）
{
  const s = S(true);
  s.addShape(pres.ShapeType.ellipse, { x: 4.7, y: 0.6, w: 6.3, h: 6.3, fill: { color: "342F28" } });
  s.addText("経絡に触れ、\nこころに耳を傾ける。", {
    x: 1.5, y: 2.4, w: 10.3, h: 1.9, isTextBox: true, margin: 0, align: "center",
    fontFace: MIN, fontSize: 40, bold: true, color: ONDARK, lineSpacing: 66,
  });
  s.addText("これこそが、鍼灸のチカラだと思っています。", {
    x: 1.5, y: 4.6, w: 10.3, h: 0.5, isTextBox: true, margin: 0, align: "center",
    fontFace: GO, fontSize: 16, color: MUTED_D,
  });
  note(s, "【17:00-18:00】\nこの講演を貫く一行。ここで一度置いて、最後（53枚目）で同じ言葉に戻る。\n読んだあと一拍おいてから次へ。急がない。");
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
  note(s, "【18:00-19:30】\n言い切る。ここで『え、違うの？』という顔をさせたい。\n特に③は多くの人が『我慢すること』だと思っている。");
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
  note(s, "【19:30-21:30】\nロジャーズをそのまま説明すると眠くなる。必ず鍼灸の言葉に翻訳して話す。\n『証を立てる前に』のくだりが、この講演でいちばん鍼灸師に刺さる言い回し。");
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
  note(s, ["【21:30-23:00】沈黙は『気を待つ』のと同じ、と重ねると鍼灸師には一発で伝わる。",
    "【23:00-24:30】『診立てはあとからでも間に合う』は、責任感の強い人ほど救われる一言。",
    "【24:30-26:00】ここで『訓練』という言葉を初めて出す。後半のサロンの話への伏線。"][i]);
});

// 18 つなぎ
{
  const s = S(true);
  s.addText("ここまでは、知識です。", { x: M, y: 2.5, w: CW, h: 0.8, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 34, bold: true, color: MUTED_D });
  s.addText("知識では、1ミリも変わりません。", { x: M, y: 3.4, w: CW, h: 0.9, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 40, bold: true, color: ONDARK });
  s.addText("では、やってみましょう。", { x: M, y: 4.6, w: CW, h: 0.6, isTextBox: true, margin: 0, fontFace: GO, fontSize: 20, color: SHU });
  note(s, "【26:00-27:00】\nワークへの切り替え。ここで声のトーンを一段上げる。");
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
  note(s, "【27:00-28:00】\n★制約を『わざとそうした』に変える。ここで場の空気が一段変わる。\n『カメラはオフのままで大丈夫』は必ず言う。言うほど参加率が上がる。");
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
  s.addShape(pres.ShapeType.ellipse, { x: 4.87, y: 1.35, w: 3.6, h: 3.6, fill: { color: INK }, line: { color: SHU, width: 3 } });
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
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 4.75, w: CW, h: 0.95, rectRadius: 0.1, fill: { color: MORI } });
  s.addText("▶ チャットへ　「何個 書けましたか？」　数字だけ送ってください", {
    x: M + 0.5, y: 4.75, w: CW - 1.0, h: 0.95, isTextBox: true, margin: 0, valign: "middle",
    fontFace: GO, fontSize: 17, bold: true, color: "FFFFFF",
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
  s.addShape(pres.ShapeType.ellipse, { x: 5.17, y: 1.75, w: 3.0, h: 3.0, fill: { color: INK }, line: { color: k <= 3 ? SHU : "4A4239", width: 2.5 } });
  s.addText(String(k), {
    x: 5.17, y: 1.75, w: 3.0, h: 3.0, isTextBox: true, margin: 0, align: "center", valign: "middle",
    fontFace: MIN, fontSize: 88, bold: true, color: k <= 3 ? SHU : ONDARK,
  });
  s.addText("沈黙のワーク", { x: 1.5, y: 5.35, w: 10.3, h: 0.4, isTextBox: true, margin: 0, align: "center", fontFace: GO, fontSize: 13, color: "6B5F51" });
  if (k === 10) note(s, "★このスライドから10枚が、10秒のカウントダウンです。\n\n【設定方法】\n1. スライド一覧でこの10枚を選択\n2.「画面切り替え」タブ →「自動的に切り替え」に 00:01 と入力\n3.「クリック時」のチェックを外す\n\nこれで自動的に10秒進みます。数字を画面に出すことで、\nアーカイブ視聴者に『配信事故』と誤解されるのを防げます。\n\nこの10秒、絶対に話しかけないこと。");
  else note(s, "無言。（自動切り替え 1秒）");
}

// 35 沈黙のあと
{
  const s = S(false);
  eyebrow(s, "体験ワーク ②", false);
  title(s, "どうでしたか");
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 1.95, w: CW, h: 0.95, rectRadius: 0.1, fill: { color: MORI } });
  s.addText("▶ チャットへ　落ち着かなかった方は「1」を送ってください", {
    x: M + 0.5, y: 1.95, w: CW - 1.0, h: 0.95, isTextBox: true, margin: 0, valign: "middle",
    fontFace: GO, fontSize: 17, bold: true, color: "FFFFFF",
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
  s.addShape(pres.ShapeType.roundRect, { x: 9.6, y: 4.45, w: 2.9, h: 2.35, rectRadius: 0.08, fill: { color: "35302A" }, line: { color: MUTED_D, width: 1, dashType: "dash" } });
  s.addText("▶ ここに\n公式LINEのQRを\n貼ってください", { x: 9.7, y: 4.75, w: 2.7, h: 1.1, isTextBox: true, margin: 0, align: "center", fontFace: GO, fontSize: 11, color: MUTED_D, lineSpacing: 18 });
  s.addText("先に撮っておきたい方はどうぞ", { x: 9.6, y: 6.25, w: 2.9, h: 0.35, isTextBox: true, margin: 0, align: "center", fontFace: GO, fontSize: 10, color: "6B5F51" });
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

// 45 症例
{
  const s = S(false);
  eyebrow(s, "明日からの臨床", false);
  title(s, "実際に、こう変わりました");
  card(s, M, 2.05, CW, 3.1);
  s.addText("▶ ご自身の症例を1つ（個人が特定されない形で）", {
    x: M + 0.6, y: 2.4, w: CW - 1.2, h: 0.45, isTextBox: true, margin: 0, fontFace: GO, fontSize: 13, bold: true, color: MUTED,
  });
  body(s, "語るのは、症状の改善ではなく——\n\n・関係が、どう変わったか\n・通院が、どう続くようになったか\n・その方が、何を話せるようになったか",
    M + 0.6, 2.9, CW - 1.2, 2.1, { fontSize: 16, lineSpacing: 29 });
  s.addText("鍼灸症例データベース構築プロジェクト（株式会社ケアクル／東京有明医療大学／玉川大学／ここちめいど）", {
    x: M, y: 5.4, w: CW, h: 0.5, isTextBox: true, margin: 0, fontFace: GO, fontSize: 11.5, color: MUTED,
  });
  note(s, "【48:00-49:00】\n★ここはご自身の症例に差し替えてください。掲載許可の確認を忘れずに。\n★症状の改善で語ると『すごい人の話』になる。関係と継続で語ると『自分にもできそう』になる。\n時間が押していたら、このスライドは一言で流してよい。");
}

// 46 私の失敗（濃色）
{
  const s = S(true);
  eyebrow(s, "明日からの臨床", true);
  title(s, "ただ、私は一度 失敗しています", true);
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 2.1, w: CW, h: 2.3, rectRadius: 0.1, fill: { color: "35302A" } });
  body(s, "聴きすぎて、自分が消耗した時期があります。\nカウンセラーになろうとして、施術がぶれた時期もありました。\n\n傾聴には、自分を守る技術がセットで必要です。",
    M + 0.6, 2.5, CW - 1.2, 1.6, { fontSize: 18, color: ONDARK, lineSpacing: 34 });
  s.addText("ひとりでやると、壊れます。", {
    x: M, y: 4.75, w: CW, h: 0.7, isTextBox: true, margin: 0, fontFace: MIN, fontSize: 30, bold: true, color: SHU,
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
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 2.7, w: CW, h: 1.75, rectRadius: 0.12, fill: { color: SHU } });
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
  s.addShape(pres.ShapeType.ellipse, { x: 4.7, y: 0.6, w: 6.3, h: 6.3, fill: { color: "342F28" } });
  s.addText("経絡に触れ、\nこころに耳を傾ける。", {
    x: 1.5, y: 2.3, w: 10.3, h: 1.9, isTextBox: true, margin: 0, align: "center",
    fontFace: MIN, fontSize: 40, bold: true, color: ONDARK, lineSpacing: 66,
  });
  s.addText("それが、私たちの持っているチカラです。", {
    x: 1.5, y: 4.5, w: 10.3, h: 0.6, isTextBox: true, margin: 0, align: "center",
    fontFace: MIN, fontSize: 22, bold: true, color: SHU,
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
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 5.75, w: CW, h: 0.95, rectRadius: 0.1, fill: { color: MORI } });
  s.addText("録画でご覧の方も、同じように送ってください。全部読みます。　／　本日のスライドをお送りします。", {
    x: M + 0.4, y: 5.75, w: CW - 0.8, h: 0.95, isTextBox: true, margin: 0, valign: "middle",
    fontFace: GO, fontSize: 15, bold: true, color: "FFFFFF",
  });
  note(s, "【57:40-60:00】\n★このスライドを30秒 黙って映す。読み取り待ち。ここを急ぐと激減します。\n★Q&Aの間も、このスライドを出しっぱなしにする。\n★チャットにも同じリンクを投下（事前にメモ帳に用意してコピペ／共同ホストに依頼）。\n　ただしチャットは録画に残らないので、画面表示が本体です。\n★短縮URLは口頭でも読み上げる（音声だけで聴いている人がいます）。\n\n【Q&Aのコツ】\n質問には即答せず、まず「その患者さん、なんて仰ってました？」と聴き返す。\n傾聴の実演が、いちばん強い営業になります。");
}

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
    if (it.k === "text") s.addText(it.t, o); else s.addShape(it.st, o);
  }
  if (rec.num > 1) s.addText(String(rec.num), {
    x: W - 0.85, y: H - 0.5, w: 0.5, h: 0.3, isTextBox: true, margin: 0,
    align: "right", fontFace: GO, fontSize: 10, color: rec.dark ? "6B5F51" : "B7A991",
  });
  if (rec.notes) s.addNotes(rec.notes);
}

const OUT = process.argv[2] || "APNET-keicho-60min.pptx";
pres.writeFile({ fileName: OUT }).then(() => console.log("slides:", n, "->", OUT));
