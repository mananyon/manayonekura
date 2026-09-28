# スライド

| フォルダ / ファイル | 用途 |
|---|---|
| `20261016APNET/` | **APNET講演（2026.10.16）** 全63枚・60分。経絡に触れ、こころに耳を傾ける |
| `keicho-seminar.pptx` | **傾聴セミナー（サロン募集用）** 全47枚・60分。LINE登録／無料セミナーへの導線つき |
| `README-seminar.md` | 上のセミナー資料の手引き |
| `build-seminar.js` | セミナー資料の生成元 |
| `assets/` | スライドで使う画像 |

どちらも配色は「ここちめいど」のロゴ由来（深緑・ミント・ピンク）。
詳細は `20261016APNET/README.md` の「デザイン」節に書いてあります。

## 作り直すとき

```bash
npm install pptxgenjs

# APNET講演
cd 20261016APNET && node build.js APNET-keicho-60min.pptx

# 傾聴セミナー
node build-seminar.js keicho-seminar.pptx
```
