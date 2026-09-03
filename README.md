# 米倉まな プロフィールサイト

メンタル鍼灸師・米倉まな（よねくら まな）の実績プロフィールページ。
静的HTML 1枚 + 画像のみ。ビルド不要。

**公開URL**: `https://<GitHubユーザー名>.github.io/mana-yonekura-profile/`

## 構成

```
index.html          サイト本体（これ1枚で完結）
images/             写真・ロゴ（11点）
.nojekyll           Jekyll処理を無効化（GitHub Pages用）
CLAUDE.md           Claude Code 向けの作業コンテキスト
docs/HANDOFF.md     引き継ぎ資料（設計思想・編集ガイド）
docs/WORKLOG.md     これまでの作業ログ
docs/CONTENT.md     掲載情報の一次ソース（事実・URL一覧）
docs/DEPLOY.md      GitHub Pages 公開手順
src/profile.dc.html 元データ（参考・編集対象ではない）
```

## ローカルで確認する

```bash
python3 -m http.server 8000
# → http://localhost:8000
```

`index.html` をブラウザで直接開いてもほぼ同じ表示になります
（Googleフォントは要ネットワーク接続）。

## デプロイ

`main` ブランチに push するだけで GitHub Pages が自動更新されます。
反映まで 1〜2分。

Settings → Pages → Source: `Deploy from a branch` / Branch: `main` / Folder: `/ (root)`

## 編集するとき

まず `CLAUDE.md` と `docs/HANDOFF.md` を読んでください。
デザイントークン（色・フォント・余白）とセクション構成が書いてあります。
