# GitHub Pages 公開手順

リポジトリ名: `mana-yonekura-profile`（public）

## A. ブラウザだけで済ませる場合

1. https://github.com/new を開く
2. Repository name に `mana-yonekura-profile`、**Public** を選択
3. README等のチェックは全部外して **Create repository**
4. 次の画面の **uploading an existing file** をクリック
5. ダウンロードしたzipを展開し、**中身**（`index.html` / `images` /
   `docs` / `CLAUDE.md` / `README.md` / `.nojekyll`）をまとめてドラッグ&ドロップ
   - ※ `site` フォルダごとではなく、その中身を入れること
   - ※ `.nojekyll` が見えない場合は Finder で `Cmd + Shift + .`
6. **Commit changes**
7. Settings → Pages → Source: `Deploy from a branch` /
   Branch: `main` / Folder: `/ (root)` → Save
8. 1〜2分で `https://<ユーザー名>.github.io/mana-yonekura-profile/` が公開される

## B. コマンドラインの場合

`gh` CLI が入っていれば一撃。

```bash
cd <展開したsiteフォルダ>

git init
git add -A
git commit -m "米倉まな プロフィールサイト 初回公開"
git branch -M main

gh repo create mana-yonekura-profile --public --source=. --push

# Pages を有効化
gh api -X POST repos/:owner/mana-yonekura-profile/pages \
  -f "source[branch]=main" -f "source[path]=/"
```

`gh` がない場合は手順Aの1〜3でリポジトリだけ作ってから:

```bash
git remote add origin https://github.com/<ユーザー名>/mana-yonekura-profile.git
git push -u origin main
```
→ そのあと Settings → Pages で有効化。

## C. 別端末の Claude Code で続きをやる場合

```bash
git clone https://github.com/<ユーザー名>/mana-yonekura-profile.git
cd mana-yonekura-profile
claude
```

`CLAUDE.md` が自動で読み込まれるので、そのまま編集を頼めば文脈が引き継がれる。
最初に「`docs/HANDOFF.md` と `docs/CONTENT.md` も読んで」と伝えるとより確実。

## 公開後にやると良いこと

- `index.html` の `<link rel="canonical">` を実URLに更新
  （現在はプレースホルダなし、必要なら追加）
- カスタムドメインを当てる場合: リポジトリ直下に `CNAME` ファイルを追加し、
  DNS の CNAME レコードを `<ユーザー名>.github.io` に向ける
