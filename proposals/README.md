# アポプラスステーション様 ご提案スライド（導入パート）

5分尺・全10枚。米倉まなの自己紹介と課題提起。

| ファイル | 中身 |
|---|---|
| `apoplus-station-intro.html` | Web版スライド（ブラウザで開く。← → で送り、`N` でスピーカーノート） |
| `build-pptx.js` | PowerPoint版の生成スクリプト |
| `build.sh` | 上記を実行するためのショートカット |
| `米倉まな_アポプラスステーション様ご提案_導入.pptx` | 生成済みのPowerPoint |

## PowerPointを書き出す

```bash
./build.sh
```

Dropbox の `合同会社ここちめいど / 未病 / アポプラスステーション` を自動で探して、
そこに `.pptx` を書き出します。見つからない場合はこのフォルダに出力します。

出力先を指定したいときは:

```bash
./build.sh "/Users/あなた/Desktop"
```

## 中身を直す

文言はすべて `build-pptx.js` の中にそのまま書いてあります。
Claude Code を起動して「3枚目のstand upの説明を◯◯に変えて」と伝えれば、
スクリプトを直してから書き出しまでやります。

```bash
claude
```

Web版（`apoplus-station-intro.html`）とPowerPoint版は別ファイルなので、
文言を変えるときは両方を直す必要があります。Claude Code に頼めば両方まとめて直します。
