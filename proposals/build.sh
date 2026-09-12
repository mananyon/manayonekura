#!/usr/bin/env bash
# アポプラスステーション様ご提案スライド（PowerPoint）を書き出す
#   使い方:  ./build.sh              → Dropbox の所定フォルダへ書き出し
#            ./build.sh /出力/先     → 指定フォルダへ書き出し
set -euo pipefail
cd "$(dirname "$0")"

if ! command -v node >/dev/null 2>&1; then
  echo "❌ Node.js が見つかりません。https://nodejs.org からインストールしてください。" >&2
  exit 1
fi

if [ ! -d node_modules/pptxgenjs ]; then
  echo "📦 pptxgenjs を準備しています..."
  npm install --silent pptxgenjs
fi

node build-pptx.js "$@"
