#!/usr/bin/env bash
# 画面モニター 公開スクリプト（Windows は publish.bat をダブルクリック、Mac はターミナルで ./publish.sh）
#
#   ./publish.sh                 … 変更をまとめて公開（説明は自動）
#   ./publish.sh "直した内容"     … 説明つきで公開
#
# 1. GitHub の最新を確認
# 2. 変更したファイルがあれば、版（APP_VERSION / SCRIPT_VERSION）を今の日時に書き換えてコミット
#    ※ すでに公開しているファイルだけが対象（新しいファイルは自動では公開しない）
# 3. GitHub 側の新しい変更（Claude が直接反映した分など）を取り込む
# 4. GitHub に送る → 1分ほどで GitHub Pages に反映。開いている先生用ページも自動で新しい版に切り替わる
#
# 公開用の Git 管理情報を作業フォルダの外（~/.kyotango-deploy.git）に置いている場合は自動でそれを使う。
# 別の場所なら環境変数 KYOTANGO_GIT_DIR で指定する。
set -euo pipefail

cd "$(dirname "$0")"
WT="$(pwd)"
GD="${KYOTANGO_GIT_DIR:-$HOME/.kyotango-deploy.git}"
# 作業フォルダには別の Git（バックアップ用）があることがあるため、公開用の管理情報があればそちらを優先する
if [ -d "$GD" ]; then
  g() { git --git-dir="$GD" --work-tree="$WT" "$@"; }
  echo "公開用の管理情報：$GD"
else
  g() { git "$@"; }
fi
# 送り先が画面モニターのリポジトリであることを確かめる（別のリポジトリに公開しないため）
URL="$(g remote get-url origin 2>/dev/null || true)"
if ! echo "$URL" | grep -qi 'hirokimu1124/kyotango'; then
  echo "✖ 送り先が画面モニターのリポジトリではありません（$URL）。何もしていません。"; exit 1
fi
BRANCH=main
MSG="${1:-}"
stop() { echo; echo "✖ $1"; exit 1; }

echo "▶ GitHub の最新を確認しています…"
g fetch -q origin "$BRANCH" || stop "GitHub につながりません。ネットワークを確認してください。"

# 途中で止まった取り込みが残っていれば知らせる
GITDIR="$(g rev-parse --absolute-git-dir)"
if [ -d "$GITDIR/rebase-merge" ] || [ -d "$GITDIR/rebase-apply" ]; then
  stop "前回の取り込みが途中で止まっています。内容を確認して解決するか、git rebase --abort で元に戻してください。"
fi

CHANGED="$(g diff --name-only HEAD)"
UNTRACKED="$(g ls-files --others --exclude-standard | head -20)"

if [ -n "$CHANGED" ]; then
  VER="$(date '+%Y-%m-%d.%H%M')"
  if echo "$CHANGED" | grep -qx 'teacher.html'; then
    sed -i.bak -E "s/const APP_VERSION = '[^']*';/const APP_VERSION = '$VER';/" teacher.html && rm -f teacher.html.bak
  fi
  if echo "$CHANGED" | grep -qx 'gas/Code.gs'; then
    sed -i.bak -E "s/const SCRIPT_VERSION = '[^']*';/const SCRIPT_VERSION = '$VER';/" gas/Code.gs && rm -f gas/Code.gs.bak
  fi
  echo "▶ 変更したファイル："
  echo "$CHANGED" | sed 's/^/    /'
  g add -u
  g commit -q -m "${MSG:-更新 $VER}"
  echo "  → 版 $VER としてまとめました"
else
  echo "▶ このパソコンでの変更はありません"
fi

if [ -n "$UNTRACKED" ]; then
  echo "（参考）まだ公開していない新しいファイルがあります。公開するときは git add で追加してください："
  echo "$UNTRACKED" | sed 's/^/    /'
fi

BEHIND="$(g rev-list --count "HEAD..origin/$BRANCH")"
if [ "$BEHIND" != "0" ]; then
  echo "▶ GitHub 側の新しい変更を $BEHIND 件取り込みます…"
  if ! g rebase -q "origin/$BRANCH"; then
    g rebase --abort || true
    stop "GitHub 側の変更と、このパソコンの変更が同じ場所でぶつかりました。何も公開していません（このパソコンの変更はコミットとして残っています）。Claude に相談してください。"
  fi
fi

AHEAD="$(g rev-list --count "origin/$BRANCH..HEAD")"
if [ "$AHEAD" = "0" ]; then
  echo
  echo "✔ 最新の状態です（公開するものはありません）"
  exit 0
fi

echo "▶ GitHub に送っています…"
g push -q origin "HEAD:$BRANCH" || stop "GitHub に送れませんでした。少し待ってからもう一度実行してください。"
echo
echo "✔ 公開しました（1分ほどで https://hirokimu1124.github.io/KYOTANGO/ に反映）"
echo "  開いている先生用ページは、つながっている人がいなくなったところで自動で新しい版に切り替わります。"
