# 自宅 PC（Windows 11）セットアップ手順

このリポジトリ（`HS1962TNK/daily-report`）を自宅の Windows 11 PC で
編集・保存（コミット）・GitHub へ反映（プッシュ）できるようにする手順です。
上から順にそのまま進めれば設定が終わります。

- リポジトリ: <https://github.com/HS1962TNK/daily-report>
- 中身: 現場で使う HTML ツール一式（作業日報 `index.html`、現場情報 `genba.html`、
  勤怠 `kintai.html`、写真台帳 `photo-album.html`、図面計測 `blueprint-measure.html`）

---

## 0. 用意するもの

- Windows 11 PC（インストール済み）
- GitHub アカウント（このリポジトリの所有者アカウント `HS1962TNK`）
- インターネット接続

### 編集のやり方は 3 通り（どれか 1 つでよい）

| やり方 | 必要なもの | 向いている場面 |
| --- | --- | --- |
| **A. Claude Code アプリに頼む** | Git + Claude Code デスクトップアプリ | いちばん手数が少ない。自然な言葉で修正を頼める |
| **B. Claude（web / スマホ）に頼む** | PC 側は Git だけ（確認用） | 外出先やスマホからも頼める。PC では `git pull` して確認するだけ |
| **C. 自分で編集する** | Git + VS Code などのエディタ | 細かい文言直しを自分でやりたいとき |

**A または B を使うなら VS Code は不要です。** Git だけ入れておけば足ります。

> **アプリの名前の違いに注意**
> - **Claude Code**（デスクトップアプリ / ターミナル）… PC 内のフォルダを開いて、**ファイルの編集・コミット・プッシュまで**できます。これがあれば VS Code の代わりになります。
> - **Claude**（claude.ai のデスクトップアプリ）… 基本はチャットです。PC 内のファイルを直接書き換えたり `git push` したりはしません。この場合は上の **B**（Claude に直してもらい、PC では取り込んで確認する）か **C** になります。

---

## 1. Git を入れる（必須）

スタートボタンを右クリック →「**ターミナル**」を開き、次を貼り付けて実行します。

```powershell
winget install --id Git.Git -e --source winget
```

インストールが終わったら、**ターミナルを一度閉じて開き直します**（PATH を反映させるため）。

確認:

```powershell
git --version
```

バージョン番号が出れば成功です。

> Git は「変更の履歴を記録して GitHub とやり取りする」ための土台で、
> Claude Code もコミットやプッシュのときに裏でこれを使います。ここは省略できません。
> `winget` が使えない場合は <https://git-scm.com/download/win> から手動でインストールしてください。

### （任意）VS Code を入れる

自分で HTML を編集したい場合だけ入れてください。A / B のやり方なら不要です。

```powershell
winget install --id Microsoft.VisualStudioCode -e --source winget
```

おすすめ拡張機能: **Live Server**（保存するたびにブラウザへ即反映）、
**Japanese Language Pack for Visual Studio Code**（画面の日本語化）

---

## 2. Git の初期設定

名前とメールアドレスは、コミットの記録に残る「作業者名」です。GitHub アカウントと同じものを入れてください。

```powershell
git config --global user.name "あなたの名前"
git config --global user.email "あなたのメールアドレス"

# 改行コードを勝手に書き換えさせない（HTML の差分が全行変更になるのを防ぐ）
git config --global core.autocrlf false

# 日本語のファイル名をそのまま表示する
git config --global core.quotepath false

# 既定のブランチ名を main にそろえる
git config --global init.defaultBranch main
```

---

## 3. リポジトリを PC に取り込む（クローン）

ドキュメントフォルダの下に置く例です。

```powershell
cd $env:USERPROFILE\Documents
git clone https://github.com/HS1962TNK/daily-report.git
cd daily-report
```

**初回は GitHub のログイン画面（ブラウザ）が自動で開きます。** GitHub アカウントでサインインして許可してください。
一度許可すれば、次回以降はパスワード入力なしで使えます（Git Credential Manager が記憶します）。

取り込めたか確認:

```powershell
git status
dir
```

`index.html` などのファイルが並んで見えれば成功です。

---

## 4. Claude Code アプリからこのフォルダを開く（やり方 A）

1. Claude Code デスクトップアプリを起動する
2. 作業フォルダとして `C:\Users\<ユーザー名>\Documents\daily-report` を選ぶ
3. 初回は Claude アカウントでのログインを求められるのでサインインする
4. あとは日本語でそのまま頼む。例:
   - 「作業日報の保存ボタンの文字を大きくして」
   - 「現場情報の一覧をあいうえお順に並べて」
   - 「直したらコミットして GitHub にプッシュして」

ターミナルから使いたい場合はこちらでも同じことができます。

```powershell
winget install --id OpenJS.NodeJS.LTS -e --source winget
# ターミナルを閉じて開き直してから
npm install -g @anthropic-ai/claude-code
cd $env:USERPROFILE\Documents\daily-report
claude
```

---

## 5. 動作確認（ブラウザで開く）

### かんたんな方法
エクスプローラーで `index.html` をダブルクリック → Edge / Chrome で開きます。

### VS Code を入れた場合
`index.html` を開き、右クリックの「**Open with Live Server**」を選ぶと
`http://127.0.0.1:5500/index.html` で開き、保存するたびに自動で再読み込みされます。

> これらのツールはスマホ幅（画面の横幅 480px 想定）で作ってあります。
> PC のブラウザでは画面中央に細長く表示されるのが正常です。
> スマホでの見た目を確認したいときは、ブラウザで `F12` →左上のスマホのアイコン（デバイスツールバー）を押します。

---

## 6. 変更を保存して GitHub に反映する

### Claude Code に頼む場合（やり方 A）
「**変更をコミットしてプッシュして**」と伝えれば、そのまま反映まで行います。

### 自分でコマンドを打つ場合

```powershell
cd $env:USERPROFILE\Documents\daily-report

git pull                     # 先に最新を取り込む（重要）
git status                   # 変更されたファイルを確認
git add .                    # 変更をすべて対象にする
git commit -m "作業日報: ○○を修正"
git push
```

### Claude（web / スマホ）に直してもらった場合（やり方 B）
PC 側は取り込むだけです。

```powershell
cd $env:USERPROFILE\Documents\daily-report
git pull
```

> **作業を始める前にはいつも `git pull`** を実行してください。
> スマホ側や他の PC、Claude が加えた変更を先に取り込んでおくと、衝突（コンフリクト）を防げます。

---

## 7. 公開ページ（GitHub Pages）について

このリポジトリが GitHub Pages で公開されている場合、`main` ブランチに反映（プッシュ）すると
1〜2 分後に公開ページへ反映されます。

- 公開 URL（想定）: `https://hs1962tnk.github.io/daily-report/`
- 設定の確認場所: GitHub のリポジトリ →「Settings」→「Pages」

反映されないときは、ブラウザで `Ctrl + F5`（キャッシュを無視して再読み込み）を試してください。

---

## 8. 困ったときは

| 症状 | 対処 |
| --- | --- |
| `git` が見つからない | ターミナルを閉じて開き直す。それでも駄目なら Git を再インストール |
| `push` が「rejected（拒否）」になる | `git pull` を実行してから、もう一度 `git push` |
| 認証画面が繰り返し出る | `git credential-manager github login` を実行し直す |
| 変更していないのに全行が変更扱いになる | `git config --global core.autocrlf false` が設定されているか確認 |
| 変更を元に戻したい | `git restore ファイル名`（コミット前の変更を破棄） |
| どこまで進んだか分からない | `git status` と `git log --oneline -5` で状態を確認 |

---

## 覚えておく 3 つのコマンド

```powershell
git pull      # 始める前：最新を取り込む
git status    # 途中：今どうなっているか見る
git add . ; git commit -m "メモ" ; git push   # 終わり：GitHub に反映する
```
