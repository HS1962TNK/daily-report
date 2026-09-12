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

---

## 1. Git と VS Code を入れる

スタートボタンを右クリック →「**ターミナル**」を開き、以下を 1 行ずつ貼り付けて実行します。

```powershell
winget install --id Git.Git -e --source winget
winget install --id Microsoft.VisualStudioCode -e --source winget
```

インストールが終わったら、**ターミナルを一度閉じて開き直します**（PATH を反映させるため）。

確認:

```powershell
git --version
code --version
```

それぞれバージョン番号が出れば成功です。

> `winget` が使えない場合は、手動でダウンロードしてインストールしてください。
> - Git for Windows: <https://git-scm.com/download/win>
> - VS Code: <https://code.visualstudio.com/>

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

## 4. VS Code で開く

```powershell
code .
```

左側にファイル一覧が出ます。ここで HTML を編集します。

おすすめの拡張機能（VS Code の左側「拡張機能」から検索してインストール）:

- **Live Server** … HTML を保存するたびにブラウザへ即反映して確認できる
- **Japanese Language Pack for Visual Studio Code** … VS Code の画面を日本語にする

---

## 5. 動作確認（ブラウザで開く）

### かんたんな方法
エクスプローラーで `index.html` をダブルクリック → Edge / Chrome で開きます。

### おすすめの方法（Live Server）
VS Code で `index.html` を開き、右下または右クリックの「**Open with Live Server**」を選ぶと、
`http://127.0.0.1:5500/index.html` で開きます。保存するたびに自動で再読み込みされます。

> これらのツールはスマホ幅（画面の横幅 480px 想定）で作ってあります。
> PC のブラウザでは画面中央に細長く表示されるのが正常です。
> スマホでの見た目を確認したいときは、ブラウザで `F12` →左上のスマホのアイコン（デバイスツールバー）を押します。

---

## 6. 変更を保存して GitHub に反映する

### VS Code の画面から行う（おすすめ）
1. 左端の「**ソース管理**」（枝アイコン）を開く
2. 変更したファイルの `+` を押す（ステージング）
3. 上のメッセージ欄に「何を直したか」を日本語で書く
4. 「**コミット**」→「**変更の同期**（Sync Changes）」を押す

### コマンドで行う場合

```powershell
cd $env:USERPROFILE\Documents\daily-report

git pull                     # 先に最新を取り込む（重要）
git status                   # 変更されたファイルを確認
git add .                    # 変更をすべて対象にする
git commit -m "作業日報: ○○を修正"
git push
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

## 8. （任意）自宅 PC でも Claude Code を使う

PC 上で Claude に修正を頼みたい場合の準備です。

```powershell
winget install --id OpenJS.NodeJS.LTS -e --source winget
```

ターミナルを開き直してから:

```powershell
npm install -g @anthropic-ai/claude-code
cd $env:USERPROFILE\Documents\daily-report
claude
```

初回はブラウザでのログイン（Claude アカウント）を求められます。

---

## 9. 困ったときは

| 症状 | 対処 |
| --- | --- |
| `git` が見つからない | ターミナルを閉じて開き直す。それでも駄目なら Git を再インストール |
| `push` が「rejected（拒否）」になる | `git pull` を実行してから、もう一度 `git push` |
| 認証画面が繰り返し出る | `git credential-manager github login` を実行し直す |
| 変更していないのに全行が変更扱いになる | `git config --global core.autocrlf false` が設定されているか確認 |
| 間に合わせで変更を元に戻したい | `git restore ファイル名`（コミット前の変更を破棄） |
| どこまで進んだか分からない | `git status` と `git log --oneline -5` で状態を確認 |

---

## 覚えておく 3 つのコマンド

```powershell
git pull      # 始める前：最新を取り込む
git status    # 途中：今どうなっているか見る
git add . ; git commit -m "メモ" ; git push   # 終わり：GitHub に反映する
```
