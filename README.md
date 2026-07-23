# NOTE-PON

NOTE-PONは、Windowsデスクトップ版Microsoft PowerPointで実行中のスライドショーから、現在のスライドに対応するスピーカーノートを取得し、別ウィンドウへ大きく表示する小さなWindowsアプリです。

発表中のPowerPoint操作はそのままに、ノート返し用モニターへノート本文だけを表示できます。通信機能、Webサーバー、PowerPointのスライド操作機能はありません。

## 主な機能

- 実行中のPowerPointへ自動接続
- 現在スライドの番号、総スライド数、スピーカーノートを表示
- スライド変更時に対応するノートへ自動切り替え
- 日本語、改行、空行を維持した大文字表示
- ノート本文の強調情報を舞台向けの大文字表示へ反映
  - 太字、斜体、下線、文字色、上付き・下付き
  - 箇条書き、段落のインデント、左・中央・右・両端揃え
  - NOTE-PONの基本フォント、文字サイズ、行間は維持し、PowerPoint側の値は使用しません
  - 黒背景で読みにくい文字色は、色味を残しながら読み取れる明るさへ補正します
- PowerPointが前面でも動作する固定グローバルホットキー
  - `Ctrl + Alt + F10`：ノートを約4行上へスクロール
  - `Ctrl + Alt + F11`：ノートを約4行下へスクロール
  - 音量アップキー：ノートを1行上へスクロール
  - 音量ダウンキー：ノートを1行下へスクロール
  - ミュートキー：現在の表示領域から1行だけ重ねて、次の原稿へスクロール
- 下に原稿が残っている間、ノート領域の下中央に`▼`を表示
- NOTE-PON起動中、音量アップ／ダウン／ミュートキーはノート操作に使用され、Windowsの音量・ミュート操作には使用されません
- NOTE-PONのウィンドウはクリックしてもキーボードフォーカスを取得せず、通常キー、マウスホイール、スクロールバー、タッチ操作をノートスクロールに使用しません
- 文字サイズ変更、常に手前表示
- PowerPoint終了後の待機と、再起動後の自動再接続

## 入力仕様

NOTE-PONのノートをスクロールする入力は、次の5つだけです。

- `Ctrl + Alt + F10`：約4行上へスクロール
- `Ctrl + Alt + F11`：約4行下へスクロール
- 音量アップキー：1行上へスクロール
- 音量ダウンキー：1行下へスクロール
- ミュートキー：現在の表示領域から1行だけ重ねて、次の原稿へスクロール

ミュートキーは1回の押下につき1ページだけ送り、長押しによる連続ページ送りは行いません。下に原稿が残っている間は、言語に依存しない`▼`をノート領域の下中央へ表示し、末尾に到達すると消します。ページ番号はPowerPointのスライド番号と混同しないよう表示しません。

上下矢印キーを含む通常のキーボード入力では、NOTE-PONのノートはスクロールしません。NOTE-PONのウィンドウや操作ボタンをクリックしてもキーボードフォーカスを取得しないため、PowerPointを前面にしたまま使用できます。

マウスホイール、スクロールバー、タッチ操作もNOTE-PONのノートスクロールには使用しません。現在の最小実装では、NOTE-PON上で受けたマウスホイール入力をPowerPointへ転送しないため、マウスポインターがNOTE-PON上にある間はPowerPointもホイールに反応しない場合があります。

## 必要環境

- Windows 11 バージョン24H2（ビルド26100）以降
- .NET 10 Desktop Runtime 10.0.10以降
- Windowsデスクトップ版Microsoft PowerPoint（Microsoft Office）

Microsoft PowerPoint / Microsoft Officeは本リポジトリおよびNOTE-PONには含まれません。利用するPCへ別途正規にインストールされている必要があります。Microsoft Office関連DLL、Microsoft製品本体、ライセンスキーなども同梱しません。

PowerPoint for the webには対応していません。

## 障害耐性

- PowerPoint監視はNOTE-PON本体とは別の子プロセスで実行します。
- PowerPointから2秒間応答がない場合、画面とスクロールを維持したまま監視プロセスだけを再起動します。連続失敗時は2秒、5秒、10秒、30秒の順で再起動間隔を延ばします。
- NOTE-PONの多重起動を防止します。
- 予期しない例外は`%LOCALAPPDATA%\NOTE-PON\note-pon.log`へ記録します。ログは1MBで世代交代します。
- UIで予期しない例外が発生した場合、壊れた状態で継続せずNOTE-PONだけを安全に終了します。
- NOTE-PONはPowerPointを読み取るだけで、スライド送りや編集は行いません。
- ノートの書式取得に失敗しても本文をプレーンテキストで表示し、PowerPointの操作を妨げません。

## インストール

[GitHub Releases](https://github.com/pondashicom/note-pon/releases)から`NOTE-PON-Setup-0.2.3.exe`をダウンロードして実行します。

現在のインストーラーにはコード署名がありません。Windowsから発行元を確認できない旨の警告が表示される場合は、ダウンロード元とReleaseに記載されたSHA-256を確認してから実行してください。

インストーラーに.NETランタイムやMicrosoft Office関連DLLは含まれていません。必要環境は別途インストールしてください。

## ビルド

リポジトリのルートディレクトリで実行します。

```powershell
dotnet build .\NotePon.csproj --configuration Release
```

ビルド結果は次に生成されます。

```text
bin\Release\net10.0-windows\NOTE-PON.exe
```

## インストーラーのビルド

インストーラーの作成には.NET 10 SDKとInno Setup 6が必要です。

```powershell
winget install --id JRSoftware.InnoSetup --exact
powershell.exe -NoProfile -ExecutionPolicy Bypass -File '.\installer\build-installer.ps1'
```

生成先：

```text
publish\installer\NOTE-PON-Setup-0.2.3.exe
```

## 障害耐性テスト

PowerPoint監視プロセスを意図的に停止し、2秒後の自動再起動、画面側プロセスの応答、多重起動防止、終了時の子プロセス回収を確認します。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File '.\scripts\test-fault-tolerance.ps1'
```

## ノート書式テスト

一時的なPowerPoint資料を作成し、太字、斜体、下線、文字色、上付き・下付き、箇条書き、空行、インデント、中央揃えの取得と表示を確認します。既存資料へ接続しないよう、PowerPointを終了してから実行してください。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File '.\scripts\test-formatted-notes.ps1'
```

検証後の画面キャプチャは`%TEMP%\note-pon-formatted-notes.png`へ保存されます。

## 手動テスト

1. 長いスピーカーノートが入ったPowerPointを開き、スライドショーを開始します。
2. NOTE-PONを起動し、現在のスライドに対応するノートが表示されることを確認します。
3. 下に原稿が残っているときだけ、ノート領域の下中央に`▼`が表示されることを確認します。
4. 音量アップ／ダウンキーで、ノートが1行ずつ上下へ滑らかにスクロールすることを確認します。
5. `Ctrl + Alt + F10`／`F11`で、ノートが約4行ずつ上下へスクロールすることを確認します。
6. ミュートキーを短く1回押し、表示中の最終1行を残して次の原稿へ1ページだけ進むことを確認します。
7. 原稿の末尾まで進み、`▼`が消えることを確認します。
8. NOTE-PONをクリックしてから上下矢印キーを押し、NOTE-PONのノートがスクロールしないことを確認します。
9. NOTE-PON上でマウスホイールを回し、NOTE-PONのノートがスクロールしないことを確認します。
10. 太字、斜体、下線、文字色、上付き・下付き、箇条書き、インデント、段落配置が反映されることを確認します。
11. `A−`、`A＋`、「常に手前」が操作できることを確認します。
12. NOTE-PONを終了し、音量アップ／ダウン／ミュートキーがWindowsの音量操作へ戻ることを確認します。

## 実行

1. Microsoft PowerPointでプレゼンテーションを開きます。
2. スライドショーを開始します。
3. NOTE-PONを起動します。
4. NOTE-PONのウィンドウをノート返し用モニターへ移動し、最大化します。

```powershell
Start-Process -FilePath '.\bin\Release\net10.0-windows\NOTE-PON.exe'
```

NOTE-PONはPowerPointのスライド送りや戻しを行いません。現在スライドの正本は常にPowerPointです。

## 技術構成

- C# / .NET 10
- WPF
- PowerPoint COM連携
- Win32 `RegisterHotKey`
- 外部NuGetパッケージなし
- アプリアイコン原本：`assets\note-pon-icon.svg`

## ライセンス

[MIT License](LICENSE)

Microsoft、PowerPoint、Microsoft OfficeおよびWindowsはMicrosoft Corporationの商標または登録商標です。本プロジェクトはMicrosoft Corporationによる公式製品ではありません。
