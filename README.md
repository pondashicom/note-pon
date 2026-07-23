# NOTE-PON

NOTE-PONは、Windowsデスクトップ版Microsoft PowerPointで実行中のスライドショーから、現在のスライドに対応するスピーカーノートを取得し、別ウィンドウへ大きく表示する小さなWindowsアプリです。

発表中のPowerPoint操作はそのままに、ノート返し用モニターへノート本文だけを表示できます。通信機能、Webサーバー、PowerPointのスライド操作機能はありません。

## 主な機能

- 実行中のPowerPointへ自動接続
- 現在スライドの番号、総スライド数、スピーカーノートを表示
- スライド変更時に対応するノートへ自動切り替え
- 日本語、改行、空行を維持した大文字表示
- マウスホイールとスクロールバーによる縦スクロール
- PowerPointが前面でも動作する固定グローバルホットキー
  - `Ctrl + Alt + F10`：ノートを約4行上へスクロール
  - `Ctrl + Alt + F11`：ノートを約4行下へスクロール
  - 音量アップキー：ノートを1行上へスクロール
  - 音量ダウンキー：ノートを1行下へスクロール
- NOTE-PON起動中、音量アップ／ダウンキーはノート操作に使用され、Windowsの音量変更には使用されません
- 文字サイズ変更、常に手前表示
- PowerPoint終了後の待機と、再起動後の自動再接続

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

## インストール

[GitHub Releases](https://github.com/pondashicom/note-pon/releases)から`NOTE-PON-Setup-0.2.0.exe`をダウンロードして実行します。

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
publish\installer\NOTE-PON-Setup-0.2.0.exe
```

## 障害耐性テスト

PowerPoint監視プロセスを意図的に停止し、2秒後の自動再起動、画面側プロセスの応答、多重起動防止、終了時の子プロセス回収を確認します。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File '.\scripts\test-fault-tolerance.ps1'
```

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
