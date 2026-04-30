# vba-template

VBA で各種マクロを作成するときのベースプロジェクトです。

Excel マクロブックに共通して入れておきたいユーティリティ、アプリケーションコンテキスト、進捗表示、スケジュール実行、シート再計算監視、JSON 変換、外部通知 API などを `modules/` 以下にまとめています。

## 構成

- `template.xlsm`: 新しいマクロを作成するときのベースブック。
- `modules/`: VBE からエクスポートした標準モジュール、クラスモジュール、フォーム。
- `MainModule.bas`: 実装開始用のサンプル処理。
- `AppContext.cls`: `App.Utility` や `App.Constants` など、共通機能へのアクセスポイント。
- `Utility.cls`: Excel 操作、ファイル操作、CSV、ADODB、OnTime、営業日計算などの共通処理。
- `AppEventHandler.cls`: `SheetCalculate` を監視して、対象セルの値が変わったときに指定プロシージャを実行。
- `StaticUserFunctions.bas`: 監視中タスクや OnTime ジョブの状態確認用 UDF。
- `ProgressBarController.cls` / `ProgressBarForm.frm`: 進捗表示とステータスバー更新。
- `JsonHelper.cls`: JSON の parse / stringify。
- `ApiSlack.cls` / `ApiGoogleChat.cls` / `ApiHolidaysJp.cls`: 外部サービス連携用のクラス。

## 使い方

1. `template.xlsm` をコピーして、新しいマクロブックの雛形にします。
2. `MainModule.sample` を起点に、対象案件のメイン処理を実装します。
3. 共通処理は `App.Utility`、定数は `App.Constants`、進捗表示は `App.ProgressBarController` から利用します。

例:

```vb
Call App.ProgressBarController.start("処理中")

If Not App.Utility.isExistsSheet("Sheet1") Then
    Err.Raise vbObjectError, , "Sheet1 がありません。"
End If

Call App.ProgressBarController.finish
```

## スケジュール実行と監視

`Application.OnTime` を使う処理は `App.Utility.scheduleOnTime` / `cancelOnTime` で管理します。登録済みジョブは `GetRunningJobStatus` や `GetScheduledTasks` で確認できます。

シート再計算後に特定セルの値変化を検知したい場合は、`App.AppEventHandler.addWatch` で監視対象を登録します。手続き名は `ProcName` と `Book.xlsm!ProcName` のどちらの形式でも状態確認できるようにしています。

## 外部通知

Slack / Google Chat の送信クラスは、本文を `JsonHelper.ConvertToJson` で JSON エスケープしてから送信します。本文に改行やダブルクォートが含まれていても JSON が壊れない前提です。

利用する場合は、各 API クラス内の URL、ID、トークンなどを実際の値に置き換えてください。

## 開発メモ

- `modules/` の `.bas` / `.cls` / `.frm` は CP932 前提です。
- `.gitattributes` で VBA モジュールの `working-tree-encoding=CP932` を指定しています。
- 文字化けを避けるため、VBA モジュールを編集するときはエンコーディングに注意してください。
