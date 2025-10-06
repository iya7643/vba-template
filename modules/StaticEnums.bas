Attribute VB_Name = "StaticEnums"
Option Explicit
Option Private Module

Public Enum WeekDayType
    Sunday_ = 1 ' vbSunday
    Monday_ = 2
    Tuesday_ = 3
    Wednesday_ = 4
    Thursday_ = 5
    Friday_ = 6
    Saturday_ = 7
End Enum

Public Enum ProgressBarType
    ProgressForm_ = 1
    StatusBar_
    Both_
End Enum

' Stream.Type
Public Enum AdType
    Binary_ = 1     ' バイナリ
    Text_ = 2       ' テキスト
End Enum

' Stream.ReadText
Public Enum AdReadText
    All_ = -1       ' 全文読み込み
    Line_ = -2      ' 1行だけ読み込み
End Enum

' Stream.SaveToFile
Public Enum AdSaveToFile
    CreateNotExist_ = 1     ' 新規作成
    CreateOverWrite_ = 2    ' 上書き保存
End Enum

' Connection.State
Public Enum AdConnectionState
    Closed_ = 0         ' 閉じている
    Open_ = 1           ' 開いている
    Connecting_ = 2     ' 接続中
    Executing_ = 4      ' 実行中
    Fetching_ = 8       ' データ取得中
End Enum

' Recordset.CursorType
Public Enum AdCursorType
    ForwardOnly_ = 0    ' 前方のみ（最速・読み取り専用）
    Keyset_ = 1         ' キーセット（他ユーザーの追加削除は見える）
    Dynamic_ = 2        ' ダイナミック（すべての変更が反映）
    Static_ = 3         ' スナップショット（静的コピー）
End Enum

' Recordset.LockType
Public Enum AdLock
    ReadOnly_ = 1           ' 読み取り専用
    Pessimistic_ = 2        ' 悲観的ロック（編集中すぐロック）
    Optimistic_ = 3         ' 楽観的ロック（更新時にロック）
    BatchOptimistic_ = 4    ' バッチ更新用の楽観的ロック
End Enum

' Connection.CursorLocation
Public Enum AdCursorLocation
    UseServer_ = 2      ' サーバー側（既定）
    UseClient_ = 3      ' クライアント側（柔軟）
End Enum

' Connection.Open
Public Enum AdConnectOption
    AsyncConnect_ = 16      ' 非同期接続
End Enum

