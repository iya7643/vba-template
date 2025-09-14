Attribute VB_Name = "StaticModule"
''''''''''''''''''''''''''''''
' Enumやクラスに定義できないプロシージャを定義する用モジュール
''''''''''''''''''''''''''''''
Option Explicit
Option Private Module

Private Const MODULE_NAME = "StaticModule"

' --------------------------------------------------
' Type
' --------------------------------------------------
' === VBA-UTC Headers
#If Mac Then

#If VBA7 Then
Public Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As LongPtr
End Type

#Else

Public Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As Long
End Type

#End If

#Else

Public Type utc_SYSTEMTIME
    utc_wYear As Integer
    utc_wMonth As Integer
    utc_wDayOfWeek As Integer
    utc_wDay As Integer
    utc_wHour As Integer
    utc_wMinute As Integer
    utc_wSecond As Integer
    utc_wMilliseconds As Integer
End Type

Public Type utc_TIME_ZONE_INFORMATION
    utc_Bias As Long
    utc_StandardName(0 To 31) As Integer
    utc_StandardDate As utc_SYSTEMTIME
    utc_StandardBias As Long
    utc_DaylightName(0 To 31) As Integer
    utc_DaylightDate As utc_SYSTEMTIME
    utc_DaylightBias As Long
End Type

#End If
' === End VBA-UTC

Public Type json_Options
    ' VBA only stores 15 significant digits, so any numbers larger than that are truncated
    ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
    ' See: http://support.microsoft.com/kb/269370
    '
    ' By default, VBA-JSON will use String for numbers longer than 15 characters that contain only digits
    ' to override set `JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True`
    UseDoubleForLargeNumbers As Boolean

    ' The JSON standard requires object keys to be quoted (" or '), use this option to allow unquoted keys
    AllowUnquotedKeys As Boolean

    ' The solidus (/) is not required to be escaped, use this option to escape them as \/ in ConvertToJson
    EscapeSolidus As Boolean
End Type

' --------------------------------------------------
' Enum
' --------------------------------------------------
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


' --------------------------------------------------
' Global変数
' --------------------------------------------------
Public glob As Globals
Public JsonOptions As json_Options

Public Sub init()
    Set glob = New Globals
End Sub

Public Sub finish()
    Set glob = Nothing
End Sub

'''''
' 進捗バーユーザーフォームを更新します。（Application.onTimeで実行するため、標準モジュールに定義する必要がある）
'''''
Public Sub updateProgressBar()
    On Error Resume Next
    
    If glob Is Nothing Then Exit Sub
    If Not glob.progress_bar_active Then Exit Sub
    
    Call glob.progress_bar_controller.update
    glob.progress_bar_next_run_time = Now + TimeValue("00:00:01")
    Call Application.OnTime(glob.progress_bar_next_run_time, glob.consts.UPDATE_PROGRESS_BAR_PROC_NAME)
End Sub

' --------------------------------------------------
' Property
' --------------------------------------------------
Public Property Get DBL_QUOTE() As String
    DBL_QUOTE = Chr(34)    ' 「"」を返却します。
End Property

Public Property Get SL_QUOTE() As String
    SL_QUOTE = "'"        ' 「'」を返却します。
End Property
