Attribute VB_Name = "StaticModule"
''''''''''''''''''''''''''''''
' Enumやクラスに定義できないプロシージャを定義する用モジュール
''''''''''''''''''''''''''''''
Option Explicit
Option Private Module

Private Const MODULE_NAME = "StaticModule"

Public glob As Globals

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
' --------------------------------------------------


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
