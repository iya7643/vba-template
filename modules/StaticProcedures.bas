Attribute VB_Name = "StaticProcedures"
Option Explicit
Option Private Module

Public Function App() As AppContext
    If m_app Is Nothing Then Set m_app = New AppContext
    Set App = m_app
End Function

'''''
' 進捗バーユーザーフォームを更新します。（Application.onTimeで実行するため、標準モジュールに定義する必要がある）
'''''
Public Sub updateProgressBar()
    On Error Resume Next
    
    If Not StaticVariables.progress_bar_active Then Exit Sub
    
    Call App.ProgressBarController.update
    StaticVariables.progress_bar_next_run_time = Now + TimeValue("00:00:01")
    Call Application.OnTime(StaticVariables.progress_bar_next_run_time, App.Constants.UPDATE_PROGRESS_BAR_PROC_NAME)
End Sub



Public Property Get DBL_QUOTE() As String
    DBL_QUOTE = Chr(34)    ' 「"」を返却します。
End Property

Public Property Get SL_QUOTE() As String
    SL_QUOTE = "'"        ' 「'」を返却します。
End Property


