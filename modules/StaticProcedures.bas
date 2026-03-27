Attribute VB_Name = "StaticProcedures"
Option Explicit
Option Private Module
Private Const MODULE_NAME = "StaticProcedures"

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



'''''
' すべてのApplicationイベント（SheetCalculate）の監視とApplication.OnTimeのジョブを削除します。
'''''
Public Sub clearWatchItemsAndJobs()
On Error Resume Next
    Call App.AppEventHandler.clearWatches

    If StaticVariables.on_time_tasks Is Nothing Then GoTo Finally
    If StaticVariables.on_time_tasks.Count = 0 Then GoTo Finally
    
    Dim target_proc_name As Variant
    For Each target_proc_name In StaticVariables.on_time_tasks.Keys
        Call App.Utility.cancelOnTime(target_proc_name)
    Next
    
Finally:
    On Error GoTo 0
    Set StaticVariables.on_time_tasks = CreateObject("Scripting.Dictionary")
    Application.Calculate
    Call MsgBox("すべての実行中タイマーを停止しました。", vbInformation + vbOKOnly)
End Sub

'''''
' Workbook_BeforeCloseから実行します。
'''''
Public Function finalizeBeforeClose() As Boolean
    Dim is_cancel As Boolean: is_cancel = False
    If StaticVariables.IS_TEST Then
        If MsgBox( _
            "IS_TESTがTrueになっています。開発テスト用チャンネルのみに通知する状態です。" & vbCrLf & _
            "終了してよろしいですか？", _
            vbYesNo + vbQuestion _
        ) = vbNo Then
            is_cancel = True
            GoTo Finally
        End If
    End If

    ' すべてのスケジュール実行を停止します。
    If Not StaticVariables.on_time_tasks Is Nothing Then
        Dim running_proc_name As Variant
        For Each running_proc_name In StaticVariables.on_time_tasks.Keys
            If IsEmpty(running_proc_name) Then GoTo ContinueForProcName
            Call App.Utility.cancelOnTime(running_proc_name)
        
ContinueForProcName:
        Next
    End If
    
Finally:
    finalizeBeforeClose = is_cancel

End Function
