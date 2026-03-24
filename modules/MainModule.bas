Attribute VB_Name = "MainModule"
Option Explicit
Option Private Module

Private Const MODULE_NAME = "MainModule"

'''''
'
'''''
Public Sub sample()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

On Error GoTo Except
    ' 必須シートの存在チェックをします。
    Dim sheet_name As Variant
    For Each sheet_name In App.Constants.THIS_SHEET_NAMES.Items
        If App.Utility.isExistsSheet(sheet_name) = False Then
            Call MsgBox(sheet_name & "シートがありません。", vbExclamation + vbOKOnly)
            GoTo Finally
        End If
    Next
    
    Call App.ProgressBarController.start("処理中")
    
    ' メイン処理
    
    Call App.ProgressBarController.finish
    Call MsgBox("処理が完了しました。", vbInformation + vbOKOnly)
    GoTo Finally

Except:
    Dim err_msg As String
    err_msg = "エラーが発生しました。" & vbCrLf & _
        "Number: " & Err.Number & vbCrLf & _
        "Source: " & Err.Source & vbCrLf & _
        "Description: " & Err.Description
    Call App.ProgressBarController.finish
    Call MsgBox(err_msg, vbExclamation)

Finally:
    On Error Resume Next
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Err.Clear

End Sub

'''''
' サブルーチン例
'''''
Private Sub privateProc1()
    Const PROC_NAME = MODULE_NAME & ".privateProc1"

On Error GoTo Except
    
    GoTo Finally

Except:
    Err.Raise Err.Number, Err.Source & " " & PROC_NAME, Err.Description

Finally:

End Sub

'''''
' サブルーチン例
'''''
Private Sub privateProc2()
    Const PROC_NAME = MODULE_NAME & ".privateProc2"

On Error GoTo Except
    
    GoTo Finally

Except:
    Err.Raise Err.Number, Err.Source & " " & PROC_NAME, Err.Description

Finally:

End Sub

'''''
' サブルーチン例
'''''
Private Sub privateProc3()
    Const PROC_NAME = MODULE_NAME & ".privateProc3"

On Error GoTo Except
    
    GoTo Finally

Except:
    Err.Raise Err.Number, Err.Source & " " & PROC_NAME, Err.Description

Finally:

End Sub

