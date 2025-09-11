Attribute VB_Name = "MainModule"
Option Explicit
Option Private Module

Private Const MODULE_NAME = "MainModule"

'''''
' メイン処理
'''''
Public Sub run()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Call StaticModule.init

On Error GoTo Except
    ' 必須シートの存在チェックをします。
    Dim sheet_name_key As Variant, sheet_name As String
    For Each sheet_name_key In glob.consts.THIS_SHEET_NAMES
        If glob.util.isExistsSheet(glob.consts.sheet_names(sheet_name_key)) = False Then
            Call MsgBox(glob.consts.sheet_names(sheet_name_key) & "シートがありません。", vbExclamation + vbOKOnly)
            GoTo Finally
        End If
    Next
    
    ' 処理1
'    If MsgBox("1秒ごとに更新する例。よろしいですか？", vbInformation + vbYesNo) <> vbYes Then GoTo Finally
'    Call glob.progress_bar_controller.start("1秒ごとに更新中")
'
'    Call privateProc1

'    Call glob.progress_bar_controller.hide
    
    ' 処理2
'    If MsgBox("ループ処理でその都度更新する例。よろしいですか？", vbInformation + vbYesNo) <> vbYes Then GoTo Finally
'    Call glob.progress_bar_controller.start
'    Call privateProc2
'
'    Call glob.progress_bar_controller.finish
'    Call MsgBox("処理が完了しました。", vbInformation)
    
    
    GoTo Finally

Except:
    Dim err_msg As String
    err_msg = "エラーが発生しました。" & vbCrLf & _
        "Number: " & Err.Number & vbCrLf & _
        "Source: " & Err.Source & vbCrLf & _
        "Description: " & Err.Description
    Call glob.progress_bar_controller.finish
    Call MsgBox(err_msg, vbExclamation)

Finally:
    On Error Resume Next
    Call StaticModule.finish
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
    ' 処理
    Dim end_time As Double
    end_time = Timer + 1
    Do While Timer < end_time
        DoEvents
    Loop
    
    Dim i As Long, rng As Range
    Set rng = ActiveSheet.Range("A1")
    For i = 1 To 3000
        rng.Value = "更新中 " & Now
        rng.Font.Bold = True
        rng.Interior.Color = RGB(200, 230, 200)

        DoEvents
    Next i
    ActiveSheet.UsedRange.Clear
    
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
    ' 処理
    Call glob.progress_bar_controller.setStatus("ループ処理でその都度更新中")
    
    Dim end_time As Double
    end_time = Timer + 2
    Do While Timer < end_time
        Call glob.progress_bar_controller.update
    Loop
    GoTo Finally

Except:
    Err.Raise Err.Number, Err.Source & " " & PROC_NAME, Err.Description

Finally:

End Sub

