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
    Dim obj As Object
    Set obj = glob.json_helper.ParseJson("{""a"":123,""b"":[1,2,3,4],""c"":{""d"":456}}")
    
    
    Dim json_str As String
    json_str = glob.json_helper.ConvertToJson(obj)
    Debug.Print json_str
    
    
    Set obj = CreateObject("Scripting.Dictionary")
    
    obj.Add "a", New Collection
    obj("a").Add CreateObject("Scripting.Dictionary")
    obj("a").Item(1).Add "a1", "aaa"
    obj("a").Item(1).Add "a2", 123
    obj("a").Add New Collection
    obj("a").Item(2).Add "aaa"
    obj("a").Item(2).Add "bbb"
    obj("a").Item(2).Add "ccc"
    
    obj.Add "b", CreateObject("Scripting.Dictionary")
    obj("b").Add "b1", 123
    obj("b").Add "b2", "bbbbb"
    obj("b").Add "b3", CreateObject("Scripting.Dictionary")
    obj("b")("b3").Add "c1", 123
    obj("b")("b3").Add "c2", "ccccc"
    
    
    
    Debug.Print glob.json_helper.ConvertToJson(obj)
    

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
        rng.value = "更新中 " & Now
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

