VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiHolidaysJp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

''''''''''''''''''''''''''''''
' Holidays JP APIクラス
''''''''''''''''''''''''''''''
Option Explicit

Private Const MODULE_NAME = "ApiHolidaysJp"

Private Const BASE_URL As String = "https://holidays-jp.github.io/api/v1"

Private http As Object
'Private xmlhttp As Object

Private Sub Class_Initialize()
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
'    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
End Sub

Private Sub Class_Terminate()
    Set http = Nothing
'    Set xmlhttp = Nothing
End Sub

'''''
' 祝日一覧を取得します。
'''''
Public Function fetch() As Object
    Const PROC_NAME = MODULE_NAME & ".tokenAuthUser"
    Dim holidays As Object
    Set holidays = CreateObject("Scripting.Dictionary")

On Error GoTo Except
    Dim url As String
    url = BASE_URL & "/date.csv"
    http.Open "GET", url, False
    http.Send
    
    If http.status = 200 Then
        Dim lines() As String, tmp As Variant, i As Long
        lines = Split(http.ResponseText, vbLf)
        For i = 0 To UBound(lines)
            tmp = Split(lines(i), ",")
            If UBound(tmp) <= 0 Then Exit For
            holidays.Add tmp(0), Replace(tmp(1), """", "")
        Next
    Else
        Err.Raise 999, http.status & "エラー", "Holidays JPから祝日情報を取得できませんでした。"
    End If
    
    GoTo Finally

Except:
    Err.Raise Err.Number, Err.Source & " " & PROC_NAME, Err.Description

Finally:
    Set fetch = holidays
    Set holidays = Nothing

End Function
