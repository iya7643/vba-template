VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiSlack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''
' Slack APIクラス
''''''''''''''''''''''''''''''
Option Explicit

Private Const MODULE_NAME = "ApiSlack"

Private Const BASE_URL As String = "https://hooks.slack.com/services/"
Private Const WORKSPACE_ID As String = "xxxxx"
Private Const APP_ID = "xxxxx"
Private Const TOKEN = "xxxxx"

Private http As Object

Private Sub Class_Initialize()
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
End Sub

Private Sub Class_Terminate()
    Set http = Nothing
End Sub

'''''
' Slackチャンネルへメッセージを送信します。
'''''
Public Sub sendMessage(ByVal text As String)
    Const PROC_NAME = MODULE_NAME & ".sendMessage"

On Error GoTo Except
    Dim post_data As String
    post_data = "{""text"": ""<TEXT>""}"
    post_data = Replace(post_data, "<TEXT>", text)
    
    Dim url As String
    url = BASE_URL & WORKSPACE_ID & "/" & APP_ID & "/" & TOKEN
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send post_data

    If http.status = 200 Then
    Else
        Err.Raise 999, "Slackへの通知に失敗しました。¥n" & http.ResponseText
    End If
    
    GoTo Finally

Except:
    Err.Raise Err.Number, Err.Source & " " & PROC_NAME, Err.Description

Finally:

End Sub


