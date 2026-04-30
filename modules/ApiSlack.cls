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

Private APP_ID As String
Private TOKEN As String
Private http As Object

Private Sub Class_Initialize()
    If IS_TEST Then
        APP_ID = "xxxxx"
        TOKEN = "xxxxx"
    Else
        APP_ID = "xxxxx"
        TOKEN = "xxxxx"
    End If
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
    post_data = "{""text"":" & App.JsonHelper.ConvertToJson(text) & "}"
    
    Dim url As String
    url = BASE_URL & WORKSPACE_ID & "/" & APP_ID & "/" & TOKEN
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send post_data

    If http.status = 200 Then
    Else
        Err.Raise vbObjectError + 1000, PROC_NAME, "Slackへの通知に失敗しました。" & vbCrLf & http.ResponseText
    End If
    
    GoTo Finally

Except:
    Err.Raise Err.Number, Err.Source & " " & PROC_NAME, Err.Description

Finally:

End Sub


