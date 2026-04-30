VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiGoogleChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''
' GoogleChat APIクラス
''''''''''''''''''''''''''''''
Option Explicit

Private Const MODULE_NAME = "ApiGoogleChat"

Private Const BASE_URL As String = "https://chat.googleapis.com/v1/spaces/${SPACE_ID}/messages?key=${KEY}"
Private Const SPACE_ID = "xxxxx"
Private Const KEY = "xxxxx"

Private http As Object

Private Sub Class_Initialize()
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
End Sub

Private Sub Class_Terminate()
    Set http = Nothing
End Sub

'''''
' GoogleChatスペースへメッセージを送信します。
'''''
Public Sub sendMessage(ByVal text As String)
    Const PROC_NAME = MODULE_NAME & ".sendMessage"

On Error GoTo Except
    Dim post_data As String
    post_data = "{""text"":" & App.JsonHelper.ConvertToJson(text) & "}"
    
    Dim url As String
    url = Replace(BASE_URL, "${SPACE_ID}", SPACE_ID)
    url = Replace(url, "${KEY}", KEY)
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send post_data

    If http.status = 200 Then
    Else
        Err.Raise vbObjectError + 1000, PROC_NAME, "GoogleChatへの通知に失敗しました。" & vbCrLf & http.ResponseText
    End If
    
    GoTo Finally

Except:
    Err.Raise Err.Number, Err.Source & " " & PROC_NAME, Err.Description

Finally:

End Sub

