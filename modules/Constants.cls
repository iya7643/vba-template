VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Constants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''
' 定数
''''''''''''''''''''''''''''''
Option Explicit

Private Const MODULE_NAME = "Constants"

Public UPDATE_PROGRESS_BAR_PROC_NAME As String

Public WEEKDAY_JP As Object
Private Const WEEKDAY_JP_JSON = "{" & _
    " '" & vbSunday & "': '日', " & _
    " '" & vbMonday & "': '月', " & _
    " '" & vbTuesday & "': '火', " & _
    " '" & vbWednesday & "': '水', " & _
    " '" & vbThursday & "': '木', " & _
    " '" & vbFriday & "': '金', " & _
    " '" & vbSaturday & "': '土' " & _
    "}"

Public COL_LETTERS As Object
Private Const COL_LETTERS_JSON = "{" & _
    " 'LONG': { " & _
        " 'SRC': { 'SYMBOL': 'U', 'NOTIONAL': 'V' }, " & _
        " 'DST': { 'SYMBOL': 'Z', 'NOTIONAL': 'AA' } " & _
    " }, " & _
    " 'SHORT': { " & _
        " 'SRC': { 'SYMBOL': 'U', 'NOTIONAL': 'V' }, " & _
        " 'DST': { 'SYMBOL': 'Z', 'NOTIONAL': 'AA' } " & _
    " } " & _
    "}"

Private Sub Class_Initialize()
    UPDATE_PROGRESS_BAR_PROC_NAME = "updateProgressBar"
End Sub

Public Sub init(ByVal json_helper As JsonHelper)
    Set WEEKDAY_JP = json_helper.ParseJson(Replace(WEEKDAY_JP_JSON, "'", DBL_QUOTE))
    Set COL_LETTERS = json_helper.ParseJson(Replace(COL_LETTERS_JSON, "'", DBL_QUOTE))
End Sub

Private Sub Class_Terminate()
    Set COL_LETTERS = Nothing
End Sub

