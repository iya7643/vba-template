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

Public THIS_SHEET_NAMES As Object
Private Const THIS_SHEET_NAMES_JSON = "{" & _
    " 'BELL': 'Pr', " & _
    " 'MARKET_INDEX': '指数一覧' " & _
"}"

Public FILE_PATTERNS As Object
Private Const FILE_PATTERNS_JSON = "{" & _
    " 'PHILIP': '^504890047_BELL INVESTMENTS_\\d{8}\\.xlsx$', " & _
    " 'TORA': '^EOD_REPORT_BELLINVEST_Phillip_\\d{8}\\.csv$' " & _
"}"


Private Sub Class_Initialize()
    UPDATE_PROGRESS_BAR_PROC_NAME = "updateProgressBar"
End Sub

Public Sub init(ByVal json_helper As JsonHelper)
    Set WEEKDAY_JP = json_helper.ParseJson(Replace(WEEKDAY_JP_JSON, SL_QUOTE, DBL_QUOTE))
    Set THIS_SHEET_NAMES = json_helper.ParseJson(Replace(THIS_SHEET_NAMES_JSON, SL_QUOTE, DBL_QUOTE))
    Set FILE_PATTERNS = json_helper.ParseJson(Replace(FILE_PATTERNS_JSON, SL_QUOTE, DBL_QUOTE))
End Sub

Private Sub Class_Terminate()
    Set THIS_SHEET_NAMES = Nothing
    Set FILE_PATTERNS = Nothing
End Sub

