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

Public COL_LETTERS As Object
' {
'   "LONG": {
'       "SRC": {"SYMBOL": str, "NOTIONAL": str},
'       "DST": {"SYMBOL": str, "NOTIONAL": str},
'   },
'   "SHORT": {
'       "SRC": {"SYMBOL": str, "NOTIONAL": str},
'       "DST": {"SYMBOL": str, "NOTIONAL": str},
'   },
' }

Private Sub Class_Initialize()
    UPDATE_PROGRESS_BAR_PROC_NAME = "updateProgressBar"
    
    Set WEEKDAY_JP = CreateObject("Scripting.Dictionary")
    WEEKDAY_JP.Add vbSunday, "日"
    WEEKDAY_JP.Add vbMonday, "月"
    WEEKDAY_JP.Add vbTuesday, "火"
    WEEKDAY_JP.Add vbWednesday, "水"
    WEEKDAY_JP.Add vbThursday, "木"
    WEEKDAY_JP.Add vbFriday, "金"
    WEEKDAY_JP.Add vbSaturday, "土"

    ' --------------------------------------------------
    ' COL_LETTERS
    ' --------------------------------------------------
    Set COL_LETTERS = CreateObject("Scripting.Dictionary")
    
    COL_LETTERS.Add "LONG", CreateObject("Scripting.Dictionary")
    COL_LETTERS("LONG").Add "SRC", CreateObject("Scripting.Dictionary")
    COL_LETTERS("LONG").Add "DST", CreateObject("Scripting.Dictionary")
    COL_LETTERS.Add "SHORT", CreateObject("Scripting.Dictionary")
    COL_LETTERS("SHORT").Add "SRC", CreateObject("Scripting.Dictionary")
    COL_LETTERS("SHORT").Add "DST", CreateObject("Scripting.Dictionary")
    
    COL_LETTERS("LONG")("SRC").Add "SYMBOL", "U"
    COL_LETTERS("LONG")("SRC").Add "NOTIONAL", "V"
    COL_LETTERS("LONG")("DST").Add "SYMBOL", "Z"
    COL_LETTERS("LONG")("DST").Add "NOTIONAL", "AA"

    COL_LETTERS("SHORT")("SRC").Add "SYMBOL", "U"
    COL_LETTERS("SHORT")("SRC").Add "NOTIONAL", "V"
    COL_LETTERS("SHORT")("DST").Add "SYMBOL", "Z"
    COL_LETTERS("SHORT")("DST").Add "NOTIONAL", "AA"
    ' --------------------------------------------------

End Sub

Private Sub Class_Terminate()
    Set COL_LETTERS = Nothing
End Sub

