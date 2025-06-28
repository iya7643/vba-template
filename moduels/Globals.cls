VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Globals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''
' プロジェクト全体から使用するグローバル変数
''''''''''''''''''''''''''''''
Option Explicit

Private Const MODULE_NAME = "Globals"

Public consts As Constants
Public util As Utility
Public progress_bar_controller As ProgressBarController
Public progress_bar_next_run_time As Date
Public progress_bar_active As Boolean

Public api_slack As ApiSlack
Public api_holidays As ApiHolidaysJp


Private Sub Class_Initialize()
    Set consts = New Constants
    Set util = New Utility
    Set progress_bar_controller = New ProgressBarController
    Set api_slack = New ApiSlack
    Set api_holidays = New ApiHolidaysJp
End Sub

Private Sub Class_Terminate()
    Set consts = Nothing
    Set util = Nothing
    Set progress_bar_controller = Nothing
    Set api_slack = Nothing
    Set api_holidays = Nothing
End Sub

