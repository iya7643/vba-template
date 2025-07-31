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

Public json_helper As JsonHelper

Public consts As Constants
Public util As Utility
Public progress_bar_controller As ProgressBarController
Public progress_bar_next_run_time As Date
Public progress_bar_active As Boolean


Private Sub Class_Initialize()
    Set json_helper = New JsonHelper
    Set consts = New Constants
    Call consts.init(json_helper)
    Set util = New Utility
    Set progress_bar_controller = New ProgressBarController
End Sub

Private Sub Class_Terminate()
    Set json_helper = Nothing
    Set util = Nothing
    Set consts = Nothing
    Set progress_bar_controller = Nothing
End Sub

