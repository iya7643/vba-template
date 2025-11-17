VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressBarController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''
' プログレスバー/ステータスバーコントローラー
''''''''''''''''''''''''''''''
Option Explicit

Private Const MODULE_NAME = "ProgressBarController"

Private Const max_len As Integer = 30
Private Const WINDOW_TITLE As String = "Please wait..."
Private Const bar_color As Long = 11197857 ' RGB(161, 221, 170)

Private progress_type_ As ProgressBarType
Private frm_ As ProgressBarForm
Private progress_index_ As Integer

Private Sub Class_Terminate()
    Call finish
End Sub

Public Sub start(Optional ByVal status As String = "", Optional ByVal progress_type As ProgressBarType = ProgressBarType.Both_)
    progress_type_ = progress_type
    
    If progress_type_ <> ProgressBarType.ProgressForm_ Then Application.StatusBar = status
    
    If progress_type_ <> ProgressBarType.StatusBar_ Then
        Set frm_ = New ProgressBarForm
        frm_.Caption = WINDOW_TITLE
        Call frm_.initUI(bar_color)
        frm_.currentStatus = status
        progress_index_ = 1
        
        frm_.Show vbModeless
        frm_.Repaint
        
        progress_bar_active = True
        progress_bar_next_run_time = Now + TimeValue("00:00:01")
        Call Application.OnTime(progress_bar_next_run_time, App.Constants.UPDATE_PROGRESS_BAR_PROC_NAME)
    End If
End Sub

Public Sub update()
    Dim s As String
    If progress_type_ = ProgressBarType.ProgressForm_ Then
        s = frm_.currentStatus
    Else
        s = CStr(Application.StatusBar)
    End If
    
    If progress_type_ <> ProgressBarType.ProgressForm_ Then
        s = IIf(Len(s) >= max_len, Replace(s, "...", ""), s & "...")
        Application.StatusBar = s
    End If
    
    If progress_type_ <> ProgressBarType.StatusBar_ Then
        If frm_ Is Nothing Then Exit Sub
        Call frm_.updateStatus(s, max_len)
        
        Dim segs As Integer: segs = frm_.segmentCount
        If segs > 0 Then
            Call frm_.drawProgress(progress_index_)
            progress_index_ = progress_index_ + 1
            If progress_index_ > segs Then progress_index_ = 1
        End If
    End If
    
    DoEvents
End Sub

Public Sub setStatus(ByVal status As String)
    If progress_type_ <> ProgressBarType.ProgressForm_ Then Application.StatusBar = status
    
    If progress_type_ <> ProgressBarType.StatusBar_ Then
        frm_.currentStatus = status
        frm_.Repaint
    End If
End Sub

Public Sub hide()
    If progress_type_ <> ProgressBarType.ProgressForm_ Then Application.StatusBar = False
    
    If progress_type_ <> ProgressBarType.StatusBar_ Then
        On Error Resume Next
        If Not frm_ Is Nothing Then frm_.hide
    End If
End Sub

Public Sub finish()
    If progress_type_ <> ProgressBarType.ProgressForm_ Then Application.StatusBar = False
    
    If progress_type_ <> ProgressBarType.StatusBar_ Then
        On Error Resume Next
        progress_bar_active = False
        If progress_bar_next_run_time > 0 Then
            Application.OnTime EarliestTime:=progress_bar_next_run_time, _
                Procedure:=App.Constants.UPDATE_PROGRESS_BAR_PROC_NAME, _
                Schedule:=False
        End If
        
        If Not frm_ Is Nothing Then
            Unload frm_
            Set frm_ = Nothing
        End If
    End If
End Sub
