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

Private Const max_len = 30
Private Const window_title = "Please wait..."
Private Const bar_color = 11197857 ' RGB(161, 221, 170)

Private progress_type_ As ProgressBarType
Private form As ProgressBarForm
Private lbls As Collection
Private progress_index_ As Integer

Private Sub Class_Initialize()
    Application.StatusBar = False
    
    Set form = New ProgressBarForm
    Set lbls = New Collection
    
    form.Height = 120
    form.Width = 334
    
    With Application
        form.Left = .Left + (.Width - form.Width) / 2
        form.Top = .Top + (.Height - form.Height) / 2
    End With
    
    Dim i As Integer, lbl As MSForms.Label
    For i = 1 To 10
        Set lbl = form.Controls("ProgressLabel" & i)
        lbl.BackColor = bar_color
        lbls.Add lbl
    Next
End Sub

Private Sub Class_Terminate()
    Call finish
End Sub

Public Sub start(Optional ByVal status As String = "", Optional ByVal progress_type As ProgressBarType = ProgressBarType.Both_)
    progress_type_ = progress_type
    
    If progress_type_ <> ProgressBarType.ProgressForm_ Then
        Application.StatusBar = status
    End If
    
    If progress_type_ <> ProgressBarType.StatusBar_ Then
        progress_index_ = 1
        form.Caption = window_title
        form.StatusLabel.Caption = status
        Call form.Show(vbModeless)
        Call form.Repaint
        glob.progress_bar_active = True
        glob.progress_bar_next_run_time = Now + TimeValue("00:00:01")
        Call Application.OnTime(glob.progress_bar_next_run_time, glob.consts.UPDATE_PROGRESS_BAR_PROC_NAME)
    End If
End Sub

Public Sub update()
    Dim status As String
    If progress_type_ = ProgressBarType.ProgressForm_ Then
        status = form.StatusLabel.Caption
    Else
        status = Application.StatusBar
    End If
    
    If Len(status) >= max_len Then
        status = Replace(status, "...", "")
    Else
        status = status & "..."
    End If
    
    If progress_type_ <> ProgressBarType.ProgressForm_ Then
        Application.StatusBar = status
        DoEvents
    End If
    
    If progress_type_ <> ProgressBarType.StatusBar_ Then
        If form Is Nothing Then Exit Sub
        
        form.StatusLabel.Caption = status
        
        Dim i As Integer
        For i = 1 To lbls.Count
            lbls(i).Visible = (i <= progress_index_)
        Next
        
        progress_index_ = progress_index_ + 1
        If progress_index_ > lbls.Count Then
            progress_index_ = 0
        End If
    End If
End Sub

Public Sub setStatus(ByVal status As String)
    If progress_type_ <> ProgressBarType.ProgressForm_ Then
        Application.StatusBar = status
    End If
    
    If progress_type_ <> ProgressBarType.StatusBar_ Then
        form.StatusLabel.Caption = status
    End If
End Sub

Public Sub hide()
    If progress_type_ <> ProgressBarType.ProgressForm_ Then
        Application.StatusBar = False
    End If
    
    If progress_type_ <> ProgressBarType.StatusBar_ Then
        Call form.hide
    End If
End Sub

Public Sub finish()
    If progress_type_ <> ProgressBarType.ProgressForm_ Then
        Application.StatusBar = False
    End If
    
    If progress_type_ <> ProgressBarType.StatusBar_ Then
        On Error Resume Next
        
        glob.progress_bar_active = False
        Call Application.OnTime(glob.progress_bar_next_run_time, glob.consts.UPDATE_PROGRESS_BAR_PROC_NAME, , False)
        Call Unload(form)
        Set form = Nothing
    End If
End Sub
