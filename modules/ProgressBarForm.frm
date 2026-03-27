VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBarForm 
   ClientHeight    =   2040
   ClientLeft      =   400
   ClientTop       =   1240
   ClientWidth     =   9330.001
   OleObjectBlob   =   "ProgressBarForm.frx":0000
End
Attribute VB_Name = "ProgressBarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private lbls As Collection

Public Sub initUI(ByVal bar_color As Long)
    Set lbls = New Collection
    Dim candidates As Collection
    Set candidates = New Collection
    
    Dim c As MSForms.Control
    For Each c In Me.Controls
        If TypeName(c) <> "Label" Then GoTo ContinueForControl
        If Left$(c.Name, Len("ProgressLabel")) = "ProgressLabel" Then
            candidates.Add c
        End If
    
ContinueForControl:
    Next
    
    Dim i As Integer, j As Integer, idx_i As Integer, idx_j As Integer
    For i = 1 To candidates.Count
        Dim ins_at As Integer: ins_at = lbls.Count + 1
        idx_i = CInt(Mid$(candidates.Item(i).Name, Len("ProgressLabel") + 1))
        For j = 1 To lbls.Count
            idx_j = CInt(Mid$(lbls.Item(j).Name, Len("ProgressLabel") + 1))
            If idx_i < idx_j Then
                ins_at = j
                Exit For
            End If
        Next
        
        If ins_at > lbls.Count Then
            lbls.Add candidates.Item(i)
        Else
            lbls.Add candidates.Item(i), , ins_at
        End If
    Next
    
    For i = 1 To lbls.Count
        lbls.Item(i).BackColor = bar_color
        lbls.Item(i).Visible = False
    Next
    
    With Application
        Me.Height = 120
        Me.Width = 334
        Me.Left = .Left + (.Width - Me.Width) / 2
        Me.Top = .Top + (.Height - Me.Height) / 2
    End With
End Sub

Public Sub updateStatus(ByVal text As String, ByVal max_len As Integer)
    Dim s As String: s = text
    s = IIf(Len(s) >= max_len, Replace(s, "...", ""), s & "...")
    
    Me.StatusLabel.Caption = s
'    Me.Repaint
End Sub

Public Sub drawProgress(ByVal progress_index As Integer)
    Dim i As Integer
    For i = 1 To lbls.Count
        lbls.Item(i).Visible = (i <= progress_index)
    Next
    Me.Repaint
End Sub

Public Property Get segmentCount() As Integer
    If lbls Is Nothing Then
        segmentCount = 0
    Else
        segmentCount = lbls.Count
    End If
End Property

Public Property Get currentStatus() As String
    currentStatus = Me.StatusLabel.Caption
End Property

Public Property Let currentStatus(ByVal text As String)
    Me.StatusLabel.Caption = text
End Property

Private Sub UserForm_QueryClose(cancel As Integer, close_mode As Integer)
    If Not App.ProgressBarController Is Nothing Then Call App.ProgressBarController.finish
End Sub
