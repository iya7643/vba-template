VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBarForm 
   ClientHeight    =   2040
   ClientLeft      =   396
   ClientTop       =   1236
   ClientWidth     =   9336.001
   OleObjectBlob   =   "ProgressBarForm.frx":0000
End
Attribute VB_Name = "ProgressBarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_QueryClose(cancel As Integer, close_mode As Integer)
    If Not glob Is Nothing Then
        If Not glob.progress_bar_controller Is Nothing Then Call glob.progress_bar_controller.finish
    End If
End Sub
