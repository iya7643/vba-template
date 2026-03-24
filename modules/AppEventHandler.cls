VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''
' AppEventHandlerクラス
' Applicationイベント（SheetCalculate）をフックし、指定した複数セルの再計算を監視するクラス。
''''''''''''''''''''''''''''''

Option Explicit

Private Const MODULE_NAME = "AppEventHandler"

Private WithEvents m_app As Application
Attribute m_app.VB_VarHelpID = -1

Private watch_items As Object   ' { ブックのパス|シート名: { セルアドレス: { "VALUE": 値, "PROC": 実行するプロシージャ名 } } }

'''''
' 初期化します。
'''''
Private Sub init()
    Set m_app = Application
    If watch_items Is Nothing Then Set watch_items = CreateObject("Scripting.Dictionary")
End Sub

'''''
' キーを作成します。
'''''
Private Function buildWatchItemKey(ByVal wb_path As String, ByVal ws_name As String) As String
    buildWatchItemKey = LCase$(wb_path & "|" & ws_name)
End Function

'''''
' 監視対象を追加します。
'''''
Public Sub addWatch( _
    ByVal target_wb_path As String, ByVal ws_name As String, ByVal target_cell_addr As String, ByVal target_proc_name As String _
)
    Call init
    
    Dim wb_path As String, cell_addr As String
    wb_path = LCase$(target_wb_path)
    cell_addr = UCase$(target_cell_addr)
    
    Dim watch_item_key As String
    watch_item_key = buildWatchItemKey(wb_path, ws_name)
    
    If Not watch_items.Exists(watch_item_key) Then
        Set watch_items(watch_item_key) = CreateObject("Scripting.Dictionary")
    End If
    If watch_items(watch_item_key).Exists(cell_addr) Then GoTo Finally
    
    Dim watch_item_value As Variant
    With New CodeBlock
        Dim wb_target As Workbook, wb As Workbook
        For Each wb In m_app.Workbooks
            If LCase$(wb.FullName) = LCase$(wb_path) Then
                Set wb_target = wb
                Exit For
            End If
        Next
    
        If wb_target Is Nothing Then GoTo Finally
        
        watch_item_value = wb_target.Worksheets(ws_name).Range(cell_addr).Value
    End With
    
    Set watch_items(watch_item_key)(cell_addr) = CreateObject("Scripting.Dictionary")
    watch_items(watch_item_key)(cell_addr)("VALUE") = watch_item_value
    watch_items(watch_item_key)(cell_addr)("PROC") = target_proc_name
    
Finally:

End Sub

'''''
' 監視対象を削除します。
'''''
Public Sub removeWatch(ByVal target_wb_path As String, ByVal ws_name As String, Optional ByVal target_cell_addr As String = "")
    Call init
    
    Dim wb_path As String, cell_addr As String
    wb_path = LCase$(target_wb_path)
    cell_addr = UCase$(target_cell_addr)
    
    Dim watch_item_key As String
    watch_item_key = buildWatchItemKey(wb_path, ws_name)
    If Not watch_items.Exists(watch_item_key) Then GoTo Finally
    
    If cell_addr = "" Then
        watch_items.Remove watch_item_key
    Else
        If Not watch_items(watch_item_key).Exists(cell_addr) Then GoTo Finally
        watch_items(watch_item_key).Remove cell_addr
        
        If watch_items(watch_item_key).Count = 0 Then watch_items.Remove watch_item_key
    End If

Finally:

End Sub

'''''
' 監視対象を全削除します。
'''''
Public Sub clearWatches()
    Call init
    watch_items.RemoveAll
End Sub

Private Sub m_app_SheetCalculate(ByVal ws As Object)
    Call init

On Error GoTo Finally
    Dim wb_path As String, ws_name As String, watch_item_key As String
    wb_path = LCase$(ws.Parent.FullName)
    ws_name = LCase$(ws.Name)
    watch_item_key = buildWatchItemKey(wb_path, ws_name)
    If Not watch_items.Exists(watch_item_key) Then GoTo Finally
    
    Dim cell_addr As Variant
    For Each cell_addr In watch_items(watch_item_key).Keys
        Dim prev_value As Variant, curr_value As Variant, target_proc_name As String
        prev_value = watch_items(watch_item_key)(cell_addr)("VALUE")
        curr_value = ws.Range(cell_addr).Value
        If IsError(curr_value) Then GoTo ContinueForCellAddr
        
        If CStr(prev_value) <> CStr(curr_value) Then
            Debug.Print "calculate: " & watch_item_key & "!" & cell_addr & ": " & prev_value & " -> " & curr_value
            target_proc_name = watch_items(watch_item_key)(cell_addr)("PROC")
            m_app.Run target_proc_name
            watch_items(watch_item_key)(cell_addr)("VALUE") = curr_value
        End If
    
ContinueForCellAddr:
    Next

Finally:

End Sub
