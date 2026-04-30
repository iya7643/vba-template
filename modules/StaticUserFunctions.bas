Attribute VB_Name = "StaticUserFunctions"
Option Explicit
Option Private Module
Private Const MODULE_NAME = "StaticUserFunctions"

Private Function getProcName(ByVal target_proc_name As String) As String
    Dim proc_names As Variant
    proc_names = Split(target_proc_name, "!")
    getProcName = CStr(proc_names(UBound(proc_names)))
End Function

'''''
' Applicationイベント（SheetCalculate）の監視状態を返却します。
'''''
Public Function GetWatchItemStatus(ByVal target_proc_name As String, Optional ByVal dummy As Variant) As Variant
    Dim result As Variant
    ReDim result(1 To 2, 1 To 1) As Variant
    result(2, 1) = "値が更新されたとき"
    Dim is_watched As Boolean: is_watched = False
    
    If StaticVariables.watch_items Is Nothing Then GoTo Finally
    
    Dim watch_item_key As Variant, cell_addr As Variant
    For Each watch_item_key In StaticVariables.watch_items.Keys
        For Each cell_addr In StaticVariables.watch_items(watch_item_key)
            Dim proc_name_watched As String
            proc_name_watched = getProcName(StaticVariables.watch_items(watch_item_key)(cell_addr)("PROC"))
            If proc_name_watched <> getProcName(target_proc_name) Then GoTo ContinueForCellAddr
            
            is_watched = True
            GoTo Finally
        
ContinueForCellAddr:
        Next
    Next
    
Finally:
    result(1, 1) = IIf(is_watched, "実行中", "停止中")
    GetWatchItemStatus = result
    
End Function

'''''
' Application.onTimeジョブの状態を返却します。
'''''
Public Function GetRunningJobStatus(ByVal target_proc_name As String, Optional ByVal dummy As Variant) As Variant
    Dim result As Variant
    ReDim result(1 To 2, 1 To 1) As Variant
    
    If StaticVariables.on_time_tasks Is Nothing Then
        result(1, 1) = "停止中"
        result(2, 1) = "-"
        GoTo Finally
    End If
    
    Dim on_time_task_key As Variant, running_proc_name As String, run_at As Variant
    For Each on_time_task_key In StaticVariables.on_time_tasks.Keys
        running_proc_name = getProcName(CStr(on_time_task_key))
        If running_proc_name = getProcName(target_proc_name) Then
            run_at = StaticVariables.on_time_tasks(on_time_task_key)
            Exit For
        End If
    Next
    
    If IsEmpty(run_at) Then
        result(1, 1) = "停止中"
        result(2, 1) = ""
    Else
        result(1, 1) = "実行中"
        result(2, 1) = Format(run_at, "hh:mm:ss")
    End If

Finally:
    GetRunningJobStatus = result

End Function

'''''
' Applicationイベント（SheetCalculate）の監視データとApplication.onTimeジョブのデータを返却します。
'''''
Public Function GetScheduledTasks(Optional ByVal dummy As Variant) As Variant
    Dim watch_item_records As Collection
    Set watch_item_records = getWatchItems
    
    Dim running_job_records As Collection
    Set running_job_records = getRunningJobs
    
    Dim records As New Collection, record As Collection
    For Each record In watch_item_records
        records.Add record
    Next
    For Each record In running_job_records
        records.Add record
    Next
    
    If records.Count = 0 Then
        GetScheduledTasks = ""
    Else
        GetScheduledTasks = App.Utility.nestedCollectionToArray(records)
    End If
End Function

'''''
' Applicationイベント（SheetCalculate）の監視対象データ一覧を取得します。
'''''
Private Function getWatchItems() As Collection
    Dim records As New Collection

    If StaticVariables.watch_items Is Nothing Then GoTo Finally

    Dim watch_item_key As Variant, record As Collection
    For Each watch_item_key In StaticVariables.watch_items.Keys
        Dim watch_item_keys As Variant, wb_name As String, ws_name As String
        watch_item_keys = Split(watch_item_key, "|")
        wb_name = watch_item_keys(0)
        ws_name = watch_item_keys(1)
        
        Dim cell_addr As Variant
        For Each cell_addr In StaticVariables.watch_items(watch_item_key)
            Set record = New Collection
            records.Add getProcAttributes(StaticVariables.watch_items(watch_item_key)(cell_addr)("PROC"), wb_name, ws_name)
        Next
    Next

Finally:
    Set getWatchItems = records

End Function

'''''
' Application.onTimeジョブ一覧を取得します。
'''''
Private Function getRunningJobs() As Collection
    Dim records As New Collection
    
    Dim is_empty_jobs As Boolean: is_empty_jobs = False
    is_empty_jobs = (StaticVariables.on_time_tasks Is Nothing)
    If Not is_empty_jobs Then is_empty_jobs = (StaticVariables.on_time_tasks.Count = 0)
    If is_empty_jobs Then GoTo Finally
    
    Dim target_proc_name_full As Variant
    For Each target_proc_name_full In StaticVariables.on_time_tasks
        records.Add getProcAttributes(target_proc_name_full)
    Next
    
Finally:
    Set getRunningJobs = records

End Function

'''''
' プロシージャ名に対応する情報を取得します。
'''''
Private Function getProcAttributes( _
    ByVal target_proc_name As String, Optional ByVal wb_name As String = "", Optional ByVal ws_name As String = "" _
) As Collection
    Dim result As New Collection
    
    If wb_name = "" And ws_name = "" Then
        result.Add "OnTime"
    Else
        result.Add "SheetCalculate"
    End If
    result.Add getProcName(target_proc_name)
    result.Add wb_name
    result.Add ws_name
    
Finally:
    Set getProcAttributes = result

End Function
