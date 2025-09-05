VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Utility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''
' ユーティリティクラス
''''''''''''''''''''''''''''''
Option Explicit

Private Const MODULE_NAME = "Utility"

'''''
' 最終行を取得します。
'''''
Public Function lastRow(ByVal ws As Worksheet, Optional ByVal col_value As Variant) As Long
    Dim result As Long: result = 1
    
    If IsMissing(col_value) Then
        result = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        GoTo Finally
    End If
    
    Dim col_num As Long
    If IsNumeric(col_value) Then
        col_num = CLng(col_value)
    Else
        col_num = ws.Range(col_value & "1").Column
    End If
    result = ws.Cells(ws.Rows.Count, col_num).End(xlUp).Row
    
Finally:
    lastRow = result
    
End Function

'''''
' 最終列を取得します。
'''''
Public Function lastCol(ByVal ws As Worksheet, Optional ByVal row_no As Long = 0) As Long
    Dim rng As Range
    If row_no > 0 Then
        Set rng = ws.Rows(row_no)
    Else
        Set rng = ws.UsedRange
    End If
    
    Dim last_cell As Range
    Set last_cell = rng.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If last_cell Is Nothing Then
        lastCol = 0
    Else
        lastCol = last_cell.Column
    End If
End Function

'''''
' 列番号から列文字を取得します。
'''''
Public Function colLetter(ByVal col_no As Integer) As String
    colLetter = Split(Cells(1, col_no).Address(True, False), "$")(0)
End Function

'''''
' シートがあるかチェックします。
'''''
Public Function isExistsSheet(ByVal sheet_name As String) As Boolean
    Const PROC_NAME = MODULE_NAME & ".isExistsSheet"
    Dim ws As Worksheet
    Dim is_exists As Boolean: is_exists = False
    
On Error GoTo Except
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheet_name Then
            is_exists = True
            Exit For
        End If
    Next
    
    GoTo Finally

Except:
    Err.Raise Err.Number, Err.Source & " " & PROC_NAME, Err.Description

Finally:
    isExistsSheet = is_exists

End Function

'''''
' すべてのシートがあるかチェックします。
'''''
Public Function isExistsSheets(ByVal sheet_names As Variant) As Boolean
    Const PROC_NAME = MODULE_NAME & ".isExistsSheets"
    Dim ws As Worksheet
    Dim is_exists As Boolean: is_exists = True
    
On Error GoTo Except
    Dim sheet_name As Variant
    For Each sheet_name In sheet_names
        If Not isExistsSheet(sheet_name) Then
            is_exists = False
            Exit For
        End If
    Next
    
    GoTo Finally

Except:
    Err.Raise Err.Number, Err.Source & " " & PROC_NAME, Err.Description

Finally:
    isExistsSheets = is_exists

End Function

'''''
' フォルダがあるかチェックします。
'''''
Public Function isExistsFolder(ByVal dir_path As String) As Boolean
    isExistsFolder = (Dir(dir_path, vbDirectory) <> "")
End Function

'''''
' ファイルがあるかチェックします。
'''''
Public Function isExistsFile(ByVal file_path As String) As Boolean
    isExistsFile = (Dir(file_path, vbNormal) <> "")
End Function

'''''
' フォルダパスとファイル名の正規表現から、該当するファイルのパスを取得します。
' @return array[エラーメッセージ, ファイルパス]
'''''
Public Function getFileByFilePattern(ByVal dir_path As String, ByVal file_pattern As String) As Variant
    Const PROC_NAME = MODULE_NAME & ".getFileByFilePattern"
    Dim result As Variant

On Error GoTo Except
    Dim target_dir_path As String
    target_dir_path = OneDriveUrlToLocalPath(dir_path)
    target_dir_path = IIf(Right(target_dir_path, 1) = "¥", Left(target_dir_path, Len(target_dir_path) - 1), target_dir_path)
    If Not glob.util.isExistsFolder(target_dir_path) Then
        result = Array("以下のフォルダが存在しません。" & vbCrLf & target_dir_path, "")
        GoTo Finally
    End If

    Dim fso As Object, folder As Object, regex As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(dir_path)
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = file_pattern
    regex.Global = False
    regex.IgnoreCase = True
    
    Dim file As Object
    For Each file In folder.Files
        If regex.Test(file.Name) Then Exit For
    Next
    
    If file Is Nothing Then
        result = Array("ファイル名が以下のパターンのファイルがありません。" & vbCrLf & regex.Pattern, "")
        GoTo Finally
    End If
    
    result = Array("", file.Path)
    
    GoTo Finally
    
Except:
    Err.Raise Err.Number, Err.Source & " " & PROC_NAME, Err.Description
    
Finally:
    getFileByFilePattern = result
    
End Function

'''''
' ファイルが開かれているか判定します。
'''''
Public Function isFileOpen(ByVal file_path As String) As Boolean
    Dim f As Integer
    
    On Error Resume Next
    f = FreeFile
    Open file_path For Binary Access Read Lock Read As #f
    If Err.Number <> 0 Then
        isFileOpen = True
        Err.Clear
    Else
        isFileOpen = False
        Close #f
    End If
    On Error GoTo 0
End Function

'''''
' ファイルを移動します。
'''''
Public Sub moveFile(ByVal src_file_path, ByVal dst_dir_path, Optional is_overwrite = True)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim file_name As Variant, dst_file_path As String
    file_name = Split(src_file_path, "¥")
    file_name = file_name(UBound(file_name))
    
    If Right(dst_dir_path, 1) = "¥" Then
        dst_file_path = dst_dir_path & file_name
    Else
        dst_file_path = dst_dir_path & "¥" & file_name
    End If
    
    If is_overwrite Then
        If fso.FileExists(dst_file_path) Then Call fso.deleteFile(dst_file_path, True)
    End If
    
    Call fso.moveFile(src_file_path, dst_file_path)

End Sub

'''''
' 非表示で別のブックを開きます。
'''''
Public Function openBookInvisible(ByVal file_path As String, Optional ByVal pass As String = "") As Workbook
    Dim prev_sec As Long
    prev_sec = Application.AutomationSecurity

    Application.AutomationSecurity = 3  ' msoAutomationSecurityForceDisable
    Dim wb As Workbook
    
    If isFileOpen(file_path) Then
        Set wb = Workbooks(Dir(file_path))
    Else
        If pass = "" Then
            Set wb = Workbooks.Open(FileName:=file_path, ReadOnly:=True)
        Else
            Set wb = Workbooks.Open(FileName:=file_path, ReadOnly:=True, Password:=pass)
        End If
    End If
    wb.Windows(1).Visible = False
    
    Set openBookInvisible = wb
    Application.AutomationSecurity = prev_sec
End Function

'''''
' ブックを閉じます。
'''''
Public Sub closeBook(ByVal wb As Workbook)
    Application.DisplayAlerts = False
    wb.Close SaveChanges:=False
    Set wb = Nothing
    Application.DisplayAlerts = True
End Sub

'''''
' Dictionaryの値をキーの配列から取得します。
'''''
Public Function dictor(ByVal dict As Object, ByVal dict_keys As Variant)
    Const PROC_NAME = MODULE_NAME & ".dictor"
    Dim i As Integer, key As Variant, tmp As Object
    
On Error GoTo Except
    Set tmp = dict
    For i = 0 To UBound(dict_keys) - 1
        key = dict_keys(i)
        Set tmp = tmp.Item(key)
    Next

    key = dict_keys(UBound(dict_keys))
    If VarType(tmp.Item(key)) = vbObject Then
        Set dictor = tmp.Item(key)
    Else
        dictor = tmp.Item(key)
    End If
    
    GoTo Finally

Except:
    Err.Raise Err.Number, Err.Source & " " & PROC_NAME, Err.Description

Finally:
    Set tmp = Nothing

End Function

'''''
' 値が数値かチェックします。
'''''
Public Function isNumericValue(ByVal tmp As Variant) As Boolean
    isNumericValue = IsNumeric(tmp) And Not IsEmpty(tmp)
End Function

'''''
' 全角・半角スペースをトリムします。
'''''
Public Function fullTrim(ByVal txt As String) As String
    fullTrim = Trim(Replace(txt, "　", ""))
End Function

'''''
' 文字列が含まれているか判定します。
'''''
Public Function isMatch(ByVal txt As String, ByVal reg As String, Optional ByVal ignore_case As Boolean = True) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    
    re.Global = False
    re.IgnoreCase = ignore_case
    re.Pattern = reg

    isMatch = re.Test(txt)
End Function

'''''
' すべてのシートのA1セルを選択して先頭シートを表示します。
'''''
Public Sub scrollAndSelectA1AllSheets()
    Dim i As Integer
    With Application.ActiveWorkbook
        For i = .Worksheets.Count To 1 Step -1
            .Worksheets(i).Activate
            .Worksheets(i).Range("A1").Select
            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = 1
        Next
    End With
End Sub

'''''
' 指定の文字・文字数で右埋めまたは左埋めします。
'''''
Public Function padText(ByVal str As String, ByVal total_width As Long, ByVal pad_char As String, Optional ByVal is_left As Boolean = False) As String
    Const PROC_NAME = MODULE_NAME & ".padText"
    Dim result As String

On Error GoTo Except
    ' 表示上の文字数（全角は2文字とする）を計算
    Dim ch As String, i As Long, display_len As Long
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        If AscW(ch) > 255 Then
            display_len = display_len + 2
        Else
            display_len = display_len + 1
        End If
    Next i

    ' 必要な埋め文字数を計算
    Dim pad_count As Long
    pad_count = total_width - display_len
    If pad_count < 0 Then pad_count = 0

    ' 埋め文字を追加
    Dim padding As String
    padding = String(pad_count, Left(pad_char, 1))
    If is_left Then
        result = padding & str
    Else
        result = str & padding
    End If

    GoTo Finally

Except:
    Err.Raise Err.Number, Err.Source & " " & PROC_NAME, Err.Description

Finally:
    padText = result

End Function

'''''
' LSEG Workspaceから値を取得できていないセルがあるか判定します。
'''''
Public Function isUnloadedLsegWorkspace(ByVal ws As Worksheet) As Boolean
    Dim result As Boolean: result = False

    Dim needle As Variant, unloaded_cell As Range
    For Each needle In Array("Not Signed in", "Retrieving...")
        Set unloaded_cell = ws.UsedRange.Find(What:=needle, LookIn:=xlValues, LookAt:=xlPart)
        If Not unloaded_cell Is Nothing Then
            result = True
            Exit For
        End If
    Next
    
    isUnloadedLsegWorkspace = result
    
End Function

'''''
' Collection[Collection]をインデックスが1から始まる2次元配列へ変換します。
'''''
Public Function nestedCollectionToArray(ByVal records As Collection, Optional ByVal is_start_one = True) As Variant
    Dim result As Variant: result = Array()
    
    If records.Count = 0 Then GoTo Finally
    
    Dim i As Long, j As Long, col_count As Long
    col_count = records(1).Count
    For i = 1 To records.Count
        If col_count <> records.Item(i).Count Then
            Err.Raise 999, "", "inner collectionの要素数が行によって異なります。"
        End If
    Next
    
    Dim item_offset As Integer: item_offset = -1
    If is_start_one Then
        item_offset = 0
    End If
    
    ReDim result(1 + item_offset To records.Count + item_offset, 1 + item_offset To records(1).Count + item_offset)
    For i = 1 To records.Count + item_offset
        For j = 1 To records(1).Count + item_offset
            result(i + item_offset, j + item_offset) = records.Item(i).Item(j)
        Next
    Next
    
Finally:
    nestedCollectionToArray = result
    ' NOTE: ActiveSheet.Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
End Function

'''''
' フォルダ選択ダイアログを表示して、選択したフォルダパスを返します。
'''''
Public Function selectDir(Optional ByVal is_multi_select = False) As String
    Const PROC_NAME = MODULE_NAME & ".selectDir"
    Dim dir_path As String: dir_path = ""

On Error GoTo Except
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .title = "フォルダを選択してください。"
        .AllowMultiSelect = is_multi_select
        If .Show = -1 Then
            dir_path = .SelectedItems(1)
        End If
    End With
    
    GoTo Finally

Except:
    Err.Raise Err.Number, Err.Source & " " & PROC_NAME, Err.Description

Finally:
    selectDir = dir_path

End Function

''''
' CSVを読み取ってCollection[Collection]で返却します。
''''
Public Function readCsv( _
    ByVal file_path As String, ByVal csv_charset As String, _
    Optional ByVal has_comma_in_value As Boolean = False, Optional ByVal start_row_idx As Integer = 0 _
) As Collection
    Const PROC_NAME = MODULE_NAME & ".readCsv"
    Dim records As New Collection

On Error GoTo Except
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    Dim txt As String
    With stm
        .Type = 2                       ' 1: バイナリ, 2: テキスト
        .Charset = csv_charset
        .Open
        .LoadFromFile file_path
        txt = .ReadText(-1)             ' -1: 現在のストリーム位置（Position）から末尾まで一気に文字列として取り込む。
'        .Position = 0

        If has_comma_in_value Then
            Dim reg As Object
            Set reg = CreateObject("VBScript.RegExp")
            reg.Pattern = """[^""]*"""
            reg.Global = True
            
            Dim m As Object
            For Each m In reg.Execute(txt)
                txt = Replace(txt, m.Value, Replace(m.Value, ",", "#####"))
            Next
        End If
        
        Dim lines As Variant
        lines = Split(txt, vbCrLf)
        
        Dim row_idx As Long, col_idx As Long, line As Variant, record As Collection
        For row_idx = start_row_idx To UBound(lines)
            If Trim(lines(row_idx)) = "" Then GoTo ContinueForRowIdx
            
            line = Split(lines(row_idx), ",")
            
            Set record = New Collection
            For col_idx = 0 To UBound(line)
                If has_comma_in_value Then
                    record.Add Replace(line(col_idx), "#####", ",")
                Else
                    record.Add line(col_idx)
                End If
            Next
            
            records.Add record
        
ContinueForRowIdx:
        Next
        
        .Close
    End With
    
    GoTo Finally

Except:
    Err.Raise Err.Number, Err.Source & " " & PROC_NAME, Err.Description

Finally:
    Set readCsv = records

End Function

'''''
' シートのオートフィルターがセットされている行を取得して、解除します。
'''''
Public Function clearAutoFilter(ByVal ws As Worksheet) As Long
    Const PROC_NAME = MODULE_NAME & ".clearAutoFilter"
    Dim filter_row_no As Long
    filter_row_no = 0

On Error GoTo Except
    If ws.AutoFilterMode = False Then GoTo Finally
    
    filter_row_no = ws.AutoFilter.Range.Rows(1).Row
    
    ws.AutoFilterMode = False
    
    GoTo Finally

Except:
    Err.Raise Err.Number, Err.Source & " " & PROC_NAME, Err.Description

Finally:
    clearAutoFilter = filter_row_no

End Function

'''''
' シートのオートフィルターをセットします。
'''''
Public Sub setAutoFilter(ByVal ws As Worksheet, ByVal filter_row_no As Long, Optional ByVal filter_col_no_from = 1)
    Const PROC_NAME = MODULE_NAME & ".setAutoFilter"

On Error GoTo Except
    Dim last_row As Long, last_col As Long
    last_row = lastRow(ws)
    last_col = lastCol(ws)
    
    ws.Range(ws.Cells(filter_row_no, filter_col_no_from), ws.Cells(last_row, last_col).Address).AutoFilter
    
    
    GoTo Finally

Except:
    Err.Raise Err.Number, Err.Source & " " & PROC_NAME, Err.Description

Finally:

End Sub

'''''
' ADODBでSQLを発行します。
'''''
Public Function query(ByVal wb_path As String, ByVal sql As String, ByVal tmp_ws As Worksheet) As Collection
    Const PROC_NAME = MODULE_NAME & ".query"
    Const AD_ASYNC_CONNECTION As Long = 16
    Const AD_STATE_OPEN As Long = 1
    Const AD_OPEN_STATIC As Long = 3
    Const AD_LOCK_READONLY As Long = 1
    Const AD_USE_CLIENT As Long = 3
    
    Dim result As New Collection
    Dim is_thisworkbook As Boolean: is_thisworkbook = False
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    Dim rs As Object: Set rs = CreateObject("ADODB.RecordSet")

On Error GoTo Except
    Dim data_src_path As String
    data_src_path = wb_path
    
    ' ThisWorkbookだと接続が不安定なため、TEMPディレクトリにコピーしてそこから読み取ります。
    If StrComp(wb_path, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        is_thisworkbook = True
        data_src_path = Environ$("TEMP") & "¥ado_tmp_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
        ThisWorkbook.SaveCopyAs data_src_path
        SetAttr data_src_path, vbReadOnly
    End If

    Dim cs As String
    cs = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & data_src_path & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1;ReadOnly=1"";"
        
    conn.ConnectionTimeout = 30
    conn.CommandTimeout = 60
    conn.Open cs, , , AD_ASYNC_CONNECTION
    
    Dim t0 As Date: t0 = Now
    Do While conn.State <> AD_STATE_OPEN
        DoEvents
        If DateDiff("s", t0, Now) > 30 Then Err.Raise vbObjectError + 513, PROC_NAME, "Connection open timeout."
    Loop
    
    rs.CursorLocation = AD_USE_CLIENT
    rs.Open sql, conn, AD_OPEN_STATIC, AD_LOCK_READONLY
    
    If rs.EOF And rs.BOF Then
        rs.Close: Set rs = Nothing
        conn.Close: Set conn = Nothing
        GoTo Finally
    End If
    
    Dim field_names As New Collection
    Dim i As Long
    For i = 1 To rs.Fields.Count
        field_names.Add rs.Fields(i - 1).Name
    Next
    
    tmp_ws.Cells.ClearContents
    tmp_ws.Range("A1").CopyFromRecordset rs
    
    Set rs.ActiveConnection = Nothing
    rs.Close: Set rs = Nothing
    conn.Close: Set conn = Nothing
    
    Dim last_row As Long
    last_row = lastRow(tmp_ws)
    
    Dim data As Variant
    data = tmp_ws.Range( _
        "A1:" & tmp_ws.Cells(last_row, field_names.Count).Address _
    ).Value
    
    Dim row_no As Long, col_no As Long, record As Object
    For row_no = 1 To UBound(data, 1)
        Set record = CreateObject("Scripting.Dictionary")
        For col_no = 1 To UBound(data, 2)
            record.Add field_names(col_no), data(row_no, col_no)
        Next
        result.Add record
    Next
    
    tmp_ws.Cells.ClearContents
    
    If is_thisworkbook Then
        Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
        Call fso.deleteFile(data_src_path, True)
    End If
    
    GoTo Finally

Except:
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State <> 0 Then conn.Close
        Set conn = Nothing
    End If
    Err.Raise Err.Number, Err.Source & " " & PROC_NAME, Err.Description

Finally:
    Set query = result

'    Dim sql As String
'    sql = ""
'    sql = sql & "SELECT A.ID, A.ValueA1, B.ValueB2, C.ValueC3, D.ValueD3 "
'    sql = sql & "FROM ((([Sheet1$] A "
'    sql = sql & "INNER JOIN [Sheet2$] B ON A.ID = B.ID) "
'    sql = sql & "INNER JOIN [Sheet3$] C ON B.ID = C.ID) "
'    sql = sql & "INNER JOIN [Sheet4$] D ON C.ID = D.ID) "

End Function


' [VBA]OneDriveで同期しているファイルまたはフォルダのURLをローカルパスに変換する関数
' Copyright (c) 2020-2025  黒箱
' This software is released under the GPLv3.
' このソフトウェアはGNU GPLv3の下でリリースされています。
'* @fn Public Function OneDriveUrlToLocalPath(ByRef Url As String) As String
'* @brief OneDriveのファイルURL又はフォルダURLをローカルパスに変換します。
'* @param[in] Url OneDrive内に保存されたのファイル又はフォルダのURL
'* @return Variant ローカルパスを返します。引数Urlにローカルパスに"https://"以外から始まる文字列を指定した場合、引数Urlを返します。
'* @details OneDriveのファイルURL又はフォルダURLをローカルパスに変換します。本関数は、ExcelブックがOneDrive内に格納されている場合に、Workbook.Path又はWorkbook.FullNameがURLを返す問題を解決するためのものです。
'*
Public Function OneDriveUrlToLocalPath(ByRef url As String) As String
    Const OneDriveCommercialUrlPattern As String = "*my.sharepoint.com*" '法人向けOneDriveのURLか否かを判定するためのLike右辺値

    '引数がURLでない場合、引数はローカルパスと判断してそのまま返す。
    If Not (url Like "https://*") Then
        OneDriveUrlToLocalPath = url
        Exit Function
    End If
    
    'OneDriveのパスを取得しておく(パフォーマンス優先)。
    Static PathSeparator As String
    Static OneDriveCommercialPath As String
    Static OneDriveConsumerPath As String
    
    If (PathSeparator = "") Then
        'PowerPointではApplication.PathSeparatorが定義されていないため修正。2023/06/20
        'PathSeparator = Application.PathSeparator
        #If Mac Then
            PathSeparator = "/"
        #Else
            PathSeparator = Chr(92)
        #End If
        
        '法人向けOneDrive(OneDrive for Business)のパス
        OneDriveCommercialPath = Environ("OneDriveCommercial")
        If (OneDriveCommercialPath = "") Then OneDriveCommercialPath = Environ("OneDrive")
        
        '個人向けOneDriveのパス
        OneDriveConsumerPath = Environ("OneDriveConsumer")
        If (OneDriveConsumerPath = "") Then OneDriveConsumerPath = Environ("OneDrive")

    End If
    
    '法人向けOneDrive：URL＝"https://会社名-my.sharepoint.com/personal/ユーザー名_domain_com/Documentsファイルパス")
    Dim FilePathPos As Long
    If (url Like OneDriveCommercialUrlPattern) Then
        FilePathPos = InStr(1, url, "/Documents") + 10 '10 = Len("/Documents")
        OneDriveUrlToLocalPath = OneDriveCommercialPath & Replace(Mid(url, FilePathPos), "/", PathSeparator)
        
    '個人向けOneDrive：URL＝"https://d.docs.live.net/CID番号/ファイルパス"
    Else
        FilePathPos = InStr(9, url, "/") '9 == Len("https://") + 1
        FilePathPos = InStr(FilePathPos + 1, url, "/")

        If (FilePathPos = 0) Then
            OneDriveUrlToLocalPath = OneDriveConsumerPath
        Else
            OneDriveUrlToLocalPath = OneDriveConsumerPath & Replace(Mid(url, FilePathPos), "/", PathSeparator)
        End If
    End If

End Function

