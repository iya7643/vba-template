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
' フォルダがあるかチェックします。
'''''
Public Function isExistsFolder(ByVal path As String) As Boolean
    isExistsFolder = (Dir(path, vbDirectory) <> "")
End Function

'''''
' ファイルがあるかチェックします。
'''''
Public Function isExistsFile(ByVal path As String) As Boolean
    isExistsFile = (Dir(path, vbNormal) <> "")
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
' 非表示で別のブックを開きます。
'''''
Public Function openBookInvisible(ByVal file_path As String, Optional ByVal pass As String = "") As Workbook
    Dim wb As Workbook
    
    If isFileOpen(file_path) Then
        Set wb = Workbooks(Dir(file_path))
    Else
        If pass = "" Then
            Set wb = Workbooks.Open(Filename:=file_path, ReadOnly:=True)
        Else
            Set wb = Workbooks.Open(Filename:=file_path, ReadOnly:=True, Password:=pass)
        End If
    End If
    wb.Windows(1).Visible = False
    
    Set openBookInvisible = wb
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
Public Function dictor(ByVal dict As Object, ByVal keys As Variant)
    Const PROC_NAME = MODULE_NAME & ".dictor"
    Dim i As Integer, key As Variant, tmp As Object
    
On Error GoTo Except
    Set tmp = dict
    For i = 0 To UBound(keys) - 1
        key = keys(i)
        Set tmp = tmp.Item(key)
    Next

    key = keys(UBound(keys))
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
Public Function readCsv(ByVal file_path As String, ByVal csv_charset As String, Optional ByVal start_row_idx As Integer = 0) As Collection
    Const PROC_NAME = MODULE_NAME & ".readCsv"
    Dim records As New Collection

On Error GoTo Except
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    Dim txt As String, lines As Variant, line As Variant, record As Collection
    Dim row_idx As Long, col_idx As Long
    With stm
        .Charset = csv_charset
        .Open
        .LoadFromFile file_path
        txt = .ReadText(-1)             ' -1: 現在のストリーム位置（Position）から末尾まで一気に文字列として取り込む。
'        .Position = 0
        lines = Split(txt, vbCrLf)
        For row_idx = start_row_idx To UBound(lines)
            line = Split(lines(row_idx), ",")
            
            Set record = New Collection
            For col_idx = 0 To UBound(line)
                record.Add line(col_idx)
            Next
            
            records.Add record
        Next
        .Close
    End With
    
    GoTo Finally

Except:
    Err.Raise Err.Number, Err.Source & " " & PROC_NAME, Err.Description

Finally:
    Set readCsv = records

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

