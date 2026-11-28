Attribute VB_Name = "modMentionCsvImport"
Option Explicit

Private Const INBOX_NAME As String = "書き込み用"
Private Const OUTBOX_NAME As String = "書き出し用"
Private Const SUMMARY_SHEET As String = "集計"
Private Const HISTORY_SHEET As String = "取込履歴"
Private Const CSV_COLUMNS As Long = 21
Private Const MAX_REQUESTS As Long = 3
Private Const CSV_QUOTE As String = """"
Private Const IMPORT_ERROR As Long = vbObjectError + 1000
Private Const IN_PROGRESS As String = "対応中"
Private Const COMPLETED As String = "完了"
Private Const IN_PROGRESS_PREFIX As String = IN_PROGRESS & "_"
Private Const COMPLETED_PREFIX As String = COMPLETED & "_"

Private Const CSV_HEADERS As String = _
    "送信日時|拠点|依頼者|依頼No.|組織|CA名|出社状況|" & _
    "代理CA1|代理CA2|代理CA3|代理CA4|代理CA5|代理CA6|代理CA7|代理CA8|" & _
    "オプション|メールメモ|処理日時|期日または面接日程|タイプ|RequestID"

Public Sub ImportMentionRequests()
    Dim book As Workbook
    Dim summary As Worksheet, history As Worksheet
    Dim files As Object, candidate As Object, archiveDays As Object
    Dim newNames As New Collection, archiveNames As New Collection
    Dim requests As New Collection, fileRequests As Collection
    Dim statusNames As New Collection, completedNames As New Collection
    Dim requestRows As Object, statusPaths As Object, statusIds As Object, oldStatuses As Object
    Dim requestRow As Variant, fileName As Variant
    Dim fields As Variant, currentStatus As Variant, requestId As String, row As Long
    Dim statusPath As String, isCompleted As Boolean, isLegacy As Boolean
    Dim statusCell As Range
    Dim inbox As String, outbox As String, logPath As String, importDay As String
    Dim dataTarget As Range, historyTarget As Range
    Dim oldEvents As Boolean, oldCancelKey As XlEnableCancelKey
    Dim restoreApplication As Boolean, saved As Boolean
    Dim failure As String

    Set book = ThisWorkbook
    If book.ReadOnly Then Exit Sub
    If Len(book.Path) = 0 Then Exit Sub
    On Error GoTo Failed

    Set files = CreateObject("Scripting.FileSystemObject")
    logPath = files.BuildPath(book.Path, "集計エラー.log")
    inbox = files.BuildPath(book.Path, INBOX_NAME)
    If Not files.FolderExists(inbox) Then
        Err.Raise IMPORT_ERROR, , "フォルダが見つかりません " & inbox
    End If
    importDay = Format$(Date, "yyyymmdd")

    oldEvents = Application.EnableEvents
    oldCancelKey = Application.EnableCancelKey
    restoreApplication = True
    Application.EnableEvents = False
    Application.EnableCancelKey = xlErrorHandler

    Set summary = PrepareSheet(book, SUMMARY_SHEET, CSV_HEADERS & "|対応状況")
    Set history = PrepareSheet(book, HISTORY_SHEET, "CSVファイル名|集計日")
    Set archiveDays = ReadHistory(history)
    Set statusPaths = CreateObject("Scripting.Dictionary")
    statusPaths.CompareMode = vbTextCompare

    ' 作業用Excelが参照するCSVを残し、未取込の同名ファイルは書き出し用を優先する。
    outbox = files.BuildPath(book.Path, OUTBOX_NAME)
    If files.FolderExists(outbox) Then
        For Each candidate In files.GetFolder(outbox).Files
            If StrComp(files.GetExtensionName(candidate.Name), "csv", vbTextCompare) = 0 Then
                fileName = CStr(candidate.Name)
                If Not archiveDays.Exists(fileName) Then
                    QueueStatusCsv CStr(fileName), CStr(candidate.Path), statusPaths, statusNames, completedNames
                End If
            End If
        Next candidate
    End If

    For Each candidate In files.GetFolder(inbox).Files
        If StrComp(files.GetExtensionName(candidate.Name), "csv", vbTextCompare) = 0 Then
            fileName = CStr(candidate.Name)
            If Not QueueStatusCsv(CStr(fileName), CStr(candidate.Path), statusPaths, statusNames, completedNames) Then
                If archiveDays.Exists(fileName) Then
                    archiveNames.Add fileName
                Else
                    Set fileRequests = ReadCsv(CStr(candidate.Path), logPath)
                    If Not fileRequests Is Nothing Then
                        For Each requestRow In fileRequests
                            requests.Add requestRow
                        Next requestRow
                        newNames.Add fileName
                        archiveDays.Add fileName, importDay
                        archiveNames.Add fileName
                    End If
                End If
            End If
        End If
    Next candidate

    If newNames.Count > 0 Then
        Set dataTarget = summary.Cells(NextRow(summary.Range("A:V")), 1) _
            .Resize(requests.Count, CSV_COLUMNS)
        dataTarget.Value2 = RequestValues(requests)
    End If

    ' 同じ回の新規依頼にも反映し、完了を対応中より後に処理する。
    For Each fileName In completedNames
        statusNames.Add fileName
    Next fileName
    If statusNames.Count > 0 Then
        Set requestRows = ReadRequestRows(summary)
        Set statusIds = CreateObject("Scripting.Dictionary")
        Set oldStatuses = CreateObject("Scripting.Dictionary")
        For Each fileName In statusNames
            statusPath = CStr(statusPaths(fileName))
            isCompleted = Left$(fileName, Len(COMPLETED_PREFIX)) = COMPLETED_PREFIX
            isLegacy = StrComp(files.GetParentFolderName(statusPath), inbox, vbTextCompare) = 0
            fields = ReadStatusCsv(statusPath, isCompleted, logPath)
            If IsArray(fields) Then
                requestId = fields(1)
                If archiveDays.Exists(fileName) Then
                    If isLegacy Then statusIds.Add fileName, requestId
                Else
                    row = 0
                    If requestRows.Exists(requestId) Then row = requestRows(requestId)
                    If row = 0 Then
                        LogProblem logPath, "読込 " & fileName, "RequestIDを一意に照合できません。"
                    Else
                        Set statusCell = summary.Cells(row, CSV_COLUMNS + 1)
                        currentStatus = statusCell.Value2
                        If IsError(currentStatus) Then currentStatus = Empty
                        If isCompleted Or CStr(currentStatus) <> COMPLETED Then
                            If Not oldStatuses.Exists(row) Then
                                oldStatuses.Add row, Array(statusCell.HasFormula, statusCell.Formula, statusCell.Value2)
                            End If
                            statusCell.Value2 = StatusCellValue(fields(UBound(fields)))
                        End If
                        newNames.Add fileName
                        archiveDays.Add fileName, importDay
                        If isLegacy Then statusIds.Add fileName, requestId
                    End If
                End If
            End If
        Next fileName
    End If

    ' 集計行、対応状況、CSVの取込履歴を同じ保存で確定する。
    If newNames.Count > 0 Then
        Set historyTarget = history.Cells(NextRow(history.Range("A:B")), 1) _
            .Resize(newNames.Count, 2)
        historyTarget.Value2 = HistoryValues(newNames, importDay)
    End If

    If Not book.Saved Then
        book.Save
        If Not book.Saved Then Err.Raise IMPORT_ERROR, , "上書き保存を確認できません。"
    End If
    saved = True

    For Each fileName In archiveNames
        MoveCsv files, inbox, CStr(fileName), CStr(archiveDays(fileName)), logPath
    Next fileName
    If statusNames.Count > 0 Then
        ArchiveStatusCsvs files, inbox, statusNames, statusIds, archiveDays, logPath
    End If

CleanUp:
    On Error Resume Next
    If restoreApplication Then
        Application.EnableCancelKey = oldCancelKey
        Application.EnableEvents = oldEvents
    End If
    Exit Sub

Failed:
    failure = Err.Description
    On Error Resume Next
    If Not saved Then
        If Not oldStatuses Is Nothing Then RestoreStatuses summary, oldStatuses
        If Not dataTarget Is Nothing Then dataTarget.ClearContents
        If Not historyTarget Is Nothing Then historyTarget.ClearContents
    End If
    LogProblem logPath, "集計", failure
    GoTo CleanUp
End Sub

Private Function QueueStatusCsv(ByVal fileName As String, ByVal filePath As String, _
                                ByVal statusPaths As Object, ByVal statusNames As Collection, _
                                ByVal completedNames As Collection) As Boolean
    Dim matchingNames As Collection
    If Left$(fileName, Len(IN_PROGRESS_PREFIX)) = IN_PROGRESS_PREFIX Then
        Set matchingNames = statusNames
    ElseIf Left$(fileName, Len(COMPLETED_PREFIX)) = COMPLETED_PREFIX Then
        Set matchingNames = completedNames
    Else
        Exit Function
    End If

    QueueStatusCsv = True
    If statusPaths.Exists(fileName) Then Exit Function
    statusPaths.Add fileName, filePath
    matchingNames.Add fileName
End Function

Private Function PrepareSheet(ByVal book As Workbook, ByVal sheetName As String, _
                              ByVal headerText As String) As Worksheet
    Dim ws As Worksheet, expected As Variant, actual As Variant, column As Long
    On Error Resume Next
    Set ws = book.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count))
        ws.Name = sheetName
    End If

    expected = Split(headerText, "|")
    actual = ws.Cells(1, 1).Resize(1, UBound(expected) + 1).Value2
    For column = 1 To UBound(expected) + 1
        If Not IsEmpty(actual(1, column)) Then
            If CStr(actual(1, column)) <> expected(column - 1) Then
                Err.Raise IMPORT_ERROR, , "シート「" & sheetName & "」の見出しを確認してください。列 " & column
            End If
        End If
    Next column
    For column = 1 To UBound(expected) + 1
        If IsEmpty(actual(1, column)) Then ws.Cells(1, column).Value2 = expected(column - 1)
    Next column
    Set PrepareSheet = ws
End Function

Private Function NextRow(ByVal searchRange As Range) As Long
    Dim lastCell As Range
    Set lastCell = searchRange.Find(What:="*", After:=searchRange.Cells(1, 1), _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    NextRow = 2
    If Not lastCell Is Nothing Then NextRow = lastCell.Row + 1
End Function

Private Function ReadHistory(ByVal ws As Worksheet) As Object
    Dim days As Object, entries As Variant, lastRow As Long, row As Long
    Dim fileName As String, importDay As String
    Set days = CreateObject("Scripting.Dictionary")
    days.CompareMode = vbTextCompare

    lastRow = NextRow(ws.Range("A:B")) - 1
    If lastRow >= 2 Then
        entries = ws.Cells(2, 1).Resize(lastRow - 1, 2).Value2
        For row = 1 To UBound(entries, 1)
            fileName = CStr(entries(row, 1))
            importDay = CStr(entries(row, 2))
            If Len(fileName) > 0 Or Len(importDay) > 0 Then
                If Len(fileName) = 0 Or Not (importDay Like "########") Then
                    Err.Raise IMPORT_ERROR, , "取込履歴を確認してください。行 " & (row + 1)
                End If
                If days.Exists(fileName) Then
                    Err.Raise IMPORT_ERROR, , "取込履歴に同じファイル名があります " & fileName
                End If
                days.Add fileName, importDay
            End If
        Next row
    End If
    Set ReadHistory = days
End Function

Private Function RequestValues(ByVal requests As Collection) As Variant
    Dim values() As Variant, fields As Variant, row As Long, column As Long
    ReDim values(1 To requests.Count, 1 To CSV_COLUMNS)
    For row = 1 To requests.Count
        fields = requests(row)
        For column = 1 To CSV_COLUMNS
            ' 文字列として渡し、先頭のゼロや記号を保持する。
            If Len(fields(column)) > 0 Then values(row, column) = "'" & fields(column)
        Next column
    Next row
    RequestValues = values
End Function

Private Function HistoryValues(ByVal names As Collection, ByVal importDay As String) As Variant
    Dim values() As Variant, row As Long
    ReDim values(1 To names.Count, 1 To 2)
    For row = 1 To names.Count
        values(row, 1) = "'" & names(row)
        values(row, 2) = "'" & importDay
    Next row
    HistoryValues = values
End Function

Private Function ReadCsv(ByVal filePath As String, ByVal logPath As String) As Collection
    Dim csvText As String
    On Error GoTo Failed
    csvText = ReadCsvText(filePath)
    Set ReadCsv = ParseCsv(csvText)
    Exit Function

Failed:
    If Err.Number = 18 Then Err.Raise 18
    LogProblem logPath, "読込 " & filePath, Err.Description
End Function

Private Function ReadCsvText(ByVal filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "shift_jis"
    stream.Open
    stream.LoadFromFile filePath
    ReadCsvText = stream.ReadText(-1)
    stream.Close
End Function

Private Function ParseCsv(ByRef csvText As String) As Collection
    Dim rows As New Collection
    Dim position As Long, column As Long, fields As Variant, headers As Variant
    position = 1
    fields = ReadCsvRecord(csvText, position)
    headers = Split(CSV_HEADERS, "|")
    For column = 1 To CSV_COLUMNS
        If fields(column) <> headers(column - 1) Then
            Err.Raise IMPORT_ERROR, , "CSVの見出しが一致しません。列 " & column
        End If
    Next column

    Do While position <= Len(csvText)
        If rows.Count = MAX_REQUESTS Then Err.Raise IMPORT_ERROR, , "依頼データが3件を超えています。"
        fields = ReadCsvRecord(csvText, position)
        rows.Add fields
    Loop
    If rows.Count = 0 Then Err.Raise IMPORT_ERROR, , "依頼データがありません。"
    Set ParseCsv = rows
End Function

Private Function ReadCsvRecord(ByRef csvText As String, ByRef position As Long, _
                               Optional ByVal columnCount As Long = CSV_COLUMNS) As Variant
    Dim fields() As String
    Dim column As Long
    ReDim fields(1 To columnCount)
    For column = 1 To columnCount
        fields(column) = ReadCsvField(csvText, position)
        If Len(fields(column)) > 32767 Or InStr(fields(column), vbNullChar) > 0 Then
            Err.Raise IMPORT_ERROR, , "Excelのセルに収まらない値があります。列 " & column
        End If
        If column < columnCount Then
            If Mid$(csvText, position, 1) <> "," Then
                Err.Raise IMPORT_ERROR, , "CSVの列数が" & columnCount & "列に満たない行があります。"
            End If
            position = position + 1
        End If
    Next column

    Select Case Mid$(csvText, position, 1)
        Case vbCr
            position = position + 1
            If Mid$(csvText, position, 1) = vbLf Then position = position + 1
        Case vbLf
            position = position + 1
        Case vbNullString
            ' ファイル末尾。最終行は改行なしでも受け付ける。
        Case Else
            Err.Raise IMPORT_ERROR, , "CSVの" & columnCount & "列目に続く区切りが正しくありません。"
    End Select
    ReadCsvRecord = fields
End Function

Private Function ReadCsvField(ByRef csvText As String, ByRef position As Long) As String
    Dim startAt As Long, quoteAt As Long
    If Mid$(csvText, position, 1) = CSV_QUOTE Then
        startAt = position + 1
        quoteAt = InStr(startAt, csvText, CSV_QUOTE)
        Do While quoteAt > 0
            If Mid$(csvText, quoteAt + 1, 1) <> CSV_QUOTE Then
                ReadCsvField = Replace(Mid$(csvText, startAt, quoteAt - startAt), _
                                       CSV_QUOTE & CSV_QUOTE, CSV_QUOTE)
                position = quoteAt + 1
                Exit Function
            End If
            quoteAt = InStr(quoteAt + 2, csvText, CSV_QUOTE)
        Loop
        Err.Raise IMPORT_ERROR, , "CSVの引用符が閉じていません。"
    End If

    startAt = position
    Do While position <= Len(csvText)
        Select Case Mid$(csvText, position, 1)
            Case ",", vbCr, vbLf
                Exit Do
            Case CSV_QUOTE
                Err.Raise IMPORT_ERROR, , "CSVの引用符の位置が正しくありません。"
        End Select
        position = position + 1
    Loop
    ReadCsvField = Mid$(csvText, startAt, position - startAt)
End Function

Private Function MoveCsv(ByVal files As Object, ByVal inbox As String, ByVal fileName As String, _
                         ByVal importDay As String, ByVal logPath As String, _
                         Optional ByVal requestId As String = "") As Boolean
    Dim oldFolder As String, dayFolder As String, destination As String
    On Error GoTo Failed
    oldFolder = files.BuildPath(inbox, "OLD")
    dayFolder = files.BuildPath(oldFolder, importDay)
    If Not files.FolderExists(oldFolder) Then files.CreateFolder oldFolder
    If Not files.FolderExists(dayFolder) Then files.CreateFolder dayFolder
    If Len(requestId) > 0 And Left$(fileName, Len(IN_PROGRESS_PREFIX)) = IN_PROGRESS_PREFIX Then
        dayFolder = files.BuildPath(dayFolder, IN_PROGRESS)
        If Not files.FolderExists(dayFolder) Then files.CreateFolder dayFolder
    End If
    destination = files.BuildPath(dayFolder, fileName)
    If Len(requestId) > 0 And Left$(fileName, Len(COMPLETED_PREFIX)) = COMPLETED_PREFIX Then
        If files.FileExists(destination) Then Err.Raise IMPORT_ERROR, , "移動先に同名のCSVがあります。"
        ' 削除に失敗した場合は完了CSVを入力側に残し、保存済みの履歴で再試行する。
        DeleteInProgressCsvs files, oldFolder, requestId
    End If
    files.MoveFile files.BuildPath(inbox, fileName), destination
    MoveCsv = True
    Exit Function

Failed:
    If Err.Number = 18 Then Err.Raise 18
    LogProblem logPath, "移動 " & fileName, Err.Description
End Function

Private Function ReadStatusCsv(ByVal filePath As String, ByVal isCompleted As Boolean, _
                               ByVal logPath As String) As Variant
    Dim csvText As String, position As Long, fields As Variant
    Dim columnCount As Long
    On Error GoTo Failed
    csvText = ReadCsvText(filePath)
    position = 1
    columnCount = 2
    If isCompleted Then columnCount = 3
    fields = ReadCsvRecord(csvText, position, columnCount)
    If position <= Len(csvText) Or Len(fields(1)) = 0 Then
        Err.Raise IMPORT_ERROR, , "対応状況CSVにはRequestIDを含む1行が必要です。"
    End If
    If isCompleted Then
        If fields(3) <> COMPLETED Then Err.Raise IMPORT_ERROR, , "完了CSVのC1を確認してください。"
    End If
    ReadStatusCsv = fields
    Exit Function

Failed:
    If Err.Number = 18 Then Err.Raise 18
    LogProblem logPath, "読込 " & filePath, Err.Description
End Function

Private Function ReadRequestRows(ByVal summary As Worksheet) As Object
    Dim rows As Object, values As Variant, row As Long, lastRow As Long
    Dim requestId As String
    Set rows = CreateObject("Scripting.Dictionary")
    rows.CompareMode = vbBinaryCompare
    lastRow = NextRow(summary.Columns(CSV_COLUMNS)) - 1
    If lastRow >= 2 Then
        ' 見出しから取得して、依頼が1件の場合も配列で受け取る。
        values = summary.Cells(1, CSV_COLUMNS).Resize(lastRow, 1).Value2
        For row = 2 To lastRow
            If Not IsError(values(row, 1)) Then
                requestId = CStr(values(row, 1))
                If Len(requestId) > 0 Then
                    If rows.Exists(requestId) Then
                        rows(requestId) = 0
                    Else
                        rows.Add requestId, row
                    End If
                End If
            End If
        Next row
    End If
    Set ReadRequestRows = rows
End Function

Private Function StatusCellValue(ByVal value As Variant) As Variant
    StatusCellValue = value
    If VarType(value) = vbString Then
        If Len(value) > 0 Then StatusCellValue = "'" & value
    End If
End Function

Private Sub RestoreStatuses(ByVal summary As Worksheet, ByVal oldStatuses As Object)
    Dim row As Variant, previous As Variant
    On Error Resume Next
    For Each row In oldStatuses.Keys
        previous = oldStatuses(row)
        With summary.Cells(CLng(row), CSV_COLUMNS + 1)
            If previous(0) Then
                .Formula = previous(1)
            Else
                .Value2 = StatusCellValue(previous(2))
            End If
        End With
    Next row
End Sub

Private Sub ArchiveStatusCsvs(ByVal files As Object, ByVal inbox As String, ByVal statusNames As Collection, _
                              ByVal statusIds As Object, ByVal archiveDays As Object, ByVal logPath As String)
    Dim blocked As Object, fileName As Variant, requestId As String
    Set blocked = CreateObject("Scripting.Dictionary")
    For Each fileName In statusNames
        If statusIds.Exists(fileName) Then
            requestId = statusIds(fileName)
            If Left$(fileName, Len(IN_PROGRESS_PREFIX)) = IN_PROGRESS_PREFIX Then
                If Not MoveCsv(files, inbox, CStr(fileName), CStr(archiveDays(fileName)), logPath, requestId) Then
                    blocked(requestId) = True
                End If
            ElseIf Not blocked.Exists(requestId) Then
                MoveCsv files, inbox, CStr(fileName), CStr(archiveDays(fileName)), logPath, requestId
            End If
        End If
    Next fileName
End Sub

Private Sub DeleteInProgressCsvs(ByVal files As Object, ByVal oldFolder As String, ByVal requestId As String)
    Dim dayFolder As Object, candidate As Object, paths As New Collection, path As Variant
    Dim progressFolder As String, prefix As String, wrappedPrefix As String
    Dim csvText As String, archivedId As String, position As Long
    prefix = IN_PROGRESS_PREFIX & requestId
    wrappedPrefix = IN_PROGRESS_PREFIX & "(" & requestId & ")"
    For Each dayFolder In files.GetFolder(oldFolder).SubFolders
        If dayFolder.Name Like "########" Then
            progressFolder = files.BuildPath(dayFolder.Path, IN_PROGRESS)
            If files.FolderExists(progressFolder) Then
                For Each candidate In files.GetFolder(progressFolder).Files
                    If StrComp(files.GetExtensionName(candidate.Name), "csv", vbTextCompare) = 0 Then
                        If Left$(candidate.Name, Len(prefix)) = prefix Or _
                           Left$(candidate.Name, Len(wrappedPrefix)) = wrappedPrefix Then
                            csvText = ReadCsvText(CStr(candidate.Path))
                            position = 1
                            archivedId = ReadCsvField(csvText, position)
                            If archivedId = requestId And Mid$(csvText, position, 1) = "," Then
                                paths.Add CStr(candidate.Path)
                            End If
                        End If
                    End If
                Next candidate
            End If
        End If
    Next dayFolder
    For Each path In paths
        files.DeleteFile CStr(path)
    Next path
End Sub

Private Sub LogProblem(ByVal logPath As String, ByVal operation As String, ByVal detail As String)
    Dim files As Object, output As Object, message As String
    On Error Resume Next
    message = Format$(Now, "yyyy-mm-dd hh:nn:ss") & "  " & operation & "  " & detail
    Debug.Print message
    Set files = CreateObject("Scripting.FileSystemObject")
    Set output = files.OpenTextFile(logPath, 8, True, -1) ' 追記、Unicode
    If Not output Is Nothing Then
        output.WriteLine message
        output.Close
    End If
End Sub
