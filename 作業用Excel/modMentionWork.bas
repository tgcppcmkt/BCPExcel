Attribute VB_Name = "modMentionWork"
Option Explicit

' 導入手順
' 編集できる状態の作業用 .xlsm に、日本語版WindowsのVBAエディターから
' このCP932の .bas ファイルをインポートします。
' 更新の場合は、既存の modMentionWork と modMentionWorkStorage を取り外してから、
' 新しい modMentionWork を1本だけ取り込みます。
' 作業用 .xlsm、「メンション集計用」Excel、「書き込み用」フォルダを同じ階層に置きます。
' 下の2つのイベントは各行の先頭のコメント記号を外し、作業用ブックの
' ThisWorkbook に貼り付けます。既存の同名イベントがある場合は置き換えます。
' 取り込み後は［デバッグ］→［VBAProject のコンパイル］で確認し、.xlsm を保存します。
'
' Private Sub Workbook_Open()
'     modMentionWork.RefreshMentionWork
' End Sub
'
' Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
'     modMentionWork.PickMentionRow Target
' End Sub

Private Const SUMMARY_BOOK As String = "メンション集計用"
Private Const SUMMARY_SHEET As String = "集計"
Private Const WORK_SHEET As String = "依頼一覧"
Private Const STATE_HEADER As String = "対応状況"
Private Const SOURCE_HEADERS As String = _
    "送信日時|拠点|依頼者|依頼No.|組織|CA名|出社状況|" & _
    "代理CA1|代理CA2|代理CA3|代理CA4|代理CA5|代理CA6|代理CA7|代理CA8|" & _
    "オプション|メールメモ|処理日時|期日または面接日程|タイプ|RequestID"
Private Const OUTBOX As String = "書き込み用"
Private Const UNASSIGNED As String = "未対応"
Private Const IN_PROGRESS As String = "対応中"
Private Const COMPLETED As String = "完了"
Private Const PICKABLE As String = "選択可"
Private Const PICKED As String = "選択済"
Private Const PICKED_ROW_FORMULA As String = "=$C2=""選択済"""
Private Const SOURCE_COLS As Long = 22
Private Const SOURCE_ID As Long = 21
Private Const PICK_COL As Long = 3
Private Const DISPLAY_STATE As Long = 4
Private Const FIRST_DATA_COL As Long = 5
Private Const LAST_DATA_COL As Long = 25
Private Const WORK_ERROR As Long = vbObjectError + 2100
Private Const BUTTON_ROW_HEIGHT As Double = 42
Private Const REFRESH_BUTTON_COLUMN_WIDTH As Double = 15
Private Const SELECTION_BUTTON_COLUMN_WIDTH As Double = 19
Private Const BUTTON_PADDING As Double = 5

Private mStates As Object
Private mPicked As Object
Private mSourceRows As Variant
Private mAccount As String
Private mSelecting As Boolean

Private Const STORAGE_ERROR As Long = vbObjectError + 1600
Private Const STATUS_EXTENSION As String = ".csv"
Private Const QUOTE As String = """"
Private Const MAX_PATH As Long = 260

Private Const GENERIC_WRITE As Long = &H40000000
Private Const DELETE_ACCESS As Long = &H10000
Private Const OPEN_ALWAYS As Long = 4
Private Const FILE_ATTRIBUTE_HIDDEN As Long = 2
Private Const FILE_FLAG_DELETE_ON_CLOSE As Long = &H4000000
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_SHARING_VIOLATION As Long = 32
Private Const ERROR_LOCK_VIOLATION As Long = 33
Private Const ERROR_FILE_EXISTS As Long = 80
Private Const ERROR_ALREADY_EXISTS As Long = 183
Private Const ERROR_DELETE_PENDING As Long = 303

Private Declare PtrSafe Function CreateFileW Lib "kernel32" ( _
    ByVal fileName As LongPtr, ByVal access As Long, ByVal shareMode As Long, _
    ByVal securityAttributes As LongPtr, ByVal creationDisposition As Long, _
    ByVal flagsAndAttributes As Long, ByVal templateFile As LongPtr) As LongPtr
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal handle As LongPtr) As Long
Private Declare PtrSafe Function MoveFileW Lib "kernel32" ( _
    ByVal existingName As LongPtr, ByVal newName As LongPtr) As Long
Private Declare PtrSafe Function DeleteFileW Lib "kernel32" (ByVal fileName As LongPtr) As Long
Private Declare PtrSafe Function GetTempFileNameW Lib "kernel32" ( _
    ByVal directory As LongPtr, ByVal prefix As LongPtr, ByVal unique As Long, _
    ByVal fileNameBuffer As LongPtr) As Long

Public Sub RefreshMentionWork()
    Dim failure As String
    If Not RefreshView(failure) Then MsgBox failure, vbExclamation, "更新"
End Sub

Public Sub ToggleMentionSelection()
    Dim failure As String

    On Error GoTo Failed
    If mSelecting Then
        ConfirmMentionSelection
        Exit Sub
    End If
    If Not RefreshView(failure) Then
        MsgBox failure, vbExclamation, "選択開始"
        Exit Sub
    End If

    Set mPicked = NewRequestMap()
    mSelecting = True
    If PaintSelection() = 0 Then
        EndMentionSelection
        Exit Sub
    End If
    If ActiveCell.Column = PICK_COL Then ActiveCell.Offset(0, 1).Select
    SetSelectionButtonCaption "選択を確定"
    Exit Sub

Failed:
    failure = Err.Description
    On Error Resume Next
    EndMentionSelection
    On Error GoTo 0
    MsgBox failure, vbExclamation, "選択操作"
End Sub

Public Sub PickMentionRow(ByVal target As Range)
    Dim requestId As String, action As String

    If Not mSelecting Then Exit Sub
    If target.Worksheet.Name <> WORK_SHEET Then Exit Sub
    If target.CountLarge <> 1 Then Exit Sub
    If target.Column <> PICK_COL Or target.Row < 2 Then Exit Sub
    On Error GoTo Failed
    requestId = TextValue(target.Worksheet.Cells(target.Row, LAST_DATA_COL).Value2)
    If Len(requestId) = 0 Or Not mStates.Exists(requestId) Then Exit Sub
    action = AvailableAction(mStates(requestId))
    If Len(action) = 0 Then Exit Sub

    If mPicked.Exists(requestId) Then
        mPicked.Remove requestId
        target.Value2 = PICKABLE
        target.Offset(0, 1).Select
    Else
        mPicked.Add requestId, action
        target.Value2 = PICKED
        target.Resize(1, LAST_DATA_COL - PICK_COL + 1).Select
        ' 行を選択したままD列をアクティブにし、C列を再クリックで解除できるようにする。
        target.Offset(0, 1).Activate
    End If
    ThisWorkbook.Saved = True
    Exit Sub

Failed:
    MsgBox Err.Description, vbExclamation, "依頼の選択"
End Sub

Private Sub ConfirmMentionSelection()
    Dim requestId As Variant, action As String
    Dim starts As Long, finishes As Long, skipped As Long, failedCount As Long
    Dim startCount As Long, finishCount As Long
    Dim fresh As Object, data As Variant, outboxPath As String
    Dim details As String, failure As String, reason As String, result As String
    Dim oldCancelKey As XlEnableCancelKey
    Dim screenUpdates As Boolean, restoreApplication As Boolean

    If mPicked.Count = 0 Then
        EndMentionSelection
        Exit Sub
    End If

    ' 選んだ時点の操作を固定し、確定前後に状態が変わっても別の操作へ切り替えない。
    For Each requestId In mPicked.Keys
        If mPicked(requestId) = IN_PROGRESS Then
            startCount = startCount + 1
        Else
            finishCount = finishCount + 1
        End If
    Next requestId
    If finishCount > 0 Then
        If MsgBox("担当開始 " & startCount & "件・完了 " & finishCount & "件" & vbCrLf & _
                  "選択した依頼を確定しますか？", _
                  vbQuestion Or vbYesNo Or vbDefaultButton2, "選択を確定") <> vbYes Then Exit Sub
    End If

    On Error GoTo Failed
    oldCancelKey = Application.EnableCancelKey
    screenUpdates = Application.ScreenUpdating
    restoreApplication = True
    Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False

    outboxPath = GetOutbox()
    Set fresh = ReadWorkSnapshot(outboxPath, data)
    For Each requestId In mPicked.Keys
        action = mPicked(requestId)
        reason = vbNullString
        result = CommitRequest(outboxPath, CStr(requestId), action, fresh, reason)
        Select Case result
            Case "OK"
                If action = COMPLETED Then
                    finishes = finishes + 1
                Else
                    starts = starts + 1
                End If
            Case "SKIP"
                skipped = skipped + 1
            Case Else
                failedCount = failedCount + 1
        End Select
        If Len(reason) > 0 And Len(details) < 480 Then
            details = details & vbCrLf & Left$(CStr(requestId) & "  " & reason, 140)
        End If
    Next requestId
    result = "担当開始 " & starts & "件・完了 " & finishes & "件"
    If skipped > 0 Then result = result & vbCrLf & "スキップ " & skipped & "件"
    If failedCount > 0 Then result = result & vbCrLf & "書き出しエラー " & failedCount & "件"
    If Not RefreshView(failure) Then details = details & vbCrLf & "一覧の更新 " & Left$(failure, 140)
    Application.ScreenUpdating = screenUpdates
    Application.EnableCancelKey = oldCancelKey
    restoreApplication = False
    If Len(details) > 0 Then
        MsgBox result & vbCrLf & details, vbExclamation, "確定結果"
    Else
        MsgBox result, vbInformation, "確定結果"
    End If
    Exit Sub

Failed:
    failure = Err.Description
    On Error Resume Next
    If restoreApplication Then
        Application.ScreenUpdating = screenUpdates
        Application.EnableCancelKey = oldCancelKey
    End If
    On Error GoTo 0
    If starts + finishes > 0 Then
        MsgBox "担当開始 " & starts & "件・完了 " & finishes & "件を保存済みです。" & _
               vbCrLf & failure, vbExclamation, "確定を中断"
    Else
        MsgBox failure, vbExclamation, "確定を中断"
    End If
End Sub

Private Function CommitRequest(ByVal outboxPath As String, ByVal requestId As String, _
                               ByVal action As String, ByVal fresh As Object, _
                               ByRef reason As String) As String
    Dim current As Variant

    On Error GoTo Failed
    CommitRequest = "SKIP"
    reason = "担当または対応状況が変わりました。"
    If Not fresh.Exists(requestId) Then Exit Function
    current = fresh(requestId)
    If AvailableAction(current) <> action Then Exit Function

    If Not WorkPublishStatus(outboxPath, requestId, action, mAccount) Then
        reason = "依頼の状態が変わったか、CSVが作成中または作成済みです。"
        Exit Function
    End If
    reason = vbNullString
    CommitRequest = "OK"
    Exit Function

Failed:
    If Err.Number = 18 Then Err.Raise 18
    reason = Err.Description
    CommitRequest = "ERROR"
End Function

Private Function RefreshView(ByRef failure As String) As Boolean
    Dim data As Variant, states As Object, screenUpdates As Boolean

    screenUpdates = Application.ScreenUpdating
    On Error GoTo Failed
    Application.ScreenUpdating = False

    If Not ThisWorkbook.ReadOnly Then ThisWorkbook.ChangeFileAccess xlReadOnly
    EnsureWorkList
    mAccount = Environ$("USERNAME")
    If Len(mAccount) = 0 Then Err.Raise WORK_ERROR, , "PCのアカウント名を取得できません。"

    Set states = ReadWorkSnapshot(GetOutbox(), data)
    RenderWorkList data, states
    mSourceRows = data
    Set mStates = states
    EndMentionSelection
    RefreshView = True

Finished:
    Application.ScreenUpdating = screenUpdates
    Exit Function

Failed:
    failure = Err.Description
    On Error Resume Next
    EndMentionSelection
    On Error GoTo 0
    GoTo Finished
End Function

Private Function WorkList() As Worksheet
    Set WorkList = ThisWorkbook.Worksheets(WORK_SHEET)
End Function

Private Sub EnsureWorkList()
    Dim ws As Worksheet, restoreAlerts As Boolean
    Dim failureNumber As Long, failureSource As String, failureDescription As String

    On Error Resume Next
    Set ws = WorkList()
    On Error GoTo 0
    If Not ws Is Nothing Then
        UpdateWorkListLayout ws
        Exit Sub
    End If

    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = WORK_SHEET
    FormatWorkList ws
    ThisWorkbook.Saved = True
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    If Not ws Is Nothing Then
        ' 初期化できなかったシートだけを片付け、次の更新で作成を再試行できるようにする。
        restoreAlerts = Application.DisplayAlerts
        On Error Resume Next
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = restoreAlerts
        On Error GoTo 0
    End If
    Err.Raise failureNumber, failureSource, "「" & WORK_SHEET & "」シートを作成できません。" & vbCrLf & failureDescription
End Sub

Private Sub UpdateWorkListLayout(ByVal ws As Worksheet)
    Dim pickedRows As Range, condition As Object
    Dim hasPickedRows As Boolean, hasOldHelp As Boolean

    ws.Unprotect
    ResizeWorkButtons ws
    If ws.Range("A2").MergeArea.Address = "$A$2:$B$2" Then
        ws.Range("A2:B2").UnMerge
        ws.Range("A2:B2").Clear
        ws.Rows(2).RowHeight = ws.StandardHeight
    End If

    Set pickedRows = ws.Range("C2:Y" & ws.Rows.Count)
    For Each condition In pickedRows.Cells(1, 1).FormatConditions
        If condition.Type = xlExpression Then
            If condition.Formula1 = PICKED_ROW_FORMULA And condition.AppliesTo.Address = pickedRows.Address Then
                hasPickedRows = True
                Exit For
            End If
        End If
    Next condition
    hasOldHelp = ws.Range("A3").Value2 = "［選択開始］のあと、C列の「選択可」をクリック。" And _
                 ws.Range("A4").Value2 = "未対応は担当開始、自分の対応中は完了。"
    If hasPickedRows And Not hasOldHelp Then Exit Sub

    If Not hasPickedRows Then AddPickedRowHighlight pickedRows
    If hasOldHelp Then
        ws.Range("A3:B4").UnMerge
        ws.Range("A3:B4").Clear
        ws.Rows("3:4").RowHeight = ws.StandardHeight
    End If
End Sub

Private Sub FormatWorkList(ByVal ws As Worksheet)
    Dim headers As Variant, widths As Variant, column As Long
    Dim lastRow As Long

    headers = Split(SOURCE_HEADERS, "|")
    widths = Array(21, 12, 15, 14, 16, 15, 14, 13, 13, 13, 13, 13, 13, 13, 13, 23, 42, 21, 24, 15, 39)
    With ws
        .Tab.Color = RGB(19, 121, 109)
        .Rows(1).RowHeight = BUTTON_ROW_HEIGHT
        .Columns("A").ColumnWidth = REFRESH_BUTTON_COLUMN_WIDTH
        .Columns("B").ColumnWidth = SELECTION_BUTTON_COLUMN_WIDTH
        .Columns("C").ColumnWidth = 11
        .Columns("D").ColumnWidth = 13
        .Range("C:Y").NumberFormat = "@"
        .Columns("C:D").HorizontalAlignment = xlCenter

        .Cells(1, PICK_COL).Value2 = "選択"
        .Cells(1, DISPLAY_STATE).Value2 = STATE_HEADER
        For column = LBound(headers) To UBound(headers)
            .Cells(1, FIRST_DATA_COL + column).Value2 = headers(column)
            .Columns(FIRST_DATA_COL + column).ColumnWidth = widths(column)
        Next column
        With .Range("A1:Y1")
            .Interior.Color = RGB(32, 59, 81)
            .Font.Color = vbWhite
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        .Range("A1:B1").Interior.Color = RGB(20, 43, 63)
        .Range("C1:D1").Interior.Color = RGB(19, 121, 109)

        AddWorkButton .Range("A1"), "更新", "RefreshMentionWork"
        AddWorkButton .Range("B1"), "選択開始", "ToggleMentionSelection"
        lastRow = .Rows.Count
        ColorWorkCells .Range("C2:C" & lastRow), "=$C2=""選択可""", RGB(229, 244, 240), RGB(24, 84, 77)
        ColorWorkCells .Range("D2:D" & lastRow), "=$D2=""未対応""", RGB(237, 242, 247), RGB(48, 78, 102)
        ColorWorkCells .Range("D2:D" & lastRow), "=$D2=""対応中""", RGB(255, 242, 206), RGB(113, 80, 17)
        ColorWorkCells .Range("D2:D" & lastRow), "=$D2=""完了""", RGB(231, 237, 243), RGB(67, 88, 107)
        AddPickedRowHighlight .Range("C2:Y" & lastRow)
    End With

    If Application.ActiveSheet Is ws Then
        With Application.ActiveWindow
            .DisplayGridlines = False
            .SplitRow = 1
            .SplitColumn = 4
            .FreezePanes = True
        End With
    End If
End Sub

Private Sub AddWorkButton(ByVal anchor As Range, ByVal caption As String, ByVal action As String)
    Dim button As Object

    Set button = anchor.Worksheet.Buttons.Add(anchor.Left, anchor.Top, anchor.Width, anchor.Height)
    With button
        .Caption = caption
        .OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!modMentionWork." & action
    End With
    FitWorkButton button, anchor
End Sub

Private Sub ResizeWorkButtons(ByVal ws As Worksheet)
    With ws
        If .Buttons.Count < 2 Then Exit Sub
        If .Rows(1).RowHeight < BUTTON_ROW_HEIGHT Then .Rows(1).RowHeight = BUTTON_ROW_HEIGHT
        If .Columns("A").ColumnWidth < REFRESH_BUTTON_COLUMN_WIDTH Then
            .Columns("A").ColumnWidth = REFRESH_BUTTON_COLUMN_WIDTH
        End If
        If .Columns("B").ColumnWidth < SELECTION_BUTTON_COLUMN_WIDTH Then
            .Columns("B").ColumnWidth = SELECTION_BUTTON_COLUMN_WIDTH
        End If
        FitWorkButton .Buttons(1), .Range("A1")
        FitWorkButton .Buttons(2), .Range("B1")
    End With
End Sub

Private Sub FitWorkButton(ByVal button As Object, ByVal anchor As Range)
    With button
        .Left = anchor.Left + BUTTON_PADDING
        .Top = anchor.Top + BUTTON_PADDING
        .Width = anchor.Width - BUTTON_PADDING * 2
        .Height = anchor.Height - BUTTON_PADDING * 2
    End With
End Sub

Private Sub ColorWorkCells(ByVal target As Range, ByVal formula As String, _
                           ByVal fillColor As Long, ByVal fontColor As Long)
    With target.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
        .Interior.Color = fillColor
        .Font.Color = fontColor
        .Font.Bold = True
    End With
End Sub

Private Sub AddPickedRowHighlight(ByVal pickedRows As Range)
    With pickedRows.FormatConditions.Add(Type:=xlExpression, Formula1:=PICKED_ROW_FORMULA)
        .Interior.Color = RGB(200, 233, 223)
        .Font.Color = RGB(23, 60, 56)
        .SetFirstPriority
        .StopIfTrue = True
    End With
End Sub

Private Function ReadWorkSnapshot(ByVal outboxPath As String, ByRef data As Variant) As Object
    Dim result As Object, receipts As Object, receipt As Variant
    Dim row As Long, requestId As String, owner As String, status As String

    ' 集計側は保存してからCSVを移動するため、CSVの確認を先に行う。
    Set receipts = WorkReadStatus(outboxPath)
    data = ReadSummaryData()
    Set result = NewRequestMap()
    For row = 2 To UBound(data, 1)
        requestId = TextValue(data(row, SOURCE_ID))
        If Len(requestId) > 0 Then
            If result.Exists(requestId) Then
                Err.Raise WORK_ERROR, , "集計に重複したRequestIDがあります。" & vbCrLf & requestId
            End If
            owner = TextValue(data(row, SOURCE_COLS))
            Select Case owner
                Case COMPLETED
                    status = COMPLETED
                    owner = vbNullString
                Case UNASSIGNED, vbNullString
                    status = UNASSIGNED
                    owner = vbNullString
                Case Else
                    status = IN_PROGRESS
            End Select
            If status <> COMPLETED And receipts.Exists(requestId) Then
                receipt = receipts(requestId)
                status = receipt(0)
                owner = receipt(1)
            End If
            result.Add requestId, Array(status, owner)
        End If
    Next row
    Set ReadWorkSnapshot = result
End Function

Private Function ReadSummaryData() As Variant
    Dim files As Object, source As String, sourceToOpen As String, tempPath As String
    Dim book As Workbook, existing As Workbook, summary As Worksheet
    Dim lastRow As Long, col As Long, data As Variant, headers As Variant
    Dim prevEvents As Boolean, prevSecurity As Long, changedSettings As Boolean
    Dim createdTempFolder As Boolean
    Dim failure As String

    On Error GoTo Failed
    Set files = CreateObject("Scripting.FileSystemObject")
    source = FindSummaryFile(files)
    sourceToOpen = source
    For Each existing In Application.Workbooks
        If StrComp(existing.Name, files.GetFileName(source), vbTextCompare) = 0 Then
            ' 同名ブックが開いている場合は、ディスク上の保存済みの版を別名で読む。
            tempPath = files.BuildPath(files.GetSpecialFolder(2).Path, files.GetTempName())
            files.CreateFolder tempPath
            createdTempFolder = True
            sourceToOpen = files.BuildPath(tempPath, files.GetBaseName(tempPath) & "." & _
                                              files.GetExtensionName(source))
            files.CopyFile source, sourceToOpen, False
            Exit For
        End If
    Next existing

    prevEvents = Application.EnableEvents
    prevSecurity = Application.AutomationSecurity
    changedSettings = True
    Application.EnableEvents = False
    Application.AutomationSecurity = 3 ' msoAutomationSecurityForceDisable
    Set book = Application.Workbooks.Open(Filename:=sourceToOpen, UpdateLinks:=0, _
                    ReadOnly:=True, IgnoreReadOnlyRecommended:=True, AddToMru:=False, Notify:=False)
    On Error Resume Next
    Set summary = book.Worksheets(SUMMARY_SHEET)
    On Error GoTo Failed
    If summary Is Nothing Then Err.Raise WORK_ERROR, , "集計用Excelに「集計」シートがありません。"

    lastRow = summary.Cells(summary.Rows.Count, SOURCE_ID).End(xlUp).Row
    data = summary.Cells(1, 1).Resize(lastRow, SOURCE_COLS).Value2
    headers = Split(SOURCE_HEADERS, "|")
    For col = 1 To SOURCE_ID
        If TextValue(data(1, col)) <> headers(col - 1) Then
            Err.Raise WORK_ERROR, , "「集計」シートの見出しを確認してください。列 " & col
        End If
    Next col
    If TextValue(data(1, SOURCE_COLS)) <> STATE_HEADER Then
        Err.Raise WORK_ERROR, , "「集計」シートのV1に「対応状況」が必要です。"
    End If

    ReadSummaryData = data

CleanUp:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    If changedSettings Then
        Application.AutomationSecurity = prevSecurity
        Application.EnableEvents = prevEvents
    End If
    If createdTempFolder Then files.DeleteFolder tempPath, True
    On Error GoTo 0
    If Len(failure) > 0 Then Err.Raise WORK_ERROR, , failure
    Exit Function

Failed:
    failure = Err.Description
    Resume CleanUp
End Function

Private Function FindSummaryFile(ByVal files As Object) As String
    Dim extension As Variant, candidate As String

    For Each extension In Array(".xlsm", ".xlsx", ".xlsb", ".xls")
        candidate = files.BuildPath(ThisWorkbook.Path, SUMMARY_BOOK & extension)
        If files.FileExists(candidate) Then
            If Len(FindSummaryFile) > 0 Then
                Err.Raise WORK_ERROR, , "同名の「メンション集計用」ファイルが複数あります。"
            End If
            FindSummaryFile = candidate
        End If
    Next extension
    If Len(FindSummaryFile) = 0 Then
        Err.Raise WORK_ERROR, , "このExcelと同じ階層に「メンション集計用」が見つかりません。"
    End If
End Function

Private Function GetOutbox() As String
    Dim files As Object
    Set files = CreateObject("Scripting.FileSystemObject")
    If Len(ThisWorkbook.Path) = 0 Or InStr(1, ThisWorkbook.Path, "://", vbTextCompare) > 0 Then
        Err.Raise WORK_ERROR, , "Windowsの共有フォルダから作業用Excelを開いてください。"
    End If
    GetOutbox = files.BuildPath(ThisWorkbook.Path, OUTBOX)
    If Not files.FolderExists(GetOutbox) Then
        Err.Raise WORK_ERROR, , "このExcelと同じ階層に「書き込み用」フォルダがありません。"
    End If
End Function

Private Sub RenderWorkList(ByRef data As Variant, ByVal states As Object)
    Dim display() As Variant, item As Variant
    Dim row As Long, col As Long, rowCount As Long, previousRow As Long
    Dim requestId As String, cellText As String

    rowCount = UBound(data, 1) - 1
    If rowCount > 0 Then
        ReDim display(1 To rowCount, 1 To LAST_DATA_COL - DISPLAY_STATE + 1)
        For row = 2 To rowCount + 1
            requestId = TextValue(data(row, SOURCE_ID))
            If Len(requestId) > 0 Then
                item = states(requestId)
                display(row - 1, 1) = item(0)
            End If
            For col = 1 To SOURCE_ID
                cellText = TextValue(data(row, col))
                If Len(cellText) > 0 Then display(row - 1, col + 1) = "'" & cellText
            Next col
        Next row
    End If

    With WorkList()
        previousRow = .Cells(.Rows.Count, LAST_DATA_COL).End(xlUp).Row
        If previousRow > 1 Then
            .Range(.Cells(2, PICK_COL), .Cells(previousRow, LAST_DATA_COL)).ClearContents
        End If
        If rowCount > 0 Then
            .Cells(2, DISPLAY_STATE).Resize(rowCount, LAST_DATA_COL - DISPLAY_STATE + 1).Value2 = display
        End If
    End With
End Sub

Private Sub SetSelectionButtonCaption(ByVal caption As String)
    With WorkList().Buttons(2)
        If .Caption <> caption Then .Caption = caption
    End With
    ThisWorkbook.Saved = True
End Sub

Private Function PaintSelection() As Long
    Dim row As Long, requestId As String, values() As Variant
    If Not IsArray(mSourceRows) Then Exit Function
    If UBound(mSourceRows, 1) <= 1 Then Exit Function
    ReDim values(1 To UBound(mSourceRows, 1) - 1, 1 To 1)
    If mSelecting Then
        For row = 2 To UBound(mSourceRows, 1)
            requestId = TextValue(mSourceRows(row, SOURCE_ID))
            If Len(requestId) > 0 Then
                If Len(AvailableAction(mStates(requestId))) > 0 Then
                    values(row - 1, 1) = PICKABLE
                    PaintSelection = PaintSelection + 1
                End If
            End If
        Next row
    End If
    WorkList().Cells(2, PICK_COL).Resize(UBound(values, 1), 1).Value2 = values
    ThisWorkbook.Saved = True
End Function

Private Sub EndMentionSelection()
    Dim wasSelecting As Boolean

    wasSelecting = mSelecting
    mSelecting = False
    Set mPicked = Nothing
    PaintSelection
    If wasSelecting Then
        If Application.ActiveSheet Is WorkList() Then ActiveCell.Select
    End If
    SetSelectionButtonCaption "選択開始"
End Sub

Private Function AvailableAction(ByVal item As Variant) As String
    Select Case item(0)
        Case UNASSIGNED
            AvailableAction = IN_PROGRESS
        Case IN_PROGRESS
            If StrComp(item(1), mAccount, vbTextCompare) = 0 Then AvailableAction = COMPLETED
    End Select
End Function

Private Function NewRequestMap() As Object
    Set NewRequestMap = CreateObject("Scripting.Dictionary")
    NewRequestMap.CompareMode = vbTextCompare
End Function

Private Function TextValue(ByVal value As Variant) As String
    If IsError(value) Then Err.Raise WORK_ERROR, , "集計の依頼データにセルエラーがあります。"
    If Not IsEmpty(value) Then TextValue = CStr(value)
End Function

' CSV ファイルの入出力
Private Function WorkReadStatus(ByVal outbox As String) As Object
    Dim files As Object, statuses As Object

    Set files = OpenOutbox(outbox)
    Set statuses = NewRequestMap()
    CollectFolderStatuses files.GetFolder(outbox), statuses
    Set WorkReadStatus = statuses
End Function

Private Sub CollectFolderStatuses(ByVal folder As Object, ByVal statuses As Object, _
                                  Optional ByVal requestId As String = "")
    Dim candidate As Object, fields As Variant
    Dim fileName As String, state As String, suffix As String

    suffix = "_" & requestId & STATUS_EXTENSION
    For Each candidate In folder.Files
        fileName = CStr(candidate.Name)
        state = StatusFileState(fileName)
        If Len(state) > 0 Then
            If Len(requestId) = 0 Or InStr(1, fileName, requestId, vbTextCompare) > 0 Then
                ' 新旧どちらの名前でも、IDと担当者はCSV本文のA1/B1から照合する。
                fields = ReadStatusFields(CStr(candidate.Path), state)
                If IsArray(fields) Then
                    If Len(requestId) = 0 Or StrComp(fields(0), requestId, vbTextCompare) = 0 Then
                        AddReceiptStatus statuses, CStr(fields(0)), state, CStr(fields(1))
                    End If
                ElseIf Len(requestId) > 0 Then
                    ' 対象の名前と一致するが読めないCSVがある場合は、担当の確定を保留する。
                    If StrComp(Right$(fileName, Len(suffix)), suffix, vbTextCompare) = 0 Then
                        AddReceiptStatus statuses, requestId, state, vbNullString
                    End If
                End If
            End If
        End If
    Next candidate
End Sub

Private Sub AddReceiptStatus(ByVal statuses As Object, ByVal requestId As String, _
                             ByVal state As String, ByVal account As String)
    Dim previous As Variant

    If statuses.Exists(requestId) Then
        previous = statuses(requestId)
        If previous(0) = COMPLETED Then Exit Sub
        If state = IN_PROGRESS Then
            If StrComp(CStr(previous(1)), account, vbTextCompare) <> 0 Then account = vbNullString
        End If
    End If
    statuses(requestId) = Array(state, account)
End Sub

Private Function StatusFileState(ByVal fileName As String) As String
    If StrComp(Right$(fileName, Len(STATUS_EXTENSION)), STATUS_EXTENSION, vbTextCompare) <> 0 Then Exit Function
    If Left$(fileName, Len(IN_PROGRESS) + 1) = IN_PROGRESS & "_" Then StatusFileState = IN_PROGRESS
    If Left$(fileName, Len(COMPLETED) + 1) = COMPLETED & "_" Then StatusFileState = COMPLETED
End Function

Private Function StatusFileName(ByVal state As String, ByVal account As String, ByVal requestId As String) As String
    StatusFileName = state & "_" & account & "_" & requestId & STATUS_EXTENSION
End Function

Private Function WorkPublishStatus(ByVal outbox As String, ByVal requestId As String, _
                                  ByVal state As String, ByVal account As String) As Boolean
    Dim files As Object
    Dim fileName As String, destination As String, lockPath As String, tempPath As String, csvText As String
    Dim lockHandle As LongPtr, windowsError As Long, failure As Long, detail As String
    Dim previous As Variant

    lockHandle = -1
    On Error GoTo Failed
    Set files = OpenOutbox(outbox)
    If Not ValidFileNamePart(requestId) Then
        Err.Raise STORAGE_ERROR, , "ファイル名に使えないRequestIDです。"
    End If
    If Not ValidFileNamePart(account) Then
        Err.Raise STORAGE_ERROR, , "ファイル名に使えないPCのアカウント名です。"
    End If
    If state <> IN_PROGRESS And state <> COMPLETED Then
        Err.Raise STORAGE_ERROR, , "書き出す対応状況を確認してください。"
    End If

    fileName = StatusFileName(state, account, requestId)
    If Len(fileName) > 255 Then Err.Raise STORAGE_ERROR, , "書き出すファイル名が長すぎます。"
    destination = files.BuildPath(outbox, fileName)
    lockPath = files.BuildPath(outbox, "~_" & requestId & ".lock")
    If Len(destination) >= MAX_PATH Or Len(lockPath) >= MAX_PATH Or Len(outbox) > MAX_PATH - 14 Then
        Err.Raise STORAGE_ERROR, , "書き出し先のパスが長すぎます。"
    End If

    ' 同じ依頼の開始・完了を直列化する。Excel終了時はOSがハンドルを閉じる。
    lockHandle = CreateFileW(StrPtr(lockPath), GENERIC_WRITE Or DELETE_ACCESS, 0, _
                             0, OPEN_ALWAYS, FILE_ATTRIBUTE_HIDDEN Or FILE_FLAG_DELETE_ON_CLOSE, 0)
    windowsError = Err.LastDllError
    If lockHandle = -1 Then
        Select Case windowsError
            Case ERROR_SHARING_VIOLATION, ERROR_LOCK_VIOLATION, ERROR_DELETE_PENDING
                GoTo CleanUp
            Case ERROR_ACCESS_DENIED
                If files.FileExists(lockPath) Then GoTo CleanUp
        End Select
        RaiseWindowsError "依頼を書き出すためのロックを取得できません。", windowsError
    End If

    previous = ReadRequestReceipt(files, outbox, requestId)
    If IsArray(previous) Then
        If state = IN_PROGRESS Or previous(0) = COMPLETED Then GoTo CleanUp
        If StrComp(CStr(previous(1)), account, vbTextCompare) <> 0 Then GoTo CleanUp
    End If

    csvText = CsvField(requestId) & "," & CsvField(account)
    If state = COMPLETED Then csvText = csvText & "," & CsvField(COMPLETED)
    tempPath = ReserveTempFile(outbox)
    SaveCsv tempPath, csvText & vbCrLf

    ' 完成したCSVを同じフォルダ内で公開する。同名CSVは上書きしない。
    If MoveFileW(StrPtr(tempPath), StrPtr(destination)) = 0 Then
        windowsError = Err.LastDllError
        If windowsError = ERROR_FILE_EXISTS Or windowsError = ERROR_ALREADY_EXISTS Then GoTo CleanUp
        RaiseWindowsError "CSVを確定できません。", windowsError
    End If
    tempPath = ""
    WorkPublishStatus = True

CleanUp:
    On Error Resume Next
    If Len(tempPath) > 0 Then DeleteFileW StrPtr(tempPath)
    If lockHandle <> -1 Then CloseHandle lockHandle
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, "メンション作業用", detail
    Exit Function

Failed:
    failure = Err.Number
    detail = Err.Description
    Resume CleanUp
End Function

Private Function OpenOutbox(ByVal outbox As String) As Object
    Dim files As Object
    Set files = CreateObject("Scripting.FileSystemObject")
    If Not files.FolderExists(outbox) Then
        Err.Raise STORAGE_ERROR, , "「書き込み用」フォルダが見つかりません。" & vbCrLf & outbox
    End If
    Set OpenOutbox = files
End Function

Private Function ReadRequestReceipt(ByVal files As Object, ByVal outbox As String, _
                                    ByVal requestId As String) As Variant
    Dim statuses As Object, dayFolder As Object
    Dim archivePath As String, progressPath As String

    Set statuses = NewRequestMap()
    CollectFolderStatuses files.GetFolder(outbox), statuses, requestId

    ' 集計がCSVを移動した直後も、同じ依頼の受付記録を確認できるようにする。
    archivePath = files.BuildPath(outbox, "OLD")
    If files.FolderExists(archivePath) Then
        For Each dayFolder In files.GetFolder(archivePath).SubFolders
            If dayFolder.Name Like "########" Then
                CollectFolderStatuses dayFolder, statuses, requestId
                progressPath = files.BuildPath(dayFolder.Path, IN_PROGRESS)
                If files.FolderExists(progressPath) Then
                    CollectFolderStatuses files.GetFolder(progressPath), statuses, requestId
                End If
            End If
        Next dayFolder
    End If
    If statuses.Exists(requestId) Then ReadRequestReceipt = statuses(requestId)
End Function

Private Function ValidFileNamePart(ByVal value As String) As Boolean
    Dim position As Long
    If Len(value) = 0 Or HasControlCharacters(value) Then Exit Function
    For position = 1 To Len(value)
        If InStr(1, "<>:""/\|?*", Mid$(value, position, 1), vbBinaryCompare) > 0 Then Exit Function
    Next position
    ValidFileNamePart = True
End Function

Private Function HasControlCharacters(ByVal value As String) As Boolean
    Dim position As Long, code As Long
    For position = 1 To Len(value)
        code = AscW(Mid$(value, position, 1)) And &HFFFF&
        If code < 32 Then
            HasControlCharacters = True
            Exit Function
        End If
    Next position
End Function

Private Function ReserveTempFile(ByVal outbox As String) As String
    Dim buffer As String, prefix As String, windowsError As Long
    buffer = String$(MAX_PATH, vbNullChar)
    prefix = "mwc"
    If GetTempFileNameW(StrPtr(outbox), StrPtr(prefix), 0, StrPtr(buffer)) = 0 Then
        windowsError = Err.LastDllError
        RaiseWindowsError "CSVの一時ファイルを作成できません。", windowsError
    End If
    ReserveTempFile = Left$(buffer, InStr(buffer, vbNullChar) - 1)
End Function

Private Sub RaiseWindowsError(ByVal description As String, ByVal windowsError As Long)
    Err.Raise STORAGE_ERROR, , description & " (Windows " & windowsError & ")"
End Sub

Private Function CsvField(ByVal value As String) As String
    CsvField = QUOTE & Replace(value, QUOTE, QUOTE & QUOTE) & QUOTE
End Function

Private Function OpenTextStream() As Object
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "shift_jis"
    stream.Open
    Set OpenTextStream = stream
End Function

Private Sub SaveCsv(ByVal path As String, ByVal text As String)
    Dim stream As Object
    Set stream = OpenTextStream()
    stream.WriteText text
    stream.Position = 0
    If StrComp(stream.ReadText(-1), text, vbBinaryCompare) <> 0 Then
        Err.Raise STORAGE_ERROR, , "RequestIDまたはPCのアカウント名にShift_JISで保存できない文字が含まれています。"
    End If
    stream.Position = 0
    stream.SaveToFile path, 2 ' adSaveCreateOverWrite、自分の一時ファイル
    stream.Close
End Sub

Private Function ReadStatusFields(ByVal path As String, ByVal state As String) As Variant
    Dim stream As Object, csvText As String, fields As Variant
    Dim columnCount As Long
    On Error GoTo InvalidReceipt

    Set stream = OpenTextStream()
    stream.LoadFromFile path
    csvText = stream.ReadText(-1)
    stream.Close
    columnCount = 2
    If state = COMPLETED Then columnCount = 3
    fields = ReadRecord(csvText, columnCount)
    If Not IsArray(fields) Then Exit Function
    If Len(fields(0)) = 0 Or HasControlCharacters(fields(0)) Then Exit Function
    If Len(fields(1)) = 0 Or HasControlCharacters(fields(1)) Then Exit Function
    If state = COMPLETED Then
        If fields(2) <> COMPLETED Then Exit Function
    End If
    ReadStatusFields = fields
    Exit Function
InvalidReceipt:
    If Err.Number = 18 Then Err.Raise 18
End Function

Private Function ReadRecord(ByVal csvText As String, ByVal columnCount As Long) As Variant
    Dim fields() As String, position As Long, column As Long
    Dim size As Long, character As String, field As String

    If Right$(csvText, 2) = vbCrLf Then
        csvText = Left$(csvText, Len(csvText) - 2)
    ElseIf Right$(csvText, 1) = vbLf Or Right$(csvText, 1) = vbCr Then
        csvText = Left$(csvText, Len(csvText) - 1)
    End If
    If InStr(csvText, vbCr) > 0 Or InStr(csvText, vbLf) > 0 Then Exit Function

    size = Len(csvText)
    position = 1
    ReDim fields(0 To columnCount - 1)
    For column = LBound(fields) To UBound(fields)
        field = ""
        If Mid$(csvText, position, 1) = QUOTE Then
            position = position + 1
            Do
                If position > size Then Exit Function
                character = Mid$(csvText, position, 1)
                position = position + 1
                If character = QUOTE Then
                    If Mid$(csvText, position, 1) <> QUOTE Then Exit Do
                    position = position + 1
                End If
                field = field & character
            Loop
        Else
            Do While position <= size
                character = Mid$(csvText, position, 1)
                If character = "," Then Exit Do
                If character = QUOTE Then Exit Function
                field = field & character
                position = position + 1
            Loop
        End If
        fields(column) = field
        If column < UBound(fields) Then
            If Mid$(csvText, position, 1) <> "," Then Exit Function
            position = position + 1
        End If
    Next column
    If position > size Then ReadRecord = fields
End Function
