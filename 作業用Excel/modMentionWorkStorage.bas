Attribute VB_Name = "modMentionWorkStorage"
Option Explicit

Private Const WORK_ERROR As Long = vbObjectError + 1600
Private Const IN_PROGRESS As String = "対応中"
Private Const COMPLETED As String = "完了"
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

Public Function WorkReadStatus(ByVal outbox As String) As Object
    Dim files As Object, statuses As Object, candidate As Object
    Dim fileName As String, state As String, requestId As String
    Dim key As Variant, receipt As Variant, owner As String

    Set files = OpenOutbox(outbox)
    Set statuses = CreateObject("Scripting.Dictionary")
    statuses.CompareMode = vbTextCompare

    For Each candidate In files.GetFolder(outbox).Files
        fileName = CStr(candidate.Name)
        If StrComp(files.GetExtensionName(fileName), "csv", vbTextCompare) = 0 Then
            state = ""
            If Left$(fileName, Len(IN_PROGRESS) + 1) = IN_PROGRESS & "_" Then state = IN_PROGRESS
            If Left$(fileName, Len(COMPLETED) + 1) = COMPLETED & "_" Then state = COMPLETED
            If Len(state) > 0 And Len(fileName) > Len(state) + 1 + Len(STATUS_EXTENSION) Then
                requestId = Mid$(fileName, Len(state) + 2, Len(fileName) - Len(state) - 1 - Len(STATUS_EXTENSION))
                If state = COMPLETED Or Not statuses.Exists(requestId) Then
                    statuses(requestId) = Array(state, CStr(candidate.Path))
                End If
            End If
        End If
    Next candidate

    ' 完了済みの対応中CSVは開かず、担当中のものだけ担当者を照合する。
    For Each key In statuses.Keys
        receipt = statuses(key)
        owner = ""
        If receipt(0) = IN_PROGRESS Then owner = ReadReceiptOwner(CStr(receipt(1)), CStr(key), IN_PROGRESS)
        statuses(key) = Array(receipt(0), owner)
    Next key
    Set WorkReadStatus = statuses
End Function

Public Function WorkReadOneStatus(ByVal outbox As String, ByVal requestId As String) As Variant
    Dim files As Object
    If Not ValidRequestId(requestId) Then Exit Function
    Set files = OpenOutbox(outbox)
    WorkReadOneStatus = ReadOneStatus(files, outbox, requestId)
End Function

Public Function WorkPublishStatus(ByVal outbox As String, ByVal requestId As String, _
                                  ByVal state As String, ByVal account As String) As Boolean
    Dim files As Object
    Dim destination As String, lockPath As String, tempPath As String, csvText As String
    Dim lockHandle As LongPtr, windowsError As Long, failure As Long, detail As String
    Dim previous As Variant

    lockHandle = -1
    On Error GoTo Failed
    Set files = OpenOutbox(outbox)
    If Not ValidRequestId(requestId) Then
        Err.Raise WORK_ERROR, , "ファイル名に使えないRequestIDです。"
    End If
    If Len(account) = 0 Or HasControlCharacters(account) Then
        Err.Raise WORK_ERROR, , "PCのアカウント名を取得できません。"
    End If
    If state <> IN_PROGRESS And state <> COMPLETED Then
        Err.Raise WORK_ERROR, , "書き出す対応状況を確認してください。"
    End If

    destination = StatusPath(files, outbox, requestId, state)
    lockPath = files.BuildPath(outbox, "~_" & requestId & ".lock")
    If Len(destination) >= MAX_PATH Or Len(lockPath) >= MAX_PATH Or Len(outbox) > MAX_PATH - 14 Then
        Err.Raise WORK_ERROR, , "書き出し先のパスが長すぎます。"
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

    previous = ReadOneStatus(files, outbox, requestId)
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
        Err.Raise WORK_ERROR, , "「書き出し用」フォルダが見つかりません。" & vbCrLf & outbox
    End If
    Set OpenOutbox = files
End Function

Private Function StatusPath(ByVal files As Object, ByVal outbox As String, _
                            ByVal requestId As String, ByVal state As String) As String
    StatusPath = files.BuildPath(outbox, state & "_" & requestId & STATUS_EXTENSION)
End Function

Private Function ReadOneStatus(ByVal files As Object, ByVal outbox As String, _
                               ByVal requestId As String) As Variant
    Dim path As String, state As String
    state = COMPLETED
    path = StatusPath(files, outbox, requestId, state)
    If Not files.FileExists(path) Then
        state = IN_PROGRESS
        path = StatusPath(files, outbox, requestId, state)
        If Not files.FileExists(path) Then Exit Function
        ' 対応中を確認した直後に完了が公開されていれば、完了を優先する。
        If files.FileExists(StatusPath(files, outbox, requestId, COMPLETED)) Then
            state = COMPLETED
            path = StatusPath(files, outbox, requestId, state)
        End If
    End If
    ReadOneStatus = Array(state, ReadReceiptOwner(path, requestId, state))
End Function

Private Function ValidRequestId(ByVal requestId As String) As Boolean
    Dim position As Long
    If Len(requestId) = 0 Or Len(requestId) > 255 - Len("対応中_.csv") Then Exit Function
    If HasControlCharacters(requestId) Then Exit Function
    For position = 1 To Len(requestId)
        If InStr(1, "<>:""/\|?*", Mid$(requestId, position, 1), vbBinaryCompare) > 0 Then Exit Function
    Next position
    ValidRequestId = True
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
    Err.Raise WORK_ERROR, , description & " (Windows " & windowsError & ")"
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
        Err.Raise WORK_ERROR, , "RequestIDまたはPCのアカウント名にShift_JISで保存できない文字が含まれています。"
    End If
    stream.Position = 0
    stream.SaveToFile path, 2 ' adSaveCreateOverWrite、自分の一時ファイル
    stream.Close
End Sub

Private Function ReadReceiptOwner(ByVal path As String, ByVal requestId As String, _
                                  ByVal state As String) As String
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
    If StrComp(fields(0), requestId, vbBinaryCompare) <> 0 Then Exit Function
    If Len(fields(1)) = 0 Or HasControlCharacters(fields(1)) Then Exit Function
    If state = COMPLETED Then
        If fields(2) <> COMPLETED Then Exit Function
    End If
    ReadReceiptOwner = fields(1)
InvalidReceipt:
    ' 読めないCSVも確保済みとして扱う。担当者の判定は保留する。
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
