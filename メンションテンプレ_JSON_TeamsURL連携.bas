Attribute VB_Name = "modTemplateJsonTeamsUrl"
Option Explicit

' ============================================================
' メンションテンプレ JSON Teams / SharePoint URL連携
'
' 【構成】
'   ・元の「メンションテンプレ_JSON自動読込.bas」とは独立して動作
'   ・元モジュールのSub/Functionは呼び出さない
'   ・初回だけ元の「メンションテンプレ_JSON版を作成」を手動実行
'   ・以後はTeams版がJSON同期＋既存UserForm表示を担当
'
' 【前提】
'   Excelブックはローカル、またはOneDrive同期済みのローカルパスから開くこと。
'   ThisWorkbook.Path が https:// 形式の場合は同期できません。
' ============================================================

Private Const FORM_NAME As String = "frmTemplateSelectorJson"
Private Const CONFIG_NAME As String = "MentionTemplateJsonSourceUrl"
Private Const JSON_FILE_NAME As String = "テンプレート.json"
Private Const HTTP_TIMEOUT_MS As Long = 15000

Private mLastError As String
Private mLastStatus As Long
Private mLastTargetPath As String

' ============================================================
' 公開マクロ
' ============================================================
Public Sub メンションテンプレ_TeamsJSON版を初期設定()

    MsgBox _
        "初回だけ先に、Alt+F8から「メンションテンプレ_JSON版を作成」を実行してください。" & vbCrLf & vbCrLf & _
        "UserForm作成後、Teams / SharePoint上のJSON URLを設定します。", _
        vbInformation, _
        "Teams JSON 初期設定"

    TeamsJSON_URLを設定

End Sub

Public Sub メンションテンプレ_TeamsJSON版を開く()

    Dim sourceUrl As String

    sourceUrl = TeamsJsonSourceUrl()

    If Len(sourceUrl) > 0 Then
        If Not TeamsJsonSync(False) Then
            MsgBox _
                "Teams / SharePointから最新JSONを取得できませんでした。" & vbCrLf & vbCrLf & _
                mLastError & vbCrLf & vbCrLf & _
                "既存のローカルJSONがある場合は、その内容でフォームを開きます。", _
                vbExclamation, _
                "Teams JSON"
        End If
    End If

    TeamsJsonShowExistingForm

End Sub

Public Sub TeamsJSON_URLを設定()

    Dim currentUrl As String
    Dim inputUrl As String

    currentUrl = TeamsJsonSourceUrl()

    inputUrl = Trim$(InputBox( _
        "Teamsチャネルのテンプレート.jsonをSharePointで開き、" & vbCrLf & _
        "SharePoint側のファイルURLを貼り付けてください。" & vbCrLf & vbCrLf & _
        "URLが長い場合はセルへ貼り付けて、" & vbCrLf & _
        "TeamsJSON_URLを選択セルから設定 を使用してください。", _
        "Teams JSON URL設定", _
        currentUrl))

    If Len(inputUrl) = 0 Then Exit Sub

    If Len(inputUrl) >= 250 Then
        MsgBox _
            "URLが長いため、入力欄では途中で切れる可能性があります。" & vbCrLf & vbCrLf & _
            "URLをセルへ貼り付け、そのセルを選択してから" & vbCrLf & _
            "TeamsJSON_URLを選択セルから設定 を実行してください。", _
            vbExclamation, _
            "Teams JSON URL設定"
        Exit Sub
    End If

    If Not TeamsJsonValidateSourceUrl(inputUrl) Then Exit Sub

    TeamsJsonSaveSourceUrl inputUrl

    MsgBox _
        "JSONの参照URLを保存しました。", _
        vbInformation, _
        "Teams JSON URL設定"

End Sub

Public Sub TeamsJSON_URLを選択セルから設定()

    Dim inputUrl As String

    On Error GoTo CellError

    If TypeName(Selection) <> "Range" Then
        MsgBox _
            "SharePointのURLを入力したセルを選択してください。", _
            vbExclamation, _
            "Teams JSON URL設定"
        Exit Sub
    End If

    If Selection.Cells.CountLarge <> 1 Then
        MsgBox _
            "URLを入力したセルを1つだけ選択してください。", _
            vbExclamation, _
            "Teams JSON URL設定"
        Exit Sub
    End If

    inputUrl = Trim$(CStr(Selection.Value2))

    If Len(inputUrl) = 0 Then
        MsgBox _
            "選択セルにURLが入力されていません。", _
            vbExclamation, _
            "Teams JSON URL設定"
        Exit Sub
    End If

    If Not TeamsJsonValidateSourceUrl(inputUrl) Then Exit Sub

    TeamsJsonSaveSourceUrl inputUrl

    MsgBox _
        "選択セルからJSONの参照URLを保存しました。", _
        vbInformation, _
        "Teams JSON URL設定"

    Exit Sub

CellError:
    MsgBox _
        "選択セルからURLを設定できませんでした。" & vbCrLf & vbCrLf & _
        Err.Description, _
        vbExclamation, _
        "Teams JSON URL設定"

End Sub

Public Sub TeamsJSON_URL設定をクリア()

    On Error Resume Next
    ThisWorkbook.Names(CONFIG_NAME).Delete
    On Error GoTo 0

    mLastError = vbNullString
    mLastStatus = 0
    mLastTargetPath = vbNullString

    MsgBox _
        "Teams JSON URL設定をクリアしました。", _
        vbInformation, _
        "Teams JSON URL設定"

End Sub

Public Sub TeamsJSONを今すぐ同期()

    Call TeamsJsonSync(True)

End Sub

Public Sub TeamsJSON_接続状況を確認()

    Dim sourceUrl As String
    Dim statusText As String
    Dim targetPath As String
    Dim localJsonText As String

    sourceUrl = TeamsJsonSourceUrl()
    targetPath = TeamsJsonTargetPath(False)

    If mLastStatus > 0 Then
        statusText = CStr(mLastStatus)
    Else
        statusText = "未実行"
    End If

    If Len(targetPath) > 0 And TeamsJsonFileExists(targetPath) Then
        localJsonText = "あり"
    Else
        localJsonText = "なし"
    End If

    MsgBox _
        "【Teams JSON接続状況】" & vbCrLf & vbCrLf & _
        "設定URL: " & IIf(Len(sourceUrl) > 0, sourceUrl, "未設定") & vbCrLf & _
        "ローカルJSON: " & localJsonText & vbCrLf & _
        "保存先: " & IIf(Len(targetPath) > 0, targetPath, "取得できません") & vbCrLf & vbCrLf & _
        "直近HTTP状態: " & statusText & vbCrLf & _
        "直近エラー: " & IIf(Len(mLastError) > 0, mLastError, "なし"), _
        vbInformation, _
        "Teams JSON接続状況"

End Sub

' ============================================================
' UserForm表示
' ============================================================
Private Sub TeamsJsonShowExistingForm()

    On Error GoTo FormMissing

    VBA.UserForms.Add(FORM_NAME).Show
    Exit Sub

FormMissing:
    MsgBox _
        "JSON版UserFormがまだ作成されていません。" & vbCrLf & vbCrLf & _
        "Alt+F8から「メンションテンプレ_JSON版を作成」を一度だけ手動実行してください。" & vbCrLf & vbCrLf & _
        "その後は「メンションテンプレ_TeamsJSON版を開く」だけでOKです。", _
        vbExclamation, _
        "UserForm未作成"

End Sub

' ============================================================
' URL設定
' ============================================================
Public Function TeamsJsonSourceUrl() As String

    Dim nameObject As Name
    Dim refersText As String

    On Error Resume Next
    Set nameObject = ThisWorkbook.Names(CONFIG_NAME)
    On Error GoTo 0

    If nameObject Is Nothing Then Exit Function

    refersText = CStr(nameObject.RefersTo)

    If Left$(refersText, 1) = "=" Then
        refersText = Mid$(refersText, 2)
    End If

    If Len(refersText) >= 2 Then
        If Left$(refersText, 1) = Chr$(34) And Right$(refersText, 1) = Chr$(34) Then
            refersText = Mid$(refersText, 2, Len(refersText) - 2)
            refersText = Replace(refersText, Chr$(34) & Chr$(34), Chr$(34))
        End If
    End If

    TeamsJsonSourceUrl = Trim$(refersText)

End Function

Private Sub TeamsJsonSaveSourceUrl(ByVal sourceUrl As String)

    Dim refersText As String

    sourceUrl = Trim$(sourceUrl)

    refersText = "=" & Chr$(34) & _
                 Replace(sourceUrl, Chr$(34), Chr$(34) & Chr$(34)) & _
                 Chr$(34)

    On Error Resume Next
    ThisWorkbook.Names(CONFIG_NAME).Delete
    On Error GoTo SaveError

    ThisWorkbook.Names.Add _
        Name:=CONFIG_NAME, _
        RefersTo:=refersText, _
        Visible:=False

    Exit Sub

SaveError:
    Err.Raise _
        vbObjectError + 2400, _
        , _
        "JSON URL設定をブックへ保存できませんでした。" & vbCrLf & Err.Description

End Sub

Private Function TeamsJsonValidateSourceUrl(ByVal sourceUrl As String) As Boolean

    sourceUrl = Trim$(sourceUrl)

    If InStr(1, sourceUrl, "teams.microsoft.com", vbTextCompare) > 0 Then
        MsgBox _
            "Teamsアプリの画面URLではなく、対象ファイルをSharePointで開き、" & vbCrLf & _
            "SharePoint側のファイルURLを設定してください。", _
            vbExclamation, _
            "Teams JSON URL設定"
        Exit Function
    End If

    If InStr(1, sourceUrl, "http://", vbTextCompare) <> 1 _
       And InStr(1, sourceUrl, "https://", vbTextCompare) <> 1 Then

        MsgBox _
            "http:// または https:// で始まるURLを設定してください。", _
            vbExclamation, _
            "Teams JSON URL設定"

        Exit Function
    End If

    TeamsJsonValidateSourceUrl = True

End Function

' ============================================================
' 同期
' ============================================================
Public Function TeamsJsonSync( _
    Optional ByVal showResult As Boolean = False) As Boolean

    Dim sourceUrl As String
    Dim requestUrl As String
    Dim targetPath As String
    Dim tempPath As String
    Dim jsonText As String

    On Error GoTo SyncError

    mLastError = vbNullString
    mLastStatus = 0
    mLastTargetPath = vbNullString

    sourceUrl = TeamsJsonSourceUrl()

    If Len(sourceUrl) = 0 Then
        Err.Raise _
            vbObjectError + 2401, _
            , _
            "Teams JSON URLが未設定です。"
    End If

    targetPath = TeamsJsonTargetPath(True)
    mLastTargetPath = targetPath

    requestUrl = TeamsJsonNormalizeDownloadUrl(sourceUrl)

    jsonText = TeamsJsonDownloadUtf8(requestUrl)
    TeamsJsonValidateDownloadedText jsonText

    tempPath = targetPath & ".download.tmp"

    TeamsJsonDeleteIfExists tempPath
    TeamsJsonWriteUtf8Text tempPath, jsonText
    TeamsJsonReplaceFileSafely tempPath, targetPath

    TeamsJsonSync = True

    If showResult Then
        MsgBox _
            "Teams / SharePointからJSONを同期しました。" & vbCrLf & vbCrLf & _
            "HTTP: " & CStr(mLastStatus) & vbCrLf & _
            "保存先: " & targetPath, _
            vbInformation, _
            "Teams JSON同期"
    End If

    Exit Function

SyncError:
    mLastError = Err.Description

    TeamsJsonDeleteIfExists tempPath
    TeamsJsonSync = False

    If showResult Then
        MsgBox _
            "Teams / SharePointからJSONを同期できませんでした。" & vbCrLf & vbCrLf & _
            mLastError & vbCrLf & vbCrLf & _
            TeamsJsonAuthenticationAdvice(), _
            vbExclamation, _
            "Teams JSON同期"
    End If

End Function

Private Function TeamsJsonTargetPath( _
    ByVal raiseOnError As Boolean) As String

    Dim bookFolder As String

    bookFolder = Trim$(ThisWorkbook.Path)

    If Len(bookFolder) = 0 Then
        If raiseOnError Then
            Err.Raise _
                vbObjectError + 2440, _
                , _
                "先にExcelブックを*.xlsmとして保存してください。"
        End If
        Exit Function
    End If

    If TeamsJsonIsWebPath(bookFolder) Then
        If raiseOnError Then
            Err.Raise _
                vbObjectError + 2441, _
                , _
                "ExcelブックがSharePointのURL形式で開かれています。" & vbCrLf & _
                "OneDriveで同期したローカル側のExcelを開いてから実行してください。"
        End If
        Exit Function
    End If

    If Right$(bookFolder, 1) = Application.PathSeparator Then
        TeamsJsonTargetPath = bookFolder & JSON_FILE_NAME
    Else
        TeamsJsonTargetPath = bookFolder & Application.PathSeparator & JSON_FILE_NAME
    End If

End Function

Private Function TeamsJsonIsWebPath(ByVal valueText As String) As Boolean

    Dim lowerText As String

    lowerText = LCase$(Trim$(valueText))

    TeamsJsonIsWebPath = _
        (Left$(lowerText, 7) = "http://" Or _
         Left$(lowerText, 8) = "https://")

End Function

' ============================================================
' HTTP取得
' ============================================================
Private Function TeamsJsonDownloadUtf8(ByVal requestUrl As String) As String

    Dim httpObject As Object
    Dim responseBody As Variant
    Dim firstError As String

    On Error GoTo XmlHttpError

    Set httpObject = CreateObject("MSXML2.XMLHTTP.6.0")

    With httpObject
        .Open "GET", requestUrl, False
        .setRequestHeader "Accept", "application/json,text/plain,*/*"
        .setRequestHeader "Cache-Control", "no-cache"
        .send

        mLastStatus = CLng(.Status)

        If mLastStatus < 200 Or mLastStatus >= 300 Then
            Err.Raise _
                vbObjectError + 2410, _
                , _
                "HTTP " & CStr(mLastStatus) & " が返されました。"
        End If

        responseBody = .responseBody
    End With

    TeamsJsonDownloadUtf8 = TeamsJsonBytesToUtf8(responseBody)
    Exit Function

XmlHttpError:
    firstError = Err.Description
    Err.Clear
    Set httpObject = Nothing

    On Error GoTo WinHttpError

    Set httpObject = CreateObject("WinHttp.WinHttpRequest.5.1")

    With httpObject
        .Open "GET", requestUrl, False
        .SetTimeouts _
            HTTP_TIMEOUT_MS, _
            HTTP_TIMEOUT_MS, _
            HTTP_TIMEOUT_MS, _
            HTTP_TIMEOUT_MS
        .SetAutoLogonPolicy 0
        .SetRequestHeader "Accept", "application/json,text/plain,*/*"
        .SetRequestHeader "Cache-Control", "no-cache"
        .Send

        mLastStatus = CLng(.Status)

        If mLastStatus < 200 Or mLastStatus >= 300 Then
            Err.Raise _
                vbObjectError + 2411, _
                , _
                "HTTP " & CStr(mLastStatus) & " が返されました。"
        End If

        responseBody = .ResponseBody
    End With

    TeamsJsonDownloadUtf8 = TeamsJsonBytesToUtf8(responseBody)
    Exit Function

WinHttpError:
    If mLastStatus = 401 Or mLastStatus = 403 Then
        Err.Raise _
            vbObjectError + 2412, _
            , _
            "SharePoint側でMicrosoft 365認証が必要です（HTTP " & _
            CStr(mLastStatus) & "）。"
    End If

    Err.Raise _
        vbObjectError + 2413, _
        , _
        "URLからJSONを取得できませんでした。" & vbCrLf & _
        "XMLHTTP: " & firstError & vbCrLf & _
        "WinHTTP: " & Err.Description

End Function

Private Function TeamsJsonBytesToUtf8( _
    ByVal responseBody As Variant) As String

    Dim streamObject As Object
    Dim loadedText As String

    Set streamObject = CreateObject("ADODB.Stream")

    With streamObject
        .Type = 1
        .Open
        .Write responseBody
        .Position = 0
        .Type = 2
        .Charset = "utf-8"
        loadedText = .ReadText(-1)
        .Close
    End With

    If Len(loadedText) > 0 Then
        If AscW(Left$(loadedText, 1)) = &HFEFF Then
            loadedText = Mid$(loadedText, 2)
        End If
    End If

    TeamsJsonBytesToUtf8 = loadedText

End Function

Private Function TeamsJsonNormalizeDownloadUrl( _
    ByVal sourceUrl As String) As String

    Dim normalizedUrl As String

    normalizedUrl = Trim$(sourceUrl)

    normalizedUrl = Replace( _
        normalizedUrl, _
        "&amp;", _
        "&", _
        1, _
        -1, _
        vbTextCompare)

    normalizedUrl = Replace( _
        normalizedUrl, _
        "?web=1", _
        "?download=1", _
        1, _
        -1, _
        vbTextCompare)

    normalizedUrl = Replace( _
        normalizedUrl, _
        "&web=1", _
        "&download=1", _
        1, _
        -1, _
        vbTextCompare)

    If InStr(1, normalizedUrl, "download=1", vbTextCompare) = 0 Then
        If InStr(1, normalizedUrl, "?", vbBinaryCompare) > 0 Then
            normalizedUrl = normalizedUrl & "&download=1"
        Else
            normalizedUrl = normalizedUrl & "?download=1"
        End If
    End If

    TeamsJsonNormalizeDownloadUrl = normalizedUrl

End Function

Private Sub TeamsJsonValidateDownloadedText(ByVal jsonText As String)

    Dim trimmedText As String
    Dim lowerText As String

    trimmedText = Trim$(jsonText)

    If Len(trimmedText) = 0 Then
        Err.Raise _
            vbObjectError + 2420, _
            , _
            "取得したファイルが空です。"
    End If

    lowerText = LCase$(Left$(trimmedText, 500))

    If Left$(trimmedText, 1) = "<" _
       Or InStr(1, lowerText, "<!doctype", vbTextCompare) > 0 _
       Or InStr(1, lowerText, "<html", vbTextCompare) > 0 _
       Or InStr(1, lowerText, "login.microsoftonline.com", vbTextCompare) > 0 Then

        Err.Raise _
            vbObjectError + 2421, _
            , _
            "JSONではなくMicrosoft 365のサインイン画面またはWebページが返されました。"
    End If

    If Left$(trimmedText, 1) <> "[" _
       Or Right$(trimmedText, 1) <> "]" Then

        Err.Raise _
            vbObjectError + 2422, _
            , _
            "取得内容がテンプレートJSONの配列形式ではありません。"
    End If

End Sub

' ============================================================
' ファイル保存
' ============================================================
Private Sub TeamsJsonWriteUtf8Text( _
    ByVal filePath As String, _
    ByVal sourceText As String)

    Dim streamObject As Object

    Set streamObject = CreateObject("ADODB.Stream")

    With streamObject
        .Type = 2
        .Charset = "utf-8"
        .Open
        .WriteText sourceText
        .SaveToFile filePath, 2
        .Close
    End With

End Sub

Private Sub TeamsJsonReplaceFileSafely( _
    ByVal tempPath As String, _
    ByVal targetPath As String)

    Dim backupPath As String

    backupPath = targetPath & ".bak"

    On Error GoTo ReplaceError

    If TeamsJsonFileExists(targetPath) Then
        TeamsJsonDeleteIfExists backupPath
        FileCopy targetPath, backupPath
    End If

    TeamsJsonDeleteIfExists targetPath
    Name tempPath As targetPath

    Exit Sub

ReplaceError:
    On Error Resume Next

    If Not TeamsJsonFileExists(targetPath) _
       And TeamsJsonFileExists(backupPath) Then
        FileCopy backupPath, targetPath
    End If

    On Error GoTo 0

    Err.Raise _
        vbObjectError + 2430, _
        , _
        "取得したJSONを保存先へ反映できませんでした。" & vbCrLf & _
        targetPath

End Sub

Private Function TeamsJsonFileExists(ByVal filePath As String) As Boolean

    If Len(Trim$(filePath)) = 0 Then Exit Function

    On Error Resume Next

    TeamsJsonFileExists = _
        (Len(Dir$(filePath, _
            vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0)

    On Error GoTo 0

End Function

Private Sub TeamsJsonDeleteIfExists(ByVal filePath As String)

    If Len(Trim$(filePath)) = 0 Then Exit Sub

    On Error Resume Next

    If TeamsJsonFileExists(filePath) Then
        Kill filePath
    End If

    On Error GoTo 0

End Sub

Private Function TeamsJsonAuthenticationAdvice() As String

    If mLastStatus = 401 _
       Or mLastStatus = 403 _
       Or InStr(1, mLastError, "サインイン", vbTextCompare) > 0 _
       Or InStr(1, mLastError, "認証", vbTextCompare) > 0 Then

        TeamsJsonAuthenticationAdvice = _
            "Microsoft 365認証が必要なリンクでは、VBA単体のURL取得が失敗する場合があります。" & vbCrLf & _
            "その場合はTeamsのチャネルフォルダーをOneDriveで同期し、従来のローカル読込方式を使用してください。"
    Else
        TeamsJsonAuthenticationAdvice = _
            "URL、ネットワーク接続、SharePoint側のアクセス権を確認してください。"
    End If

End Function
