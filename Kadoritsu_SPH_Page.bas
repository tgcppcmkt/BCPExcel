Attribute VB_Name = "Kadoritsu_SPH_Page"
Option Explicit

' 稼働率とSPHの差し込み用ページをPowerPointで作成
' 実行場所：Excel / PowerPoint どちらでも可
' 実行マクロ：CreateKadoritsuSPHDeck / AddKadoritsuSPHPageToActivePresentation

Private Const FONT_JP As String = "Meiryo"
Private Const ppLayoutBlank As Long = 12
Private Const ppAlignLeft As Long = 1
Private Const ppAlignCenter As Long = 2
Private Const ppAlignRight As Long = 3
Private Const ppAutoSizeNone As Long = 0

Public Sub CreateKadoritsuSPHDeck()
    Dim ppApp As Object, pres As Object, sld As Object
    Set ppApp = GetPowerPointApp()
    If ppApp Is Nothing Then Exit Sub

    Set pres = ppApp.Presentations.Add
    With pres.PageSetup
        .SlideWidth = PtIn(13.333333)
        .SlideHeight = PtIn(7.5)
    End With

    Do While pres.Slides.Count > 0
        pres.Slides(1).Delete
    Loop

    Set sld = pres.Slides.Add(1, ppLayoutBlank)
    BuildKadoritsuSPHSlide sld, 1
    pres.Slides(1).Select
    MsgBox "稼働率とSPHの差し込み用ページを作成しました。", vbInformation
End Sub

Public Sub AddKadoritsuSPHPageToActivePresentation()
    Dim ppApp As Object, pres As Object, sld As Object, pageNo As Long
    Set ppApp = GetPowerPointApp()
    If ppApp Is Nothing Then Exit Sub

    If ppApp.Presentations.Count = 0 Then
        Set pres = ppApp.Presentations.Add
        With pres.PageSetup
            .SlideWidth = PtIn(13.333333)
            .SlideHeight = PtIn(7.5)
        End With
    Else
        Set pres = ppApp.ActivePresentation
    End If

    pageNo = pres.Slides.Count + 1
    Set sld = pres.Slides.Add(pageNo, ppLayoutBlank)
    BuildKadoritsuSPHSlide sld, pageNo
    sld.Select
    MsgBox "稼働率とSPHのページを追加しました。", vbInformation
End Sub

Private Function GetPowerPointApp() As Object
    Dim ppApp As Object
    On Error Resume Next
    Set ppApp = GetObject(, "PowerPoint.Application")
    On Error GoTo 0

    If ppApp Is Nothing Then
        On Error Resume Next
        Set ppApp = CreateObject("PowerPoint.Application")
        On Error GoTo 0
    End If

    If ppApp Is Nothing Then
        MsgBox "PowerPointを起動できませんでした。" & vbCrLf & _
               "Microsoft PowerPointがインストールされているか確認してください。", vbExclamation
        Set GetPowerPointApp = Nothing
        Exit Function
    End If

    ppApp.Visible = True
    Set GetPowerPointApp = ppApp
End Function

Private Sub BuildKadoritsuSPHSlide(ByVal sld As Object, ByVal slideNo As Long)
    ApplyBase sld, slideNo

    AddText sld, 0.55, 0.3, 12, 0.52, "稼働率とSPHの見方", 22, True, ColWhite, ppAlignLeft
    AddText sld, 0.58, 0.88, 11.83, 0.31, "稼働率は、Aチーム時間の中で「実際に処理へ使えた時間」を見る指標", 13.5, False, ColSub, ppAlignLeft

    AddRoundBox sld, 0.75, 1.34, 7.45, 1.32, ColPanel, ColCyan, 1.4
    AddText sld, 0.95, 1.5, 7.05, 0.44, "稼働率 ＝ Aチームとしての処理時間合計 ÷ Aチームとしてらぼろぐを選択していた時間", 15.2, True, ColWhite, ppAlignCenter
    AddText sld, 0.95, 2.05, 7.05, 0.33, "※処理時間は「平均処理時間 × 処理件数」で暫定換算する", 11.8, False, ColSub, ppAlignCenter

    AddRoundBox sld, 8.45, 1.34, 4.05, 1.32, ColGreen, ColGreenBright, 1.4
    AddText sld, 8.68, 1.47, 3.6, 0.34, "目安の稼働率：83％", 16, True, ColWhite, ppAlignCenter
    AddText sld, 8.68, 1.91, 3.6, 0.25, "平均処理時間 2.5分 × 20件 ＝ 約50分", 11.8, True, ColWhite, ppAlignCenter
    AddText sld, 8.68, 2.3, 3.6, 0.2, "50分 ÷ 60分 ≒ 83％", 10.8, False, ColWhite, ppAlignCenter

    AddRoundBox sld, 0.75, 2.92, 11.75, 0.42, ColGrayBox, ColBorder, 1.4
    AddText sld, 0.92, 2.98, 11.4, 0.28, "SPH × 稼働率の4象限で、サポート投入・解除・育成対象を見極める", 13.3, True, ColWhite, ppAlignCenter

    AddRoundBox sld, 0.75, 3.54, 0.88, 1.28, ColStepBlue, ColCyan, 1.4
    AddText sld, 0.83, 3.82, 0.72, 0.72, "SPH" & vbCrLf & "高", 14.5, True, ColWhite, ppAlignCenter
    AddRoundBox sld, 0.75, 5.02, 0.88, 1.28, ColGrayBox, ColBorder, 1.4
    AddText sld, 0.83, 5.3, 0.72, 0.72, "SPH" & vbCrLf & "低", 14.5, True, ColWhite, ppAlignCenter

    AddRoundBox sld, 1.83, 3.37, 5.08, 0.32, ColStepBlue, ColCyan, 1.4
    AddText sld, 1.98, 3.4, 4.78, 0.22, "稼働率：高", 12.5, True, ColWhite, ppAlignLeft
    AddRoundBox sld, 7.22, 3.37, 5.08, 0.32, ColGrayBox, ColBorder, 1.4
    AddText sld, 7.37, 3.4, 4.78, 0.22, "稼働率：低", 12.5, True, ColWhite, ppAlignLeft

    AddRoundBox sld, 1.83, 3.78, 5.08, 1.18, ColGreen, ColGreenBright, 1.4
    AddText sld, 1.99, 3.9, 4.76, 0.38, "理想的な形", 15.2, True, ColWhite, ppAlignLeft
    AddText sld, 2.05, 4.36, 4.64, 0.5, "・処理速度も、実処理時間も高い" & vbCrLf & "・SLA達成に向きやすい" & vbCrLf & "・ただし逼迫状態に注意", 11.1, False, ColWhite, ppAlignLeft

    AddRoundBox sld, 7.22, 3.78, 5.08, 1.18, ColBrown, ColOrange, 1.4
    AddText sld, 7.38, 3.9, 4.76, 0.38, "切替忘れに注意", 15.2, True, ColWhite, ppAlignLeft
    AddText sld, 7.44, 4.36, 4.64, 0.5, "・処理件数は多いが、手隙時間も長い" & vbCrLf & "・らぼろぐ切替忘れに注意" & vbCrLf & "・切替済ならSPHはもっと高い", 10.6, False, ColWhite, ppAlignLeft

    AddRoundBox sld, 1.83, 5.25, 5.08, 1.18, ColStepBlue, ColCyan, 1.4
    AddText sld, 1.99, 5.37, 4.76, 0.38, "サポート対象の見極め", 14.7, True, ColWhite, ppAlignLeft
    AddText sld, 2.05, 5.83, 4.64, 0.5, "・らぼろぐ切替はできている" & vbCrLf & "・ただしSPHが低く、サポート効果は薄い可能性" & vbCrLf & "・依頼対象にするか個別判断", 10.5, False, ColWhite, ppAlignLeft

    AddRoundBox sld, 7.22, 5.25, 5.08, 1.18, ColRed, ColRed, 1.4
    AddText sld, 7.38, 5.37, 4.76, 0.38, "業務姿勢の改善対象", 14.7, True, ColWhite, ppAlignLeft
    AddText sld, 7.44, 5.83, 4.64, 0.5, "・サポート効果が薄く、切替忘れもある" & vbCrLf & "・サポート依頼ではなく、日常の業務姿勢を確認" & vbCrLf & "・記録・手隙時行動を改善", 10.1, False, ColWhite, ppAlignLeft

    AddRoundBox sld, 0.75, 6.7, 11.75, 0.42, ColPanel, ColOrange, 1.4
    AddText sld, 0.92, 6.77, 11.4, 0.27, "SPHだけで判断せず、稼働率・SLA・品質・残業時間をセットで確認する", 12.2, True, ColWhite, ppAlignCenter
End Sub

Private Sub ApplyBase(ByVal sld As Object, ByVal slideNo As Long)
    sld.FollowMasterBackground = False
    With sld.Background.Fill
        .Visible = True
        .Solid
        .ForeColor.RGB = ColBg
    End With

    AddRect sld, 0, 0, 13.333333, 0.05, ColCyan, ColCyan, 0
    AddRect sld, 0, 7.43, 13.333333, 0.07, ColOrange, ColOrange, 0
    AddText sld, 12.38, 7.05, 0.31, 0.24, CStr(slideNo), 8.25, False, ColSub, ppAlignRight
End Sub

Private Function AddText(ByVal sld As Object, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal txt As String, ByVal fontSize As Double, ByVal isBold As Boolean, ByVal fontColor As Long, ByVal align As Long) As Object
    Dim sh As Object
    Set sh = sld.Shapes.AddTextbox(1, PtIn(x), PtIn(y), PtIn(w), PtIn(h))
    With sh
        .Fill.Visible = False
        .Line.Visible = False
        With .TextFrame
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .WordWrap = True
            .AutoSize = ppAutoSizeNone
            .VerticalAnchor = 3
            .TextRange.Text = txt
            .TextRange.ParagraphFormat.Alignment = align
            With .TextRange.Font
                .Name = FONT_JP
                On Error Resume Next
                .NameFarEast = FONT_JP
                On Error GoTo 0
                .Size = fontSize
                .Bold = isBold
                .Color.RGB = fontColor
            End With
        End With
    End With
    Set AddText = sh
End Function

Private Function AddRect(ByVal sld As Object, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal fillColor As Long, ByVal lineColor As Long, ByVal lineWeight As Double) As Object
    Dim sh As Object
    Set sh = sld.Shapes.AddShape(1, PtIn(x), PtIn(y), PtIn(w), PtIn(h))
    With sh
        .Fill.Solid
        .Fill.ForeColor.RGB = fillColor
        If lineWeight <= 0 Then
            .Line.Visible = False
        Else
            .Line.Visible = True
            .Line.ForeColor.RGB = lineColor
            .Line.Weight = lineWeight
        End If
    End With
    Set AddRect = sh
End Function

Private Function AddRoundBox(ByVal sld As Object, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal fillColor As Long, ByVal lineColor As Long, ByVal lineWeight As Double) As Object
    Dim sh As Object
    Set sh = sld.Shapes.AddShape(5, PtIn(x), PtIn(y), PtIn(w), PtIn(h))
    With sh
        .Fill.Solid
        .Fill.ForeColor.RGB = fillColor
        If lineWeight <= 0 Then
            .Line.Visible = False
        Else
            .Line.Visible = True
            .Line.ForeColor.RGB = lineColor
            .Line.Weight = lineWeight
        End If
    End With
    Set AddRoundBox = sh
End Function

Private Function PtIn(ByVal inches As Double) As Single
    PtIn = CSng(inches * 72)
End Function

Private Function ColBg() As Long
    ColBg = RGB(13, 23, 40)
End Function
Private Function ColPanel() As Long
    ColPanel = RGB(17, 29, 49)
End Function
Private Function ColCyan() As Long
    ColCyan = RGB(51, 181, 229)
End Function
Private Function ColOrange() As Long
    ColOrange = RGB(243, 155, 23)
End Function
Private Function ColWhite() As Long
    ColWhite = RGB(255, 255, 255)
End Function
Private Function ColSub() As Long
    ColSub = RGB(182, 196, 216)
End Function
Private Function ColStepBlue() As Long
    ColStepBlue = RGB(14, 99, 139)
End Function
Private Function ColGrayBox() As Long
    ColGrayBox = RGB(42, 50, 67)
End Function
Private Function ColBorder() As Long
    ColBorder = RGB(51, 65, 88)
End Function
Private Function ColGreen() As Long
    ColGreen = RGB(18, 107, 53)
End Function
Private Function ColGreenBright() As Long
    ColGreenBright = RGB(34, 197, 94)
End Function
Private Function ColBrown() As Long
    ColBrown = RGB(163, 71, 11)
End Function
Private Function ColRed() As Long
    ColRed = RGB(179, 40, 40)
End Function
