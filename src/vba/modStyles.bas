Attribute VB_Name = "modStyles"
Option Explicit

' =====================================================================
' modStyles
' Installs (or refreshes) the OI-* named styles in the active document,
' so downstream code can assign paragraphs to canonical styles.
' =====================================================================

Public Sub InstallOrRefresh(ByVal doc As Document)
    EnsureBody doc
    EnsureHeading doc, R_STY_H1, 1
    EnsureHeading doc, R_STY_H2, 2
    EnsureHeading doc, R_STY_H3, 3
    EnsureHeading doc, R_STY_H4, 4
    EnsureHeading doc, R_STY_H5, 5
    EnsureTitle doc
    EnsureTitleBlock doc
    EnsureAttachmentTitle doc
    EnsureBullet doc, R_STY_BULLET_L1, 1
    EnsureBullet doc, R_STY_BULLET_L2, 2
    EnsureBullet doc, R_STY_BULLET_L3, 3
    EnsureBullet doc, R_STY_BULLET_L4, 4
End Sub

' ---------- individual style builders --------------------------------

Private Sub EnsureBody(ByVal doc As Document)
    Dim s As Style
    Set s = GetOrAdd(doc, R_STY_BODY, wdStyleTypeParagraph)
    With s
        .BaseStyle = vbNullString
        .NextParagraphStyle = R_STY_BODY
        .Font.Name = R_BODY_FONT
        .Font.Size = R_BODY_SIZE
        .Font.Bold = False
        .Font.Italic = False
        .Font.Color = wdColorAutomatic
        With .ParagraphFormat
            .LineSpacingRule = wdLineSpaceSingle
            .SpaceBefore = 0
            .SpaceAfter = R_SPACE_AFTER_PT
            .FirstLineIndent = 0
            .LeftIndent = 0
            .Alignment = wdAlignParagraphLeft
            .WidowControl = True
        End With
    End With
End Sub

Private Sub EnsureHeading(ByVal doc As Document, _
                          ByVal name As String, _
                          ByVal level As Long)
    Dim s As Style
    Set s = GetOrAdd(doc, name, wdStyleTypeParagraph)
    With s
        .BaseStyle = R_STY_BODY
        .NextParagraphStyle = R_STY_BODY
        .Font.Name = R_HEADING_FONT
        .Font.Size = R_HEADING_SIZE
        .Font.Bold = True
        .Font.Italic = False
        With .ParagraphFormat
            .LineSpacingRule = wdLineSpaceSingle
            .SpaceBefore = IIf(level = 1, 12, 6)
            .SpaceAfter = R_SPACE_AFTER_PT
            .OutlineLevel = level
            .KeepWithNext = True
            .WidowControl = True
            .LeftIndent = 0
        End With
    End With
End Sub

Private Sub EnsureTitle(ByVal doc As Document)
    Dim s As Style
    Set s = GetOrAdd(doc, R_STY_TITLE, wdStyleTypeParagraph)
    With s
        .BaseStyle = vbNullString
        .NextParagraphStyle = R_STY_BODY
        .Font.Name = R_HEADING_FONT
        .Font.Size = 14
        .Font.Bold = True
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .SpaceBefore = 12
            .SpaceAfter = 12
        End With
    End With
End Sub

Private Sub EnsureTitleBlock(ByVal doc As Document)
    Dim s As Style
    Set s = GetOrAdd(doc, R_STY_TITLEBLOCK, wdStyleTypeParagraph)
    With s
        .BaseStyle = vbNullString
        .NextParagraphStyle = R_STY_TITLEBLOCK
        .Font.Name = R_TITLEBLOCK_FONT
        .Font.Size = R_TITLEBLOCK_SIZE
        .Font.Bold = False
        With .ParagraphFormat
            .LineSpacingRule = wdLineSpaceSingle
            .SpaceBefore = 0
            .SpaceAfter = 0
            .Alignment = wdAlignParagraphLeft
        End With
    End With
End Sub

Private Sub EnsureAttachmentTitle(ByVal doc As Document)
    Dim s As Style
    Set s = GetOrAdd(doc, R_STY_ATTACH_TITLE, wdStyleTypeParagraph)
    With s
        .BaseStyle = R_STY_BODY
        .NextParagraphStyle = R_STY_BODY
        .Font.Bold = True
        .Font.AllCaps = True
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .PageBreakBefore = True
            .SpaceAfter = 12
            .OutlineLevel = 1
            .KeepWithNext = True
        End With
    End With
End Sub

Private Sub EnsureBullet(ByVal doc As Document, _
                         ByVal name As String, _
                         ByVal level As Long)
    Dim s As Style
    Set s = GetOrAdd(doc, name, wdStyleTypeParagraph)
    With s
        .BaseStyle = R_STY_BODY
        .NextParagraphStyle = name
        With .ParagraphFormat
            .LeftIndent = InchesToPts(0.25 * level)
            .FirstLineIndent = InchesToPts(-0.25)
            .SpaceAfter = 3
        End With
    End With
End Sub

' ---------- util ------------------------------------------------------

Private Function GetOrAdd(ByVal doc As Document, _
                          ByVal name As String, _
                          ByVal styType As WdStyleType) As Style
    On Error Resume Next
    Set GetOrAdd = doc.Styles(name)
    On Error GoTo 0
    If GetOrAdd Is Nothing Then
        Set GetOrAdd = doc.Styles.Add(Name:=name, Type:=styType)
    End If
End Function
