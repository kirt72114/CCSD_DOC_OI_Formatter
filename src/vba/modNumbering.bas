Attribute VB_Name = "modNumbering"
Option Explicit

' =====================================================================
' modNumbering
' Builds and applies the T&Q / DAFMAN 90-161 multi-level list:
'   1.    1.1.    1.1.1.    1.1.1.1.    1.1.1.1.1.
' Bound to the five OI heading styles.
' =====================================================================

Public Sub Apply(ByVal doc As Document)
    Dim lt As ListTemplate
    Set lt = BuildOrRefreshTemplate(doc)

    BindHeadingToLevel doc, R_STY_H1, lt, 1
    BindHeadingToLevel doc, R_STY_H2, lt, 2
    BindHeadingToLevel doc, R_STY_H3, lt, 3
    BindHeadingToLevel doc, R_STY_H4, lt, 4
    BindHeadingToLevel doc, R_STY_H5, lt, 5
End Sub

Private Function BuildOrRefreshTemplate(ByVal doc As Document) As ListTemplate
    Const TEMPLATE_NAME As String = "OI Numbering"
    Dim lt As ListTemplate
    Dim found As Boolean

    For Each lt In doc.ListTemplates
        If lt.Name = TEMPLATE_NAME Then
            found = True
            Exit For
        End If
    Next lt

    If Not found Then
        Set lt = doc.ListTemplates.Add(OutlineNumbered:=True, Name:=TEMPLATE_NAME)
    End If

    Dim i As Long
    For i = 1 To R_MAX_NUMBER_DEPTH
        With lt.ListLevels(i)
            .NumberFormat = BuildLevelFormat(i)
            .TrailingCharacter = wdTrailingTab
            .NumberStyle = wdListNumberStyleArabic
            .NumberPosition = InchesToPts(0)
            .Alignment = wdListLevelAlignLeft
            .TextPosition = InchesToPts(0.25 * i)
            .TabPosition = InchesToPts(0.25 * i)
            .ResetOnHigher = i - 1
            .StartAt = 1
            .Font.Name = R_HEADING_FONT
            .Font.Size = R_HEADING_SIZE
            .Font.Bold = True
            .LinkedStyle = HeadingStyleForLevel(i)
        End With
    Next i

    Set BuildOrRefreshTemplate = lt
End Function

' Produces "%1.", "%1.%2.", "%1.%2.%3.", ... using Word's level placeholders.
Private Function BuildLevelFormat(ByVal level As Long) As String
    Dim i As Long
    Dim s As String
    For i = 1 To level
        s = s & Chr(i) & "."
    Next i
    BuildLevelFormat = s
End Function

Private Function HeadingStyleForLevel(ByVal level As Long) As String
    Select Case level
        Case 1: HeadingStyleForLevel = R_STY_H1
        Case 2: HeadingStyleForLevel = R_STY_H2
        Case 3: HeadingStyleForLevel = R_STY_H3
        Case 4: HeadingStyleForLevel = R_STY_H4
        Case 5: HeadingStyleForLevel = R_STY_H5
    End Select
End Function

Private Sub BindHeadingToLevel(ByVal doc As Document, _
                               ByVal styleName As String, _
                               ByVal lt As ListTemplate, _
                               ByVal level As Long)
    Dim p As Paragraph
    For Each p In doc.Paragraphs
        If p.Style = styleName Then
            p.Range.ListFormat.ApplyListTemplate _
                ListTemplate:=lt, _
                ContinuePreviousList:=True, _
                ApplyTo:=wdListApplyToWholeList, _
                DefaultListBehavior:=wdWord10ListBehavior
        End If
    Next p
End Sub
