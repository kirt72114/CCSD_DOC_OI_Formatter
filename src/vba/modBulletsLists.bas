Attribute VB_Name = "modBulletsLists"
Option Explicit

' =====================================================================
' modBulletsLists
' Normalizes bullet characters to the T&Q sequence based on indent depth.
' Any ASCII-ish bullet glyph at the start of a paragraph is replaced
' with the canonical one for its level. Paragraphs whose style is a
' heading (OI Heading N) are skipped.
' =====================================================================

Public Sub Apply(ByVal doc As Document)
    Dim p As Paragraph
    Dim lvl As Long
    Dim leadChar As String
    Dim rng As Range

    For Each p In doc.Paragraphs
        If IsHeadingStyle(CStr(p.Style)) Then GoTo NextP
        Set rng = p.Range
        If rng.Characters.Count < 2 Then GoTo NextP

        leadChar = Left$(rng.Text, 1)
        If Not LooksLikeBullet(leadChar) Then GoTo NextP

        lvl = LevelFromIndent(p)
        ReplaceLeadingBullet p, BulletForLevel(lvl)
        AssignBulletStyle doc, p, lvl
NextP:
    Next p
End Sub

Private Function IsHeadingStyle(ByVal styleName As String) As Boolean
    IsHeadingStyle = (styleName = R_STY_H1) Or _
                     (styleName = R_STY_H2) Or _
                     (styleName = R_STY_H3) Or _
                     (styleName = R_STY_H4) Or _
                     (styleName = R_STY_H5) Or _
                     (styleName = R_STY_TITLE) Or _
                     (styleName = R_STY_ATTACH_TITLE)
End Function

Private Function LooksLikeBullet(ByVal ch As String) As Boolean
    Select Case ch
        Case "-", "*", "o", ChrW(8226), ChrW(8211), ChrW(8212), ChrW(187), ChrW(9642)
            LooksLikeBullet = True
    End Select
End Function

Private Function LevelFromIndent(ByVal p As Paragraph) As Long
    Dim pts As Single
    pts = p.LeftIndent
    ' 0.25" per level; clamp 1..4
    Dim lvl As Long
    lvl = Int(pts / InchesToPts(0.25)) + 1
    If lvl < 1 Then lvl = 1
    If lvl > 4 Then lvl = 4
    LevelFromIndent = lvl
End Function

Private Function BulletForLevel(ByVal lvl As Long) As String
    Select Case lvl
        Case 1: BulletForLevel = R_BULLET_L1
        Case 2: BulletForLevel = R_BULLET_L2
        Case 3: BulletForLevel = R_BULLET_L3
        Case Else: BulletForLevel = R_BULLET_L4
    End Select
End Function

Private Sub ReplaceLeadingBullet(ByVal p As Paragraph, ByVal newChar As String)
    Dim rng As Range
    Set rng = p.Range.Duplicate
    rng.End = rng.Start + 1
    rng.Text = newChar & " "
    ' Collapse any extra whitespace after.
    Do While Mid$(p.Range.Text, 3, 1) = " "
        p.Range.Characters(3).Delete
    Loop
End Sub

Private Sub AssignBulletStyle(ByVal doc As Document, _
                              ByVal p As Paragraph, _
                              ByVal lvl As Long)
    Select Case lvl
        Case 1: p.Style = doc.Styles(R_STY_BULLET_L1)
        Case 2: p.Style = doc.Styles(R_STY_BULLET_L2)
        Case 3: p.Style = doc.Styles(R_STY_BULLET_L3)
        Case Else: p.Style = doc.Styles(R_STY_BULLET_L4)
    End Select
End Sub
