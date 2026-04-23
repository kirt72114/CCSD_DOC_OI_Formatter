Attribute VB_Name = "modHeaderBlock"
Option Explicit

' =====================================================================
' modHeaderBlock
' Builds the DAFMAN 90-161 Figure A2.2 compliant OI title block.
' Rendered as a 2-column table at the top of the document, then the
' horizontal rule and the COMPLIANCE banner, then the OPR/Supersedes
' block. All text uses OI TitleBlock style (Arial 10).
' =====================================================================

Public Type OIMeta
    Unit As String              ' e.g. "442D MAINTENANCE SQUADRON"
    UnitShort As String         ' e.g. "442 MXS"
    OINumber As String          ' e.g. "CCSD OI 36-1"
    DateStr As String           ' e.g. "23 April 2026"
    Category As String          ' e.g. "Personnel"
    Subject As String           ' e.g. "Personnel Actions"
    OPR As String               ' e.g. "CCSD/CCC"
    Supersedes As String        ' e.g. "CCSD OI 36-1, 1 Jan 2024"
    CertifiedBy As String       ' e.g. "Col Jane Doe, Commander"
    Pages As String             ' e.g. "12"
    Accessibility As String     ' optional
    Releasability As String     ' optional
End Type

Public Sub Rebuild(ByVal doc As Document, ByRef meta As OIMeta)
    RemoveExistingTitleBlock doc

    Dim anchor As Range
    Set anchor = doc.Range(0, 0)
    anchor.Collapse wdCollapseStart

    InsertTopTable doc, anchor, meta
    InsertHorizontalRule doc, anchor
    InsertComplianceLine doc, anchor
    InsertAccessReleaseBlock doc, anchor, meta
    InsertOPRBlock doc, anchor, meta
    InsertHorizontalRule doc, anchor
    anchor.InsertParagraphAfter
End Sub

' Detects and removes anything the previous run (or prior editor) left
' before the first real heading, so we rebuild cleanly.
Private Sub RemoveExistingTitleBlock(ByVal doc As Document)
    Dim firstHeading As Range
    Set firstHeading = FindFirstHeading(doc)
    If firstHeading Is Nothing Then Exit Sub

    Dim scrub As Range
    Set scrub = doc.Range(0, firstHeading.Start)
    scrub.Delete
End Sub

Private Function FindFirstHeading(ByVal doc As Document) As Range
    Dim p As Paragraph
    For Each p In doc.Paragraphs
        Select Case CStr(p.Style)
            Case R_STY_H1, R_STY_H2, R_STY_H3, R_STY_H4, R_STY_H5, _
                 "Heading 1", "Heading 2", "Heading 3"
                Set FindFirstHeading = p.Range
                Exit Function
        End Select
    Next p
End Function

' Two-column header table: left = BY ORDER / unit, right = OI no / date
' / category / subject.
Private Sub InsertTopTable(ByVal doc As Document, _
                           ByRef anchor As Range, _
                           ByRef meta As OIMeta)
    Dim tbl As Table
    Set tbl = doc.Tables.Add(Range:=anchor, NumRows:=1, NumColumns:=2)
    tbl.Borders.Enable = False
    tbl.AutoFitBehavior wdAutoFitFixed
    tbl.Columns(1).PreferredWidth = InchesToPts(3.25)
    tbl.Columns(2).PreferredWidth = InchesToPts(3.25)

    Dim L As Range, R As Range
    Set L = tbl.Cell(1, 1).Range
    Set R = tbl.Cell(1, 2).Range

    WriteLines L, Array(R_LBL_BYORDER, UCase$(NonEmpty(meta.Unit, "UNIT"))), True
    WriteLines R, Array( _
        UCase$(NonEmpty(meta.OINumber, "UNIT OPERATING INSTRUCTION XX-X")), _
        NonEmpty(meta.DateStr, Format$(Date, "d mmmm yyyy")), _
        "", _
        NonEmpty(meta.Category, "Category"), _
        NonEmpty(meta.Subject, "Subject") _
    ), False

    ApplyTitleBlockStyle tbl.Range
    Set anchor = tbl.Range
    anchor.Collapse wdCollapseEnd
End Sub

Private Sub InsertComplianceLine(ByVal doc As Document, ByRef anchor As Range)
    anchor.InsertParagraphAfter
    anchor.Collapse wdCollapseEnd
    anchor.Text = R_LBL_COMPLIANCE
    anchor.Style = doc.Styles(R_STY_TITLEBLOCK)
    anchor.Bold = True
    anchor.ParagraphFormat.Alignment = wdAlignParagraphCenter
    anchor.InsertParagraphAfter
    anchor.Collapse wdCollapseEnd
End Sub

Private Sub InsertAccessReleaseBlock(ByVal doc As Document, _
                                     ByRef anchor As Range, _
                                     ByRef meta As OIMeta)
    Dim acc As String, rel As String
    acc = NonEmpty(meta.Accessibility, R_DEFAULT_ACCESSIBILITY)
    rel = NonEmpty(meta.Releasability, R_DEFAULT_RELEASABILITY)

    AppendTitleBlockLine doc, anchor, R_LBL_ACCESSIBILITY & "  " & acc
    AppendTitleBlockLine doc, anchor, R_LBL_RELEASABILITY & "  " & rel
End Sub

Private Sub InsertOPRBlock(ByVal doc As Document, _
                           ByRef anchor As Range, _
                           ByRef meta As OIMeta)
    Dim tbl As Table
    anchor.InsertParagraphAfter
    anchor.Collapse wdCollapseEnd

    Set tbl = doc.Tables.Add(Range:=anchor, NumRows:=2, NumColumns:=2)
    tbl.Borders.Enable = False
    tbl.Cell(1, 1).Range.Text = R_LBL_OPR & " " & NonEmpty(meta.OPR, "OPR")
    tbl.Cell(1, 2).Range.Text = R_LBL_CERTIFIED_BY & " " & NonEmpty(meta.CertifiedBy, "TBD")
    tbl.Cell(2, 1).Range.Text = R_LBL_SUPERSEDES & " " & NonEmpty(meta.Supersedes, "N/A")
    tbl.Cell(2, 2).Range.Text = R_LBL_PAGES & " " & NonEmpty(meta.Pages, "TBD")
    ApplyTitleBlockStyle tbl.Range

    Set anchor = tbl.Range
    anchor.Collapse wdCollapseEnd
End Sub

Private Sub InsertHorizontalRule(ByVal doc As Document, ByRef anchor As Range)
    anchor.InsertParagraphAfter
    anchor.Collapse wdCollapseEnd
    With anchor.ParagraphFormat.Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth050pt
    End With
End Sub

' ---------- helpers --------------------------------------------------

Private Sub WriteLines(ByRef target As Range, _
                       ByVal lines As Variant, _
                       ByVal boldFirst As Boolean)
    Dim i As Long
    Dim out As String
    For i = LBound(lines) To UBound(lines)
        If i > LBound(lines) Then out = out & vbCr
        out = out & lines(i)
    Next i
    target.Text = out
    If boldFirst Then
        Dim firstLine As Range
        Set firstLine = target.Duplicate
        firstLine.End = firstLine.Start + Len(CStr(lines(LBound(lines))))
        firstLine.Bold = True
    End If
End Sub

Private Sub AppendTitleBlockLine(ByVal doc As Document, _
                                 ByRef anchor As Range, _
                                 ByVal txt As String)
    anchor.InsertParagraphAfter
    anchor.Collapse wdCollapseEnd
    anchor.Text = txt
    anchor.Style = doc.Styles(R_STY_TITLEBLOCK)
    anchor.InsertParagraphAfter
    anchor.Collapse wdCollapseEnd
End Sub

Private Sub ApplyTitleBlockStyle(ByVal rng As Range)
    rng.Font.Name = R_TITLEBLOCK_FONT
    rng.Font.Size = R_TITLEBLOCK_SIZE
End Sub

Private Function NonEmpty(ByVal s As String, ByVal fallback As String) As String
    If Len(Trim$(s)) = 0 Then
        NonEmpty = fallback
    Else
        NonEmpty = s
    End If
End Function
