Attribute VB_Name = "modPageSetup"
Option Explicit

' =====================================================================
' modPageSetup
' Enforces margins, paper, orientation, and page numbering.
' =====================================================================

Public Sub Apply(ByVal doc As Document)
    Dim ps As PageSetup
    Set ps = doc.PageSetup

    ps.TopMargin = InchesToPts(R_MARGIN_IN)
    ps.BottomMargin = InchesToPts(R_MARGIN_IN)
    ps.LeftMargin = InchesToPts(R_MARGIN_IN)
    ps.RightMargin = InchesToPts(R_MARGIN_IN)
    ps.PageWidth = InchesToPts(R_PAGE_WIDTH_IN)
    ps.PageHeight = InchesToPts(R_PAGE_HEIGHT_IN)
    ps.Orientation = wdOrientPortrait
    ps.HeaderDistance = InchesToPts(0.5)
    ps.FooterDistance = InchesToPts(0.5)
    ps.DifferentFirstPageHeaderFooter = True  ' title page suppresses number

    InsertPageNumbers doc
End Sub

Private Sub InsertPageNumbers(ByVal doc As Document)
    Dim sec As Section
    Dim ftr As HeaderFooter

    For Each sec In doc.Sections
        ' Suppress number on first (title) page but include on the rest.
        Set ftr = sec.Footers(wdHeaderFooterPrimary)
        ftr.Range.Text = vbNullString
        With ftr.PageNumbers
            .NumberStyle = wdPageNumberStyleArabic
            .HeadingLevelForChapter = 0
            .RestartNumberingAtSection = False
            .StartingNumber = 1
            If .Count = 0 Then
                .Add PageNumberAlignment:=wdAlignPageNumberCenter, _
                     FirstPage:=False
            End If
        End With
    Next sec
End Sub
