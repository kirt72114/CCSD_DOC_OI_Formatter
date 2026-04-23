Attribute VB_Name = "modFormatter"
Option Explicit

' =====================================================================
' modFormatter
' Orchestrates the full formatting pipeline against one document.
' Called from frmMain (GUI), modCLI (headless), and modBatch (folder).
' =====================================================================

Public Sub FormatDocument(ByVal doc As Document, ByRef meta As OIMeta)
    Dim prevScreenUpdating As Boolean
    prevScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    On Error GoTo Cleanup

    modReport.BeginRun doc
    modReport.Note "page", "Applying margins and page setup"
    modPageSetup.Apply doc

    modReport.Note "styles", "Installing OI styles"
    modStyles.InstallOrRefresh doc

    modReport.Note "header", "Rebuilding DAFMAN 90-161 title block"
    modHeaderBlock.Rebuild doc, meta

    modReport.Note "walk", "Classifying paragraphs"
    ClassifyParagraphs doc

    modReport.Note "numbering", "Applying 1. / 1.1. / 1.1.1. numbering"
    modNumbering.Apply doc

    modReport.Note "bullets", "Normalizing bullets"
    modBulletsLists.Apply doc

    modReport.Note "acronyms", "Collecting acronyms"
    modAcronyms.Apply doc

    modReport.Note "attachments", "Rebuilding attachment titles"
    modAttachments.Apply doc

    modReport.Note "hygiene", "Text hygiene (whitespace, quotes, dashes)"
    CleanText doc

    modReport.Note "finalize", "Widow/orphan, field update"
    FinalizePolish doc

    modReport.FinishRun doc

Cleanup:
    Application.ScreenUpdating = prevScreenUpdating
    If Err.Number <> 0 Then
        modReport.Note "ERROR", "#" & Err.Number & " " & Err.Description
        modReport.FinishRun doc
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

' Runs the full pipeline and saves a copy as <name>_formatted.docx.
Public Sub FormatAndSave(ByVal doc As Document, ByRef meta As OIMeta)
    FormatDocument doc, meta
    Dim outPath As String
    outPath = FormattedOutputPath(doc.FullName)
    doc.SaveAs2 FileName:=outPath, FileFormat:=wdFormatXMLDocument
End Sub

Public Function FormattedOutputPath(ByVal docPath As String) As String
    Dim dot As Long
    dot = InStrRev(docPath, ".")
    If dot > 0 Then
        FormattedOutputPath = Left$(docPath, dot - 1) & R_OUTPUT_SUFFIX & _
                              Mid$(docPath, dot)
    Else
        FormattedOutputPath = docPath & R_OUTPUT_SUFFIX
    End If
End Function

' ---------- classification -------------------------------------------

Private Sub ClassifyParagraphs(ByVal doc As Document)
    Dim p As Paragraph
    Dim text As String
    Dim trimmed As String

    For Each p In doc.Paragraphs
        text = p.Range.Text
        trimmed = Trim$(Replace(Replace(text, vbCr, ""), vbLf, ""))
        If Len(trimmed) = 0 Then GoTo NextP

        ' Skip rows inside tables (title block uses them).
        If p.Range.Information(wdWithInTable) Then GoTo NextP

        Dim lvl As Long
        lvl = LeadingNumberDepth(trimmed)
        If lvl > 0 And lvl <= R_MAX_NUMBER_DEPTH Then
            p.Style = doc.Styles(HeadingStyleForLevel(lvl))
        ElseIf IsAllCapsHeading(trimmed) Then
            p.Style = doc.Styles(R_STY_H1)
        Else
            p.Style = doc.Styles(R_STY_BODY)
        End If
NextP:
    Next p
End Sub

' Returns the depth implied by a leading "1.", "1.2.", "1.2.3." etc,
' or 0 if the paragraph does not start with a numbered prefix.
Private Function LeadingNumberDepth(ByVal s As String) As Long
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^(\d+(?:\.\d+){0,4})\.?\s+"
    If Not re.Test(s) Then Exit Function
    Dim hit As String
    hit = re.Execute(s)(0).SubMatches(0)
    LeadingNumberDepth = 1 + CountChar(hit, ".")
End Function

Private Function CountChar(ByVal s As String, ByVal ch As String) As Long
    Dim i As Long, n As Long
    For i = 1 To Len(s)
        If Mid$(s, i, 1) = ch Then n = n + 1
    Next i
    CountChar = n
End Function

Private Function IsAllCapsHeading(ByVal s As String) As Boolean
    If Len(s) < 3 Or Len(s) > 120 Then Exit Function
    If s <> UCase$(s) Then Exit Function
    ' At least one letter.
    Dim i As Long
    For i = 1 To Len(s)
        If Mid$(s, i, 1) Like "[A-Z]" Then
            IsAllCapsHeading = True
            Exit Function
        End If
    Next i
End Function

Private Function HeadingStyleForLevel(ByVal level As Long) As String
    Select Case level
        Case 1: HeadingStyleForLevel = R_STY_H1
        Case 2: HeadingStyleForLevel = R_STY_H2
        Case 3: HeadingStyleForLevel = R_STY_H3
        Case 4: HeadingStyleForLevel = R_STY_H4
        Case 5: HeadingStyleForLevel = R_STY_H5
        Case Else: HeadingStyleForLevel = R_STY_BODY
    End Select
End Function

' ---------- text hygiene ---------------------------------------------

Private Sub CleanText(ByVal doc As Document)
    ReplaceAll doc, ".  ", ". "                   ' double space after period
    ReplaceAll doc, "  ", " "                     ' collapse double spaces
    ReplaceAll doc, ChrW(8220), """"              ' “
    ReplaceAll doc, ChrW(8221), """"              ' ”
    ReplaceAll doc, ChrW(8216), "'"               ' ‘
    ReplaceAll doc, ChrW(8217), "'"               ' ’
End Sub

Private Sub ReplaceAll(ByVal doc As Document, ByVal findText As String, _
                       ByVal replText As String)
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replText
        .Forward = True
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Private Sub FinalizePolish(ByVal doc As Document)
    Dim p As Paragraph
    For Each p In doc.Paragraphs
        p.Format.WidowControl = True
    Next p
    doc.Fields.Update
End Sub
