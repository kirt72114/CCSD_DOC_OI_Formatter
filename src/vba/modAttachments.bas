Attribute VB_Name = "modAttachments"
Option Explicit

' =====================================================================
' modAttachments
' Rebuilds attachment titles so they read:
'   Attachment N—TITLE IN ALL CAPS
' If the document has no Attachment 1, inserts a Glossary attachment
' containing the acronyms collected by modAcronyms.
' =====================================================================

Public Sub Apply(ByVal doc As Document)
    NormalizeExistingAttachmentHeadings doc

    If Not HasAttachment1(doc) Then
        InsertGlossaryAttachment doc
    End If
End Sub

Private Sub NormalizeExistingAttachmentHeadings(ByVal doc As Document)
    Dim p As Paragraph
    Dim text As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "^\s*attachment\s+(\d+)\s*[-" & ChrW(8211) & ChrW(8212) & _
                 ":.]?\s*(.*)$"

    For Each p In doc.Paragraphs
        text = RTrimCR(p.Range.Text)
        If re.Test(text) Then
            Dim matches As Object
            Set matches = re.Execute(text)
            Dim num As String, rest As String
            num = matches(0).SubMatches(0)
            rest = UCase$(Trim$(CStr(matches(0).SubMatches(1))))
            p.Range.Text = R_ATTACH_PREFIX & num & R_ATTACH_SEP & rest & vbCr
            p.Style = doc.Styles(R_STY_ATTACH_TITLE)
        End If
    Next p
End Sub

Private Function HasAttachment1(ByVal doc As Document) As Boolean
    Dim p As Paragraph
    For Each p In doc.Paragraphs
        If InStr(1, RTrimCR(p.Range.Text), R_ATTACH_PREFIX & "1" & R_ATTACH_SEP, _
                 vbBinaryCompare) = 1 Then
            HasAttachment1 = True
            Exit Function
        End If
    Next p
End Function

Private Sub InsertGlossaryAttachment(ByVal doc As Document)
    Dim tail As Range
    Set tail = doc.Range
    tail.Collapse wdCollapseEnd

    tail.InsertParagraphAfter
    tail.Collapse wdCollapseEnd
    tail.Text = R_ATTACH_PREFIX & "1" & R_ATTACH_SEP & R_GLOSSARY_TITLE
    tail.Style = doc.Styles(R_STY_ATTACH_TITLE)
    tail.InsertParagraphAfter
    tail.Collapse wdCollapseEnd

    WriteGlossaryEntries doc, tail
End Sub

Private Sub WriteGlossaryEntries(ByVal doc As Document, ByRef anchor As Range)
    Dim dict As Object
    Set dict = modAcronyms.GetCollectedAcronyms()
    If dict Is Nothing Then Exit Sub
    If dict.Count = 0 Then Exit Sub

    ' Sort keys alphabetically.
    Dim keys As Variant
    keys = SortedKeys(dict)

    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        Dim line As String
        line = CStr(keys(i)) & vbTab & CStr(dict(keys(i)))
        anchor.InsertParagraphAfter
        anchor.Collapse wdCollapseEnd
        anchor.Text = line
        anchor.Style = doc.Styles(R_STY_BODY)
    Next i
End Sub

Private Function SortedKeys(ByVal dict As Object) As Variant
    Dim arr() As String
    Dim n As Long: n = dict.Count
    ReDim arr(0 To n - 1)

    Dim i As Long, k As Variant
    i = 0
    For Each k In dict.Keys
        arr(i) = CStr(k)
        i = i + 1
    Next k

    Dim j As Long, tmp As String
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If arr(j) < arr(i) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i

    SortedKeys = arr
End Function

Private Function RTrimCR(ByVal s As String) As String
    Do While Len(s) > 0 And (Right$(s, 1) = vbCr Or Right$(s, 1) = vbLf _
                             Or Right$(s, 1) = Chr(7))
        s = Left$(s, Len(s) - 1)
    Loop
    RTrimCR = s
End Function
