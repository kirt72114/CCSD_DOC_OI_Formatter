Attribute VB_Name = "modReport"
Option Explicit

' =====================================================================
' modReport
' Collects a before/after change log for a single document and writes
' it to <input>_changes.txt next to the output file.
' =====================================================================

Private m_lines As Object      ' Collection of String
Private m_startedAt As String
Private m_docName As String
Private m_preSnapshot As Object

Public Sub BeginRun(ByVal doc As Document)
    Set m_lines = New Collection
    m_startedAt = NowStamp()
    m_docName = doc.FullName
    Set m_preSnapshot = Snapshot(doc)
    Log "=== USAF OI Formatter run ==="
    Log "Document:  " & m_docName
    Log "Started:   " & m_startedAt
End Sub

Public Sub Log(ByVal msg As String)
    If m_lines Is Nothing Then Set m_lines = New Collection
    m_lines.Add msg
End Sub

Public Sub Note(ByVal stage As String, ByVal detail As String)
    Log "[" & stage & "] " & detail
End Sub

Public Sub FinishRun(ByVal doc As Document)
    CompareSnapshots doc
    Log "Finished:  " & NowStamp()
    WriteSidecar doc
End Sub

Private Sub CompareSnapshots(ByVal doc As Document)
    Dim post As Object
    Set post = Snapshot(doc)

    Dim k As Variant
    For Each k In post.Keys
        Dim before As String, after As String
        after = CStr(post(k))
        If m_preSnapshot.Exists(k) Then
            before = CStr(m_preSnapshot(k))
        Else
            before = "(unset)"
        End If
        If before <> after Then
            Log "CHANGED " & k & ": " & before & "  =>  " & after
        End If
    Next k
End Sub

Private Function Snapshot(ByVal doc As Document) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    With doc.PageSetup
        d.Add "margin.top_pt", CStr(.TopMargin)
        d.Add "margin.bottom_pt", CStr(.BottomMargin)
        d.Add "margin.left_pt", CStr(.LeftMargin)
        d.Add "margin.right_pt", CStr(.RightMargin)
        d.Add "page.width_pt", CStr(.PageWidth)
        d.Add "page.height_pt", CStr(.PageHeight)
        d.Add "page.orientation", CStr(.Orientation)
    End With

    d.Add "paragraph.count", CStr(doc.Paragraphs.Count)
    d.Add "style.body.exists", CStr(StyleExists(doc, R_STY_BODY))
    d.Add "style.h1.exists", CStr(StyleExists(doc, R_STY_H1))

    Set Snapshot = d
End Function

Private Function StyleExists(ByVal doc As Document, ByVal name As String) As Boolean
    On Error Resume Next
    Dim s As Style
    Set s = doc.Styles(name)
    StyleExists = (Err.Number = 0) And Not (s Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub WriteSidecar(ByVal doc As Document)
    Dim path As String
    path = SidecarPath(doc.FullName)

    Dim fnum As Integer
    fnum = FreeFile
    Open path For Output As #fnum
    Dim i As Long
    For i = 1 To m_lines.Count
        Print #fnum, m_lines(i)
    Next i
    Close #fnum
End Sub

Public Function SidecarPath(ByVal docPath As String) As String
    Dim base As String
    Dim dot As Long
    dot = InStrRev(docPath, ".")
    If dot > 0 Then
        base = Left$(docPath, dot - 1)
    Else
        base = docPath
    End If
    SidecarPath = base & R_REPORT_SUFFIX
End Function

Public Function RenderText() As String
    Dim s As String
    Dim i As Long
    If m_lines Is Nothing Then Exit Function
    For i = 1 To m_lines.Count
        s = s & CStr(m_lines(i)) & vbCrLf
    Next i
    RenderText = s
End Function
