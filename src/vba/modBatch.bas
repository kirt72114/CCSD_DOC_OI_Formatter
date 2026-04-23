Attribute VB_Name = "modBatch"
Option Explicit

' =====================================================================
' modBatch
' Iterates a folder of .docx files and runs the full formatter on each.
' Honors a recurse flag; writes a master log next to the folder.
' =====================================================================

Public Sub RunFolder(ByVal folderPath As String, _
                     ByVal recurse As Boolean, _
                     ByRef meta As OIMeta, _
                     Optional ByVal outputDir As String = vbNullString)
    If Right$(folderPath, 1) <> Application.PathSeparator Then
        folderPath = folderPath & Application.PathSeparator
    End If

    Dim logPath As String
    logPath = folderPath & "batch_" & Format$(Now, "yyyymmdd_hhnnss") & ".log"

    Dim fnum As Integer
    fnum = FreeFile
    Open logPath For Output As #fnum
    Print #fnum, "USAF OI Formatter batch run " & modRules.NowStamp()
    Print #fnum, "Folder: " & folderPath
    Print #fnum, "Recurse: " & CStr(recurse)
    Print #fnum, ""

    ProcessFolder folderPath, recurse, meta, outputDir, fnum

    Close #fnum
End Sub

Private Sub ProcessFolder(ByVal folderPath As String, _
                          ByVal recurse As Boolean, _
                          ByRef meta As OIMeta, _
                          ByVal outputDir As String, _
                          ByVal fnum As Integer)
    Dim name As String
    name = Dir(folderPath & "*.docx")
    Do While Len(name) > 0
        If Not IsFormatterOutput(name) Then
            ProcessOne folderPath & name, meta, outputDir, fnum
        End If
        name = Dir()
    Loop

    If recurse Then
        Dim sub_ As String
        sub_ = Dir(folderPath, vbDirectory)
        Do While Len(sub_) > 0
            If sub_ <> "." And sub_ <> ".." Then
                If (GetAttr(folderPath & sub_) And vbDirectory) = vbDirectory Then
                    ProcessFolder folderPath & sub_ & Application.PathSeparator, _
                                  True, meta, outputDir, fnum
                End If
            End If
            sub_ = Dir()
        Loop
    End If
End Sub

Private Function IsFormatterOutput(ByVal fileName As String) As Boolean
    IsFormatterOutput = (InStr(1, fileName, R_OUTPUT_SUFFIX & ".docx", _
                               vbTextCompare) > 0)
End Function

Private Sub ProcessOne(ByVal path As String, _
                       ByRef meta As OIMeta, _
                       ByVal outputDir As String, _
                       ByVal fnum As Integer)
    Dim doc As Document
    On Error GoTo Fail
    Set doc = Documents.Open(FileName:=path, ReadOnly:=False, AddToRecentFiles:=False)

    modFormatter.FormatDocument doc, meta

    Dim outPath As String
    outPath = ResolveOutputPath(path, outputDir)
    doc.SaveAs2 FileName:=outPath, FileFormat:=wdFormatXMLDocument
    doc.Close SaveChanges:=False
    Print #fnum, "OK    " & path & "  ->  " & outPath
    Exit Sub

Fail:
    Print #fnum, "FAIL  " & path & "  #" & Err.Number & " " & Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=False
    Err.Clear
End Sub

Private Function ResolveOutputPath(ByVal srcPath As String, _
                                   ByVal outputDir As String) As String
    Dim formatted As String
    formatted = modFormatter.FormattedOutputPath(srcPath)

    If Len(outputDir) = 0 Then
        ResolveOutputPath = formatted
        Exit Function
    End If

    If Right$(outputDir, 1) <> Application.PathSeparator Then
        outputDir = outputDir & Application.PathSeparator
    End If

    Dim leaf As String
    Dim slash As Long
    slash = InStrRev(formatted, Application.PathSeparator)
    If slash > 0 Then
        leaf = Mid$(formatted, slash + 1)
    Else
        leaf = formatted
    End If
    ResolveOutputPath = outputDir & leaf
End Function
