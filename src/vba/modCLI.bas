Attribute VB_Name = "modCLI"
Option Explicit

' =====================================================================
' modCLI
' Entry points invoked by the PowerShell driver via Application.Run.
' Metadata is passed as a single ";;"-delimited "k=v;;k=v" string so we
' don't depend on marshalling VBA types across COM.
' =====================================================================

' Single file. `metaKvp` is the ";;"-separated key/value pairs from PS.
Public Sub FormatOne(ByVal path As String, ByVal metaKvp As String)
    Dim meta As OIMeta
    ParseMeta metaKvp, meta

    Dim doc As Document
    Set doc = Documents.Open(FileName:=path, ReadOnly:=False, _
                             AddToRecentFiles:=False)
    modFormatter.FormatDocument doc, meta

    Dim outPath As String
    outPath = modFormatter.FormattedOutputPath(path)
    doc.SaveAs2 FileName:=outPath, FileFormat:=wdFormatXMLDocument
    doc.Close SaveChanges:=False
End Sub

' Folder. `recurse01` is "0" or "1".
Public Sub FormatFolder(ByVal folderPath As String, _
                        ByVal recurse01 As String, _
                        ByVal outputDir As String, _
                        ByVal metaKvp As String)
    Dim meta As OIMeta
    ParseMeta metaKvp, meta
    modBatch.RunFolder folderPath, (recurse01 = "1"), meta, outputDir
End Sub

' Handy hook for running on the user's currently open document without
' ever saving — good for in-Word keyboard shortcut.
Public Sub FormatActive()
    Dim meta As OIMeta
    ' Fall back to sensible defaults; the GUI form supersedes these.
    meta.DateStr = Format$(Date, "d mmmm yyyy")
    modFormatter.FormatDocument ActiveDocument, meta
End Sub

' ---------- metadata KVP parser --------------------------------------

Private Sub ParseMeta(ByVal kvp As String, ByRef meta As OIMeta)
    If Len(kvp) = 0 Then Exit Sub
    Dim pairs() As String
    pairs = Split(kvp, ";;")

    Dim i As Long, eq As Long, k As String, v As String
    For i = LBound(pairs) To UBound(pairs)
        eq = InStr(pairs(i), "=")
        If eq > 0 Then
            k = LCase$(Trim$(Left$(pairs(i), eq - 1)))
            v = Mid$(pairs(i), eq + 1)
            Select Case k
                Case "unit":         meta.Unit = v
                Case "unitshort":    meta.UnitShort = v
                Case "oinumber":     meta.OINumber = v
                Case "date":         meta.DateStr = v
                Case "category":     meta.Category = v
                Case "subject":      meta.Subject = v
                Case "opr":          meta.OPR = v
                Case "supersedes":   meta.Supersedes = v
                Case "certifiedby":  meta.CertifiedBy = v
                Case "pages":        meta.Pages = v
                Case "accessibility": meta.Accessibility = v
                Case "releasability": meta.Releasability = v
            End Select
        End If
    Next i
End Sub
