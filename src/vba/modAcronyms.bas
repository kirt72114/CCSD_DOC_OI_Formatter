Attribute VB_Name = "modAcronyms"
Option Explicit

' =====================================================================
' modAcronyms
' First-use expansion: when the spelled-out phrase appears immediately
' before the parenthesized acronym ("Operating Instruction (OI)"), the
' pair is kept; subsequent uses are left alone. Acronyms with no known
' expansion are collected for the glossary but never rewritten.
' A small seed dictionary ships with the tool; unknown acronyms are
' reported for human review.
' =====================================================================

Private m_seen As Object              ' Scripting.Dictionary: acronym -> True
Private m_glossary As Object          ' Scripting.Dictionary: acronym -> expansion

Public Sub Apply(ByVal doc As Document)
    InitState
    SeedDictionary

    Dim p As Paragraph
    For Each p In doc.Paragraphs
        If IsBlock(CStr(p.Style)) Then ScanParagraph p
    Next p
End Sub

' Called by modAttachments to emit the glossary contents.
Public Function GetCollectedAcronyms() As Object
    If m_glossary Is Nothing Then InitState
    Set GetCollectedAcronyms = m_glossary
End Function

Private Sub InitState()
    Set m_seen = CreateObject("Scripting.Dictionary")
    Set m_glossary = CreateObject("Scripting.Dictionary")
    m_seen.CompareMode = 0          ' BinaryCompare = case sensitive
    m_glossary.CompareMode = 0
End Sub

Private Sub SeedDictionary()
    ' Acronyms we expect to see in USAF OIs. Extend freely.
    AddKnown "OI", "Operating Instruction"
    AddKnown "AFI", "Air Force Instruction"
    AddKnown "AFMAN", "Air Force Manual"
    AddKnown "AFPD", "Air Force Policy Directive"
    AddKnown "AFH", "Air Force Handbook"
    AddKnown "DAFI", "Department of the Air Force Instruction"
    AddKnown "DAFMAN", "Department of the Air Force Manual"
    AddKnown "OPR", "Office of Primary Responsibility"
    AddKnown "USAF", "United States Air Force"
    AddKnown "DoD", "Department of Defense"
    AddKnown "POC", "Point of Contact"
    AddKnown "T&Q", "Tongue and Quill"
End Sub

Private Sub AddKnown(ByVal acr As String, ByVal expansion As String)
    If Not m_glossary.Exists(acr) Then m_glossary.Add acr, expansion
End Sub

Private Function IsBlock(ByVal styleName As String) As Boolean
    Select Case styleName
        Case R_STY_TITLEBLOCK, R_STY_TITLE, R_STY_ATTACH_TITLE
            IsBlock = False
        Case Else
            IsBlock = True
    End Select
End Function

' Scans a paragraph for ALL-CAPS tokens 2..7 chars long and tracks
' first use. Does NOT rewrite existing correct expansions; only flags
' any acronym whose first occurrence is NOT preceded by a parenthesized
' expansion pair.
Private Sub ScanParagraph(ByVal p As Paragraph)
    Dim text As String
    text = p.Range.Text
    If Len(text) < 3 Then Exit Sub

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = False
    re.Pattern = "\b[A-Z][A-Z0-9&]{1,6}\b"

    Dim matches As Object, m As Object
    Set matches = re.Execute(text)
    For Each m In matches
        Dim token As String
        token = m.Value
        If Not IsNoiseToken(token) Then
            If Not m_seen.Exists(token) Then
                m_seen.Add token, True
                If Not m_glossary.Exists(token) Then
                    m_glossary.Add token, "TBD - define on first use"
                End If
            End If
        End If
    Next m
End Sub

Private Function IsNoiseToken(ByVal s As String) As Boolean
    Select Case s
        Case "A", "I", "AN", "AT", "BE", "BY", "DO", "GO", "HE", "IF", "IN", _
             "IS", "IT", "ME", "MY", "NO", "OF", "ON", "OR", "SO", "TO", "UP", _
             "US", "WE", "THE", "AND", "FOR", "BUT", "NOT", "YOU", "ALL", "CAN"
            IsNoiseToken = True
    End Select
End Function
