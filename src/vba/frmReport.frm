VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReport
   Caption         =   "USAF OI Formatter - Change Report"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   OleObjectBlob   =   "frmReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
' frmReport
' Read-only viewer for the change log produced by modReport.
'
' Expected controls:
'   txtLog     (TextBox, MultiLine=True, ScrollBars=2, Locked=True)
'   cmdCopy    (CommandButton - copies to clipboard)
'   cmdClose   (CommandButton)
' =====================================================================

Public Sub SetText(ByVal s As String)
    txtLog.Value = s
End Sub

Private Sub cmdCopy_Click()
    Dim dobj As Object
    Set dobj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dobj.SetText txtLog.Value
    dobj.PutInClipboard
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
