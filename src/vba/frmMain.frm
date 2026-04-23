VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain
   Caption         =   "USAF OI Formatter"
   ClientHeight    =   10200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8700
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
' frmMain
' Tabbed UserForm for USAF OI Formatter.
' The form is built at design time; this module only wires up events.
'
' Expected controls (MultiPage named 'tabs', with pages Single, Batch,
' Meta, Options):
'
'   Single page:
'     txtFilePath   (TextBox)
'     cmdBrowseFile (CommandButton)
'
'   Batch page:
'     txtFolderPath (TextBox)
'     cmdBrowseFolder (CommandButton)
'     chkRecurse    (CheckBox)
'     txtOutputDir  (TextBox)
'     cmdBrowseOut  (CommandButton)
'
'   Meta page:
'     txtUnit, txtUnitShort, txtOINumber, txtDate, txtCategory, txtSubject
'     txtOPR, txtSupersedes, txtCertifiedBy, txtPages
'     txtAccessibility, txtReleasability
'
'   Options page:
'     chkSkipAcronyms, chkOpenReport
'
'   Footer:
'     cmdRun, cmdCancel
'
' If you edit the form in Word's VBA editor, re-export and commit both
' frmMain.frm and frmMain.frx.
' =====================================================================

Private Sub UserForm_Initialize()
    LoadDefaults
End Sub

Private Sub cmdBrowseFile_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "Word documents", "*.docx;*.docm", 1
    If fd.Show = -1 Then txtFilePath.Value = fd.SelectedItems(1)
End Sub

Private Sub cmdBrowseFolder_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    If fd.Show = -1 Then txtFolderPath.Value = fd.SelectedItems(1)
End Sub

Private Sub cmdBrowseOut_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    If fd.Show = -1 Then txtOutputDir.Value = fd.SelectedItems(1)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRun_Click()
    Dim meta As OIMeta
    BuildMeta meta
    SaveDefaults

    On Error GoTo Fail
    Select Case tabs.Value
        Case 0   ' Single
            RunSingle meta
        Case 1   ' Batch
            RunBatch meta
        Case Else
            MsgBox "Pick a Single or Batch tab first.", vbExclamation
            Exit Sub
    End Select

    If chkOpenReport.Value Then ShowReport
    Unload Me
    Exit Sub

Fail:
    MsgBox "Formatter failed: #" & Err.Number & " " & Err.Description, _
           vbCritical, "USAF OI Formatter"
End Sub

Private Sub RunSingle(ByRef meta As OIMeta)
    If Len(Trim$(txtFilePath.Value)) = 0 Then
        MsgBox "Pick a .docx file first.", vbExclamation: Exit Sub
    End If
    Dim doc As Document
    Set doc = Documents.Open(FileName:=txtFilePath.Value, _
                             ReadOnly:=False, AddToRecentFiles:=False)
    modFormatter.FormatAndSave doc, meta
End Sub

Private Sub RunBatch(ByRef meta As OIMeta)
    If Len(Trim$(txtFolderPath.Value)) = 0 Then
        MsgBox "Pick a folder first.", vbExclamation: Exit Sub
    End If
    modBatch.RunFolder txtFolderPath.Value, _
                       CBool(chkRecurse.Value), _
                       meta, _
                       CStr(txtOutputDir.Value)
End Sub

Private Sub BuildMeta(ByRef meta As OIMeta)
    meta.Unit = txtUnit.Value
    meta.UnitShort = txtUnitShort.Value
    meta.OINumber = txtOINumber.Value
    meta.DateStr = txtDate.Value
    meta.Category = txtCategory.Value
    meta.Subject = txtSubject.Value
    meta.OPR = txtOPR.Value
    meta.Supersedes = txtSupersedes.Value
    meta.CertifiedBy = txtCertifiedBy.Value
    meta.Pages = txtPages.Value
    meta.Accessibility = txtAccessibility.Value
    meta.Releasability = txtReleasability.Value
End Sub

Private Sub LoadDefaults()
    On Error Resume Next
    txtUnit.Value = GetSetting("USAF_OI_Formatter", "Meta", "Unit", "")
    txtUnitShort.Value = GetSetting("USAF_OI_Formatter", "Meta", "UnitShort", "")
    txtOINumber.Value = GetSetting("USAF_OI_Formatter", "Meta", "OINumber", "")
    txtDate.Value = Format$(Date, "d mmmm yyyy")
    txtCategory.Value = GetSetting("USAF_OI_Formatter", "Meta", "Category", "")
    txtOPR.Value = GetSetting("USAF_OI_Formatter", "Meta", "OPR", "")
    txtCertifiedBy.Value = GetSetting("USAF_OI_Formatter", "Meta", "CertifiedBy", "")
    txtAccessibility.Value = GetSetting("USAF_OI_Formatter", "Meta", "Accessibility", _
                                        R_DEFAULT_ACCESSIBILITY)
    txtReleasability.Value = GetSetting("USAF_OI_Formatter", "Meta", "Releasability", _
                                        R_DEFAULT_RELEASABILITY)
    chkOpenReport.Value = True
    On Error GoTo 0
End Sub

Private Sub SaveDefaults()
    On Error Resume Next
    SaveSetting "USAF_OI_Formatter", "Meta", "Unit", CStr(txtUnit.Value)
    SaveSetting "USAF_OI_Formatter", "Meta", "UnitShort", CStr(txtUnitShort.Value)
    SaveSetting "USAF_OI_Formatter", "Meta", "OINumber", CStr(txtOINumber.Value)
    SaveSetting "USAF_OI_Formatter", "Meta", "Category", CStr(txtCategory.Value)
    SaveSetting "USAF_OI_Formatter", "Meta", "OPR", CStr(txtOPR.Value)
    SaveSetting "USAF_OI_Formatter", "Meta", "CertifiedBy", CStr(txtCertifiedBy.Value)
    SaveSetting "USAF_OI_Formatter", "Meta", "Accessibility", CStr(txtAccessibility.Value)
    SaveSetting "USAF_OI_Formatter", "Meta", "Releasability", CStr(txtReleasability.Value)
    On Error GoTo 0
End Sub

Private Sub ShowReport()
    frmReport.SetText modReport.RenderText()
    frmReport.Show
End Sub
