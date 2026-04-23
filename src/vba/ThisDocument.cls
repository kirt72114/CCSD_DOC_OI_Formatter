VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' =====================================================================
' ThisDocument (for USAF_OI_Formatter.dotm)
' AutoExec adds a toolbar button to the Add-Ins ribbon so users can
' launch the GUI from any Word session once the template is installed
' in %APPDATA%\Microsoft\Word\STARTUP.
' =====================================================================

Private Sub Document_Open()
    On Error Resume Next
    InstallCommandBar
End Sub

Private Sub InstallCommandBar()
    Const BAR_NAME As String = "USAF OI Formatter"
    Dim bar As CommandBar
    On Error Resume Next
    Set bar = Application.CommandBars(BAR_NAME)
    On Error GoTo 0
    If bar Is Nothing Then
        Set bar = Application.CommandBars.Add( _
            Name:=BAR_NAME, Position:=msoBarTop, Temporary:=True)
    End If
    bar.Visible = True

    EnsureButton bar, "Format OI...", "modCLI.FormatActive", _
                 "Run the full USAF OI formatter on the active document"
    EnsureButton bar, "Open Formatter GUI", "ShowFormatterGUI", _
                 "Open the tabbed USAF OI Formatter dialog"
End Sub

Private Sub EnsureButton(ByVal bar As CommandBar, _
                         ByVal caption As String, _
                         ByVal onAction As String, _
                         ByVal tip As String)
    Dim btn As CommandBarButton
    Dim ctl As CommandBarControl
    For Each ctl In bar.Controls
        If ctl.Caption = caption Then
            Set btn = ctl
            Exit For
        End If
    Next ctl
    If btn Is Nothing Then
        Set btn = bar.Controls.Add(Type:=msoControlButton)
    End If
    btn.Caption = caption
    btn.Style = msoButtonCaption
    btn.OnAction = onAction
    btn.TooltipText = tip
End Sub

Public Sub ShowFormatterGUI()
    frmMain.Show
End Sub
