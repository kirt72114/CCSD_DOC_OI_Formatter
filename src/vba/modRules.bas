Attribute VB_Name = "modRules"
Option Explicit

' =====================================================================
' modRules
' Single source of truth for every USAF OI formatting constant.
' Citations live in docs/rules.md.
' Sources: AFH 33-337 (Tongue and Quill), DAFMAN 90-161.
' =====================================================================

' ---- Fonts -----------------------------------------------------------
Public Const R_BODY_FONT              As String = "Times New Roman"
Public Const R_BODY_SIZE              As Single = 12
Public Const R_HEADING_FONT           As String = "Times New Roman"
Public Const R_HEADING_SIZE           As Single = 12
Public Const R_TITLEBLOCK_FONT        As String = "Arial"
Public Const R_TITLEBLOCK_SIZE        As Single = 10

' ---- Page setup (points; 72 pt = 1 inch) -----------------------------
Public Const R_MARGIN_IN              As Single = 1#
Public Const R_PAGE_WIDTH_IN          As Single = 8.5
Public Const R_PAGE_HEIGHT_IN         As Single = 11#

' ---- Spacing ---------------------------------------------------------
Public Const R_LINE_SPACING_RULE      As Long = 0    ' wdLineSpaceSingle
Public Const R_SPACE_AFTER_PT         As Single = 6

' ---- Numbering -------------------------------------------------------
Public Const R_MAX_NUMBER_DEPTH       As Long = 5

' ---- Bullet sequence (T&Q Ch. 10) ------------------------------------
' L2..L4 use Unicode glyphs so they're exposed as Property Get because
' VBA Const expressions can only be compile-time literals.
Public Const R_BULLET_L1              As String = "-"
Public Property Get R_BULLET_L2() As String: R_BULLET_L2 = ChrW(8226):  End Property
Public Property Get R_BULLET_L3() As String: R_BULLET_L3 = ChrW(8211):  End Property
Public Property Get R_BULLET_L4() As String: R_BULLET_L4 = ChrW(187):   End Property

' ---- Required title-block labels (DAFMAN 90-161 Fig A2.2) ------------
Public Const R_LBL_BYORDER            As String = "BY ORDER OF THE COMMANDER"
Public Const R_LBL_COMPLIANCE         As String = "COMPLIANCE WITH THIS PUBLICATION IS MANDATORY"
Public Const R_LBL_ACCESSIBILITY      As String = "ACCESSIBILITY:"
Public Const R_LBL_RELEASABILITY      As String = "RELEASABILITY:"
Public Const R_LBL_OPR                As String = "OPR:"
Public Const R_LBL_SUPERSEDES         As String = "Supersedes:"
Public Const R_LBL_CERTIFIED_BY       As String = "Certified by:"
Public Const R_LBL_PAGES              As String = "Pages:"

' Default boilerplate if the user leaves optional fields blank.
Public Const R_DEFAULT_ACCESSIBILITY  As String = _
    "Publications and forms are available for downloading or ordering on the " & _
    "e-Publishing website at www.e-Publishing.af.mil."
Public Const R_DEFAULT_RELEASABILITY  As String = _
    "There are no releasability restrictions on this publication."

' ---- Attachment conventions -----------------------------------------
Public Const R_ATTACH_PREFIX          As String = "Attachment "
Public Property Get R_ATTACH_SEP() As String: R_ATTACH_SEP = ChrW(8212): End Property ' em dash
Public Const R_GLOSSARY_TITLE         As String = _
    "GLOSSARY OF REFERENCES AND SUPPORTING INFORMATION"

' ---- Style names (we install / refresh these) -----------------------
Public Const R_STY_BODY               As String = "OI Body"
Public Const R_STY_H1                 As String = "OI Heading 1"
Public Const R_STY_H2                 As String = "OI Heading 2"
Public Const R_STY_H3                 As String = "OI Heading 3"
Public Const R_STY_H4                 As String = "OI Heading 4"
Public Const R_STY_H5                 As String = "OI Heading 5"
Public Const R_STY_TITLE              As String = "OI Title"
Public Const R_STY_TITLEBLOCK         As String = "OI TitleBlock"
Public Const R_STY_ATTACH_TITLE       As String = "OI Attachment Title"
Public Const R_STY_BULLET_L1          As String = "OI Bullet 1"
Public Const R_STY_BULLET_L2          As String = "OI Bullet 2"
Public Const R_STY_BULLET_L3          As String = "OI Bullet 3"
Public Const R_STY_BULLET_L4          As String = "OI Bullet 4"

' ---- Output naming --------------------------------------------------
Public Const R_OUTPUT_SUFFIX          As String = "_formatted"
Public Const R_REPORT_SUFFIX          As String = "_changes.txt"

' ---- Helpers ---------------------------------------------------------
Public Function InchesToPts(ByVal inches As Single) As Single
    InchesToPts = inches * 72#
End Function

Public Function NowStamp() As String
    NowStamp = Format$(Now, "yyyy-mm-dd HH:nn:ss")
End Function
