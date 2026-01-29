VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' =============================================================================
' ThisWorkbook Module - Creo Parameter Manager
' =============================================================================
' Handles workbook-level events for auto-refresh functionality
' =============================================================================

Private Sub Workbook_Activate()
    ' Auto-refresh XML file list when workbook is activated
    ' This catches new files added while workbook was in background

    On Error Resume Next
    RefreshXMLFileList
    On Error GoTo 0
End Sub

Private Sub Workbook_Open()
    ' Initialize both lists when workbook is opened

    On Error Resume Next
    RefreshAllLists
    On Error GoTo 0
End Sub
