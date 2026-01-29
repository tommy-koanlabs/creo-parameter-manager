VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' =============================================================================
' Sheet1 (Manager) Module - Creo Parameter Manager
' =============================================================================
' Handles UI events for the Manager sheet controls
' =============================================================================

' Note: If using Form buttons, assign these macros directly:
'   Import button -> modParamManager.btnImport_Click
'   Export button -> modParamManager.btnExport_Click
'   Refresh button -> modParamManager.btnRefresh_Click
'
' If using ActiveX buttons, use the event handlers below:

Private Sub CommandButton1_Click()
    ' Import button (if using ActiveX CommandButton)
    btnImport_Click
End Sub

Private Sub CommandButton2_Click()
    ' Export button (if using ActiveX CommandButton)
    btnExport_Click
End Sub

Private Sub CommandButton3_Click()
    ' Refresh button (if using ActiveX CommandButton)
    btnRefresh_Click
End Sub

' Alternative: Form button click handlers (assign in button properties)
' These are public subs that can be assigned to Form buttons

Public Sub Button_Import()
    btnImport_Click
End Sub

Public Sub Button_Export()
    btnExport_Click
End Sub

Public Sub Button_Refresh()
    btnRefresh_Click
End Sub
