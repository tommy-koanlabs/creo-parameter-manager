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

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' Handle cell changes to update formatting for user-added data in partial fields
    ' When user fills an empty cell in a partial additional field, change it from grey to yellow

    Dim cell As Range
    Dim markerValue As String
    Dim hasValue As Boolean
    Dim isPartialField As Boolean

    On Error Resume Next

    ' Only process data sheets (not Manager sheet)
    If Sh.Index = 1 Then Exit Sub

    ' Only process if marker row exists
    If Sh.Rows(2).Hidden = False And Sh.Cells(2, 1).Value <> "F" And Sh.Cells(2, 1).Value <> "P" Then Exit Sub
    If Sh.Rows(2).Hidden = True Or Sh.Cells(2, 1).Value = "F" Or Sh.Cells(2, 1).Value = "P" Then
        ' Marker row exists, process changes
    Else
        Exit Sub
    End If

    Application.EnableEvents = False ' Prevent recursive calls

    For Each cell In Target
        ' Only process cells in data rows (row 3+)
        If cell.row >= 3 Then
            ' Check if this column is a partial field
            markerValue = Sh.Cells(2, cell.Column).Value
            isPartialField = (markerValue = "P")

            If isPartialField Then
                hasValue = (Trim(CStr(cell.Value)) <> "")

                ' If cell now has value and wasn't blue (original data), make it yellow
                If hasValue And cell.Interior.Color <> RGB(204, 204, 255) Then
                    ' User added data to an empty cell
                    cell.Interior.Color = RGB(255, 255, 204) ' Light yellow
                ElseIf Not hasValue Then
                    ' Cell was cleared, return to grey (conditional formatting will handle this)
                    cell.Interior.ColorIndex = xlNone ' Remove direct formatting to let conditional formatting show
                End If
            End If
        End If
    Next cell

    Application.EnableEvents = True
    On Error GoTo 0
End Sub
