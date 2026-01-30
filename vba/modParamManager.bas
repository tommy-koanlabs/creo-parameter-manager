Attribute VB_Name = "modParamManager"
Option Explicit

' =============================================================================
' Creo Parameter Manager - Main Module
' =============================================================================
' Handles XML import/export and list management for Creo parameter files
' =============================================================================

' Priority fields for column ordering (PTC_WM_NAME always first)
Private Const PRIORITY_FIELDS As String = "PTC_WM_NAME,CAGE_CODE,PART_NUMBER,DESCRIPTION_1,DESCRIPTION_2"

' Required fields that should warn if blank on export
Private Const REQUIRED_FIELDS As String = "CAGE_CODE,DESCRIPTION_1,PART_NUMBER,PTC_WM_NAME"

' =============================================================================
' UTILITIES
' =============================================================================

Private Function CreateXMLDocument() As Object
    ' Creates an MSXML DOM document, trying multiple versions for compatibility
    Dim xmlDoc As Object

    On Error Resume Next

    ' Try MSXML 6.0 first (preferred)
    Set xmlDoc = CreateObject("MSXML2.DOMDocument60")
    If xmlDoc Is Nothing Then
        ' Try MSXML 3.0
        Set xmlDoc = CreateObject("MSXML2.DOMDocument30")
    End If
    If xmlDoc Is Nothing Then
        ' Try generic version
        Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    End If
    If xmlDoc Is Nothing Then
        ' Try Microsoft.XMLDOM as last resort
        Set xmlDoc = CreateObject("Microsoft.XMLDOM")
    End If

    On Error GoTo 0

    If Not xmlDoc Is Nothing Then
        xmlDoc.Async = False
        xmlDoc.validateOnParse = False
    End If

    Set CreateXMLDocument = xmlDoc
End Function

Private Function GetLocalPath(ByVal pathStr As String) As String
    ' Converts SharePoint/OneDrive URL to local sync folder path

    Dim localPath As String
    Dim docPos As Long
    Dim relativePath As String
    Dim oneDrivePath As String
    Dim fso As Object
    Dim userFolder As Object
    Dim subFolder As Object

    ' If not a URL, return as-is
    If Left(pathStr, 8) <> "https://" And Left(pathStr, 7) <> "http://" Then
        GetLocalPath = pathStr
        Exit Function
    End If

    ' Extract path after /Documents/
    docPos = InStr(1, pathStr, "/Documents/", vbTextCompare)
    If docPos > 0 Then
        relativePath = Mid(pathStr, docPos + Len("/Documents/"))
        ' Replace forward slashes with backslashes
        relativePath = Replace(relativePath, "/", "\")
    Else
        ' No /Documents/ found - can't convert
        GetLocalPath = pathStr
        Exit Function
    End If

    ' Try OneDrive environment variables (Business first, then Personal)
    oneDrivePath = Environ("OneDriveCommercial")
    If oneDrivePath = "" Then oneDrivePath = Environ("OneDriveConsumer")
    If oneDrivePath = "" Then oneDrivePath = Environ("OneDrive")

    ' If env var found, construct path
    If oneDrivePath <> "" Then
        localPath = oneDrivePath & "\" & relativePath
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(localPath) Then
            GetLocalPath = localPath
            Exit Function
        End If
    End If

    ' Fallback: scan user profile for OneDrive folders
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set userFolder = fso.GetFolder(Environ("USERPROFILE"))

    For Each subFolder In userFolder.SubFolders
        If Left(subFolder.Name, 8) = "OneDrive" Then
            localPath = subFolder.Path & "\" & relativePath
            If fso.FolderExists(localPath) Then
                GetLocalPath = localPath
                Exit Function
            End If
        End If
    Next subFolder

    ' Could not resolve - return original
    GetLocalPath = pathStr
End Function

' =============================================================================
' LIST REFRESH PROCEDURES
' =============================================================================

Public Sub RefreshXMLFileList()
    ' Populates ListBox1 with .xml files in workbook directory, sorted by date (newest first)

    Dim ws As Worksheet
    Dim lb As MSForms.ListBox
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim xmlFiles As Collection
    Dim fileDates As Collection
    Dim workbookPath As String
    Dim i As Long, j As Long
    Dim tempName As String
    Dim tempDate As Date

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Sheets(1)
    Set lb = ws.OLEObjects("ListBox1").Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set xmlFiles = New Collection
    Set fileDates = New Collection

    workbookPath = ThisWorkbook.Path
    If workbookPath = "" Then
        lb.Clear
        lb.AddItem "(Save workbook first to see XML files)"
        Exit Sub
    End If

    ' Convert SharePoint/OneDrive URL to local path
    workbookPath = GetLocalPath(workbookPath)

    Set folder = fso.GetFolder(workbookPath)

    ' Collect all XML files
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "xml" Then
            xmlFiles.Add file.Name
            fileDates.Add file.DateLastModified
        End If
    Next file

    ' Sort by date descending (bubble sort - fine for small lists)
    For i = 1 To xmlFiles.Count - 1
        For j = i + 1 To xmlFiles.Count
            If fileDates(j) > fileDates(i) Then
                ' Swap
                tempName = xmlFiles(i)
                tempDate = fileDates(i)
                xmlFiles.Remove i
                xmlFiles.Add tempName, , , i - 1
                fileDates.Remove i
                fileDates.Add tempDate, , , i - 1

                xmlFiles.Remove i
                xmlFiles.Add xmlFiles(j - 1), , i
                fileDates.Remove i
                fileDates.Add fileDates(j - 1), , i

                xmlFiles.Remove j
                xmlFiles.Add tempName, , j
                fileDates.Remove j
                fileDates.Add tempDate, , j
            End If
        Next j
    Next i

    ' Populate listbox
    lb.Clear
    For i = 1 To xmlFiles.Count
        lb.AddItem xmlFiles(i)
    Next i

    If xmlFiles.Count = 0 Then
        lb.AddItem "(No XML files found)"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error refreshing XML file list: " & Err.Description, vbExclamation
End Sub

Public Sub RefreshSheetList()
    ' Populates ListBox2 with data sheets (excluding Manager), sorted newest first

    Dim ws As Worksheet
    Dim lb As MSForms.ListBox
    Dim dataSheets As Collection
    Dim sheet As Worksheet
    Dim i As Long, j As Long
    Dim tempName As String

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Sheets(1)
    Set lb = ws.OLEObjects("ListBox2").Object
    Set dataSheets = New Collection

    ' Collect data sheets (all except first sheet)
    For Each sheet In ThisWorkbook.Sheets
        If sheet.Index > 1 Then
            dataSheets.Add sheet.Name
        End If
    Next sheet

    ' Sort alphabetically descending (timestamps in name = newest first)
    For i = 1 To dataSheets.Count - 1
        For j = i + 1 To dataSheets.Count
            If dataSheets(j) > dataSheets(i) Then
                tempName = dataSheets(i)
                dataSheets.Remove i
                dataSheets.Add tempName, , , i - 1
                dataSheets.Remove i
                dataSheets.Add dataSheets(j - 1), , i
                dataSheets.Remove j
                dataSheets.Add tempName, , j
            End If
        Next j
    Next i

    ' Populate listbox
    lb.Clear
    For i = 1 To dataSheets.Count
        lb.AddItem dataSheets(i)
    Next i

    If dataSheets.Count = 0 Then
        lb.AddItem "(No data sheets)"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error refreshing sheet list: " & Err.Description, vbExclamation
End Sub

Public Sub RefreshAllLists()
    ' Refresh both list boxes
    RefreshXMLFileList
    RefreshSheetList
End Sub

' =============================================================================
' XML IMPORT
' =============================================================================

Public Sub ImportXML()
    ' Import selected XML file to a new worksheet

    Dim ws As Worksheet
    Dim lb As MSForms.ListBox
    Dim xmlDoc As Object
    Dim xmlPath As String
    Dim selectedFile As String
    Dim newSheetName As String
    Dim newSheet As Worksheet
    Dim paramNodes As Object
    Dim fieldNames As Collection
    Dim orderedFields As Collection
    Dim lockedFields As Collection
    Dim paramData As Collection
    Dim cadObjectData As Object
    Dim i As Long, row As Long, col As Long

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Sheets(1)
    Set lb = ws.OLEObjects("ListBox1").Object

    ' Check selection
    If lb.ListIndex < 0 Then
        MsgBox "Please select an XML file from the list.", vbExclamation
        Exit Sub
    End If

    selectedFile = lb.List(lb.ListIndex)
    If selectedFile = "(No XML files found)" Or selectedFile = "(Save workbook first to see XML files)" Then
        MsgBox "No valid XML file selected.", vbExclamation
        Exit Sub
    End If

    xmlPath = GetLocalPath(ThisWorkbook.Path) & "\" & selectedFile

    ' Load XML
    Set xmlDoc = CreateXMLDocument()
    If xmlDoc Is Nothing Then
        MsgBox "Could not create XML parser. MSXML may not be installed.", vbCritical
        Exit Sub
    End If

    If Not xmlDoc.Load(xmlPath) Then
        MsgBox "Failed to load XML file: " & xmlDoc.parseError.reason, vbCritical
        Exit Sub
    End If

    ' Get all Parameter nodes
    Set paramNodes = xmlDoc.SelectNodes("//Parameter")
    If paramNodes.Length = 0 Then
        MsgBox "No Parameter elements found in XML.", vbExclamation
        Exit Sub
    End If

    ' Detect field names by finding first duplicate
    Set fieldNames = DetectFieldNames(paramNodes)
    If fieldNames.Count = 0 Then
        MsgBox "Could not detect parameter field structure.", vbCritical
        Exit Sub
    End If

    ' Detect locked fields
    Set lockedFields = DetectLockedFields(paramNodes)

    ' Order fields by priority
    Set orderedFields = OrderFieldsByPriority(fieldNames)

    ' Parse parameter data into CAD object groups
    Set paramData = ParseParameterData(paramNodes, fieldNames)

    ' Create new sheet
    newSheetName = CreateSheetName(selectedFile)
    Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    newSheet.Name = newSheetName

    ' Write headers
    For col = 1 To orderedFields.Count
        newSheet.Cells(1, col).Value = orderedFields(col)
        newSheet.Cells(1, col).Font.Bold = True
    Next col

    ' Write data rows
    row = 2
    For i = 1 To paramData.Count
        Set cadObjectData = paramData(i)
        For col = 1 To orderedFields.Count
            If cadObjectData.Exists(orderedFields(col)) Then
                newSheet.Cells(row, col).Value = cadObjectData(orderedFields(col))
            End If
        Next col
        row = row + 1
    Next i

    ' Format sheet
    FormatDataSheet newSheet, orderedFields, lockedFields, paramData.Count + 1

    ' Refresh sheet list
    RefreshSheetList

    MsgBox "Imported " & paramData.Count & " CAD objects to sheet: " & newSheetName, vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Error importing XML: " & Err.Description, vbCritical
End Sub

Private Function DetectFieldNames(paramNodes As Object) As Collection
    ' Detect field names by iterating until first duplicate

    Dim fieldNames As Collection
    Dim node As Object
    Dim paramName As String
    Dim i As Long

    Set fieldNames = New Collection

    For i = 0 To paramNodes.Length - 1
        Set node = paramNodes.Item(i)
        paramName = node.getAttribute("Name")

        ' Check if already seen
        If CollectionContains(fieldNames, paramName) Then
            Exit For ' Found duplicate - we have complete field list
        End If

        fieldNames.Add paramName, paramName
    Next i

    Set DetectFieldNames = fieldNames
End Function

Private Function DetectLockedFields(paramNodes As Object) As Collection
    ' Detect which fields have Access=Locked anywhere in the XML

    Dim lockedFields As Collection
    Dim node As Object
    Dim paramName As String
    Dim accessNode As Object
    Dim i As Long

    Set lockedFields = New Collection

    For i = 0 To paramNodes.Length - 1
        Set node = paramNodes.Item(i)
        paramName = node.getAttribute("Name")

        ' Check for Access element
        Set accessNode = node.SelectSingleNode("Access")
        If Not accessNode Is Nothing Then
            If accessNode.Text = "Locked" Then
                ' Add to locked fields if not already there
                If Not CollectionContains(lockedFields, paramName) Then
                    lockedFields.Add paramName, paramName
                End If
            End If
        End If
    Next i

    Set DetectLockedFields = lockedFields
End Function

Private Function OrderFieldsByPriority(fieldNames As Collection) As Collection
    ' Order fields: priority fields first (in order), then alphabetical

    Dim orderedFields As Collection
    Dim priorityArr() As String
    Dim remainingFields As Collection
    Dim fieldName As Variant
    Dim i As Long, j As Long
    Dim tempName As String

    Set orderedFields = New Collection
    Set remainingFields = New Collection

    priorityArr = Split(PRIORITY_FIELDS, ",")

    ' Add priority fields first (if they exist)
    For i = LBound(priorityArr) To UBound(priorityArr)
        If CollectionContains(fieldNames, priorityArr(i)) Then
            orderedFields.Add priorityArr(i), priorityArr(i)
        End If
    Next i

    ' Collect non-priority fields
    For Each fieldName In fieldNames
        If Not CollectionContains(orderedFields, CStr(fieldName)) Then
            remainingFields.Add CStr(fieldName)
        End If
    Next fieldName

    ' Sort remaining fields alphabetically
    For i = 1 To remainingFields.Count - 1
        For j = i + 1 To remainingFields.Count
            If remainingFields(j) < remainingFields(i) Then
                tempName = remainingFields(i)
                remainingFields.Remove i
                remainingFields.Add tempName, , , i - 1
                remainingFields.Remove i
                remainingFields.Add remainingFields(j - 1), , i
                remainingFields.Remove j
                remainingFields.Add tempName, , j
            End If
        Next j
    Next i

    ' Add sorted remaining fields
    For i = 1 To remainingFields.Count
        orderedFields.Add remainingFields(i)
    Next i

    Set OrderFieldsByPriority = orderedFields
End Function

Private Function ParseParameterData(paramNodes As Object, fieldNames As Collection) As Collection
    ' Parse parameter nodes into collection of dictionaries (one per CAD object)

    Dim paramData As Collection
    Dim cadObject As Object
    Dim node As Object
    Dim paramName As String
    Dim paramValue As String
    Dim valueNode As Object
    Dim fieldCount As Long
    Dim i As Long

    Set paramData = New Collection
    fieldCount = fieldNames.Count

    For i = 0 To paramNodes.Length - 1
        ' Start new CAD object at beginning of each group
        If i Mod fieldCount = 0 Then
            Set cadObject = CreateObject("Scripting.Dictionary")
        End If

        Set node = paramNodes.Item(i)
        paramName = node.getAttribute("Name")

        ' Get value
        Set valueNode = node.SelectSingleNode("Value")
        If Not valueNode Is Nothing Then
            paramValue = valueNode.Text
        Else
            paramValue = ""
        End If

        cadObject(paramName) = paramValue

        ' End of group - add to collection
        If (i + 1) Mod fieldCount = 0 Then
            paramData.Add cadObject
        End If
    Next i

    Set ParseParameterData = paramData
End Function

Private Sub FormatDataSheet(ws As Worksheet, orderedFields As Collection, lockedFields As Collection, rowCount As Long)
    ' Format the data sheet: text format, freeze panes, protect locked columns

    Dim rng As Range
    Dim col As Long
    Dim fieldName As String
    Dim colCount As Long

    colCount = orderedFields.Count

    ' Format all data cells as text
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(rowCount, colCount))
    rng.NumberFormat = "@"

    ' Auto-fit columns
    rng.Columns.AutoFit

    ' Freeze first row and first column
    ws.Activate
    ws.Cells(2, 2).Select
    ActiveWindow.FreezePanes = True

    ' Lock all locked fields and add conditional formatting for missing fields
    Dim hasEmptyCells As Boolean
    Dim r As Long
    Dim cellRng As Range
    Dim fc As FormatCondition

    For col = 1 To colCount
        fieldName = CStr(orderedFields(col))
        If CollectionContains(lockedFields, fieldName) Then
            ws.Columns(col).Locked = True
            ws.Columns(col).Interior.Color = RGB(240, 240, 240) ' Light grey for locked
        Else
            ws.Columns(col).Locked = False

            ' Check if this column has empty cells (field missing in some objects)
            hasEmptyCells = False
            For r = 2 To rowCount
                If Trim(CStr(ws.Cells(r, col).Value)) = "" Then
                    hasEmptyCells = True
                    Exit For
                End If
            Next r

            ' Add conditional formatting for columns with partial field presence
            If hasEmptyCells Then
                Set cellRng = ws.Range(ws.Cells(2, col), ws.Cells(rowCount, col))
                cellRng.FormatConditions.Delete ' Clear existing conditions

                ' Rule 1 (Priority 1): Non-empty cells = pale yellow (parameter present/added)
                Set fc = cellRng.FormatConditions.Add(Type:=xlExpression, Formula1:="=LEN(TRIM(" & cellRng.Cells(1, 1).Address(False, False) & "))>0")
                fc.Interior.Color = RGB(255, 255, 204) ' Pale yellow for present/added fields
                fc.StopIfTrue = False

                ' Rule 2 (Priority 2): Empty cells = grey (parameter missing)
                Set fc = cellRng.FormatConditions.Add(Type:=xlExpression, Formula1:="=LEN(TRIM(" & cellRng.Cells(1, 1).Address(False, False) & "))=0")
                fc.Interior.Color = RGB(240, 240, 240) ' Light grey for missing fields
                fc.StopIfTrue = False
            End If
        End If
    Next col

    ' Protect sheet with locked columns
    ws.Protect Password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True, _
               AllowFormattingColumns:=True, AllowFormattingRows:=True

    ' Header formatting
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(200, 200, 200)
End Sub

Private Function CreateSheetName(xmlFileName As String) As String
    ' Create sheet name: filename-yyyymmdd_hhmmss.xml

    Dim baseName As String
    Dim timestamp As String
    Dim fullName As String

    ' Remove .xml extension if present
    baseName = xmlFileName
    If LCase(Right(baseName, 4)) = ".xml" Then
        baseName = Left(baseName, Len(baseName) - 4)
    End If

    ' Create timestamp
    timestamp = Format(Now, "yyyymmdd_hhmmss")

    ' Combine (Excel sheet names max 31 chars)
    fullName = baseName & "-" & timestamp & ".xml"
    If Len(fullName) > 31 Then
        fullName = Left(baseName, 31 - Len("-" & timestamp & ".xml")) & "-" & timestamp & ".xml"
    End If

    CreateSheetName = fullName
End Function

' =============================================================================
' XML EXPORT
' =============================================================================

Public Sub ExportXML()
    ' Export selected sheet back to XML file

    Dim ws As Worksheet
    Dim lb As MSForms.ListBox
    Dim dataSheet As Worksheet
    Dim selectedSheet As String
    Dim xmlDoc As Object
    Dim rootNode As Object
    Dim paramNode As Object
    Dim childNode As Object
    Dim headers As Collection
    Dim fieldOrder As Collection
    Dim originalXmlPath As String
    Dim outputXmlPath As String
    Dim row As Long, col As Long
    Dim lastRow As Long, lastCol As Long
    Dim fieldName As String
    Dim cellValue As String
    Dim blankWarnings As String
    Dim proceedExport As VbMsgBoxResult

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Sheets(1)
    Set lb = ws.OLEObjects("ListBox2").Object

    ' Check selection
    If lb.ListIndex < 0 Then
        MsgBox "Please select a sheet from the list.", vbExclamation
        Exit Sub
    End If

    selectedSheet = lb.List(lb.ListIndex)
    If selectedSheet = "(No data sheets)" Then
        MsgBox "No valid sheet selected.", vbExclamation
        Exit Sub
    End If

    Set dataSheet = ThisWorkbook.Sheets(selectedSheet)

    ' Get dimensions
    lastRow = dataSheet.Cells(dataSheet.Rows.Count, 1).End(xlUp).row
    lastCol = dataSheet.Cells(1, dataSheet.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        MsgBox "Sheet has no data rows.", vbExclamation
        Exit Sub
    End If

    ' Read headers and determine field order for XML (alphabetical by parameter name)
    ' Also detect which columns are locked
    Set headers = New Collection
    Set fieldOrder = New Collection
    Dim lockedFields As Collection
    Set lockedFields = New Collection

    For col = 1 To lastCol
        fieldName = CStr(dataSheet.Cells(1, col).Value)
        headers.Add fieldName, fieldName

        ' Check if column is locked
        If dataSheet.Columns(col).Locked Then
            lockedFields.Add fieldName, fieldName
        End If
    Next col

    ' Sort headers alphabetically for XML output
    Set fieldOrder = SortCollectionAlphabetically(headers)

    ' Validate data and collect warnings
    blankWarnings = ValidateExportData(dataSheet, headers, lastRow, lastCol)

    If blankWarnings <> "" Then
        proceedExport = MsgBox("The following required fields have blank values:" & vbCrLf & vbCrLf & _
                               blankWarnings & vbCrLf & _
                               "Do you want to continue with export?", _
                               vbYesNo + vbExclamation, "Validation Warning")
        If proceedExport = vbNo Then
            Exit Sub
        End If
    End If

    ' Create XML document
    Set xmlDoc = CreateXMLDocument()
    If xmlDoc Is Nothing Then
        MsgBox "Could not create XML parser. MSXML may not be installed.", vbCritical
        Exit Sub
    End If

    ' Add XML declaration
    Dim xmlDecl As Object
    Set xmlDecl = xmlDoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
    xmlDoc.appendChild xmlDecl

    ' Create root element
    Set rootNode = xmlDoc.createElement("CreoParamSet")
    xmlDoc.appendChild rootNode

    ' Write parameters in alphabetical order within each CAD object group
    For row = 2 To lastRow
        For Each fieldName In fieldOrder
            ' Find column for this field
            col = GetColumnForField(dataSheet, fieldName, lastCol)
            If col > 0 Then
                cellValue = CStr(dataSheet.Cells(row, col).Value)

                ' Create Parameter node
                Set paramNode = xmlDoc.createElement("Parameter")
                paramNode.setAttribute "Name", fieldName

                ' DataType (always String for these parameters)
                Set childNode = xmlDoc.createElement("DataType")
                childNode.Text = "String"
                paramNode.appendChild childNode

                ' Value
                Set childNode = xmlDoc.createElement("Value")
                childNode.Text = cellValue
                paramNode.appendChild childNode

                ' Add Access=Locked for locked fields
                If CollectionContains(lockedFields, fieldName) Then
                    Set childNode = xmlDoc.createElement("Access")
                    childNode.Text = "Locked"
                    paramNode.appendChild childNode
                End If

                rootNode.appendChild paramNode
            End If
        Next fieldName
    Next row

    ' Save XML file (same name as sheet)
    outputXmlPath = GetLocalPath(ThisWorkbook.Path) & "\" & selectedSheet
    If LCase(Right(outputXmlPath, 4)) <> ".xml" Then
        outputXmlPath = outputXmlPath & ".xml"
    End If

    xmlDoc.Save outputXmlPath

    ' Refresh XML list
    RefreshXMLFileList

    MsgBox "Exported " & (lastRow - 1) & " CAD objects to: " & vbCrLf & outputXmlPath, vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Error exporting XML: " & Err.Description, vbCritical
End Sub

Private Function ValidateExportData(ws As Worksheet, headers As Collection, lastRow As Long, lastCol As Long) As String
    ' Validate data and return warning string for blank required fields

    Dim warnings As String
    Dim requiredArr() As String
    Dim row As Long, col As Long
    Dim fieldName As String
    Dim cellValue As String
    Dim blankCount As Long
    Dim i As Long

    warnings = ""
    requiredArr = Split(REQUIRED_FIELDS, ",")

    For i = LBound(requiredArr) To UBound(requiredArr)
        fieldName = requiredArr(i)
        col = GetColumnForField(ws, fieldName, lastCol)

        If col > 0 Then
            blankCount = 0
            For row = 2 To lastRow
                cellValue = Trim(CStr(ws.Cells(row, col).Value))
                If cellValue = "" Then
                    blankCount = blankCount + 1
                End If
            Next row

            If blankCount > 0 Then
                warnings = warnings & "  - " & fieldName & ": " & blankCount & " blank value(s)" & vbCrLf
            End If
        End If
    Next i

    ValidateExportData = warnings
End Function

Private Function GetColumnForField(ws As Worksheet, fieldName As String, lastCol As Long) As Long
    ' Find column number for a field name in row 1

    Dim col As Long

    For col = 1 To lastCol
        If CStr(ws.Cells(1, col).Value) = fieldName Then
            GetColumnForField = col
            Exit Function
        End If
    Next col

    GetColumnForField = 0
End Function

Private Function SortCollectionAlphabetically(col As Collection) As Collection
    ' Sort a collection of strings alphabetically

    Dim sorted As Collection
    Dim arr() As String
    Dim i As Long, j As Long
    Dim tempStr As String

    Set sorted = New Collection

    ' Convert to array for sorting
    ReDim arr(1 To col.Count)
    For i = 1 To col.Count
        arr(i) = col(i)
    Next i

    ' Bubble sort
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                tempStr = arr(i)
                arr(i) = arr(j)
                arr(j) = tempStr
            End If
        Next j
    Next i

    ' Convert back to collection
    For i = LBound(arr) To UBound(arr)
        sorted.Add arr(i)
    Next i

    Set SortCollectionAlphabetically = sorted
End Function

' =============================================================================
' UTILITY FUNCTIONS
' =============================================================================

Private Function CollectionContains(col As Collection, key As String) As Boolean
    ' Check if a collection contains a key

    Dim item As Variant

    On Error Resume Next
    item = col(key)
    CollectionContains = (Err.Number = 0)
    On Error GoTo 0
End Function

' =============================================================================
' BUTTON CLICK HANDLERS
' =============================================================================

Public Sub btnImport_Click()
    ImportXML
End Sub

Public Sub btnExport_Click()
    ExportXML
End Sub

Public Sub btnRefresh_Click()
    RefreshAllLists
End Sub
