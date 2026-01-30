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
    Dim workbookPath As String
    Dim i As Long, j As Long
    Dim fileCount As Long
    Dim fileNames() As String
    Dim fileDates() As Date
    Dim tempName As String
    Dim tempDate As Date

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Sheets(1)
    Set lb = ws.OLEObjects("ListBox1").Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    workbookPath = ThisWorkbook.Path
    If workbookPath = "" Then
        lb.Clear
        lb.AddItem "(Save workbook first to see XML files)"
        Exit Sub
    End If

    ' Convert SharePoint/OneDrive URL to local path
    workbookPath = GetLocalPath(workbookPath)

    Set folder = fso.GetFolder(workbookPath)

    ' Count XML files first
    fileCount = 0
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "xml" Then
            fileCount = fileCount + 1
        End If
    Next file

    ' Handle case with no files
    lb.Clear
    If fileCount = 0 Then
        lb.AddItem "(No XML files found)"
        Exit Sub
    End If

    ' Collect all XML files into arrays
    ReDim fileNames(1 To fileCount)
    ReDim fileDates(1 To fileCount)
    i = 1
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "xml" Then
            fileNames(i) = file.Name
            fileDates(i) = file.DateLastModified
            i = i + 1
        End If
    Next file

    ' Sort by date descending (bubble sort)
    For i = 1 To fileCount - 1
        For j = i + 1 To fileCount
            If fileDates(j) > fileDates(i) Then
                ' Swap names
                tempName = fileNames(i)
                fileNames(i) = fileNames(j)
                fileNames(j) = tempName
                ' Swap dates
                tempDate = fileDates(i)
                fileDates(i) = fileDates(j)
                fileDates(j) = tempDate
            End If
        Next j
    Next i

    ' Populate listbox
    For i = 1 To fileCount
        lb.AddItem fileNames(i)
    Next i

    Exit Sub

ErrorHandler:
    MsgBox "Error refreshing XML file list: " & Err.Description, vbExclamation
End Sub

Public Sub RefreshSheetList()
    ' Populates ListBox2 with data sheets (excluding Manager), sorted newest first

    Dim ws As Worksheet
    Dim lb As MSForms.ListBox
    Dim sheet As Worksheet
    Dim i As Long, j As Long
    Dim sheetCount As Long
    Dim sheetNames() As String
    Dim tempName As String

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Sheets(1)
    Set lb = ws.OLEObjects("ListBox2").Object

    ' Count data sheets (all except first sheet)
    sheetCount = 0
    For Each sheet In ThisWorkbook.Sheets
        If sheet.Index > 1 Then
            sheetCount = sheetCount + 1
        End If
    Next sheet

    ' Handle case with no data sheets
    lb.Clear
    If sheetCount = 0 Then
        lb.AddItem "(No data sheets)"
        Exit Sub
    End If

    ' Collect data sheets into array
    ReDim sheetNames(1 To sheetCount)
    i = 1
    For Each sheet In ThisWorkbook.Sheets
        If sheet.Index > 1 Then
            sheetNames(i) = sheet.Name
            i = i + 1
        End If
    Next sheet

    ' Sort alphabetically descending (timestamps in name = newest first)
    For i = 1 To sheetCount - 1
        For j = i + 1 To sheetCount
            If sheetNames(j) > sheetNames(i) Then
                tempName = sheetNames(i)
                sheetNames(i) = sheetNames(j)
                sheetNames(j) = tempName
            End If
        Next j
    Next i

    ' Populate listbox
    For i = 1 To sheetCount
        lb.AddItem sheetNames(i)
    Next i

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

    ' Write data rows (starting at row 3, row 2 will be marker row)
    row = 3
    For i = 1 To paramData.Count
        Set cadObjectData = paramData(i)
        For col = 1 To orderedFields.Count
            If cadObjectData.Exists(orderedFields(col)) Then
                newSheet.Cells(row, col).Value = cadObjectData(orderedFields(col))
            End If
        Next col
        row = row + 1
    Next i

    ' Add marker row (row 2) to track which fields are partial
    ' "F" = Full field (all objects have it), "P" = Partial field (some have it)
    Dim hasEmpty As Boolean, hasFilled As Boolean
    For col = 1 To orderedFields.Count
        hasEmpty = False
        hasFilled = False
        For row = 3 To paramData.Count + 2
            If Trim(CStr(newSheet.Cells(row, col).Value)) = "" Then
                hasEmpty = True
            Else
                hasFilled = True
            End If
        Next row

        ' Mark column type
        If hasEmpty And hasFilled Then
            newSheet.Cells(2, col).Value = "P" ' Partial field
        Else
            newSheet.Cells(2, col).Value = "F" ' Full field
        End If
    Next col

    ' Format sheet
    FormatDataSheet newSheet, orderedFields, lockedFields, paramData.Count + 2 ' +2 for header and marker row

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
    ' Format the data sheet with comprehensive conditional formatting and styling

    Dim rng As Range
    Dim col As Long, row As Long
    Dim fieldName As String
    Dim colCount As Long
    Dim isStandardField As Boolean
    Dim isPartialField As Boolean
    Dim priorityArr() As String
    Dim cellRng As Range
    Dim fc As FormatCondition
    Dim dataStartRow As Long

    colCount = orderedFields.Count
    dataStartRow = 3 ' Data starts at row 3 (row 1 = header, row 2 = marker)
    priorityArr = Split(PRIORITY_FIELDS, ",")

    ' Format all cells as text
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(rowCount, colCount))
    rng.NumberFormat = "@"

    ' === DIRECT FORMATTING (applied first, can be overridden by conditional) ===

    ' Default white fill for all data cells
    Set rng = ws.Range(ws.Cells(dataStartRow, 1), ws.Cells(rowCount, colCount))
    rng.Interior.Color = RGB(255, 255, 255)

    ' First column: Bold text, grey fill
    ws.Columns(1).Font.Bold = True
    Set rng = ws.Range(ws.Cells(dataStartRow, 1), ws.Cells(rowCount, 1))
    rng.Interior.Color = RGB(200, 200, 200)

    ' First row: Bold (already done in import)
    ws.Rows(1).Font.Bold = True

    ' Borders: Left and right on columns, all borders on data cells
    For col = 1 To colCount
        Set rng = ws.Range(ws.Cells(dataStartRow, col), ws.Cells(rowCount, col))
        ' Left and right borders for column
        With rng.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With rng.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        ' All borders for each cell
        With rng.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next col

    ' === CONDITIONAL FORMATTING ===

    For col = 1 To colCount
        fieldName = CStr(orderedFields(col))

        ' Check if this is a standard (priority) field
        isStandardField = CollectionContains(lockedFields, fieldName) Or _
                          InStrInArray(fieldName, priorityArr) >= 0

        ' Check if this is a partial field (from marker row)
        isPartialField = (ws.Cells(2, col).Value = "P")

        ' Set up cell range for data rows
        Set cellRng = ws.Range(ws.Cells(dataStartRow, col), ws.Cells(rowCount, col))
        cellRng.FormatConditions.Delete ' Clear existing conditions

        If isStandardField And Not isPartialField Then
            ' === STANDARD FIELD (full presence) ===
            ' Green for filled, red for blank

            ' Rule 1: Non-empty = light green
            Set fc = cellRng.FormatConditions.Add(Type:=xlExpression, _
                Formula1:="=LEN(TRIM(" & cellRng.Cells(1, 1).Address(False, False) & "))>0")
            fc.Interior.Color = RGB(144, 238, 144) ' Light green
            fc.StopIfTrue = False

            ' Rule 2: Empty = light red
            Set fc = cellRng.FormatConditions.Add(Type:=xlExpression, _
                Formula1:="=LEN(TRIM(" & cellRng.Cells(1, 1).Address(False, False) & "))=0")
            fc.Interior.Color = RGB(255, 204, 203) ' Light red
            fc.StopIfTrue = False

        ElseIf isStandardField And isPartialField Then
            ' === STANDARD FIELD with partial presence (shouldn't happen but handle it) ===
            ' Non-empty = green, empty = dark red with bold

            ' Rule 1: Non-empty = light green
            Set fc = cellRng.FormatConditions.Add(Type:=xlExpression, _
                Formula1:="=LEN(TRIM(" & cellRng.Cells(1, 1).Address(False, False) & "))>0")
            fc.Interior.Color = RGB(144, 238, 144) ' Light green
            fc.StopIfTrue = False

            ' Rule 2: Empty = dark red, bold
            Set fc = cellRng.FormatConditions.Add(Type:=xlExpression, _
                Formula1:="=LEN(TRIM(" & cellRng.Cells(1, 1).Address(False, False) & "))=0")
            fc.Interior.Color = RGB(255, 0, 0) ' Dark red
            fc.Font.Bold = True
            fc.StopIfTrue = False

        ElseIf Not isStandardField And isPartialField Then
            ' === ADDITIONAL FIELD with partial presence ===
            ' Filled originally = light blue, empty = light grey, user-added = light yellow

            ' For partial additional fields, use direct light blue for cells with original data
            For row = dataStartRow To rowCount
                If Trim(CStr(ws.Cells(row, col).Value)) <> "" Then
                    ws.Cells(row, col).Interior.Color = RGB(173, 216, 230) ' Light blue
                End If
            Next row

            ' Rule 1: Empty = light grey
            Set fc = cellRng.FormatConditions.Add(Type:=xlExpression, _
                Formula1:="=LEN(TRIM(" & cellRng.Cells(1, 1).Address(False, False) & "))=0")
            fc.Interior.Color = RGB(240, 240, 240) ' Light grey
            fc.StopIfTrue = False

            ' Rule 2: Non-empty = light yellow (will show for user-added, overridden by direct blue for original)
            ' Actually, we can't distinguish user-added from original with conditional formatting alone
            ' The light blue direct format will remain for original data
            ' When user types in an empty (grey) cell, they should manually change color or we need event handler

        Else
            ' === ADDITIONAL FIELD (full presence) ===
            ' Same as standard fields
            ' Rule 1: Non-empty = light green
            Set fc = cellRng.FormatConditions.Add(Type:=xlExpression, _
                Formula1:="=LEN(TRIM(" & cellRng.Cells(1, 1).Address(False, False) & "))>0")
            fc.Interior.Color = RGB(144, 238, 144) ' Light green
            fc.StopIfTrue = False

            ' Rule 2: Empty = light red
            Set fc = cellRng.FormatConditions.Add(Type:=xlExpression, _
                Formula1:="=LEN(TRIM(" & cellRng.Cells(1, 1).Address(False, False) & "))=0")
            fc.Interior.Color = RGB(255, 204, 203) ' Light red
            fc.StopIfTrue = False
        End If

        ' Lock columns with locked fields
        If CollectionContains(lockedFields, fieldName) Then
            ws.Columns(col).Locked = True
        Else
            ws.Columns(col).Locked = False
        End If
    Next col

    ' Hide marker row
    ws.Rows(2).Hidden = True

    ' Auto-fit columns
    ws.Cells.Columns.AutoFit

    ' Freeze first row (header) and first column
    ws.Activate
    ws.Cells(dataStartRow, 2).Select
    ActiveWindow.FreezePanes = True

    ' Protect sheet with locked columns
    ws.Protect Password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True, _
               AllowFormattingColumns:=True, AllowFormattingRows:=True
End Sub

Private Function InStrInArray(searchFor As String, arr() As String) As Long
    ' Returns index of string in array, or -1 if not found
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = searchFor Then
            InStrInArray = i
            Exit Function
        End If
    Next i
    InStrInArray = -1
End Function

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
