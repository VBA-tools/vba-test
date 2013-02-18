Attribute VB_Name = "Utils"
''
' Utils v1.0.0
' (c) Tim Hall - https://github.com/timhall/ExcelHelpers
'
' General utilities for specs
'
' @dependencies
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

''
' Check if named range exists and return sheet index if it does
'
' @param {String} rangeName
' @param {String} [WB] Workbook to check or active workbook
' @return {Integer} Index of sheet that named range is found on or -1
' --------------------------------------------- '

Public Function NamedRangeExists(rangeName As String, Optional wb As Workbook) As Integer
    Dim rngTest As Range, i As Long
     
    If wb Is Nothing Then: Set wb = ActiveWorkbook
    With wb
        On Error Resume Next
        ' Loop through all sheets in workbook. In VBA, you MUST specify
        ' the worksheet name which the named range is found on. Using
        ' Named Ranges in worksheet functions DO work across sheets
        ' without explicit reference.
        For i = 1 To .Sheets.Count Step 1
            ' Try to set our variable as the named range.
            Set rngTest = .Sheets(i).Range(rangeName)
             
            ' If there is no error then the name exists.
            If Err = 0 Then
                ' Set the function to TRUE & exit
                NamedRangeExists = i
                Exit Function
            Else
                ' Clear the error and keep trying
                Err.Clear
            End If
        Next i
    End With
    
    ' No range found, return -1
    NamedRangeExists = -1
End Function

''
' Check if sheet exists in current workbook
'
' @param {String} sheetName
' @param {Workbook} [WB] Workbook to check or active workbook
' @return {Boolean}
' --------------------------------------------- '

Public Function SheetExists(sheetName As String, Optional wb As Workbook) As Boolean
    Dim sheet As Worksheet
    
    If wb Is Nothing Then: Set wb = ActiveWorkbook
    If Not wb Is Nothing Then
        For Each sheet In wb.Sheets
            If sheet.name = sheetName Then
                SheetExists = True
                Exit Function
            End If
        Next sheet
    End If
End Function

''
' Check if sheet is visible in current workbook
'
' @param {String} sheetName
' @param {Workbook} [WB] Workbook to check or active workbook
' @return {Boolean}
' --------------------------------------------- '

Public Function SheetIsVisible(sheetName As String, Optional wb As Workbook) As Boolean
    
    If wb Is Nothing Then: Set wb = ActiveWorkbook
    If SheetExists(sheetName, wb) Then
        Dim sheet As Worksheet
        Set sheet = wb.Sheets(sheetName)
        
        Select Case wb.Sheets(sheetName).Visible
        Case XlSheetVisibility.xlSheetVisible: SheetIsVisible = True
        End Select
    End If
End Function

''
' Check if workbook is protected
'
' @param {Workbook} [WB] Workbook to check or active workbook
' @return {Boolean}
' --------------------------------------------- '

Public Function WBIsProtected(Optional wb As Workbook) As Boolean
    
    If wb Is Nothing Then: Set wb = ActiveWorkbook
    If wb.ProtectWindows Then WBIsProtected = True
    If wb.ProtectStructure Then WBIsProtected = True
End Function

''
' Check if sheet is protected
'
' @param {String} sheetName
' @param {Workbook} [WB] Workbook to check or active workbook
' @return {Boolean}
' --------------------------------------------- '
Public Function SheetIsProtected(sheetName As String, Optional wb As Workbook) As Boolean
    
    If wb Is Nothing Then: Set wb = ActiveWorkbook
    If Me.wb.Sheets(sheetName).ProtectContents Then SheetIsProtected = True
    If Me.wb.Sheets(sheetName).ProtectDrawingObjects Then SheetIsProtected = True
    If Me.wb.Sheets(sheetName).ProtectScenarios Then SheetIsProtected = True
End Function

''
' Check if file exists
'
' @param {String} filePath
' @return {Boolean}
' --------------------------------------------- '
Public Function FileExists(filePath As String) As Boolean
    On Error GoTo ErrorHandling
    If Not Dir(filePath, vbDirectory) = vbNullString Then FileExists = True
    
ErrorHandling:
    On Error GoTo 0
End Function

''
' Create SheetCell helper
'
' @param {String} sheetName
' @param {Integer} row
' @param {Integer} col
' @return {Dictionary}
' --------------------------------------------- '
Public Function SheetCell(sheetName As String, row As Integer, col As Integer) As Dictionary
    Set SheetCell = New Dictionary
    SheetCell.Add "sheetName", sheetName
    SheetCell.Add "row", row
    SheetCell.Add "col", col
End Function

''
' Combine collections
'
' @param {Collection} collection1
' @param {Collection} collection2
' @return {Collection}
' --------------------------------------------- '
Public Function CombineCollections(collection1 As Collection, collection2 As Collection) As Collection
    Dim combined As New Collection
    Dim value As Variant
    
    For Each value In collection1
        combined.Add value
    Next value
    For Each value In collection2
        combined.Add value
    Next value
    
    Set CombineCollections = combined
End Function

''
' Get last row for sheet
'
' @param {Worksheet} sheet
' @return {Integer}
' --------------------------------------------- '
Public Function lastRow(sheet As Worksheet) As Integer
    Dim numRows As Integer
    numRows = sheet.UsedRange.Rows.Count
    lastRow = sheet.UsedRange.Rows(numRows).row
End Function
