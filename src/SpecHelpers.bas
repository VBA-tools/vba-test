Attribute VB_Name = "SpecHelpers"
''
' SpecHelpers v1.1.0
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
' @param {String} RangeName
' @param {String} [WB] Workbook to check or active workbook
' @returns {Integer} Index of sheet that named range is found on or -1
' --------------------------------------------- '

Public Function NamedRangeExists(RangeName As String, Optional WB As Workbook) As Integer
    Dim rngTest As Range, i As Long
     
    If WB Is Nothing Then: Set WB = ActiveWorkbook
    With WB
        On Error Resume Next
        ' Loop through all sheets in workbook. In VBA, you MUST specify
        ' the worksheet name which the named range is found on. Using
        ' Named Ranges in worksheet functions DO work across sheets
        ' without explicit reference.
        For i = 1 To .Sheets.Count Step 1
            ' Try to set our variable as the named range.
            Set rngTest = .Sheets(i).Range(RangeName)
             
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
' @returns {Boolean}
' --------------------------------------------- '

Public Function SheetExists(SheetName As String, Optional WB As Workbook) As Boolean
    Dim Sheet As Worksheet
    
    If WB Is Nothing Then: Set WB = ActiveWorkbook
    If Not WB Is Nothing Then
        For Each Sheet In WB.Sheets
            If Sheet.Name = SheetName Then
                SheetExists = True
                Exit Function
            End If
        Next Sheet
    End If
End Function

''
' Check if sheet is visible in current workbook
'
' @param {String} sheetName
' @param {Workbook} [WB] Workbook to check or active workbook
' @returns {Boolean}
' --------------------------------------------- '

Public Function SheetIsVisible(SheetName As String, Optional WB As Workbook) As Boolean
    
    If WB Is Nothing Then: Set WB = ActiveWorkbook
    If SheetExists(SheetName, WB) Then
        Dim Sheet As Worksheet
        Set Sheet = WB.Sheets(SheetName)
        
        Select Case WB.Sheets(SheetName).Visible
        Case XlSheetVisibility.xlSheetVisible: SheetIsVisible = True
        End Select
    End If
End Function

''
' Check if workbook is protected
'
' @param {Workbook} [WB] Workbook to check or active workbook
' @returns {Boolean}
' --------------------------------------------- '

Public Function WBIsProtected(Optional WB As Workbook) As Boolean
    
    If WB Is Nothing Then: Set WB = ActiveWorkbook
    If WB.ProtectWindows Then WBIsProtected = True
    If WB.ProtectStructure Then WBIsProtected = True
End Function

''
' Check if sheet is protected
'
' @param {String} sheetName
' @param {Workbook} [WB] Workbook to check or active workbook
' @returns {Boolean}
' --------------------------------------------- '

Public Function SheetIsProtected(SheetName As String, Optional WB As Workbook) As Boolean
    
    If WB Is Nothing Then: Set WB = ActiveWorkbook
    If WB.Sheets(SheetName).ProtectContents Then SheetIsProtected = True
    If WB.Sheets(SheetName).ProtectDrawingObjects Then SheetIsProtected = True
    If WB.Sheets(SheetName).ProtectScenarios Then SheetIsProtected = True
End Function

''
' Check if file exists
'
' @param {String} filePath
' @returns {Boolean}
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
' @returns {Dictionary}
' --------------------------------------------- '

Public Function SheetCell(SheetName As String, Row As Integer, Col As Integer) As Dictionary
    Set SheetCell = New Dictionary
    SheetCell.Add "sheetName", SheetName
    SheetCell.Add "row", Row
    SheetCell.Add "col", Col
End Function

''
' Combine collections
'
' @param {Collection} collection1
' @param {Collection} collection2
' @returns {Collection}
' --------------------------------------------- '

Public Function CombineCollections(collection1 As Collection, collection2 As Collection) As Collection
    Dim combined As New Collection
    Dim Value As Variant
    
    For Each Value In collection1
        combined.Add Value
    Next Value
    For Each Value In collection2
        combined.Add Value
    Next Value
    
    Set CombineCollections = combined
End Function

''
' Get last row for sheet
'
' @param {Worksheet} sheet
' @returns {Integer}
' --------------------------------------------- '

Public Function LastRow(Sheet As Worksheet) As Integer
    Dim NumRows As Integer
    NumRows = Sheet.UsedRange.Rows.Count
    LastRow = Sheet.UsedRange.Rows(NumRows).Row
End Function
