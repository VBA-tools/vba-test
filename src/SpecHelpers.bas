Attribute VB_Name = "SpecHelpers"
''
' SpecHelpers v1.2.1
' (c) Tim Hall - https://github.com/timhall/Excel-TDD
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
    SheetCell.Add "SheetName", SheetName
    SheetCell.Add "Row", Row
    SheetCell.Add "Col", Col
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

''
' Check if workbook is open
'
' @param {String} Path
' @returns {Boolean}
' --------------------------------------------- '

Public Function WorkbookIsOpen(Path As String) As Boolean
    On Error Resume Next
    Dim WB As Workbook
    Set WB = Application.Workbooks(Filename)
    On Error GoTo 0
    
    ' If failed to load already open workbook, open it
    If Err.Number = 0 Then
        WorkbookIsOpen = True
    End If
    
    Set WB = Nothing
    Err.Clear
End Function

''
' Toggle screen updating and return previous updating value
'
' @param {Boolean} [Updating=False]
' @param {Boolean} [ToggleEvents=True]
'
' Example:
' Dim PrevUpdating As Boolean
' PrevUpdating = SpecHelpers.ToggleUpdating()
'
' ... Do screen-intensive stuff
'
' ' Restore previous updating status after hard work
' ToggleUpdating PrevUpdating
'
' --------------------------------------- '
Public Function ToggleUpdating(Optional Updating As Boolean = False, Optional ToggleEvents As Boolean = True) As Boolean
    ToggleUpdating = Application.ScreenUpdating
    
    Application.ScreenUpdating = Updating
    If Updating Or Events Then
        Application.EnableEvents = Updating
    End If
End Function

''
' Run scenario using given scenario, sheet name, and IWBProxy
'
' @param {IScenario} Scenario
' @param {IWBProxy} WB to use for scenario
' @param {String} SheetName to load scenario from
' --------------------------------------------- '

Public Function RunScenario(Scenario As IScenario, WB As IWBProxy, SheetName As String) As SpecSuite
    If SpecHelpers.SheetExists(SheetName, ThisWorkbook) Then
        Scenario.Load SheetName
        Set RunScenario = Scenario.RunScenario(WB)
    Else
        MsgBox "Warning" & vbNewLine & "No sheet was found for the following scenario: " & SheetName, Title:="Scenario sheet not found"
    End If
End Function

''
' Run scenarios using given scenario, sheet name, and IWBProxy
'
' @param {IScenario} Scenario
' @param {IWBProxy} WB to use for scenario
' @param {String} ... Pass scenario sheet names as additional arguments
'
' Example:
' RunScenarios(Scenario, WB, "Scenario 1", "Scenario 2", "Scenario 3")
' --------------------------------------------- '

Public Function RunScenarios(Scenario As IScenario, WB As IWBProxy, ParamArray SheetNames() As Variant) As Collection
    Dim i As Integer
    Dim SheetName As String
    Dim Spec As SpecSuite
    Set RunScenarios = New Collection
    
    For i = LBound(SheetNames) To UBound(SheetNames)
        SheetName = SheetNames(i)
        Set Spec = SpecHelpers.RunScenario(Scenario, WB, SheetName)
        
        If Not Spec Is Nothing Then
            RunScenarios.Add Spec
        End If
    Next i
End Function

''
' Run scenarios using given scenario and IWBProxy by matcher
'
' @param {IScenario} Scenario
' @param {IWBProxy} WB to use for scenario
' @param {String} Matcher to compare all sheet names to
' @param {Boolean} [MatchCase=False]
'
' Example:
' RunScenarios(Scenario, WB, "Scenario")
' Sheet Names: Spec Runner, Mapping, Scenario 1, and Advanced Scenario
' -> Runs scenarios for Scenario 1 and Advanced Scenario
' --------------------------------------------- '

Public Function RunScenariosByMatcher(Scenario As IScenario, WB As IWBProxy, Matcher As String, _
    Optional MatchCase As Boolean = False, Optional IgnoreBlank As Boolean = True) As Collection
    
    Set RunScenariosByMatcher = New Collection
    
    Dim Sheet As Worksheet
    For Each Sheet In ThisWorkbook.Sheets
        If Sheet.Name = "Blank Scenario" Then
            If Not IgnoreBlank Then
                RunScenariosByMatcher.Add SpecHelpers.RunScenario(Scenario, WB, Sheet.Name)
            End If
        ElseIf MatchCase Then
            If InStr(Sheet.Name, Matcher) Then
                RunScenariosByMatcher.Add SpecHelpers.RunScenario(Scenario, WB, Sheet.Name)
            End If
        Else
            If InStr(UCase(Sheet.Name), UCase(Matcher)) Then
                RunScenariosByMatcher.Add SpecHelpers.RunScenario(Scenario, WB, Sheet.Name)
            End If
        End If
    Next Sheet
End Function

''
' Get value from workbook for provided mapping and key
'
' @param {Workbook} WB
' @param {Dictionary} Mapping
' @param {String} Key
' @returns {Variant} Value from workbook
' --------------------------------------------- '

Public Function GetValue(WB As Workbook, Mapping As Dictionary, Key As String) As Variant
    Dim RangeRef As Range
    
    Set RangeRef = GetRange(WB, Mapping, Key)
    If Not RangeRef Is Nothing Then
        GetValue = RangeRef.Value
    End If
End Function

''
' Set value in workbook for provided mapping and key
'
' @param {Workbook} WB
' @param {Dictionary} Mapping
' @param {String} Key
' @param {Variant} Value
' --------------------------------------------- '

Public Function SetValue(WB As Workbook, Mapping As Dictionary, Key As String, Value As Variant)
    Dim RangeRef As Range
    
    Set RangeRef = GetRange(WB, Mapping, Key)
    If Not RangeRef Is Nothing Then
        RangeRef.Value = Value
    End If
End Function

''
' Get reference to range from workbook for provided mapping and key
'
' @param {Workbook} WB
' @param {Dictionary} Mapping
' @param {String} Key
' @returns {Range} Range from workbook
' --------------------------------------------- '

Public Function GetRange(WB As Workbook, Mapping As Dictionary, Key As String) As Range
    Dim MappingValue As Dictionary
    Dim NamedRangeSheetIndex As Integer
    
    If Mapping.Exists(Key) Then
        ' If mapping contains entry for key, use it to find range
        Set MappingValue = Mapping.Item(Key)
        Set GetRange = WB.Sheets(MappingValue("SheetName")) _
            .Cells(MappingValue("Row"), MappingValue("Col"))
    Else
        ' Check for named range matching mapping key
        NamedRangeSheetIndex = SpecHelpers.NamedRangeExists(Key, WB)
        If NamedRangeSheetIndex > 0 Then
            Set GetRange = WB.Sheets(NamedRangeSheetIndex).Range(Key)
        End If
    End If
End Function

''
' Set range in workbook for provided mapping and key
'
' @param {Workbook} WB
' @param {Dictionary} Mapping
' @param {String} Key
' @param {Variant} Value
' --------------------------------------------- '

Public Function SetRange(WB As Workbook, Mapping As Dictionary, Key As String, Value As Range)
    Dim RangeRef As Range
    
    Set RangeRef = GetRange(WB, Mapping, Key)
    If Not IsEmpty(RangeRef) Then
        Set RangeRef = Value
    End If
End Function

''
' Open the workbook specified in the workbook proxy
' (Opens a temporary copy if the workbook is currently open)
'
' @param {Variant} WBOrInArray IWBProxy directly or in array
' --------------------------------------------- '

Public Sub OpenIWBProxy(WBOrInArray As Variant)
    Dim WB As IWBProxy
    
    If TypeOf WBOrInArray Is IWBProxy Then
        Set WB = WBOrInArray
    Else
        Set WB = WBOrInArray(0)
    End If

    ' TODO temporary copy
    Dim PrevUpdating As Boolean
    PrevUpdating = SpecHelpers.ToggleUpdating
    
    If WB.Path <> "" Then
        Set WB.Instance = Workbooks.Open(WB.Path, UpdateLinks:=False, Password:=WB.Password)
    Else
        Err.Raise vbObjectError + 1, "Specs", "Error: No workbook path defined"
    End If
    
    SpecHelpers.ToggleUpdating PrevUpdating
End Sub

''
' Close the workbook specified in the workbook proxy
'
' @param {Variant} WBOrInArray IWBProxy directly or in array
' --------------------------------------------- '

Public Sub CloseIWBProxy(WBOrInArray As Variant)
    Dim WB As IWBProxy
    
    If TypeOf WBOrInArray Is IWBProxy Then
        Set WB = WBOrInArray
    Else
        Set WB = WBOrInArray(0)
    End If

    If Not WB.Instance Is Nothing Then
        WB.Instance.Close False
        Set WB.Instance = Nothing
    End If
End Sub

''
' Close and reopen the workbook specified in the workbook proxy
'
' @param {Variant} WBOrInArray IWBProxy directly or in array
' --------------------------------------------- '

Public Sub ReloadIWBProxy(WBOrInArray As Variant)
    Dim WB As IWBProxy
    
    If TypeOf WBOrInArray Is IWBProxy Then
        Set WB = WBOrInArray
    Else
        Set WB = WBOrInArray(0)
    End If

    SpecHelpers.CloseIWBProxy WB
    SpecHelpers.OpenIWBProxy WB
End Sub
