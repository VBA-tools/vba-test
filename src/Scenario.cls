VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Scenario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Scenario v1.0.0
' (c) Tim Hall - https://github.com/timhall/Excel-TDD
'
' Generic implementation of scenario
'
' @dependencies
'   Microsoft Scripting Runtime
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IScenario

Private inputValues As Dictionary
Private expectedValues As Dictionary

''
' Generate a spec suite to run through a scenario for the given wb
'
' @param {IWBProxy} wb to perform scenario on
' @param {String} name of scenario
' @return {SpecSuite}
' --------------------------------------------- '

Private Function IScenario_RunScenario(wb As IWBProxy, Optional name As String = "") As SpecSuite
    
    Dim specs As New SpecSuite
    Dim keyValue As Variant
    
    specs.wb = wb
    
    With specs.It("Should pass scenario")
        ' Load input values
        For Each keyValue In inputValues.Keys
            specs.wb.value(CStr(keyValue)) = inputValues.Item(CStr(keyValue))
        Next keyValue
        
        ' Run expectations
        For Each keyValue In expectedValues.Keys
            .Expect(specs.wb.value(CStr(keyValue))).ToEqual expectedValues.Item(CStr(keyValue))
        Next keyValue
    End With
    
    specs.wb.CloseWB
    Set IScenario_RunScenario = specs
End Function

''
' Load the scenario from a sheet
'
' @param {String} sheetName
' --------------------------------------------- '

Private Sub IScenario_Load(sheetName As String)
    
    Dim inputRow As Integer
    Dim inputCol As Integer
    Dim expectedRow As Integer
    Dim expectedCol As Integer
    
    inputRow = 4
    inputCol = 1
    expectedRow = 4
    expectedCol = 3
    
    Set inputValues = New Dictionary
    Set expectedValues = New Dictionary
    
    If Utils.SheetExists(sheetName, ThisWorkbook) Then
        
        Do While ThisWorkbook.Sheets(sheetName).Cells(inputRow, inputCol) <> ""
            inputValues.Add _
                CStr(ThisWorkbook.Sheets(sheetName).Cells(inputRow, inputCol)), _
                ThisWorkbook.Sheets(sheetName).Cells(inputRow, inputCol + 1).value
            
            inputRow = inputRow + 1
        Loop
        
        Do While ThisWorkbook.Sheets(sheetName).Cells(expectedRow, expectedCol) <> ""
            expectedValues.Add _
                CStr(ThisWorkbook.Sheets(sheetName).Cells(expectedRow, expectedCol)), _
                ThisWorkbook.Sheets(sheetName).Cells(expectedRow, expectedCol + 1).value
            
            expectedRow = expectedRow + 1
        Loop
        
    End If
    
End Sub
