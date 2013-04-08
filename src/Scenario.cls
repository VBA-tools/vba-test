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
' Scenario v1.1.0
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

Private Const ScenarioInputStartRow As Integer = 4
Private Const ScenarioInputStartCol As Integer = 1
Private Const ScenarioExpectedStartRow As Integer = 4
Private Const ScenarioExpectedStartCol As Integer = 3

Private InputValues As Dictionary
Private ExpectedValues As Dictionary

''
' Generate a spec suite to run through a scenario for the given wb
'
' @param {IWBProxy} WB to perform scenario on
' @param {String} Name of scenario
' @returns {SpecSuite}
' --------------------------------------------- '

Private Function IScenario_RunScenario(WB As IWBProxy, Optional Name As String = "") As SpecSuite
    
    Dim Specs As New SpecSuite
    Dim keyValue As Variant
    
    Specs.WB = WB
    
    With Specs.It("Should pass scenario")
        ' Load input values
        For Each keyValue In InputValues.Keys
            Specs.WB.Value(CStr(keyValue)) = InputValues.Item(CStr(keyValue))
        Next keyValue
        
        ' Run expectations
        For Each keyValue In ExpectedValues.Keys
            .Expect(Specs.WB.Value(CStr(keyValue))).ToEqual ExpectedValues.Item(CStr(keyValue))
        Next keyValue
    End With
    
    Specs.WB.CloseWB
    Set IScenario_RunScenario = Specs
End Function

''
' Load the scenario from a sheet
'
' @param {String} SheetName
' --------------------------------------------- '

Private Sub IScenario_Load(SheetName As String)
        
    Dim ScenarioSheet As Worksheet
    Dim Row As Integer
    Set InputValues = New Dictionary
    Set ExpectedValues = New Dictionary
    
    If SpecHelpers.SheetExists(SheetName, ThisWorkbook) Then
        Set ScenarioSheet = ThisWorkbook.Sheets(SheetName)
        
        Row = ScenarioInputStartRow
        Do While ScenarioSheet.Cells(Row, ScenarioInputStartCol) <> ""
            InputValues.Add _
                CStr(ScenarioSheet.Cells(Row, ScenarioInputStartCol)), _
                ScenarioSheet.Cells(Row, ScenarioInputStartCol + 1).Value
            
            Row = Row + 1
        Loop
        
        Row = ScenarioExpectedStartRow
        Do While ScenarioSheet.Cells(Row, ScenarioExpectedStartCol) <> ""
            ExpectedValues.Add _
                CStr(ScenarioSheet.Cells(Row, ScenarioExpectedStartCol)), _
                ScenarioSheet.Cells(Row, ScenarioExpectedStartCol + 1).Value
            
            Row = Row + 1
        Loop
        
    End If
    
End Sub
