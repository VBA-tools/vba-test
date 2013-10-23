VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Scenario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' Scenario v1.2.2
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

Private Description As String
Private InputValues As Dictionary
Private ExpectedValues As Dictionary

''
' Generate a spec suite to run through a scenario for the given wb
'
' @param {IWBProxy} WB to perform scenario on
' @param {String} Name of scenario
' @returns {SpecSuite}
' --------------------------------------------- '

Private Function IScenario_RunScenario(WB As IWBProxy) As SpecSuite
    
    Dim Specs As New SpecSuite
    Dim KeyValue As Variant
    
    If Description = "" Then
        Description = "should pass scenario"
    End If
    
    SpecHelpers.OpenIWBProxy WB
    
    With Specs.It(Description)
        ' Load input values
        For Each KeyValue In InputValues.Keys
            WB.Value(CStr(KeyValue)) = InputValues.Item(CStr(KeyValue))
        Next KeyValue
        
        ' Run expectations
        For Each KeyValue In ExpectedValues.Keys
            .Expect(WB.Value(CStr(KeyValue))).ToEqual ExpectedValues.Item(CStr(KeyValue))
        Next KeyValue
    End With
    
    SpecHelpers.CloseIWBProxy WB
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
        
        Description = ScenarioSheet.Cells(1, ScenarioExpectedStartCol)
        
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
