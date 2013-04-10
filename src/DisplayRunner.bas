Attribute VB_Name = "DisplayRunner"
''
' DisplayRunner v1.1.0
' (c) Tim Hall - https://github.com/timhall/Excel-TDD
'
' Runner for outputting results of specs to worksheet
'
' @dependencies
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Private Const RunnerSheetName As String = "Spec Runner"
Private Const OutputStartRow As Integer = 6
Private Const IdCol As Integer = 1
Private Const DescCol As Integer = 2
Private Const ResultCol As Integer = 3

' Get/set path of workbook to run specs on
Public Property Get WBPath() As String
    WBPath = RunnerSheet.[Filename].Value
End Property
Public Property Let WBPath(Value As String)
    RunnerSheet.[Filename].Value = Value
End Property

' Get the runner sheet
Public Property Get RunnerSheet() As Worksheet
    If SpecHelpers.SheetExists(RunnerSheetName, ThisWorkbook) Then
        Set RunnerSheet = ThisWorkbook.Sheets(RunnerSheetName)
    Else
        Err.Raise vbObjectError + 1, "DisplayRunner", "Unable to find runner sheet"
    End If
End Property


' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Methods
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Sub RunSpecs()
    
    ' *****
    ' Overwrite this Sub with your code to call specs
    ' *****
    
    Debug.Print "Notice: DisplayRunner.RunSpecs needs to be linked to your specs code in order to run"
    
End Sub

''
' Run the given suite
'
' @param {SpecSuite} Specs
' --------------------------------------------- '

Public Sub RunSuite(Specs As SpecSuite)
    ' Simply add to empty collection and call RunSuites
    Dim SuiteCol As New Collection
    
    SuiteCol.Add Specs
    RunSuites SuiteCol
End Sub

''
' Run the given collection of spec suites
'
' @param {Collection} of SpecSuite
' --------------------------------------------- '

Public Sub RunSuites(SuiteCol As Collection)
    
    Dim Suite As SpecSuite
    Dim Spec As SpecDefinition
    Dim Row As Integer
    
    ' 0. Disable screen updating
    Dim PrevUpdating As Boolean
    PrevUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' 1. Clear existing output
    ClearOutput
    
    ' 2. Loop through Suites and output specs
    Row = OutputStartRow
    For Each Suite In SuiteCol
        If Not Suite Is Nothing Then
            For Each Spec In Suite.SpecsCol
                OutputSpec Spec, Row
            Next Spec
        End If
    Next Suite
    
    ' Finally, restore screen updating
    Application.ScreenUpdating = PrevUpdating
    
End Sub

''
' Browse for the workbook to run specs on
' --------------------------------------------- '

Public Sub BrowseForWB()
    Dim BrowseWB As String

    BrowseWB = Application.GetOpenFilename( _
        FileFilter:="Excel Workbooks (*.xls; *.xlsx; *.xlsm), *.xls, *.xlsx, *.xlsm", _
        Title:="Select the Excel Workbook to Test", _
        MultiSelect:=False _
    )

    If BrowseWB <> "" And BrowseWB <> "False" Then
        WBPath = BrowseWB
    End If
End Sub


' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Internal
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Private Sub OutputSpec(Spec As SpecDefinition, ByRef Row As Integer)
    
    RunnerSheet.Cells(Row, IdCol) = Spec.Id
    RunnerSheet.Cells(Row, DescCol) = "It " & Spec.Description
    RunnerSheet.Cells(Row, ResultCol) = Spec.ResultName
    Row = Row + 1
    
    If Spec.FailedExpectations.Count > 0 Then
        Dim Exp As SpecExpectation
        For Each Exp In Spec.FailedExpectations
            RunnerSheet.Cells(Row, DescCol) = "X  " & Exp.FailureMessage
            Row = Row + 1
        Next Exp
    End If
    
End Sub

Private Sub ClearOutput()
    Dim EndRow As Integer
    
    Dim PrevUpdating As Boolean
    PrevUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    EndRow = SpecHelpers.LastRow(RunnerSheet)
    
    If EndRow >= OutputStartRow Then
        RunnerSheet.Range(Cells(OutputStartRow, IdCol), Cells(EndRow, ResultCol)).ClearContents
    End If
    
    Application.ScreenUpdating = PrevUpdating
End Sub

