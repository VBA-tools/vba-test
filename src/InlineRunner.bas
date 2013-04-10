Attribute VB_Name = "InlineRunner"
''
' InlineRunner v1.1.0
' (c) Tim Hall - https://github.com/timhall/Excel-TDD
'
' Runner for outputting results of specs to Immediate window
'
' @dependencies
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

''
' Run the given suite
'
' @param {SpecSuite} Specs
' @param {Boolean} [Condensed] Hide failed expectations
' --------------------------------------------- '

Public Sub RunSuite(Specs As SpecSuite, Optional Condensed As Boolean = False)
    Dim SuiteCol As New Collection
    
    SuiteCol.Add Specs
    RunSuites SuiteCol, Condensed
End Sub

''
' Run the given collection of spec suites
'
' @param {Collection} of SpecSuite
' @param {Boolean} [Condensed] Hide failed expectations
' --------------------------------------------- '

Public Sub RunSuites(SuiteCol As Collection, Optional Condensed As Boolean = False)
    Dim Suite As SpecSuite
    Dim Spec As SpecDefinition
    Dim TotalCount As Integer
    Dim SuccessfulCount As Integer
    Dim FailedSpecs As New Collection
    Dim i As Integer
    
    For Each Suite In SuiteCol
        If Not Suite Is Nothing Then
            TotalCount = TotalCount + Suite.SpecsCol.Count
        
            For Each Spec In Suite.SpecsCol
                If Spec.Result = SpecResult.Fail Then
                    FailedSpecs.Add Spec
                End If
            Next Spec
        End If
    Next Suite
    
    Debug.Print
    Debug.Print vbNewLine & "= " & SummaryMessage(TotalCount, FailedSpecs) & " = " & Now & " ========================="
    If FailedSpecs.Count > 0 Then
        For Each Spec In FailedSpecs
            Debug.Print FailureMessage(Spec, Condensed)
        Next Spec
        Debug.Print "==="
    End If
End Sub

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Internal Methods
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Private Function SummaryMessage(TotalCount As Integer, FailedSpecs As Collection) As String
    If FailedSpecs.Count = 0 Then
        SummaryMessage = "PASS (" & TotalCount & " of " & TotalCount & " passed)"
    Else
        SummaryMessage = "FAIL (" & FailedSpecs.Count & " of " & TotalCount & " failed)"
    End If
End Function

Private Function FailureMessage(Spec As SpecDefinition, Condensed As Boolean) As String
    Dim FailedExpectation As SpecExpectation
    Dim i As Integer
    
    FailureMessage = "X "
    
    If Spec.Id <> "" Then
        FailureMessage = FailureMessage & Spec.Id & ": "
    End If
    
    FailureMessage = FailureMessage & "It " & Spec.Description
    
    If Not Condensed Then
        FailureMessage = FailureMessage & vbNewLine
        
        For Each FailedExpectation In Spec.FailedExpectations
            FailureMessage = FailureMessage & "  " & FailedExpectation.FailureMessage
            
            If i + 1 <> Spec.FailedExpectations.Count Then: FailureMessage = FailureMessage & vbNewLine
            i = i + 1
        Next FailedExpectation
    End If
End Function
