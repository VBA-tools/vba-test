VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImmediateReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' ImmediateReporter v2.0.0-beta.3
' (c) Tim Hall - https://github.com/VBA-tools/VBA-TDD
'
' Report results to Immediate Window
'
' @class ImmediateReporter
' @author tim.hall.engr@gmail.com
' @license MIT (https://opensource.org/licenses/MIT)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private WithEvents pSuite As TestSuite
Attribute pSuite.VB_VarHelpID = -1
Private Finished As Boolean

' ============================================= '
' Public Methods
' ============================================= '

''
' Listen to given TestSuite
'
' @method ListenTo
' @param {TestSuite} Suite
''
Public Sub ListenTo(Suite As TestSuite)
    If Not pSuite Is Nothing Then
        Done
    End If
    
    Debug.Print "===" & IIf(Suite.Description <> "", " " & Suite.Description & " ===", "")
    Set pSuite = Suite
    Finished = False
End Sub

''
' Finish report for SpecSuite
'
' @method Done
''
Public Function Done()
    Finished = True
    
    Debug.Print "= " & Summary & " = " & Now & " =" & vbNewLine
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Function ResultTypeToString(ResultType As TestResultType) As String
    Select Case ResultType
    Case TestResultType.Pass
        ResultTypeToString = "+"
    Case TestResultType.Fail
        ResultTypeToString = "X"
    Case TestResultType.Pending
        ResultTypeToString = "."
    End Select
End Function

Private Function Summary() As String
    Dim Total As Long
    Dim Passed As Long
    Dim Failed As Long
    Dim Pending As Long
    Dim Skipped As Long
    
    Total = pSuite.Tests.Count
    Passed = pSuite.PassedTests.Count
    Failed = pSuite.FailedTests.Count
    Pending = pSuite.PendingTests.Count
    Skipped = pSuite.SkippedTests.Count
    
    Dim SummaryMessage As String
    If Failed > 0 Then
        SummaryMessage = "FAIL (" & Failed & " of " & Total & " failed"
    Else
        SummaryMessage = "PASS (" & Passed & " of " & Total & " passed"
    End If
    If Pending > 0 Then
        SummaryMessage = SummaryMessage & ", " & Pending & " pending"
    End If
    If Skipped > 0 Then
        SummaryMessage = SummaryMessage & ", " & Skipped & " skipped)"
    Else
        SummaryMessage = SummaryMessage & ")"
    End If
    
    Summary = SummaryMessage
End Function

Private Sub pSuite_Result(Test As TestCase)
    If Test.Result = TestResultType.Skipped Then
        Exit Sub
    End If

    Debug.Print ResultTypeToString(Test.Result) & " " & Test.Name
    
    If Test.Result = TestResultType.Fail Then
        Dim Failure As Variant
        For Each Failure In Test.Failures
            Debug.Print "  " & Failure
        Next Failure
    End If
End Sub

Private Sub Class_Terminate()
    If Not Finished Then
        Done
    End If
End Sub
