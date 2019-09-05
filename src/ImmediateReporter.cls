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
' # ImmediateReporter
'
' Report results to Immediate Window
'
' ```vba
' Dim Suite As New TestSuite
' ...
'
' Dim Reporter As New ImmediateReporter
' Reporter.ListenTo Suite
' ```
'
' @class ImmediateReporter
' @author tim.hall.engr@gmail.com
' @repository https://github.com/vba-tools/vba-test
' @license MIT
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Private WithEvents pSuite As TestSuite
Attribute pSuite.VB_VarHelpID = -1
Private Finished As Boolean

''
' Listen to given TestSuite
''
Public Sub ListenTo(Suite As TestSuite)
    If Not pSuite Is Nothing Then
        ' If already listening to suite,
        ' report summary before moving on to next suite
        PrintSummary
    End If
    
    Finished = False
    Set pSuite = Suite
    PrintHeader Suite

    ' Report any tests added prior to listening
    Dim Test As TestCase
    For Each Test In Suite.Tests
        PrintResult Test
    Next Test
End Sub

' ============================================= '

Private Sub PrintHeader(Suite As TestSuite)
    Debug.Print "===" & IIf(Suite.Description <> "", " " & Suite.Description & " ===", "")
End Sub

Private Sub PrintResult(Test As TestCase)
    If Test.Result = TestResultType.Skipped Then
        Exit Sub
    End If

    Debug.Print ResultTypeToString(Test.Result) & " " & Test.Description
    
    If Test.Result = TestResultType.Fail Then
        Dim Failure As Variant
        For Each Failure In Test.Failures
            Debug.Print "  " & Failure
        Next Failure
    End If
End Sub

Private Sub PrintSummary()
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
    
    Dim Summary As String
    If Failed > 0 Then
        Summary = "FAIL (" & Failed & " of " & Total & " failed"
    Else
        Summary = "PASS (" & Passed & " of " & Total & " passed"
    End If
    If Pending > 0 Then
        Summary = Summary & ", " & Pending & " pending"
    End If
    If Skipped > 0 Then
        Summary = Summary & ", " & Skipped & " skipped)"
    Else
        Summary = Summary & ")"
    End If

    Debug.Print "= " & Summary & " = " & Now & " =" & vbNewLine
End Sub

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

Private Sub pSuite_Group(Suite As TestSuite)
    PrintHeader Suite
End Sub

Private Sub pSuite_Result(Test As TestCase)
    PrintResult Test
End Sub

Private Sub Class_Terminate()
    If Not Finished Then
        PrintSummary
    End If
End Sub
