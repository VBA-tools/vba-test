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
' ImmediateReporter v2.0.0-alpha
' (c) Tim Hall - https://github.com/VBA-tools/Excel-TDD
'
' Report results to Immediate Window
'
' @class ImmediateReporter
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Private WithEvents pSpecs As SpecSuite
Attribute pSpecs.VB_VarHelpID = -1
Private Finished As Boolean

''
' Listen to given SpecSuite
'
' @method ListenTo
' @param {SpecSuite} Specs
''
Public Sub ListenTo(Specs As SpecSuite)
    If Not pSpecs Is Nothing Then
        Done
    End If
    
    Debug.Print "= " & Now & " ="
    Set pSpecs = Specs
    Finished = False
End Sub

''
' Finish report for SpecSuite
'
' @method Done
''
Public Function Done()
    Finished = True
    Debug.Print "= DONE = " & pSpecs.Results.Count & " ="
End Function

Private Function ResultTypeToString(ResultType As SpecResultType) As String
    Select Case ResultType
    Case SpecResultType.Pass
        ResultTypeToString = "PASS"
    Case SpecResultType.Fail
        ResultTypeToString = "FAIL"
    Case SpecResultType.Pending
        ResultTypeToString = "pending"
    End Select
End Function

Private Sub Class_Terminate()
    If Not Finished Then
        Done
    End If
End Sub

Private Sub pSpecs_Result(Description As String, ResultType As SpecResultType, Expectations As Collection, Id As String)
    Debug.Print ResultTypeToString(ResultType) & " " & Description
End Sub
