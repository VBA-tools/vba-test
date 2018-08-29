VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' TestSuite v2.0.0-beta.3
' (c) Tim Hall - https://github.com/vba-tools/vba-test
'
' A collection of tests, with events and results
'
' @class TestSuite
' @author tim.hall.engr@gmail.com
' @license MIT (https://opensource.org/licenses/MIT)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Types, Events, and Properties
' --------------------------------------------- '

Public Enum TestResultType
    Pass
    Fail
    Pending
    Skipped
End Enum

Public Event BeforeEach(Test As TestCase)
Public Event Result(Test As TestCase)
Public Event AfterEach(Test As TestCase)

''
' (Optional) description of suite for display in runners
'
' @property Description
' @type String
''
Public Description As String

''
' @property Tests
' @type Collection<TestCase>
''
Public Tests As VBA.Collection

''
' Compute suite result from tests
'
' @property Result
' @type SpecResultType
''
Public Property Get Result() As TestResultType
    Result = TestResultType.Pending
    
    Dim Test As TestCase
    For Each Test In Me.Tests
        If Test.Result = TestResultType.Pass Then
            Result = TestResultType.Pass
        ElseIf Test.Result = TestResultType.Fail Then
            Result = TestResultType.Fail
            Exit For
        End If
    Next Test
End Property

''
' @property PassedTests
' @type Collection<TestCase>
''
Public Property Get PassedTests() As VBA.Collection
    Set PassedTests = GetTestsByType(TestResultType.Pass)
End Property

''
' @property FailedTests
' @type Collection<TestCase>
''
Public Property Get FailedTests() As VBA.Collection
    Set FailedTests = GetTestsByType(TestResultType.Fail)
End Property

''
' @property PendingTests
' @type Collection<TestCase>
''
Public Property Get PendingTests() As VBA.Collection
    Set PendingTests = GetTestsByType(TestResultType.Pending)
End Property

''
' @property SkippedTests
' @type Collection<TestCase>
''
Public Property Get SkippedTests() As VBA.Collection
    Set SkippedTests = GetTestsByType(TestResultType.Skipped)
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Create a new test case with name
'
' @method Test
' @param {String} Name
' @returns {TestCase}
''
Public Function Test(Name As String) As TestCase
    Dim Instance As New TestCase
    
    Instance.Name = Name
    Set Instance.Suite = Me
    
    RaiseEvent BeforeEach(Instance)
    
    Set Test = Instance
End Function

Public Sub TestComplete(Test As TestCase)
    Tests.Add Test

    RaiseEvent Result(Test)
    RaiseEvent AfterEach(Test)
End Sub

' ============================================= '
' Private Functions
' ============================================= '

Private Function GetTestsByType(ResultType As TestResultType) As Collection
    Dim Test As TestCase
    Dim Filtered As New VBA.Collection
    For Each Test In Me.Tests
        If Test.Result = ResultType Then
            Filtered.Add Test
        End If
    Next Test

    Set GetTestsByType = Filtered
End Function


Private Sub Class_Initialize()
    Set Tests = New VBA.Collection
End Sub
