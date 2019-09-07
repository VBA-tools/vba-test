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
' # TestSuite
'
' A collection of tests, with events and results
'
' ```vba
' Dim Suite As New TestSuite
' Suite.Description = "Name"
'
' With Suite.Test("A")
'   .IsEqual 2 + 2, 4
' End With
' ```
'
' @class TestSuite
' @author tim.hall.engr@gmail.com
' @repository https://github.com/vba-tools/vba-test
' @license MIT
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Public Enum TestResultType
    Pass
    Fail
    Pending
    Skipped
End Enum

Public Event BeforeEach(Test As TestCase)
Public Event Result(Test As TestCase)
Public Event AfterEach(Test As TestCase)
Public Event Group(Suite As TestSuite)

''
' (Optional) description of suite for display in runners
''
Public Description As String

''
' @type Collection<TestCase>
''
Public Tests As VBA.Collection

''
' @internal
''
Public Parent As TestSuite

''
' Compute suite result from tests
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
' @type Collection<TestCase>
''
Public Property Get PassedTests() As VBA.Collection
    Set PassedTests = GetTestsByType(TestResultType.Pass)
End Property

''
' @type Collection<TestCase>
''
Public Property Get FailedTests() As VBA.Collection
    Set FailedTests = GetTestsByType(TestResultType.Fail)
End Property

''
' @type Collection<TestCase>
''
Public Property Get PendingTests() As VBA.Collection
    Set PendingTests = GetTestsByType(TestResultType.Pending)
End Property

''
' @type Collection<TestCase>
''
Public Property Get SkippedTests() As VBA.Collection
    Set SkippedTests = GetTestsByType(TestResultType.Skipped)
End Property

Public Property Get Self() As TestSuite
    Set Self = Me
End Property

''
' Create a new test case with description
''
Public Function Test(Description As String) As TestCase
    Dim Instance As New TestCase
    
    Instance.Description = Description
    Set Instance.Suite = Me
    
    OnTestBefore Instance
    
    Set Test = Instance
End Function

''
' Create a new group of tests with description
''
Public Function Group(Description As String) As TestSuite
    Dim Instance As New TestSuite
    
    Instance.Description = Description
    Set Instance.Parent = Me
    
    RaiseEvent Group(Instance)
    
    Set Group = Instance
End Function

''
' @internal
''
Public Sub TestComplete(Test As TestCase)
    OnTestResult Test
    OnTestAfter Test
End Sub

''
' @internal
''
Public Sub OnTestBefore(Test As TestCase)
    If Not Me.Parent Is Nothing Then
        Me.Parent.OnTestBefore Test
    End If
    
    RaiseEvent BeforeEach(Test)
End Sub

''
' @internal
''
Public Sub OnTestResult(Test As TestCase)
    Tests.Add Test
    RaiseEvent Result(Test)
    
    If Not Me.Parent Is Nothing Then
        Me.Parent.OnTestResult Test
    End If
End Sub

''
' @internal
''
Public Sub OnTestAfter(Test As TestCase)
    RaiseEvent AfterEach(Test)
    
    If Not Me.Parent Is Nothing Then
        Me.Parent.OnTestAfter Test
    End If
End Sub

' ============================================= '

Private Function GetTestsByType(ResultType As TestResultType) As VBA.Collection
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
