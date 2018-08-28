VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Test_Fixture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private WithEvents pSuite As TestSuite
Attribute pSuite.VB_VarHelpID = -1

Public BeforeEachCallCount As Long
Public ResultCalls As Collection
Public AfterEachCallCount As Long

Public Sub ListenTo(Suite As TestSuite)
    Set pSuite = Suite
End Sub

Private Sub pSuite_BeforeEach(Test As TestCase)
    BeforeEachCallCount = BeforeEachCallCount + 1
End Sub

Private Sub pSuite_Result(Test As TestCase)
    Me.ResultCalls.Add Test
End Sub

Private Sub pSuite_AfterEach(Test As TestCase)
    AfterEachCallCount = AfterEachCallCount + 1
End Sub

Private Sub Class_Initialize()
    Set Me.ResultCalls = New Collection
End Sub

