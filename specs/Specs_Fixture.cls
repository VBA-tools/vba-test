VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Specs_Fixture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private WithEvents pSuite As SpecSuite
Attribute pSuite.VB_VarHelpID = -1

Public BeforeEachCallCount As Long
Public ResultCalls As Collection
Public AfterEachCallCount As Long

Public Sub ListenTo(Suite As SpecSuite)
    Set pSuite = Suite
End Sub

Private Sub pSuite_BeforeEach()
    BeforeEachCallCount = BeforeEachCallCount + 1
End Sub

Private Sub pSuite_Result(Spec As SpecDefinition)
    Me.ResultCalls.Add Spec
End Sub

Private Sub pSuite_AfterEach()
    AfterEachCallCount = AfterEachCallCount + 1
End Sub

Private Sub Class_Initialize()
    Set Me.ResultCalls = New Collection
End Sub

