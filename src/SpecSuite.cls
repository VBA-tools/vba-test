VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpecSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' SpecSuite v1.1.0
' (c) Tim Hall - https://github.com/timhall/Excel-TDD
'
' A collection of specs with the workbook that they act on
'
' @dependencies
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Private pSpecsCol As Collection
Private pWB As IWBProxy
Private pTempWorkbookPath As String

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Properties
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public TeardownWB As Boolean

Public Property Get WB() As IWBProxy
    Set WB = pWB
End Property
Public Property Let WB(Value As IWBProxy)
    Set pWB = Value
End Property

Public Property Get SpecsCol() As Collection
    If pSpecsCol Is Nothing Then: Set pSpecsCol = New Collection
    Set SpecsCol = pSpecsCol
End Property
Public Property Let SpecsCol(Value As Collection)
    Set pSpecsCol = Value
End Property


' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Public Methods
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

''
' Create a new spec definition with description
'
' @param {String} Description
' @param {String} [SpecId] Useful for identifying specific specs
' @returns {SpecDefinition} Initialized Spec Definition
' --------------------------------------------- '

Public Function It(Description As String, Optional SpecId As String = "") As SpecDefinition
    Dim Spec As New SpecDefinition
    
    ' Reset the workbook before spec
    If Not Me.WB Is Nothing Then
        Me.WB.ResetWB
    End If
    
    ' Initialize spec
    Spec.Description = Description
    Spec.Id = SpecId
    Me.SpecsCol.Add Spec
    Set It = Spec
End Function

Private Sub Class_Initialize()
    Me.TeardownWB = True
End Sub

Private Sub Class_Terminate()
    If TeardownWB = True And Not Me.WB Is Nothing Then
        Me.WB.CloseWB
    End If
End Sub
