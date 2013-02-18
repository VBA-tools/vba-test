VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpecSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' SpecSuite v1.0.0
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


' --------------------------------------------- '
' Properties
' --------------------------------------------- '


Public TeardownWB As Boolean

Public Property Get wb() As IWBProxy
    Set wb = pWB
End Property
Public Property Let wb(value As IWBProxy)
    Set pWB = value
End Property

Public Property Get SpecsCol() As Collection
    If pSpecsCol Is Nothing Then: Set pSpecsCol = New Collection
    Set SpecsCol = pSpecsCol
End Property
Public Property Let SpecsCol(value As Collection)
    Set pSpecsCol = value
End Property


' ============================================= '
' Public Methods
' ============================================= '


''
' Create a new spec definition with description
'
' @param {String} Description
' @param {String} [SpecId] Useful for identifying specific specs
' @return {SpecDefinition} Initialized Spec Definition
' --------------------------------------------- '

Public Function It(Description As String, Optional SpecId As String = "") As SpecDefinition
    Dim spec As New SpecDefinition
    
    ' Reload the workbook
    ' TODO: Separate setup + teardown
    If Not wb Is Nothing Then
        wb.ReloadWB
    End If
    
    ' Initialize spec
    spec.Description = Description
    spec.Id = SpecId
    Me.SpecsCol.Add spec
    Set It = spec
End Function

Private Sub Class_Initialize()
    Me.TeardownWB = True
End Sub

Private Sub Class_Terminate()
    If TeardownWB = True And Not Me.wb Is Nothing Then
        Me.wb.CloseWB
    End If
End Sub
