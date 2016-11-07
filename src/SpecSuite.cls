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
' SpecSuite v2.0.0-alpha
' (c) Tim Hall - https://github.com/VBA-tools/Excel-TDD
'
' A collection of specs and results
'
' @class SpecSuite
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Types, Events, and Properties
' --------------------------------------------- '

Public Enum SpecResultType
    Pass
    Fail
    Pending
End Enum

Public Event BeforeEach()
Public Event Result(Description As String, ResultType As SpecResultType, Expectations As VBA.Collection, Id As String)
Public Event AfterEach()

''
' (Optional) description of suite for display in runners
'
' @property Description
' @type String
''
Public Description As String

''
' @property Results
' @type Collection
''
Public Results As VBA.Collection

' ============================================= '
' Public Methods
' ============================================= '

''
' Create a new spec definition with description
'
' @method It
' @param {String} Description
' @param {String} [Id = ""] Useful for identifying specific specs
' @returns {SpecDefinition}
''
Public Function It(Description As String, Optional Id As String = "") As SpecDefinition
    Dim Spec As New SpecDefinition
    
    RaiseEvent BeforeEach
    
    ' Prepare Spec
    Spec.Description = Description
    Spec.Id = Id
    Set Spec.Suite = Me

    Set It = Spec
End Function

''
' Called at end of SpecDefinition with results
'
' @internal
' @method AddResult
' @param {SpecResultType} ResultType
' @param {Collection} Expectations
''
Public Sub AddResult(Description As String, ResultType As SpecResultType, Expectations As VBA.Collection, Id As String)
    Me.Results.Add Array(Description, ResultType, Expectations, Id)
    
    RaiseEvent Result(Description, ResultType, Expectations, Id)
    RaiseEvent AfterEach
End Sub

' ============================================= '
' Private Functions
' ============================================= '

Private Sub Class_Initialize()
    Set Me.Results = New VBA.Collection
End Sub
