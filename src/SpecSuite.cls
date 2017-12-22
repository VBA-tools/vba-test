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
' SpecSuite v2.0.0-beta
' (c) Tim Hall - https://github.com/VBA-tools/VBA-TDD
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
Public Event Result(Definition As SpecDefinition)
Public Event AfterEach()

''
' (Optional) description of suite for display in runners
'
' @property Description
' @type String
''
Public Description As String

''
' @property Specs
' @type Collection<SpecDefinition>
''
Public Specs As VBA.Collection

''
' @property Suites
' @type Collection<SpecSuite>
''
Public Suites As VBA.Collection

''
' @property Result
' @type SpecResultType
''
Public Property Get Result() As SpecResultType
    Result = SpecResultType.Pending

    Dim Spec As SpecDefinition
    For Each Spec In Me.Specs
        If Spec.Result = SpecResultType.Pass Then
            Result = SpecResultType.Pass
        ElseIf Spec.Result = SpecResultType.Fail Then
            Result = SpecResultType.Fail
            Exit For
        End If
    Next Spec
End Property

''
' @property PassedSpecs
' @type Collection<SpecDefinition>
''
Public Property Get PassedSpecs() As Collection
    Dim Spec As SpecDefinition
    Dim Filtered As New Collection
    For Each Spec In Me.Specs
        If Spec.Result = SpecResultType.Pass Then
            Filtered.Add Spec
        End If
    Next Spec

    Set PassedSpecs = Filtered
End Property

''
' @property FailedSpecs
' @type Collection<SpecDefinition>
''
Public Property Get FailedSpecs() As Collection
    Dim Spec As SpecDefinition
    Dim Filtered As New Collection
    For Each Spec In Me.Specs
        If Spec.Result = SpecResultType.Fail Then
            Filtered.Add Spec
        End If
    Next Spec
    
    Set FailedSpecs = Filtered
End Property

''
' @property PendingSpecs
' @type Collection<SpecDefinition>
''
Public Property Get PendingSpecs() As Collection
    Dim Spec As SpecDefinition
    Dim Filtered As New Collection
    For Each Spec In Me.Specs
        If Spec.Result = SpecResultType.Pending Then
            Filtered.Add Spec
        End If
    Next Spec
    
    Set PendingSpecs = Filtered
End Property
''
'

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
' Called at completion of SpecDefinition
'
' @internal
' @method SpecDone
' @param {SpecDefinition} Spec
''
Public Sub SpecDone(Spec As SpecDefinition)
    Me.Specs.Add Spec
    RaiseEvent Result(Spec)
    RaiseEvent AfterEach
End Sub

' ============================================= '
' Private Functions
' ============================================= '

Private Sub Class_Initialize()
    Set Me.Specs = New VBA.Collection
    Set Me.Suites = New VBA.Collection
End Sub
