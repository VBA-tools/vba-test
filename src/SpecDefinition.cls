VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpecDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' SpecDefinition v1.1.0
' (c) Tim Hall - https://github.com/timhall/Excel-TDD
'
' Provides helpers and acts as workbook proxy
'
' @dependencies
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '


Private pExpectations As Collection
Private pFailedExpectations As Collection


' --------------------------------------------- '
' Types
' --------------------------------------------- '

Public Enum SpecResult
    Pass
    Fail
    Pending
End Enum


' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public Description As String
Public Id As String

Public Property Get Expectations() As Collection
    If pExpectations Is Nothing Then
        Set pExpectations = New Collection
    End If
    Set Expectations = pExpectations
End Property
Private Property Let Expectations(Value As Collection)
    Set pExpectations = Value
End Property

Public Property Get FailedExpectations() As Collection
    If pFailedExpectations Is Nothing Then
        Set pFailedExpectations = New Collection
    End If
    Set FailedExpectations = pFailedExpectations
End Property
Private Property Let FailedExpectations(Value As Collection)
    Set pFailedExpectations = Value
End Property


' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Public Methods
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

''
' Create a new expectation to test the defined value
'
' @param {Variant} value Value to be tested by expectation
' @returns {Expectation}
' --------------------------------------------- '

Public Function Expect(Optional Value As Variant) As Expectation
    Dim Exp As New Expectation
    
    If VarType(Value) = vbObject Then
        Set Exp.ExpectValue = Value
    Else
        Exp.ExpectValue = Value
    End If
    Me.Expectations.Add Exp
    
    Set Expect = Exp
End Function

''
' Run each expectation, store failed expectations, and return result
'
' @returns {SpecResult} Pass/Fail/Pending
' --------------------------------------------- '

Public Function Result() As SpecResult
    Dim Exp As Expectation
    
    ' Reset failed expectations
    FailedExpectations = New Collection
    
    ' If no expectations have been defined, return pending
    If Me.Expectations.Count < 1 Then
        Result = Pending
    Else
        ' Loop through all expectations
        For Each Exp In Me.Expectations
            ' If expectation fails, store it
            If Exp.Result = Fail Then
                FailedExpectations.Add Exp
            End If
        Next Exp
        
        ' If no expectations failed, spec passes
        If Me.FailedExpectations.Count > 0 Then
            Result = Fail
        Else
            Result = Pass
        End If
    End If
End Function

''
' Helper to get result name (i.e. "Pass", "Fail", "Pending")
'
' @returns {String}
' --------------------------------------------- '

Public Function ResultName() As String
    Select Case Me.Result
        Case Pass: ResultName = "Pass"
        Case Fail: ResultName = "Fail"
        Case Pending: ResultName = "Pending"
    End Select
End Function
