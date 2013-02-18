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
' SpecDefinition v1.0.0
' (c) Tim Hall - https://github.com/timhall/Excel-TDD
'
' Provides helpers and acts as workbook proxy
'
' @dependencies
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

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
Private Property Let Expectations(value As Collection)
    Set pExpectations = value
End Property

Public Property Get FailedExpectations() As Collection
    If pFailedExpectations Is Nothing Then
        Set pFailedExpectations = New Collection
    End If
    Set FailedExpectations = pFailedExpectations
End Property
Private Property Let FailedExpectations(value As Collection)
    Set pFailedExpectations = value
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Create a new expectation to test the defined value
'
' @param {Variant} value Value to be tested by expectation
' @return {Expectation}
' --------------------------------------------- '

Public Function Expect(value As Variant) As Expectation
    Dim exp As New Expectation
    
    exp.ExpectValue = value
    Me.Expectations.Add exp
    
    Set Expect = exp
End Function

''
' Run each expectation, store failed expectations, and return result
'
' @return {SpecResult} Pass/Fail/Pending
' --------------------------------------------- '

Public Function Result() As SpecResult
    Dim exp As Expectation
    
    ' Reset failed expectations
    FailedExpectations = New Collection
    
    ' If no expectations have been defined, return pending
    If Me.Expectations.Count < 1 Then
        Result = Pending
    Else
        ' Loop through all expectations
        For Each exp In Me.Expectations
            ' If expectation fails, store it
            If exp.Result = Fail Then
                FailedExpectations.Add exp
            End If
        Next exp
        
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
' @return {String}
' --------------------------------------------- '

Public Function ResultName() As String
    Select Case Me.Result
        Case Pass: ResultName = "Pass"
        Case Fail: ResultName = "Fail"
        Case Pending: ResultName = "Pending"
    End Select
End Function
