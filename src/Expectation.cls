VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Expectation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Expectation v1.0.0
' (c) Tim Hall - https://github.com/timhall/Excel-TDD
'
' Provides various tests that can be performed for a provided value
'
' @dependencies
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '


Public Enum ExpectResult
    Pass
    Fail
End Enum


' --------------------------------------------- '
' Public Properties
' --------------------------------------------- '


Public ExpectValue As Variant
Public Result As ExpectResult
Public FailureMessage As String


' ============================================= '
' Public Methods
' ============================================= '


''
' Check if the defined value is equal to the test value
'
' @param {Variant} Value
' --------------------------------------------- '

Public Sub ToEqual(value As Variant)
    Dim failed As Boolean
    failed = False
    
    If IsError(Me.ExpectValue) Or IsError(value) Then
        failed = True
    
    ' If both the values are doubles, use decimal notequal comparison
    ElseIf VarType(Me.ExpectValue) = vbDouble And VarType(value) = vbDouble Then
        If CDec(Me.ExpectValue) <> CDec(value) Then
            failed = True
        End If
        
    ' Otherwise use standard notequal comparison
    Else
        If Me.ExpectValue <> value Then
            failed = True
        End If
    End If
    
    ' If test fails, create failure message
    If failed Then
        Me.Result = ExpectResult.Fail
        Me.FailureMessage = CreateFailureMessage("to equal", value)
    Else
        Me.Result = ExpectResult.Pass
    End If
End Sub

''
' Check if the defined value is not equal to the test value
'
' @param {Variant} Value
' --------------------------------------------- '

Public Sub ToNotEqual(value As Variant)
    If Not IsError(Me.ExpectValue) And Not IsError(value) And Me.ExpectValue <> value Then
        Me.Result = ExpectResult.Pass
    Else
        Me.Result = ExpectResult.Fail
        Me.FailureMessage = CreateFailureMessage("to not equal", value)
    End If
End Sub

''
' Check if the test value is defined
' --------------------------------------------- '

Public Sub ToBeDefined()
    ' Make sure the value isn't empty or null
    ' (There may be other things that need to checked as well)
    If Not IsEmpty(Me.ExpectValue) And Not IsNull(Me.ExpectValue) Then
        Me.Result = ExpectResult.Pass
    Else
        Me.Result = ExpectResult.Fail
        Me.FailureMessage = CreateFailureMessage("to be defined")
    End If
End Sub

''
' Check if the test value is undefined
' --------------------------------------------- '

Public Sub ToBeUndefined()
    ' If the value is empty or null, it is undefined
    If IsEmpty(Me.ExpectValue) Or IsNull(Me.ExpectValue) Then
        Me.Result = ExpectResult.Pass
    Else
        Me.Result = ExpectResult.Fail
        Me.FailureMessage = CreateFailureMessage("to be undefined")
    End If
End Sub

''
' Check if the defined value is less than the test value
'
' @param {Variant} Value
' --------------------------------------------- '

Public Sub ToBeLessThan(value As Variant)
    If IsError(Me.ExpectValue) Or IsError(value) Or Me.ExpectValue >= value Then
        Me.Result = ExpectResult.Fail
        Me.FailureMessage = CreateFailureMessage("to be less than", value)
    Else
        Me.Result = ExpectResult.Pass
    End If
End Sub
Public Sub ToBeLT(value As Variant)
    Call ToBeLessThan(value)
End Sub

''
' Check if the defined value is less than or equal to the test value
'
' @param {Variant} Value
' --------------------------------------------- '

Public Sub ToBeLessThanOrEqualTo(value As Variant)
    If IsError(Me.ExpectValue) Or IsError(value) Or Me.ExpectValue > value Then
        Me.Result = ExpectResult.Fail
        Me.FailureMessage = CreateFailureMessage("to be less than or equal to", value)
    Else
        Me.Result = ExpectResult.Pass
    End If
End Sub
Public Sub ToBeLTE(value As Variant)
    Call ToBeLessThanOrEqualTo(value)
End Sub

''
' Check if the defined value is greater than the test value
'
' @param {Variant} Value
' --------------------------------------------- '

Public Sub ToBeGreaterThan(value As Variant)
    If IsError(Me.ExpectValue) Or IsError(value) Or Me.ExpectValue <= value Then
        Me.Result = ExpectResult.Fail
        Me.FailureMessage = CreateFailureMessage("to be greater than", value)
    Else
        Me.Result = ExpectResult.Pass
    End If
End Sub
Public Sub ToBeGT(value As Variant)
    Call ToBeGreaterThan(value)
End Sub

''
' Check if the defined value is greater than or equal to the test value
'
' @param {Variant} Value
' --------------------------------------------- '

Public Sub ToBeGreaterThanOrEqualTo(value As Variant)
    If IsError(Me.ExpectValue) Or IsError(value) Or Me.ExpectValue < value Then
        Me.Result = ExpectResult.Fail
        Me.FailureMessage = CreateFailureMessage("to be greater than or equal to", value)
    Else
        Me.Result = ExpectResult.Pass
    End If
End Sub
Public Sub ToBeGTE(value As Variant)
    Call ToBeGreaterThanOrEqualTo(value)
End Sub


' ============================================= '
' Private Methods
' ============================================= '


Private Function GetStringForValue(value As Variant) As String
    GetStringForValue = CStr(value)
    If GetStringForValue = "" Then
        GetStringForValue = "(Undefined)"
    End If
End Function

Private Function CreateFailureMessage(message As String, Optional value As Variant) As String
    CreateFailureMessage = "Expected " & GetStringForValue(Me.ExpectValue) & " " & message
    If Not IsMissing(value) Then
        CreateFailureMessage = CreateFailureMessage & " " & GetStringForValue(value)
    End If
End Function

