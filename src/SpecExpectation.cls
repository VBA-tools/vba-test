VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpecExpectation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' SpecExpectation v2.0.0-beta
' (c) Tim Hall - https://github.com/VBA-tools/Excel-TDD
'
' Provides various tests that can be performed for a provided value
'
' @class SpecExpectation
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

''
' @feature RunMatcher
' @compatibility: Platforms / applications that support Application.Run
'   Platforms: Windows
'   Applications: Excel, Word, Access
''
#Const EnableRunMatcher = True

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public Actual As Variant
Public Expected As Variant
Public Passed As Boolean
Public FailureMessage As String

' ============================================= '
' Public Methods
' ============================================= '

''
' Check if the actual value is equal to the expected value
'
' @method ToEqual
' @param {Variant} Expected value
''
Public Sub ToEqual(Expected As Variant)
    Check IsEqual(Me.Actual, Expected), "to equal", Expected:=Expected
End Sub

''
' @method ToNotEqual
' @param {Variant} Expected value
''
Public Sub ToNotEqual(Expected As Variant)
    Check IsEqual(Me.Actual, Expected), "to not equal", Expected:=Expected, Inverse:=True
End Sub

Private Function IsEqual(Actual As Variant, Expected As Variant) As Variant
    If VBA.IsError(Actual) Or VBA.IsError(Expected) Then
        IsEqual = False
    ElseIf VBA.IsObject(Actual) Or VBA.IsObject(Expected) Then
        IsEqual = "Unsupported: Can't compare objects"
    ElseIf VBA.VarType(Actual) = vbDouble And VBA.VarType(Expected) = vbDouble Then
        ' It is inherently difficult/almost impossible to Check equality of Double
        ' http://support.microsoft.com/kb/78113
        '
        ' Compare up to 15 significant figures
        ' -> Format as scientific notation with 15 significant figures and then compare strings
        IsEqual = IsCloseTo(Actual, Expected, 15)
    Else
        IsEqual = Actual = Expected
    End If
End Function

''
' Check if the actual value is undefined
' (Nothing, Empty, Null, or Missing)
'
' @method ToBeUndefined
''
Public Sub ToBeUndefined()
    Check IsUndefined(Me.Actual), "to be undefined"
End Sub

''
' Check if the actual value is not undefined
' (not Nothing, Empty, Null, or Missing)
'
' @method ToNotBeUndefined
''
Public Sub ToNotBeUndefined()
    Check IsUndefined(Me.Actual), "to not be undefined", Inverse:=True
End Sub

Private Function IsUndefined(Actual As Variant) As Variant
    IsUndefined = IsNothing(Actual) Or VBA.IsEmpty(Actual) Or VBA.IsNull(Actual) Or VBA.IsMissing(Actual)
End Function

''
' Check if the actual value is Nothing
'
' @method ToBeNothing
''
Public Sub ToBeNothing()
    Check IsNothing(Me.Actual), "to be nothing"
End Sub

''
' Check if the actual value is not Nothing
'
' @method ToNotBeNothing
''
Public Sub ToNotBeNothing()
    Check IsNothing(Me.Actual), "to not be nothing", Inverse:=True
End Sub

Private Function IsNothing(Actual As Variant) As Variant
    If VBA.IsObject(Actual) Then
        If Actual Is Nothing Then
            IsNothing = True
        Else
            IsNothing = False
        End If
    Else
        IsNothing = False
    End If
End Function

''
' Check if the actual value is empty
'
' @method ToBeEmpty
''
Public Sub ToBeEmpty()
    Check VBA.IsEmpty(Me.Actual), "to be empty"
End Sub

''
' Check if the actual value is not empty
'
' @method ToNotBeEmpty
''
Public Sub ToNotBeEmpty()
    Check VBA.IsEmpty(Me.Actual), "to not be empty", Inverse:=True
End Sub

''
' Check if the actual value is null
'
' @method ToBeNull
''
Public Sub ToBeNull()
    Check VBA.IsNull(Me.Actual), "to be null"
End Sub

''
' Check if the actual value is not null
'
' @method ToNotBeNull
''
Public Sub ToNotBeNull()
    Check VBA.IsNull(Me.Actual), "to not be null", Inverse:=True
End Sub

''
' Check if the actual value is missing
'
' @method ToBeMissing
''
Public Sub ToBeMissing()
    Check VBA.IsMissing(Me.Actual), "to be missing"
End Sub

''
' Check if the actual value is not missing
'
' @method ToNotBeMissing
''
Public Sub ToNotBeMissing()
    Check VBA.IsMissing(Me.Actual), "to not be missing", Inverse:=True
End Sub

''
' Check if the actual value is less than the expected value
'
' @method ToBeLessThan / ToBeLT
' @param {Variant} Expected value
''
Public Sub ToBeLessThan(Expected As Variant)
    Check IsLT(Me.Actual, Expected), "to be less than", Expected:=Expected
End Sub
Public Sub ToBeLT(Expected As Variant)
    ToBeLessThan Expected
End Sub

Private Function IsLT(Actual As Variant, Expected As Variant) As Variant
    If VBA.IsError(Actual) Or VBA.IsError(Expected) Or Actual >= Expected Then
        IsLT = False
    Else
        IsLT = True
    End If
End Function

''
' Check if the actual value is less than or equal to the expected value
'
' @method ToBeLessThanOrEqualTo / ToBeLTE
' @param {Variant} Expected value
''
Public Sub ToBeLessThanOrEqualTo(Expected As Variant)
    Check IsLTE(Me.Actual, Expected), "to be less than or equal to", Expected:=Expected
End Sub
Public Sub ToBeLTE(Expected As Variant)
    ToBeLessThanOrEqualTo Expected
End Sub

Private Function IsLTE(Actual As Variant, Expected As Variant) As Variant
    If VBA.IsError(Actual) Or VBA.IsError(Expected) Or Actual > Expected Then
        IsLTE = False
    Else
        IsLTE = True
    End If
End Function

''
' Check if the actual value is greater than the expected value
'
' @method ToBeGreaterThan / ToBeGT
' @param {Variant} Expected value
''
Public Sub ToBeGreaterThan(Expected As Variant)
    Check IsGT(Me.Actual, Expected), "to be greater than", Expected:=Expected
End Sub
Public Sub ToBeGT(Expected As Variant)
    ToBeGreaterThan Expected
End Sub

Private Function IsGT(Actual As Variant, Expected As Variant) As Variant
    If VBA.IsError(Actual) Or VBA.IsError(Expected) Or Actual <= Expected Then
        IsGT = False
    Else
        IsGT = True
    End If
End Function

''
' Check if the actual value is greater than or equal to the expected value
'
' @method ToBeGreaterThanOrEqualTo / ToBeGTE
' @param {Variant} Expected
''
Public Sub ToBeGreaterThanOrEqualTo(Expected As Variant)
    Check IsGTE(Me.Actual, Expected), "to be greater than or equal to", Expected:=Expected
End Sub
Public Sub ToBeGTE(Expected As Variant)
    ToBeGreaterThanOrEqualTo Expected
End Sub

Private Function IsGTE(Actual As Variant, Expected As Variant) As Variant
    If VBA.IsError(Actual) Or VBA.IsError(Expected) Or Actual < Expected Then
        IsGTE = False
    Else
        IsGTE = True
    End If
End Function

''
' Check if the actual value is close to the expected value
'
' @method ToBeCloseTo
' @param {Variant} Expected value
' @param {Integer} SignificantFigures (1-15)
''
Public Sub ToBeCloseTo(Expected As Variant, SignificantFigures As Integer)
    Check IsCloseTo(Me.Actual, Expected, SignificantFigures), "to be close to", Expected:=Expected
End Sub

''
' Check if the actual value is not close to the expected value
'
' @method ToNotBeCloseTo
' @param {Variant} Expected value
' @param {Integer} SignificantFigures (1-15
''
Public Sub ToNotBeCloseTo(Expected As Variant, SignificantFigures As Integer)
    Check IsCloseTo(Me.Actual, Expected, SignificantFigures), "to be close to", Expected:=Expected, Inverse:=True
End Sub

Private Function IsCloseTo(Actual As Variant, Expected As Variant, SignificantFigures As Integer) As Variant
    Dim ActualAsString As String
    Dim ExpectedAsString As String
    
    If SignificantFigures < 1 Or SignificantFigures > 15 Then
        IsCloseTo = "ToBeCloseTo/ToNotBeClose to can only compare from 1 to 15 significant figures"""
    ElseIf Not VBA.IsError(Actual) And Not VBA.IsError(Expected) Then
        ' Convert values to scientific notation strings and then compare strings
        If Actual > 1 Then
            ActualAsString = VBA.Format$(Actual, VBA.Left$("0.00000000000000", SignificantFigures + 1) & "e+0")
        Else
            ActualAsString = VBA.Format$(Actual, VBA.Left$("0.00000000000000", SignificantFigures + 1) & "e-0")
        End If
        
        If Expected > 1 Then
            ExpectedAsString = VBA.Format$(Expected, VBA.Left$("0.00000000000000", SignificantFigures + 1) & "e+0")
        Else
            ExpectedAsString = VBA.Format$(Expected, VBA.Left$("0.00000000000000", SignificantFigures + 1) & "e-0")
        End If
        
        IsCloseTo = ActualAsString = ExpectedAsString
    End If
End Function

''
' Check if the actual value array contains the expected value
'
' @method ToContain
' @param {Variant} Expected value
''
Public Sub ToContain(Expected As Variant)
    Check Contains(Me.Actual, Expected), "to contain", Expected:=Expected
End Sub

''
' Check if the actual value array does not contain the expected value
'
' @method ToNotContain
' @param {Variant} Expected value
''
Public Sub ToNotContain(Expected As Variant)
    Check Contains(Me.Actual, Expected), "to not contain", Expected:=Expected, Inverse:=True
End Sub

Private Function Contains(Actual As Variant, Expected As Variant) As Variant
    If Not IsArrayOrCollection(Actual) Then
        Contains = "Error: Actual needs to be an Array or Collection for ToContain/ToNotContain"
    Else
        Dim i As Integer
        If TypeOf Actual Is Collection Then
            For i = 1 To Actual.Count
                If Actual.Item(i) = Expected Then
                    Contains = True
                    Exit Function
                End If
            Next i
        Else
            For i = LBound(Actual) To UBound(Actual)
                If Actual(i) = Expected Then
                    Contains = True
                    Exit Function
                End If
            Next i
        End If
    End If
End Function

''
' Check if the actual value string has a match for the expected value substring
' (Only checks if the actual contains the expected string currently)
'
' @method ToMatch
' @param {Variant} Expected value
''
Public Sub ToMatch(Expected As Variant)
    Check Matches(Me.Actual, Expected), "to match", Expected:=Expected
End Sub

''
' Check if the actual value string does not have a match for the expected value substring
' (Only checks if the actual does not contain the expected string currently)
'
' @method ToNotMatch
' @param {Variant} Expected value
''
Public Sub ToNotMatch(Expected As Variant)
    Check Matches(Me.Actual, Expected), "to not match", Expected:=Expected, Inverse:=True
End Sub

Private Function Matches(Actual As Variant, Expected As Variant) As Variant
    If VBA.InStr(Actual, Expected) > 0 Then
        Matches = True
    Else
        Matches = False
    End If
End Function

''
' Run custom matcher
'
' @example
' .Expect(100).RunMatcher "Module.ToBeWithin", "to be within", 90, 110
'
' Module:
' Public Function ToBeWithin(Actual As Variant, Args As Variant) As Variant
'     If UBound(Args) - LBound(Args) < 1 Then
'         ' Return string for specific failure message
'         ToBeWithin = "Need to pass in upper-bound to ToBeWithin"
'     Else
'         If Actual >= Args(0) And Actual <= Args(1) Then
'             ' Return true for pass
'             ToBeWithin = True
'         Else
'             ' Return false for fail or custom failure message
'             ToBeWithin = False
'         End If
'     End If
' End Function
'
' @method RunMatcher
' @param {String} Name of function for matcher
' @param {String} Message
' @param {...} Arguments for custom matcher
''
#If EnableRunMatcher Then
Public Sub RunMatcher(Name As String, Message As String, ParamArray Arguments())
    Dim Expected As String
    Dim i As Integer
    Dim HasArguments As Boolean
        
    HasArguments = UBound(Arguments) >= 0
    For i = LBound(Arguments) To UBound(Arguments)
        If Expected = "" Then
            Expected = GetStringForValue(Arguments(i))
        ElseIf i = UBound(Arguments) Then
            If (UBound(Arguments) > 1) Then
                Expected = Expected & ", and " & GetStringForValue(Arguments(i))
            Else
                Expected = Expected & " and " & GetStringForValue(Arguments(i))
            End If
        Else
            Expected = Expected & ", " & GetStringForValue(Arguments(i))
        End If
    Next i
    
    If HasArguments Then
        Check Application.Run(Name, Me.Actual, Arguments), Message, Expected:=Expected
    Else
        Check Application.Run(Name, Me.Actual), Message
    End If
End Sub
#End If

' ============================================= '
' Private Methods
' ============================================= '

Private Sub Check(Result As Variant, Message As String, Optional Expected As Variant, Optional Inverse As Boolean = False)
    If Not VBA.IsMissing(Expected) Then
        If VBA.IsObject(Expected) Then
            Set Me.Expected = Expected
        Else
            Me.Expected = Expected
        End If
    End If
    
    If VBA.VarType(Result) = VBA.vbString Then
        Fails CStr(Result)
    Else
        If Inverse Then
            Result = Not Result
        End If
        
        If Result Then
            Passes
        Else
            Fails CreateFailureMessage(Message, Expected)
        End If
    End If
End Sub

Private Sub Passes()
    Me.Passed = True
End Sub

Private Sub Fails(Message As String)
    Me.Passed = False
    Me.FailureMessage = Message
End Sub

Private Function CreateFailureMessage(Message As String, Optional Expected As Variant) As String
    CreateFailureMessage = "Expected " & GetStringForValue(Me.Actual) & " " & Message
    If Not VBA.IsMissing(Expected) Then
        CreateFailureMessage = CreateFailureMessage & " " & GetStringForValue(Expected)
    End If
End Function

Private Function GetStringForValue(Value As Variant) As String
    If VBA.IsMissing(Value) Then
        GetStringForValue = "(Missing)"
        Exit Function
    End If

    Select Case VBA.VarType(Value)
    Case VBA.vbObject
        If Value Is Nothing Then
            GetStringForValue = "(Nothing)"
        Else
            GetStringForValue = "(Object)"
        End If
    Case VBA.vbArray To VBA.vbArray + VBA.vbByte
        GetStringForValue = "(Array)"
    Case VBA.vbEmpty
        GetStringForValue = "(Empty)"
    Case VBA.vbNull
        GetStringForValue = "(Null)"
    Case VBA.vbString
        GetStringForValue = """" & Value & """"
    Case Else
        GetStringForValue = CStr(Value)
    End Select
    
    If GetStringForValue = "" Then
        GetStringForValue = "(Undefined)"
    End If
End Function

Private Function IsArrayOrCollection(Value As Variant) As Boolean
    Select Case VBA.VarType(Value)
    Case VBA.vbArray To VBA.vbArray + VBA.vbByte
        IsArrayOrCollection = True
    Case VBA.vbObject
        If TypeOf Value Is Collection Then
            IsArrayOrCollection = True
        End If
    End Select
End Function

