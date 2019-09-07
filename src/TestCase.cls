VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' # TestCase
'
' Verify a single test case with assertions
'
' ```vba
' Dim Test As New TestCase
'
' Test.Description = "should add many numbers
' With Test
'   .IsEqual Add(2, 2), 4
'   .IsEqual Add(1, 2, 3), 6
' End With
' ```
'
' @class TestCase
' @author tim.hall.engr@gmail.com
' @repository https://github.com/vba-tools/vba-test
' @license MIT
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Private pFailures As VBA.Collection

''
' Set test description as displayed in reporter
''
Public Description As String

''
' Get/set values for test context
' (useful for passing values to BeforeEach/AfterEach)
''
Public Context As Dictionary

''
' @internal
''
Public Planned As Long

''
' @internal
''
Public Successes As Long

''
' @internal
''
Public Skipped As Boolean

''
' @internal
''
Public Suite As TestSuite

Public Property Get Result() As TestResultType
    If Me.Skipped Then
        Result = TestResultType.Skipped
    ElseIf Me.Successes = 0 And Me.Failures.Count = 0 Then
        Result = TestResultType.Pending
    ElseIf Me.Failures.Count > 0 Then
        Result = TestResultType.Fail
    Else
        Result = TestResultType.Pass
    End If
End Property

Public Property Get Failures() As VBA.Collection
    Dim Total As Long
    Total = Me.Successes + pFailures.Count
    
    If Me.Planned > 0 And Me.Planned <> Total Then
        Dim Message As String
        Dim Failure As Variant
        
        Set Failures = New VBA.Collection
        For Each Failure In pFailures
            Failures.Add Failure
        Next Failure
        
        Message = "Total assertions, {0}, does not equal planned, {1}"
        Failures.Add FormatMessage(Message, Total, Me.Planned)
    Else
        Set Failures = pFailures
    End If
End Property

''
' Access test instance (e.g. for passing to custom matchers)
''
Public Property Get Self() As TestCase
    Self = Me
End Property

''
' Check if two values are deep equal (including Array, Collection, and Dictionary)
''
Public Sub IsEqual(A As Variant, B As Variant, Optional Message As String = _
    "Expected {0} to equal {1}")

    Check IsDeepEqual(A, B), Message, A, B
End Sub

''
' Check if two values are not deep equal (including Array, Collection, and Dictionary)
''
Public Sub NotEqual(A As Variant, B As Variant, Optional Message As String = _
    "Expected {0} to not equal {1}")

    Check Not IsDeepEqual(A, B), Message, A, B
End Sub

''
' Check if a value is "truthy"
'
' From https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/if-then-else-statement
'
' Must evaluate to True or False, or to a data type that is implicitly convertible to Boolean.
' If the expression is a Nullable Boolean variable that evaluates to Nothing, the condition is treated as if the expression is False.
''
Public Sub IsOk(Value As Variant, Optional Message As String = _
    "Expected {0} to be ok")

    Check Value, Message, Value
End Sub

''
' Check if a value is not "truthy" (See .IsOk)
''
Public Sub NotOk(Value As Variant, Optional Message As String = _
    "Expected {0} to not be ok")

    Check Not CBool(Value), Message, Value
End Sub

''
' Check if a value is "undefined": Nothing, Empty, Null, or Missing
''
Public Sub IsUndefined(Optional Value As Variant, Optional Message As String = _
    "Expected {0} to be undefined")

    Check IsNothing(Value) Or VBA.IsEmpty(Value) Or VBA.IsNull(Value) Or VBA.IsMissing(Value), Message, Value
End Sub

''
' Check if a value is not "undefined": Nothing, Empty, Null, or Missing
''
Public Sub NotUndefined(Value As Variant, Optional Message As String = _
    "Expected {0} to not be undefined")

    Check Not IsNothing(Value) And Not VBA.IsEmpty(Value) And Not VBA.IsNull(Value) And Not VBA.IsMissing(Value), Message, Value
End Sub

''
' Check if the current Err value contains an error with values (if given)
''
Public Sub IsError(Optional Number As Long, Optional Source As String, Optional Description As String, Optional Message As String = _
    "Expected {0} to be an error (with Number = {1}, Source = {2}, Description = {3}")
    
    If Err.Number = 0 Then
        pFailures.Add FormatMessage(Message, "[Error Number=0]", Number, Source, Description)
        Exit Sub
    End If
    
    Check (Number = 0 Or Err.Number = Number) _
        And (Source = "" Or Err.Source = Source) _
        And (Description = "" Or Err.Description = Description), Message, FormattedErr, Number, Source, Description
End Sub

''
' Check if the current Err value does not contain an error
''
Public Sub NotError(Optional Message As String = "Expected {0} to not be an error")
    Check Err.Number = 0, Message, FormattedErr
End Sub

''
' Check if a value is included in an arbitrarily nested Array or Collection
''
Public Sub Includes(Values As Variant, Value As Variant, Optional Message As String = _
    "Expected {1} to be included in {0}")

    If IsCollection(Values) Then
        Check CollectionIncludes(Values, Value), Message, Values, Value
    ElseIf IsArray(Values) Then
        Check ArrayIncludes(Values, Value), Message, Values, Value
    Else
        pFailures.Add FormatMessage(Message, Values, Value) & " (Incompatible type for Values)"
    End If
End Sub

''
' Check if a value is not included in an arbitrarily nested Array or Collection
''
Public Sub NotIncludes(Values As Variant, Value As Variant, Optional Message As String = _
    "Expected {1} not to be included in {0}")
    
    If IsCollection(Values) Then
        Check Not CollectionIncludes(Values, Value), Message, Values, Value
    ElseIf IsArray(Values) Then
        Check Not ArrayIncludes(Values, Value), Message, Values, Value
    Else
        pFailures.Add FormatMessage(Message, Values, Value) & " (Incompatible type for Values)"
    End If
End Sub

''
' Check if two values are approximately equal, up to the given amount of significant figures
'
' ```vba
' .IsApproximate 1.001, 1.002, 3
'
' ' Equivalent to .IsEqual 1.00e+0, 1.00e+0
' ```
''
Public Sub IsApproximate(A As Variant, B As Variant, SignificantFigures As Integer, Optional Message As String = _
    "Expected {0} to be approximately equal to {1} (with {2} significant figures of precision)")

    If SignificantFigures < 1 Or SignificantFigures > 15 Then
        pFailures.Add "IsApproximate can only compare from 1 to 15 significant figures"
    Else
        Check IsApproximatelyEqual(A, B, SignificantFigures), Message, A, B, SignificantFigures
    End If
End Sub

''
' Check if two values are approximately equal, up to the given amount of significant figures
'
' ```vba
' .NotApproximate 1.001, 1.009, 3
'
' ' Equivalent to .IsEqual 1.00e+0, 1.01e+0
' ```
''
Public Sub NotApproximate(A As Variant, B As Variant, SignificantFigures As Integer, Optional Message As String = _
    "Expected {0} to not be approximately equal to {1} (with {2} significant figures of precision)")

    If SignificantFigures < 1 Or SignificantFigures > 15 Then
        pFailures.Add "NotApproximate can only compare from 1 to 15 significant figures"
    Else
        Check Not IsApproximatelyEqual(A, B, SignificantFigures), Message, A, B, SignificantFigures
    End If
End Sub

''
' Mark the test as passing
''
Public Sub Pass()
    Me.Successes = 1
    Set pFailures = New VBA.Collection
End Sub

''
' Mark the test as failing
''
Public Sub Fail(Optional Message As String = _
    "Test failed unexpectedly")
    
    pFailures.Add Message
End Sub

''
' Set the planned number of assertions for the test
''
Public Sub Plan(Count As Long)
    Planned = Count
End Sub

''
' Mark the test as skipped
''
Public Sub Skip()
    Me.Skipped = True
End Sub

' ============================================= '

Private Sub Check(Assertion As Variant, Message As String, ParamArray Values() As Variant)
    If Assertion Then
        Me.Successes = Me.Successes + 1
    Else
        pFailures.Add FormatMessage(Message, Values)
    End If
End Sub

Private Function IsDeepEqual(A As Variant, B As Variant) As Boolean
    Dim AType As VbVarType
    Dim BType As VbVarType
    
    AType = VBA.VarType(A)
    BType = VBA.VarType(B)

    If VBA.IsError(A) Or VBA.IsError(B) Then
        IsDeepEqual = False
        
    ElseIf VBA.IsArray(A) And VBA.IsArray(B) Then
        IsDeepEqual = IsArrayEqual(A, B)
    
    ElseIf AType = VBA.vbObject Or BType = VBA.vbObject Then
        If AType <> BType Or VBA.TypeName(A) <> VBA.TypeName(B) Then
            IsDeepEqual = False
        ElseIf VBA.TypeName(A) = "Collection" Then
            IsDeepEqual = IsCollectionEqual(A, B)
        ElseIf VBA.TypeName(A) = "Dictionary" Then
            IsDeepEqual = IsDictionaryEqual(A, B)
        Else
            IsDeepEqual = A Is B
        End If
    
    ElseIf VBA.VarType(A) = VBA.vbDouble Or VBA.VarType(B) = VBA.vbDouble Then
        ' It is inherently difficult/almost impossible to check equality of Double
        ' http://support.microsoft.com/kb/78113
        '
        ' -> Compare up to 15 significant figures
        IsDeepEqual = IsApproximatelyEqual(A, B, 15)
    
    Else
        IsDeepEqual = A = B
    End If
End Function

Private Function IsArrayEqual(A As Variant, B As Variant) As Boolean
    If UBound(A) <> UBound(B) Then
        IsArrayEqual = False
        Exit Function
    End If
    
    Dim i As Long
    For i = LBound(A) To UBound(A)
        If Not IsDeepEqual(A(i), B(i)) Then
            IsArrayEqual = False
            Exit Function
        End If
    Next i
    
    IsArrayEqual = True
End Function

Private Function IsCollectionEqual(A As Variant, B As Variant) As Boolean
    If A.Count <> B.Count Then
        IsCollectionEqual = False
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To A.Count
        If Not IsDeepEqual(A(i), B(i)) Then
            IsCollectionEqual = False
            Exit Function
        End If
    Next i

    IsCollectionEqual = True
End Function

Private Function IsDictionaryEqual(A As Variant, B As Variant) As Boolean
    If UBound(A.Keys) <> UBound(B.Keys) Then
        IsDictionaryEqual = False
        Exit Function
    End If
    
    Dim AKeys As Variant
    Dim BKeys As Variant
    Dim i As Long
    
    AKeys = A.Keys
    BKeys = B.Keys
    
    For i = LBound(AKeys) To UBound(AKeys)
        If AKeys(i) <> BKeys(i) Or A.Item(AKeys(i)) <> B.Item(BKeys(i)) Then
            IsDictionaryEqual = False
            Exit Function
        End If
    Next i

    IsDictionaryEqual = True
End Function

Private Function IsCollection(Value As Variant) As Boolean
    IsCollection = VBA.VarType(Value) = VBA.vbObject And VBA.TypeName(Value) = "Collection"
End Function

Private Function IsNothing(Value As Variant) As Boolean
    If VBA.IsObject(Value) Then
        IsNothing = Value Is Nothing
    Else
        IsNothing = False
    End If
End Function

Private Function ArrayIncludes(Values As Variant, Value As Variant) As Boolean
    Dim i As Long
    For i = LBound(Values) To UBound(Values)
        If VBA.IsArray(Values(i)) Then
            If ArrayIncludes(Values(i), Value) Then
                ArrayIncludes = True
                Exit Function
            End If
        ElseIf IsCollection(Values(i)) Then
            If CollectionIncludes(Values(i), Value) Then
                ArrayIncludes = True
                Exit Function
            End If
        ElseIf IsDeepEqual(Values(i), Value) Then
            ArrayIncludes = True
            Exit Function
        End If
    Next i

    ArrayIncludes = False
End Function

Private Function CollectionIncludes(Values As Variant, Value As Variant) As Boolean
    Dim Item As Variant
    For Each Item In Values
        If VBA.IsArray(Item) Then
            If ArrayIncludes(Item, Value) Then
                CollectionIncludes = True
                Exit Function
            End If
        ElseIf IsCollection(Item) Then
            If CollectionIncludes(Item, Value) Then
                CollectionIncludes = True
                Exit Function
            End If
        ElseIf IsDeepEqual(Item, Value) Then
            CollectionIncludes = True
            Exit Function
        End If
    Next Item
    
    CollectionIncludes = False
End Function

Private Function IsApproximatelyEqual(A As Variant, B As Variant, SignificantFigures As Integer) As Boolean
    If SignificantFigures < 1 Or SignificantFigures > 15 Or VBA.IsError(A) Or VBA.IsError(B) Then
        IsApproximatelyEqual = False
        Exit Function
    End If
    
    Dim AValue As String
    Dim BValue As String
    
    AValue = VBA.Format$(A, VBA.Left$("0.00000000000000", SignificantFigures + 1) & IIf(A > 1, "e+0", "e-0"))
    BValue = VBA.Format$(B, VBA.Left$("0.00000000000000", SignificantFigures + 1) & IIf(B > 1, "e+0", "e-0"))
    
    IsApproximatelyEqual = AValue = BValue
End Function

Private Function FormatMessage(Message As String, ParamArray Values() As Variant) As String
    Dim Value As Variant
    Dim Index As Long
    
    FormatMessage = Message
    For Each Value In IIf(VBA.IsArray(Values(0)), Values(0), Values)
        FormatMessage = VBA.Replace(FormatMessage, "{" & Index & "}", PrettyPrint(Value))
        Index = Index + 1
    Next Value
End Function

Private Function PrettyPrint(Value As Variant, Optional Indentation As Long = 0) As String
    If VBA.IsMissing(Value) Then
        PrettyPrint = "[Missing]"
        Exit Function
    End If
    
    Dim i As Long
    Dim Indented As String
    Indented = VBA.String$(Indentation + 1, "  ")
    
    Select Case VBA.VarType(Value)
    Case VBA.vbObject
        ' Nothing
        If Value Is Nothing Then
            PrettyPrint = "[Nothing]"
        
        ' Collection
        ElseIf VBA.TypeName(Value) = "Collection" Then
            PrettyPrint = "[Collection [" & vbNewLine
            
            For i = 1 To Value.Count
                PrettyPrint = PrettyPrint & Indent(Indentation + 1) & _
                    PrettyPrint(Value(i), Indentation + 1) & _
                    IIf(i <> Value.Count, ",", "") & vbNewLine
            Next i
            
            PrettyPrint = PrettyPrint & Indent(Indentation) & "]"
        
        ' Dictionary
        ElseIf VBA.TypeName(Value) = "Dictionary" Then
            PrettyPrint = "[Dictionary {" & vbNewLine
            
            For i = LBound(Value.Keys) To UBound(Value.Keys)
                PrettyPrint = PrettyPrint & Indent(Indentation + 1) & _
                    Value.Keys(i) & ": " & _
                    PrettyPrint(Value.Item(Value.Keys(i)), Indentation + 1) & _
                    IIf(i <> Value.Count, ",", "") & vbNewLine
            Next i
            
            PrettyPrint = PrettyPrint & Indent(Indentation) & "}]"
        
        ' Object
        Else
            PrettyPrint = "[" & VBA.TypeName(Value) & "]"
        End If
        
    ' Array
    Case VBA.vbArray To VBA.vbArray + VBA.vbByte
        PrettyPrint = "[" & vbNewLine
        
        For i = LBound(Value) To UBound(Value)
            PrettyPrint = PrettyPrint & Indent(Indentation + 1) & _
                PrettyPrint(Value(i), Indentation + 1) & _
                IIf(i <> UBound(Value), ",", "") & vbNewLine
        Next i
        
        PrettyPrint = PrettyPrint & Indent(Indentation) & "]"
    
    ' Empty
    Case VBA.vbEmpty
        PrettyPrint = "[Empty]"
    
    ' Null
    Case VBA.vbNull
        PrettyPrint = "[Null]"
    
    ' String
    Case VBA.vbString
        PrettyPrint = """" & Value & """"
    
    ' Everything else
    Case Else
        PrettyPrint = CStr(Value)
    End Select
End Function

Private Function FormattedErr() As String
    Dim ErrNumberDetails As String
    
    ErrNumberDetails = IIf(Err.Number < 0, " (" & (Err.Number - vbObjectError) & " / " & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "")
    FormattedErr = "[Error Number=" & Err.Number & ErrNumberDetails & ", Source=" & Err.Source & ", Description=" & Err.Description & "]"
End Function

Private Function Indent(Optional Indentation As Long)
    Indent = VBA.String$(Indentation, "  ")
End Function

Private Sub Class_Initialize()
    Set Me.Context = New Dictionary
    Set pFailures = New VBA.Collection
End Sub

Private Sub Class_Terminate()
    Me.Suite.TestComplete Me
    Set Me.Context = Nothing
End Sub
