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
' TestCase v2.0.0-beta
' (c) Tim Hall - https://github.com/vba-tools/vba-test
'
' Verify a single test case with assertions
'
' @class TestCase
' @author tim.hall.engr@gmail.com
' @license MIT (https://opensource.org/licenses/MIT)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Events and Properties
' --------------------------------------------- '

Public Name As String
Public Context As Dictionary

Public Planned As Long
Public Successes As Long
Public Failures As VBA.Collection
Public Skipped As Boolean

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

Public Property Get Self() As TestCase
    Self = Me
End Property

' ============================================= '
' Public Methods
' ============================================= '

Public Sub IsEqual(A As Variant, B As Variant, Optional Message As String = _
    "Expected ${1} to equal ${2}")

    If IsDeepEqual(A, B) Then
        Me.Successes = Me.Successes + 1
    Else
        Me.Failures.Add FormatMessage(Message, A, B)
    End If
End Sub

Public Sub NotEqual(A As Variant, B As Variant, Optional Message As String = _
    "Expected ${1} to not equal ${2}")

    If Not IsDeepEqual(A, B) Then
        Me.Successes = Me.Successes + 1
    Else
        Me.Failures.Add FormatMessage(Message, A, B)
    End If
End Sub

Public Sub IsTrue(Value As Variant, Optional Message As String = _
    "Expected ${1} to be true")

    If Value Then
        Me.Successes = Me.Successes + 1
    Else
        Me.Failures.Add FormatMessage(Message, Value)
    End If
End Sub

Public Sub IsFalse(Value As Variant, Optional Message As String = _
    "Expected ${1} to be false")

    If Not Value Then
        Me.Successes = Me.Successes + 1
    Else
        Me.Failures.Add FormatMessage(Message, Value)
    End If
End Sub

Public Sub IsUndefined(Value As Variant, Optional Message As String = _
    "Expected ${1} to be undefined")

    If IsNothing(Value) Or VBA.IsEmpty(Value) Or VBA.IsNull(Value) Or VBA.IsMissing(Value) Then
        Me.Successes = Me.Successes + 1
    Else
        Me.Failures.Add FormatMessage(Message, Value)
    End If
End Sub

Public Sub NotUndefined(Value As Variant, Optional Message As String = _
    "Expected ${1} to not be undefined")

    If Not IsNothing(Value) And Not VBA.IsEmpty(Value) And Not VBA.IsNull(Value) And Not VBA.IsMissing(Value) Then
        Me.Successes = Me.Successes + 1
    Else
        Me.Failures.Add FormatMessage(Message, Value)
    End If
End Sub

Public Sub Includes(Values As Variant, Value As Variant, Optional Message As String = _
    "Expected ${2} to be included in ${1}")

    If isCollection(Values) Then
        If CollectionIncludes(Values, Value) Then
            Me.Successes = Me.Successes + 1
        Else
            Me.Failures.Add FormatMessage(Message, Values, Value)
        End If
    ElseIf IsArray(Values) Then
        If ArrayIncludes(Values, Value) Then
            Me.Successes = Me.Successes + 1
        Else
            Me.Failures.Add FormatMessage(Message, Values, Value)
        End If
    Else
        Me.Failures.Add FormatMessage(Message, Values, Value) & " (Incompatible type for Values)"
    End If
End Sub

Public Sub IsApproximate(A As Variant, B As Variant, SignificantFigures As Integer, Optional Message As String = _
    "Expected ${1} to be approximately equal to ${2} (with ${3} significant figures of precision)")

    If SignificantFigures < 1 Or SignificantFigures > 15 Then
        Me.Failures.Add "IsApproximate can only compare from 1 to 15 significant figures"
    ElseIf IsApproximatelyEqual(A, B, SignificantFigures) Then
        Me.Successes = Me.Successes + 1
    Else
        Me.Failures.Add FormatMessage(Message, A, B, SignificantFigures)
    End If
End Sub

Public Sub Pass()
    Me.Successes = 1
    Set Me.Failures = New Collection
End Sub

Public Sub Fail(Optional Message As String = _
    "Test failed unexpectedly")
    
    Me.Failures.Add Message
End Sub

Public Sub Plan(Count As Long)
    Planned = Count
End Sub

Public Sub Skip()
    Me.Skipped = True
End Sub

' ============================================= '
' Private Functions
' ============================================= '

Private Function IsDeepEqual(A As Variant, B As Variant) As Boolean
    If VBA.VarType(A) <> VBA.VarType(B) Then
        IsDeepEqual = False
        Exit Function
    End If
    
    Select Case VBA.VarType(A)
    Case VBA.vbError
        IsDeepEqual = False
    
    Case VBA.vbArray To VBA.vbArray + VBA.vbByte
        IsDeepEqual = IsArrayEqual(A, B)
    
    Case VBA.vbObject
        If VBA.TypeName(A) <> VBA.TypeName(B) Then
            IsDeepEqual = False
        ElseIf VBA.TypeName(A) = "Collection" Then
            IsDeepEqual = IsCollectionEqual(A, B)
        ElseIf VBA.TypeName(A) = "Dictionary" Then
            IsDeepEqual = IsDictionaryEqual(A, B)
        Else
            IsDeepEqual = A Is B
        End If
    
    Case VBA.vbDouble
        ' It is inherently difficult/almost impossible to check equality of Double
        ' http://support.microsoft.com/kb/78113
        '
        ' -> Compare up to 15 significant figures
        IsDeepEqual = IsApproximatelyEqual(A, B, 15)
    
    Case Else
        IsDeepEqual = A = B
    End Select
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
    
    Dim i As Long
    For i = LBound(A.Keys) To UBound(A.Keys)
        If A.Keys(i) <> B.Keys(i) Or A.Item(A.Keys(i)) <> B.Item(B.Keys(i)) Then
            IsDictionaryEqual = False
            Exit Function
        End If
    Next i

    IsDictionaryEqual = True
End Function

Private Function isCollection(Value As Variant) As Boolean
    isCollection = VBA.VarType(Value) = VBA.vbObject And VBA.TypeName(Value) = "Collection"
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
        ElseIf isCollection(Values(i)) Then
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
        ElseIf isCollection(Item) Then
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
    For Each Value In Values
        Index = Index + 1
        FormatMessage = VBA.Replace(FormatMessage, "${" & Index + "}", PrettyPrint(Value))
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
                PrettyPrint = PrettyPrint & Indented & _
                    PrettyPrint(Value(i), Indentation + 1) & _
                    IIf(i + 1 <> Value.Count, ",", "") & vbNewLine
            Next i
            
            PrettyPrint = PrettyPrint & "]"
        
        ' Dictionary
        ElseIf VBA.TypeName(Value) = "Dictionary" Then
            PrettyPrint = "[Dictionary {" & vbNewLine
            
            For i = LBound(Value.Keys) To UBound(Value.Keys)
                PrettyPrint = PrettyPrint & Indented & _
                    Value.Keys(i) & ": " & _
                    PrettyPrint(Value.Item(Value.Keys(i)), Indentation + 1) & _
                    IIf(i + 1 <> Value.Count, ",", "") & vbNewLine
            Next i
            
            PrettyPrint = PrettyPrint & "}]"
        
        ' Object
        Else
            PrettyPrint = "[" & VBA.TypeName(Value) & " (Object)]"
        End If
        
    ' Array
    Case VBA.vbArray To VBA.vbArray + VBA.vbByte
        PrettyPrint = "[" & vbNewLine
        
        For i = LBound(Value) To UBound(Value)
            PrettyPrint = PrettyPrint & Indented & _
                PrettyPrint(Value(i), Indentation + 1) & _
                IIf(i + 1 <> UBound(Value), ",", "") & vbNewLine
        Next i
        
        PrettyPrint = PrettyPrint & "]"
    
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

Private Sub Class_Initialize()
    Set Me.Context = New Dictionary
    Set Me.Failures = New VBA.Collection
End Sub

Private Sub Class_Terminate()
    Dim Total As Long
    Dim Message As String

    Total = Me.Successes + Me.Failures.Count
    If Me.Planned > 0 And Me.Planned <> Total Then
        Message = "Total assertions, ${1}, does not equal planned, ${1}"
        Me.Failures.Add FormatMessage(Message, Total, Me.Planned)
    End If

    Me.Suite.TestComplete Me
    Set Me.Context = Nothing
    Set Me.Failures = Nothing
End Sub
