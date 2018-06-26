Attribute VB_Name = "Specs_SpecExpectation"
Option Explicit
Public Function Specs() As SpecSuite
    Dim Expectation As SpecExpectation
    
    Set Specs = New SpecSuite
    Specs.Description = "SpecExpectation"
    
    Dim Reporter As New ImmediateReporter
    Reporter.ListenTo Specs
    
    With Specs.It("ToEqual/ToNotEqual")
        .Expect("A").ToEqual "A"
        .Expect(2).ToEqual 2
        .Expect(3.14).ToEqual 3.14
        .Expect(1.50000000000001).ToEqual 1.50000000000001
        .Expect(True).ToEqual True
        
        .Expect("B").ToNotEqual "A"
        .Expect(1).ToNotEqual 2
        .Expect(3.145).ToNotEqual 3.14
        .Expect(1.5).ToNotEqual 1.50000000000001
        .Expect(False).ToNotEqual True
    End With
    
    With Specs.It("ToEqual/ToNotEqual with Double")
        ' Compare to 15 significant figures
        .Expect(123456789012345#).ToEqual 123456789012345#
        .Expect(1.50000000000001).ToEqual 1.50000000000001
        .Expect(Val("1234567890123450")).ToEqual Val("1234567890123451")
        .Expect(Val("0.1000000000000010")).ToEqual Val("0.1000000000000011")
        
        .Expect(123456789012344#).ToNotEqual 123456789012345#
        .Expect(1.5).ToNotEqual 1.50000000000001
        .Expect(Val("1234567890123454")).ToNotEqual Val("1234567890123456")
        .Expect(Val("0.1000000000000014")).ToNotEqual Val("0.1000000000000016")
    End With
    
    With Specs.It("ToBeUndefined/ToNotBeUndefined")
        .Expect(Nothing).ToBeUndefined
        .Expect(Empty).ToBeUndefined
        .Expect(Null).ToBeUndefined
        .Expect().ToBeUndefined
        
        Dim Test As SpecExpectation
        .Expect(Test).ToBeUndefined
        
        .Expect("A").ToNotBeUndefined
        .Expect(2).ToNotBeUndefined
        .Expect(3.14).ToNotBeUndefined
        .Expect(True).ToNotBeUndefined
        
        Set Test = New SpecExpectation
        .Expect(Test).ToNotBeUndefined
    End With
    
    With Specs.It("ToBeNothing/ToNotBeNothing")
        .Expect(Nothing).ToBeNothing
        
        Dim Test2 As SpecExpectation
        .Expect(Test2).ToBeNothing
        
        .Expect(Null).ToNotBeNothing
        .Expect(Empty).ToNotBeNothing
        .Expect().ToNotBeNothing
        .Expect("A").ToNotBeNothing
        
        Set Test2 = New SpecExpectation
        .Expect(Test2).ToNotBeUndefined
    End With
    
    With Specs.It("ToBeEmpty/ToNotBeEmpty")
        .Expect(Empty).ToBeEmpty
        
        .Expect(Nothing).ToNotBeEmpty
        .Expect(Null).ToNotBeEmpty
        .Expect().ToNotBeEmpty
        .Expect("A").ToNotBeEmpty
    End With
    
    With Specs.It("ToBeNull/ToNotBeNull")
        .Expect(Null).ToBeNull
        
        .Expect(Nothing).ToNotBeNull
        .Expect(Empty).ToNotBeNull
        .Expect().ToNotBeNull
        .Expect("A").ToNotBeNull
    End With
    
    With Specs.It("ToBeMissing/ToNotBeMissing")
        .Expect().ToBeMissing
        
        .Expect(Nothing).ToNotBeMissing
        .Expect(Null).ToNotBeMissing
        .Expect(Empty).ToNotBeMissing
        .Expect("A").ToNotBeMissing
    End With
    
    With Specs.It("ToBeLessThan")
        .Expect(1).ToBeLessThan 2
        .Expect(1.49999999999999).ToBeLessThan 1.5
        
        .Expect(1).ToBeLT 2
        .Expect(1.49999999999999).ToBeLT 1.5
    End With
    
    With Specs.It("ToBeLessThanOrEqualTo")
        .Expect(1).ToBeLessThanOrEqualTo 2
        .Expect(1.49999999999999).ToBeLessThanOrEqualTo 1.5
        .Expect(2).ToBeLessThanOrEqualTo 2
        .Expect(1.5).ToBeLessThanOrEqualTo 1.5
        
        .Expect(1).ToBeLTE 2
        .Expect(1.49999999999999).ToBeLTE 1.5
        .Expect(2).ToBeLTE 2
        .Expect(1.5).ToBeLTE 1.5
    End With
    
    With Specs.It("ToBeGreaterThan")
        .Expect(2).ToBeGreaterThan 1
        .Expect(1.5).ToBeGreaterThan 1.49999999999999
        
        .Expect(2).ToBeGT 1
        .Expect(1.5).ToBeGT 1.49999999999999
    End With
    
    With Specs.It("ToBeGreaterThanOrEqualTo")
        .Expect(2).ToBeGreaterThanOrEqualTo 1
        .Expect(1.5).ToBeGreaterThanOrEqualTo 1.49999999999999
        .Expect(2).ToBeGreaterThanOrEqualTo 2
        .Expect(1.5).ToBeGreaterThanOrEqualTo 1.5
        
        .Expect(2).ToBeGTE 1
        .Expect(1.5).ToBeGTE 1.49999999999999
        .Expect(2).ToBeGTE 2
        .Expect(1.5).ToBeGTE 1.5
    End With
    
    With Specs.It("ToBeCloseTo")
        .Expect(3.1415926).ToNotBeCloseTo 2.78, 3
        
        .Expect(3.1415926).ToBeCloseTo 2.78, 1
    End With
    
    
    
    Dim CollectionABC As New Collection
    CollectionABC.Add "A"
    CollectionABC.Add "B"
    CollectionABC.Add "C"
    
    Dim CollectionBC As New Collection
    CollectionBC.Add "B"
    CollectionBC.Add "C"
    
    With Specs.It("ToContain/ToNotContain")
        .Expect(Array("A", "B", "C")).ToContain "B"
        .Expect(Array("A", "B", "C")).ToContain Array("B", "C")
        .Expect(Array("A", "B", "C")).ToContain CollectionBC
        
        .Expect(CollectionABC).ToContain "B"
        .Expect(CollectionABC).ToContain Array("B", "C")
        .Expect(CollectionABC).ToContain CollectionBC
        
        .Expect(Array("A", "B", "C")).ToNotContain "D"
        .Expect(Array("A", "B", "C")).ToNotContain Array("D", "E")
        .Expect(Array("A", "B", "C")).ToNotContain Array("C", "D")
        .Expect(Array("A", "B")).ToNotContain Array("A", "B", "C")

        .Expect(CollectionABC).ToNotContain "D"
        .Expect(CollectionABC).ToNotContain Array("D", "E")
        .Expect(CollectionBC).ToNotContain CollectionABC
        
    End With
    
    With Specs.It("ToBeIn/ToNotBeIn")
        .Expect("B").ToBeIn Array("A", "B", "C")
        .Expect(Array("B", "C")).ToBeIn Array("A", "B", "C")
        .Expect(CollectionBC).ToBeIn Array("A", "B", "C")
        
        .Expect("B").ToBeIn CollectionABC
        .Expect(Array("B", "C")).ToBeIn CollectionABC
        .Expect(CollectionBC).ToBeIn CollectionABC
        
        .Expect("D").ToNotBeIn Array("A", "B", "C")
        .Expect(Array("D", "E")).ToNotBeIn Array("A", "B", "C")
        .Expect(Array("C", "D")).ToNotBeIn Array("A", "B", "C")
        .Expect(Array("A", "B", "C")).ToNotBeIn Array("A", "B")

        .Expect("D").ToNotBeIn CollectionABC
        .Expect(Array("D", "E")).ToNotBeIn CollectionABC
        .Expect(CollectionABC).ToNotBeIn CollectionBC
    End With
    
    With Specs.It("ToMatch")
        .Expect("abcde").ToMatch "bcd"
        
        .Expect("abcde").ToNotMatch "xyz"
    End With
    
    With Specs.It("ToMatchRegEx")
        .Expect("person@place.com").ToMatchRegEx "^([a-zA-Z0-9_\-\.]+)\@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"

    End With
    
    With Specs.It("RunMatcher")
        .Expect(100).RunMatcher "Specs_SpecExpectation.ToBeWithin", "to be within", 90, 110
        .Expect(Nothing).RunMatcher "Specs_SpecExpectation.ToBeNothing", "to be nothing"
    End With
    
    With Specs.It("should set Passed")
        Set Expectation = New SpecExpectation
        Expectation.Actual = 4
        Expectation.ToEqual 4
        
        .Expect(Expectation.Passed).ToEqual True
        
        Expectation.ToEqual 3
        .Expect(Expectation.Passed).ToEqual False
    End With
    
    With Specs.It("should set FailureMessage")
        Set Expectation = New SpecExpectation
        Expectation.Actual = 4
        
        Expectation.ToEqual 4
        .Expect(Expectation.FailureMessage).ToEqual ""
        
        Expectation.ToEqual 3
        .Expect(Expectation.FailureMessage).ToEqual "Expected 4 to equal 3"
    End With
End Function

Public Function ToBeWithin(Actual As Variant, Args As Variant) As Variant
    If UBound(Args) - LBound(Args) < 1 Then
        ' Return string for specific failure message
        ToBeWithin = "Need to pass in upper-bound to ToBeWithin"
    Else
        If Actual >= Args(0) And Actual <= Args(1) Then
            ' Return true for pass
            ToBeWithin = True
        Else
            ' Return false for fail or custom failure message
            ToBeWithin = False
        End If
    End If
End Function

Public Function ToBeNothing(Actual As Variant) As Variant
    If VBA.IsObject(Actual) Then
        If Actual Is Nothing Then
            ToBeNothing = True
        Else
            ToBeNothing = False
        End If
    Else
        ToBeNothing = False
    End If
End Function
