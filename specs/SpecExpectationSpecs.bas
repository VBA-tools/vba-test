Attribute VB_Name = "SpecExpectationSpecs"
Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "SpecExpectation"
    
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
        .Expect(3.1415926).ToNotBeCloseTo 2.78, 2
        
        .Expect(3.1415926).ToBeCloseTo 2.78, 0
    End With
    
    With Specs.It("ToContain")
        .Expect(Array("A", "B", "C")).ToContain "B"
        
        Dim Test3 As New Collection
        Test3.Add "A"
        Test3.Add "B"
        Test3.Add "C"
        .Expect(Test3).ToContain "B"
        
        .Expect(Array("A", "B", "C")).ToNotContain "D"
        .Expect(Test3).ToNotContain "D"
    End With
    
    With Specs.It("ToMatch")
        .Expect("abcde").ToMatch "bcd"
        
        .Expect("abcde").ToNotMatch "xyz"
    End With
    
    InlineRunner.RunSuite Specs
End Function
