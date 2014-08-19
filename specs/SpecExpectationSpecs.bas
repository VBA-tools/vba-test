Attribute VB_Name = "SpecExpectationSpecs"
Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "SpecExpectation"
    
    With Specs.It("toEqual")
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
    
    ' TODO
    ' Separate Nothing, Empty, and Null
    ' -> Deprecate Undefined (more relevant for javascript)
    With Specs.It("toBeUndefined")
        .Expect(Nothing).ToBeUndefined
        .Expect(Empty).ToBeUndefined
        .Expect(Null).ToBeUndefined
        .Expect().ToBeUndefined
        
        Dim Sheet As Worksheet
        .Expect(Sheet).ToBeUndefined
        
        .Expect("A").ToBeDefined
        .Expect(2).ToBeDefined
        .Expect(3.14).ToBeDefined
        .Expect(True).ToBeDefined
        
        Set Sheet = ThisWorkbook.Sheets(1)
        .Expect(Sheet).ToBeDefined
    End With
    
    With Specs.It("toBeLessThan")
        .Expect(1).ToBeLessThan 2
        .Expect(1.49999999999999).ToBeLessThan 1.5
        
        .Expect(1).ToBeLT 2
        .Expect(1.49999999999999).ToBeLT 1.5
    End With
    
    With Specs.It("toBeLessThanOrEqualTo")
        .Expect(1).ToBeLessThanOrEqualTo 2
        .Expect(1.49999999999999).ToBeLessThanOrEqualTo 1.5
        .Expect(2).ToBeLessThanOrEqualTo 2
        .Expect(1.5).ToBeLessThanOrEqualTo 1.5
        
        .Expect(1).ToBeLTE 2
        .Expect(1.49999999999999).ToBeLTE 1.5
        .Expect(2).ToBeLTE 2
        .Expect(1.5).ToBeLTE 1.5
    End With
    
    With Specs.It("toBeGreaterThan")
        .Expect(2).ToBeGreaterThan 1
        .Expect(1.5).ToBeGreaterThan 1.49999999999999
        
        .Expect(2).ToBeGT 1
        .Expect(1.5).ToBeGT 1.49999999999999
    End With
    
    With Specs.It("toBeGreaterThanOrEqualTo")
        .Expect(2).ToBeGreaterThanOrEqualTo 1
        .Expect(1.5).ToBeGreaterThanOrEqualTo 1.49999999999999
        .Expect(2).ToBeGreaterThanOrEqualTo 2
        .Expect(1.5).ToBeGreaterThanOrEqualTo 1.5
        
        .Expect(2).ToBeGTE 1
        .Expect(1.5).ToBeGTE 1.49999999999999
        .Expect(2).ToBeGTE 2
        .Expect(1.5).ToBeGTE 1.5
    End With
    
    With Specs.It("toBeCloseTo")
        .Expect(3.1415926).ToNotBeCloseTo 2.78, 2
        
        .Expect(3.1415926).ToBeCloseTo 2.78, 0
    End With
    
    ' TODO
    ' toMatch for matching substring (and possibly regex)
    ' toContain is for checking if array contains element
    With Specs.It("toContain")
        .Expect("abcde").ToContain "bcd"
        
        .Expect("abcde").ToNotContain "xyz"
    End With
    
    ' TODO
    ' Add not
    
    InlineRunner.RunSuite Specs
End Function
