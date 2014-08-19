Attribute VB_Name = "SpecDefinitionSpecs"
Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "SpecDefinition"
    
    Dim TestSuite As New SpecSuite
    Dim Definition As SpecDefinition
    Dim Expectation As SpecExpectation
    
    With Specs.It("should pass if all expectations pass")
        Set Definition = TestSuite.It("should pass")
        With Definition
            .Expect("A").ToEqual "A"
            .Expect(2).ToEqual 2
            .Expect("pass").ToEqual "pass"
        End With
        
        .Expect(Definition.Result).ToEqual SpecResult.Pass
    End With
    
    With Specs.It("should fail if any expectation fails")
        Set Definition = TestSuite.It("should fail")
        With Definition
            .Expect("A").ToEqual "A"
            .Expect(2).ToEqual 2
            .Expect("pass").ToEqual "fail"
        End With
        
        .Expect(Definition.Result).ToEqual SpecResult.Fail
    End With
    
    With Specs.It("should contain collection of failed expectations")
        Set Definition = TestSuite.It("should fail")
        With Definition
            .Expect("A").ToEqual "A"
            .Expect(2).ToEqual 1
            .Expect("pass").ToEqual "fail"
            .Expect(True).ToEqual False
        End With
        
        .Expect(Definition.Result).ToEqual SpecResult.Fail
        .Expect(Definition.FailedExpectations(1).ExpectValue).ToEqual 2
        .Expect(Definition.FailedExpectations(1).Result).ToEqual ExpectResult.Fail
        .Expect(Definition.FailedExpectations(2).ExpectValue).ToEqual "pass"
        .Expect(Definition.FailedExpectations(2).Result).ToEqual ExpectResult.Fail
        .Expect(Definition.FailedExpectations(3).ExpectValue).ToEqual True
        .Expect(Definition.FailedExpectations(3).Result).ToEqual ExpectResult.Fail
    End With
    
    With Specs.It("should be pending if there are no expectations")
        Set Definition = TestSuite.It("pending")
        .Expect(Definition.Result).ToEqual SpecResult.Pending
    End With
    
    InlineRunner.RunSuite Specs
End Function
