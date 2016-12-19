Attribute VB_Name = "Specs_SpecDefinition"
Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "SpecDefinition"
    
    Dim Reporter As New ImmediateReporter
    Reporter.ListenTo Specs

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

        .Expect(Definition.Result).ToEqual SpecResultType.Pass
    End With

    With Specs.It("should fail if any expectation fails")
        Set Definition = TestSuite.It("should fail")
        With Definition
            .Expect("A").ToEqual "A"
            .Expect(2).ToEqual 2
            .Expect("pass").ToEqual "fail"
        End With

        .Expect(Definition.Result).ToEqual SpecResultType.Fail
    End With

    With Specs.It("should contain collection of failed expectations")
        Set Definition = TestSuite.It("should fail")
        With Definition
            .Expect("A").ToEqual "A"
            .Expect(2).ToEqual 1
            .Expect("pass").ToEqual "fail"
            .Expect(True).ToEqual False
        End With

        .Expect(Definition.Result).ToEqual SpecResultType.Fail
        .Expect(Definition.FailedExpectations(1).Actual).ToEqual 2
        .Expect(Definition.FailedExpectations(1).Passed).ToEqual False
        .Expect(Definition.FailedExpectations(2).Actual).ToEqual "pass"
        .Expect(Definition.FailedExpectations(2).Passed).ToEqual False
        .Expect(Definition.FailedExpectations(3).Actual).ToEqual True
        .Expect(Definition.FailedExpectations(3).Passed).ToEqual False
    End With

    With Specs.It("should be pending if there are no expectations")
        Set Definition = TestSuite.It("pending")
        .Expect(Definition.Result).ToEqual SpecResultType.Pending
    End With
End Function
