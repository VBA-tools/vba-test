Attribute VB_Name = "Specs_SpecSuite"
Public Function Specs() As SpecSuite
    Dim Suite As SpecSuite

    Set Specs = New SpecSuite
    Specs.Description = "SpecSuite"
    
    Dim Reporter As New ImmediateReporter
    Reporter.ListenTo Specs
    
    Dim Fixture As New Specs_Fixture
    Fixture.ListenTo Specs
    
    With Specs.It("should fire BeforeEach event", "id")
        .Expect(Fixture.BeforeEachCallCount).ToEqual 1
        .Expect(1 + 1).ToEqual 2
    End With
    
    With Specs.It("should fire Result event")
        .Expect(Fixture.ResultCalls(1).Description).ToEqual "should fire BeforeEach event"
        .Expect(Fixture.ResultCalls(1).Result).ToEqual SpecResultType.Pass
        .Expect(Fixture.ResultCalls(1).Expectations.Count).ToEqual 2
        .Expect(Fixture.ResultCalls(1).Id).ToEqual "id"
    End With
    
    With Specs.It("should fire AfterEach event")
        .Expect(Fixture.AfterEachCallCount).ToEqual 2
    End With
    
    With Specs.It("should store specs")
        Set Suite = New SpecSuite
        With Suite.It("(pass)", "(1)")
            .Expect(4).ToEqual 4
        End With
        With Suite.It("(fail)", "(2)")
            .Expect(4).ToEqual 3
        End With
        With Suite.It("(pending)", "(3)")
        End With
        
        .Expect(Suite.Specs.Count).ToEqual 3
        .Expect(Suite.PassedSpecs.Count).ToEqual 1
        .Expect(Suite.FailedSpecs.Count).ToEqual 1
        .Expect(Suite.PendingSpecs.Count).ToEqual 1
        
        .Expect(Suite.PassedSpecs(1).Description).ToEqual "(pass)"
        .Expect(Suite.FailedSpecs(1).Description).ToEqual "(fail)"
        .Expect(Suite.PendingSpecs(1).Description).ToEqual "(pending)"
    End With
    
    With Specs.It("should have overall result")
        Set Suite = New SpecSuite
        
        .Expect(Suite.Result).ToEqual SpecResultType.Pending
        
        With Suite.It("(pending)", "(1)")
        End With
        
        .Expect(Suite.Result).ToEqual SpecResultType.Pending
        
        With Suite.It("(pass)", "(2)")
            .Expect(4).ToEqual 4
        End With
        
        .Expect(Suite.Result).ToEqual SpecResultType.Pass
        
        With Suite.It("(fail)", "(3)")
            .Expect(4).ToEqual 3
        End With
        
        .Expect(Suite.Result).ToEqual SpecResultType.Fail
        
        With Suite.It("(pass)", "(4)")
            .Expect(4).ToEqual 4
        End With
        
        .Expect(Suite.Result).ToEqual SpecResultType.Fail
    End With
End Function
