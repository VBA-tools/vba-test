Attribute VB_Name = "Tests_TestSuite"
Public Function Tests() As TestSuite
    Dim Suite As New TestSuite
    
    Set Tests = New TestSuite
    Tests.Description = "TestSuite"
    
    Dim Reporter As New ImmediateReporter
    Reporter.ListenTo Tests
    
    Dim Fixture As New Test_Fixture
    Fixture.ListenTo Tests
    
    With Tests.Test("should fire BeforeEach event")
        .IsEqual Fixture.BeforeEachCallCount, 1
    End With
    
    With Tests.Test("should fire Result event")
        .IsEqual Fixture.ResultCalls(1).Name, "should fire BeforeEach event"
        .IsEqual Fixture.ResultCalls(1).Result, TestResultType.Pass
    End With
    
    With Tests.Test("should fire AfterEach event")
        .IsEqual Fixture.AfterEachCallCount, 2
    End With
    
    With Tests.Test("should store specs")
        Set Suite = New TestSuite
        With Suite.Test("(pass)")
            .IsEqual 4, 4
        End With
        With Suite.Test("(fail)")
            .IsEqual 4, 3
        End With
        With Suite.Test("(pending)")
        End With
        With Suite.Test("(skipped)")
            .Skip
        End With

        .IsEqual Suite.Tests.Count, 4
        .IsEqual Suite.PassedTests.Count, 1
        .IsEqual Suite.FailedTests.Count, 1
        .IsEqual Suite.PendingTests.Count, 1
        .IsEqual Suite.SkippedTests.Count, 1
        
        .IsEqual Suite.PassedTests(1).Name, "(pass)"
        .IsEqual Suite.FailedTests(1).Name, "(fail)"
        .IsEqual Suite.PendingTests(1).Name, "(pending)"
        .IsEqual Suite.SkippedTests(1).Name, "(skipped)"
    End With

    With Tests.Test("should have overall result")
        Set Suite = New TestSuite

        .IsEqual Suite.Result, TestResultType.Pending

        With Suite.Test("(pending)")
        End With

        .IsEqual Suite.Result, TestResultType.Pending

        With Suite.Test("(pass)")
            .IsEqual 4, 4
        End With

        .IsEqual Suite.Result, TestResultType.Pass

        With Suite.Test("(fail)")
            .IsEqual 4, 3
        End With

        .IsEqual Suite.Result, TestResultType.Fail

        With Suite.Test("(pass)")
            .IsEqual 2, 2
        End With

        .IsEqual Suite.Result, TestResultType.Fail
    End With

End Function
