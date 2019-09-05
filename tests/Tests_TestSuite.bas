Attribute VB_Name = "Tests_TestSuite"
Public Sub RunTests(Suite As TestSuite)
    Dim Tests As TestSuite
    Dim Fixture As New Test_Fixture
    Fixture.ListenTo Suite
    
    With Suite.Test("should fire BeforeEach event")
        .IsEqual Fixture.BeforeEachCallCount, 1
    End With
    
    With Suite.Test("should fire Result event")
        .IsEqual Fixture.ResultCalls(1).Description, "should fire BeforeEach event"
        .IsEqual Fixture.ResultCalls(1).Result, TestResultType.Pass
    End With
    
    With Suite.Test("should fire AfterEach event")
        .IsEqual Fixture.AfterEachCallCount, 2
    End With
    
    With Suite.Test("should store specs")
        Set Tests = New TestSuite
        With Tests.Test("(pass)")
            .IsEqual 4, 4
        End With
        With Tests.Test("(fail)")
            .IsEqual 4, 3
        End With
        With Tests.Test("(pending)")
        End With
        With Tests.Test("(skipped)")
            .Skip
        End With

        .IsEqual Tests.Tests.Count, 4
        .IsEqual Tests.PassedTests.Count, 1
        .IsEqual Tests.FailedTests.Count, 1
        .IsEqual Tests.PendingTests.Count, 1
        .IsEqual Tests.SkippedTests.Count, 1
        
        .IsEqual Tests.PassedTests(1).Description, "(pass)"
        .IsEqual Tests.FailedTests(1).Description, "(fail)"
        .IsEqual Tests.PendingTests(1).Description, "(pending)"
        .IsEqual Tests.SkippedTests(1).Description, "(skipped)"
    End With

    With Suite.Test("should have overall result")
        Set Tests = New TestSuite

        .IsEqual Tests.Result, TestResultType.Pending

        With Tests.Test("(pending)")
        End With

        .IsEqual Tests.Result, TestResultType.Pending

        With Tests.Test("(pass)")
            .IsEqual 4, 4
        End With

        .IsEqual Tests.Result, TestResultType.Pass

        With Tests.Test("(fail)")
            .IsEqual 4, 3
        End With

        .IsEqual Tests.Result, TestResultType.Fail

        With Tests.Test("(pass)")
            .IsEqual 2, 2
        End With

        .IsEqual Tests.Result, TestResultType.Fail
    End With

End Sub
