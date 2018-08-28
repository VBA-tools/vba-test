Attribute VB_Name = "Tests_TestCase"
Public Function Tests() As TestSuite
    Set Tests = New TestSuite
    Tests.Description = "TestCase"
    
    Dim Reporter As New ImmediateReporter
    Reporter.ListenTo Tests
    
    Dim Suite As New TestSuite
    Dim Test As TestCase
    
    With Tests.Test("should pass if all assertions pass")
        Set Test = Suite.Test("should pass")
        With Test
            .IsEqual "A", "A"
            .IsEqual 2, 2
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    With Tests.Test("should fail if any assertion fails")
        Set Test = Suite.Test("should fail")
        With Test
            .IsEqual "A", "A"
            .IsEqual 2, 1
        End With
        
        .IsEqual Test.Result, TestResultType.Fail
    End With
    
    With Tests.Test("should contain collection of failures")
        Set Test = Suite.Test("should have failures")
        With Test
            .IsEqual "A", "A"
            .IsEqual 2, 1
            .IsEqual True, False
        End With
        
        .IsEqual Test.Failures(1), "Expected 2 to equal 1"
        .IsEqual Test.Failures(2), "Expected True to equal False"
    End With
    
    With Tests.Test("should be pending if there are no assertions")
        Set Test = Suite.Test("pending")
        .IsEqual Test.Result, TestResultType.Pending
    End With
    
    With Tests.Test("should skip even with failed assertions")
        Set Test = Suite.Test("skipped")
        With Test
            .IsEqual 2, 1
            .Skip
        End With
        
        .IsEqual Test.Result, TestResultType.Skipped
    End With
End Function

'
'    With Specs.It("should be pending if there are no expectations")
'        Set Definition = TestSuite.It("pending")
'        .Expect(Definition.Result).ToEqual SpecResultType.Pending
'    End With
'End Function
