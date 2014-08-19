Attribute VB_Name = "SpecSuiteSpecs"
Dim NumBeforeCalls As Integer
Dim MostRecentArgs As Variant

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "SpecSuite"
    
    Specs.BeforeEach "Before", "A", 3.14, True
    NumBeforeCalls = 0
    
    With Specs.It("should call BeforeEach with arguments")
        .Expect(NumBeforeCalls).ToEqual 1
        .Expect(MostRecentArgs(0)).ToEqual "A"
        .Expect(MostRecentArgs(1)).ToEqual 3.14
        .Expect(MostRecentArgs(2)).ToEqual True
    End With
    
    With Specs.It("should add spec with description and id to spec collection", "Spec-Id")
        .Expect(Specs.SpecsCol.Count).ToEqual 2
        .Expect(Specs.SpecsCol(1).Description).ToEqual "should call BeforeEach with arguments"
        .Expect(Specs.SpecsCol(2).Description).ToEqual "should add spec with description and id to spec collection"
        .Expect(Specs.SpecsCol(2).Id).ToEqual "Spec-Id"
    End With
    
    InlineRunner.RunSuite Specs
End Function

Public Sub Before(Args As Variant)
    NumBeforeCalls = NumBeforeCalls + 1
    MostRecentArgs = Args
End Sub
