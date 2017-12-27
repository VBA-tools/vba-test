Attribute VB_Name = "Specs"
Public Sub RunSpecs()
    Dim Reporter As New WorkbookReporter
    Reporter.ConnectTo SpecRunner
    
    Reporter.Start NumSuites:=3
    Reporter.Output Specs_SpecDefinition.Specs
    Reporter.Output Specs_SpecExpectation.Specs
    Reporter.Output Specs_SpecSuite.Specs
    Reporter.Done
End Sub
