Attribute VB_Name = "Tests"
Public Sub Run(Optional Output As Variant)
    Dim Suite As New TestSuite
    Suite.Description = "vba-test"

    Dim Immediate As New ImmediateReporter
    Immediate.ListenTo Suite

    If Not IsMissing(Output) And CStr(Output) <> "" Then
        Dim Reporter As New FileReporter
        Reporter.WriteTo Output
        Reporter.ListenTo Suite
    End If

    Tests_TestSuite.RunTests Suite.Group("TestSuite")
    Tests_TestCase.RunTests Suite.Group("TestCase")
End Sub
