VBA-TDD
=======

Bring the reliability of other programming realms to VBA with Test-Driven Development (TDD) for VBA on Windows and Mac.

Quick example:

```vb
Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "Add"

    ' Report results to the Immediate Window
    ' (ctrl + g or View > Immediate Window)
    Dim Reporter As New ImmediateReporter
    Reporter.ListenTo Specs

    ' Describe the desired behavior
    With Specs.It("should add two numbers")
        ' Test the desired behavior
        .Expect(Add(2, 2)).ToEqual 4
        .Expect(Add(3, -1)).ToEqual 2
        .Expect(Add(-1, -2)).ToEqual -3
    End With

    With Specs.It("should add any number of numbers")
        .Expect(Add(1, 2, 3)).ToEqual 6
        .Expect(Add(1, 2, 3, 4)).ToEqual 10
    End With
End Sub

Public Function Add(ParamArray Values() As Variant) As Double
    Dim i As Integer
    Add = 0
    
    For i = LBound(Values) To UBound(Values)
        Add = Add + Values(i)
    Next i
End Function

' Immediate Window:
'
' === Add ===
' + should add two numbers
' + should add any number of numbers
' = PASS (2 of 2 passed) =
```

For details of the process of reaching this example, see the [TDD Example](https://github.com/VBA-tools/VBA-TDD/wiki/TDD-Example)

### Advanced Example

For an advanced example of what is possible with VBA-TDD, check out the [specs for VBA-Web](https://github.com/VBA-tools/VBA-Web/tree/master/specs)

### Getting Started

1. Download the [latest release (v2.0.0-beta)](https://github.com/VBA-tools/VBA-TDD/releases)
2. Add `src/SpecSuite.cls`, `src/SpecDefinition.cls`, `src/SpecExpectation.cls`, add `src/ImmediateReporter.cls` to your project
3. If you're starting from scratch with Excel, you can use `VBA-TDD - Blank.xlsm`

### It and Expect

`It` is how you describe desired behavior and once a collection of specs is written, it should read like a list of requirements.

```vb
With Specs.It("should allow user to continue if they are authorized and up-to-date")
    ' ...
End With

With Specs.It("should show an X when the user rolls a strike")
    ' ...
End With
```

`Expect` is how you test desired behavior 

```vb
With Specs.It("should check values")
    .Expect(2 + 2).ToEqual 4
    .Expect(2 + 2).ToNotEqual 5
    .Expect(2 + 2).ToBeLessThan 7
    .Expect(2 + 2).ToBeLT 6
    .Expect(2 + 2).ToBeLessThanOrEqualTo 5
    .Expect(2 + 2).ToBeLTE 4
    .Expect(2 + 2).ToBeGreaterThan 1
    .Expect(2 + 2).ToBeGT 2
    .Expect(2 + 2).ToBeGreaterThanOrEqualTo 3
    .Expect(2 + 2).ToBeGTE 4
    .Expect(2 + 2).ToBeCloseTo 3.9, 0
End With

With Specs.It("should check Nothing, Empty, Missing, and Null")
    .Expect(Nothing).ToBeNothing
    .Expect(Empty).ToBeEmpty
    .Expect().ToBeMissing
    .Expect(Null).ToBeNull
    
    ' `ToBeUndefined` checks if it's Nothing or Empty or Missing or Null

    .Expect(Nothing).ToBeUndefined
    .Expect(Empty).ToBeUndefined
    .Expect().ToBeUndefined
    .Expect(Null).ToBeUndefined
    
    ' Classes are undefined until they are instantiated
    Dim Sheet As Worksheet
    .Expect(Sheet).ToBeNothing
    
    .Expect("Howdy!").ToNotBeUndefined
    .Expect(4).ToNotBeUndefined
    
    Set Sheet = ThisWorkbook.Sheets(1)
    .Expect(Sheet).ToNotBeUndefined
End With

With Specs.It("should test complex things")
    .Expect(ThisWorkbook.Sheets("Hidden").Visible).ToNotEqual XlSheetVisibility.xlSheetVisible
    .Expect(ThisWorkbook.Sheets("Main").Cells(1, 1).Interior.Color).ToEqual RGB(255, 0, 0)
End With
```

### ImmediateReporter

With your specs defined, the easiest way to display the test results is with `ImmediateReporter`. This outputs results to the Immediate Window (`ctrl+g` or View > Immediate Window) and is useful for running your tests without leaving the VBA editor.

```vb
Public Function Specs As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "..."

    ' Create reporter and attach it to these specs
    Dim Reporter As New ImmediateReporter
    Reporter.ListenTo Specs

    ' -> Reporter will now output results as they are generated
End Function
```

### RunMatcher

For VBA applications that support `Application.Run` (which is at least Windows Excel, Word, and Access), you can create custom expect functions with `RunMatcher`.

```vb
Public Function Specs As SpecSuite
    Set Specs = New SpecSuite

    With Specs.It("should be within 1 and 100")
        .Expect(50).RunMatcher "ToBeWithin", "to be within", 1, 100
        '       ^ Actual
        '                      ^ Public Function to call
        '                                    ^ message for matcher
        '                                                    ^ 0+ Args to pass to matcher
    End With
End Function

Public Function ToBeWithin(Actual As Variant, Args As Variant) As Variant
    If UBound(Args) - LBound(Args) < 1 Then
        ' Return string for specific failure message
        ToBeWithin = "Need to pass in upper-bound to ToBeWithin"
    Else
        If Actual >= Args(0) And Actual <= Args(1) Then
            ' Return true for pass
            ToBeWithin = True
        Else
            ' Return false for fail or custom failure message
            ToBeWithin = False
        End If
    End If
End Function
```

To avoid compilation issues on unsupported applications, the compiler constant `EnableRunMatcher` in `SpecExpectation.cls` should be set to `False`.

For more details, check out the [Wiki](https://github.com/VBA-tools/VBA-TDD/wiki)

- Design based heavily on the [Jasmine](https://jasmine.github.io/)
- Author: Tim Hall
- License: MIT
