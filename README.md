Excel-TDD: Excel Testing Library
================================

Bring the reliability of other programming realms to Excel with Test-Driven Development for Excel.

Quick example:

```VB
Sub Specs()
    On Error Resume Next

    ' Create a new collection of specs
    Dim Specs As New SpecSuite

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

    InlineRunner.RunSuite Specs
End Sub

Public Function Add(ParamArray Values() As Variant) As Double
    Dim i As Integer
    Add = 0
    
    For i = LBound(Values) To UBound(Values)
        Add = Add + Values(i)
    Next i
End Function

' Open the Immediate Window (Ctrl+g or View > Immediate Window) and Run Specs (F5)'
' = PASS (2 of 2 passed) ==========================
```

For details of the process of reaching this example, see the [TDD Example](https://github.com/timhall/Excel-TDD/wiki/TDD-Example)

### Advanced Example

For an advanced example of what is possible with Excel-TDD, check out the [specs for Excel-REST](https://github.com/timhall/Excel-REST/tree/master/specs)

Methods used in these specs:

- Using `BeforeEach` to reset before each spec is run
- Testing VBA modules and classes
- Setting up a custom `DisplayRunner` and `InlineRunner`
- Waiting for and handling async behavior

### Getting Started

For testing macros:

- The lightweight Inline Runner is recommended and should be added directly to the workbook that is being tested
- Add `InlineRunner.bas`, `SpecDefinition.cls`, `SpecExpectation.cls`, and `SpecSuite.cls` to your workbook
- If starting from scratch, the `Excel-TDD - Blank - Inline.xlsm` workbook includes all of the required classes and modules

For testing workbooks:

- The full Workbook Runner is recommended in order to keep testing behavior separate from the workbook that is being tested
- Use the `Excel-TDD - Blank.xlsm` workbook
- See the [Workbook Runner Example](https://github.com/timhall/Excel-TDD/wiki/Workbook-Runner-Example) for details

### Inline Runner

The inline runner is a lightweight test runner that is intended to be loaded directly into the workbook that is being tested and is for testing macros and simple behaviors in the workbook
All results are displayed in the Immediate Window (Ctrl+g or View > Immediate Window) and the runner requires no setup to run test suites

```VB
InlineRunner.RunSuite Specs

' = PASS (2 of 2 passed) ==========================

' Configurable
InlineRunner.RunSuite Specs, ShowFailureDetails:=True, ShowPassed:=True, ShowSuiteDetails:=True

' = PASS (2 of 2 passed) ==========================
' + 2 specs
'   + should add two numbers
'   + should add any number of numbers
' ===
```

### Workbook Runner

The workbook runner is a full test runner that is intended to be used separately of the workbook that is being tested to keep testing behavior separate. 
It is for testing advanced workbook behaviors and allows for reseting the test workbook between tests, using scenarios for tests (see below), and running tests against different test workbooks.
See the [Workbook Runner Example](https://github.com/timhall/Excel-TDD/wiki/Workbook-Runner-Example) for details

### It and Expect

`It` is how you describe desired behavior and once a collection of specs is written, it should read like a list of requirements.

```VB
With Specs.It("should allow user to continue if they are authorized and up-to-date")
    ' ...
End With

With Specs.It("should show an X when the user rolls a strike")
    ' ...
End With
```

`Expect` is how you test desired behavior 

```VB
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

For more details, check out the [Wiki](https://github.com/timhall/Excel-TDD/wiki)

- Design based heavily on the [Jasmine](http://pivotal.github.com/jasmine/)
- Author: Tim Hall
- License: MIT
