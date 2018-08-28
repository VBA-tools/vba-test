# vba-test

vba-test (formerly VBA-TDD) adds testing to VBA on Windows and Mac.

## Example

```vb
Function AddTests() As TestSuite
  Set AddTests = New TestSuite
  AddTests.Description = "Add"

  ' Report results to the Immediate Window
  ' (ctrl + g or View > Immediate Window)
  Dim Reporter As New ImmediateReporter
  Reporter.ListenTo AddTests

  With AddTests.Test("should add two numbers")
    .IsEqual Add(2, 2), 4
    .IsEqual Add(3, -1), 2
    .IsEqual Add(-1, -2), -3
  End With

  With AddTests.Test("should add any number of numbers")
    .IsEqual Add(1, 2, 3), 6
    .IsEqual Add(1, 2, 3, 4), 10
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

## Advanced Example

For an advanced example of what is possible with vba-test, check out the [tests for VBA-Web](https://github.com/VBA-tools/VBA-Web/tree/master/specs)

## Getting Started

1. Download the [latest release (v2.0.0-beta.2)](https://github.com/vba-tools/vba-test/releases)
2. Add `src/TestSuite.cls`, `src/TestCase.cls`, add `src/ImmediateReporter.cls` to your project
3. If you're starting from scratch with Excel, you can use `vba-test-blank.xlsm`

## TestSuite

A test suite groups tests together, runs test hooks for actions that should be run before and after tests, and is responsible for passing test results to reporters.

```vb
' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "Module Name"

' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("Test Name")
Test.IsEqual ' ...

' or create and use test using With
With Suite.Test("Test Name")
  .IsEqual '...
End With
```

__TestSuite API__

- `Description`
- `Test(Name) As TestCase`
- _Event_ `BeforeEach(Test)`
- _Event_ `Result(Test)`
- _Event_ `AfterEach(Test)`

## TestCase

A test case uses assertions to test a specific part of your application.

```vb
With Suite.Test("specific part of your application")
  .IsEqual A, B, "(optional message, e.g. result should be 12)"
  .NotEqual B, C

  .IsOk C > B
  .NotOk B > C

  .IsUndefined ' Checks Nothing, Empty, Missing, or Null
  .NotUndefined

  .Includes Array(1, 2, 3), 2
  .NotIncludes Array(1, 2, 3), 4
  .IsApproximate 1.001, 1.002, 2
  .NotApproximate 1.001, 1.009, 3

  .Pass
  .Fail "e.g. should not have gotten here" 
  .Plan 4 ' Should only be 4 assertions, more or less fails
  .Skip ' skip this test
End With

With Suite.Test("complex things")
  .IsEqual _
    ThisWorkbook.Sheets("Hidden").Visible, _
    XlSheetVisibility.xlSheetVisible
  .IsEqual _
    ThisWorkbook.Sheets("Main").Cells(1, 1).Interior.Color, _
    RGB(255, 0, 0)
End With
```

In addition to these basic assertions, custom assertions can be made by passing the `TestCase` to an assertion function

```vb
Sub ToBeWithin(Test As TestCase, Value As Variant, Min As Variant, Max As Variant)
  Dim Message As String
  Message = "Expected " & Value & " to be within " & Min & " and " & Max

  Test.IsOk Value >= Min, Message
  Test.IsOk Value <= Max, Message
End Sub

With Suite.Test("...")
  ToBeWithin(.Self, Value, 0, 100)
End With
```

__TestCase API__

- `Test.Name`
- `Test.Self` - Reference to test case (useful inside of `With`)
- `Test.Context` - `Dictionary` holding test context (useful for `BeforeEach`/`AfterEach`)
- `Test.IsEqual(A, B, [Message])`
- `Test.NotEqual(A, B, [Message])`
- `Test.IsOk(Value, [Message])`
- `Test.NotOk(Value, [Message])`
- `Test.IsUndefined(Value, [Message])`
- `Test.NotUndefined(Value, [Message])`
- `Test.Includes(Values, Value, [Message])` - Check if value is included in array or `Collection`
- `Test.NotIncludes(Values, Value, [Message])`
- `Test.IsApproximate(A, B, SignificantFigures, [Message])` - Check if two values are close to each other (useful for `Double` values)
- `Test.NotApproximate(A, B, SignificantFigures, [Message])`
- `Test.Pass()` - Explicitly pass the test
- `Test.Fail([Message])` - Explicitly fail the test
- `Test.Plan(Count)` - For tests with loops and branches, it is important to catch if any assertions are skipped or extra
- `Test.Skip()` - Notify suite to skip this test

Generally, more advanced assertions should be added with custom assertions functions (detailed above), but there are common assertions that will be added (e.g. `IsApproximate` = close within significant fixtures, `Includes` = array/collection includes value, )

## ImmediateReporter

With your tests defined, the easiest way to display the test results is with `ImmediateReporter`. This outputs results to the Immediate Window (`ctrl+g` or View > Immediate Window) and is useful for running your tests without leaving the VBA editor.

```vb
Public Function Suite As TestSuite
  Set Suite = New TestSuite
  Suite.Description = "..."

  ' Create reporter and attach it to these specs
  Dim Reporter As New ImmediateReporter
  Reporter.ListenTo Suite

  ' -> Reporter will now output results as they are generated
End Function
```

## Context / Lifecycle Hooks

`TestSuite` includes events for setup and teardown before tests and a `Context` object for passing values into tests that are properly torn down between tests.

```vb
' Class TestFixture
Private WithEvents pSuite As TestSuite

Public Sub ListenTo(Suite As TestSuite)
  Set pSuite = Suite
End Sub

Private Sub pSuite_BeforeEach(Test As TestCase)
  Test.Context.Add "fixture", New Collection
End Sub

Private Sub pSuite_AfterEach(Test As TestCase)
  ' Context is cleared automatically,
  ' but can manually cleanup here
End Sub

' Elsewhere

Dim Suite As New TestSuite

Dim Fixture As New TestFixture
Fixture.ListenTo Suite

With Suite.Test("...")
  .Context("fixture").Add "..."
End With
```
