Excel-TDD: Excel Testing Library
================================

Bring the reliability of other programming realms to Excel.

(API design based heavily on [Jasmine](http://pivotal.github.com/jasmine/))

Example:

```VB
Function GeneralSpecs(wb As IWBProxy) As SpecSuite

    Dim specs As New SpecsSuite
    
    With specs.It("should test something simple")
        ' Set up the test by setting values in the workbook
        wb.Value("NamedRangeA") = 2
        wb.Value("MappingKeyB") = 2
        
        ' Then check that it matches what is expected
        .Expect(wb.Value("Sum")).toEqual(4)
    End With
    
    With specs.It("has lots of ways to check values!")
        .Expect(2 + 2).toEqual(4)
        .Expect(2 + 2).toNotEqual(5)
        .Expect("Howdy!").toBeDefined()
        .Expect(Nothing).toBeUndefined()
        .Expect(2 + 2).toBeLessThan(10) ' Alias: .toBeLT()
        .Expect(2 + 2).toBeLessThanOrEqualTo(4) ' Alias: .toBeLTE()
        .Expect(2 + 2).toBeGreaterThan(2) ' Alias: .toBeGT()
        .Expect(2 + 2).toBeGreaterThanOrEqualTo(4) ' Alias: .toBeGTE()
    End With
    
    With specs.It("should test something complex")
        .Expect(wb.Instance().Sheets("Hidden").toNotEqual(XlSheetVisibility.xlSheetVisible)
        .Expect(wb.CellRef("Red").Interior.Color).toEqual(RGB(255,0,0))
    End With
    
    With specs.It("shouldn"t carryover between tests")
        wb.Value("A") = 4
        wb.Value("B") = 3
        .Expect(wb.Value("Sum")).toEqual(7)
    End With
    With specs.It("should be a fresh start")
        wb.Value("B") = 4
        .Expect(wb.Value("Sum")).toEqual(8) ' => False, it"s actually 0 + 4 = 4
    End With
    
    ' Finally, return the suite. Happy testing!
    Set GeneralSpecs = specs
    
End Function
```