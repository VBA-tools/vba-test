Option Explicit

Dim Args
Dim WorkbookPath
Dim RunnerType
Dim InlineModules
Dim DisplayModules
Dim Excel
Dim Workbook
Dim i, j
Dim KeepExcelOpen
Dim KeepWorkbookOpen

' Setup workbooks for import
' Optionally, pass workbook for import as argument
Set Args = Wscript.Arguments
If Args.Length > 0 Then
  WorkbookPath = Args(0)
  RunnerType = Args(1)
Else
  WorkbookPath = ""
End If

' Include all standard Excel-REST modules
DisplayModules = Array("DisplayRunner.bas", "SpecDefinition.cls", "SpecExpectation.cls", "SpecSuite.cls", "SpecHelpers.bas", "IScenario.cls", "Scenario.cls", "IWBProxy.cls", "WBProxy.cls")
InlineModules = Array("InlineRunner.bas", "SpecDefinition.cls", "SpecExpectation.cls", "SpecSuite.cls", "SpecHelpers.bas")

' Open Excel
KeepExcelOpen = OpenExcel(Excel)
Excel.Visible = True
Excel.DisplayAlerts = False

If WorkbookPath <> "" Then
  KeepWorkbookOpen = OpenWorkbook(Excel, FullPath(WorkbookPath), Workbook)

  Select Case UCase(RunnerType)
  Case "DISPLAY"
    WScript.Echo "Importing display modules for Excel-TDD into " & WorkbookPath
    ImportModules Workbook, ".\src\", DisplayModules
  Case Else
    WScript.Echo "Importing inline modules for Excel-TDD into " & WorkbookPath
    ImportModules Workbook, ".\src\", InlineModules
  End Select

  CloseWorkbook Workbook, KeepWorkbookOpen
Else
  WScript.Echo "Importing inline modules for Excel-TDD into " & "Excel-TDD - Blank - Inline.xlsm"
  KeepWorkbookOpen = OpenWorkbook(Excel, FullPath("Excel-TDD - Blank - Inline.xlsm"), Workbook)
  ImportModules Workbook, ".\src\", InlineModules
  CloseWorkbook Workbook, KeepWorkbookOpen

  WScript.Echo "Importing display modules for Excel-TDD into " & "Excel-TDD - Blank.xlsm"
  KeepWorkbookOpen = OpenWorkbook(Excel, FullPath("Excel-TDD - Blank.xlsm"), Workbook)
  ImportModules Workbook, ".\src\", DisplayModules
  CloseWorkbook Workbook, KeepWorkbookOpen

  WScript.Echo "Importing inline modules for Excel-TDD into " & "examples\Excel-TDD - Example - Inline.xlsm"
  KeepWorkbookOpen = OpenWorkbook(Excel, FullPath("examples\Excel-TDD - Example - Inline.xlsm"), Workbook)
  ImportModules Workbook, ".\src\", InlineModules
  CloseWorkbook Workbook, KeepWorkbookOpen

  WScript.Echo "Importing display modules for Excel-TDD into " & "examples\Excel-TDD - Example - Runner.xlsm"
  KeepWorkbookOpen = OpenWorkbook(Excel, FullPath("examples\Excel-TDD - Example - Runner.xlsm"), Workbook)
  ImportModules Workbook, ".\src\", DisplayModules
  CloseWorkbook Workbook, KeepWorkbookOpen
End If

CloseExcel Excel, KeepExcelOpen

Set Workbook = Nothing
Set Excel = Nothing


''
' Module helpers
' ------------------------------------ '

Function RemoveModule(Workbook, Name)
  Dim Module
  Set Module = GetModule(Workbook, Name)

  If Not Module Is Nothing Then
    Workbook.VBProject.VBComponents.Remove Module
  End If
End Function

Function GetModule(Workbook, Name)
  Dim Module
  Set GetModule = Nothing

  For Each Module In Workbook.VBProject.VBComponents
    If Module.Name = Name Then
      Set GetModule = Module
      Exit Function
    End If
  Next
End Function

Sub ImportModule(Workbook, Folder, Filename)
  If VarType(Workbook) = vbObject Then
    RemoveModule Workbook, RemoveExtension(Filename)
    Workbook.VBProject.VBComponents.Import FullPath(Folder & Filename)
  End If
End Sub

Sub ImportModules(Workbook, Folder, Filenames)
  Dim i
  For i = LBound(Filenames) To UBound(Filenames)
    ImportModule Workbook, Folder, Filenames(i)
  Next
End Sub


''
' Excel helpers
' ------------------------------------ '

Function OpenWorkbook(Excel, Path, ByRef Workbook)
  On Error Resume Next

  Set Workbook = Excel.Workbooks(GetFilename(Path))

  If Workbook Is Nothing Or Err.Number <> 0 Then
    Set Workbook = Excel.Workbooks.Open(Path)
    OpenWorkbook = False
  Else
    OpenWorkbook = True
  End If

  Err.Clear
End Function

Function OpenExcel(Excel)
  On Error Resume Next
  
  Set Excel = GetObject(, "Excel.Application")

  If Excel Is Nothing Or Err.Number <> 0 Then
    Set Excel = CreateObject("Excel.Application")
    OpenExcel = False
  Else
    OpenExcel = True
  End If

  Err.Clear
End Function

Sub CloseWorkbook(ByRef Workbook, KeepWorkbookOpen)
  If Not KeepWorkbookOpen And VarType(Workbook) = vbObject Then
    Workbook.Close True
  End If

  Set Workbook = Nothing
End Sub

Sub CloseExcel(ByRef Excel, KeepExcelOpen)
  If Not KeepExcelOpen Then
    Excel.Quit
  End If

  Set Excel = Nothing
End Sub


''
' Filesystem helpers
' ------------------------------------ '

Function FullPath(Path)
  Dim FSO
  Set FSO = CreateObject("Scripting.FileSystemObject")
  FullPath = FSO.GetAbsolutePathName(Path)
End Function

Function GetFilename(Path)
  Dim Parts
  Parts = Split(Path, "\")

  GetFilename = Parts(UBound(Parts))
End Function

Function RemoveExtension(Name)
    Dim Parts
    Parts = Split(Name, ".")
    
    If UBound(Parts) > LBound(Parts) Then
        ReDim Preserve Parts(UBound(Parts) - 1)
    End If
    
    RemoveExtension = Join(Parts, ".")
End Function
