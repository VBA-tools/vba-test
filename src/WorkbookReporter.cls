VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' DisplayReporter v2.0.0-beta
' (c) Tim Hall - https://github.com/VBA-tools/VBA-TDD
'
' Report results to Worksheet
'
' @class DisplayReporter
' @compatibility
'   Platforms: Windows and Mac
'   Applications: Excel-only
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Const ProgressWidth As Long = 128
Private pSheet As Worksheet
Private pCount As Long
Private pTotal As Long
Private pSuites As Collection


' --------------------------------------------- '
' Public Variables
' --------------------------------------------- '

''
' Whether or not to Indent nested SpecSuites
' and by how much
'
' @Boolean IndentActive
' @Integer IndentSpaces
''
Public IndentActive As Boolean
Public IndentSpaces As Integer

''
' Whether or not to create groups of rows
'
' @Boolean CreateGroups
''
Public CreateGroups As Boolean


' ============================================= '
' Public Methods
' ============================================= '

''
' Connect the display runner to a Worksheet to output results
'
' The given Worksheet should have names for:
' - "Progress" (Shape with width)
' - "ProgressBorder" (Shape)
' - "Result" (Cell) - Cell to output overall result
' - "Output" (Cell) - First cell to output results
'
' @method ConnectTo
' @param {Worksheet} Sheet
''
Public Sub ConnectTo(Sheet As Worksheet)
    Set pSheet = Sheet
End Sub

''
' Call this at the beginning of a test run to reset the worksheet
' (pass overall number of test suites that will be run to display progress)
'
' @method Start
' @param {Long} [NumSuites = 0]
''
Public Sub Start(Optional NumSuites As Long = 0)
    pCount = 0
    pTotal = NumSuites

    ClearResults
    ShowProgress
    DisplayResult "Running"
End Sub

''
' Output the given suite
'
' @method Output
' @param {SpecSuite} Suite
''
Public Sub Output(Suite As SpecSuite)
    pCount = pCount + 1
    pSuites.Add Suite
    
    ShowProgress
    DisplayResults Suite
End Sub

''
' After outputing all suites, display overall result
'
' @method Done
''
Public Sub Done()
    Dim Failed As Boolean
    Dim Suite As SpecSuite
    For Each Suite In pSuites
        If Suite.Result = SpecResultType.Fail Then
            Failed = True
            Exit For
        End If
    Next Suite
    
    DisplayResult IIf(Failed, "FAIL", "PASS")
End Sub

' ============================================= '
' Private Functions
' ============================================= '

Private Sub ShowProgress()
    If pTotal <= 0 Then
        HideProgress
        Exit Sub
    End If

    Dim Percent As Double
    Percent = pCount / pTotal
    
    If Percent > 1 Then
        Debug.Print "WARNING: DisplayRunner has output more suites than specified in Start"
        Percent = 1
    End If

    pSheet.Shapes("Progress").Width = ProgressWidth * Percent
    pSheet.Shapes("Progress").Visible = True
    pSheet.Shapes("ProgressBorder").Visible = True
End Sub

Private Sub HideProgress()
    pSheet.Shapes("Progress").Visible = False
    pSheet.Shapes("ProgressBorder").Visible = False
End Sub

Private Sub DisplayResult(Value As String)
    With pSheet.Range("Result")
        .Font.Size = IIf(Value = "Running", 12, 14)
        .Value = Value
    End With
End Sub

Private Sub ClearResults()
    Dim StartRow As Long
    Dim StartColumn As Long
    StartRow = pSheet.Range("Output").Row
    StartColumn = pSheet.Range("Output").Column
    
    pSheet.AutoFilterMode = False
    pSheet.UsedRange.Rows.ClearOutline
    
    Dim LastRow As Long
    LastRow = StartRow
    Do While pSheet.Cells(LastRow + 1, StartColumn).Value <> ""
        LastRow = LastRow + 1
    Loop
    
    With pSheet.Range(pSheet.Cells(StartRow, StartColumn), pSheet.Cells(LastRow, StartColumn + 1))
        .Value = ""
        .Font.Bold = False
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
End Sub

Private Sub DisplayResults(Suite As SpecSuite, Optional IndentLevel As Integer = 0)
    Dim Rows As New Collection
    Dim Dividers As New Collection
    Dim Headings As New Collection
    Dim Groups As New Collection
    
    Dim DisplayPacket As Collection
    Set DisplayPacket = DisplayBuilder(Suite, IndentLevel)
    
    Set Rows = DisplayPacket("Rows")
    Set Dividers = DisplayPacket("Dividers")
    Set Headings = DisplayPacket("Headings")
    Set Groups = DisplayPacket("Groups")
    
    Dim OutputValues() As String
    Dim Row As Variant
    Dim i As Long
    ReDim OutputValues(Rows.Count - 1, 1)
    i = 0
    For Each Row In Rows
        OutputValues(i, 0) = Row(0)
        OutputValues(i, 1) = Row(1)
        i = i + 1
    Next Row
    
    Dim StartRow As Long
    Dim StartColumn As Long
    StartRow = pSheet.Rows(pSheet.UsedRange.Rows.Count).End(xlUp).Row + 1
    StartColumn = pSheet.Range("output").Column
    
    
    pSheet.Range(pSheet.Cells(StartRow, StartColumn), pSheet.Cells(StartRow + Rows.Count - 1, StartColumn + 1)).Value = OutputValues
    
    Dim Divider As Variant
    For Each Divider In Dividers
        With pSheet.Range(pSheet.Cells(StartRow + Divider, StartColumn), pSheet.Cells(StartRow + Divider, StartColumn + 1)).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Color = VBA.RGB(191, 191, 191)
            .Weight = xlThin
        End With
    Next Divider
    
    Dim Heading As Variant
    For Each Heading In Headings
        pSheet.Cells(StartRow + Heading, StartColumn).Font.Bold = True
    Next Heading
    
    If CreateGroups Then
        Dim Group As Variant
        For Each Group In Groups
            If Group(0) < Group(1) Then
                pSheet.Rows(Group(0) + StartRow & ":" & Group(1) + StartRow - 1).Group
            End If
        Next Group
    End If
End Sub

Private Function DisplayBuilder(Suite As SpecSuite, Optional IndentLevel As Integer = 0, Optional DisplayPacket As Collection) As Collection
    Dim Rows As New Collection
    Dim Dividers As New Collection
    Dim Headings As New Collection
    Dim Groups As New Collection
  
    If Not DisplayPacket Is Nothing Then
        Set Rows = DisplayPacket("Rows")
        Set Dividers = DisplayPacket("Dividers")
        Set Headings = DisplayPacket("Headings")
        Set Groups = DisplayPacket("Groups")
    End If
  
    Dim Spec As SpecDefinition
    Dim Expectation As SpecExpectation
    Dim GroupStartRow As Long
    Dim GroupEndRow As Long
  
    If Rows.Count > 0 Then
        Dividers.Add Rows.Count
    End If
    If Suite.Description <> "" Then
        Headings.Add Rows.Count
        Rows.Add Array(IndentBuilder(IndentLevel) & Suite.Description, ResultTypeToString(Suite.Result))
    End If
    GroupStartRow = Rows.Count
  
    For Each Spec In Suite.Specs
        Rows.Add Array(IndentBuilder(IndentLevel + 1) & Spec.Description, " " & ResultTypeToString(Spec.Result) & " ")
  
        For Each Expectation In Spec.FailedExpectations
            Rows.Add Array(IndentBuilder(IndentLevel + 2) & Expectation.FailureMessage, "")
        Next Expectation
    Next Spec
  
    Set DisplayPacket = New Collection
    DisplayPacket.Add Rows, "Rows"
    DisplayPacket.Add Dividers, "Dividers"
    DisplayPacket.Add Headings, "Headings"
    DisplayPacket.Add Groups, "Groups"
        
    Dim SubSuite As SpecSuite
    For Each SubSuite In Suite.Suites
        Set DisplayPacket = DisplayBuilder(SubSuite, IndentLevel + 1, DisplayPacket)
    Next SubSuite
    GroupEndRow = DisplayPacket("Rows").Count
  
    DisplayPacket("Groups").Add Array(GroupStartRow, GroupEndRow)
    Set DisplayBuilder = DisplayPacket
End Function

Private Function ResultTypeToString(ResultType As SpecResultType) As String
    Select Case ResultType
        Case SpecResultType.Pass
            ResultTypeToString = "Pass"
        Case SpecResultType.Fail
            ResultTypeToString = "Fail"
        Case SpecResultType.Pending
            ResultTypeToString = "Pending"
    End Select
End Function

Private Sub Class_Initialize()
    Set pSuites = New Collection
End Sub

Private Function Repeat(str As String, Count As Integer) As String
    Repeat = WorksheetFunction.Rept(str, Count)
End Function

Private Function IndentBuilder(level As Integer) As String
    If level = 0 Then Exit Function
    IndentBuilder = Repeat(Repeat(" ", IndentSpaces), level)
End Function

