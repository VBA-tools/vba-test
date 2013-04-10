VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WBProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' WBProxy v1.1.0
' (c) Tim Hall - https://github.com/timhall/Excel-TDD
'
' Generic implementation of workbook proxy
'
' @dependencies
'   Microsoft Scripting Runtime
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWBProxy

Private Const MappingSheetCellsStartRow As Integer = 4
Private Const MappingSheetCellsStartCol As Integer = 1
Private Const MappingValuesStartRow As Integer = 4
Private Const MappingValuesStartCol As Integer = 5

Private pMapping As Dictionary
Private pInstance As Workbook
Private pPath As String
Private pPassword As String

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Properties
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Property Get Mapping() As Dictionary
    Set Mapping = IWBProxy_Mapping
End Property
Private Property Get IWBProxy_Mapping() As Scripting.IDictionary
    If pMapping Is Nothing Then: Set pMapping = New Dictionary
    Set IWBProxy_Mapping = pMapping
End Property
Public Property Let Mapping(Value As Dictionary)
    Set IWBProxy_Mapping = Value
End Property
Private Property Let IWBProxy_Mapping(Value As Scripting.IDictionary)
    Set pMapping = Value
End Property

Public Property Get Instance() As Workbook
    Set Instance = IWBProxy_Instance
End Property
Private Property Get IWBProxy_Instance() As Workbook
    Set IWBProxy_Instance = pInstance
End Property
Public Property Let Instance(Value As Workbook)
    IWBProxy_Instance = Value
End Property
Private Property Let IWBProxy_Instance(Value As Workbook)
    Set pInstance = Value
End Property

''
' Get value from workbook for provided mapping key
'
' @param {Workbook} WB
' @param {String} mappingKey
' @returns {Variant} value from WB
' --------------------------------------------- '

Public Property Get Value(MappingKey As String) As Variant
    Value = IWBProxy_Value(MappingKey)
End Property
Private Property Get IWBProxy_Value(MappingKey As String) As Variant
    
    Dim CellRef As Variant
    Set CellRef = IWBProxy_CellRef(MappingKey)
    If Not CellRef Is Nothing Then
        IWBProxy_Value = CellRef.Value
    End If
    
End Property

''
' Set value in workbook for provided mapping key
'
' @param {Workbook} WB
' @param {String} mappingKey
' @param {Variant} newValue Value to set in WB
' --------------------------------------------- '

Public Property Let Value(MappingKey As String, ValueVal As Variant)
    IWBProxy_Value(MappingKey) = ValueVal
End Property
Private Property Let IWBProxy_Value(MappingKey As String, Value As Variant)

    Dim CellRef As Variant
    Set CellRef = IWBProxy_CellRef(MappingKey)
    If Not CellRef Is Nothing Then
        CellRef.Value = Value
    End If
    
End Property

''
' Get reference to cell from workbook for provided mapping key
'
' @param {Workbook} WB
' @param {String} mappingKey
' @returns {Variant} cell from WB
' --------------------------------------------- '

Public Property Get CellRef(MappingKey As String) As Range
    Set CellRef = IWBProxy_CellRef(MappingKey)
End Property
Private Property Get IWBProxy_CellRef(MappingKey As String) As Range
    Dim MappingValue As Dictionary
    Dim NamedRangeSheetIndex As Integer

    If pMapping.Exists(MappingKey) Then
        ' If mapping contains entry for key, use it to find value
        Set MappingValue = pMapping.Item(MappingKey)
        Set IWBProxy_CellRef = pInstance.Sheets(MappingValue("sheetName")) _
            .Cells(MappingValue("row"), MappingValue("col"))
    Else
        ' Check for named range matching mapping key
        NamedRangeSheetIndex = NamedRangeExists(MappingKey, pInstance)
        If NamedRangeSheetIndex > 0 Then
            Set IWBProxy_CellRef = pInstance.Sheets(NamedRangeSheetIndex).Range(MappingKey)
        End If
    End If
End Property

''
' Set cell in workbook for provided mapping key
'
' @param {Workbook} WB
' @param {String} mappingKey
' @param {Variant} newValue Value to set in WB
' --------------------------------------------- '

Public Property Set CellRef(MappingKey As String, Value As Range)
    Set IWBProxy_CellRef(MappingKey) = Value
End Property
Private Property Set IWBProxy_CellRef(MappingKey As String, Value As Range)
    
    Dim CellRef As Variant
    Set CellRef = IWBProxy_CellRef(MappingKey)
    If Not IsEmpty(CellRef) Then
        Set CellRef = Value
    End If

End Property


' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Methods
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Sub DefineWB(Path As String, Optional Password As String)
    Call IWBProxy_DefineWB(Path, Password)
End Sub
Private Sub IWBProxy_DefineWB(Path As String, Optional Password As String)
    pPath = Path
    pPassword = Password
End Sub

Public Sub LoadWB()
    IWBProxy_LoadWB
End Sub
Private Sub IWBProxy_LoadWB()
    Application.ScreenUpdating = False

    ' First, close workbook if it is open
    IWBProxy_CloseWB
    
    ' Open and unprotect workbook
    ' Set pInstance = Workbooks.Open(pPath, False, True, Password:=pPassword)
    
    ' Turn screen updating off in case WB turned it on
    Application.ScreenUpdating = False
End Sub

Public Sub ReloadWB()
    IWBProxy_ReloadWB
End Sub
Private Sub IWBProxy_ReloadWB()
    IWBProxy_LoadWB
End Sub

Public Sub ResetWB()
    IWBProxy_ResetWB
End Sub
Private Sub IWBProxy_ResetWB()
    Call IWBProxy_ReloadWB
End Sub

Public Sub CloseWB()
    IWBProxy_CloseWB
End Sub
Private Sub IWBProxy_CloseWB()
    If Not pInstance Is Nothing Then
        'pInstance.Close False
        ' Set pInstance = Nothing
    End If
End Sub

Public Sub DefineMapping(SheetName As String)
    Call IWBProxy_DefineMapping(SheetName)
End Sub
Private Sub IWBProxy_DefineMapping(SheetName As String)
    
    Dim MappingSheet As Worksheet
    Dim Row As Integer
    Set pMapping = New Dictionary
    
    If SpecHelpers.SheetExists(SheetName) Then
        Set MappingSheet = ThisWorkbook.Sheets(SheetName)
        
        Row = MappingSheetCellsStartRow
        Do While MappingSheet.Cells(Row, MappingSheetCellsStartCol) <> ""
            Call pMapping.Add( _
                Trim(CStr(MappingSheet.Cells(Row, MappingSheetCellsStartCol))), _
                SheetCell( _
                    MappingSheet.Cells(Row, MappingSheetCellsStartCol + 1), _
                    MappingSheet.Cells(Row, MappingSheetCellsStartCol + 2), _
                    MappingSheet.Cells(Row, MappingSheetCellsStartCol + 3) _
                ) _
            )
            
            Row = Row + 1
        Loop
        
        Row = MappingValuesStartRow
        Do While MappingSheet.Cells(Row, MappingValuesStartCol) <> ""
            Call pMapping.Add( _
                Trim(CStr(MappingSheet.Cells(Row, MappingValuesStartCol))), _
                MappingSheet.Cells(Row, MappingValuesStartCol + 1) _
            )
            
            Row = Row + 1
        Loop
    End If
    
End Sub

Private Sub Class_Terminate()
    Set pInstance = Nothing
End Sub
