VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WBProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' WBProxy v1.0.0
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

Private pMapping As Dictionary
Private pInstance As Workbook
Private pPath As String
Private pPassword As String


' --------------------------------------------- '
' Properties
' --------------------------------------------- '


Private Property Get IWBProxy_mapping() As Scripting.IDictionary
    If pMapping Is Nothing Then: Set pMapping = New Dictionary
    Set mapping = pMapping
End Property
Private Property Let IWBProxy_mapping(value As Scripting.IDictionary)
    Set pMapping = value
End Property

Private Property Get IWBProxy_instance() As Workbook
    Set instance = pInstance
End Property
Private Property Let IWBProxy_instance(value As Workbook)
    Set pInstance = value
End Property

''
' Get value from workbook for provided mapping key
'
' @param {Workbook} WB
' @param {String} mappingKey
' @return {Variant} value from WB
' --------------------------------------------- '

Private Property Get IWBProxy_Value(mappingKey As String) As Variant
    
    Dim cellRef As Variant
    Set cellRef = IWBProxy_CellRef(mappingKey)
    If Not cellRef Is Nothing Then
        IWBProxy_Value = cellRef.value
    End If
    
End Property

''
' Set value in workbook for provided mapping key
'
' @param {Workbook} WB
' @param {String} mappingKey
' @param {Variant} newValue Value to set in WB
' --------------------------------------------- '

Private Property Let IWBProxy_Value(mappingKey As String, value As Variant)

    Dim cellRef As Variant
    Set cellRef = IWBProxy_CellRef(mappingKey)
    If Not cellRef Is Nothing Then
        cellRef.value = value
    End If
    
End Property

''
' Get reference to cell from workbook for provided mapping key
'
' @param {Workbook} WB
' @param {String} mappingKey
' @return {Variant} cell from WB
' --------------------------------------------- '

Private Property Get IWBProxy_CellRef(mappingKey As String) As Variant
    Dim mappingValue As Dictionary
    Dim namedRangeSheetIndex As Integer

    If pMapping.Exists(mappingKey) Then
        ' If mapping contains entry for key, use it to find value
        Set mappingValue = pMapping.Item(mappingKey)
        Set IWBProxy_CellRef = pInstance.Sheets(mappingValue("sheetName")) _
            .Cells(mappingValue("row"), mappingValue("col"))
    Else
        ' Check for named range matching mapping key
        namedRangeSheetIndex = NamedRangeExists(mappingKey, pInstance)
        If namedRangeSheetIndex > 0 Then
            Set IWBProxy_CellRef = pInstance.Sheets(namedRangeSheetIndex).Range(mappingKey)
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

Private Property Let IWBProxy_CellRef(mappingKey As String, value As Variant)
    
    Dim cellRef As Variant
    cellRef = IWBProxy_CellRef(mappingKey)
    If Not IsEmpty(cellRef) Then
        Set cellRef = value
    End If

End Property


' --------------------------------------------- '
' Methods
' --------------------------------------------- '


Private Sub IWBProxy_DefineWB(path As String, Optional password As String)
    pPath = path
    pPassword = password
End Sub

Private Sub IWBProxy_LoadWB()
    Application.ScreenUpdating = False

    ' First, close workbook if it is open
    IWBProxy_CloseWB
    
    ' Open and unprotect workbook
    Set pInstance = Workbooks.Open(pPath, False, True, password:=pPassword)
    
    ' Turn screen updating off in case WB turned it on
    Application.ScreenUpdating = False
End Sub

Private Sub IWBProxy_ReloadWB()
    IWBProxy_LoadWB
End Sub

Private Sub IWBProxy_CloseWB()
    If Not pInstance Is Nothing Then
        pInstance.Close False
        Set pInstance = Nothing
    End If
End Sub

Private Sub IWBProxy_DefineMapping(sheetName As String)
    
    Dim sheetCellRow As Integer
    Dim sheetCellCol As Integer
    Dim valueRow As Integer
    Dim valueCol As Integer
    
    sheetCellRow = 4
    sheetCellCol = 1
    valueRow = 4
    valueCol = 5
    Set pMapping = New Dictionary
    
    Do While ThisWorkbook.Sheets(sheetName).Cells(sheetCellRow, sheetCellCol) <> ""
        Call pMapping.Add( _
            Trim(CStr(ThisWorkbook.Sheets(sheetName).Cells(sheetCellRow, sheetCellCol))), _
            SheetCell( _
                ThisWorkbook.Sheets(sheetName).Cells(sheetCellRow, sheetCellCol + 1), _
                ThisWorkbook.Sheets(sheetName).Cells(sheetCellRow, sheetCellCol + 2), _
                ThisWorkbook.Sheets(sheetName).Cells(sheetCellRow, sheetCellCol + 3) _
            ) _
        )
        
        sheetCellRow = sheetCellRow + 1
    Loop
    
    Do While ThisWorkbook.Sheets(sheetName).Cells(valueRow, valueCol) <> ""
        Call pMapping.Add( _
            Trim(CStr(ThisWorkbook.Sheets(sheetName).Cells(valueRow, valueCol))), _
            ThisWorkbook.Sheets(sheetName).Cells(valueRow, valueCol + 1) _
        )
        
        valueRow = valueRow + 1
    Loop
    
End Sub

Private Sub Class_Terminate()
    Set pInstance = Nothing
End Sub
