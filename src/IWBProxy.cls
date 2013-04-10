VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWBProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' IWBProxy v1.1.0
' (c) Tim Hall - https://github.com/timhall/Excel-TDD
'
' Interface for generic workbook proxies
' that adds mapping functionality to workbooks
'
' @dependencies
'   Microsoft Scripting Runtime
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '


Public Property Get Mapping() As Dictionary
End Property
Public Property Let Mapping(Value As Dictionary)
End Property

Public Property Get Instance() As Workbook
End Property
Public Property Let Instance(Value As Workbook)
End Property

''
' Get value from workbook for provided mapping key
'
' @param {String} mappingKey
' @returns {Variant} value from WB
' --------------------------------------------- '

Public Property Get Value(MappingKey As String) As Variant
End Property

''
' Set value in workbook for provided mapping key
'
' @param {String} mappingKey
' @param {Variant} newValue Value to set in WB
' --------------------------------------------- '

Public Property Let Value(MappingKey As String, NewValue As Variant)
End Property

''
' Get reference to cell from workbook for provided mapping key
'
' @param {String} mappingKey
' @returns {Variant} cell from WB
' --------------------------------------------- '

Public Property Get CellRef(MappingKey As String) As Range
End Property

''
' Set cell in workbook for provided mapping key
'
' @param {String} mappingKey
' @param {Variant} newValue Value to set in WB
' --------------------------------------------- '

Public Property Set CellRef(MappingKey As String, NewValue As Range)
End Property

''
' Define mapping
'
' @param {String} sheetName to load mapping from
' --------------------------------------------- '

Public Sub DefineMapping(SheetName As String)
End Sub

''
' Define workbook
'
' @param {String} path full path to workbook
' @param {String} [password]
' --------------------------------------------- '

Public Sub DefineWB(Path As String, Optional Password As String)
End Sub

''
' Load instance of defined workbook
' --------------------------------------------- '

Public Sub LoadWB()
End Sub

''
' Reload workbook into instance
' --------------------------------------------- '

Public Sub ReloadWB()
End Sub

''
' Reset workbook
' --------------------------------------------- '

Public Sub ResetWB()
End Sub

''
' Close current workbook instance
' --------------------------------------------- '

Public Sub CloseWB()
End Sub
