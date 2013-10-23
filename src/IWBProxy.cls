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
' IWBProxy v1.2.2
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

Public Property Get Path() As String
End Property
Public Property Let Path(Value As String)
End Property

Public Property Get TempPath() As String
End Property
Public Property Let TempPath(Value As String)
End Property

Public Property Get Password() As String
End Property
Public Property Let Password(Value As String)
End Property

Public Property Get Mapping() As Dictionary
End Property
Public Property Set Mapping(Value As Dictionary)
End Property

Public Property Get Instance() As Workbook
End Property
Public Property Set Instance(Value As Workbook)
End Property

''
' Get value from workbook for provided mapping key
'
' @param {String} MappingKey
' @returns {Variant} Value from workbook
' --------------------------------------------- '

Public Property Get Value(MappingKey As String) As Variant
End Property

''
' Set value in workbook for provided mapping key
'
' @param {String} MappingKey
' @param {Variant} NewValue Value to set in workbook
' --------------------------------------------- '

Public Property Let Value(MappingKey As String, NewValue As Variant)
End Property

''
' Get reference to range from workbook for provided mapping key
'
' @param {String} MappingKey
' @returns {Range} Range from workbook
' --------------------------------------------- '

Public Property Get Range(MappingKey As String) As Range
End Property

''
' Set underlying range in workbook for provided mapping key
'
' @param {String} MappingKey
' @param {Variant} NewValue Value to set in workbook
' --------------------------------------------- '

Public Property Set Range(MappingKey As String, NewValue As Range)
End Property

''
' Define mapping
'
' @param {String} SheetName to load mapping from
' --------------------------------------------- '

Public Sub DefineMapping(SheetName As String)
End Sub

