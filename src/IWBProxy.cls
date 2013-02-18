VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWBProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' IWBProxy v1.0.0
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


Public Property Get mapping() As Dictionary
Public Property Let mapping(value As Dictionary)

Public Property Get instance() As Workbook
Public Property Let instance(value As Workbook)

''
' Get value from workbook for provided mapping key
'
' @param {String} mappingKey
' @return {Variant} value from WB
' --------------------------------------------- '

Public Property Get value(mappingKey As String) As Variant

''
' Set value in workbook for provided mapping key
'
' @param {String} mappingKey
' @param {Variant} newValue Value to set in WB
' --------------------------------------------- '

Public Property Let value(mappingKey As String, newValue As Variant)

''
' Get reference to cell from workbook for provided mapping key
'
' @param {String} mappingKey
' @return {Variant} cell from WB
' --------------------------------------------- '

Public Property Get cellRef(mappingKey As String) As Variant

''
' Set cell in workbook for provided mapping key
'
' @param {String} mappingKey
' @param {Variant} newValue Value to set in WB
' --------------------------------------------- '

Public Property Let cellRef(mappingKey As String, newValue As Variant)

''
' Define mapping
'
' @param {String} sheetName to load mapping from
' --------------------------------------------- '

Public Sub DefineMapping(sheetName As String)

''
' Define workbook
'
' @param {String} path full path to workbook
' @param {String} [password]
' --------------------------------------------- '

Public Sub DefineWB(path As String, Optional password As String)

''
' Load instance of defined workbook
' --------------------------------------------- '

Public Sub LoadWB()

''
' Reload workbook into instance
' --------------------------------------------- '

Public Sub ReloadWB()

''
' Close current workbook instance
' --------------------------------------------- '

Public Sub CloseWB()
