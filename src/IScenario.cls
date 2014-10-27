VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IScenario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' IScenario v1.4.0
' (c) Tim Hall - https://github.com/timhall/Excel-TDD
'
' Interface for creating and running scenarios on workbooks
'
' @dependencies
'   Microsoft Scripting Runtime
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

''
' Generate a spec suite to run through a scenario for the given wb
'
' @param {IWBProxy} WB to perform scenario on
' @returns {SpecSuite}
' --------------------------------------------- '

Public Function RunScenario(WB As IWBProxy) As SpecSuite
End Function

''
' Load scenario from given sheet
'
' @param {String} SheetName to load scenario from
' --------------------------------------------- '

Public Sub Load(SheetName As String)
End Sub
