VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlankIWBProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' BlankIWBProxy v1.2.3
' (c) Tim Hall - https://github.com/timhall/Excel-TDD
'
' Blank implementation of IWBProxy
'
' @dependencies
'   Microsoft Scripting Runtime
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWBProxy

Private pPath As String
Private pPassword As String
Private pInstance As Workbook
Private pMapping As Dictionary

Public Sub DefineMapping(SheetName As String)
        
End Sub

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Common to all IWBProxy implementations
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Private Sub IWBProxy_DefineMapping(SheetName As String)
    Call Me.DefineMapping(SheetName)
End Sub

Public Property Get Range(MappingKey As String) As Range
    Set Range = SpecHelpers.GetRange(Instance, Mapping, MappingKey)
End Property
Private Property Get IWBProxy_Range(MappingKey As String) As Range
    Set IWBProxy_Range = Range(MappingKey)
End Property
Public Property Set Range(MappingKey As String, Value As Range)
    Call SpecHelpers.SetRange(Instance, Mapping, MappingKey, Value)
End Property
Private Property Set IWBProxy_Range(MappingKey As String, Value As Range)
    Set Range(MappingKey) = Value
End Property

Public Property Get Value(MappingKey As String) As Variant
    Value = SpecHelpers.GetValue(Instance, Mapping, MappingKey)
End Property
Private Property Get IWBProxy_Value(MappingKey As String) As Variant
    IWBProxy_Value = Value(MappingKey)
End Property
Public Property Let Value(MappingKey As String, NewValue As Variant)
    Call SpecHelpers.SetValue(Instance, Mapping, MappingKey, NewValue)
End Property
Private Property Let IWBProxy_Value(MappingKey As String, NewValue As Variant)
    Value(MappingKey) = NewValue
End Property

Public Property Get Instance() As Workbook
    Set Instance = pInstance
End Property
Private Property Get IWBProxy_Instance() As Workbook
    Set IWBProxy_Instance = Instance
End Property
Public Property Set Instance(Value As Workbook)
    Set pInstance = Value
End Property
Private Property Set IWBProxy_Instance(Value As Workbook)
    Set Instance = Value
End Property

Public Property Get Mapping() As Scripting.IDictionary
    Set Mapping = IWBProxy_Mapping
End Property
Private Property Get IWBProxy_Mapping() As Scripting.IDictionary
    If pMapping Is Nothing Then: Set pMapping = New Dictionary
    Set IWBProxy_Mapping = pMapping
End Property
Public Property Set Mapping(Value As Scripting.IDictionary)
    Set IWBProxy_Mapping = Value
End Property
Private Property Set IWBProxy_Mapping(Value As Scripting.IDictionary)
    Set pMapping = Value
End Property

Public Property Get Password() As String
    Password = pPassword
End Property
Private Property Get IWBProxy_Password() As String
    IWBProxy_Password = Password
End Property
Public Property Let Password(Value As String)
    pPassword = Value
End Property
Private Property Let IWBProxy_Password(Value As String)
    Password = Value
End Property

Public Property Get Path() As String
    Path = pPath
End Property
Private Property Get IWBProxy_Path() As String
    IWBProxy_Path = Path
End Property
Public Property Let Path(Value As String)
    pPath = Value
End Property
Private Property Let IWBProxy_Path(Value As String)
    Path = Value
End Property

Public Property Get TempPath() As String
    TempPath = pTempPath
End Property
Public Property Get IWBProxy_TempPath() As String
    IWBProxy_TempPath = TempPath
End Property
Public Property Let TempPath(Value As String)
    pTempPath = Value
End Property
Public Property Let IWBProxy_TempPath(Value As String)
    TempPath = Value
End Property

Private Sub Class_Terminate()
    Set Me.Instance = Nothing
End Sub
