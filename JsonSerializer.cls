VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JsonSerializer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IgnoreNulls As Boolean

Private Sub Class_Initialize()
    IgnoreNulls = True
End Sub

Public Function EscapeString(ByVal str As String) As String
    str = Replace(str, "\", "\\")
    str = Replace(str, """", "\""")
    str = Replace(str, vbCrLf, "\n")
    str = Replace(str, vbCr, "\n")
    str = Replace(str, vbLf, "\n")
    EscapeString = str
End Function

Public Function ConvertToJson(ByVal obj As Object) As String
    ConvertToJson = DictionaryToJson(ClassToDictionary(obj))
End Function

Private Function ClassToDictionary(ByVal obj As Object) As Object
    If obj Is Nothing Then Exit Function
    'Dim tName As String
    'tName = TypeName(obj)
    If TypeName(obj) = "Dictionary" Then
        Set ClassToDictionary = obj
        Exit Function
    End If

    Dim propInfo As TLI.MemberInfo
    Dim classObject As Object
    Dim propName As String
    Dim propValue As Variant
    Dim ti As TLI.typeInfo
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Use typelib to inspect object properties
    
    Set ti = TLI.InterfaceInfoFromObject(obj)
    
    For Each propInfo In ti.Members
        If propInfo Is Nothing Then GoTo nextProp
        If propInfo.InvokeKind <> INVOKE_PROPERTYGET Then GoTo nextProp
    'propInfo.ReturnType.VarType
    
        propName = propInfo.name
        
        'Dim kind As Object
        'Set kind = propInfo.InvokeKind
        
        On Error GoTo dictLet
        Call SafeGetValue(propValue, CallByName(obj, propName, VbGet))
        GoTo dictNow
dictLet:
        Set propValue = CallByName(obj, propName, VbGet)
        GoTo dictNow:
        
        
dictNow:
        If Not dict.Exists(propName) Then
            ' Check if property is another class (nested class)
            If isObject(propValue) Then
                Dim childDict As Object
                If Not propValue Is Nothing Then
                    Set classObject = propValue
                    Set childDict = ClassToDictionary(classObject)
                End If
                
                dict.Add propName, childDict
            Else
                dict.Add propName, propValue
            End If
        End If
nextProp:
    Next propInfo
    
    Set ClassToDictionary = dict
End Function

Private Function ToJson(ByVal obj As Variant) As String
    Dim t As Integer
    t = VarType(obj)
    Select Case t
        Case vbString
            ToJson = """" & EscapeString(obj) & """"
        Case vbBoolean
            ToJson = IIf(obj, "true", "false")
        Case vbNull, vbEmpty
            ToJson = "null"
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbByte
            ToJson = CStr(obj)
        Case vbArray
            ToJson = ArrayToJson(obj)
        Case vbObject
            If Not obj Is Nothing Then
                ToJson = DictionaryToJson(ClassToDictionary(obj))
            Else
                ToJson = "null"
            End If
        Case Else
            Dim ub As Integer
            ub = UBound(obj)
            If ub > 0 Then
                ToJson = ArrayToJson(obj)
            Else
                Err.Raise vbObjectError + 1, "ToJson", "Unsupported data type"
            End If
    End Select
End Function

Private Function ArrayToJson(ByVal arr As Variant) As String
    'If arr Is Nothing Then Exit Function

    Dim result As String
    Dim i As Long
    
    result = "["
    For i = LBound(arr) To UBound(arr)
        Dim val As Variant
        Call SafeSetValue(val, arr(i))
        
        result = result & ToJson(val) & ","
    Next i
    
    If Len(result) > 1 Then
        result = Left(result, Len(result) - 1) ' Remove trailing comma
    End If
    
    result = result & "]"
    ArrayToJson = result
End Function

Public Sub SafeSetValue(ByRef source, ByRef value)
    
    On Error GoTo notArray
    Dim i As Integer
    Dim ub As Integer
    Dim lb As Integer
    ub = UBound(value)
    lb = LBound(value)
    ReDim source(lb To ub)
    For i = lb To ub
        source(i) = value(i)
    Next i
    Exit Sub
    
notArray:
    If isObject(value) Then
        Set source = value
    Else
        source = value
    End If
End Sub

Public Function SafeGetValue(ByRef change, ByRef source) As Variant

    Dim t As Integer
    t = VarType(source)
    
    On Error Resume Next
    Dim ub As Integer
    ub = UBound(source)
    
    If (t = vbEmpty Or t = vbNull) And ub = 0 Then Exit Function
    
    If t = vbObject Then
        Set change = source
    Else
        change = source
    End If

End Function

Private Function EncodeKey(ByVal source As String) As String
    EncodeKey = Replace(source, "_", " ")
End Function


Private Function DictionaryToJson(ByVal dict As Object) As String
    If dict Is Nothing Then Exit Function
    
    Dim result As String
    Dim key As Variant
    
    result = "{"
    
    For Each key In dict.Keys
    
        Dim json As String
        json = ToJson(dict(key))
        
        If IgnoreNulls Then
            If json <> "null" Then
                result = result & """" & EncodeKey(key) & """: " & json & ","
            End If
        Else
            result = result & """" & EncodeKey(key) & """: " & json & ","
        End If
        
    Next key
    
    If Len(result) > 1 Then
        result = Left(result, Len(result) - 1) ' Remove trailing comma
    End If
    
    result = result & "}"
    DictionaryToJson = result
End Function
