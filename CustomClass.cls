VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public name As String

Private mNested As Variant ' Infinitely Nestable Object
Public test_Spaced_Name As String

Private helper As New JsonSerializer

Public Property Let nested(ByRef value)
   Call helper.SafeSetValue(mNested, value)
End Property

Public Property Get nested()
   Call helper.SafeGetValue(nested, mNested)
End Property

Private Sub Class_Initialize()
    name = "Test"
End Sub
