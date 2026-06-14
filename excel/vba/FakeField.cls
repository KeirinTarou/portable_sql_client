VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum DataTypeEnum
    adInteger = 3
    adDouble = 5
    adDBTimeStamp = 135
    adVarChar = 200
    adVarWChar = 202
End Enum

Private m_Name As String
Private m_RaiseNameError As Boolean
Private m_Value As Variant
Private m_RaiseValueError As Boolean

Public Property Get Value() As Variant
    If m_RaiseValueError Then
        Call Err.Raise( _
            Number:=vbObjectError + 6000 + 1, _
            Source:="FakeField@Value(Get)", _
            Description:="FieldオブジェクトValueプロパティアクセスエラー")
    End If
    Value = m_Value
End Property
Public Property Let Value(a_Value As Variant)
    m_Value = a_Value
End Property

Public Property Get Name() As String
    If m_RaiseNameError Then
        Call Err.Raise( _
            Number:=vbObjectError + 6000 + 2, _
            Source:="FakeField@Name(Get)", _
            Description:="FieldオブジェクトNameプロパティアクセスエラー")
    End If
    Name = m_Name
End Property
Public Property Let Name(ByVal a_Value As String)
    m_Name = a_Value
End Property

Public Sub SetRaiseValueError( _
   Optional ByVal a_Raise As Boolean = True)
    m_RaiseValueError = a_Raise
End Sub

Public Sub SetRaiseNameError( _
   Optional ByVal a_Raise As Boolean = True)
    m_RaiseNameError = a_Raise
End Sub

