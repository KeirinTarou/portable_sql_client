VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnitTestLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
'   Dependency
'       - StringFormatter
' =============================================================================

Private sf As StringFormatter

Private m_Succeeded As Boolean
Private m_Message As String
Private m_Expected As String
Private m_Actual As String
Private m_Comment As String

Public Property Get Succeeded() As Boolean
    Succeeded = m_Succeeded
End Property
Public Property Let Succeeded(ByVal a_Succeeded As Boolean)
    m_Succeeded = a_Succeeded
End Property

Public Property Get Message() As String
    Message = m_Message
End Property
Public Property Let Message(ByVal a_Message As String)
    m_Message = a_Message
End Property

Public Property Get expected() As String
    expected = m_Expected
End Property
Public Property Let expected(ByVal a_Expected As String)
    m_Expected = a_Expected
End Property

Public Property Get Actual() As String
    Actual = m_Actual
End Property
Public Property Let Actual(ByVal a_Actual As String)
    m_Actual = a_Actual
End Property

Public Property Get Comment() As String
    Comment = m_Comment
End Property

Private Sub Class_Initialize()
    Set sf = New StringFormatter
End Sub

Public Sub Init( _
            ByVal a_Expected As String, _
            ByVal a_Actual As String, _
   Optional ByVal a_Comment As String = "")
    m_Expected = a_Expected
    m_Actual = a_Actual
    m_Comment = a_Comment
End Sub

Public Function BuildMessage() As String
    Dim prefix As String
    prefix = IIf(m_Succeeded, "[PASSED]", "[FAILED]")
    
    m_Message = _
        prefix & sf.NL() & _
        sf.IND() & "Expected: " & m_Expected & sf.NL() & _
        sf.IND() & "Actual  : " & m_Actual & sf.NL()
        
    If m_Comment <> "" Then
        m_Message = m_Message & sf.IND() & "Comment : " & m_Comment
    End If
    BuildMessage = m_Message
End Function

' 1件分のログを文字列化して返す
Public Function ToString( _
   Optional ByVal a_MaxRecords = 1) As String
    Dim ret As String
    ret = "Succeeded=" & CStr(m_Succeeded) & ";" & _
          "Expected=" & CStr(m_Expected) & ";" & _
          "Actual=" & CStr(m_Actual)
    ' コメントがある場合は追記
    If m_Comment <> "" Then _
        ret = ret & ";" & "Comment=" & m_Comment
    ToString = ret
End Function

