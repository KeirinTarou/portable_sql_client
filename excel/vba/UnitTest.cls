VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnitTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' =============================================================================
'   Dependency
'       - Stringifier
'       - StringFormatter
'       - UnitTestLog
' =============================================================================

Const ERR_BASE_NUM As Long = vbObjectError + 5000

Private Enum ErrorTypeEnum
    etMIN = 0
    etMAX = 10
End Enum

Public Enum TestResultTristate
    utTest_Undefined = -1
    utTestFailed = 0
    utTestPassed = 1
End Enum

Private sf As New StringFormatter
Private strf As New Stringifier
Private m_TestLogs As Collection

Private m_HasFailed As Boolean

Public Property Get TestLogs() As Collection
    Set TestLogs = m_TestLogs
End Property

Public Property Get HasFailed() As Boolean
    HasFailed = m_HasFailed
End Property

Public Property Get LastResult() As TestResultTristate
    Dim lastLog As UnitTestLog
    ' テスト未実施 -> Exit
    If m_TestLogs.Count = 0 Then _
        LastResult = utTest_Undefined: Exit Property
    Set lastLog = m_TestLogs.Item(m_TestLogs.Count)
    LastResult = IIf(lastLog.Succeeded, utTestPassed, utTestFailed)
End Property

Private Sub Class_Initialize()
    Set m_TestLogs = New Collection
    ' 不要だが明示しておく
    m_HasFailed = False
End Sub

Public Sub AssertEqual( _
            ByVal a_Expected As Variant, _
            ByVal a_Actual As Variant, _
   Optional ByVal a_Comment As String = "")
    ' 引数をシリアル化
    Dim e As String, a As String
    e = strf.ToString(a_Expected, False)
    a = strf.ToString(a_Actual, False)
    
    Call AssertCore((e = a), e, a, a_Comment)
End Sub

Public Sub AssertNotEqual( _
            ByVal a_Expected As Variant, _
            ByVal a_Actual As Variant, _
   Optional ByVal a_Comment As String = "")
    ' 引数をシリアル化
    Dim e As String, a As String
    e = strf.ToString(a_Expected, False)
    a = strf.ToString(a_Actual, False)

    Call AssertCore((e <> a), e, a, a_Comment)
End Sub

Public Sub AssertTrue( _
            ByVal a_Actual As Boolean, _
   Optional ByVal a_Comment As String = "")
    Call AssertCore( _
        a_Result:=(a_Actual = True), _
        a_ExpectedSerial:="True", _
        a_ActualSerial:=CStr(a_Actual), _
        a_Comment:=a_Comment)
End Sub

Public Sub AssertFalse( _
            ByVal a_Actual As Boolean, _
   Optional ByVal a_Comment As String = "")
    Call AssertCore( _
        a_Result:=(a_Actual = False), _
        a_ExpectedSerial:="False", _
        a_ActualSerial:=CStr(a_Actual), _
        a_Comment:=a_Comment)
End Sub

Public Sub AssertError( _
            ByVal a_ExpectedErrNumber As Long, _
            ByVal a_ActualErrNumber As Long, _
   Optional ByVal a_Comment As String = "")
        
    ' UnitTestLogインスタンスを用意
    Dim result As Boolean
    result = (a_ExpectedErrNumber = a_ActualErrNumber)
    Dim testLog As UnitTestLog
    Set testLog = _
        BuildTestLog( _
            a_Result:=result, _
            a_ExpectedSerial:=CStr(a_ExpectedErrNumber), _
            a_ActualSerial:=CStr(a_ActualErrNumber), _
            a_Comment:=a_Comment)
    ' テストが失敗していたとき
    If Not result Then m_HasFailed = True
    ' UnitTestLogインスタンスをコレクションに追加
    Call m_TestLogs.Add(Item:=testLog)
End Sub

' テストログをクリア
Public Sub ClearTestLogs()
    Set m_TestLogs = New Collection
End Sub

' =============================================================================
'   Helper Functions
' =============================================================================
Private Sub AssertCore( _
            ByVal a_Result As Boolean, _
            ByVal a_ExpectedSerial As String, _
            ByVal a_ActualSerial As String, _
            ByVal a_Comment As String)
    ' UnitTestLogインスタンス作成
    Dim testLog As UnitTestLog
    Set testLog = _
        BuildTestLog( _
            a_Result:=a_Result, _
            a_ExpectedSerial:=a_ExpectedSerial, _
            a_ActualSerial:=a_ActualSerial, _
            a_Comment:=a_Comment)
    ' UnitTestLogインスタンスをコレクションに追加
    ' テストが失敗していたとき
    If Not a_Result Then m_HasFailed = True
    Call m_TestLogs.Add(Item:=testLog)
End Sub

Private Function BuildTestLog( _
            a_Result As Boolean, _
            a_ExpectedSerial As String, _
            a_ActualSerial As String, _
            a_Comment As String) As UnitTestLog
    Dim utl As UnitTestLog
    Set utl = New UnitTestLog
    
    ' メッセージの共通部分を作成
    Call utl.Init( _
        a_Expected:=a_ExpectedSerial, _
        a_Actual:=a_ActualSerial, _
        a_Comment:=a_Comment)
    
    ' テスト結果をセット
    utl.Succeeded = a_Result
    ' メッセージ組み立て
    Call utl.BuildMessage

    Set BuildTestLog = utl
End Function
