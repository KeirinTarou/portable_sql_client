Attribute VB_Name = "Test_FakeDataFactory"
Option Explicit

Private Const MODULE_NAME As String = "Test_FakeDataFactory"

Private Const vbext_pk_Proc As Long = 0

Private ut As UnitTest
Private sf As New StringFormatter
Private strf As New Stringifier
Private fdf As New FakeDataFactory
Private Const ERR_BASE_NUM As Long = vbObjectError + 5000

Private m_SucceededCount As Long
Private m_FailedCount As Long

Private m_TestProcNames As Object

' 区切り線 <- Setup()で内容をセット
Private BL As String
' テストプロシージャがないときの中断メッセージ
Private TEST_NOT_FOUND As String

Private Type TestRunConfig
    StopIfFailed As Boolean
    PrintLog As Boolean
End Type

Private Sub AA_EntryPoint(): End Sub
' =============================================================================
'   Entry Point
' =============================================================================

' 全てのテストを実行 -> 2026-04-19 08:35 All green!!
Public Sub StartTest()
    Dim cfg As TestRunConfig
    
    ' 失敗即Stopあり × 詳細ログ出力あり
    cfg.StopIfFailed = True
    cfg.PrintLog = True
    ' 失敗即Stopあり × 詳細ログ出力なし
'    cfg.StopIfFailed = True
'    cfg.PrintLog = False
'    ' 失敗即Stopなし × 詳細ログ出力あり
'    cfg.StopIfFailed = False
'    cfg.PrintLog = True
'    ' 失敗即Stopなし × 詳細ログ出力なし
'    cfg.StopIfFailed = False
'    cfg.PrintLog = False
    
    Call RunAllTests(cfg)
End Sub

' テストを名前を指定して実行
Public Sub StartTestByName()
    Dim cfg As TestRunConfig
    cfg.StopIfFailed = False
    ' 詳細ログ出力あり
    cfg.PrintLog = True
    ' 詳細ログ出力なし
'    cfg.PrintLog = False
    
    ' Call RunTestsByName(cfg, "Test_Error*")
    ' Call RunTestsByName(cfg, "Test_GenRow*")
    Call RunTestsByName(cfg, "Test_AddRow*")
    ' Call RunTestsByName(cfg, "Test_Error_GenRow*")
    ' Call RunTestsByName(cfg, "Test_GenScalar_Empty2Null")
End Sub

Private Sub ZZ_TestCodeHere(): End Sub
' =============================================================================
'   Test Procedures
' =============================================================================
' GenScalar()
' -----------------------------------------------------------------------------
' 正常系
' -----------------------------------------------------------------------------
' スカラ値（数値）をセット
Private Sub Test_GenScalar_IntValue()
    Dim dic As Object
    Set dic = fdf.GenScalar("pachinko", 123)
    Call ut.AssertEqual( _
        123, dic("pachinko"), _
        "スカラ値（数値）をセットする")
End Sub
' スカラ値（文字列）をセット
Private Sub Test_GenScalar_StringValue()
    Dim dic As Object
    Set dic = fdf.GenScalar("pachinko", "123")
    Call ut.AssertEqual( _
        "123", dic("pachinko"), _
        "スカラ値（文字列）をセットする")
End Sub
' Emptyをセット -> Nullに寄せる
Private Sub Test_GenScalar_Empty2Null()
    Dim dic As Object
    Set dic = fdf.GenScalar("pachinko", Empty)
    Call ut.AssertEqual( _
        Null, dic("pachinko"), _
        "EmptyはNullに寄せる")
End Sub
' Nullをセット -> そのままNull
Private Sub Test_GenScalar_Null()
    Dim dic As Object
    Set dic = fdf.GenScalar("pachinko", Null)
    Call ut.AssertEqual( _
        Null, dic("pachinko"), _
        "NullはNull")
End Sub
' Booleanはそのまま
Private Sub Test_GenScalar_Boolean()
    Dim dic As Object
    Set dic = fdf.GenScalar("pachinko", True)
    Call ut.AssertEqual( _
        True, dic("pachinko"), _
        "Booleanを渡す")
End Sub
' GenRow()
Private Sub Test_GenRow_StringIntRow()
    Dim dic As Object, e As String
    Set dic = fdf.GenRow("pachinko", "123", "slot", 123)
    e = "{""pachinko"": ""123"", ""slot"": 123}"
    Call ut.AssertEqual( _
        e, strf.ToString(dic, False), _
        "文字列・数値の1行データを渡す")
End Sub
Private Sub Test_GenRow_StringDateRow()
    Dim dic As Object, e As Date
    Set dic = fdf.GenRow("pachinko", "123", "Now()", CDate("2026/04/18 06:56:15"))
    e = CDate("2026/04/18 06:56:15")
    Call ut.AssertEqual( _
        e, dic("Now()"), _
        "文字列・日付の1行データを渡す")
End Sub
' AddRow()
' 行を2つ追加する
Private Sub Test_AddRow_Add2Rows()
    Dim dics As Variant, e As String
    'ReDim dics(0 To 0)
    dics = fdf.AddRow("pachinko", "123")
    dics = fdf.AddRow("slot", "123")
    e = "[0: {""pachinko"": ""123""}, 1: {""slot"": ""123""}]"
    Call ut.AssertEqual( _
        e, strf.ToString(dics, False), _
        "文字列の1行データを2つ追加する")
End Sub

Private Sub ZZ_Test_Error(): End Sub
' -----------------------------------------------------------------------------
' 異常系
' -----------------------------------------------------------------------------
' GenScalar()
' キーとして配列を渡す -> 13エラー
Sub Test_Error_GenScalar_InvalidKey_Array()
    On Error Resume Next
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call fdf.GenScalar(Array(), "pachinko")
    Call ut.AssertError( _
        a_ExpectedErrNumber:=13, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="キーに配列を渡した -> 「型が違います」エラーを期待")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' キーとして文字列変換できないオブジェクトを渡す -> 438
Sub Test_Error_GenScalar_InvalidKey_Object()
    On Error Resume Next
    Dim o As Object
    Set o = CreateObject("Scripting.FileSystemObject")
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    ' VBAが`o`を文字列として評価しようとする
    ' デフォルトプロパティがない
    '   -> 「オブジェクトはこのプロパティまたはメソッドを……」になる
    Call fdf.GenScalar(o, "pachinko")
    Call ut.AssertError( _
        a_ExpectedErrNumber:=438, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="文字列変換できないオブジェクトを渡した -> 438エラー")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' キーとしてEmptyを渡す
Sub Test_Error_GenScalar_InvalidKey_Empty()
    Const errBlankKey As Long = ERR_BASE_NUM + 8
    On Error Resume Next
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call fdf.GenScalar(Empty, "pachinko")
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errBlankKey, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="Emptyを渡した -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' キーとしてNullを渡す -> 94エラー
Sub Test_Error_GenScalar_InvalidKey_Null()
    On Error Resume Next
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call fdf.GenScalar(Null, "pachinko")
    Call ut.AssertError( _
        a_ExpectedErrNumber:=94, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="Nullを渡した -> 94エラー")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' キーとして空文字を渡す
Sub Test_Error_GenScalar_InvalidKey_BlankString()
    Const errBlankKey As Long = ERR_BASE_NUM + 8
    On Error Resume Next
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call fdf.GenScalar("", "pachinko")
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errBlankKey, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="空文字を渡した -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' キーとして半角スペースを渡す
Sub Test_Error_GenScalar_InvalidKey_SingleSpace()
    Const errBlankKey As Long = ERR_BASE_NUM + 8
    On Error Resume Next
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call fdf.GenScalar(" ", "pachinko")
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errBlankKey, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="半角スペースを渡した -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' キーとして半角スペースを2個渡す
Sub Test_Error_GenScalar_InvalidKey_TwoSpaces()
    Const errBlankKey As Long = ERR_BASE_NUM + 8
    On Error Resume Next
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call fdf.GenScalar("  ", "pachinko")
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errBlankKey, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="半角スペース2個を渡した -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' キーとしてタブ文字を渡す
Sub Test_Error_GenScalar_InvalidKey_Tab()
    Const errBlankKey As Long = ERR_BASE_NUM + 8
    On Error Resume Next
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call fdf.GenScalar(vbTab, "pachinko")
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errBlankKey, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="タブを渡した -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' キーとして全角スペースを渡す
Sub Test_Error_GenScalar_InvalidKey_DBSpace()
    Const errBlankKey As Long = ERR_BASE_NUM + 8
    On Error Resume Next
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call fdf.GenScalar("　", "pachinko")
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errBlankKey, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="全角スペースを渡した -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' キーとしてNBSPを渡した
Sub Test_Error_GenScalar_InvalidKey_NBSP()
    Const errBlankKey As Long = ERR_BASE_NUM + 8
    On Error Resume Next
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call fdf.GenScalar(Chr(160), "pachinko")
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errBlankKey, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="NBSPを渡した -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' 256文字のキーを渡した -> セーフ
Sub Test_Error_GenScalar_ValidKey_256Chars()
    On Error Resume Next
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call fdf.GenScalar(String(256, "a"), "pachinko")
    Call ut.AssertError( _
        a_ExpectedErrNumber:=0, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="256文字のキーを渡した -> 例外なし")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' 257文字のキーを渡した -> アウト！
Sub Test_Error_GenScalar_ValidKey_Over256Chars()
    Const errTooLongKey As Long = ERR_BASE_NUM + 9
    On Error Resume Next
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call fdf.GenScalar(String(257, "a"), "pachinko")
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errTooLongKey, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="257文字のキーを渡した -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' TODO: 禁止文字検出テストは別途作成
' キーに禁止文字が含まれている
Sub Test_Error_GenScalar_HasInvalidChar_chr_0()
    Const errHasInvalidChar As Long = ERR_BASE_NUM + 10
    On Error Resume Next
    Dim s As String
    s = "pachinko" & Chr(0)
    Call fdf.GenScalar(s, "pachinko")
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errHasInvalidChar, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="キーにChr(0)が混入 -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' キーに制御文字が含まれていない
Sub Test_Error_GenScalar_HasNoInvalidChar_chr_32()
    On Error Resume Next
    Dim s As String
    s = "pachinko" & Chr(32)
    Call fdf.GenScalar(s, "pachinko")
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call ut.AssertError( _
        a_ExpectedErrNumber:=0, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="キーにChr(32)が混入 -> 例外なし")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' 値が配列 -> カスタム例外
Sub Test_Error_GenScalar_InvalidValue_Array()
    Const errInvalidValue As Long = ERR_BASE_NUM + 4
    On Error Resume Next
    Dim a As Variant
    a = Array("pachinko", "123")
    Call fdf.GenScalar("pachinko", a)
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errInvalidValue, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="値が配列 -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' 値がオブジェクト -> カスタム例外
Sub Test_Error_GenScalar_InvalidValue_Object()
    Const errInvalidValue As Long = ERR_BASE_NUM + 4
    On Error Resume Next
    Dim o As Object
    Set o = CreateObject("Scripting.FileSystemObject")
    o = Array("pachinko", "123")
    Call fdf.GenScalar("pachinko", o)
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errInvalidValue, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="値がオブジェクト -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' GenRow()
' キー被り（小文字どうし） -> カスタム例外
Sub Test_Error_GenRow_DuplicateKey_CaseSensitive()
    Const errDuplicateKey As Long = ERR_BASE_NUM + 5
    On Error Resume Next
    Dim dic As Object
    Set dic = fdf.GenRow("pachinko", "123", "pachinko", "123")
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errDuplicateKey, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="キー重複（小文字どうし） -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' キー被り（大文字と小文字） -> カスタム例外
Sub Test_Error_GenRow_DuplicateKey_CaseInSensitive()
    Const errDuplicateKey As Long = ERR_BASE_NUM + 5
    On Error Resume Next
    Dim dic As Object
    Set dic = fdf.GenRow("pachinko", "123", "PACHINKO", "123")
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errDuplicateKey, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="キー重複（大文字と小文字） -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' 引数なしで呼び出す -> カスタム例外
Sub Test_Error_GenRow_NoArg()
    Const errNotEvenNumberOfArgs As Long = ERR_BASE_NUM + 2
    On Error Resume Next
    Dim dic As Object
    Set dic = fdf.GenRow()
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errNotEvenNumberOfArgs, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="引数なし -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' 引数が奇数個 -> カスタム例外
Sub Test_Error_GenRow_OddNumberOfArgs()
    Const errNotEvenNumberOfArgs As Long = ERR_BASE_NUM + 2
    On Error Resume Next
    Dim dic As Object
    Set dic = fdf.GenRow()
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errNotEvenNumberOfArgs, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="引数が奇数個 -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' 奇数番目の引数に非文字列 -> カスタム例外
Sub Test_Error_GenRow_OddNumberOfArgsNotString()
    Const errInvalidKey As Long = ERR_BASE_NUM + 3
    On Error Resume Next
    Dim dic As Object
    Set dic = fdf.GenRow("pachinko", 123, 456, 789)
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errInvalidKey, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="奇数個目の引数に非文字列 -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' 奇数番目の引数にEmpty -> カスタム例外
Sub Test_Error_GenRow_OddNumberOfArgsEmpty()
    Const errEmptyKey As Long = ERR_BASE_NUM + 6
    On Error Resume Next
    Dim dic As Object
    Set dic = fdf.GenRow("pachinko", 123, Empty, 789)
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errEmptyKey, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="奇数個目の引数にEmpty -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' 奇数番目の引数にNull -> カスタム例外
Sub Test_Error_GenRow_OddNumberOfArgsNull()
    Const errNullKey As Long = ERR_BASE_NUM + 7
    On Error Resume Next
    Dim dic As Object
    Set dic = fdf.GenRow("pachinko", 123, Null, 789)
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errNullKey, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="奇数個目の引数にNull -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' 奇数番目の引数に空白文字 -> カスタム例外
Sub Test_Error_GenRow_OddNumberOfArgsBlank()
    Const errBlankKey As Long = ERR_BASE_NUM + 8
    On Error Resume Next
    Dim dic As Object
    Set dic = fdf.GenRow("pachinko", 123, "   ", 789)
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errBlankKey, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="奇数個目の引数にブランク -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' キーの正規化 -> 重複チェックが効くか
Sub Test_Error_GenRow_NormalizedDuplicateKey()
    Const errDuplicateKey As Long = ERR_BASE_NUM + 5
    On Error Resume Next
    Dim dic As Object
    Set dic = fdf.GenRow("pachinko", 123, " Pachinko", 456)
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errDuplicateKey, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="正規化したキーが重複 -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' 偶数番目の引数に配列
Sub Test_Error_GenRow_EvenNumberOfArgArray()
    Const errInvalidValue As Long = ERR_BASE_NUM + 4
    On Error Resume Next
    Dim dic As Object
    Set dic = fdf.GenRow("pachinko", 123, " Pachinko", Array("pachinko", "123"))
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errInvalidValue, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="偶数番目の引数が配列 -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub
' 偶数番目の引数にオブジェクト
Sub Test_Error_GenRow_EvenNumberOfArgObject()
    Const errInvalidValue As Long = ERR_BASE_NUM + 4
    On Error Resume Next
    Dim dic As Object, o As Object
    Set o = CreateObject("Scripting.FileSystemObject")
    Set dic = fdf.GenRow("pachinko", 123, " Pachinko", o)
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call ut.AssertError( _
        a_ExpectedErrNumber:=errInvalidValue, _
        a_ActualErrNumber:=Err.Number, _
        a_Comment:="偶数番目の引数がオブジェクト -> カスタム例外")
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call Err.Clear
End Sub

Private Sub CallStartTest(): Call StartTest: End Sub
Private Sub CallStartTestByName(): Call StartTestByName: End Sub

Private Sub ZZ_Main(): End Sub
' =============================================================================
'   Main Procedures errHasInvalidChar
' =============================================================================
Private Sub RunAllTests( _
            ByRef a_TestRunConfig As TestRunConfig)
    Call RunTests(a_TestRunConfig)
End Sub

Private Sub RunTestsByName( _
            ByRef a_TestRunConfig As TestRunConfig, _
            ByVal a_TestName As String)
    Call RunTests(a_TestRunConfig, a_TestName)
End Sub

Private Sub RunTests( _
            ByRef a_TestRunConfig As TestRunConfig, _
   Optional ByVal a_TestName As String = "")
    ' セットアップ
    Call Setup
    ' テストプロシージャ名をリストアップ
    Call ListUpTestProcNames(a_TestName)
    
    ' ここでテストプロシージャがなかったらイミディエイトに表示してExit
    Dim testProcCnt As Long
    testProcCnt = m_TestProcNames.Count
    If testProcCnt = 0 Then
        Debug.Print TEST_NOT_FOUND
        Exit Sub
    End If
    
    ' イミディエイトの表示
    Dim tmp As String
    tmp = IIf((testProcCnt < 2), " Test...", " Tests...")
    Debug.Print BL
    Debug.Print "Running " & CStr(testProcCnt) & tmp
    Debug.Print BL
    
    Dim k As Variant
    For Each k In m_TestProcNames.Keys
        Debug.Print "-> " & k
        ' テスト実行
        Call BeforeEach
        Call Application.Run(MODULE_NAME & "." & k)
        ' 失敗即停止モードのとき
        If a_TestRunConfig.StopIfFailed Then
            If ut.HasFailed Then Exit For
        End If
    Next
    
    Call PrintSummary(a_TestRunConfig)
    
    If a_TestRunConfig.PrintLog Then Call PrintTestLogs
    
    If Not ut.HasFailed Then
        Debug.Print sf.BreakLine(a_Length:=50, a_Char:="*")
        Debug.Print "All tests completed successfully!!"
        Debug.Print sf.BreakLine(a_Length:=50, a_Char:="*")
    End If
    
    ' あとしまつ
    Call Teardown
End Sub

Private Sub PrintSummary( _
            ByRef a_TestRunConfig As TestRunConfig)
    ' 結果表示
    Dim totalCnt As Long
    totalCnt = ut.TestLogs.Count
    
    Dim tmp As String
    tmp = IIf((totalCnt = 1), " Procedure.", " Procedures.")
    ' テスト結果の集計
    Dim tl As UnitTestLog
    For Each tl In ut.TestLogs
        If tl.Succeeded Then
            m_SucceededCount = m_SucceededCount + 1
        Else
            m_FailedCount = m_FailedCount + 1
        End If
    Next
    
    ' 結果（サマリ）出力
    Dim summary As String
    summary = _
        Join( _
            Array( _
                "Tested " & CStr(totalCnt) & tmp, _
                "[PASSED] " & CStr(m_SucceededCount), _
                "[FAILED] " & CStr(m_FailedCount) _
             ), _
            sf.NL _
        )
    ' テスト件数が1件だけだったら中断したかどうかは関係ない
    Dim testCnt As Long
    testCnt = m_TestProcNames.Count
    If testCnt = 1 Then GoTo Finally
    ' テスト全実行モードの場合はそもそも中断しない
    If Not a_TestRunConfig.StopIfFailed Then GoTo Finally
    ' テスト失敗がなかった場合はテスト中断表示は不要
    If Not ut.HasFailed Then GoTo Finally
    
    ' 中断した場合のみ、テストが中断した旨を表示
    Dim stopped As String
    stopped = "[STOPPED] Test execution aborted due to failure." & sf.NL
    summary = stopped & sf.NL & summary
    Dim stoppedComment As String
    stoppedComment = _
        "- Executed: " & CStr(totalCnt) & " / " & CStr(testCnt) & " tests"
    If testCnt > totalCnt Then
        stoppedComment = _
            stoppedComment & sf.NL & "- Remaining tests were skipped."
    End If
    summary = summary & sf.NL & stoppedComment
Finally:
    Debug.Print BL
    Debug.Print summary
    Debug.Print BL
End Sub

Private Sub PrintTestLogs()
    Dim tl As UnitTestLog
    For Each tl In ut.TestLogs
        Debug.Print tl.BuildMessage()
        Debug.Print BL
    Next
End Sub

Private Sub ZZ_Configs(): End Sub
' =============================================================================
'   Configuration Procedures
' =============================================================================
Private Sub Setup()
    ' カウンタをリセット
    Call ResetCounts
    ' UnitTestインスタンスを取得
    Set ut = New UnitTest
    
    ' 擬似イミディエイトウィンドウクリア
    Debug.Print sf.NL(2)
    ' 開始メッセージ書き込み
    Debug.Print sf.BreakLine(a_Length:=50, a_Char:="*")
    Debug.Print "Test start: " & Format(Now(), "yyyy-mm-dd HH:mm:ss")
    Debug.Print sf.BreakLine(a_Length:=50, a_Char:="*")
    
    ' 区切り線の設定（デフォルトは長さ40、文字が`-`）
    BL = sf.BreakLine(a_Length:=50, a_Char:="-")
    ' テストプロシージャがないときのメッセージ
    TEST_NOT_FOUND = _
        BL & sf.NL & _
        "No test procedure found..." & sf.NL & _
        BL
    ' セットアップ用コードをここに実装
    Set m_TestProcNames = CreateObject("Scripting.Dictionary")
End Sub

Private Sub BeforeEach()
    ' テストごとに実行したい処理はここに追加
    '   -> インスタンスの初期化など
    ' なければ空でよい
    Set fdf = New FakeDataFactory
End Sub

Private Sub Teardown()
    ' 設定破棄用コードをここに実装
    ' UnitTestオブジェクトのログを破棄
    Call ut.ClearTestLogs
End Sub

Private Sub ZZ_HelperFunctions(): End Sub
' =============================================================================
'   Helper Functions
' =============================================================================
' 内部成功/失敗カウンタをリセット
Private Sub ResetCounts()
    m_SucceededCount = 0
    m_FailedCount = 0
End Sub

' `Test_`で始まるプロシージャ名を`m_TestProcNames`に格納する
Private Sub ListUpTestProcNames( _
   Optional ByVal a_TestName As String = "")
    ' CodeModuleインスタンス取得
    Dim cm As Object
    Set cm = GetSelfCodeModule()
    
    Dim ln As Long, proc As String
    Dim startLine As Long, procLines As Long
    ln = 1
    ' CodeModule.CountOfLine
    '   - CodeModule内のコード行数を返す
    Do While ln <= cm.CountOfLines
        ' CodeModule.ProcOfLine(<Long: Line>, <vbext_ProcKind: ProcKind>)
        '   - 指定した行が属するプロシージャ名を返す
        proc = cm.ProcOfLine(ln, vbext_pk_Proc)
        ' プロシージャ名が返らない -> 行カウンタをインクリメントしてスキップ
        If proc = "" Then ln = ln + 1: GoTo Continue
        
        ' プロシージャ名が返った -> プロシージャの開始行番号をキャッシュ
        ' CodeModule.ProcStartLine(<String: ProcName>, <vbext_ProcKindProcKind>)
        '   - 指定したプロシージャの開始行番号を返す
        startLine = cm.ProcStartLine(proc, vbext_pk_Proc)
        ' 現在の行位置と開始行番号が不一致
        '   同一プロシージャの2行目以降 -> スキップ
        If startLine <> ln Then GoTo Skip
        ' `Test_`で始まらない -> スキップ
        If Left$(proc, 5) <> "Test_" Then GoTo Skip
        ' 引数`a_TestName`が空白でなく、かつ`proc`にマッチしない -> スキップ
        If a_TestName <> "" And Not (proc Like a_TestName) Then GoTo Skip
        ' ディクショナリに追加済み -> スキップ
        If m_TestProcNames.Exists(proc) Then GoTo Skip
               
        ' ここまでたどり着いたらディクショナリに追加
        Call m_TestProcNames.Add(proc, True)
Skip:
        ' CodeModule.ProcCountLines(<String: ProcName>, <vbext_ProcKind: ProcKind>)
        ' 指定したプロシージャの行数を返す
        ' プロシージャ開始行番号 + プロシージャの行数 = プロシージャの次の行
        procLines = cm.ProcCountLines(proc, vbext_pk_Proc)
        ln = startLine + procLines
        GoTo Continue
Continue:
    Loop
End Sub

' 自分自身のCodeModuleインスタンスを取得
Private Function GetSelfCodeModule() As Object
    Dim ret As Object
    Set ret = ThisWorkbook.VBProject.VBComponents(MODULE_NAME).CodeModule
    Set GetSelfCodeModule = ret
End Function

