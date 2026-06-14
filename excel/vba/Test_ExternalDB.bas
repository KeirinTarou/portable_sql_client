Attribute VB_Name = "Test_ExternalDB"
Option Explicit

Private Const MODULE_NAME As String = "Test_ExternalDB"

Private Const vbext_pk_Proc As Long = 0

Private ut As UnitTest
Private sf As New StringFormatter
Private strf As New Stringifier
Private edb As New ExternalDB
Private fdf As New FakeDataFactory
Private ff As New FakeField

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

' 期待値・実測値・コメント用変数
Private e As Variant, a As Variant, c As String
' 接続文字列
Private CONN_STR As String

' テスト用クエリ
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' 家庭用
Private Const QUERY_SCALAR_TEXT_NO_PARAM As String = _
    "SELECT 'IRON MAIDEN' AS ""name"";"
Private Const QUERY_SCALAR_TEXT_WITH_PARAM As String = _
    "SELECT 'IRON MAIDEN' AS ""name"" WHERE ? = 1;"
Private Const QUERY_SCALAR_TEXT_WITH_PARAMS As String = _
    "SELECT 'IRON MAIDEN' AS ""name"" WHERE ? = 'pachinko' AND ? = 123;"
Private Const QUERY_SCALAR_NUMBER_NO_PARAM As String = _
    "SELECT 666 AS ""number_of_the_beast"";"
Private Const QUERY_SCALAR_NUMBER_WITH_PARAM As String = _
    "SELECT 666 AS ""number_of_the_beast"" WHERE ? = 1"
Private Const QUERY_SCALAR_DATE_NO_PARAM As String = _
    "SELECT STR_TO_DATE('1973-03-23', '%Y-%m-%d') AS ""date"";"
Private Const QUERY_SCALAR_DATE_WITH_PARAM As String = _
    "SELECT STR_TO_DATE('1973/03/23', '%Y/%m/%d') AS ""date"" WHERE ? = 1;"
Private Const QUERY_SCALAR_DATETIME As String = _
    "SELECT STR_TO_DATE('2026/04/19 23:28:43', '%Y/%m/%d %H:%i:%s') AS ""datetime"";"
Private Const QUERY_SCALAR_BLANK As String = _
    "SELECT '' AS ""blank"";"
Private Const QUERY_SCALAR_NULL As String = _
    "SELECT NULL AS ""Null"";"
Private Const QUERY_NO_RECORD As String = _
    "SELECT 1 WHERE 1 = 0;"
Private Const QUERY_MULTI_RECORDS As String = _
    "SELECT 'Steve Harris' AS ""name"", 70 AS ""age"" UNION ALL " & _
    "SELECT 'Bruce Dickinson' AS ""name"", 65 AS ""age"";"
Private Const QUERY_MULTI_COLUMNS As String = _
    "SELECT 'Steve Harris' AS ""name"", 70 AS ""age"", " & _
    "STR_TO_DATE('1956/03/12', '%Y/%m/%d') AS ""birthday"";"
Private Const ERR_QUERY_WITH_SINGLE_PARAM As String = _
    "SLECT ? AS ""name"";"
Private Const ERR_QUERY_WITH_MULTI_PARAMS As String = _
    "SLECT 'Steve' AS ""name"" WHERE ? = 'pachinko' AND ? = '123';"
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' 職場用
'Private Const QUERY_SCALAR_TEXT_NO_PARAM As String = _
'    "SELECT 'IRON MAIDEN' AS ""name"" FROM dual;"
'Private Const QUERY_SCALAR_TEXT_WITH_PARAM As String = _
'    "SELECT 'IRON MAIDEN' AS ""name"" FROM dual WHERE ? = 1;"
'Private Const QUERY_SCALAR_TEXT_WITH_PARAMS As String = _
'    "SELECT 'IRON MAIDEN' AS ""name"" FROM dual WHERE ? = 'pachinko' AND ? = 123;"
'Private Const QUERY_SCALAR_NUMBER_NO_PARAM As String = _
'    "SELECT 666 AS ""number_of_the_beast"" FROM dual;"
'Private Const QUERY_SCALAR_NUMBER_WITH_PARAM As String = _
'    "SELECT 666 AS ""number_of_the_beast"" FROM dual WHERE ? = 1"
'Private Const QUERY_SCALAR_DATE_NO_PARAM As String = _
'    "SELECT TO_DATE('1973/03/23', 'YYYY/MM/DD') AS ""date"" FROM dual;"
'Private Const QUERY_SCALAR_DATE_WITH_PARAM As String = _
'    "SELECT TO_DATE('1973/03/23', 'YYYY/MM/DD') AS ""date"" FROM dual WHERE ? = 1;"
'Private Const QUERY_SCALAR_DATETIME As String = _
'    "SELECT TO_DATE('2026/04/19 23:28:43', 'YYYY/MM/DD HH24/MI/SS') AS ""datetime"" FROM dual;"
'Private Const QUERY_SCALAR_BLANK As String = _
'    "SELECT '' AS ""blank"" FROM dual;"
'Private Const QUERY_SCALAR_NULL As String = _
'    "SELECT NULL AS ""Null"" FROM dual;"
'Private Const QUERY_NO_RECORD As String = _
'    "SELECT 1 FROM dual WHERE 1 = 0;"
'Private Const QUERY_MULTI_RECORDS As String = _
'    "SELECT 'Steve Harris' AS ""name"", 70 AS ""age"" FROM dual UNION ALL " & _
'    "SELECT 'Bruce Dickinson' AS ""name"", 65 AS ""age"" FROM dual;"
'Private Const QUERY_MULTI_COLUMNS As String = _
'    "SELECT 'Steve Harris' AS ""name"", 70 AS ""age"", TO_DATE('1956/03/12', 'YYYY/MM/DD') AS ""birthday"" " & _
'    "FROM dual;"
'Private Const ERR_QUERY_WITH_SINGLE_PARAM As String = _
'    "SLECT ? AS ""name"" FROM dual;"
'Private Const ERR_QUERY_WITH_MULTI_PARAMS As String = _
'    "SLECT 'Steve' AS ""name"" FROM dual WHERE ? = 'pachinko' AND ? = '123';"
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Private Const QUERY_NOT_EXIST_TABLE As String = _
    "SELECT 1 FROM not_exist;"

Private Sub AA_EntryPoint(): End Sub
' =============================================================================
'   Entry Point
' =============================================================================

' 全てのテストを実行
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
    'cfg.PrintLog = False
    
    'Call RunTestsByName(cfg, "Test_*")
    'Call RunTestsByName(cfg, "Test_FetchOne*")
    'Call RunTestsByName(cfg, "Test_FetchAll*")
    'Call RunTestsByName(cfg, "Test_FetchAll_NoRecord")
    Call RunTestsByName(cfg, "Test_Error_*")
End Sub

Private Sub AA_Configs(): End Sub
' =============================================================================
'   Configuration Procedures
' =============================================================================
Private Sub Setup()
    ' カウンタをリセット
    Call ResetCounts
    ' UnitTestインスタンスを取得
    Set ut = New UnitTest
    CONN_STR = edb.CONNECTION_STRING
    
    ' 期待値データ作成用FakeDataFactoryを初期化
    Set fdf = New FakeDataFactory
    
    ' 擬似イミディエイトウィンドウクリア
    Debug.Print sf.NL(10)
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
    ' テストプロシージャの名前を格納するディクショナリ
    Set m_TestProcNames = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Teardown()
    ' 設定破棄用コードをここに実装
    ' UnitTestオブジェクトのログを破棄
    Call ut.ClearTestLogs
End Sub

Private Sub ZZ_TestCodeHere(): End Sub
' =============================================================================
'   Test Procedures
' =============================================================================
Private Sub Test_FetchScalar_Text()
    e = "IRON MAIDEN"
    a = edb.FetchScalar(CONN_STR, QUERY_SCALAR_TEXT_NO_PARAM)
    c = "テキストのスカラ値を返す"
    Call ut.AssertEqual(e, a, c)
End Sub
' 文字列のスカラ値（1つのパラメータ）
Private Sub Test_FetchScalar_Text_WithParam()
    e = "IRON MAIDEN"
    a = _
        edb.FetchScalar(CONN_STR, QUERY_SCALAR_TEXT_WITH_PARAM, Array(1))
    c = "テキストのスカラ値を返す（パラメータ1つ使用）"
    Call ut.AssertEqual(e, a, c)
End Sub
' 文字列のスカラ値（複数のパラメータ）
Private Sub Test_FetchScalar_Text_WithParams()
    e = "IRON MAIDEN"
    a = _
        edb.FetchScalar(CONN_STR, QUERY_SCALAR_TEXT_WITH_PARAMS, Array("pachinko", 123))
    c = "テキストのスカラ値を返す（パラメータ複数使用）"
    Call ut.AssertEqual(e, a, c)
End Sub
' 数値のスカラ値
Private Sub Test_FetchScalar_Number()
    e = 666
    a = edb.FetchScalar(CONN_STR, QUERY_SCALAR_NUMBER_NO_PARAM)
    c = "数値のスカラ値を返す"
    Call ut.AssertEqual(e, a, c)
End Sub
' 日付のスカラ値
Private Sub Test_FetchScalar_Date()
    e = CDate("1973/03/23")
    a = edb.FetchScalar(CONN_STR, QUERY_SCALAR_DATE_NO_PARAM)
    c = "日付のスカラ値を返す"
    Call ut.AssertEqual(e, a, c)
End Sub
' 日付時刻のスカラ値
Private Sub Test_FetchScalar_Datetime()
    e = CDate("2026/04/19 23:28:43")
    a = edb.FetchScalar(CONN_STR, QUERY_SCALAR_DATETIME)
    c = "日付時刻のスカラ値を返す"
    Call ut.AssertEqual(e, a, c)
End Sub
' 空文字のスカラ値
Private Sub Test_FetchScalar_Blank()
    ' 家庭用と職場用で返り値を切り替える
    e = ""
    If (InStr(1, QUERY_SCALAR_BLANK, "dual") > 0) Then e = Null
    a = edb.FetchScalar(CONN_STR, QUERY_SCALAR_BLANK)
    c = "空文字のスカラ値を返す"
    Call ut.AssertEqual(e, a, c)
End Sub
' Null
Private Sub Test_FetchScalar_Null()
    e = Null
    a = edb.FetchScalar(CONN_STR, QUERY_SCALAR_NULL)
    c = "Nullを返す"
    Call ut.AssertEqual(e, a, c)
End Sub
' レコードセットが返らなかった
Private Sub Test_FetchScalar_NoRecord()
    e = Null
    a = edb.FetchScalar(CONN_STR, QUERY_NO_RECORD)
    c = "レコードセットなし -> Null"
    Call ut.AssertEqual(e, a, c)
End Sub
' 複数列・行を返すクエリの場合、TopLeftを返す
Private Sub Test_FetchScalar_MultiValues()
    e = "Steve Harris"
    a = edb.FetchScalar(CONN_STR, QUERY_MULTI_RECORDS)
    c = "複数列・行を返すクエリ -> TopLeftを採用 & 警告"
    Call ut.AssertEqual(e, a, c)
End Sub
' FetchOne()
' 文字列・数値・日付カラムを持つレコード
Private Sub Test_FetchOne_TextNumberDate()
    Set e = _
        fdf.GenRow( _
            "name", "Steve Harris", _
            "age", 70, _
            "birthday", CDate("1956/03/12") _
        )
    Set a = edb.FetchOne(CONN_STR, QUERY_MULTI_COLUMNS)
    c = "文字列・数値・日付カラムを持つレコード"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_FetchOne_NoRecord()
    Set e = CreateObject("Scripting.Dictionary")
    Set a = edb.FetchOne(CONN_STR, QUERY_NO_RECORD)
    c = "レコードセットなし -> 空ディクショナリ"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_FetchOne_MultiRows()
    Set e = fdf.GenRow("name", "Steve Harris", "age", 70)
    Set a = edb.FetchOne(CONN_STR, QUERY_MULTI_RECORDS)
    c = "複数行を返すクエリ -> 1行目を採用 & 警告"
    Call ut.AssertEqual(e, a, c)
End Sub
' FetchAll()
Private Sub Test_FetchAll_MultiRows()
    e = fdf.AddRow("name", "Steve Harris", "age", 70)
    e = fdf.AddRow("name", "Bruce Dickinson", "age", 65)
    a = edb.FetchAll(CONN_STR, QUERY_MULTI_RECORDS)
    c = "複数行のレコードを返すクエリ"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_FetchAll_NoRecord()
    e = Array()
    a = edb.FetchAll(CONN_STR, QUERY_NO_RECORD)
    c = "レコードセットなし -> 空の配列"
    Call ut.AssertEqual(e, a, c)
End Sub
' GetFieldValueOrNull()
Private Sub Test_GetFieldValueOrNull_NoError_String()
    ' テスト用ニセFieldオブジェクトに値を設定
    Call ff.SetRaiseValueError(False)
    ff.Name = "Test": ff.Value = "Test"
    e = "Test"
    a = edb.GetFieldValueOrNull(ff, True)
    c = "Fieldオブジェクトから文字列の値が取り出せる"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_GetFieldValueOrNull_NoError_Int()
    ' テスト用ニセFieldオブジェクトに値を設定
    Call ff.SetRaiseValueError(False)
    ff.Name = "Test": ff.Value = 123
    e = 123
    a = edb.GetFieldValueOrNull(ff, True)
    c = "Fieldオブジェクトから整数型の値が取り出せる"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_GetFieldValueOrNull_NoError_Date()
    ' テスト用ニセFieldオブジェクトに値を設定
    Call ff.SetRaiseValueError(False)
    ff.Name = "Test": ff.Value = #6/5/2026#
    e = #6/5/2026#
    a = edb.GetFieldValueOrNull(ff, True)
    c = "Fieldオブジェクトから日付型の値が取り出せる"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_GetFieldValueOrNull_NoError_Null()
    ' テスト用ニセFieldオブジェクトに値を設定
    Call ff.SetRaiseValueError(False)
    ff.Name = "Test": ff.Value = Null
    e = Null
    a = edb.GetFieldValueOrNull(ff, True)
    c = "FieldオブジェクトからNull値が取り出せる"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_GetFieldValueOrNull_WithError()
    ' テスト用ニセFieldオブジェクトに値を設定
    Call ff.SetRaiseValueError(True)
    ff.Name = "Test": ff.Value = "pachinko123"
    e = Null
    a = edb.GetFieldValueOrNull(ff, True)
    c = "Value取り出し時にエラー > Nullにフォールバック＆デバッグ出力"
    Call ut.AssertEqual(e, a, c)
End Sub

' 異常系
Private Sub Test_Error_FetchScalar_SingleParams()
    On Error Resume Next
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call edb.FetchScalar( _
        CONN_STR, ERR_QUERY_WITH_SINGLE_PARAM, Array("Steve"))
    e = -2147217900
    a = Err.Number
    c = "パラメータのクエリで例外発生 -> パラメータを表示"
    Call ut.AssertError(e, a, c)
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    e = "(Steve)"
    a = Right(Err.Description, Len(e))
    c = "エラーメッセージに`Steve`が含まれる"
    Call ut.AssertEqual(e, a, c)
    Call Err.Clear
    On Error GoTo 0
End Sub
Private Sub Test_Error_FetchScalar_MultiParams()
    On Error Resume Next
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call edb.FetchScalar( _
        CONN_STR, ERR_QUERY_WITH_MULTI_PARAMS, Array("pachinko", "123"))
    e = -2147217900
    a = Err.Number
    c = "複数パラメータのクエリで例外発生 -> パラメータをカンマつなぎで表示"
    Call ut.AssertError(e, a, c)
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    e = "(pachinko,123)"
    a = Right(Err.Description, Len(e))
    c = "エラーメッセージに`pachinko,123`が含まれる"
    Call ut.AssertEqual(e, a, c)
    Call Err.Clear
    On Error GoTo 0
End Sub

Private Sub CallStartTest(): Call StartTest: End Sub
Private Sub CallStartTestByName(): Call StartTestByName: End Sub
' =============================================================================
' Test Procedures Sample: 不要ならば消してしまってOK
'   - プロシージャ名の先頭に`Test_`を付ける
'   - 引数なしのSubプロシージャ限定
' =============================================================================
Private Sub AA_SampleTestProcedures(): End Sub
' Sample >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'Sub Test_Add2Numbers_NormalArgs()
'    Call ut.AssertEqual( _
'        a_Expected:=3, _
'        a_Actual:=Add2Numbers(1, 2), _
'        a_Comment:="1 + 2 = 3")
'End Sub
'
'Sub Test_Add2Numbers_Fail()
'    Call ut.AssertEqual( _
'        a_Expected:=1, _
'        a_Actual:=Add2Numbers(1, 1), _
'        a_Comment:="Should be failed.")
'End Sub
'
'Sub Test_Add2Numbers_NotEqual()
'    Call ut.AssertNotEqual(3, Add2Numbers(1, 1), "Not equal, so succeed.")
'End Sub
'
'Public Function Add2Numbers( _
'            ByVal x As Long, _
'            ByVal y As Long) As Long
'    Add2Numbers = x + y
'End Function

Private Sub ZZ_SampleErrorTestProcedures(): End Sub
' -----------------------------------------------------------------------------
'   Error test procedure template:
'       On Error Resume Next
'       (Call your proc)
'       Call ut.AssertError(ExpectedErrNum, Err.Number, "Comment")
'       Call Err.Clear
' -----------------------------------------------------------------------------
'Sub Test_DivideNumber_ErrorDivByZero()
'    On Error Resume Next
'' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'    Call DivideNumber(3, 0)
'    Call ut.AssertError( _
'        a_ExpectedErrNumber:=11, _
'        a_ActualErrNumber:=Err.Number, _
'        a_Comment:="Expected zero div error.")
'' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'    Call Err.Clear
'End Sub
'
'Sub Test_DivideNumber_FailErrorDivByZero()
'    On Error Resume Next
'' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'    Call DivideNumber(9, 3)
'    Call ut.AssertError( _
'        a_ExpectedErrNumber:=11, _
'        a_ActualErrNumber:=Err.Number, _
'        a_Comment:="No error, test should be failed.")
'' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'    Call Err.Clear
'End Sub
'
'Public Function DivideNumber( _
'            ByVal x As Long, _
'            ByVal y As Long) As Long
'    DivideNumber = Int(x / y)
'End Function
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Sample

Private Sub AA_Main(): End Sub
' =============================================================================
'   Main Procedures
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

Private Sub AA_HelperFunctions(): End Sub
' =============================================================================
'   Helper Functions
' =============================================================================
' PrettyPrint文字列を作る
Private Function PP(ParamArray a_Lines()) As String
    PP = Join(a_Lines, vbCrLf)
End Function

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


