Attribute VB_Name = "Test_RecordsetJsonConverter"
Option Explicit

Private Const MODULE_NAME As String = "Test_RecordsetJsonConverter"

Private Const vbext_pk_Proc As Long = 0

Private ut As UnitTest
Private sf As New StringFormatter

Private m_SucceededCount As Long
Private m_FailedCount As Long

Private m_TestProcNames As Object

Private rjc As New RecordsetJsonConverter
' テスト用のADODB.Recordsetインスタンス格納用
'   - インスタンスはSetup()で作成
Private rs As Object

' ADODB.Recordset用の列挙体
Private Enum ObjectStateEnum
    adStateClosed = 0
    adStateOpen = 1
    adStateConnecting = 2
    adStateExecuting = 4
    adStateFetching = 8
End Enum

Private Enum DataTypeEnum
    adInteger = 3
    adDouble = 5
    adDBTimeStamp = 135
    adVarChar = 200
    adVarWChar = 202
End Enum

Private Const adFldIsNullable As Long = &H20

' 区切り線 <- Setup()で内容をセット
Private BL As String
' テストプロシージャがないときの中断メッセージ
Private TEST_NOT_FOUND As String

Private Type TestRunConfig
    StopIfFailed As Boolean
    PrintLog As Boolean
End Type

Private e As Variant
Private a As Variant
Private c As String

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
'    cfg.PrintLog = False
    
    Call RunTestsByName(cfg, "Test_*")
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
    ' テスト用Recordsetインスタンス
    Set rs = CreateObject("ADODB.Recordset")
    
    ' テストプロシージャの名前を格納するディクショナリ
    Set m_TestProcNames = CreateObject("Scripting.Dictionary")
End Sub

Private Sub BeforeEach()
    ' テストごとに実行したい処理はここに追加
    '   -> インスタンスの初期化など
    ' なければ空でよい
    ' Recordsetインスタンスが開きっぱなしだったら閉じる
    If Not rs Is Nothing Then
        If rs.State <> adStateClosed Then Call rs.Close
        Set rs = Nothing
    End If
    ' テスト用Recordsetインスタンス
    Set rs = CreateObject("ADODB.Recordset")
End Sub

Private Sub Teardown()
    ' 設定破棄用コードをここに実装
    ' RecordsetインスタンスのStateがadStateClose以外だったらClose
    If Not rs Is Nothing Then
        If rs.State <> adStateClosed Then Call rs.Close
        Set rs = Nothing
    End If
    ' UnitTestオブジェクトのログを破棄
    Call ut.ClearTestLogs
End Sub

Private Sub AA_TestCodeHere(): End Sub
' =============================================================================
'   Test Procedures
' =============================================================================
' ToJSONLiteral()メソッド
Private Sub Test_ToJSONLiteral_Null()
    e = "null"
    a = rjc.ToJSONLiteral(Null)
    c = "ToJSONLiteral(): NULL -> null"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_ToJSONLiteral_Empty()
    e = "null"
    a = rjc.ToJSONLiteral(Empty)
    c = "ToJSONLiteral(): Empty -> null"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_ToJSONLiteral_Blank()
    e = """"""
    a = rjc.ToJSONLiteral("")
    c = "ToJSONLiteral(): """" -> """""
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_ToJSONLiteral_String()
    e = """pachinko123"""
    a = rjc.ToJSONLiteral("pachinko123")
    c = "ToJSONLiteral(): pachinko123 -> pachinko123"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_ToJSONLiteral_Japanese()
    e = """ち～ん（笑）"""
    a = rjc.ToJSONLiteral("ち～ん（笑）")
    c = "ToJSONLiteral(): ち～ん（笑） -> ち～ん（笑）"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_ToJSONLiteral_Int()
    e = "123"
    a = rjc.ToJSONLiteral(123)
    c = "ToJSONLiteral(): 123 -> ""123"""
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_ToJSONLiteral_Float()
    e = "1.23"
    a = rjc.ToJSONLiteral(1.23)
    c = "ToJSONLiteral(): 1.23 -> ""1.23"""
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_ToJSONLiteral_Date()
    e = """2026/06/13"""
    a = rjc.ToJSONLiteral(#6/13/2026#)
    c = "ToJSONLiteral(): 2026/06/13 -> ""2026/06/13"""
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_ToJSONLiteral_DateTime()
    e = """2026/06/13 12:34:56"""
    a = rjc.ToJSONLiteral(#6/13/2026 12:34:56 PM#)
    c = "ToJSONLiteral(): 2026/06/13 12:34:56 -> ""2026/06/13 12:34:56"""
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_ToJSONLiteral_EscapeDoubleQuote()
    e = """Steve \""Harris\"""""
    a = rjc.ToJSONLiteral("Steve ""Harris""")
    c = "ToJSONLiteral(): Steve ""Harris"" -> Steve \""Harris\"""
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_ToJSONLiteral_EscapeBackSlash()
    e = """Steve \\ Harris"""
    a = rjc.ToJSONLiteral("Steve \ Harris")
    c = "ToJSONLiteral(): Steve \ Harris -> Steve \\ Harris"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_ToJSONLiteral_EscapeTab()
    e = """Steve\tHarris"""
    a = rjc.ToJSONLiteral("Steve" & vbTab & "Harris")
    c = "ToJSONLiteral(): Steve<Tab>Harris -> Steve\tHarris"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_ToJSONLiteral_EscapeCrLf()
    e = """Steve\r\nHarris"""
    a = rjc.ToJSONLiteral("Steve" & vbCrLf & "Harris")
    c = "ToJSONLiteral(): Steve<CrLf>Harris -> Steve\r\nHarris"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_ToJSONLiteral_EscapeCr()
    e = """Steve\rHarris"""
    a = rjc.ToJSONLiteral("Steve" & vbCr & "Harris")
    c = "ToJSONLiteral(): Steve<Cr>Harris -> Steve\r\nHarris"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_ToJSONLiteral_EscapeLf()
    e = """Steve\nHarris"""
    a = rjc.ToJSONLiteral("Steve" & vbLf & "Harris")
    c = "ToJSONLiteral(): Steve<Lf>Harris -> Steve\r\nHarris"
    Call ut.AssertEqual(e, a, c)
End Sub
' BuildJsonCore()メソッド
Private Sub Test_BuildJSONCore_OneRow()
    Dim cols As Variant
    cols = Array("ID", "NAME")
    Dim Rows As Variant
    Rows = Array(Array(1, "Steve"))
    e = "{""columns"": [""ID"", ""NAME""], ""rows"": [[1, ""Steve""]]}"
    a = rjc.BuildJSONCore(cols, Rows)
    c = "1行分のレコードがJSON文字列化される。"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_BuildJSONCore_OneRowNull()
    Dim cols As Variant
    cols = Array("ID", "NAME")
    Dim Rows As Variant
    Rows = Array(Array(1, Null))
    e = "{""columns"": [""ID"", ""NAME""], ""rows"": [[1, null]]}"
    a = rjc.BuildJSONCore(cols, Rows)
    c = "カラムの値がNULL -> `rows`が`[[1, null]]`になる"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_BuildJSONCore_NoRows()
    Dim cols As Variant
    cols = Array("ID", "NAME")
    Dim Rows As Variant
    Rows = Array()
    e = "{""columns"": [""ID"", ""NAME""], ""rows"": []}"
    a = rjc.BuildJSONCore(cols, Rows)
    c = "結果セットが0件のとき -> `rows`が`[]`になる"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_BuildJSONCore_MultiRows()
    Dim cols As Variant
    cols = Array("ID", "NAME")
    Dim Rows As Variant
    Rows = Array(Array(1, "Steve"), Array(2, "Bruce"))
    e = "{""columns"": [""ID"", ""NAME""], ""rows"": [[1, ""Steve""], [2, ""Bruce""]]}"
    a = rjc.BuildJSONCore(cols, Rows)
    c = "2行分のレコードが正しくJSON文字列化される"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_BuildJSONCore_MultiRowsWithNull()
    Dim cols As Variant
    cols = Array("ID", "NAME")
    Dim Rows As Variant
    Rows = Array(Array(1, "Steve"), Array(2, Null), Array(3, "Bruce"))
    e = "{""columns"": [""ID"", ""NAME""], ""rows"": [[1, ""Steve""], [2, null], [3, ""Bruce""]]}"
    a = rjc.BuildJSONCore(cols, Rows)
    c = "複数のレコードに値がNULLのレコードがあっても正しくJSON文字列化される"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_BuildJSONCore_ThreeColumns()
    Dim cols As Variant
    cols = Array("ID", "NAME", "AGE")
    Dim Rows As Variant
    Rows = Array(Array(1, "Steve", 70), Array(2, "Bruce", 69))
    e = "{""columns"": [""ID"", ""NAME"", ""AGE""], ""rows"": [[1, ""Steve"", 70], [2, ""Bruce"", 69]]}"
    a = rjc.BuildJSONCore(cols, Rows)
    c = "3列のレコードが正しくJSON文字列化される"
    Call ut.AssertEqual(e, a, c)
End Sub
' ExtractColumns()メソッド
Private Sub Test_ExtractColumns_OneColumn()
    ' カラムの設定
    With rs
        With .Fields
            Call .Append("ID", adInteger)
        End With
        Call .Open
    End With
    ' RecordsetJSONConverterインスタンス初期化
    Call rjc.Init(rs)
    Dim arr As Variant
    arr = rjc.ExtractColumns()
    e = "ID"
    a = arr(0)
    c = "フィールド名 -> ID"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_ExtractColumns_TwoColumns()
    ' カラムの設定
    With rs
        With .Fields
            Call .Append("ID", adInteger)
            Call .Append("NAME", adVarChar, 100)
        End With
        Call .Open
    End With
    ' RecordsetJSONConverterインスタンス初期化
    Call rjc.Init(rs)
    Dim arr As Variant
    arr = rjc.ExtractColumns()
    e = "ID|NAME"
    a = arr(0) & "|" & arr(1)
    c = "フィールド名 -> ID|NAME"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_ExtractColumns_ThreeColumns()
    ' カラムの設定
    With rs
        With .Fields
            Call .Append("ID", adInteger)
            Call .Append("NAME", adVarChar, 100)
            Call .Append("AGE", adInteger)
        End With
        Call .Open
    End With
    ' RecordsetJSONConverterインスタンス初期化
    Call rjc.Init(rs)
    Dim arr As Variant
    arr = rjc.ExtractColumns()
    e = "ID|NAME|AGE"
    a = arr(0) & "|" & arr(1) & "|" & arr(2)
    c = "フィールド名 -> ID|NAME|AGE"
    Call ut.AssertEqual(e, a, c)
End Sub
' ExtractRows()メソッド
Private Sub Test_ExtractRows_OneRow()
    ' カラムの設定
    With rs
        With .Fields
            Call .Append("ID", adInteger)
            Call .Append("NAME", adVarChar, 100)
        End With
        Call .Open
    End With
    ' レコード追加
    Call rs.AddNew
    rs("ID") = 1
    rs("NAME") = "Steve"
    Call rs.Update
    ' RecordsetJSONConverterインスタンス初期化
    Call rjc.Init(rs)
    Dim arr As Variant
    arr = rjc.ExtractRows()
    Dim cnt As Long
    cnt = UBound(arr) - LBound(arr) + 1
    e = "1 rows affected. 1|Steve"
    a = CStr(cnt) & " rows affected. " & arr(0)(0) & "|" & arr(0)(1)
    c = "1行分の結果セット -> ID: 1, NAME: Steve"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_ExtractRows_NoRows()
    ' カラムの設定
    With rs
        With .Fields
            Call .Append("ID", adInteger)
            Call .Append("NAME", adVarChar, 100)
        End With
        Call .Open
    End With
    ' レコード追加しない
    ' RecordsetJSONConverterインスタンス初期化
    Call rjc.Init(rs)
    Dim arr As Variant
    arr = rjc.ExtractRows()
    Dim cnt As Long
    cnt = UBound(arr) - LBound(arr) + 1
    e = "0 rows affected."
    a = CStr(cnt) & " rows affected."
    c = "結果セットなし -> 0 rows affected."
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_ExtractRows_MultiRows()
    ' カラムの設定
    With rs
        With .Fields
            Call .Append("ID", adInteger)
            Call .Append("NAME", adVarChar, 100)
        End With
        Call .Open
    End With
    ' レコードを2行分追加
    Call rs.AddNew
    rs("ID") = 1
    rs("NAME") = "Steve"
    Call rs.Update
    Call rs.AddNew
    rs("ID") = 2
    rs("NAME") = "Bruce"
    Call rs.Update
    ' RecordsetJSONConverterインスタンス初期化
    Call rjc.Init(rs)
    Dim arr As Variant
    arr = rjc.ExtractRows()
    Dim cnt As Long
    cnt = UBound(arr) - LBound(arr) + 1
    e = "2 rows affected. 1|Steve, 2|Bruce"
    a = CStr(cnt) & " rows affected. " & _
        arr(0)(0) & "|" & arr(0)(1) & ", " & _
        arr(1)(0) & "|" & arr(1)(1)
    c = "1行分の結果セット -> ID: 1, NAME: Steve; ID: 2, NAME: Bruce"
    Call ut.AssertEqual(e, a, c)
End Sub
Private Sub Test_ExtractRows_MultiRowsIncludeNull()
    ' カラムの設定
    With rs
        With .Fields
            Call .Append("ID", adInteger)
            Call .Append("NAME", adVarChar, 100, adFldIsNullable)
        End With
        Call .Open
    End With
    ' レコードを2行分追加
    Call rs.AddNew
    rs("ID") = 1
    rs("NAME") = "Steve"
    Call rs.Update
    Call rs.AddNew
    rs("ID") = 2
    rs("NAME") = Null
    Call rs.Update
    Call rs.AddNew
    rs("ID") = 3
    rs("NAME") = "Bruce"
    Call rs.Update
    ' RecordsetJSONConverterインスタンス初期化
    Call rjc.Init(rs)
    Dim arr As Variant
    arr = rjc.ExtractRows()
    Dim cnt As Long
    cnt = UBound(arr) - LBound(arr) + 1
    e = "3 rows affected. 2nd row's NAME column is Null"
    a = CStr(cnt) & " rows affected. " & _
        "2nd row's NAME column is " & TypeName(arr(1)(1))
    c = "3行の結果セット中、2行目のNAMEカラムの値がNULL"
    Call ut.AssertEqual(e, a, c)
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

Private Sub AA_SampleErrorTestProcedures(): End Sub
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


