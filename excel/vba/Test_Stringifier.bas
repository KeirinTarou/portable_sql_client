Attribute VB_Name = "Test_Stringifier"
Option Explicit

Private Const MODULE_NAME As String = "Test_Stringifier"

Private Const vbext_pk_Proc As Long = 0

Private ut As UnitTest
Private sf As New StringFormatter

Private m_SucceededCount As Long
Private m_FailedCount As Long

Private m_TestProcNames As Object

' Stringifierインスタンス
Private strf As New Stringifier

Private dic As Object
Private col As Collection
Private m_Exp As String

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

    'Call RunTestsByName(cfg, "Test_*")
    Call RunTestsByName(cfg, "Test_ArrayToString2D_PP")
    'Call RunTestsByName(cfg, "Test_DictionaryToString_KStrV2DArrayPP")
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
    Set dic = CreateObject("Scripting.Dictionary")
    ' テストプロシージャの名前を格納するディクショナリ
    Set m_TestProcNames = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Teardown()
    ' 設定破棄用コードをここに実装
    ' UnitTestオブジェクトのログを破棄
    Call ut.ClearTestLogs
End Sub

Private Sub AA_TestCodeHere(): End Sub
' =============================================================================
'   Test Procedures
' =============================================================================
Private Sub Test_DictionaryToString_Blank()
    Call ut.AssertEqual("{}", strf.ToString(dic, False), _
        "空の辞書/詳細表示なし/整形なし")
End Sub
Private Sub Test_DictionaryToString_BlankPP()
    Call ut.AssertEqual("{}", strf.ToString(dic, False, True), _
        "空の辞書/詳細表示なし/整形あり")
End Sub
Private Sub Test_DictionaryToString_BlankDetail()
    Call ut.AssertEqual("<Dictionary> {}", strf.ToString(dic, True, False), _
        "空の辞書/詳細表示あり/整形なし")
End Sub
Private Sub Test_DictionaryToString_BlankDetailPP()
    Call ut.AssertEqual("<Dictionary> {}", strf.ToString(dic, True, True), _
        "空の辞書/詳細表示あり/整形あり")
End Sub
Private Sub Test_DictionaryToString_KStrVStr()
    dic("pachinko") = "123"
    Call ut.AssertEqual("{""pachinko"": ""123""}", strf.ToString(dic, False), _
        "Key: String/Val: String/詳細表示なし/整形なし")
End Sub
Private Sub Test_DictionaryToString_KStrVInt()
    dic("pachinko") = 123
    Call ut.AssertEqual("{""pachinko"": 123}", strf.ToString(dic, False), _
        "Key: String/Val: Integer/詳細表示なし/整形なし")
End Sub
Private Sub Test_DictionaryToString_KStrVIntDetail()
    dic("pachinko") = 123
    m_Exp = "<Dictionary> {<String> ""pachinko"": <Integer> 123}"
    Call ut.AssertEqual(m_Exp, strf.ToString(dic, True), _
        "Key: String/Val: Integer/詳細表示あり/整形なし")
End Sub
Private Sub Test_DictionaryToString_KStrVIntPP()
    dic("pachinko") = 123
    m_Exp = "{" & sf.NL & sf.IND & """pachinko"": 123" & sf.NL & "}"
    Call ut.AssertEqual(m_Exp, strf.ToString(dic, False, True), _
        "Key: String/Val: Integer/詳細表示なし/整形あり")
End Sub
Private Sub Test_DictionaryToString_KStrVIntDetailPP()
    dic("pachinko") = 123
    m_Exp = "<Dictionary> {" & sf.NL & _
        sf.IND & "<String> ""pachinko"": <Integer> 123" & sf.NL & "}"
    Call ut.AssertEqual(m_Exp, strf.ToString(dic, True, True), _
        "Key: String/Val: Integer/詳細表示あり/整形あり")
End Sub
Private Sub Test_DictionaryToString_KStrVInt2ElemsPP()
    dic("pachinko") = 123
    dic("slot") = 123
    m_Exp = "{" & sf.NL & sf.IND & """pachinko"": 123, " & sf.NL & _
        sf.IND & """slot"": 123" & sf.NL & "}"
    Call ut.AssertEqual(m_Exp, strf.ToString(dic, False, True), _
        "Key: String/Val: Integer/2要素/詳細表示なし/整形あり")
End Sub
Private Sub Test_DictionaryToString_KStrVArray()
    Set dic = CreateObject("Scripting.Dictionary")
    dic("sponsored") = Array("nobuta", "group")
    Call ut.AssertEqual("{""sponsored"": [0: ""nobuta"", 1: ""group""]}", strf.ToString(dic, False), _
        "Key: String/Val: Array/詳細表示なし/整形なし")
End Sub
Private Sub Test_DictionaryToString_KStrV2DArray()
    Set dic = CreateObject("Scripting.Dictionary")
    Dim arr(0 To 1, 0 To 1)
    arr(0, 0) = "pachinko": arr(0, 1) = "123"
    arr(1, 0) = "slot": arr(1, 1) = "123"
    dic("sponsored") = arr
    Call ut.AssertEqual("{""sponsored"": [[0, 0: ""pachinko"", 0, 1: ""123""], [1, 0: ""slot"", 1, 1: ""123""]]}", _
        strf.ToString(dic, False), _
        "Key: String/Val: 2DArray/詳細表示なし/整形なし")
End Sub
Private Sub Test_ArrayToString2D_PP()
    Dim arr(0 To 1, 0 To 1)
    arr(0, 0) = "pachinko": arr(0, 1) = "123"
    arr(1, 0) = "slot": arr(1, 1) = "123"
    m_Exp = _
        PP( _
            "[", _
            sf.IND & "[", _
            sf.IND(2) & "0, 0: ""pachinko"", ", _
            sf.IND(2) & "0, 1: ""123""", _
            sf.IND & "], ", _
            sf.IND & "[", _
            sf.IND(2) & "1, 0: ""slot"", ", _
            sf.IND(2) & "1, 1: ""123""", _
            sf.IND & "]", _
            "]" _
        )
    Call ut.AssertEqual(m_Exp, strf.ToString(arr, False, True), _
        "2次元配列/詳細なし/整形あり")
End Sub
Private Sub Test_DictionaryToString_KStrV2DArrayPP()
    Set dic = CreateObject("Scripting.Dictionary")
    Dim arr(0 To 1, 0 To 1)
    arr(0, 0) = "pachinko": arr(0, 1) = "123"
    arr(1, 0) = "slot": arr(1, 1) = "123"
    dic("sponsored") = arr
    m_Exp = PP( _
        "{", _
        sf.IND & """sponsored"": [", _
        sf.IND(2) & "[", _
        sf.IND(3) & "0, 0: ""pachinko"", ", _
        sf.IND(3) & "0, 1: ""123""", _
        sf.IND(2) & "], ", _
        sf.IND(2) & "[", _
        sf.IND(3) & "1, 0: ""slot"", ", _
        sf.IND(3) & "1, 1: ""123""", _
        sf.IND(2) & "]", _
        sf.IND & "]", _
        "}" _
    )
    Call ut.AssertEqual(m_Exp, strf.ToString(dic, False, True), _
        "Key: String/Val: 2DArray/詳細表示なし/整形あり")
End Sub
Private Sub Test_CollectionToString_2Elems()
    Set col = New Collection
    Call col.Add("pachinko")
    Call col.Add("123")
    m_Exp = "(""pachinko"", ""123"")"
    Call ut.AssertEqual(m_Exp, strf.ToString(col, False), _
        "コレクション/詳細表示なし/整形なし")
End Sub
Private Sub Test_CollectionToString_2ElemsPP()
    Set col = New Collection
    Call col.Add("pachinko")
    Call col.Add("123")
    m_Exp = PP( _
        "(", _
        sf.IND & """pachinko"", ", _
        sf.IND & """123""", _
        ")" _
    )
    Call ut.AssertEqual(m_Exp, strf.ToString(col, False, True), _
        "コレクション/詳細表示なし/整形あり")
End Sub


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


