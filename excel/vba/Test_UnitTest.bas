Attribute VB_Name = "Test_UnitTest"
Option Explicit

Private Const MODULE_NAME As String = "Test_UnitTest"

Private Const vbext_pk_Proc As Long = 0

Private ut As UnitTest
Private sf As New StringFormatter

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
    
'    Call RunTestsByName(cfg, "Test_AssertTrue*")
    Call RunTestsByName(cfg, "Test_AssertTrue_MultiExecution")
End Sub

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

Private Sub AA_TestCodeHere(): End Sub
' =============================================================================
'   Test Procedures
' =============================================================================

' 直近のテスト結果を返す`LastResult`プロパティの実装
' 成功 -> utTestPassed=1
Private Sub Test_LastResult_True()
    Dim lut As New UnitTest
    Call lut.AssertEqual(1, 1, "PASSするテスト。")
    
    Call ut.AssertEqual( _
        a_Expected:=utTestPassed, _
        a_Actual:=lut.LastResult, _
        a_Comment:="utTestPassed=1を返すはず。")
End Sub
' 失敗 -> utTestFailed=0
Private Sub Test_LastResult_False()
    Dim lut As New UnitTest
    Call lut.AssertEqual(1, 2, "FAILになるテスト。")
    Call ut.AssertEqual(utTestFailed, lut.LastResult, _
        "utTestFailed=0を返すはず。")
End Sub
' 未実施 -> utTestUndefined=-1
Private Sub Test_LastResult_NotTestedYet()
    Dim lut As New UnitTest
    ' テストしていない段階で呼ぶ
    Call ut.AssertEqual(utTest_Undefined, lut.LastResult, _
        "utTestUndefined=-1を返すはず。")
End Sub
' 成功 >>> 失敗の場合 -> tfTestFailed=0
Private Sub Test_LastResult_SuccessThenFail()
    Dim lut As New UnitTest
    Call lut.AssertEqual(1, 1, "PASS")
    Call lut.AssertEqual(1, 2, "FAIL")
    Call ut.AssertEqual(utTestFailed, lut.LastResult, _
        "utTestFailed=0を返すはず。")
End Sub
' 失敗 >>> 成功の場合 -> tfTestPassed=1
Private Sub Test_LastResult_FailThenSuccess()
    Dim lut As New UnitTest
    Call lut.AssertEqual(1, 2, "FAIL")
    Call lut.AssertEqual(1, 1, "PASS")
    Call ut.AssertEqual(utTestPassed, lut.LastResult, _
        "utTestPassed=1を返すはず。")
End Sub
' コメントを省略しても動作する
Private Sub Test_LastResult_NoComment()
    Dim lut As New UnitTest
    Call lut.AssertEqual(1, 1)
    Call ut.AssertEqual(utTestPassed, lut.LastResult, _
        "コメントを省略しても動作するはず。")
End Sub
' Assert系メソッドの実行がTestLogsを壊さない
Private Sub Test_LastResultDoesNotClearLogs()
    Dim lut As New UnitTest
    Call lut.AssertEqual(1, 1)
    Call lut.AssertEqual(2, 2)
    Call lut.AssertEqual(3, 4)
    Call ut.AssertEqual(3, lut.TestLogs.Count, _
        "AssertがTestLogsを消してしまわない。")
End Sub
' AssertTrue()メソッドの実装
' Trueを渡したらPASS
Private Sub Test_AssertTrue_ShouldPass_WhenTrue()
    Dim lut As New UnitTest
    Call lut.AssertTrue(True, "True -> PASS")
    Call ut.AssertEqual(utTestPassed, lut.LastResult, _
        "Trueを渡したらPASSするはず。")
End Sub
' Falseを渡したらFAIL
Private Sub Test_AssertTrue_ShouldFail_WhenFalse()
    Dim lut As New UnitTest
    Call lut.AssertTrue(False, "False -> FAIL")
    Call ut.AssertEqual(utTestFailed, lut.LastResult, _
        "Falseを渡したらFAILになるはず。")
End Sub
' ログが正しく作られるか
' 成功フラグ
Private Sub Test_AssertTrue_ShouldCreateCorrectLog_WhenTrue_SuccessFlag()
    Dim lut As New UnitTest
    Call lut.AssertTrue(True, "コメント")
    ' 直近のログを取得
    Dim lg As UnitTestLog
    Set lg = lut.TestLogs(lut.TestLogs.Count)
    Call ut.AssertEqual(True, lg.Succeeded, "成功フラグ")
End Sub
' Expected
Private Sub Test_AssertTrue_ShouldCreateCorrectLog_WhenTrue_Expected()
    Dim lut As New UnitTest
    Call lut.AssertTrue(True, "コメント")
    ' 直近のログを取得
    Dim lg As UnitTestLog
    Set lg = lut.TestLogs(lut.TestLogs.Count)
    Call ut.AssertEqual("True", CStr(lg.expected), "Expectedは""True""")
End Sub
' Actual
Private Sub Test_AssertTrue_ShouldCreateCorrectLog_WhenTrue_Actual()
    Dim lut As New UnitTest
    Call lut.AssertTrue(True, "コメント")
    ' 直近のログを取得
    Dim lg As UnitTestLog
    Set lg = lut.TestLogs(lut.TestLogs.Count)
    ' 検証
    Call ut.AssertEqual("True", CStr(lg.Actual), "Trueが渡された")
End Sub
' Comment
Private Sub Test_AssertTrue_ShouldCreateCorrectLog_WhenTrue_Comment()
    Dim lut As New UnitTest
    Call lut.AssertTrue(True, "コメント")
    ' 直近のログを取得
    Dim lg As UnitTestLog
    Set lg = lut.TestLogs(lut.TestLogs.Count)
    Call ut.AssertEqual("コメント", lg.Comment, "コメントも保存される")
End Sub
Private Sub Test_AssertTrue_MultiExecution()
'    Set ut = New UnitTest
    Dim lut1 As New UnitTest
    Call lut1.AssertTrue(True, "コメント")
    Dim lg As UnitTestLog
    Set lg = lut1.TestLogs(lut1.TestLogs.Count)
    Dim lut2 As New UnitTest
    Call lut2.AssertEqual(True, lg.Succeeded, "成功フラグ")
    Call lut2.AssertEqual("True", CStr(lg.expected), "Expectedは""True""")
    Call lut2.AssertEqual("True", CStr(lg.Actual), "Trueが渡された")
    Call lut2.AssertEqual("コメント", lg.Comment, "コメントも保存される")
    Dim expected As String
    expected = "(" & _
        "Succeeded=True;Expected=True;Actual=True;Comment=成功フラグ, " & _
        "Succeeded=True;Expected=""True"";Actual=""True"";Comment=Expectedは""True"", " & _
        "Succeeded=True;Expected=""True"";Actual=""True"";Comment=Trueが渡された, " & _
        "Succeeded=True;Expected=""コメント"";Actual=""コメント"";Comment=コメントも保存される)"
    Call ut.AssertEqual(expected, UtilStringify.ToString(lut2.TestLogs, False), "すべて想定どおり")
End Sub
' コメントを省略しても動くか
Private Sub Test_AssertTrue_NoComment()
    Dim lut As New UnitTest
    Call lut.AssertTrue(True)
    
    Dim lg As UnitTestLog
    Set lg = lut.TestLogs.Item(lut.TestLogs.Count)
    Call ut.AssertEqual("", lg.Comment, _
        "コメント省略時は空文字のはず")
End Sub
' 連続呼び出しでログが壊れないか
Private Sub Test_AssertTrue_DoesNotClearExistingLogs()
    Dim lut As New UnitTest
    Call lut.AssertTrue(True)
    Call lut.AssertTrue(True)
    Call ut.AssertEqual(2, lut.TestLogs.Count, _
        "ログが2件残るはず")
End Sub
' ログ件数は増えるか
Private Sub Test_AssertTrue_ShouldIncreaseLogCount()
    Dim lut As New UnitTest
    Dim prevCnt As Long
    prevCnt = lut.TestLogs.Count
    Call lut.AssertTrue(True)
    Call ut.AssertEqual(prevCnt + 1, lut.TestLogs.Count, _
        "ログが1件になっているはず")
End Sub
' AssertFalse()メソッドの実装
' Falseを渡したらPASS
Private Sub Test_AssertFalse_ShouldPass_WhenFalse()
    Dim lut As New UnitTest
    Call lut.AssertFalse(False, "False -> PASS")
    Call ut.AssertEqual(utTestPassed, lut.LastResult, _
        "Falseを渡したらPASSするはず。")
End Sub
' Trueを渡したらFAIL
Private Sub Test_AssertFalse_ShouldFail_WhenTrue()
    Dim lut As New UnitTest
    Call lut.AssertFalse(True, "True -> FAIL")
    Call ut.AssertEqual(utTestFailed, lut.LastResult, _
        "Trueを渡したらFAILになるはず。")
End Sub
' ここで、テスト駆動せずに実装したAssertEqual()の検証を行う
' 同じ文字列 -> PASS
Private Sub Test_AssertEqual_ShouldPass_WhenSameString()
    Dim lut As New UnitTest
    Call lut.AssertEqual("pachinko123", "pachinko123")
    Call ut.AssertEqual(utTestPassed, lut.LastResult, _
        "同じ文字列 -> PASSになるはず。")
End Sub
' 同じ数値 -> PASS
Private Sub Test_AssertEqual_ShouldPass_WhenSameNumber()
    Dim lut As New UnitTest
    Call lut.AssertEqual(123, 123)
    Call ut.AssertEqual(utTestPassed, lut.LastResult, _
        "同じ数値 -> PASSになるはず。")
End Sub
' 同じ真偽値 -> PASS
Private Sub Test_AssertEqual_ShouldPass_WhenSameBoolean()
    Dim lut As New UnitTest
    Call lut.AssertEqual(False, False)
    Call ut.AssertEqual(utTestPassed, lut.LastResult, _
        "同じ真偽値 -> PASSになるはず。")
End Sub
' 同じ配列 -> PASS
Private Sub Test_AssertEqual_ShouldPass_WhenSameArray()
    Dim lut As New UnitTest
    Call lut.AssertEqual(Array("pachinko", 123), Array("pachinko", 123))
    Call ut.AssertEqual(utTestPassed, lut.LastResult, _
        "同じ配列 -> PASSになるはず。")
End Sub
' 同じディクショナリ -> PASS
Private Sub Test_AssertEqual_ShouldPass_WhenSameDictionary()
    Dim lut As New UnitTest
    Dim dic1 As Object, dic2 As Object
    Set dic1 = CreateObject("Scripting.Dictionary")
    Set dic2 = CreateObject("Scripting.Dictionary")
    Call dic1.Add("pachinko", 123)
    Call dic2.Add("pachinko", 123)
    Call lut.AssertEqual(dic1, dic2)
    Call ut.AssertEqual(utTestPassed, lut.LastResult, _
        "同じ内容のディクショナリ -> PASSになるはず。")
End Sub
' 異なる文字列 -> FAIL
Private Sub Test_AssertEqual_ShouldFail_WhenDifferentString()
    Dim lut As New UnitTest
    Call lut.AssertEqual("pachinko123", "pachinko12")
    Call ut.AssertEqual(utTestFailed, lut.LastResult, _
        "異なる文字列 -> FAILになるはず。")
End Sub

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

Private Sub AA_Configs(): End Sub
' =============================================================================
'   Configuration Procedures
' =============================================================================
Private Sub Setup()
    ' カウンタをリセット
    Call ResetCounts
    ' UnitTestインスタンスを取得
    Set ut = New UnitTest
    
    ' イミディエイトウィンドウクリア
    Call ClearImmediateWindow
    
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

Private Sub Teardown()
    ' 設定破棄用コードをここに実装
    ' UnitTestオブジェクトのログを破棄
    Call ut.ClearTestLogs
End Sub

Private Sub AA_HelperFunctions(): End Sub
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

Private Function ClearImmediateWindow()
'    Dim win As Object
'    Set win = Application.VBE.Windows.Item(vbext_wt_Immediate)
'    win.Visible = True

    Call WinAPI.Wait(1000)
    Call Application.SendKeys("^g", 1)
    DoEvents
    
    Call WinAPI.Wait(1000)
    Call Application.SendKeys("^a", 1)
    DoEvents

    Call WinAPI.Wait(1000)
    Call Application.SendKeys("^x", 1)
    DoEvents
    
    Call WinAPI.Wait(1000)
    Call Application.SendKeys("{DEL}", 1)
    DoEvents
    
    
'    Dim vbeObject As Object
'    Set vbeObject = Application.VBE
'    Dim cb As Object, ctrl As Object
'
'    Set cb = vbeObject.CommandBars("Immediate Window")
'    Debug.Print cb.Name
'
'    For Each ctrl In cb.Controls
'        Debug.Print ctrl.ID & ":" & ctrl.Caption
'    Next
    
'    For Each cb In vbeObject.CommandBars
'        Set ctrl = cb.FindControl(ID:=3625)
'        Debug.Print cb.ID & ";" & cb.Name
'        If Not ctrl Is Nothing Then
'            Call ctrl.Execute
'            Exit For
'        End If
'    Next
'
'     On Error Resume Next
'     Call cb.Controls("Clear All").Execute
'     On Error GoTo 0
     
End Function
