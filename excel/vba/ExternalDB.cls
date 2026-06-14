VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExternalDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
'   Dependency
'       - StringFormatter.cls
' =============================================================================

Private sf As New StringFormatter

Private Enum StreamTypeEnum
  adTypeBinary = 1
  adTypeText = 2
End Enum

Private Enum CursorTypeEnum
    adOpenUnspecified = -1
    adOpenForwardOnly = 0
    adOpenKeyset = 1
    adOpenDynamic = 2
    adOpenStatic = 3
End Enum

Private Enum CommandTypeEnum
    adCmdUnSpecified = -1
    adCmdText = 1
    adCmdTable = 2
    adCmdStoredProc = 4
    adCmdUnKnown = 8
    adCmdFile = 256
    adCmdTableDirect = 512
End Enum

Private Enum ObjectStateEnum
    adStateClosed = 0
    adStateOpen = 1
    adStateConnecting = 2
    adStateExecuting = 4
    adStateFetching = 8
End Enum

Private Enum CursorLocationEnum
    adUseClient = 3
    adUseNone = 1
    adUseServer = 2
End Enum

Private Enum CursorOptionEnum
    adAddNew = &H1000400
    adApproxPosition = &H4000
    adBookmark = &H2000
    adDelete = &H1000800
    adFind = &H80000
    adHoldRecords = &H100
    adIndex = &H100000
    adMovePrevious = &H200
    adNotify = &H40000
    adResync = &H20000
    adSeek = &H200000
    adUpdate = &H1008000
    adUpdateBatch = &H10000
End Enum

Private Enum LockTypeEnum
    adLockUnspecified = -1
    adLockReadOnly = 1
    adLockPessimistic = 2
    adLockOptimistic = 3
    adLockBatchOptimistic = 4
End Enum

Private Enum DataTypeEnum
    adInteger = 3
    adDouble = 5
    adDBTimeStamp = 135
    adVarChar = 200
    adVarWChar = 202
End Enum

Private Enum ParameterDirectionEnum
    adParamUnknown = 0
    adParamInput = 1
    adParamOutput = 2
    adParamInputOutput = 3
    adParamReturnValue = 4
End Enum

Private m_FSO As Object

' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' 家庭用
Public Property Get CONNECTION_STRING() As String
    CONNECTION_STRING = "DSN=MySQL_DB_FOR_VBA;UID=user;PWD=userpass"
End Property

Public Property Get TEST_QUERY_NO_PARAM() As String
    TEST_QUERY_NO_PARAM = "SELECT * FROM Prefecturals AS pf WHERE pf.PrefecturalID = 28;"
End Property

Public Property Get TEST_QUERY_WITH_PARAM() As String
    TEST_QUERY_WITH_PARAM = "SELECT * FROM Prefecturals AS pf WHERE pf.PrefecturalID = ?;"
End Property

Public Property Get TEST_QUERY_WITH_PARAMS() As String
    TEST_QUERY_WITH_PARAMS = _
        "SELECT * FROM Prefecturals AS pf WHERE pf.PrefecturalID = ? AND pf.PrefecturalName = ?;"
End Property
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' 職場用
'Public Property Get CONNECTION_STRING() As String
'    CONNECTION_STRING = _
'        "DRIVER={Oracle in OraClient12Home1_32bit};SERVER=GTA001;UID=end_user;PWD=kps;" & _
'        "DBQ=GTA001;DBA=W;APA=T;EXC=F;XSM=Default;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;" & _
'        "BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=Me;CSR=F;FWC=F;FBS=60000;TLO=O;" & _
'        "MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;"
'End Property
'
'Public Property Get TEST_QUERY_NO_PARAM() As String
'    TEST_QUERY_NO_PARAM = "SELECT * FROM tbl_part p WHERE p.model='M5' AND p.cn = ' CA300';"
'End Property
'
'Public Property Get TEST_QUERY_WITH_PARAM() As String
'    TEST_QUERY_WITH_PARAM = "SELECT * FROM tbl_part p WHERE p.model='M5' AND p.cn = ' ' & TRIM(?);"
'End Property
'
'Public Property Get TEST_QUERY_WITH_PARAMS() As String
'    TEST_QUERY_WITH_PARAMS = "SELECT * FROM tbl_part p WHERE p.model = ? AND p.cn = ' ' & TRIM(?);"
'End Property
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' クエリの結果セットのうち、1行1列目のデータを取得する
' 基本的にスカラクエリの結果を取るために使う
Public Function FetchScalar( _
            ByVal a_ConnectionString As String, _
            ByVal a_Query As String, _
   Optional ByVal a_Params As Variant) As Variant
' === Usage ===============================================
'   Dim result As Variant
'   result = edb.FetchScalar( _
'       "DRIVER={xxx};SERVER=xxx;UID=xxx;PWD=xxx;", _
'       "SELECT p.pn FROM tbl_part p WHERE p.model = ? AND p.cn = ?;", _
'       Array("M5", " CA300"))
'   ' result: "TF309051G10"
' =========================================================
    On Error GoTo HandleError
    Dim ret As Variant
    ' 初期値を闘魂注入
    ret = Null
    
    Dim paramArr As Variant
    ' 引数`a_Params`が省略された -> Emptyにする
    If IsMissing(a_Params) Then
        paramArr = Empty
    ' 引数`a_Params`が渡された -> 中身を調べて正規化
    Else
        paramArr = NormalizeParams(a_Params)
    End If
    
    ' 1. クエリを実行してレコードセットを取得
    Dim rs As Object
    ' パラメータなしのときは`a_Params`がEmptyのはず
    If IsEmpty(paramArr) Then
        Set rs = ExecQuery(a_ConnectionString, a_Query)
    Else
        Set rs = ExecQuery(a_ConnectionString, a_Query, paramArr)
    End If
    
    ' 2. `rs`が`Nothing`だったらそのまま終了
    If rs Is Nothing Then GoTo Finally
    
    ' 2. 結果を1件取得
    If Not (rs.EOF And rs.BOF) Then
        ' 結果が複数ある場合は警告を出しておく
        If rs.Fields.Count > 1 Or _
           (rs.Supports(adApproxPosition) And rs.RecordCount > 1) Then
            Debug.Print "[Warning!!]: "
            Debug.Print "FetchScalar() detected multi row/column result."
            Debug.Print "TopLeft value fallback."
        End If
        ' フィールドの値を安全に取り出し -> 危険な値はNullにフォールバック
        ret = GetFieldValueOrNull(rs.Fields.Item(0))
    End If
    If rs.State <> adStateClosed Then Call rs.Close
    Set rs = Nothing
    GoTo Finally
HandleError:
    ' 例外をそのまま呼び出し元にスロー
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
Finally:
    FetchScalar = ret
End Function

' クエリの結果セットのうち先頭行のレコードを返す
' 基本的に、1行分だけレコードセットを返すクエリを用いる
' 1行文のレコードをカラム名と値のディクショナリに詰め込んで返す
Public Function FetchOne( _
            ByVal a_ConnectionString As String, _
            ByVal a_Query As String, _
   Optional ByVal a_Params As Variant) As Object
' === Usage ===============================================
'   Dim result As Object    ' Scripting.Dictionary object
'   set result = edb.FetchOne( _
'       "DRIVER={xxx};SERVER=xxx;UID=xxx;PWD=xxx;", _
'       "SELECT p.pn, p.part_name, p.std_span,  FROM tbl_part p WHERE p.model = ? AND p.cn = ?;", _
'       Array("M5", " CA300"))
'   ' result: {"PN": "TF309051G10", "PART_NAME": "FRAME,COMB.CHAM.", "STD_SPAN": 114}
' =========================================================
    Dim ret As Object   ' Scripting.Dictionary object
    ' 初期値: 空のディクショナリ
    Set ret = CreateObject("Scripting.Dictionary")
    
    On Error GoTo HandleError
    
    Dim paramArr As Variant
    ' 引数`a_Params`が省略された -> Emptyにする
    If IsMissing(a_Params) Then
        paramArr = Empty
    ' 引数`a_Params`が渡された -> 中身を調べて正規化
    Else
        paramArr = NormalizeParams(a_Params)
    End If
    
    Dim rs As Object
    If IsEmpty(paramArr) Then
        Set rs = ExecQuery(a_ConnectionString, a_Query)
    Else
        Set rs = ExecQuery(a_ConnectionString, a_Query, paramArr)
    End If
    
    If rs Is Nothing Then GoTo Finally
    ' EOF, BOFともにTrue -> レコードなし -> 空ディクショナリ
    If rs.EOF And rs.BOF Then GoTo Finally
    
    ' ここへ来た -> 1件以上レコードあり -> カーソルを先頭に移動
    Call rs.MoveFirst
    
    If Not rs.EOF Then
        ' 件数確認のために末尾までスキャン
        Call rs.MoveLast
        If rs.RecordCount > 1 Then
            Debug.Print "[Warning!!]: "
            Debug.Print "FetchOne() detected multi rows result."
            Debug.Print "First row fallback."
        End If
        ' カーソルを先頭に戻す
        Call rs.MoveFirst
    End If

    ' 1行分の結果セットをディクショナリに詰め込む
    '   -> この時点でカーソルが先頭にあることが保障されている
    '   -> 先頭行のフィールドのみ取得
    Dim n As Long
    For n = 0 To rs.Fields.Count - 1
        ' 各フィールドの値を安全に取り出し -> 危険な値はNullにフォールバック
        ret(rs.Fields(n).Name) = GetFieldValueOrNull(rs.Fields(n))
    Next
    GoTo Finally
    
HandleError:
    Debug.Print "FetchOne() Error:"
    Debug.Print vbTab & "Number: " & Err.Number
    Debug.Print vbTab & "Description: " & Err.Description
    Debug.Print vbTab & "Source: " & Err.Source
    ' 例外を再スロー
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
Finally:
    If Not rs Is Nothing Then
        If rs.State <> adStateClosed Then Call rs.Close
        Set rs = Nothing
    End If
    Set FetchOne = ret
End Function

' クエリの結果セット全件を返す
' 1件のレコードをDictionaryとし、
' DictionaryをVariant配列（0始まり）に詰め込んで返す
Public Function FetchAll( _
            ByVal a_ConnectionString As String, _
            ByVal a_Query As String, _
   Optional ByVal a_Params As Variant) As Variant
' === Usage ===============================================
'   Dim result As Variant    'Variant array of Scripting.Dictionary object
'   result = edb.FetchAll( _
'       "DRIVER={xxx};SERVER=xxx;UID=xxx;PWD=xxx;", _
'       "SELECT p.pn, p.part_name, p.std_span " & _
'       "FROM tbl_part p " & _
'       "WHERE p.model = ? AND p.cn LIKE ?;", _
'       Array("M5", " CA3%"))
'   result:
'          [
'               {"MODEL": "M5", "CN": " CA300", "PN": "TF309051G10"},
'               {"MODEL": "M5", "CN": " CA301", "PN": "TF309051P11"},
'               {"MODEL": "M5", "CN": " CA303", "PN": "TF309051P12"},
'               {"MODEL": "M5", "CN": " CA3T1", "PN": "TF309051G10TR1"},
'          ]
' =========================================================
    'Dim ret() As Variant    ' Variant array of Scripting.Dictionary object
    Dim ret As Variant
    ret = Array()
    
    On Error GoTo HandleError
    
    Dim paramArr As Variant
    ' 引数`a_Params`が省略された -> Emptyにする
    If IsMissing(a_Params) Then
        paramArr = Empty
    ' 引数`a_Params`が渡された -> 中身を調べて正規化
    Else
        paramArr = NormalizeParams(a_Params)
    End If
    
    ' 1. Recordset取得
    Dim rs As Object
    ' ParamArrayに何も渡さなかったとき -> paramArrがEmptyになる
    If IsEmpty(paramArr) Then
        Set rs = ExecQuery(a_ConnectionString, a_Query)
    Else
        Set rs = ExecQuery(a_ConnectionString, a_Query, paramArr)
    End If
    If rs Is Nothing Then GoTo EmptyReturn
    If rs.EOF And rs.BOF Then GoTo EmptyReturn
    
    ' 2. 行数カウント
    Dim cnt As Long
    cnt = 0
    Call rs.MoveFirst
    Do Until rs.EOF
        cnt = cnt + 1
        Call rs.MoveNext
    Loop

    ' この時点でレコード数0 -> 空の配列を返す
    If cnt = 0 Then GoTo EmptyReturn
    
    ' 3. 先頭のレコードに戻って配列に詰め込む
    Call rs.MoveFirst
    ReDim ret(0 To cnt - 1)
    
    Dim dict As Object
    Dim f As Object
    Dim n As Long
    For n = 0 To cnt - 1
        Set dict = CreateObject("Scripting.Dictionary")
        For Each f In rs.Fields
            Call dict.Add(Key:=f.Name, Item:=GetFieldValueOrNull(f))
        Next
        Set ret(n) = dict
        Call rs.MoveNext
    Next

    GoTo Finally
    
HandleError:
    Debug.Print "FetchAll() Error:"
    Debug.Print sf.IND & "Number: " & Err.Number
    Debug.Print sf.IND & "Description: " & Err.Description
    Debug.Print sf.IND & "Source: " & Err.Source
    Debug.Print "Query:"
    Debug.Print a_Query
    If Not IsEmpty(paramArr) Then _
        Debug.Print "Params: " & Join(paramArr, ", ")
    ' 例外を再スロー
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
EmptyReturn:
    ' 結果セットなし -> 空の配列を返す
    ret = Array()
    
Finally:
    If Not rs Is Nothing Then
        If rs.State <> adStateClosed Then Call rs.Close
        Set rs = Nothing
    End If
    
    FetchAll = ret
End Function

Private Sub AA_HelperFunctions(): End Sub
' =============================================================================
'   Helper Functions
' =============================================================================
' Fieldオブジェクトから安全に値を取り出す
Friend Function GetFieldValueOrNull( _
            ByVal a_Field As Object, _
   Optional ByVal a_TestMode As Boolean = False) As Variant
    Dim ret As Variant
    ' フォールバック値は`Null`
    ret = Null
    
    On Error Resume Next
    ret = a_Field.Value
    If Err.Number = 0 Then GoTo Finally
    Dim errNum As Long, errSrc As String, errDesc As String, _
        fldName As String, fldType As DataTypeEnum, fldTypeName As String
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    ' エラーは握りつぶす
    Call Err.Clear
    On Error GoTo 0
    ' Fieldオブジェクト情報を収集
    On Error Resume Next
    fldName = a_Field.Name
    If Err.Number <> 0 Then fldName = "(Error)": Call Err.Clear
    ' テストモードではTypeの検証をあきらめる
    '   - Typeプロパティが予約語の壁に阻まれるため
    If a_TestMode Then
        fldType = -2
    Else
        fldType = a_Field.Type
    End If
    If Err.Number <> 0 Then fldType = -1: Call Err.Clear
    fldTypeName = GetFieldTypeName(fldType)
    On Error GoTo 0

    ' デバッグ用出力
    Debug.Print "Number: " & CStr(errNum)
    Debug.Print "Source: " & errSrc
    Debug.Print "Desc  : " & errDesc
    Debug.Print "Name  : " & fldName
    Debug.Print "Type  : " & fldTypeName & "(" & CStr(fldType) & ")"
    
    GoTo Finally
Finally:
    GetFieldValueOrNull = ret
End Function

Private Function GetFieldTypeName( _
            ByVal a_DataType As DataTypeEnum) As String
    Dim ret As String
    
    Select Case a_DataType
        Case adInteger: ret = "Integer"
        Case adDouble: ret = "Double"
        Case adVarChar: ret = "VarChar"
        Case adVarWChar: ret = "VarWChar"
        Case adDBTimeStamp: ret = "DBTimeStamp"
        Case -1: ret = "Error"
        Case -2: ret = "Test"
        Case Else: ret = "Unknown"
    End Select
    
    GetFieldTypeName = ret
End Function

' 引数として受け取ったパラメータを正規化する
Private Function NormalizeParams( _
            ByVal a_Params As Variant) As Variant
    Dim ret As Variant
    
    If IsEmpty(a_Params) Then
        ret = Empty
        GoTo Finally
    End If
    
    If Not IsArray(a_Params) Then
        ReDim ret(0 To 0)
        ret(0) = a_Params
        GoTo Finally
    End If
    
    If UBound(a_Params) < 0 Then
        ret = Empty
        GoTo Finally
    End If
    
    ReDim ret(0 To UBound(a_Params))
    Dim i As Long
    For i = LBound(a_Params) To UBound(a_Params)
        ret(i) = a_Params(i)
    Next
    
Finally:
    NormalizeParams = ret
End Function

' クエリを実行してRecordsetインスタンスを返す
Friend Function ExecQuery( _
            ByVal a_ConnectionString As String, _
            ByVal a_Query As String, _
   Optional ByVal a_Params As Variant) As Object
    Dim ret As Object
    
    On Error GoTo HandleError
    
    ' 1. ADODB.Connectionインスタンス取得
    Dim CONN As Object
    Set CONN = CreateObject("ADODB.Connection")
    CONN.CursorLocation = adUseClient
    Call CONN.Open(a_ConnectionString)
    
    ' 2. Commandオブジェクト作成
    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = CONN
    ' プレースホルダ付きクエリを一旦セット
    cmd.CommandText = a_Query
    cmd.CommandType = adCmdText
    
    ' パラメータをプレースホルダにバインド
    Dim i As Long, paramsStr As String
    paramsStr = "(No param)"
    If IsArray(a_Params) Then
        Dim param As Object, v As Variant, s As String
        Dim lb As Long, ub As Long
        lb = LBound(a_Params): ub = UBound(a_Params)
        For i = lb To ub
            If i = lb Then paramsStr = "("
            paramsStr = paramsStr & CStr(a_Params(i))
            If i <> ub Then paramsStr = paramsStr & ","
            If i = ub Then paramsStr = paramsStr & ")"
            v = a_Params(i)
            ' CreateParameter()の引数
            '   - Name, Type, Direction, Size, Value
            If IsNull(v) Then
                Set param = _
                    cmd.CreateParameter( _
                        "param_" & Format(i, "00#"), _
                        adVarChar, _
                        adParamInput, _
                        1, _
                        Null _
                    )
            Else
                ' 日付型は文字列に寄せる
                If IsDate(v) Then v = Format$(v, "yyyy-mm-dd")
                s = CStr(v)
                Set param = _
                    cmd.CreateParameter( _
                        "param_" & Format(i, "00#"), _
                        adVarChar, _
                        adParamInput, _
                        LenB(s), _
                        s _
                    )
            End If
            Call cmd.Parameters.Append(param)
        Next
    End If
    
    ' 3. Recordsetインスタンスを作成
    Set ret = CreateObject("ADODB.Recordset")
    ret.CursorLocation = adUseClient
    ret.CursorType = adOpenStatic
    ret.LockType = adLockReadOnly
    ' 【参考】
    ' [ADODB.Recordset].Open()の引数リスト
    '   Source As String | As ADODB.Command | As ADODB.Recordset | As ADODB.Stream,
    '   [ActiveConnection As ADODB.Connection, ]
    '   [CursorType As ADODB.CursorTypeEnum, ]
    '   [LockType As ADODB.LockTypeEnum, ]
    '   [Options As ADODB.CommandTypeEnum ]
    Call ret.Open(cmd)
    
    ' 4. RecordsetをConnectionから切り離す
    Set ret.ActiveConnection = Nothing
    ' ここでConnectionインスタンスを殺してしまう
    Call CONN.Close
    Set CONN = Nothing
    
    ' 5. Commandオブジェクトも用なし -> 抹殺
    Set cmd = Nothing
    
    ' 6. 結果セットを返す
    Set ExecQuery = ret
    Exit Function
    
HandleError:
    ' 例外発生時は呼び出し元に例外をスローする
    
    ' 元の例外情報をキャッシュしておく
    '   -> あとしまつ中のエラーで上書きされたとき用
    Dim errNum As Long, errSrc As String, errDesc As String
    errNum = Err.Number
    errSrc = Err.Source
    errDesc = Err.Description & paramsStr

    Call CleanUpIfError(CONN, cmd, ret)
    
    ' 元の例外を呼び出し元にスロー
    Call Err.Raise(errNum, errSrc & "[ExecQuery()]", errDesc)
End Function

Private Sub CleanUpIfError( _
            ByRef a_Connection As Object, _
            ByRef a_Command As Object, _
            ByRef a_Recordset As Object)
    
    ' Recordsetを抹殺
    If Not a_Recordset Is Nothing Then
        On Error Resume Next
        If a_Recordset.State <> adStateClosed Then Call a_Recordset.Close
        On Error GoTo 0
        Set a_Recordset = Nothing
    End If
    ' Commandを抹殺
    If Not a_Command Is Nothing Then
        On Error Resume Next
        Set a_Command.ActiveConnection = Nothing
        On Error GoTo 0
        Set a_Command = Nothing
    End If
    ' 最後にConnectionを抹殺
    If Not a_Connection Is Nothing Then
        On Error Resume Next
        If a_Connection.State <> adStateClosed Then Call a_Connection.Close
        On Error GoTo 0
        Set a_Connection = Nothing
    End If
End Sub

