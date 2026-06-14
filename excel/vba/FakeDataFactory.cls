VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeDataFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ERR_BASE_NUMBER As Long = vbObjectError + 5000

Private Enum ErrTypeEnum
    errInvalidArg = 1
    errNotEvenNumberOfArgs = 2
    errInvalidKey = 3
    errInvalidValue = 4
    errDuplicateKey = 5
    errEmptyKey = 6
    errNullKey = 7
    errBlankKey = 8
    errTooLongKey = 9
    errHasInvalidChar = 10
    errKeyMustBeScalar = 11
End Enum

Private m_Rows As Variant

Private Sub Class_Initialize()
    m_Rows = Array()
End Sub

' 1組のkey-valueペアを格納したディクショナリを返す
Public Function GenScalar( _
            ByVal a_Key As String, _
            ByVal a_Value As Variant) As Object
    Const ERR_SOURCE As String = "GenScalar()"
    
    Dim paramArr(0 To 1) As Variant
    paramArr(0) = a_Key
    ' 2つ目の引数はオブジェクトの可能性がある
    If IsObject(a_Value) Then
        Set paramArr(1) = a_Value
    Else
        paramArr(1) = a_Value
    End If
    
    On Error Resume Next
    Dim dic As Object
    Set dic = BuildDictFromArray(paramArr)
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    If errNum <> 0 Then
        Call Err.Raise(errNum, ERR_SOURCE, errDesc)
    End If
    
    Set GenScalar = dic
End Function

' 複数組のkey-valueペアを格納したディクショナリを返す
Public Function GenRow( _
            ParamArray a_Params()) As Object
    Const ERR_SOURCE As String = "GenRow()"
    
    On Error Resume Next
    Dim dic As Object
    Set dic = BuildDictFromArray(a_Params)
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    If errNum <> 0 Then
        Call Err.Raise(errNum, ERR_SOURCE, errDesc)
    End If
    
    Set GenRow = dic
End Function

' 新たに1行分のディクショナリを作って配列に追加 -> 追加後の配列を返す
Public Function AddRow( _
            ParamArray a_Params()) As Variant
    Const ERR_SOURCE As String = "AddRow()"
    
    On Error Resume Next
    Dim dic As Object
    Set dic = BuildDictFromArray(a_Params)
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    If errNum <> 0 Then
        Call Err.Raise(errNum, ERR_SOURCE, errDesc)
    End If
    
    Dim lb As Long, ub As Long
    lb = LBound(m_Rows)
    ub = UBound(m_Rows)
    ReDim Preserve m_Rows(lb To ub + 1)
    Set m_Rows(ub + 1) = dic
    AddRow = m_Rows
End Function

Public Sub Clear()
    m_Rows = Array()
End Sub

Private Sub ZZ_HelperFunctions(): End Sub
' =============================================================================
'   Helper functions
' =============================================================================
Private Function BuildDictFromArray( _
            ByVal a_Params As Variant) As Object
    Const ERR_SOURCE As String = "BuildDictFromArray()"
    Const ERR_DESC_0 As String = "引数は偶数個（2個以上）でなければならない。"
    Const ERR_DESC_1 As String = "奇数番目の引数は`String`でなければならない。"
    Const ERR_DESC_2 As String = "偶数番目の引数はスカラ値でなければならない。"
    Const ERR_DESC_3 As String = "奇数番目の引数に重複があってはならない。"
    Const ERR_DESC_4 As String = "奇数番目の引数を空文字にすることは許さん！"
    Const ERR_DESC_5 As String = "奇数番目の引数が長すぎるのは許さん！"
    Const ERR_DESC_6 As String = "奇数番目の引数に不正文字が混入しているのは許さん！"
    Const ERR_DESC_7 As String = "奇数番目の引数はスカラ値でなければならない。"
    Const ERR_DESC_8 As String = "奇数番目の引数をEmptyにすることは許さん！"
    Const ERR_DESC_9 As String = "奇数番目の引数をNullにすることは許さん！"
    
    Dim lb As Long, ub As Long
    lb = LBound(a_Params)
    ub = UBound(a_Params)
    ' 引数の個数が奇数 -> 例外スロー
    If ub < 0 Or (ub - lb + 1) Mod 2 = 1 Then _
        Call RaiseError(errNotEvenNumberOfArgs, ERR_SOURCE, ERR_DESC_0)
    ' ディクショナリ作成
    Dim i As Long, k As Variant, v As Variant
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    For i = lb To ub - 1 Step 2
        ' キーがスカラ値でない
        If IsArray(a_Params(i)) Or IsObject(a_Params(i)) Then _
            Call RaiseError(errKeyMustBeScalar, ERR_SOURCE, ERR_DESC_7)
        ' キーがEmpty
        If IsEmpty(a_Params(i)) Then _
            Call RaiseError(errEmptyKey, ERR_SOURCE, ERR_DESC_8)
        ' キーがNull
        If IsNull(a_Params(i)) Then _
            Call RaiseError(errNullKey, ERR_SOURCE, ERR_DESC_9)
        ' キーが文字列でない -> 例外スロー
        If TypeName(a_Params(i)) <> "String" Then _
            Call RaiseError(errInvalidKey, ERR_SOURCE, ERR_DESC_1)
        k = StripString(a_Params(i))
        ' 空文字のキーは許さん！
        If k = "" Then _
            Call RaiseError(errBlankKey, ERR_SOURCE, ERR_DESC_4)
        ' 長すぎるキーは許さん！
        If LenB(k) > 256 * 2 Then _
            Call RaiseError(errTooLongKey, ERR_SOURCE, ERR_DESC_5)
        If HasInvalidChar(k) Then _
            Call RaiseError(errHasInvalidChar, ERR_SOURCE, ERR_DESC_6)
        ' 値が直値でない
        If IsArray(a_Params(i + 1)) Or IsObject(a_Params(i + 1)) Then _
            Call RaiseError(errInvalidValue, ERR_SOURCE, ERR_DESC_2)
        
        v = a_Params(i + 1)
        ' 値がEmpty -> Nullに寄せる（`""`とは別扱い）
        If IsEmpty(v) Then v = Null
        ' ディクショナリに闘魂注入
        ' キーが重複している -> 例外スロー
        If dic.Exists(k) Or dic.Exists(UCase(k)) Or dic.Exists(LCase(k)) Then _
            Call RaiseError(errDuplicateKey, ERR_SOURCE, ERR_DESC_3)
        ' 闘魂注入！
        dic(k) = v
    Next
    Set BuildDictFromArray = dic
End Function

Private Function StripString( _
            ByVal a_String As String) As String
    Dim ret As String
    ret = a_String
    ret = Trim$(ret)
    ret = Replace$(ret, vbTab, "")
    ret = Replace$(ret, vbCrLf, "")
    ret = Replace$(ret, vbCr, "")
    ret = Replace$(ret, vbLf, "")
    ' NBSP対応
    ret = Replace$(ret, Chr(160), "")
    StripString = ret
End Function

Private Function HasInvalidChar( _
            ByVal a_String As String) As Boolean
    HasInvalidChar = True
    
    Dim i As Long, cd As Long
    
    ' 1文字ずつ切り出して検証 -> 不正文字があったらTrue
    For i = 1 To Len(a_String)
        cd = AscW(Mid$(a_String, i, 1))
        ' 制御文字の検証
        '   -> Chr(0)～Chr(31), Chr(127)はExit
        If (cd >= 0 And cd <= 31) Or cd = 127 Then Exit Function
        
        ' ゼロ幅スペース等の検証
        If cd = &H200B Or cd = &H200C Or cd = &H200D Then Exit Function
        
        ' サロゲート単体（不正UTF-16）の検証
        If cd >= &HD800 And cd <= &HDFFF Then Exit Function
    Next
    ' ここまでたどり着いたら不正文字なし
    HasInvalidChar = False
End Function

Private Sub RaiseError( _
            ByVal a_ErrorType As ErrTypeEnum, _
            ByVal a_Source As String, _
            ByVal a_Description As String)
    Const CLASS_NAME As String = "FakeDataFactory"
    Call Err.Raise( _
        Number:=ERR_BASE_NUMBER + a_ErrorType, _
        Source:=CLASS_NAME & ": " & a_Source, _
        Description:=a_Description & "（クソが。）")
End Sub

