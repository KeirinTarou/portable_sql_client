VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordsetJSONConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
'   ADODB.RecordsetオブジェクトをJSON化する
'
'       Dependency
'           - StringFormatter.cls
'           - TextFile.cls
'               - TextLines.cls
' #############################################################################

Private sf As New StringFormatter

Private Enum ErrorTypeEnum
    etInvalidRecordset = 1
    etNotInitialized = 2
    etInvalidCursorType = 3
End Enum

Private Const adOpenStatic As Long = 3
Private Const adTypeText As Long = 2
Private Const adSaveCreateOverWrite As Long = 2
Private Const ERR_NUMBER_BASE As Long = vbObjectError + 5000

Private m_Recordset As Object

' =============================================================================
'   Constructor
' =============================================================================
Private Sub Class_Initialize()
    Set m_Recordset = Nothing
End Sub

Public Property Get Recordset() As Object
    Set Recordset = m_Recordset
End Property

Public Sub Init( _
            ByVal a_Recordset As Object)
    Const ERR_SOURCE As String = "Init()"
    If TypeName(a_Recordset) <> "Recordset" Then _
        Call RaiseError( _
            etInvalidRecordset, _
            ERR_SOURCE, _
            "引数は`ADODB.Recordset`オブジェクトでなければならない。")
    
    ' CursorTypeをadOpenStatic縛りにする
    '   - RecordCountプロパティ、MoveFirst()メソッド利用のため
    If a_Recordset.CursorType <> adOpenStatic Then _
        Call RaiseError( _
            etInvalidCursorType, _
            ERR_SOURCE, _
            "RecordsetのCursorTypeはadOpenStatic(=3)でなければならない。")
    
    Set m_Recordset = a_Recordset
End Sub

Public Sub Export( _
            ByVal a_FilePath As String)
    Const ERR_SOURCE As String = "Export()"
    ' m_RecordsetがNothing -> 例外スロー
    If m_Recordset Is Nothing Or TypeName(m_Recordset) <> "Recordset" Then _
        Call RaiseError( _
            etInvalidRecordset, _
            ERR_SOURCE, _
            "Recordsetオブジェクトがセットされていない。")
    
    Dim json As String
    json = BuildJSON()
    
    ' JSONファイルをエクスポート
    Dim strm As Object
    Set strm = CreateObject("ADODB.Stream")
    With strm
        .Type = adTypeText
        .Charset = "utf-8"
        Call .Open
    End With
    Call strm.WriteText(json)
    Call strm.SaveToFile(a_FilePath, adSaveCreateOverWrite)
    Call strm.Close
    Set strm = Nothing
End Sub

Private Sub AA_InternalMethods(): End Sub
' =============================================================================
'   Internal Methods
' =============================================================================
Private Function BuildJSON() As String
    Dim ret As String
    
    ret = BuildJSONCore(Me.ExtractColumns, Me.ExtractRows)
    
    BuildJSON = ret
End Function

Friend Function BuildJSONCore( _
            ByVal a_Columns As Variant, _
            ByVal a_Rows As Variant) As String
' -----------------------------------------------------------------------------
'   @params
'       - a_Columns: カラム名の1次元配列
'       - a_Rows: データ行の配列の配列
'           - 配列の各要素は〝1行分のレコードを表す1次元配列〟
'   【例】:
'       - a_Columns = Array("ID", "NAME")
'       - a_Rows = Array( _
'               Array(1, "Steve"), _
'               Array(2, Null))
'   @return
'       - {"columns": ["ID", "NAME"], "rows": [[1, "Steve"], [2, null]]}
' -----------------------------------------------------------------------------
    Dim ret As String
    ' 先頭の`"columns": `を作る
    ret = sf.Quoted("columns") & ": "
    ' 列名の配列をJSON化する
    Dim i As Long, ub1 As Long, tmp1 As Variant
    ub1 = UBound(a_Columns)
    For i = LBound(a_Columns) To ub1
        tmp1 = tmp1 & sf.Quoted(a_Columns(i))
        If i <> ub1 Then tmp1 = tmp1 & ", "
    Next
    ' 列名配列の要素をブラケットで囲んで`, `を追加
    ret = ret & sf.Bracketed(tmp1) & ", "
    'ここで`tmp1`は初期化
    tmp1 = ""
    
    ' 行データ部分の先頭`"rows": `を作る
    ret = ret & sf.Quoted("rows") & ": "
    ' 2重の配列をJSON文字列化する
    Dim j As Long, ub2 As Long, tmp2 As Variant
    ub1 = UBound(a_Rows)
    For i = LBound(a_Rows) To ub1
        ub2 = UBound(a_Rows(i))
        For j = LBound(a_Rows(i)) To ub2
            tmp2 = tmp2 & ToJSONLiteral(a_Rows(i)(j))
            If j <> ub2 Then tmp2 = tmp2 & ", "
        Next
        ' 内側の配列を全部列挙したらブラケットで囲む
        tmp2 = sf.Bracketed(tmp2)
        ' 外側の配列に追加
        tmp1 = tmp1 & tmp2
        If i <> ub1 Then tmp1 = tmp1 & ", "
        ' ここで次のループで使うために`tmp2`はクリア
        tmp2 = ""
    Next
    ' 外側のブラケットをかぶせて`..., "rows": `の後ろに連結
    ret = ret & sf.Bracketed(tmp1)
    ' 最後に全体をブレースで囲む
    ret = sf.Braced(ret)
    
    BuildJSONCore = ret
End Function

Friend Function ToJSONLiteral( _
            ByVal a_Value As Variant) As String
    Dim ret As String
    
    ' 念のためEmptyはNullにフォールバック
    If IsEmpty(a_Value) Then
        ret = "null"
        GoTo Finally
    End If
    
    ' NULL -> null
    If IsNull(a_Value) Then
        ret = "null"
        GoTo Finally
    End If
    
    ' 数値として評価できる -> そのまま文字列化
    If IsNumeric(a_Value) Then
        ret = CStr(a_Value)
        GoTo Finally
    End If
    
    ' それ以外 -> ダブルクォートで括る
    ret = EscapeJSONString(CStr(a_Value))
    ret = sf.DQ & ret & sf.DQ
    
Finally:
    ToJSONLiteral = ret
End Function

Friend Function EscapeJSONString( _
            ByVal a_Text As String) As String
    Dim ret As String
    
    ret = a_Text
    ret = Replace(ret, "\", "\\")
    ret = Replace(ret, """", "\""")
    ret = Replace(ret, vbTab, "\t")
    ret = Replace(ret, vbCrLf, "\r\n")
    ret = Replace(ret, vbCr, "\r")
    ret = Replace(ret, vbLf, "\n")
    
    EscapeJSONString = ret
End Function

Friend Function ExtractColumns() As Variant
    Dim ret As Variant
    ret = Array()
    
    Dim cnt As Long
    cnt = m_Recordset.Fields.Count
    If cnt < 1 Then GoTo Finally
    
    ReDim ret(0 To cnt - 1)
    Dim i As Long
    For i = 1 To cnt
        ret(i - 1) = m_Recordset.Fields(i - 1).Name
    Next
    
Finally:
    ExtractColumns = ret
End Function

Friend Function ExtractRows() As Variant
    Dim ret As Variant
    ret = Array()
    
    ' フィールドが1以上あることを保証
    Dim colCnt As Long
    colCnt = m_Recordset.Fields.Count
    If colCnt < 1 Then GoTo Finally
    
    ' 行が1行以上あることを保証
    Dim cnt As Long
    cnt = m_Recordset.RecordCount
    If cnt < 1 Then GoTo Finally
    
    ReDim ret(0 To cnt - 1)
    Dim r As Variant, n As Long
    ' 先頭行に移動
    Call m_Recordset.MoveFirst
    n = 0
    Do Until m_Recordset.EOF
        r = ExtractCurrentRow()
        ret(n) = r
        Call m_Recordset.MoveNext
        n = n + 1
    Loop
    
Finally:
    ExtractRows = ret
End Function

Private Function ExtractCurrentRow() As Variant
    Dim ret() As Variant
    Dim cnt As Long
    cnt = m_Recordset.Fields.Count
    ReDim ret(0 To cnt - 1)
    Dim i As Long, v As Variant
    For i = 0 To cnt - 1
        ret(i) = m_Recordset.Fields(i).Value
    Next
    ExtractCurrentRow = ret
End Function

Private Sub AA_HelperFunctions(): End Sub
' =============================================================================
'   Helper Functions
' =============================================================================
Private Sub RaiseError( _
            ByVal a_ErrorType As ErrorTypeEnum, _
            ByVal a_Source As String, _
            ByVal a_Description As String)
    Call Err.Raise( _
        Number:=ERR_NUMBER_BASE + a_ErrorType, _
        Source:="RecordsetJsonConverter: " & a_Source, _
        Description:=a_Description & "（クソが。）")
End Sub

