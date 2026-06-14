VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stringifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
'   Stringify Functions
'       Dependency:
'           - StringFormatter.cls
' =============================================================================

Private sf As StringFormatter

Private Sub Class_Initialize()
    Set sf = New StringFormatter
End Sub

' -----------------------------------------------------------------------------
'   ToString()系
'       - 直値:
'           - <String> "ち～ん（笑）"
'           - <Boolean> False
'           - <Long> 12345
'           - <Null> Null
'           - <Empty> Empty
'       - 配列・コレクション
'           - <Array> [<String> "pachinko", <Long> 123]
'           - <Collection> (<String> "pachinko", <Long> 123)
'       - ディクショナリ
'           - <Dictionary> {<String> "foo": {<String> "bar": <Long> 2000} }
'       - オブジェクト（ToString()が実装されているもの）
'           - <ClassName> ToString()の返り値
'           - ToString()メソッドには`MaxRecords`引数の設定が必要
'       - オブジェクト（値展開できないもの）
'           - <Object> CustomClass
' -----------------------------------------------------------------------------

' -----------------------------------------------------------------------------
'   司令塔
' -----------------------------------------------------------------------------
Public Function ToString( _
                 ByVal a_DataElement As Variant, _
        Optional ByVal a_ShowDetail As Boolean = True, _
        Optional ByVal a_PrettyPrint As Boolean = False, _
        Optional ByVal a_IndentLevel As Long = 0, _
        Optional ByVal a_MaxRecords As Long = 20) As String
    Const ERR_SOURCE As String = "ToString()"
    
    ' 数値指定引数のガード
    ' インデントレベルがマイナスになっていたら0に補正
    If a_IndentLevel < 0 Then a_IndentLevel = 0
    ' 頭打ち件数が0以下になっていたら1に補正
    If a_MaxRecords < 1 Then a_MaxRecords = 1
    
    Dim ret As String
    
    ' ObjectのNothing対策 -> 例外スロー
    If IsObject(a_DataElement) Then
        If a_DataElement Is Nothing Then _
            Call RaiseError(ERR_SOURCE, "Nothingを渡してはいけない。")
    End If
    
    ' Dictionaryの場合
    If TypeName(a_DataElement) = "Dictionary" Then _
        ret = DictionaryToString( _
            a_DataElement, a_MaxRecords, a_ShowDetail, a_PrettyPrint, a_IndentLevel): GoTo Finally
    ' Collectionの場合
    If TypeName(a_DataElement) = "Collection" Then _
        ret = CollectionToString( _
            a_DataElement, a_MaxRecords, a_ShowDetail, a_PrettyPrint, a_IndentLevel): GoTo Finally
    ' 上記以外のObjectの場合
    If IsObject(a_DataElement) Then _
        ret = ObjectToString(a_DataElement, a_MaxRecords, a_ShowDetail): GoTo Finally
    ' 配列の場合
    If IsArray(a_DataElement) Then _
        ret = ArrayToString( _
            a_DataElement, a_MaxRecords, a_ShowDetail, a_PrettyPrint, a_IndentLevel): GoTo Finally
    ' 上記以外
    ret = ValueToString(a_DataElement, a_ShowDetail)
Finally:
    ToString = ret
End Function

' -----------------------------------------------------------------------------
'   ディクショナリを文字列化する
' -----------------------------------------------------------------------------
Private Function DictionaryToString( _
            ByVal a_Dictionary As Variant, _
            ByVal a_MaxRecords As Long, _
   Optional ByVal a_ShowDetail As Boolean = True, _
   Optional ByVal a_PrettyPrint As Boolean = False, _
   Optional ByVal a_IndentLevel As Long = 0) As String
    Const ERR_SOURCE As String = "DictionaryToString()"
    Dim ret As String
    
    ' Dictionary以外 -> 例外スロー
    If TypeName(a_Dictionary) <> "Dictionary" Then _
        Call RaiseError(ERR_SOURCE, "Dictionary以外を渡してはいけない。")

    'ベースインデントレベルと整形フラグを取得
    Dim il As Long, PP As Boolean
    il = a_IndentLevel: PP = a_PrettyPrint
    
    ' 先頭の開きカッコ
    ret = "{"
    
    Dim tgtKey As Variant
    Dim k As String, v As String, cnt As Long
    cnt = 0
    Dim firstElem As Boolean
    firstElem = True
    For Each tgtKey In a_Dictionary.Keys
        ' 頭打ち設定用cntをインクリメント
        cnt = cnt + 1
        ' 処理レコード数がa_MaxRecordsを超えていたら、「... (+n more)」を追加してExit
        If cnt > a_MaxRecords And Not firstElem Then
            ret = ret & ", " & NEW_LINE(PP) & _
                  INDENT(PP, il + 1) & "... (" & "+" & CStr(a_Dictionary.Count - cnt + 1) & " more)"
            Exit For
        End If
        ' 先頭の要素
        If firstElem Then
            ret = ret & NEW_LINE(PP)        ' 整形ありのとき -> 改行
            firstElem = False
        ' 2番目以降の要素
        Else
            ret = ret & ", " & NEW_LINE(PP) ' 整形ありのとき -> 改行
        End If
        k = ToString(tgtKey, a_ShowDetail, PP, il + 1, a_MaxRecords)
        v = ToString(a_Dictionary(tgtKey), a_ShowDetail, PP, il + 1, a_MaxRecords)
        ' キーとアイテムを`: `でつなぐ
        ret = ret & INDENT(PP, il + 1) & k & ": " & v
    Next
    
    ' 要素が2個以上のときは閉じカッコの前に改行する
    '   -> 整形なしのときはNLには空文字が入っている
    If Not firstElem Then ret = ret & NEW_LINE(PP)
    ' インデントを1段戻す
    '   -> 整形なしのときは空文字
    ret = ret & INDENT(PP, il)
    
    ' 末尾の閉じカッコ
    ret = ret & "}"
    
Finally:
    If a_ShowDetail Then
        ret = "<Dictionary> " & ret
    End If
    DictionaryToString = ret
End Function

' -----------------------------------------------------------------------------
'   コレクションを文字列化する
' -----------------------------------------------------------------------------
Private Function CollectionToString( _
                    ByVal a_Collection As Variant, _
                    ByVal a_MaxRecords As Long, _
           Optional ByVal a_ShowDetail As Boolean = True, _
           Optional ByVal a_PrettyPrint As Boolean = False, _
           Optional ByVal a_IndentLevel As Long = 0) As String
    Const ERR_SOURCE As String = "CollectionToString()"
    Dim ret As String
    
    ' Collection以外 -> 例外スロー
    If TypeName(a_Collection) <> "Collection" Then _
        Call RaiseError(ERR_SOURCE, "Collection以外を渡してはいけない。")

    'ベースインデントレベルと整形フラグを取得
    Dim il As Long, PP As Boolean
    il = a_IndentLevel: PP = a_PrettyPrint
    
    ' 先頭の開きカッコ
    ret = "("
    
    ' Collectionオブジェクト`a_Item`の要素を1つずつ取り出して処理
    Dim itm As Variant, cnt As Long, tmp As String
    Dim firstElem As Boolean
    firstElem = True
    cnt = 0
    For Each itm In a_Collection
        ' 頭打ち件数カウント用変数をインクリメント
        cnt = cnt + 1
        ' 処理レコード数がa_MaxRecordsを超えていたら、
        '「... (+n more)」を追加してExit
        If cnt > a_MaxRecords And Not firstElem Then
            ret = ret & ", " & NEW_LINE(PP) & _
                  INDENT(PP, il + 1) & "... (" & "+" & _
                  CStr(a_Collection.Count - cnt + 1) & " more)"
            Exit For
        End If
        ' 先頭の要素のとき
        If firstElem Then
            ret = ret & NEW_LINE(PP)
            firstElem = False
        ' 2番目以降の要素
        Else
            ret = ret & ", " & NEW_LINE(PP)
        End If
        
        tmp = ToString(itm, a_ShowDetail, PP, il + 1)
        ' ここで末尾に追加
        ret = ret & INDENT(PP, il + 1) & tmp
    Next
    
    ' 要素が2個以上のときは閉じカッコの前で改行
    '   -> 整形なしのときはNLに""が入っているので改行なし
    If Not firstElem Then ret = ret & NEW_LINE(PP)
    ' 1段戻したインデントを追加
    '   -> 整形なしのときはPARENT_INDは空文字
    ret = ret & INDENT(PP, il)
    
    ' 末尾の閉じカッコ
    ret = ret & ")"
Finally:
    If a_ShowDetail Then
        ret = "<Collection> " & ret
    End If
    CollectionToString = ret
End Function

' -----------------------------------------------------------------------------
'   配列を文字列化する
' -----------------------------------------------------------------------------
Private Function ArrayToString( _
                    ByVal a_Item As Variant, _
                    ByVal a_MaxRecords As Long, _
           Optional ByVal a_ShowDetail As Boolean, _
           Optional ByVal a_PrettyPrint As Boolean = False, _
           Optional ByVal a_IndentLevel As Long = 0) As String
    Const ERR_SOURCE As String = "ArrayToString()"
    Dim ret As String
    
    ' Null, Empty禁止 -> 例外スロー
    If IsNull(a_Item) Then _
        Call RaiseError(ERR_SOURCE, "Nullが渡された。")
    If IsEmpty(a_Item) Then _
        Call RaiseError(ERR_SOURCE, "Emptyが渡された。")
    ' Object禁止
    If IsObject(a_Item) Then _
        Call RaiseError(ERR_SOURCE, "オブジェクトを渡してはいけない。")
    
    ' とにかく配列でない
    If Not IsArray(a_Item) Then _
        Call RaiseError(ERR_SOURCE, "配列でない。")
    
    ' 空の配列 -> <Array> []
    If IsEmptyArray(a_Item) Then
        ret = "[]"
        If a_ShowDetail Then ret = "<Array> " & ret
        GoTo Finally
    End If
    
    ' ここからが普通の配列への対応
    ' 次元数を調べる（1次元・2次元のみサポート）
    Dim dimCnt As Long
    dimCnt = ArrayDimension(a_Item)
    If dimCnt >= 3 Then _
        Call RaiseError(ERR_SOURCE, "3次元以上の配列はサポートしない。")
        
    ' ここで2次元以下の配列であることが保障された
    If dimCnt = 1 Then
        ret = ArrayToString1D(a_Item, a_MaxRecords, a_ShowDetail, a_PrettyPrint, a_IndentLevel)
    ElseIf dimCnt = 2 Then
        If Invalid2DimArray(a_Item) Then
            ret = "[Invalid 2D Array]"
            If a_ShowDetail Then ret = "<Array> " & ret
            GoTo Finally
        End If
        ret = ArrayToString2D(a_Item, a_MaxRecords, a_ShowDetail, a_PrettyPrint, a_IndentLevel)
    End If
    
Finally:
    ArrayToString = ret
End Function

' 2次元配列の検証 -> 異常な2次元配列はTrue
Private Function Invalid2DimArray(ByVal a_2DimArray As Variant) As Boolean
    On Error GoTo InvalidArray
    
    Dim ubR As Long, ubC As Long
    ubR = UBound(a_2DimArray, 1)
    ubC = UBound(a_2DimArray, 2)
    
    ' 添え字が異常 -> 即Trueを返す
    If ubR < LBound(a_2DimArray, 1) Then GoTo InvalidArray
    If ubC < LBound(a_2DimArray, 2) Then GoTo InvalidArray
    
    ' ここまで耐えたらFalse確定
    Invalid2DimArray = False
    Exit Function
    
InvalidArray:
    ' 例外が出ていたら握りつぶす
    If Err.Number <> 0 Then Call Err.Clear
    Invalid2DimArray = True
End Function

' 純粋に1次元配列を文字列化する
'   -> 1次元配列であることが保障された状態で呼び出す
Private Function ArrayToString1D( _
                    ByVal a_Array As Variant, _
                    ByVal a_MaxRecords As Long, _
           Optional ByVal a_ShowDetail As Boolean = True, _
           Optional ByVal a_PrettyPrint As Boolean = False, _
           Optional ByVal a_IndentLevel As Long = 0) As String
    
    Dim ret As String

    'ベースインデントレベルと整形フラグを取得
    Dim il As Long, PP As Boolean
    il = a_IndentLevel: PP = a_PrettyPrint
    
    ' 先頭の開きブラケット
    ret = "["
    
    ' 要素を順に文字列化
    Dim i As Long, ub As Long, lb As Long, tmp As String, idx As String
    ' 配列添え字の上限/下限をキャッシュ
    ub = UBound(a_Array): lb = LBound(a_Array)
    ' 頭打ち件数カウント用変数
    Dim cnt As Long
    cnt = 0
    For i = lb To ub
        ' 頭打ち件数カウント用変数をインクリメント
        cnt = cnt + 1
        ' 処理レコード数がa_MaxRecordsを超えていたら、「... (+n more)」を追加してExit
        If cnt > a_MaxRecords And i <> lb Then
            ' 整形フラグTrue -> 改行・さらに1段インデントを下げる
            ret = ret & ", " & NEW_LINE(PP) & _
                  INDENT(PP, il + 1) & "... (" & "+" & CStr((ub - lb + 1) - cnt + 1) & " more)"
            Exit For
        End If
        ' 先頭の要素のとき
        If i = lb Then
            ret = ret & NEW_LINE(PP)
        Else
            ret = ret & ", " & NEW_LINE(PP)
        End If
        ' 配列添え字を文字列化
        idx = CStr(i) & ": "
        tmp = ToString(a_Array(i), a_ShowDetail, PP, il + 1, a_MaxRecords)
        ret = ret & INDENT(PP, il + 1) & idx & tmp
    Next
    
    ' 要素が2個以上のときは閉じカッコの前で改行
    ' If Not lb = ub Then ret = ret & NL
    ret = ret & NEW_LINE(PP)
    
    ' 1段戻したインデントを追加
    ret = ret & INDENT(PP, il)
    
    ' 末尾の閉じブラケット
    ret = ret & "]"
    If a_ShowDetail Then ret = "<Array> " & ret
    
    ArrayToString1D = ret
End Function

' 純粋に2次元配列を文字列化する
Private Function ArrayToString2D( _
            ByVal a_Array As Variant, _
            ByVal a_MaxRecords As Long, _
   Optional ByVal a_ShowDetail As Boolean, _
   Optional ByVal a_PrettyPrint As Boolean = False, _
   Optional ByVal a_IndentLevel As Long = 0) As String
    Dim ret As String

    ' ベースインデントレベルと整形フラグを取得
    Dim il As Long, PP As Boolean
    il = a_IndentLevel: PP = a_PrettyPrint
    
    ' 先頭の開きブラケット
    ret = "["
    
    Dim r As Long, c As Long
    Dim lbR As Long, ubR As Long
    Dim lbC As Long, ubC As Long
    
    ' 行列方向の添え字上限/下限をキャッシュ
    lbR = LBound(a_Array, 1): ubR = UBound(a_Array, 1)
    lbC = LBound(a_Array, 2): ubC = UBound(a_Array, 2)
    
    ' 頭打ち件数カウント用変数
    Dim cntR As Long, cntC As Long
    cntR = 0: cntC = 0
        
    Dim tmp As String, idx As String
    ' スライスした配列ごとに文字列化
    For r = lbR To ubR
        ' 行方向の頭打ち件数カウント用変数インクリメント
        cntR = cntR + 1
        ' 処理レコード数がa_MaxRecordsを超えていたら、「... (+n more)」を追加してExit
        If cntR > a_MaxRecords And r <> lbR Then
            ret = ret & ", " & NEW_LINE(PP) & _
                  INDENT(PP, il) & "... (" & "+" & CStr((ubR - lbR + 1) - cntR + 1) & " more)"
            Exit For
        End If
        
        ' 先頭の行
        If r = lbR Then
            ret = ret & NEW_LINE(PP)
        Else
            ret = ret & ", " & NEW_LINE(PP)
        End If
        
        ' 行配列の処理
        ' 列方向の頭打ち件数カウント用変数初期化
        cntC = 0
        '   中の配列の開きカッコの処理
        ret = ret & INDENT(PP, il + 1) & IIf(a_ShowDetail, "<Array> ", "") & "["
        For c = lbC To ubC
            ' 行方向の頭打ち件数カウント用変数インクリメント
            cntC = cntC + 1
            ' 処理レコード数がa_MaxRecordsを超えていたら、
            '「... (+n more)」を追加してExit
            If cntC > a_MaxRecords And c <> lbC Then
                ret = ret & ", " & NEW_LINE(PP) & _
                    INDENT(PP, il + 2) & "... (" & "+" _
                    & CStr((ubC - lbC + 1) - cntC + 1) & " more)"
                Exit For
            End If
        
            ' 先頭の列
            If c = lbC Then
                ret = ret & NEW_LINE(PP)
            Else
                ret = ret & ", " & NEW_LINE(PP)
            End If
            
            ' ここで、当該の要素を文字列化
            tmp = ToString(a_Array(r, c), a_ShowDetail, PP, il + 1, a_MaxRecords)
            ' 2次元配列の添え字を作成
            idx = CStr(r) & ", " & CStr(c) & ": "
            ret = ret & INDENT(PP, il + 2) & idx & tmp
        Next
        ' 行を閉じる
        ret = ret & NEW_LINE(PP) & INDENT(PP, il + 1) & "]"
    Next
    
    ' 要素が2個以上のときは閉じカッコの前で改行
    If Not lbR = ubR Then ret = ret & NEW_LINE(PP)
    ' 1段戻したインデントを追加
    ret = ret & INDENT(PP, il)
    
    ' 最後の閉じブラケット
    ret = ret & "]"
    
    ' 型名の明示オプション対応
    If a_ShowDetail Then ret = "<Array> " & ret
    
    ArrayToString2D = ret
End Function

' 配列を行方向にスライス
Private Function SliceArrayR( _
            ByVal a_2DimArray As Variant, _
            ByVal a_SlicesAt As Long) As Variant
    Dim ret As Variant
    
    Dim lbC As Long, ubC As Long
    lbC = LBound(a_2DimArray, 2): ubC = UBound(a_2DimArray, 2)
    
    ReDim ret(lbC To ubC)
    Dim tmp As Variant
    Dim c As Long
    For c = lbC To ubC
        If IsObject(a_2DimArray(a_SlicesAt, c)) Then
            Set ret(c) = a_2DimArray(a_SlicesAt, c)
        Else
            ret(c) = a_2DimArray(a_SlicesAt, c)
        End If
    Next
    
    SliceArrayR = ret
End Function

' 配列の次元数取得
Private Function ArrayDimension(ByVal a_Array As Variant) As Long
    Const ERR_SOURCE As String = "ArrayDimension()"
    If Not IsArray(a_Array) Then _
        Call RaiseError(ERR_SOURCE, "配列でない。")
        
    Dim ret As Long
    On Error Resume Next
    
    Dim i As Long, lb As Long
    For i = 1 To 60
        lb = LBound(a_Array, i)
        If Err.Number <> 0 Then
            Call Err.Clear
            Exit For
        End If
        ret = i
    Next
    On Error GoTo 0

    ArrayDimension = ret
End Function

' 空配列判定
Private Function IsEmptyArray(ByVal a_Array As Variant) As Boolean
    Const ERR_SOURCE As String = "IsEmptyArray()"
    If Not IsArray(a_Array) Then _
        Call RaiseError(ERR_SOURCE, "配列でない。")
    ' デフォルトTrue -> 空配列確定即Return
    IsEmptyArray = True
    On Error GoTo HandleError
    Dim ub As Long
    ub = UBound(a_Array)
    
    If ub < 0 Then Exit Function
    ' ここまで来たらFalse
    IsEmptyArray = False
    Exit Function
HandleError:
    ' エラーを握りつぶしてTrueを返す
    Call Err.Clear
End Function

' -----------------------------------------------------------------------------
'   オブジェクトを文字列化する
' -----------------------------------------------------------------------------
Private Function ObjectToString( _
            ByVal a_Item As Variant, _
            ByVal a_MaxRecords As Long, _
   Optional ByVal a_ShowDetail As Boolean = True) As String
    Dim ret As String
    Const ERR_SOURCE As String = "ObjectToString()"
    
    ' Objectでない -> 例外スロー
    If Not IsObject(a_Item) Then _
        Call RaiseError(ERR_SOURCE, "オブジェクト以外を渡してはいけない。")
    
    ' TypeName取得
    Dim tn As String
    tn = TypeName(a_Item)
    ' コレクション、ディクショナリも禁止
    If tn = "Collection" Then _
        Call RaiseError(ERR_SOURCE, "Collectionを渡してはいけない。")
    If tn = "Dictionary" Then _
        Call RaiseError(ERR_SOURCE, "Dictionaryを渡してはいけない。")
    
    ' ToString()メソッドを持っていたらそれを返す
'    On Error Resume Next
    ret = a_Item.ToString(a_MaxRecords)
    If a_ShowDetail Then ret = "<" & tn & "> " & ret
    ' ToString()メソッドなし -> 例外を握りつぶして次へ
    If Err.Number <> 0 Then
        Call Err.Clear
    Else
        ' 例外なし: ToString()メソッドありオブジェクトの場合
        '   -> 必要な文字列セット済み -> return
        GoTo Finally
    End If
    On Error GoTo 0
        
    ret = tn
    
    ' 詳細情報が必要な場合はここで追加
    If a_ShowDetail Then ret = "<Object> " & ret
Finally:
    ObjectToString = ret
End Function

' -----------------------------------------------------------------------------
'   値型のアイテムを文字列化する
' -----------------------------------------------------------------------------
Private Function ValueToString( _
                ByVal a_Value As Variant, _
       Optional ByVal a_ShowDetail As Boolean = True) As String
    Const ERR_SOURCE As String = "ValueToString()"
    
    Dim ret As String

    ' オブジェクト・配列が渡されたら例外スロー
    If IsObject(a_Value) Then _
        Call RaiseError(ERR_SOURCE, "オブジェクトを渡してはいけない。")
    If IsArray(a_Value) Then _
        Call RaiseError(ERR_SOURCE, "配列を渡してはいけない。")
    
    ' NullとEmptyを先に片付ける
    If IsNull(a_Value) Then
        ret = IIf(a_ShowDetail, "<Null> Null", "Null")
        GoTo Finally
    End If
    If IsEmpty(a_Value) Then
        ret = IIf(a_ShowDetail, "<Empty> Empty", "Empty")
        GoTo Finally
    End If
    
    Select Case VarType(a_Value)
        Case vbString
            ret = """" & Replace(a_Value, """", """""") & """"
        Case vbBoolean
            ret = IIf(a_Value, "True", "False")
        Case Else
            ret = CStr(a_Value)
    End Select
    
    ' 型情報が不要ならこのままreturn
    If Not a_ShowDetail Then GoTo Finally
    
    ' 詳細表示が必要な場合は型情報を先頭に追加する
    ret = "<" & ToStringVarType(a_Value) & "> " & ret

Finally:
    ValueToString = ret
End Function

' -----------------------------------------------------------------------------
'   正確な型情報を返す（`TypeName()`の上位互換）
' -----------------------------------------------------------------------------
Private Function ToStringVarType(ByVal a_Item As Variant) As String
    Dim ret As String
    
    Dim vt As VbVarType
    vt = VarType(a_Item)
    
    Dim baseType As String

    ' 配列判定 -> 配列だったら基底型に`[]`を付けて返す
    If (vt And vbArray) <> 0 Then
        ' 配列要素の基本型を取得
        baseType = ToStringVarTypeBasic(vt And Not vbArray)
        ret = baseType & "[]"
        GoTo Finally
    End If
    
    ' 配列でない場合は型情報を返す
    ret = ToStringVarTypeBasic(vt)
        
Finally:
    ToStringVarType = ret
End Function

' -----------------------------------------------------------------------------
'   VarType の基本型部分を文字列化
' -----------------------------------------------------------------------------
Private Function ToStringVarTypeBasic(ByVal a_VarType As VbVarType) As String
    Dim ret As String
    Select Case a_VarType
        Case vbEmpty: ret = "Empty"
        Case vbNull: ret = "Null"
        Case vbByte: ret = "Byte"
        Case vbBoolean: ret = "Boolean"
        Case vbDate: ret = "Date"
        Case vbString: ret = "String"
        Case vbInteger: ret = "Integer"
        Case vbLong: ret = "Long"
        Case vbSingle: ret = "Single"
        Case vbDouble: ret = "Double"
        Case vbCurrency: ret = "Currency"
        Case vbObject: ret = "Object"
        Case vbVariant: ret = "Variant"
        Case vbUserDefinedType: ret = "UDT"
        Case Else
            ret = "Unknown(" & CStr(a_VarType) & ")"
    End Select
    ToStringVarTypeBasic = ret
End Function

' #############################################################################
'   Helper functions
' #############################################################################
Private Sub ZZ_Helpers()
End Sub

' インデント文字生成用関数
Private Function INDENT( _
            ByVal a_PrettyPrint As Boolean, _
            ByVal a_IndentLevel As Long) As String
    INDENT = ""
    ' 整形フラグ`False` -> 空文字
    If Not a_PrettyPrint Then Exit Function
    
    INDENT = sf.IND(a_IndentLevel)
End Function
' 改行文字生成用関数
Private Function NEW_LINE( _
            ByVal a_PrettyPrint As Boolean, _
   Optional ByVal a_Count As Long = 1) As String
    NEW_LINE = ""
    ' 整形フラグ`False` -> 空文字
    If Not a_PrettyPrint Then Exit Function
    
    NEW_LINE = sf.NL(a_Count)
End Function

' -----------------------------------------------------------------------------
'   テストコード生成用関数
'       - ToString用のテストコード生成
'       - PublicなのはToString()だけなので専用で良い
' -----------------------------------------------------------------------------
Public Function GenTestCodeScalar( _
                ByVal a_TestType As String, _
                ByVal a_ScalarValue As Variant, _
                ByVal a_ShowDetail As Boolean, _
                ByVal a_ExpectedResult As String, _
       Optional ByVal a_Comment As String) As String
    Dim ret As String
    
    Dim sd As String
    sd = IIf(a_ShowDetail, "True", "False")
    ret = "Private Sub test_ToString_" & a_TestType & "()" & vbNewLine
    ret = ret & sf.IND(1) & _
          "Dim v As Variant" & sf.NL(1)
    ret = ret & sf.IND(1) & "v = " & LIT(a_ScalarValue) & sf.NL(1)
    ret = ret & sf.IND(1) & "Call m_UnitTest.AssertEqual( _" & sf.NL(1)
    ret = ret & sf.IND(2) & "a_Actual:=ToString(v, " & sd & "), _" & sf.NL(1)
    ret = ret & sf.IND(2) & "a_Expected:=""" & a_ExpectedResult & """, _" & sf.NL(1)
    ret = ret & sf.IND(2) & "a_Comment:=" & """" & a_Comment & """)" & sf.NL(1)
    ret = ret & "End Sub"
    
    GenTestCodeScalar = ret
End Function

Private Function LIT(ByVal a_Value As Variant) As String
    Select Case VarType(a_Value)
        Case vbEmpty
            LIT = "Empty"
        Case vbString
            LIT = """" & Replace(a_Value, """", """""") & """"
        Case vbBoolean
            LIT = IIf(a_Value, "True", "False")
        Case vbNull
            LIT = "Null"
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency
            LIT = CStr(a_Value)
        Case vbDate
            Dim dtFmt As String
            dtFmt = "yyyy/mm/dd"
            If Len(a_Value) > 10 Then
                dtFmt = "yyyy/mm/dd hh:nn:ss"
            End If
            LIT = "#" & Format$(a_Value, dtFmt) & "#"
        Case Else
            LIT = "Nothing   'TODO: Unsupported literal"
    End Select
End Function

' -----------------------------------------------------------------------------
'   例外スロー用関数
' -----------------------------------------------------------------------------
Private Sub RaiseError( _
                ByVal a_Source As String, _
       Optional ByVal a_Description As String)
    Dim desc As String
    If a_Description = "" Then
        desc = "(　ﾟдﾟ)､ﾍﾟｯ < Objectをキーにするなど許さん！"
    Else
        desc = "(　ﾟдﾟ)､ﾍﾟｯ < " & a_Description
    End If
    
    Call Err.Raise( _
        Number:=vbObjectError + 12345, _
        Source:="DictUtil." & "DictUtil." & a_Source, _
        Description:=desc)
End Sub

