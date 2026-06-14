VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DEFAULT_INDENT_SIZE As Long = 4

Private RE As Object

Private Sub Class_Initialize()
    Set RE = CreateObject("VBScript.RegExp")
End Sub

Public Property Get DQ() As String
    ' Double Quote = Chr(34)
    DQ = Chr$(34)
End Property

' 文字列整形用
' 改行
Public Function NL(Optional ByVal a_Count As Long = 1) As String
    If a_Count <= 0 Then NL = "": Exit Function
    Dim ret As String
    Dim i As Long
    For i = 1 To a_Count
        ret = ret & vbCrLf
    Next
    NL = ret
End Function
' インデント（半角スペースによるインデント）
Public Function IND( _
   Optional ByVal a_Level As Long = 1, _
   Optional ByVal a_Size As Long = DEFAULT_INDENT_SIZE) As String
    If a_Level <= 0 Or a_Size <= 0 Then IND = "": Exit Function
    IND = String(a_Level * a_Size, " ")
End Function
' 半角スペース
Public Function SP( _
    Optional ByVal a_Count As Long = 1) As String
    If a_Count <= 0 Then SP = "": Exit Function
    SP = String(a_Count, " ")
End Function
' タブ
Public Function TB( _
   Optional ByVal a_Count As Long = 1) As String
    If a_Count <= 0 Then TB = "": Exit Function
    TB = String(a_Count, vbTab)
End Function
' ダブルクォーテーションで囲った文字列
Public Function Quoted(ByVal a_Text As String) As String
    Dim ret As String
    ret = Replace(a_Text, DQ, DQ & DQ)
    ret = DQ & ret & DQ
    Quoted = ret
End Function
' 丸括弧
Public Function Parened(ByVal a_Text As String) As String
    Dim ret As String
    ret = "(" & a_Text & ")"
    Parened = ret
End Function
' 弓括弧
Public Function Braced(ByVal a_Text As String) As String
    Dim ret As String
    ret = "{" & a_Text & "}"
    Braced = ret
End Function
' 角括弧
Public Function Bracketed(ByVal a_Text As String) As String
    Dim ret As String
    ret = "[" & a_Text & "]"
    Bracketed = ret
End Function

' 区切り線
Public Function BreakLine( _
   Optional ByVal a_Length As Long = 40, _
   Optional ByVal a_Char As String = "-") As String
    Dim c As String * 1
    c = a_Char
    BreakLine = String(a_Length, c)
End Function

' 文字列の左埋め桁揃え（半角文字限定）
Public Function PadLeft( _
            ByVal a_Text As String, _
            ByVal a_Width As Long, _
   Optional ByVal a_PadChar As String = " ") As String
    Const ERR_SOURCE As String = "PadLeft()"
    Dim ret As String
    If Not IsSingleByteString(a_Text) Then _
        Call RaiseError(ERR_SOURCE, "マルチバイト文字はサポートしない。")
    ' Padding文字を先頭の1文字に切り詰める
    Dim padChar As String * 1
    padChar = a_PadChar
    If Not IsSingleByteString(padChar) Then _
        Call RaiseError(ERR_SOURCE, "マルチバイト文字によるPaddingはサポートしない。")
    ' 指定した幅よりも文字列が長い -> そのままreturn
    If Len(a_Text) > a_Width Then ret = a_Text: GoTo Finally
    ' PadLeftを行う
    ret = String(a_Width - Len(a_Text), padChar) & a_Text
Finally:
    PadLeft = ret
End Function
' 文字列の右埋め桁揃え（半角文字限定）
Public Function PadRight( _
            ByVal a_Text As String, _
            ByVal a_Width As Long, _
   Optional ByVal a_PadChar As String = " ") As String
    Const ERR_SOURCE As String = "PadRight()"
    Dim ret As String
    If Not IsSingleByteString(a_Text) Then _
        Call RaiseError(ERR_SOURCE, "マルチバイト文字はサポートしない。")
    ' Padding文字を先頭の1文字に切り詰める
    Dim padChar As String * 1
    padChar = a_PadChar
    If Not IsSingleByteString(padChar) Then _
        Call RaiseError(ERR_SOURCE, "マルチバイト文字によるPaddingはサポートしない。")
    ' 指定した幅よりも文字列が長い -> そのままreturn
    If Len(a_Text) > a_Width Then ret = a_Text: GoTo Finally
    ' PadLeftを行う
    ret = a_Text & String(a_Width - Len(a_Text), padChar)
Finally:
    PadRight = ret
End Function

' 文字列から余分なホワイトスペースを削除する
Public Function NormalizeSpaces( _
            ByVal a_Text As String) As String
    Dim ret As String
    
    ' ホワイトスペースと全角スペースにマッチ
    RE.Pattern = "[\s　]+"
    RE.Global = True
    
    ' 1. 改行で分割
    Dim lines() As String
    ' vbCrLf -> vbLf
    ret = Replace(a_Text, vbCrLf, vbLf)
    ' vbCr -> vbLf
    ret = Replace(ret, vbCr, vbLf)
    ' これで改行文字がvbLfに統一されている
    lines = Split(ret, vbLf)
    
    ' 2. 1行ずつ処理
    Dim i As Long
    Dim tmp As String
    For i = LBound(lines) To UBound(lines)
        tmp = lines(i)
        ' NBSPを半角スペースに変換
        tmp = Replace(tmp, ChrW(&HA0), " ")
        ' ZWSPを半角スペースに変換
        tmp = Replace(tmp, ChrW(&H200B), " ")
        ' 連続するホワイトスペースを1つにする
        tmp = RE.Replace(tmp, " ")
        ' 3. 両端のホワイトスペースをトリム
        tmp = Trim$(tmp)
        lines(i) = tmp
    Next
    ' 3. vbCrLfで結合
    ret = Join(lines, vbCrLf)
    
    NormalizeSpaces = ret
End Function
            

' -----------------------------------------------------------------------------
'   Helper Functions
' -----------------------------------------------------------------------------
Private Function IsSingleByteString(ByVal a_String As String) As Boolean
    IsSingleByteString = False
    Dim i As Long, c As String * 1
    For i = 1 To Len(a_String)
        c = Mid$(a_String, i, 1)
        ' 1つでもASCII文字以外があったらFalse
        If Asc(c) < 0 Or Asc(c) > 127 Then Exit Function
    Next
    IsSingleByteString = True
End Function

Private Sub RaiseError( _
            ByVal a_Source As String, _
            ByVal a_Description As String)
    Call Err.Raise( _
        Number:=vbObjectError + 12345, _
        Source:="StringFormatter." & a_Source, _
        Description:="( ´,_ゝ`)ﾌﾟｯ < " & a_Description & "（クソが。）")
End Sub

