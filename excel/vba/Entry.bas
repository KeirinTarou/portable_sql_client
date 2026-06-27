Attribute VB_Name = "Entry"
Option Explicit

Public Sub ExecQuery( _
            ByVal a_FilePath As String, _
            ByVal a_Query As String, _
            ByVal a_Params As Variant, _
   Optional ByVal a_Timeout As Long = 30)
    Const adOpenStatic As String = 3
    Dim edb As New ExternalDB
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    On Error GoTo HandleError
    Set rs = _
        edb.ExecQuery( _
            edb.CONNECTION_STRING, _
            a_Query, _
            a_Params, _
            a_Timeout)
    Dim rjc As New RecordsetJsonConverter
    Call rjc.Init(rs)
    
    On Error GoTo HandleIOError
    Call rjc.Export(a_FilePath)
    Exit Sub
HandleError:
    Dim errDesc As String
    errDesc = Err.Description
    rs.CursorType = adOpenStatic
    Call rjc.Init(rs)
    Call rjc.Export(a_FilePath, errDesc)
    Exit Sub
HandleIOError:
    ' ファイルI/Oエラーは黙って握りつぶすしかない
    ' ワンチャン、このExcelにログを残すことを試みる
    ' この例外が出ている時点で、外部ファイルにエクスポートできる
    ' 保証がない -> だめもとでこのExcelに保存を試みる
    On Error Resume Next
    If Err.Number <> 0 Then
        Call Sh01Main.LogError(Err.Number, Err.Source, Err.Description)
        Call Err.Clear
    End If
    On Error GoTo 0
End Sub

