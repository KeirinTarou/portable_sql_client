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
    Call rjc.Export(a_FilePath)
    Exit Sub
HandleError:
    Dim errDesc As String
    errDesc = Err.Description
    rs.CursorType = adOpenStatic
    Call rjc.Init(rs)
    Call rjc.Export(a_FilePath, errDesc)
End Sub

