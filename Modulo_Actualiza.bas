Attribute VB_Name = "Modulo_Actualiza"
Global gConSql As New ADODB.Connection
Global gidesa As Integer
Global gsConexi�n As String


Sub main()
ActualizaArchivosSIAM
End Sub
Sub ActualizaArchivosSIAM()
Dim adors As New ADODB.Recordset, s As String, d As Date, su As String
Dim s1 As String
'On Error Resume Next
On Error GoTo Salir:
'If gidesa = 0 Then 'Producci�n
    gsConexi�n = "FILEDSN=c:\archivos de programa\sistema de informaci�n operativa\siam\siam.dsn;pwd=siam_desa"
'Else 'Desarrollo
'    gsConexi�n = "FILEDSN=c:\archivos de programa\sistema de informaci�n operativa\siam\siam.dsn;uid=siamdesa;pwd=siamdesa"
'End If

Call EstableceConexi�nServidor(gsConexi�n, gConSql)

adors.Open "select * from ACTUALIZAARCHIV", gConSql, adOpenStatic, adLockReadOnly
If Not adors.EOF Then
    s = adors!archivobat
    d = adors!fechaact
    su = adors!ubicaci�n
    s1 = Dir(su & s)
    If Len(s1) = 0 Or d >= Int(Date) Then
        FileCopy "\\10.33.1.51\sioactual\siam\" & s, su & s
    End If
    Shell su & s, vbHide
End If
Exit Sub
Salir:
MsgBox Err.Description, vbOKOnly, "error"
End Sub

Function EstableceConexi�nServidor(sConexi�n As String, conSQL As ADODB.Connection) As Boolean
Dim Y As Byte, adors As New ADODB.Recordset
On Error GoTo ErrorConexi�n:
With conSQL
    If .State > 0 Then .Close
    .CursorLocation = adUseClient           'La posici�n de un motor de cursores
    .CommandTimeout = 50
    '.Provider = "SQLOLEDB"
    .ConnectionString = sConexi�n
    .Open
End With
'conSQL.ConnectionTimeout = 0
'conSQL.Open sConexi�n
'conSQL.CommandTimeout = 500
'conSQL.Execute "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED"

EstableceConexi�nServidor = True

Exit Function
ErrorConexi�n:
Y = MsgBox("Error: " + Err.Description, vbAbortRetryIgnore + vbCritical, "Error durante la conexi�n (" + Str(Err.Number) + ")")
If Y = vbCancel Then
    Exit Function
ElseIf Y = vbRetry Then
    Resume
ElseIf Y = vbIgnore Then
    Resume Next
End If
End Function



