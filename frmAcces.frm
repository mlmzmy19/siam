VERSION 5.00
Begin VB.Form frmAcceso 
   Caption         =   "Acceso al sistema"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12150
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAcces.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmAcces.frx":08CA
   ScaleHeight     =   9135
   ScaleWidth      =   12150
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox ComboUsuarios 
      Height          =   315
      ItemData        =   "frmAcces.frx":1E412
      Left            =   4410
      List            =   "frmAcces.frx":1E414
      TabIndex        =   0
      Top             =   6165
      Width           =   3645
   End
   Begin VB.CommandButton cmdBotón 
      BackColor       =   &H00008000&
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   6840
      TabIndex        =   5
      Top             =   6750
      Width           =   1365
   End
   Begin VB.CommandButton cmdBotón 
      BackColor       =   &H00008000&
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   4
      Top             =   6720
      Width           =   1365
   End
   Begin VB.TextBox txtPassword 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   8340
      MaxLength       =   13
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   6150
      Width           =   1905
   End
   Begin VB.TextBox txtUsuario 
      Height          =   285
      Left            =   4410
      TabIndex        =   2
      Top             =   6165
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton cmdBotón 
      BackColor       =   &H00008000&
      Caption         =   "Ca&mbiar Contraseña"
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   8550
      TabIndex        =   1
      Top             =   6750
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Cuenta del operador del sistema:"
      Height          =   225
      Index           =   0
      Left            =   4410
      TabIndex        =   6
      Top             =   5895
      Width           =   2370
   End
   Begin VB.Label Label1 
      Caption         =   "Contraseña:"
      Height          =   225
      Index           =   2
      Left            =   8325
      TabIndex        =   7
      Top             =   5895
      Width           =   885
   End
End
Attribute VB_Name = "frmAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Y As Byte
Dim rsSQLdate As New ADODB.Recordset
Dim bCorrecta As Boolean
Dim sContraseña As String 'Contraseña esperada
Dim lID As Long 'ID del usuario según la cuenta capturada
Dim iKey As Integer 'Última tecla oprimida en Cuenta

Private Sub cmdBotón_Click(Index As Integer)
Dim i As Long, sAhora As String, y2 As Byte, s As String, archivos, ArchivoR, ArchivoW
Dim rsSQLUSR As New ADODB.Recordset, ss As String, adors As New ADODB.Recordset, yCambiaCon As Byte, i1 As Integer
'MsgBox "antes crear objeto archivo"
gsinTime = Timer
Set archivos = CreateObject("Scripting.fileSystemObject")
'MsgBox "despues crear objeto archivo"
If Index = 1 Then
    If gi = 0 Then
        End
    End If
    Unload Me
    Exit Sub
End If

If gi <> 28 Then 'No es de confirmación de contraseña
    If ComboUsuarios.ListIndex >= 0 Then
        lID = ComboUsuarios.ItemData(ComboUsuarios.ListIndex)
    Else
        lID = -1
    End If
End If



Set rsSQLUSR = Nothing
If gi = 28 And gi1 > 0 Then 'Confirmación de contraseña
    rsSQLUSR.Open "select us.*, (select count(*) from usuesp where tipo=3 and idusi=" & gil & ") as consulta from usuariossistema us where us.id=" & gi1, gConSql, adOpenStatic, adLockReadOnly
Else
    rsSQLUSR.Open "select us.*, (select count(*) from usuesp where tipo=3 and idusi=" & lID & ") as consulta from usuariossistema us where us.id=" & lID, gConSql, adOpenStatic, adLockReadOnly
End If


'Set rsSQLUSR = ObtenConsulta("select us.*,d.idtipexp from usuariossistema us,delegaciones d where us.id=" & lID & " and us.iddel=d.id(+)")
If rsSQLUSR!contraseña <> UCase(txtPassword.Text) Then
    bCorrecta = False
    If Index = 10 Then Exit Sub
    Call MsgBox("La contraseña es incorrecta", vbOKOnly, "")
    txtPassword.SetFocus
    Exit Sub
End If
'


gySoloConsulta = rsSQLUSR!consulta 'Indicador que el usuario solo tiene permisos de consulta

bCorrecta = True
If Index = 10 Then Exit Sub
'giResponsable = IIf(IsNull(rsSQLUSR!idres), -1, rsSQLUSR!idres)
giPermisos = 1 * IIf(rsSQLUSR!Registro, 2, 1) * IIf(rsSQLUSR!Análisis, 3, 1) * IIf(rsSQLUSR!Seguimiento, 5, 1) * IIf(rsSQLUSR!Reportes, 7, 1)  '* IIf(rsSQLUSR!administración, 11, 1)
If gi = 28 Then
    giUsuario = gi1
Else
    giUsuario = lID
End If
gyGrupoUsuario = rsSQLUSR!idgpo
If rsSQLUSR!responsable > 0 Then
    giResponsable = giUsuario
Else
    giResponsable = IIf(IsNull(rsSQLUSR!idres), -1, rsSQLUSR!idres)
End If
Y = 0
For i = 0 To rsSQLUSR.Fields.Count - 1
    If InStr("cuenta,contraseñaservidor,", LCase(rsSQLUSR.Fields(i).Name) + ",") > 0 Then Exit For
Next
If i < rsSQLUSR.Fields.Count Then
    If Len(gConSql.ConnectionString) > 0 Then
        s = LCase(gConSql.ConnectionString)
    Else
        s = LCase(gCadSQL)
    End If
    ss = s
Else
    rsSQLUSR.Close
End If

'Verifica si es Usu Esp
rsSQLUSR.Close
rsSQLUSR.Open "select count(*) from usuesp where idusi=" & lID & " and tipo=1 and baja=0", gConSql, adOpenStatic, adLockReadOnly

If rsSQLUSR(0) > 0 Then
    giUsuEsp = 1
End If



If rsSQLUSR.State Then rsSQLUSR.Close
rsSQLUSR.Open "select * from usuariossistema where id=" & giUsuario, gConSql, adOpenStatic, adLockReadOnly
If Not rsSQLUSR.EOF Then
    i = rsSQLUSR!idgpo
End If
If Index = 2 Or yCambiaCon = 200 Then
    If yCambiaCon = 200 Then
        frmNuevaContraseña.cmdBotón(1).Enabled = False
    End If
    frmNuevaContraseña.Show vbModal
    If yCambiaCon = 200 Then
        gConSql.Execute "update usuariossistema set fechaact=sysdate where id=" & giUsuario
    End If
End If

Set frmAcceso = Nothing
Unload Me
Exit Sub
End Sub

Private Sub Combousuarios_Click()
Dim adors As New ADODB.Recordset
cmdBotón(0).Enabled = ComboUsuarios.ListIndex >= 0 And Len(Trim(txtPassword)) > 0
If ComboUsuarios.ListIndex >= 0 Then
    If adors.State Then adors.Close
    adors.Open "select contraseña from usuariossistema where id=" & ComboUsuarios.ItemData(ComboUsuarios.ListIndex), gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        sContraseña = adors(0)
    End If
End If
If txtPassword.Visible Then txtPassword.SetFocus
End Sub

Private Sub Combousuarios_LostFocus()
If ComboUsuarios.ListIndex < 0 Then
    ComboUsuarios.ListIndex = BuscaCombo(ComboUsuarios, ComboUsuarios.Text, False, True)
    If ComboUsuarios.ListIndex < 0 Then ComboUsuarios.Text = ""
End If
End Sub

Private Sub Form_Load()
Dim rsSQLPASS As New ADODB.Recordset

If rsSQLdate.State > 0 Then rsSQLdate.Close
If gSQLACC = cyOracle Then
    rsSQLdate.Open "Select sysdate as fecha from dual", gConSql, adOpenStatic, adLockReadOnly
    'Set rsSQLdate = "Select sysdate as fecha from dual",gconsql,
Else
    rsSQLdate.Open "Select getdate() as date", gConSql, adOpenStatic, adLockReadOnly
    'Set rsSQLdate = ObtenConsulta("Select getdate() as date")
End If

LlenaCombo ComboUsuarios, "select id, descripción from usuariossistema where id<>999 and baja=0 and id>1 order by descripción", "", True
    
Me.Caption = "Acceso al Sistema de administración de Multas"
'Label5(2).Caption = "Control de acceso al " + Trim(gsSis)
If gi = 28 Then
    sContraseña = gs
    ComboUsuarios.ListIndex = BuscaCombo(ComboUsuarios, giUsuario, True)
End If
If giPermisos = 0 Then Y = 28
On Error GoTo salir:
ShockwaveFlash1.Movie = CurDir + "\LogoCondusef.swf"
Exit Sub
salir:
iValor = -1
End Sub

Private Sub txtCampo_Change()
cmdBotón(0).Enabled = ComboUsuarios.ListIndex >= 0 And Len(txtCampo) = 4
End Sub

Private Sub txtCampo_KeyPress(KeyAscii As Integer)
End Sub

Private Sub ShockwaveFlash1_OnReadyStateChange(newState As Long)

End Sub

Private Sub Label2_Click()

End Sub

Private Sub txtPassword_Change()
bCorrecta = (sContraseña = UCase(txtPassword.Text))
cmdBotón(0).Enabled = bCorrecta
cmdBotón(2).Enabled = bCorrecta
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdBotón_Click (0)
End Sub

Function ModificaDSNSIOEXE(sArchivo As String, sUID As String, sPWD As String, sWSID As String) As Boolean
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
Dim ArchivoR, ArchivoW, FArchivo, tsW, tsR
Dim s As String, Y As Byte, ss As String, d As Date, i As Integer, iTabla As Integer
d = CDate("01/01/1900")
On Error GoTo ErrorLectura:
If gbSIO_mdb Then
    For i = 0 To gdbmidb.TableDefs.Count - 1
        If gdbmidb.TableDefs(i).LastUpdated > d Then
            d = gdbmidb.TableDefs(i).LastUpdated
            iTabla = i
        End If
    Next
    ss = gdbmidb.TableDefs(iTabla).Name & "(" & d & ")"
End If

Set FArchivo = CreateObject("Scripting.fileSystemObject")
Set ArchivoR = FArchivo.GetFile(sArchivo) 'Obtiene referencia de archivo
Set tsR = ArchivoR.OpenAsTextStream(ForReading, TristateUseDefault) 'abre archivo

s = Mid(sArchivo, 1, InStrRev(sArchivo, "\")) + "dsnsioexetemp.dsn"
FArchivo.CreateTextFile s 'Crear un archivo
Set ArchivoW = FArchivo.GetFile(s)    'Obtiene referencia de archivo
Set tsW = ArchivoW.OpenAsTextStream(ForWriting, TristateUseDefault) 'abre archivo para escritura
s = tsR.ReadLine
Y = 1
Do While Len(s) > 0
    If InStr(LCase(s), "uid=") Then
        tsW.Write "UID=" + sUID + Chr(13) & Chr(10)
        Y = Y * 2
    ElseIf InStr(LCase(s), "pwd=") Then
        tsW.Write "PWD=" + sPWD + Chr(13) & Chr(10)
        Y = Y * 3
    ElseIf InStr(LCase(s), "wsid=") Then
        tsW.Write "WSID=" + sWSID + Chr(13) & Chr(10)
        Y = Y * 5
    ElseIf InStr(LCase(s), "app=") Then
        tsW.Write "APP=SIO" + gsVersión + IIf(Len(ss), " Calo: " + ss, "") + Chr(13) & Chr(10)
        Y = Y * 7
    Else
        tsW.Write s + Chr(13) & Chr(10)
    End If
    'If ts2 Then Exit Do
    s = tsR.ReadLine
Loop
aborta:
If Y Mod 2 <> 0 Then
    tsW.Write "UID=" + sUID + Chr(13) & Chr(10)
End If
If Y Mod 5 <> 0 Then
    tsW.Write "WSID=" + sWSID + Chr(13) & Chr(10)
End If
If Y Mod 7 <> 0 Then
    tsW.Write "APP=SIO" + gsVersión + IIf(Len(ss), " Calo: " + ss, "") + Chr(13) & Chr(10)
End If
tsR.Close
tsW.Close
If Len(Trim(Dir(sArchivo))) > 0 Then Kill sArchivo
s = Mid(sArchivo, 1, InStrRev(sArchivo, "\")) + "dsnsioexetemp.dsn"
FileCopy s, sArchivo
If Len(Trim(Dir(s))) > 0 Then Kill s

ModificaDSNSIOEXE = True
Exit Function
ErrorLectura:
If Err.Number = 62 Then
    GoTo aborta:
End If
Resume
End Function

Private Sub txtUsuario_GotFocus()
lID = 0
End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If InStr(",40,38,", "," & KeyCode & ",") > 0 Then
    iKey = KeyCode
    ValidaCuenta
Else
    iKey = 0
End If
End Sub

Private Sub txtUsuario_LostFocus()
If lID <= 0 Then ValidaCuenta
End Sub

Sub ValidaCuenta()
Dim adors As ADODB.Recordset
If Len(txtUsuario.Text) = 0 Then Exit Sub
Set adors = New ADODB.Recordset
If iKey = 38 Then
    adors.Open "Select id,descripción,contraseña from (select * from usuariossistema order by Replace(Replace(Replace(Replace(Replace(lower(descripción),'á','a'),'é','e'),'í','i'),'ó','o'),'ú','u') desc) us where Replace(Replace(Replace(Replace(Replace(lower(descripción),'á','a'),'é','e'),'í','i'),'ó','o'),'ú','u')<'" & Replace(Replace(Replace(Replace(Replace(LCase(txtUsuario.Text), "á", "a"), "é", "e"), "í", "i"), "ó", "o"), "ú", "u") & "' and (baja is null or baja=0) and rownum<2", gConSql, adOpenStatic, adLockReadOnly
    'Set adors = ObtenConsulta("Select id,descripción,contraseña from (select * from usuariossistema order by Replace(Replace(Replace(Replace(Replace(lower(descripción),'á','a'),'é','e'),'í','i'),'ó','o'),'ú','u') desc) us where Replace(Replace(Replace(Replace(Replace(lower(descripción),'á','a'),'é','e'),'í','i'),'ó','o'),'ú','u')<'" & Replace(Replace(Replace(Replace(Replace(LCase(txtUsuario.Text), "á", "a"), "é", "e"), "í", "i"), "ó", "o"), "ú", "u") & "' and (baja is null or baja=0) and rownum<2")
Else
    adors.Open "Select id,descripción,contraseña from (select * from usuariossistema order by Replace(Replace(Replace(Replace(Replace(lower(descripción),'á','a'),'é','e'),'í','i'),'ó','o'),'ú','u')) us where Replace(Replace(Replace(Replace(Replace(lower(descripción),'á','a'),'é','e'),'í','i'),'ó','o'),'ú','u')" & IIf(iKey = 0, ">=", IIf(iKey = 40, ">", "<")) & "'" & Replace(Replace(Replace(Replace(Replace(LCase(txtUsuario.Text), "á", "a"), "é", "e"), "í", "i"), "ó", "o"), "ú", "u") & "' and (baja is null or baja=0) and rownum<2", gConSql, adOpenStatic, adLockReadOnly
    'Set adors = ObtenConsulta("Select id,descripción,contraseña from (select * from usuariossistema order by Replace(Replace(Replace(Replace(Replace(lower(descripción),'á','a'),'é','e'),'í','i'),'ó','o'),'ú','u')) us where Replace(Replace(Replace(Replace(Replace(lower(descripción),'á','a'),'é','e'),'í','i'),'ó','o'),'ú','u')" & IIf(iKey = 0, ">=", IIf(iKey = 40, ">", "<")) & "'" & Replace(Replace(Replace(Replace(Replace(LCase(txtUsuario.Text), "á", "a"), "é", "e"), "í", "i"), "ó", "o"), "ú", "u") & "' and (baja is null or baja=0) and rownum<2")
End If
If adors.EOF Then
    MsgBox "No se encuentra cuenta de usuario"
Else
    txtUsuario.Text = adors(1)
    lID = adors(0)
    sContraseña = adors(2)
End If
End Sub


