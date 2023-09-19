Attribute VB_Name = "Comunes"
Option Explicit
'Ult. Actualizaci�n: 12 dic 09

Const conChunkSize As Integer = 16384

Global gsConexi�n As String  'Contiene la cadena de conexi�n de la base de datos de access
'''Global gs_usuario As Integer  'ID del usuario delsistema
Global gConSql As New ADODB.Connection  'Conexi�n a DB en SQL 2000 exclusivamente Consultas
Global gConSqlTrans As New ADODB.Connection  'Conexi�n a DB en sql 2000 exclusivamente transacciones

Public gsDirReportes  'Directorio re reportes (*.rpt)
Public gsFormatoFechaHora 'Formato de fecha
Public gsFormatoFecha   'Formato hora
Public gsVersi�n As String  'Versi�n del sistema
Public gsSeparador As String  'Caracteres especiales para sustituci�n
Public gsPermisos As String  'Indicadores de los permisos a cada m�dulo
Public gsPermisosRep As String  'Indicadores de los permisos a cada m�dulo
'''Public gn_Miperfil As Byte  'Indica que perfil de permisos tiene el usuario del sistema

Public gs As String
Public gs1 As String  'Variable de uso m�ltiplo
Public gs2 As String  'Variable de uso m�ltiplo
Public gs3 As String  'Variable de uso m�ltiplo
Public gs4 As String  'Variable de uso m�ltiplo
Public gi As Long
Public gi1 As Long  'Variable de uso m�ltiplo
Public gi2 As Long  'Variable de uso m�ltiplo
Public gi3 As Long  'Variable de uso m�ltiplo
Public gi4 As Long  'Variable de uso m�ltiplo
Public gidesa As Long  'Indicador que se encuentra en DESARROLLO

Public gsWWW As String  'URL de la p�gina WEB que desea verse
Public giResponsable As Integer 'id del responsable = id Usuariossistema que tiene la propiedad de responsable
Public giUsuario As Integer 'id del usuario del sistema
Public giUsuEsp As Integer 'indica si el usuario es especial para poder notificar por todas las formas no solo SINE
Public gsDirDocumentos As String 'Guarda la ruta donde se encuentran los documentos que se emiten
Public glProceso As Long
Public gySoloConsulta As Byte
Public gdVersi�n As Date 'Fecha de la versi�n del Ejecutable
Public MDI As MDI_Prin


Sub Main()
Dim adors As New ADODB.Recordset
'gdVersi�n = "17/feb/2020"
'gdVersi�n = "09/dic/2020"
'gdVersi�n = "25/mar/2021"
'gdVersi�n = "03/may/2022"
'gdVersi�n = "09/dic/2022" 'Ajuste consulta de ins activas en SINE, Despliegue de SINE en CHROME
'gdVersi�n = "03/ago/2023" 'Ajuste Informes Consulta Tabular se agreg� filtro por IF, Se agreg� informe de NO Recibidos, programaci�n de todos los reportes por UNA IF procesados al d�a siguiente
gdVersi�n = "03/ago/2023" 'SE AGREG� BOTONES DE ACTVIDAD ANTERIOR Y SIGUIENTE

gidesa = 0 '0:Producci�n, 1:Desarrollo

If gidesa = 0 Then 'Producci�n
    gsConexi�n = "DSN=SIAM;pwd=siam_desa;Provider=SQLOLEDB"
    
    'gsConexi�n = "DSN=SIAM2;DRIVER=Oracle in ORACLI;PWD=siam_desa"
    
    gsConexi�n = "FILEDSN=" & App.Path & "\siam.dsn;pwd=siam_desa;"
    
    'gsConexi�n = "FILEDSN=c:\siam.dsn;pwd=siam_desa;"
    
    'MsgBox (gsConexi�n)
    
Else 'Desarrollo
    gsConexi�n = "FILEDSN=" & App.Path & "\siamdesa.dsn;pwd=SIAMTEST;"
End If
    

Call EstableceConexi�nServidor(gsConexi�n, gConSql)
If adors.State Then adors.Close
adors.Open "select * from generales", gConSql, adOpenStatic, adLockReadOnly



'directorio de documentos y reportes
gsDirDocumentos = App.Path & "\" & adors!dirdocumentos
gsDirReportes = App.Path & "\" & adors!dirreportes
adors.Close

If adors.State Then adors.Close
adors.Open "select f_siamexe_version from dual", gConSql, adOpenStatic, adLockReadOnly
gsVersi�n = "1.06" & " (" & Format(gdVersi�n, "dd/mm/yyyy") & ")"
If adors(0) <> gdVersi�n Then
    Call MsgBox("La versi�n del executable es diferente a la especificada en base de datos. Aunque puede trabajar con esta versi�n se sugiere contactar a los responsables del SIAM para su actualizaci�n", vbOKOnly + vbInformation, "")
End If

frmAcceso.Show vbModal
If gs <> "ADIOS" Then
    'frmAcceso.Show vbModal
    Set frmAcceso = Nothing
Else
    End
End If
Set MDI = New MDI_Prin

MDI.Show

End Sub

Function EstableceConexi�nServidor(sConexi�n As String, conSQL As ADODB.Connection) As Boolean
Dim Y As Byte, adors As New ADODB.Recordset
On Error GoTo ErrorConexi�n:
With conSQL
    If .State > 0 Then .Close
    .CursorLocation = adUseClient           'La posici�n de un motor de cursores
    .CommandTimeout = 50
    '.Provider = "SQLOLEDB"
    'MsgBox (.Provider)
    
    '.Provider = "ORACLI"
    
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


Function ValidaAccesoConTra(sCon As String) As Boolean
Dim s As String, ID As Integer, iM As Integer
ID = Day(Date)
iM = Month(Date)
If ID Mod 2 = 0 Then
    If iM Mod 2 = 0 Then
        s = Chr(48 + (Asc("M") - 48 + ID) Mod 74) & Chr(48 + ID Mod 74) & Chr(48 + (Asc("K") - 48 + iM) Mod 74) & Chr(48 + iM Mod 74)
    Else
        s = Chr(48 + ID Mod 74) & Chr(48 + (Asc("I") - 48 + ID) Mod 74) & Chr(48 + iM Mod 74) & Chr(48 + (Asc("E") - 48 + iM) Mod 74)
    End If
Else
    If iM Mod 2 = 0 Then
        s = Chr(48 + (Asc("m") - 48 + ID) Mod 74) & Chr(48 + ID Mod 74) & Chr(48 + (Asc("k") - 48 + iM) Mod 74) & Chr(48 + iM Mod 74)
    Else
        s = Chr(48 + ID Mod 74) & Chr(48 + (Asc("i") - 48 + ID) Mod 74) & Chr(48 + iM Mod 74) & Chr(48 + (Asc("e") - 48 + iM) Mod 74)
    End If
End If
ValidaAccesoConTra = (s = sCon)
End Function

Sub MoverCursor(ByRef frm As Form, sMovimiento As String, ByRef adors As ADODB.Recordset, Optional i_ira As Long)
Dim yError As Integer
Dim l As Long
On Error GoTo ErrorMovimiento:
Select Case sMovimiento
Case "Primero"
    If Not adors.EOF Then adors.MoveFirst
Case "Anterior"
    If adors.Bookmark > 1 Then adors.MovePrevious
Case "Siguiente"
    If adors.Bookmark < adors.RecordCount And adors.Bookmark > 0 Then adors.MoveNext
Case "�ltimo"
    If adors.RecordCount > 0 Then adors.MoveLast
Case "Ir_a"
    If i_ira <= adors.RecordCount Then
        adors.AbsolutePosition = i_ira
    Else
        frm.txtNoReg = adors.Bookmark & " / " & adors.RecordCount
    End If
Case "Deshacer"
    If Not adors.EOF Then
        l = adors.Bookmark
        adors.MoveFirst
        adors.Bookmark = l
    End If
End Select
frm.txtNoReg = adors.Bookmark & " / " & adors.RecordCount
Exit Sub
ErrorMovimiento:
'If Err.Number < 10000 Then
'    Resume Next
'End If
yError = MsgBox(Err.Description, vbAbortRetryIgnore + vbQuestion + vbDefaultButton2, "")
If yError = vbRetry Then
    Resume
ElseIf yError = vbIgnore Then
    Resume Next
End If
End Sub


Function F_Transacci�n(sTransacci�n As String, Optional bNoMensaje As Boolean) As Boolean
Dim conSQL As New ADODB.Connection, iRows As Integer
Dim sError As String, yError As Integer
On Error GoTo ErrorTrnas:
Call EstableceConexi�nServidor(gsConexi�n, conSQL)
conSQL.Execute sTransacci�n, iRows
If iRows > 0 Then
    If Not bNoMensaje Then
        Call MsgBox("Se afectaron " & iRows & " registro(s)", vbOKOnly + vbInformation, "")
    End If
    F_Transacci�n = True
ElseIf Not bNoMensaje Then
    If conSQL.Errors.Count > 0 Then
        Call MsgBox("No se efectu� la operaci�n: " & conSQL.Errors(0).Description, vbOKOnly + vbInformation, "")
    Else
        Call MsgBox("No se efectu� la operaci�n: " & Err.Description, vbOKOnly + vbInformation, "")
    End If
End If
conSQL.Close
Set conSQL = Nothing
Exit Function
ErrorTrnas:
    sError = "Error: " + Err.Description
    If Err.Number = 3260 Then
        yError = MsgBox(sError, vbRetryCancel + vbCritical, "No se guard� la informaci�n, error (" + Str(Err.Number) + ")")
    Else
        If Err.Number = -2147217873 Then
            yError = MsgBox("Este concepto no puede ser borrado ya que ha sido utilizado y hasta que no exista ninguna relaci�n en las tablas hijas podr� ser eliminado", vbAbortRetryIgnore + vbCritical, "Error no esperado (" + Str(Err.Number) + ")")
        Else
            yError = MsgBox(sError, vbAbortRetryIgnore + vbCritical, "Error no esperado (" + Str(Err.Number) + ")")
        End If
    End If


If yError = vbRetry Then
    Resume
ElseIf yError = vbIgnore Then
    Resume Next
End If

End Function

Sub ActualizaBotones(ByRef frm As Form, yBot�n As Byte, Optional yPermiso As Byte)
Dim yPer As Byte
If yPermiso > 0 Then
    yPer = yPermiso
Else
    yPer = 2
End If
With frm
'1:Nuevo
'2:Limpiar
'3:Buscar
'4:Guardar
'5:Deshacer
'6:Eliminar
'11:Inicio
'12:Anterior
'25:Siguiente
'26:�ltimo
Select Case yBot�n
Case 1  'Normal como cuando inician o nuevo
    'txtCampo(0).Locked = False
    .Toolbar.Buttons(1).Enabled = True '(yPer = 2 Or yPer = 5 Or yPer = 7 Or yPer = 8)
    .Toolbar.Buttons(2).Enabled = True
    .Toolbar.Buttons(3).Enabled = False
    .Toolbar.Buttons(4).Enabled = True
    .Toolbar.Buttons(5).Enabled = True
    .Toolbar.Buttons(6).Enabled = False
    .Toolbar.Buttons(11).Enabled = False
    .Toolbar.Buttons(12).Enabled = False
    .Toolbar.Buttons(25).Enabled = False
    .Toolbar.Buttons(26).Enabled = False
    .txtNoReg = "Nuevo"
    .txtNoReg.Enabled = False
Case 2  'despu�s de limpiar an�logo a nuevo
    'txtCampo(0).Locked = False
    .Toolbar.Buttons(1).Enabled = True  '(yPer = 2 Or yPer = 5 Or yPer = 7 Or yPer = 8)
    .Toolbar.Buttons(2).Enabled = True
    .Toolbar.Buttons(3).Enabled = True
    .Toolbar.Buttons(4).Enabled = False
    .Toolbar.Buttons(5).Enabled = False
    .Toolbar.Buttons(6).Enabled = False
    .Toolbar.Buttons(11).Enabled = False
    .Toolbar.Buttons(12).Enabled = False
    .Toolbar.Buttons(25).Enabled = False
    .Toolbar.Buttons(26).Enabled = False
    .txtNoReg = "Buscar"
    .txtNoReg.Enabled = False
Case 3  'despu�s de buscar
    .Toolbar.Buttons(1).Enabled = True  '(yPer = 2 Or yPer = 5 Or yPer = 7 Or yPer = 8)
    .Toolbar.Buttons(2).Enabled = True
    .Toolbar.Buttons(3).Enabled = False '(yPer = 2 Or yPer = 3 Or yPer = 6 Or yPer = 7)
    .Toolbar.Buttons(4).Enabled = True ' (yPer = 2 Or yPer = 4 Or yPer = 6 Or yPer = 8)
    .Toolbar.Buttons(5).Enabled = True
    .Toolbar.Buttons(6).Enabled = True
    .Toolbar.Buttons(11).Enabled = True
    .Toolbar.Buttons(12).Enabled = True
    .Toolbar.Buttons(25).Enabled = True
    .Toolbar.Buttons(26).Enabled = True
    .txtNoReg.Enabled = True
Case Else
End Select
End With
End Sub


'Obtiene los pendientes de la persona: lpersona
'yTipo: indica si devuelve la descripci�n (0) de los pendientes o solo los ids (1)
Function Pendientes(lPersona As Long, yTipo As Byte, ByRef bMod As Boolean, Optional ByRef lAnt As Long) As String
Dim adors As New ADODB.Recordset, s As String, lmax As Long
'busca actividades pendientes
adors.Open "select count(*) from t_rhpersonalseg where n_cvepersona=" & lPersona & " and n_cveperseg is not null and f_fecha is null", gConSql, adOpenStatic, adLockReadOnly
lAnt = 0
If adors(0) > 0 Then  'Existen acts programadas
    adors.Close
    adors.Open "select s.s_seguimiento,ps.f_programado,ps.n_cveperseg,ps.n_cveant,ltrim(p.s_nombre+' ')+ltrim(p.s_paterno+' ')+p.s_materno as responsable from t_rhpersonalseg ps left join c_rhperseg s on ps.n_cveperseg=s.n_cveperseg left join t_rhpersonal p on ps.n_cveresprog=p.n_cvepersona where ps.n_cvepersona=" & lPersona & " and ps.n_cveperseg is not null and ps.f_fecha is null", gConSql, adOpenStatic, adLockReadOnly
    Do While Not adors.EOF
        If yTipo = 0 Then
            s = s & IIf(IsNull(adors(0)), "", adors(0)) & IIf(IsNull(adors(1)), "", ", el " & Format(adors(1), gsFormatoFechaHora)) & " (" & adors(4) & ");"
        Else
            s = s & adors(2) & ","
        End If
        If lAnt = 0 Then lAnt = adors(3)
        adors.MoveNext
    Loop
    If Len(s) > 0 Then s = Mid(s, 1, Len(s) - 1)
    Pendientes = s
    bMod = True
    Exit Function
End If
If adors.State > 0 Then adors.Close
adors.Open "select max(n_cveseguimiento) from t_rhpersonalseg where n_cvepersona=" & lPersona, gConSql, adOpenStatic, adLockReadOnly
If IsNull(adors(0)) Then
    'No hay actividades ... busca comienzo
    adors.Close
    adors.Open "select s.s_seguimiento,s.n_cveperseg from c_rharcosperseg ar left join c_rhperseg s on ar.n_cvedestino=s.n_cveperseg where ar.n_cveorigen=98", gConSql, adOpenStatic, adLockReadOnly
    Do While Not adors.EOF
        If yTipo = 0 Then
            s = s & IIf(IsNull(adors(0)), "", adors(0)) & ";"
        Else
            s = s & adors(1) & ","
        End If
        adors.MoveNext
    Loop
    If Len(s) > 0 Then
        s = Mid(s, 1, Len(s) - 1)
        bMod = True
    Else
        If yTipo = 0 Then
            s = "Sin pendientes"
        Else
            s = ""
        End If
        bMod = False
    End If
    Pendientes = s
ElseIf Not IsNull(adors(0)) Then
    'Existe �ltima actividad
    lmax = adors(0)
    If adors.State > 0 Then adors.Close
    adors.Open "select s.s_seguimiento,ar.n_cvedestino from c_rharcosperseg ar left join c_rhperseg s on ar.n_cvedestino=s.n_cveperseg where ar.n_cveorigen*100+ar.n_cveevento*10+n_resultado in (select n_cveperseg*100+n_cveevento*10+n_resultado from t_rhpersonalseg where n_cveseguimiento=" & lmax & ") order by n_orden", gConSql, adOpenStatic, adLockReadOnly
    If adors.EOF Then
        'No hay actividades subsecuentes...Se considera concluido
        If yTipo = 0 Then
            s = "Concluido (Sin pendientes)"
        Else
            s = ""
        End If
        bMod = False
    Else
        If yTipo = 0 Then
            s = adors(0)
        Else
            s = IIf(adors(1) <> 99, adors(1), "")
        End If
        If lAnt = 0 Then lAnt = lmax
        bMod = (adors(1) <> 99)
    End If
End If
'If Len(s) > 0 Then s = Mid(s, 1, Len(s) - 1)
Pendientes = s
Exit Function
End Function

'Obtiene los iDs de kPerSeg siguientes seg�n iPerseg (Evento) y iRes (Resultado)
Function IDPerSeg_Sig(iPerseg As Integer, iEvento As Integer, iRes As Byte) As String
Dim adors As New ADODB.Recordset, s As String
If iPerseg = 0 Then
    adors.Open "select n_cvedestino from c_rharcosperseg where n_cveorigen=98", gConSql, adOpenStatic, adLockReadOnly
Else
    adors.Open "select n_cvedestino from c_rharcosperseg where n_cveorigen=" & iPerseg & " and n_cveevento=" & iEvento & " and n_resultado=" & iRes & " order by n_orden", gConSql, adOpenStatic, adLockReadOnly
End If
Do While Not adors.EOF
    s = s & adors(0) & ","
    adors.MoveNext
Loop
If Len(s) > 0 Then s = Mid(s, 1, Len(s) - 1)
IDPerSeg_Sig = s
End Function

'Para PictureBox
Public Sub LeerBinaryPic(campoBinary As Field, unPicture As PictureBox, iNum As Long)
    'Leer la imagen del campo de la base y asignarlo al Picture
    Dim DataFile As Integer
    Dim Chunk() As Byte

    Dim lngCompensaci�n As Long
    Dim lngTama�oTotal As Long

    'Se usa un fichero temporal para guardar la imagen
    DataFile = FreeFile
    Open "pictemp" & iNum For Binary Access Write As DataFile

    lngTama�oTotal = campoBinary.ActualSize
    Do While lngCompensaci�n < lngTama�oTotal

        'Chunk() = campoBinary.GetChunk(lngCompensaci�n)
        Chunk() = campoBinary.GetChunk(conChunkSize)
        Put DataFile, , Chunk()
        lngCompensaci�n = lngCompensaci�n + conChunkSize
    Loop

    Close DataFile
    'Ahora se carga esa imagen en el control
    On Local Error Resume Next
    unPicture.Picture = LoadPicture("pictemp" & iNum)

    'Ya no necesitamos el fichero, as� que borrarlo

    If Len(Dir$("pictemp" & iNum)) Then
        Kill "pictemp" & iNum
    End If
    Err = 0
End Sub

'Para Image
Public Sub LeerBinary(campoBinary As Field, unPicture As Image, iNum As Long)
    'Leer la imagen del campo de la base y asignarlo al Picture
    Dim DataFile As Integer
    Dim Chunk() As Byte

    Dim lngCompensaci�n As Long
    Dim lngTama�oTotal As Long

    'Se usa un fichero temporal para guardar la imagen
    DataFile = FreeFile
    Open "pictemp" & iNum For Binary Access Write As DataFile

    lngTama�oTotal = campoBinary.ActualSize
    Do While lngCompensaci�n < lngTama�oTotal

        'Chunk() = campoBinary.GetChunk(lngCompensaci�n)
        Chunk() = campoBinary.GetChunk(conChunkSize)
        Put DataFile, , Chunk()
        lngCompensaci�n = lngCompensaci�n + conChunkSize
    Loop

    Close DataFile
    'Ahora se carga esa imagen en el control
    On Local Error Resume Next
    unPicture.Picture = LoadPicture("pictemp" & iNum)

    'Ya no necesitamos el fichero, as� que borrarlo

    If Len(Dir$("pictemp" & iNum)) Then
        Kill "pictemp" & iNum
    End If
    Err = 0
End Sub


Public Sub GuardarBinary(campoBinary As Field, unPicture As Image, iNum As Long)
    'Guardar el contenido del Picture en el campo de la base
    Dim i As Integer
    Dim Fragment As Integer, Fl As Long, Chunks As Integer
    Dim DataFile As Integer
    Dim Chunk() As Byte
    
    '
    'NOTA:
    '   El recordset debe estar preparado para Editar o A�adir
    '
    
    'Guardar el contenido del picture en un fichero temporal
    SavePicture unPicture.Picture, "pictemp" & iNum
    
    'Leer el fichero y guardarlo en el campo
    DataFile = FreeFile
    Open "pictemp" & iNum For Binary Access Read As DataFile
    Fl = LOF(DataFile)    ' Longitud de los datos en el archivo
    If Fl = 0 Then Close DataFile: Exit Sub
    
    Chunks = Fl \ conChunkSize
    Fragment = Fl Mod conChunkSize
    ReDim Chunk(Fragment)
    
    Get DataFile, , Chunk()
    campoBinary.AppendChunk Chunk()
    ReDim Chunk(conChunkSize)
    For i = 1 To Chunks
        Get DataFile, , Chunk()
        campoBinary.AppendChunk Chunk()
    Next i
    Close DataFile
    
    'Ya no necesitamos el fichero, as� que borrarlo
    On Local Error Resume Next
    If Len(Dir$("pictemp" & iNum)) Then
        Kill "pictemp" & iNum
    End If
    Err = 0
End Sub

Public Sub GuardarBinary2(campoBinary As Field, unPicture As PictureBox, iNum As Long)
    'Guardar el contenido del Picture en el campo de la base
    Dim i As Integer
    Dim Fragment As Integer, Fl As Long, Chunks As Integer
    Dim DataFile As Integer
    Dim Chunk() As Byte
    
    '
    'NOTA:
    '   El recordset debe estar preparado para Editar o A�adir
    '
    
    'Guardar el contenido del picture en un fichero temporal
    SavePicture unPicture.Picture, "pictemp" & iNum
    
    'Leer el fichero y guardarlo en el campo
    DataFile = FreeFile
    Open "pictemp" & iNum For Binary Access Read As DataFile
    Fl = LOF(DataFile)    ' Longitud de los datos en el archivo
    If Fl = 0 Then Close DataFile: Exit Sub
    
    Chunks = Fl \ conChunkSize
    Fragment = Fl Mod conChunkSize
    ReDim Chunk(Fragment)
    
    Get DataFile, , Chunk()
    campoBinary.AppendChunk Chunk()
    ReDim Chunk(conChunkSize)
    For i = 1 To Chunks
        Get DataFile, , Chunk()
        campoBinary.AppendChunk Chunk()
    Next i
    Close DataFile
    
    'Ya no necesitamos el fichero, as� que borrarlo
    On Local Error Resume Next
    If Len(Dir$("pictemp" & iNum)) Then
        Kill "pictemp" & iNum
    End If
    Err = 0
End Sub

'Function FU_DatosServerExt() As Boolean
'Dim S_CadenaconExt As String, S_ServerExt As String, S_BaseDatosExt As String
'Dim S_LogExt As String, S_PassExt As String
'
'FU_DatosServerExt = False
'
'S_LogExt = Fu_LeeDatosArchConfig(1, "Central")
'S_PassExt = Fu_LeeDatosArchConfig(2, "Central")
'S_BaseDatosExt = Fu_LeeDatosArchConfig(3, "Central")
'S_ServerExt = Fu_LeeDatosArchConfig(4, "Central")
''-----------------------------------------------
''S_LogExt = FUsDeCodifica(Trim(S_LogExt))
''S_PassExt = FUsDeCodifica(Trim(S_PassExt))
''--------------------------------------------------------------------------------------------
'S_CadenaconExt = "User ID=" & S_LogExt & ";Password=" & S_PassExt & ";Data Source=" & S_ServerExt & ";Initial Catalog=" & S_BaseDatosExt
'If Not FUConecta(S_CadenaconExt) Then Exit Function           'Conexi�n cerrada
'FU_DatosServerExt = True
''*******************************************************************************************
''*Arma la cadena de conexi�n
''*******************************************************************************************
'End Function

'Function Fu_LeeDatosArchConfig(I_IdLinea, S_CualArchivo) As String
'Dim S_RutaIni As String, S_MiArchivo As String, I_Linea As Integer, S_Path As String, S_Contenido As String
'S_Path = CurDir
'If UCase(Trim(S_CualArchivo)) = "CENTRAL" Then
'    S_RutaIni = Trim(S_Path) & "\conecta.ini"
'ElseIf UCase(Trim(S_CualArchivo)) = "LOCAL" Then
'    'S_RutaIni = Trim(S_Path) & "\ACSconecta.ini"
'End If
'
'S_MiArchivo = Dir(S_RutaIni, vbArchive)
'If S_MiArchivo = "" Then
'    Fu_LeeDatosArchConfig = ""
'Else
'    I_Linea = 0
'    Open S_RutaIni For Input As #1    ' Abre el archivo.
'    Do While Not EOF(1)
'        I_Linea = I_Linea + 1
'        Line Input #1, S_Contenido
'        If I_Linea = I_IdLinea Then Exit Do
'    Loop
'    Close #1
'
'    If Len(Trim(S_Contenido)) > 0 Then
'        S_Contenido = UCase(Left(S_Contenido, 1)) & LCase(Right(S_Contenido, Len(Trim(S_Contenido)) - 1))
'    End If
'    Fu_LeeDatosArchConfig = S_Contenido
'End If
''********************************************************************************************
''*Rescata nombre de la base de datos, de acuerdo a la posici�n del n�mero de linea
''*3 - Nombre de la base de Datos
''*4 - Direcci�n IP
''********************************************************************************************
'End Function
'
'
'Public Function FUConecta(sconexion, Optional S_IP As String = "99.99.99.99") As Boolean
'Static N_CuentaSinConec As Long
'On Error GoTo ErrorConexion
'Dim bynum As Byte
'
'bynum = 0
'gsConexi�n = sconexion
'With gConSql
'    .CursorLocation = adUseClient           'La posici�n de un motor de cursores
'    .CommandTimeout = 50
'    .Provider = "SQLOLEDB"
'    .ConnectionString = sconexion
'    .Open
'End With
''- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
'FUConecta = True
'Exit Function
'
'ErrorConexion:
'    N_CuentaSinConec = N_CuentaSinConec + 1
'    If InStr(Err.Description, "08001") > 0 Or (InStr(Err.Description, "01000") > 0 And InStr(Err.Description, "gethostbyname()()") > 0) Then
'        MsgBox "No se puede establecer la comunicaci�n." & vbCrLf & "Verifique la configuraci�n de la red y/o que el servidor este activo o en su defecto que exista.", vbOKOnly, "Error en la conexi�n, intente nuevamente."
'    ElseIf InStr(Err.Description, "01000") > 0 And InStr(Err.Description, "connect()") > 0 Then
'        MsgBox "No se puede establecer la comunicaci�n. Verifique la conexi�n f�sica de la red.", vbOKOnly, "Error en la conexi�n, intente nuevamente."
'    ElseIf InStr(Err.Description, "08004") > 0 Or (InStr(Err.Description, "01000") > 0 And InStr(Err.Description, "Changed database") > 0) Then
'        MsgBox "No se localiz� la base datos. Verifique el nombre de la base de datos que captur�.", vbOKOnly, "Error en la conexi�n, intente nuevamente."
'    ElseIf InStr(Err.Description, "28000") > 0 And InStr(Err.Description, "Login failed") > 0 Then
'        MsgBox "Error en la contrase�a. Verifique la informaci�n que captur�.", vbOKOnly, "Error en la conexi�n, intente nuevamente."
'    ElseIf (InStr(Err.Description, "Time") > 0 Or InStr(Err.Description, "Tiempo") > 0) And InStr(Err.Description, "S1T00") > 0 And bynum <= 2 Then
'        bynum = bynum + 1
'        Resume
'    Else
'        If Val(Err.Number) = -2147217843 Then    'Error de inicio de sesi�n del usuario
'            MsgBox "No se realiz� la conexi�n. Descripci�n del error: " & Err.Description & " IP: " & S_IP, 0 + 16, "Verificar el usuario en la local"
'        Else
'            MsgBox "No se realiz� la conexi�n. Descripci�n del error: " & Err.Description & " IP: " & S_IP, 0 + 16, "Error en la conexi�n, intente nuevamente."
'        End If
'    End If
''*******************************************************************************************
''*Sirve para conectarse al servidor central
''*******************************************************************************************
'End Function

Function GuardaBit�cora(iUsuario As Long, yTabla As Byte, lRegistro As Long, yTipo As Byte)
Dim adors As New ADODB.Recordset, lEmp As Long, lPto As Long, lPer As Long
If gConSqlTrans.State = 0 Then
    Call EstableceConexi�nServidor(gsConexi�n, gConSqlTrans)
End If
adors.Open "select dbo.f_Usuario_idpto(" & gs_usuario & "), dbo.f_Puesto_idemp(dbo.f_Usuario_idpto(" & gs_usuario & ")),dbo.f_empleado_idper(dbo.f_Puesto_idemp(dbo.f_Usuario_idpto(" & gs_usuario & ")))", gConSql, adOpenStatic, adLockReadOnly
If Not adors.EOF Then
    lPto = IIf(Not IsNull(adors(0)), adors(0), -1)
    lEmp = IIf(Not IsNull(adors(1)), adors(1), -1)
    lPer = IIf(Not IsNull(adors(2)), adors(2), -1)
End If
gConSqlTrans.Execute "insert into c_RHBitacora (n_cveusuario,n_cvetabla,n_cveRegistro,n_cveTipoMov,d_fecha,n_cvePersonal,n_cveEmpleado,n_cvePuesto) values (" & iUsuario & "," & yTabla & "," & lRegistro & "," & yTipo & ",getdate()," & lPer & "," & lEmp & "," & lPto & ")"
End Function

Function MesCorto(sMes As String) As String
Select Case sMes
  Case "01", " 1"
     MesCorto = "Ene"
 Case "02", " 2"
    MesCorto = "Feb"
 Case "03", " 3"
    MesCorto = "Mar"
Case "04", " 4"
   MesCorto = "Abr"
 Case "05", " 5"
   MesCorto = "May"
 Case "06", " 6"
   MesCorto = "Jun"
Case "07", " 7"
   MesCorto = "Jul"
Case "08", " 8"
   MesCorto = "Ago"
 Case "09", " 9"
   MesCorto = "Sep"
 Case "10", "10"
   MesCorto = "Oct"
Case "11", "11"
   MesCorto = "Nov"
Case "12", "12"
   MesCorto = "Dic"
 End Select
End Function



