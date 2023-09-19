Attribute VB_Name = "Modulo1"
Option Explicit
Const csSepara = "||"
Const cnVerde = &HC000&
Const cnRojo = &HFF&

Public miActividad As Integer 'Tiene el valor de idact de la actividad que se está registrando
Public miTarea As Integer 'Tiene el valor de idtar de la actividad que se está registrando
Public mlAnálisis As Long 'Tiene el valor de idana del Oficio que se da seguimiento
Public mlAnt As Long 'Tiene el valor de idant correspondiente al registro que se esta realizando
Public mlSeguimiento As Long 'Tiene el valor de id del avance que se está editando
Public miDesenlace As Integer 'Valor de iddes
Public miResponsable As Integer 'Valor de idres
Public msDoctos As String 'Contiene el valos de los documentos emitidos en el avance
Public msActsProg As String 'Contiene el valos de las siguientes acts. programadas
Public msObservaciones As String 'Contiene el valos de las observaciones
Public mdFecha As Date 'Fecha del avance
Public msProgResp As String 'Fecha programada y Responsable
Public miRespProg As Integer 'idusi Responsable act. programada
Public msAcuerdo As String 'Contiene información del No. de acuerdo asosiado al avance
Public msMemo As String 'Contiene información del No. de memorando asosiado al avance
Public msSanción As String 'Contiene información de los datos de la sanción
Public msCondonación As String 'Contiene información de los datos de la condonación

Public yTipoOperación As Byte '1:Agrega/Seguimiento; 2:Modifica; 0:Consulta

Public sFechaMín As String
Public sFechaMáx As String
Public bAceptar As Boolean
Public ySoloConsulta As Byte
'Public yUltimo As Byte
Public yFormaRecepción As Byte

Public Canceled As Boolean

Private dlgProgress As frmProgress
Dim TransferOperationActive As Boolean

Dim sArchivoFTP As String 'Contiene el nombre del archivo subido  a FTP para publicar en el portal de estrados
Dim sHostRemoto As String 'Contiene el nombre del servidor FTP remoto
Dim yVerificaDocto As Byte 'indicador si ya se trató visualizar el Docto.

Dim iActApo As Integer
Dim bRR As Boolean
Dim bOprimióTecla As Boolean
Dim lcuestionario As Long 'Cuestionario
Dim sCuestionario As String, sCuestionario1 As String 'Preguntas y respuestas del cuestionario
Dim sAhora As String 'Fecha/hora del servidor
Dim yUnico As Byte
Dim sQuitaNodo(2) As String
Dim sActividadesActivas As String 'Actividades activas que determinan actividades por programar, Documentos y desenlaces visibles
Dim bGuarda As Boolean
Dim bBloqueo As Boolean
Dim bPrograma As Boolean
Dim rsResponsables As Recordset 'Cursor de responsables
Dim sResponsables As String
Dim rsArcos As Recordset 'Cursor de los arcos
Public sConclusión As String
Dim yHabilita As Byte
Dim lSegundos As Long
Dim sSanción As String
Dim sCondonación As String
Dim sArchivos As String 'Contiene el nombre de los archivos que se generan por los documentos del SIO (tareas realizadas)
'Dim bModifica As Boolean 'Indicador si se está modificando la Pantalla Cualitativa
Dim bNodoSeleccionado As Boolean 'indica que paso el procedimiento de selección del nodo seleccionado
Dim adorsBloqueoAva As New ADODB.Recordset 'Recorset para bloqueo ADO
Dim rsBloqueoAva As DAO.Recordset 'Recorset para bloqueo DAO

Dim sGestiónDoctos As String
Dim SComentariosUNE As String 'Comentarios adicionales o documentación adicional a la UNE
Dim bGestiónDoctos As Boolean 'Indica si debe guardar documentos seleccionados al ejecutarse la actividad de UNES



Private Sub btnCancel_Click()
  Canceled = True
End Sub


Private Sub chkacuerdo_Click()
bOprimióTecla = True
HabilitaAceptar False
End Sub

Private Sub cmdBotón_Click(Index As Integer)
Dim i As Integer, s As String, yy As Integer, Y As Long, adors As New ADODB.Recordset, yErr As Byte, d As Date
Dim l As Long
bGuarda = False
On Error GoTo ErrorBloqueo:
If Index = 1 Then 'Cancelar
    Unload Me
    Exit Sub
End If
If yTipoOperación = 0 Then
    Unload Me
    Exit Sub
End If
bAceptar = True
HabilitaAceptar
s = ""
miTarea = 0
For i = 1 To TreeView3.Nodes.Count
    If TreeView3.Nodes(i).Checked Then
        miTarea = Val(Right(TreeView3.Nodes(i).Key, 4))
    End If
Next
If miTarea = 0 Then
    s = "Falta dato requerido (" + etiArbol3.Caption + ")"
    yy = 1
End If
If Not IsDate(txtcampo(1).Text) Then
    s = "Fecha incorrecta (" + etiTexto(1).Caption + ")"
    yy = 1
End If
If ComboResponsable.ListIndex >= 0 Then
    miResponsable = ComboResponsable.ItemData(ComboResponsable.ListIndex)
Else
    If Len(s) = 0 Then
        s = "Falta dato requerido (Responsable)"
        yy = 20
    End If
End If
If IsDate(sFechaMín) Then
    If Len(s) = 0 And CDate(txtcampo(1).Text) < CDate(Mid(sFechaMín, 20 * Y + 1, 20)) Then
        s = "La fecha de inicio no puede ser menor a la fecha de inicio de la actividad que le precede (" + Mid(sFechaMín, 20 * Y + 1, 20) + ")"
        yy = 1
    End If
End If
If IsDate(sFechaMín) Then
    If Len(s) = 0 And CDate(txtcampo(1).Text) > CDate(Mid(sFechaMáx, 20 * Y + 1, 20)) Then
        s = "La fecha de inicio no puede ser mayor a la fecha de inicio de la actividad que le sucede (" + Mid(sFechaMáx, 20 * Y + 1, 20) + ")"
        'txtCampo(Y).SetFocus
        yy = 1
    End If
End If
'dAhora = Now 'HORASERVIDOR

If Len(s) > 0 Then
    MsgBox s, vbOKOnly, "Validación"
    If yy > 10 Then
        ComboResponsable.SetFocus
    ElseIf txtcampo(yy).Visible And txtcampo(yy).Enabled Then
        txtcampo(yy).SetFocus
    End If
    Exit Sub
End If
msDoctos = ""
For i = 1 To TreeView2.Nodes.Count
    If TreeView2.Nodes(i).Checked Then
        msDoctos = msDoctos & Right(TreeView2.Nodes(i).Key, 4) & "|"
    End If
Next
msActsProg = ""
For Y = 1 To TreeView1.Nodes.Count
    i = NodoContieneFecha(TreeView1.Nodes(Y))
    If TreeView1.Nodes(Y).Checked And i > 0 Then
        s = Mid(TreeView1.Nodes(Y).Text, InStrRev(TreeView1.Nodes(Y).Text, " Resp.: ") + 8)
        If adors.State > 0 Then adors.Close
        adors.Open "select * from usuariossistema where descripción='" & Mid(s, 1, Len(s) - 1) + "'", gConSql, adOpenStatic, adLockReadOnly
        s = Mid(TreeView1.Nodes(Y).Text, i + 1, InStrRev(TreeView1.Nodes(Y).Text, " Resp.: ") - i - 1)
        If IsDate(s) Then
            s = Format(CDate(s), "dd/mm/yyyy hh:mm")
        Else
            s = ""
        End If
        If adors.RecordCount = 0 Then
            msActsProg = msActsProg & Right(TreeView1.Nodes(Y).Key, 4) & "|" & s & "||"
        Else
            msActsProg = msActsProg & Right(TreeView1.Nodes(Y).Key, 4) & "|" & s & "|" & Trim(Str(adors!ID)) & "|"
        End If
    Else 'Verifica si debe estar programada la actividad
        l = Val(Right(TreeView1.Nodes(Y).Key, 4))
        If adors.State > 0 Then adors.Close
        adors.Open "select forzarprog from relacióntareaactividad where idtar=" & miTarea & " and idact=" & l, gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            If adors(0) <> 0 Then
                MsgBox "La Actividad (" & TreeView1.Nodes(Y).Text & ") debe estar programada", vbOKOnly + vbInformation, "Validación"
                Exit Sub
            End If
        End If
    End If
Next

s = Replace(msDoctos, "|", "")
sArchivos = VerificaExistenciaDocumentos(s)
'MsgBox "comienza actualización"
'Exit Sub
'If yTipoOperación = 0 Then
'Else
'If yTipoOperación = 1 Then 'Alta de avance
    
    If cmdSanción.Visible Then 'Los datos de la sanción son obligatorios
        If InStr(msSanción, "|") > 0 Then 'debe guardar datos de la sanción, verifica consecutivo del oficio automático
'            If adors.State Then adors.Close
'            adors.Open "select f_nuevofolio(4,0," & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
'            If Not adors.EOF Then
'                If InStr(adors(0), "???") Then
'                    l = F_PreguntaConsecutivo(4, adors(0))
'                    If l < 0 Then 'Se Ejecutó cancelar
'                        Exit Sub
'                    End If
'                End If
'            End If
        Else
            MsgBox "los datos de la sanción son requeridos. Favor de capturarlos", vbOKOnly + vbInformation, "Validación"
            Exit Sub
        End If
    Else
        If InStr(msSanción, "|") > 0 Then
            msSanción = ""
        End If
    End If
    If cmdCondonación.Visible Then 'Los datos de la sanción son obligatorios
        If InStr(msCondonación, "|") > 0 Then 'debe guardar datos de la Condonación, verifica consecutivo del oficio automático
'            If adors.State Then adors.Close
'            adors.Open "select f_nuevofolio(4,0," & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
'            If Not adors.EOF Then
'                If InStr(adors(0), "???") Then
'                    l = F_PreguntaConsecutivo(4, adors(0))
'                    If l < 0 Then 'Se Ejecutó cancelar
'                        Exit Sub
'                    End If
'                End If
'            End If
        Else
            MsgBox "los datos de la Condonación son requeridos. Favor de capturarlos", vbOKOnly + vbInformation, "Validación"
            Exit Sub
        End If
    Else
        If InStr(msCondonación, "|") > 0 Then
            msCondonación = ""
        End If
    End If
    If cmdSubirDocto.Visible Then 'Es obligatorio subir Docto
        If Len(sArchivoFTP) = 0 Then 'debe subir el documento a FTP
            MsgBox "Debe subir el documento de notificación a estrados. Favor de subirlo", vbOKOnly + vbInformation, "Validación"
            Exit Sub
        End If
        If yVerificaDocto = 0 Then 'Debe verifica el documento a FTP
            MsgBox "Debe verificar si el documento se encuentra correctamente en estrados electrónicos. Favor de verificar", vbOKOnly + vbInformation, "Validación"
            Exit Sub
        End If
    Else
        If cmdVerificaDocto.Visible Then 'Debe verifica el documento a FTP
            If yVerificaDocto <= 0 Then 'Debe verifica el documento a FTP
                MsgBox "Debe verificar si el documento se encuentra correctamente en estrados electrónicos. Favor de verificar", vbOKOnly + vbInformation, "Validación"
                Exit Sub
            End If
            d = DíasHábiles(Int(AhoraServidor), 1)
            If MsgBox("¿Está seguro que se subió correctamente el documento que se publicará el día hábil siguiente?: " & Format(d, gsFormatoFecha), vbYesNo + vbQuestion + vbDefaultButton2, "Confirmación") = vbNo Then
                Exit Sub
            End If
        Else
            If Len(sArchivoFTP) > 0 Then
                sArchivoFTP = ""
            End If
        End If
    End If
    
    If txtAcuerdo.Visible Then
        If Len(Trim(txtAcuerdo.Text)) = 0 And chkacuerdo.Value = 0 Then
            MsgBox "El número de acuerdo es requerido", vbOKOnly + vbInformation, "Validación"
            Exit Sub
        End If
        msAcuerdo = txtAcuerdo.Text
        If adors.State Then adors.Close
        adors.Open "select f_analisis_oficio(s.idana) from seguimientoacuerdos sa, seguimiento s where sa.acuerdo='" & Replace(msAcuerdo, "'", "''") & "' and sa.idseg=s.id and s.id<>" & mlSeguimiento, gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            If MsgBox("Existe ya un acuerdo con ese número (oficio: " & adors(0) & "). ¿Está seguro de asignar el mismo número?", vbYesNo + vbQuestion + vbDefaultButton2, "Validación de Número de Acuerdo") = vbNo Then
                Exit Sub
            End If
        End If
        
''        If chkacuerdo.Value Then
''            msAcuerdo = ""
''            l = -1
''        Else
''            l = 1
''            If adors.State Then adors.Close
''            adors.Open "select f_nuevofolio(6,0," & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
''            If Not adors.EOF Then
''                If InStr(adors(0), "???") Then
''                    l = F_PreguntaConsecutivo(3, adors(0))
''                    If l < 0 Then 'Se Ejecutó cancelar
''                        Exit Sub
''                    End If
''                End If
''            End If
''        End If
''        'msAcuerdo = txtAcuerdo.Text
    Else
        'msAcuerdo = ""
        l = 0
    End If
    
    If txtcampo(3).Visible Then 'valida la captura de memo y fecha memo
        If Len(Trim(txtcampo(3).Text)) = 0 Then
            MsgBox "El número de memorando es requerido", vbOKOnly + vbInformation, "Validación"
            Exit Sub
        End If
        If Not IsDate(txtcampo(4).Text) Then
            MsgBox "La fecha del memorando es requerida", vbOKOnly + vbInformation, "Validación"
            Exit Sub
        End If
        msMemo = txtcampo(3).Text
        If adors.State Then adors.Close
        adors.Open "select f_analisis_oficio(s.idana) from seguimientomemos sm, seguimiento s where sm.memorando='" & Replace(msMemo, "'", "''") & "' and sm.idseg=s.id and s.id<>" & mlSeguimiento, gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            If MsgBox("Existe ya un memorando con ese número (oficio: " & adors(0) & "). ¿Está seguro de asignar el mismo número?", vbYesNo + vbQuestion + vbDefaultButton2, "Validación de Número de Acuerdo") = vbNo Then
                Exit Sub
            End If
        End If
        msMemo = msMemo & "|" & Format(txtcampo(4).Text, "dd/mm/yyyy") & "|"
    Else
        msMemo = ""
    End If
    
    
    If adors.State Then adors.Close
    'adors.Open "{call p_seguimientoguardadatos(" & mlSeguimiento & "," & mlAnt & "," & mlAnálisis & ",'" & Format(CDate(txtCampo(1).Text), "dd/mm/yyyy hh:mm:ss") & "'," & miActividad & "," & miTarea & "," & miResponsable & "," & giUsuario & "," & miDesenlace & ",'" & Replace(txtCampo(2).Text, "'", "''") & "','" & msDoctos & "','" & msActsProg & "'," & l & ",'" & msSanción & "','" & msCondonación & "','" & msMemo & "')}", gConSql, adOpenForwardOnly, adLockReadOnly
    adors.Open "{call p_seguimientoguardadatos2(" & mlSeguimiento & "," & mlAnt & "," & mlAnálisis & ",'" & Format(CDate(txtcampo(1).Text), "dd/mm/yyyy hh:mm:ss") & "'," & miActividad & "," & miTarea & "," & miResponsable & "," & giUsuario & "," & miDesenlace & ",'" & Replace(txtcampo(2).Text, "'", "''") & "','" & msDoctos & "','" & msActsProg & "','" & msAcuerdo & "','" & msSanción & "','" & msCondonación & "','" & msMemo & "')}", gConSql, adOpenForwardOnly, adLockReadOnly
    If adors(0) < 0 Then
        MsgBox "No se realizó el Alta del avance.", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
    mlSeguimiento = adors(0)
    
    For yy = 1 To Len(s) / 4 'Emite Documentos
        If adors.State > 0 Then adors.Close
        adors.Open "select * from documentos where id=" & Mid(s, (yy - 1) * 4 + 1, 4), gConSql, adOpenStatic, adLockReadOnly
        If adors.RecordCount > 0 Then
            If Len(adors!archivo) Then
                If Len(Dir(gsDirDocumentos + adors!archivo + ".doc")) > 0 Then
                    Call GeneraDocumento(adors, mlAnálisis, mlSeguimiento)
                End If
            End If
        End If
    Next
'End If
Unload Me
'Me.Hide
Exit Sub
ErrorBloqueo:
If gConSql.Errors.Count > 0 Then
    yErr = MsgBox("Error: " + Err.Description + ". vuelva a intentar", vbOKOnly + vbCritical, "Error no esperado (" + Str(IIf(Err.Number < 0, gConSql.Errors(gConSql.Errors.Count - 1).Number, Err.Number)) + ")")
Else
    yErr = MsgBox("Error: " & Err.Description & ". vuelva a intentar", vbOKOnly + vbCritical, "Error no esperado (" & IIf(Err.Number > 0, Err.Number, "???") & ")")
End If
If yErr = vbCancel Or yErr = vbAbort Then
    Exit Sub
ElseIf yErr = vbRetry Then
    Resume
ElseIf yErr = vbIgnore Then
    Resume Next
End If
End Sub

Private Sub cmdCondonación_Click()
With Condonacion
    If Not IsDate(txtcampo(1).Text) Then
        Call MsgBox("Debe capturar correctamente la fecha de la presente actividad", vbOKOnly + vbInformation, "")
        Exit Sub
    End If
    gs = msCondonación
    gs1 = mlAnálisis
    gs2 = mlSeguimiento
    If yTipoOperación = 0 Then
        .myAcción = 0
        .txtcampo(1).Locked = True
        .txtcampo(2).Locked = True
        .txtcampo(3).Locked = True
        .txtcampo(4).Locked = True
        .txtcampo(5).Locked = True
        .txtcampo(6).Locked = True
        .Check1.Enabled = False
    Else
        .myAcción = 1
    End If
    .mdFechaOficio = CDate(txtcampo(1).Text)
    .Show vbModal
    If gs <> "cancelar" Then
        msCondonación = gs
    End If
    bOprimióTecla = True
    HabilitaAceptar
End With
End Sub

Private Sub cmdSanción_Click()
With Sanción
    '.txtCampo(0).SetFocus
    If Not IsDate(txtcampo(1).Text) Then
        Call MsgBox("Debe capturar correctamente la fecha de la presente actividad", vbOKOnly + vbInformation, "")
        Exit Sub
    End If
    gs = msSanción
    gs1 = mlAnálisis
    gs2 = mlSeguimiento
    If yTipoOperación = 0 Then
        .myAcción = 0
        .txtcampo(1).Locked = True
        .txtcampo(2).Locked = True
        .txtcampo(3).Locked = True
        .ComboUnidad.Locked = True
    Else
        .myAcción = 1
    End If
    .mdFechaOficio = CDate(txtcampo(1).Text)
    .Show vbModal
    If gs <> "cancelar" Then
        msSanción = gs
    End If
    bOprimióTecla = True
    HabilitaAceptar
End With
End Sub

Private Sub cmdSubirDocto_Click()
Dim sArchivo As String, adors As New ADODB.Recordset, sArchivoOK As String
Dim s As String
Dim yErr As Byte
On Error GoTo salir:
Do While Len(Trim(sArchivo)) = 0
    If adors.State Then adors.Close
    adors.Open "select idpro from tareas where id=" & miTarea, gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        If adors(0) = 4 Then 'Sanciones Busca el oficio de sanción
            If adors.State Then adors.Close
            adors.Open "select f_sancion_oficio(" & mlAnálisis & "," & mlAnt & ") from dual", gConSql, adOpenStatic, adLockReadOnly
        Else 'Emplazamiento Busca el oficio en análisis
            If adors.State Then adors.Close
            adors.Open "select f_analisis_oficio(" & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
        End If
        If Not adors.EOF Then
            If Not IsNull(adors(0)) Then
                sArchivoOK = Replace(adors(0), "/", "_")
            End If
        End If
    End If
    Do While True
        With CommonDialog1
            .DialogTitle = "Insertar archivo PDF"
            .CancelError = True
            '.Filter = "Todos los archivos PDFs|*.pdf;Todos los archivos (*.*)|*.*"
            .Filter = "Archivos pdf|*.pdf;Todos los archivos (*.*)|*.*"
            .FileName = sArchivoOK & ".pdf"
            .FilterIndex = 1
            .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
            .ShowOpen
            sArchivo = .FileName
            If InStr(sArchivo, "\") > 0 Then
                s = Mid(sArchivo, InStrRev(sArchivo, "\") + 1)
            End If
            If LCase(s) <> LCase(sArchivoOK & ".pdf") Then
                Call MsgBox("El archivo debe contener el nombre: " & sArchivoOK & ".pdf", vbOKOnly + vbInformation, "Validación")
            Else
                Exit Do
            End If
        End With
    Loop
Loop
'If SubeFTP3(sArchivo) Then
If SubeFTP(sArchivo) Then
    sArchivoFTP = sArchivo
    cmdSubirDocto.BackColor = cnVerde
Else
    sArchivoFTP = ""
    cmdSubirDocto.BackColor = cnRojo
End If
Exit Sub
salir:
If InStr(LCase(Err.Description), "cancelar") > 0 Then
    sArchivoFTP = ""
    cmdSubirDocto.BackColor = cnRojo
    Exit Sub
End If
yErr = MsgBox("Error no esperado.", vbAbortRetryIgnore, "")
If yErr = vbRetry Then
    Resume
ElseIf yErr = vbIgnore Then
    Resume Next
End If
End Sub

Private Sub cmdVerificaDocto_Click()
yVerificaDocto = 1
If InStr(sArchivoFTP, "\") > 0 Then
    gsWWW = "http://portalif.condusef.gob.mx/estrados/admin/files1/" & Mid(sArchivoFTP, InStrRev(sArchivoFTP, "\") + 1)
Else
    gsWWW = "http://portalif.condusef.gob.mx/estrados/admin/files1/" & "/" & sArchivoFTP
End If
With Browser
    .yÚnicavez = 0
    .Show vbModal
End With
End Sub

'Dim sPantallasBorrarXconsolidar As String
'Public yPantCualitativas As Byte

'Dim sProgramarAct As String 'Cadena que contiene la programación de las siguientes actividades

Private Sub ComboDesenlaces_Click()
Dim ss As String, yy As Byte, Y As Integer, l As Long, sC As String, i As Integer
If txtEtiqueta.Visible Then txtEtiqueta.Visible = False
lSegundos = -1
If yUnico = 0 Or yUnico = 200 Then Exit Sub
bOprimióTecla = True
HabilitaAceptar True
'For i = 1 To TreeView3.Nodes.Count
'    If TreeView3.Nodes(i).Checked Then
'        ProgramaActividad (Val(Right(TreeView3.Nodes(i).Key, 4)))
'    End If
'Next
End Sub

Private Sub ComboDesenlaces_GotFocus()
'ComboDesenlaces.Enabled = (IsDate(txtCampo(2)) Or Not txtCampo(2).Visible)
End Sub

Private Sub ComboDesenlaces_LostFocus()
Dim s As String
If (Len(Trim(ComboDesenlaces.Text)) > 0 And ComboDesenlaces.ListIndex < 0) Then
    s = ComboDesenlaces.Text
    If Val(s) > 0 Then
        ComboDesenlaces.ListIndex = BuscaComboClave(ComboDesenlaces, ComboDesenlaces.Text, False, False)
    Else
        ComboDesenlaces.ListIndex = BuscaCombo(ComboDesenlaces, ComboDesenlaces.Text, False, True)
    End If
    If ComboDesenlaces.ListIndex < 0 And Len(Trim(ComboDesenlaces)) > 0 Then ComboDesenlaces = ""
End If
End Sub

Private Sub ComboResponsable_Click()
If txtEtiqueta.Visible Then txtEtiqueta.Visible = False
lSegundos = -1
HabilitaAceptar
End Sub


Private Sub ComboResponsable_LostFocus()
Dim rs As Recordset, s As String, adors As New ADODB.Recordset, i As Integer
If Len(Trim(ComboResponsable.Text)) > 0 And ComboResponsable.ListIndex < 0 Then
    s = ComboResponsable.Text
    ComboResponsable.ListIndex = BuscaCombo(ComboResponsable, ComboResponsable.Text, False, True)
End If
HabilitaAceptar
End Sub

Private Sub etiArbol1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim aValor(6) As Integer
'Call MuestraEtiqueta(etiArbol1, txtEtiqueta, 0, lSegundos, aValor)
End Sub

Private Sub etiArbol2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim aValor(6) As Integer
'Call MuestraEtiqueta(etiArbol2, txtEtiqueta, 0, lSegundos, aValor)
End Sub

Private Sub etiArbol3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim aValor(6) As Integer
'Call MuestraEtiqueta(etiArbol3, txtEtiqueta, 0, lSegundos, aValor)
End Sub

Private Sub EtiCombo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim aValor(6) As Integer
'Call MuestraEtiqueta(etiCombo(Index), txtEtiqueta, 0, lSegundos, aValor)
End Sub

Private Sub etiTexto_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim aValor(6) As Integer
'Call MuestraEtiqueta(etiTexto(Index), txtEtiqueta, 0, lSegundos, aValor)
End Sub

Private Sub Form_Activate()
Dim Y As Byte, s As String, adors As New ADODB.Recordset, i As Integer, ss As String, yy As Byte
Dim i1 As Integer
Dim s2 As String
'LicenseManager.SetLicenseKey ("7C19F6E1416C480FD3CBB509133177EE9F9F5113722D31910A08BFCE618A52807D2793EA40EF65D2BE5FD6A3D2F555C6523AB270F56522235C68936422CCE57B9A71F3A22A8B34E40B1FD1E173D6634BC954F9EB6DBAF523F064F4FD75B910C749155FE6E74C1343E621FF459619D6C8C9008580C91EB5799498C718050C9AC1EC657031FDB6A9FF573EE6DB9BF0C7CBB672EC305696ABBC8D682DB762236DF711950AB629FF589FA86B0759EEBE9F1155DD1302AB276F879C9074B0F20D5EC3444608772A9A845A2A7CBEDAC15DA1A0050DC7A687788CF1E1ECC040D50AAEC6957F385BAA7C177307D5DB091711D87B7E678459AD3FABF39E089073DE762EC3")
bOprimióTecla = True
If yUnico = 0 Then
    'Bloquea controles en caso de consulta
    If yTipoOperación = 0 Then 'Consulta
        TreeView1.Enabled = False
        TreeView2.Enabled = False
        TreeView3.Enabled = False
        txtcampo(1).Locked = True
        txtcampo(2).Locked = True
        cmdBotón(0).Enabled = True
    Else ' los habilita
        TreeView1.Enabled = True
        TreeView2.Enabled = True
        TreeView3.Enabled = True
        txtcampo(1).Locked = False
        txtcampo(2).Locked = False
        'cmdBotón(0).Enabled = True
        If yTipoOperación = 1 Then 'inicializavar
            msSanción = ""
            msCondonación = ""
            msAcuerdo = ""
        End If
    End If
    If mlSeguimiento > 0 Then
        'valores del seguimiento
        If adors.State Then adors.Close
        adors.Open "select * from seguimiento where id=" & mlSeguimiento, gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            mlAnt = adors!idant
            mlAnálisis = adors!idana
            mdFecha = adors!FECHA
            miResponsable = adors!idres
            miActividad = adors!idact
            miTarea = adors!idtar
            miDesenlace = IIf(IsNull(adors!iddes), 0, adors!iddes)
        End If
        'documentos
        If adors.State Then adors.Close
        adors.Open "select * from seguimientodoctos where idseg=" & mlSeguimiento, gConSql, adOpenStatic, adLockReadOnly
        msDoctos = ""
        Do While Not adors.EOF
            msDoctos = msDoctos & adors!iddoc & "|"
            adors.MoveNext
        Loop
        'actividades programadas
        If adors.State Then adors.Close
        adors.Open "select * from seguimientoprog where idant=" & mlSeguimiento, gConSql, adOpenStatic, adLockReadOnly
        msActsProg = ""
        Do While Not adors.EOF
            msActsProg = msActsProg & Right("0000" & adors!idact, 4) & "|"
            adors.MoveNext
        Loop
        'No de acuerdo
        If adors.State Then adors.Close
        adors.Open "select acuerdo from seguimientoacuerdos where idseg=" & mlSeguimiento, gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            msAcuerdo = adors(0)
            'txtAcuerdo.Text = msAcuerdo
        Else
            msAcuerdo = ""
        End If
        'No y fecha de memorando
        If adors.State Then adors.Close
        adors.Open "select memorando,fecha from seguimientomemos where idseg=" & mlSeguimiento, gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            msMemo = adors(0) & "|" & Format(adors(1), "dd/mm/yyyy") & "|"
        Else
            msMemo = ""
        End If
        'Datos Sanción
        If adors.State Then adors.Close
        adors.Open "select f_sancion_sanxcau(" & mlSeguimiento & ") from dual", gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            msSanción = adors(0)
        Else
            msSanción = ""
        End If
        'Datos Condonación
        If adors.State Then adors.Close
        adors.Open "select f_cond_condxcau(" & mlSeguimiento & ") from dual", gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            'msCondonación = adors!OFICIO & "|" & Format(adors!FECHA, "dd/mm/yyyy") & "|" & adors!Porcentaje & "|"
            msCondonación = adors(0)
        Else
            msCondonación = ""
        End If
        'Observaciones
        If adors.State Then adors.Close
        adors.Open "select * from seguimientoobs where idseg=" & mlSeguimiento, gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            msObservaciones = adors!Observaciones
        Else
            msObservaciones = ""
        End If
    Else
        If miActividad > 0 Then
            If adors.State Then adors.Close
            adors.Open "select f_actividad(" & miActividad & ") from dual", gConSql, adOpenStatic, adLockReadOnly
            If adors.EOF Then
                s = ""
            Else
                If IsNull(adors(0)) Then
                    s = ""
                Else
                    s = adors(0)
                End If
            End If
            txtcampo(0).Text = s
        End If
    End If
    yUnico = 200
    Call Actualiza
    'En caso de haber responsable lo coloca
    If miResponsable > 0 Then
        i = BuscaCombo(ComboResponsable, miResponsable, True)
        If i >= 0 Then
            ComboResponsable.ListIndex = i
        End If
    Else 'Es nuava act. busca el resp igual al usuario
        If adors.State Then adors.Close
        adors.Open "select count(*) from usuariossistema where id=" & giUsuario & " and responsable<>0", gConSql, adOpenStatic, adLockReadOnly
        If adors(0) > 0 Then
            ComboResponsable.ListIndex = BuscaCombo(ComboResponsable, giUsuario, True)
        Else
            If adors.State Then adors.Close
            adors.Open "select idres from usuariossistema where id=" & giUsuario, gConSql, adOpenStatic, adLockReadOnly
            If Not adors.EOF Then
                If Not IsNull(adors(0)) Then
                    ComboResponsable.ListIndex = BuscaCombo(ComboResponsable, adors(0), True)
                End If
            End If
        End If
    End If
    'Selecciona la tarea en caso de ser mayor a cero
    If miTarea > 0 Then
        For i = 1 To TreeView3.Nodes.Count
            If Val(Right(TreeView3.Nodes(i).Key, 4)) = miTarea Then
                TreeView3.Nodes(i).Checked = True
                TreeView3_NodeCheck TreeView3.Nodes(TreeView3.Nodes(i).Index)
                Exit For
            End If
        Next
    End If
    'Selecciona los documentos en su caso contenidos en la variable msDoctos
    If InStr(msDoctos, "|") > 0 Then
        For i = 1 To TreeView2.Nodes.Count
            If InStr("|" & msDoctos & "|", "|" & Val(Right(TreeView2.Nodes(i).Key, 4)) & "|") > 0 Then
                TreeView2.Nodes(i).Checked = True
            End If
        Next
    End If
    'Selecciona las programadas contenidos en la variable msActsProg
    If InStr(msActsProg, "|") > 0 Then
        For i = 1 To TreeView1.Nodes.Count
            i1 = InStr("|" & msActsProg & "|", "|" & Right(TreeView1.Nodes(i).Key, 4) & "|")
            If i1 > 0 Then 'Arcma cadena que debe asignar al nodo del árbol (Fecha dd/mm/yyyy hh:mm Resp.: Responsable)
                TreeView1.Nodes(i).Checked = True
                s = Mid(msActsProg, i1)
                If adors.State Then adors.Close
                adors.Open "select sp.fecha,us.descripción from seguimientoprog sp, usuariossistema us where sp.idant=" & mlSeguimiento & " and sp.idact=" & Val(s) & " and sp.idusi=us.id(+)", gConSql, adOpenStatic, adLockReadOnly
                
                ss = Mid(s, 1, InStr(s, "|") - 1)
                If Not adors.EOF Then
                    ss = " (" & Format(adors(0), "dd/mmm/yyyy hh:mm") & "  Resp.: " & adors(1) & ")"
                Else
                    ss = " (  Resp.: ???)"
                End If
                TreeView1.Nodes(i).Text = TreeView1.Nodes(i).Text & ss
                s = Mid(s, InStr(s, "|") + 1)
                
            End If
        Next
    Else
        Call CargaActsProg(miTarea)
    End If
    'En caso de haber desenlace lo coloca
    If miDesenlace > 0 Then
        i = BuscaCombo(ComboDesenlaces, miDesenlace, True)
        If i >= 0 Then
            ComboDesenlaces.ListIndex = i
        End If
    End If
    'Coloca fecha y observaciones en caso que existan en la variable correspondiente
    If Not IsNull(mdFecha) Then
        If Year(mdFecha) > 2000 Then
            i = 200
        End If
    End If
    If i = 200 Then
        txtcampo(1).Text = Format(mdFecha, gsFormatoFechaHora)
    Else
        txtcampo(1).Text = Format(AhoraServidor, gsFormatoFechaHora)
    End If
    If Len(Trim(msObservaciones)) > 0 Then
        txtcampo(2).Text = msObservaciones
    Else
        txtcampo(2).Text = ""
    End If
    'txtCampo(0).Text = miActividad
    If yTipoOperación <= 2 And ySoloConsulta = 0 Then
        If ComboDesenlaces.ListCount = 1 And ComboDesenlaces.ListIndex < 0 Then ComboDesenlaces.ListIndex = 0
    End If
    cmdBotón(0).Enabled = Not (yTipoOperación = 2)
'    If ySoloConsulta > 0 Then
'        For y = 0 To Controls.Count - 1
'            If LCase(Mid(Controls(y).Name, 1, 3)) = "txt" Or LCase(Mid(Controls(y).Name, 1, 5)) = "combo" Then
'                Controls(y).Locked = True
'            ElseIf LCase(Mid(Controls(y).Name, 1, 5)) = "treev" Then
'                Controls(y).Enabled = False
'            End If
'        Next
'    End If
    If TreeView3.Nodes.Count > 0 And miTarea = 0 Then
        If Not TreeView3.Nodes(1).Checked Then
            TreeView3.Nodes(1).Checked = True
            TreeView3_NodeCheck TreeView3.Nodes(1)
        End If
    End If
    If txtAcuerdo.Visible Then
        If Len(Trim(msAcuerdo)) > 0 Then
            txtAcuerdo.Text = msAcuerdo
            chkacuerdo.Enabled = (yTipoOperación = 1)
        End If
    End If
    If txtcampo(3).Visible Then
        If InStr(msMemo, "|") > 0 Then
            txtcampo(3).Text = Mid(msMemo, 1, InStr(msMemo, "|") - 1)
            s = Mid(msMemo, InStr(msMemo, "|") + 1)
            txtcampo(4).Text = Mid(s, 1, InStr(s, "|") - 1)
        End If
    End If
    yUnico = 100
    HabilitaAceptar
End If

'HabilitaAceptar
'i = 0
'If TreeView1.Enabled And TreeView3.Nodes.Count = 0 Then
'    For i = 1 To TreeView1.Nodes.Count
'        If TreeView1.Nodes(i).Checked And TreeView1.Nodes(i).Children > 0 Then
'            TreeView1.Nodes(i).Checked = False
'        ElseIf TreeView1.Nodes(i).Checked Then
'            i = 200
'            Exit For
'        End If
'    Next
'    If i <> 200 Then
'        ss = Right("000" + sActividad, 4) + ","
'        For Y = 1 To TreeView3.Nodes.Count
'            If TreeView3.Nodes(Y).Checked And i = 200 Then
'                TreeView3.Nodes(Y).Checked = False
'            ElseIf TreeView3.Nodes(Y).Checked Then
'                ss = ss + Right(TreeView3.Nodes(Y).Key, 4) + ","
'                i = 200
'            End If
'        Next
'        If Len(ss) > 0 Then ss = Mid(ss, 1, Len(ss) - 1)
'        sActividadesActivas = IIf(Len(ss) = 0, Trim(Mid(sActividad, 1, 4)), ss)
'        'Actividades por programar
'        If TreeView1.Enabled Then
'            s = "SELECT a.id,b.id,c.id,d.id,a.descripción,b.descripción,c.descripción,d.descripción FROM ((actividades a LEFT JOIN actividades b ON a.id=b.idpad) LEFT JOIN actividades c ON b.id=c.idpad) LEFT JOIN actividades d ON c.id=d.idpad WHERE a.nivel=1 and (b.nivel=2 or b.nivel is null) and (c.nivel=3 or c.nivel is null) and (a.clase<>2 and a.*Forma* and a.id in (select iddestino from Arcos where idorigen" + gsSeparador + ") or b.clase<>2 and b.*Forma* and b.id in (select iddestino from Arcos where idorigen" + gsSeparador + ") or c.clase<>2 and c.*Forma* and c.id in (select iddestino from Arcos where idorigen" + gsSeparador + ") or d.clase<>2 and d.*Forma* and d.id in (select iddestino from Arcos where idorigen" + gsSeparador + ")) ORDER BY a.descripción,b.descripción,c.descripción,d.descripción" '''''
'            If adors.State > 0 Then adors.Close
'            'adors.Open "select * from actividades where id=" & IIf(Len(ss) > 0, ss, Trim(Mid(sActividad, 1, 4))), gConSQL, adOpenStatic, adLockReadOnly
'            Set adors = ObtenConsulta("select * from actividades where id=" & IIf(Len(ss) > 0, ss, Trim(Mid(sActividad, 1, 4))))
'            If Not adors(9 + yFormaRecepción) And Not adors("Personal") And Not adors("Escrito") Then
'                s2 = ""
'                For Y = 0 To 5
'                    If adors(9 + Y) Then s2 = s2 + "z." + Trim(Mid("Personal  TelefónicaInternet  Escrito   Fax       CAT       ", Y * 10 + 1, 10)) + "<>0 or "
'                Next
'                If Len(s2) > 0 Then
'                    s2 = "(" + Mid(s2, 1, Len(s2) - 4) + ")"
'                Else
'                    s2 = "false"
'                End If
'                For Y = 1 To 4
'                    s = Replace(s, Mid("abcd", Y, 1) + ".*Forma*", Replace(s2, "z.", Mid("abcd", Y, 1) + ".")) '''''
'                Next
'            Else
'                If Not adors(9 + yFormaRecepción) Then
'                    s2 = IIf(adors("Personal"), "Personal", "Escrito")
'                Else
'                    s2 = Trim(Mid("Personal  TelefónicaInternet  Escrito   Fax       CAT       ", yFormaRecepción * 10 + 1, 10))
'                End If
'                If Len(s2) > 0 Then s2 = s2 + "<>0" '-1
'                s = Replace(s, "*Forma*", s2) '''''
'            End If
'            Call CargaDatosArbolVariosNiveles(TreeView1, Replace(s, gsSeparador, " in (" + IIf(Len(ss) > 0, ss, Trim(Mid(sActividad, 1, 4))) + ")"), 4, False, True)
'            s = sActividades(6)
'            Do While Len(s) > 0
'                If adors.State > 0 Then adors.Close
'                'adors.Open "select * from actividades where id=" & Val(Mid(s, 1, 4)), gConSQL, adOpenStatic, adLockReadOnly
'                Set adors = ObtenConsulta("select * from actividades where id=" & Val(Mid(s, 1, 4)))
'                s2 = "r" + Right("000" + Trim(Str(adors!idpad)), 4) + Mid(s, 1, 4)
'                i = nodo(TreeView1, s2)
'                If i = 0 Then
'                    s2 = Mid(s, 1, InStr(InStr(InStr(s, csSepara) + 2, s, csSepara) + 2, s, csSepara) + 2)
'                    sActividades(6) = Replace(sActividades(6), s2, "")
'                    s = Replace(s, s2, "")
'                Else
'                    TreeView1.Nodes(s2).Checked = True
'                    s = Mid(s, InStr(s, csSepara) + 2) 'Borra la clave de la actividad
'                    TreeView1.Nodes(s2).Text = TreeView1.Nodes(s2).Text + " (" + Mid(s, 1, InStr(s, csSepara) - 1)
'                    s = Mid(s, InStr(s, csSepara) + 2)  'Borra la fecha programada
'                    If adors.State > 0 Then adors.Close
'                    'adors.Open "select * from Responsables where id=" & Val(s), gConSQL, adOpenStatic, adLockReadOnly
'                    Set adors = ObtenConsulta("select * from Responsables where id=" & Val(s))
'                    TreeView1.Nodes(s2).Text = TreeView1.Nodes(s2).Text + " Resp.: " + IIf(adors.RecordCount = 0, "Desconocido", adors!descripción) + ")"
'                    s = Mid(s, InStr(s, csSepara) + 2)  'Borra la clave del responsable
'                End If
'            Loop
'        Else
'            sActividades(6) = ""
'        End If
'    End If
'End If
'If TreeView3.Nodes.Count = 2 Then
'    If Not TreeView3.Nodes(1).Checked And Not TreeView3.Nodes(2).Checked And yTipoOperación <= 2 And ySoloConsulta = 0 Then
'        TreeView3.Nodes(2).Checked = True
'        TreeView3_NodeCheck TreeView3.Nodes(2)
'    End If
'End If
'MensajeTiempo ("Tiempo al cargar la ventana de Avances: ")
End Sub

Sub HabilitaAceptar(Optional bNoRevisaActividadProgramada As Boolean)
Dim Y As Byte, i As Long, s As String
If Not bOprimióTecla Then Exit Sub
If ySoloConsulta = 2 Then
    If cmdBotón(0).Enabled Then cmdBotón(0).Enabled = False
    Exit Sub
End If
ComboDesenlaces.Enabled = ComboDesenlaces.ListCount > 0
bOprimióTecla = False
For i = 1 To TreeView3.Nodes.Count
    If TreeView3.Nodes(i).Checked And (bNodoSeleccionado = True Or Not TreeView1.Enabled) Then
        Exit For
    End If
Next
If txtAcuerdo.Visible Then 'Verifica No. de Acuerdo capturado en caso que esté visible
    If Len(Trim(txtAcuerdo.Text)) = 0 And chkacuerdo.Value = 0 Then
        If cmdBotón(0).Enabled Then cmdBotón(0).Enabled = False
        Exit Sub
    End If
End If
If cmdSanción.Visible Then 'Verifica Datos de la sanción
    If InStr(msSanción, "|") = 0 Then
        If cmdBotón(0).Enabled Then cmdBotón(0).Enabled = False
        Exit Sub
        cmdSanción.BackColor = cnRojo
    Else
        cmdSanción.BackColor = cnVerde
    End If
End If
If cmdCondonación.Visible Then 'Verifica Datos de la sanción
    If InStr(msCondonación, "|") = 0 Then
        If cmdBotón(0).Enabled Then cmdBotón(0).Enabled = False
        Exit Sub
        cmdCondonación.BackColor = cnRojo
    Else
        cmdCondonación.BackColor = cnVerde
    End If
End If
cmdBotón(0).Enabled = (ComboDesenlaces.ListCount = 0 Or ComboDesenlaces.ListIndex >= 0) And (i <= TreeView3.Nodes.Count Or TreeView3.Nodes.Count = 0)
bOprimióTecla = True
bAceptar = False
End Sub

Function ActividadPadre(iAct As Integer) As Integer
Dim Y As Byte, rs As Recordset
Dim adors As New ADODB.Recordset
Do While True
    If adors.State > 0 Then adors.Close
    adors.Open "select id,nivel,idpad from actividades where id in (select idpad from actividades where id=" + Str(iAct) + ")", gConSql, adOpenStatic, adLockReadOnly
    If adors.RecordCount = 0 Then
        ActividadPadre = -1
        Exit Function
    End If
    Y = Y + 1
    iAct = adors(0)
    If adors(1) = 1 Or adors(2) = 0 Or Y > 200 Then Exit Do
Loop
ActividadPadre = iAct
End Function

Private Sub Form_Load()
Dim rs As Recordset, adors As New ADODB.Recordset
yUnico = 0
If Not gs = "no iniciar var" Then
    ySoloConsulta = 0
    bAceptar = False
    miActividad = 0 'Tiene el valor de idact de la actividad que se está registrando
    miTarea = 0 'Tiene el valor de idtar de la actividad que se está registrando
    mlAnálisis = 0 'Tiene el valor de idana del Oficio que se da seguimiento
    mlAnt = 0 'Tiene el valor de idant correspondiente al registro que se esta realizando
    mlSeguimiento = 0 'Tiene el valor de id del avance que se está editando
    miDesenlace = 0 'Valor de iddes
    miResponsable = 0 'Valor de idres
    msDoctos = "" 'Contiene el valos de los documentos emitidos en el avance
    msActsProg = "" 'Contiene el valos de las siguientes acts. programadas
    msObservaciones = "" 'limpia observaciones
End If
End Sub

Private Sub Frame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If txtEtiqueta.Visible Then txtEtiqueta.Visible = False
lSegundos = -1
End Sub

Private Sub HScroll1_Change()

End Sub

'Private Sub SftpClient_MessageLoop(StopAction As Boolean)
'  If TransferOperationActive Then
'    If SftpClient.CurrentOperationTotalInFile <> 0 Then
'        pbProgress.Value = 100 * SftpClient.CurrentOperationProcessedInFile / SftpClient.CurrentOperationTotalInFile
'    End If
'    lProgress.Caption = CStr(SftpClient.CurrentOperationProcessedInFile) & " / " & CStr(SftpClient.CurrentOperationTotalInFile)
'  End If
'  DoEvents
'  StopAction = False
'End Sub

'Private Sub SftpClient_OnAuthenticationAttempt(ByVal AuthType As Long, ByVal AuthParam As Variant)
'  Log "Trying authentication type " & CStr(AuthType), False
'End Sub
'
'Private Sub SftpClient_OnAuthenticationFailed(ByVal AuthenticationType As SSHBBoxCli7.TxSSHAuthenticationType)
'  Log "Authentication type " & CStr(AuthenticationType) & " failed", True
'End Sub
'
'Private Sub SftpClient_OnAuthenticationKeyboard(ByVal Prompts As BaseBBox7.IElStringListX, ByVal Echo As Variant, ByVal Responses As BaseBBox7.IElStringListX)
'    Dim i
'    Dim resp As String
'    For i = 0 To Prompts.Count - 1
'        resp = InputBox(Prompts.GetString(i), "Keyboard authentication", "")
'        Responses.Add (resp)
'    Next i
'End Sub
'
'Private Sub SftpClient_OnAuthenticationStart(ByVal SupportedAuths As Long)
'  Log "Authentication started", False
'End Sub
'
'Private Sub SftpClient_OnAuthenticationSuccess()
'  Log "Authentication succeeded", False
'End Sub
'
'Private Sub SftpClient_OnCloseConnection()
'  Log "SFTP connection closed", False
'  Disconnect
'End Sub
'
'Private Sub SftpClient_OnError(ByVal ErrorCode As Long)
'  Log "SSH error: " & CStr(ErrorCode), True
'  Disconnect
'End Sub
'
'Private Sub SftpClient_OnKeyValidate(ByVal ServerKey As SSHBBoxCli7.IElSSHKeyX, Valid As Boolean)
'  Log "Server key received", False
'  Valid = True
'End Sub
'
Private Sub Timer1_Timer()
Static yy As Byte, i As Long
Dim Y As Integer, l As Long
If yHabilita = 200 Then
    HabilitaAceptar
    yHabilita = 0
End If
'MDI.sb1.Panels(3).Style = sbrTime
If TreeView1.Nodes.Count > 0 Then
    If TreeView1.Nodes(1).Checked And TreeView1.Nodes(1).Children > 0 Then sQuitaNodo(0) = TreeView1.Nodes(1).Key + ","
End If
If TreeView2.Nodes.Count > 0 Then
    If TreeView2.Nodes(1).Checked And TreeView2.Nodes(1).Children > 0 Then sQuitaNodo(1) = TreeView2.Nodes(1).Key + ","
End If
If TreeView3.Nodes.Count > 0 Then
    If TreeView3.Nodes(1).Checked And TreeView3.Nodes(1).Children > 0 Then sQuitaNodo(2) = TreeView3.Nodes(1).Key + ","
    If Not bNodoSeleccionado Then
        For i = 1 To TreeView3.Nodes.Count
            If TreeView3.Nodes(i).Checked Then sQuitaNodo(2) = TreeView3.Nodes(i).Key + ","
        Next
    End If
End If
If Not bPrograma Then
    For Y = 2 To TreeView1.Nodes.Count
        If TreeView1.Nodes(Y).Checked And NodoContieneFecha(TreeView1.Nodes(Y)) = 0 Then
            TreeView1.Nodes(Y).Checked = False
        ElseIf Not TreeView1.Nodes(Y).Checked And NodoContieneFecha(TreeView1.Nodes(Y)) > 0 Then
            l = NodoContieneFecha(TreeView1.Nodes(Y))
            TreeView1.Nodes(Y).Text = Mid(TreeView1.Nodes(Y).Text, 1, l - 2)
        End If
    Next
    yy = 0
Else
    If yy > 1 Then bPrograma = False
    yy = yy + 1
End If
For Y = 0 To 2
    If Len(sQuitaNodo(Y)) > 0 Then
        Do While InStr(sQuitaNodo(Y), ",") > 0
            If Y = 0 Then
                TreeView1.Nodes(Mid(sQuitaNodo(Y), 1, InStr(sQuitaNodo(Y), ",") - 1)).Checked = False
            ElseIf Y = 1 Then
                TreeView2.Nodes(Mid(sQuitaNodo(Y), 1, InStr(sQuitaNodo(Y), ",") - 1)).Checked = False
            Else
                If TreeView3.Enabled Then TreeView3.Nodes(Mid(sQuitaNodo(Y), 1, InStr(sQuitaNodo(Y), ",") - 1)).Checked = False
            End If
            sQuitaNodo(Y) = Mid(sQuitaNodo(Y), InStr(sQuitaNodo(Y), ",") + 1)
        Loop
    End If
Next
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
Cancel = 1
End Sub

Private Sub TreeView1_Click()
bPrograma = True
End Sub

Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If txtEtiqueta.Visible Then txtEtiqueta.Visible = False
lSegundos = -1
End Sub

'programar actividad
Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim l As Long, yActividad As Byte, s As String, ss As String, Y As Integer, s1 As String, rs As Recordset
Dim adors As New ADODB.Recordset, b As Boolean, adors1 As New ADODB.Recordset, yProceso As Byte, i As Integer
Dim adors2 As New ADODB.Recordset
bOprimióTecla = True
yActividad = Val(txtRegistro)
bPrograma = True
If Node.Checked And Node.Children = 0 Then
    gs = ""
    For Y = 0 To Forms.Count - 1
        If Forms(Y).Name = "frmProgramaActividad" Then
            If Forms(Y).Tag = "Inválido" Then 'Para los formularios que tienen la prop.hide no deben considerarse
            Else
                Exit For
            End If
        End If
    Next
    s = ""
    If Y >= Forms.Count Then
        i = NodoContieneFecha(Node)
        If Node.Checked And i > 0 Then 'Obtiene la fecha programada y el idres
            s = Mid(Node.Text, InStrRev(Node.Text, " Resp.: ") + 8)
            If adors.State > 0 Then adors.Close
            adors.Open "select * from usuariossistema where descripción='" & Mid(s, 1, Len(s) - 1) + "'", gConSql, adOpenStatic, adLockReadOnly
            s = Mid(Node.Text, i + 1, InStrRev(Node.Text, " Resp.: ") - i - 1)
            If IsDate(s) Then
                s = Format(CDate(s), "dd/mm/yyyy hh:mm")
            Else
                s = ""
            End If
            If adors.RecordCount = 0 Then
            Else
                i = adors!ID
            End If
        Else
            If ComboResponsable.ListIndex >= 0 Then
                i = ComboResponsable.ItemData(ComboResponsable.ListIndex)
            End If
        End If
        With frmProgramaActividad
            glProceso = mlAnálisis
            gs2 = 0
            gs = "no iniciar var"
            gi1 = miTarea
            gi2 = Val(Right(Node.Key, 4))
            .iPlazoEstandar = 0
            .iPlazoMáximo = 0
            .iPlazoMínimo = 0
            .sProgramada = s
            .iResponsable = i
            .bDíasNaturales = False
            If IsDate(txtcampo(1).Text) Then
                If InStr(txtcampo(1).Text, " ") Then
                    .dInicio = CDate(Mid(txtcampo(1).Text, 1, InStr(txtcampo(1).Text, " ") - 1))
                Else
                    .dInicio = CDate(txtcampo(1).Text)
                End If
            End If
            .Caption = "Programación de " + Node.Text
            .Show vbModal
        End With
    End If
    If Len(gs) > 0 And Val(gs) > 0 Then
        l = NodoContieneFecha(Node)
        If adors.State > 0 Then adors.Close
        adors.Open "select * from usuariossistema where id=" & Mid(gs, InStr(gs, "|") + 1), gConSql, adOpenStatic, adLockReadOnly
        If l > 0 Then
            Node.Text = Mid(Node.Text, 1, l) + Mid(gs, 1, InStr(gs, "|") - 1) + "  Resp.: " + IIf(adors.RecordCount = 0, "Desconocido", adors!descripción) + ")"
        Else
            Node.Text = Node.Text + " (" + Mid(gs, 1, InStr(gs, "|") - 1) + "  Resp.: " + IIf(adors.RecordCount = 0, "Desconocido", adors!descripción) + ")"
        End If
        If bRR Then
            l = Node.Index
            For i = 1 To TreeView1.Nodes.Count
                If TreeView1.Nodes(i).Checked And TreeView1.Nodes(i).Index <> l Then
                    TreeView1.Nodes(i).Checked = False
                    l = NodoContieneFecha(TreeView1.Nodes(i))
                    If l > 0 Then
                        TreeView1.Nodes(i).Text = Mid(TreeView1.Nodes(i).Text, 1, l)
                    End If
                End If
            Next
        End If
        HabilitaAceptar (True)
    Else
        l = NodoContieneFecha(Node)
        If l = 0 Then sQuitaNodo(0) = sQuitaNodo(0) + Node.Key + ","
    End If
Else
    l = NodoContieneFecha(Node)
    If l > 0 Then Node.Text = Mid(Node.Text, 1, l - 2)
End If
bPrograma = False
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
Call TreeView1_NodeCheck(Node)
End Sub

Private Sub TreeView2_BeforeLabelEdit(Cancel As Integer)
Cancel = 1
End Sub

Private Sub TreeView2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If txtEtiqueta.Visible Then txtEtiqueta.Visible = False
lSegundos = -1
End Sub

Private Sub TreeView2_NodeCheck(ByVal Node As MSComctlLib.Node)
bOprimióTecla = True
End Sub

Private Sub TreeView3_BeforeLabelEdit(Cancel As Integer)
Cancel = 1
End Sub

Private Sub TreeView3_Click()
'Dim node1 As node
'If TreeView3.Nodes.Count > 0 Then
'    Set node1 = TreeView3.Nodes(1)
'    TreeView3_NodeClick node1
'End If
End Sub

Private Sub TreeView3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If txtEtiqueta.Visible Then txtEtiqueta.Visible = False
lSegundos = -1
End Sub

Private Sub TreeView3_NodeCheck(ByVal Node As MSComctlLib.Node)
Static yúnico As Byte 'No permite la entrada del evento recursivamente
Dim ss As String, Y As Integer, s As String, i As Long, yy As Byte, s2 As String
Dim y2 As Byte, y3 As Byte, adors As New ADODB.Recordset, l As Long, lAnt As Long
bOprimióTecla = True
If yúnico = 1 And yUnico = 0 Then yúnico = 0
If yúnico = 1 Then Exit Sub
miTarea = -1
If Node.Checked Then
    bNodoSeleccionado = True
Else
    For i = 1 To TreeView3.Nodes.Count
        If TreeView3.Nodes(i).Checked Then Exit For
    Next
    If i > TreeView3.Nodes.Count Then bNodoSeleccionado = False
End If
yúnico = 1
'sActividades(6) = ""
'For Y = 1 To TreeView1.Nodes.Count
'    i = NodoContieneFecha(TreeView1.Nodes(Y))
'    If TreeView1.Nodes(Y).Checked And i > 0 Then
'        s = Mid(TreeView1.Nodes(Y).Text, InStrRev(TreeView1.Nodes(Y).Text, " Resp.: ") + 8)
'        If adors.State > 0 Then adors.Close
'        Set adors = ObtenConsulta("select * from Responsables where descripción='" & Mid(s, 1, Len(s) - 1) + "'")
'        If adors.RecordCount = 0 Then
'            sActividades(6) = sActividades(6) + Right(TreeView1.Nodes(Y).Key, 4) + csSepara + Mid(TreeView1.Nodes(Y).Text, i + 1, InStrRev(TreeView1.Nodes(Y).Text, " Resp.: ") - i - 1) + csSepara + "null" + csSepara
'        Else
'            sActividades(6) = sActividades(6) + Right(TreeView1.Nodes(Y).Key, 4) + csSepara + Mid(TreeView1.Nodes(Y).Text, i + 1, InStrRev(TreeView1.Nodes(Y).Text, " Resp.: ") - i - 1) + csSepara + Trim(Str(adors!ID)) + csSepara
'        End If
'    End If
'Next
For Y = 1 To TreeView3.Nodes.Count
    If TreeView3.Nodes(Y).Checked And Node.Index <> TreeView3.Nodes(Y).Index Then
        TreeView3.Nodes(Y).Checked = False
    ElseIf TreeView3.Nodes(Y).Checked Then
        miTarea = Val(Right(TreeView3.Nodes(Y).Key, 4))
    End If
Next
'Actividades por programar
Call CargaActsProg(miTarea)
i = -200
If ComboDesenlaces.ListIndex >= 0 Then i = ComboDesenlaces.ItemData(ComboDesenlaces.ListIndex)
'Desenlaces
Call CargaDesenlaces(miTarea)
If i < -200 Then ComboDesenlaces.ListIndex = BuscaCombo(ComboDesenlaces, Trim(Str(i)), True)
If ComboDesenlaces.ListIndex < 0 And ComboDesenlaces.ListCount = 1 Then ComboDesenlaces.ListIndex = 0
'Documentos
Call CargaDocumentos(miTarea)
'No. Acuerdo
Call CargaAcuerdo(miTarea)
'Sanción
Call CargaSanción(miTarea)
'Condonación
Call CargaCondonación(miTarea)
'Subir Archivo vía FTP
Call CargaSubirArchivoFTP(miTarea)
'Subir Archivo vía FTP
Call CargaVerificaArchivoFTP(miTarea)
'No. Memo y fecha Memo
Call CargaMemo(miTarea)
'Verifica actividades prog Automáticamente
'Call VerificaActividadProgAut(miTarea)
If Node.Checked Then
    s = Mid(Node.Text, InStr(Node.Text, " ") + 1)
    If Len(s) > 40 Then
        s = Mid(s, 1, 40) & ".."
    End If
    etiTexto(1).Caption = "Fecha (" & Trim(s) & "):"
Else
    etiTexto(1).Caption = "Fecha:"
End If
yúnico = 0
HabilitaAceptar
End Sub

Private Sub TreeView3_NodeClick(ByVal Node As MSComctlLib.Node)
Call TreeView3_NodeCheck(Node)
End Sub

Private Sub txtAcuerdo_Change()
    HabilitaAceptar
End Sub

Private Sub txtAcuerdo_KeyPress(KeyAscii As Integer)
bOprimióTecla = True
End Sub

Private Sub txtCampo_DblClick(Index As Integer)
If InStr(" 1 2", Str(Index)) > 0 Then
    If Not IsDate(txtcampo(Index)) Then
        txtcampo(Index) = Format(AhoraServidor, gsFormatoFechaHora)
    ElseIf IsDate(txtcampo(Index)) Then
        If Hour(CDate(txtcampo(Index))) = 0 Then
            txtcampo(Index) = Format(Format(CDate(txtcampo(Index)), gsFormatoFecha) + " " + Format(Time, "hh:mm:ss"), gsFormatoFechaHora)
        End If
    End If
    HabilitaAceptar
End If
End Sub

Private Sub txtcampo_GotFocus(Index As Integer)
If txtEtiqueta.Visible Then txtEtiqueta.Visible = False
lSegundos = -1
If Index = 4 And txtcampo(Index).Visible Then
    txtcampo(4) = QuitaCadena(txtcampo(4).Text, "$, ")
End If
End Sub

Private Sub txtCampo_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = TeclaOprimida(txtcampo(Index), KeyAscii, txtcampo(Index).Tag, False)
bOprimióTecla = True
bAceptar = False
HabilitaAceptar
yHabilita = 200
End Sub

Private Sub txtCampo_LostFocus(Index As Integer)
Dim df As Date
If InStr(" 1", Str(Index)) > 0 Then
    Call ValidaFecha(txtcampo(Index), 1, Me.Name)
    If IsDate(txtcampo(Index)) Then
        If Hour(CDate(txtcampo(Index))) = 0 Then
            txtcampo(Index) = Format(Format(CDate(txtcampo(Index)), gsFormatoFecha) + " " + Format(Time, "hh:mm:ss"), gsFormatoFechaHora)
        End If
    End If
End If
If txtcampo(Index).Visible Then
    If txtcampo(Index).Tag = "m" Then
        txtcampo(Index).Text = Format(Val(QuitaCadena(txtcampo(Index).Text, "$, ")), "$###,###,###.00")
    ElseIf txtcampo(Index).Tag = "f" Then
        If IsDate(txtcampo(Index).Text) Then
            txtcampo(Index).Text = Format(CDate(txtcampo(Index).Text), gsFormatoFecha)
        End If
    ElseIf txtcampo(Index).Tag = "fh" Then
        If IsDate(txtcampo(Index).Text) Then
            txtcampo(Index).Text = Format(CDate(txtcampo(Index).Text), gsFormatoFechaHora)
        End If
    End If
End If
End Sub

Private Sub txtRegistro_GotFocus()
txtRegistro.Text = Val(txtRegistro)
End Sub

' Actualiza los datos en la pantalla para la actividad del arreglo sactividades()
' yActividad es el número de actividad
' El contenido de cada una de ellas es:
' 0: Descripción de la actividad
' 1: Fecha de Inicio
' 2: Fecha de Conclusión Misma que la anterior (obsoleto)
' 3: Observaciones del avance
Sub Actualiza()
Dim Y As Integer, yy As Byte, s As String, yError As Byte, rs As Recordset, i As Long, y2 As Byte, s2 As String, y3 As Byte
Dim ss As String, d As Date
Dim adors As New ADODB.Recordset
On Error GoTo ErrorActualiza:
    
    'Coloca datos de fecha y resp programados
    If Len(Trim(msProgResp)) > 0 Then
        txtFechaProgramada.Text = msProgResp
    Else
        txtFechaProgramada.Text = ""
    End If
    'd = AhoraServidor
    'txtCampo(1).Text = Format(d, gsFormatoFecha)
'    If miRespProg > 0 Then
'        If adors.State > 0 Then adors.Close
'        adors.Open "select f_responsable(" & miRespProg & ") from dual", gConSql, adOpenStatic, adLockReadOnly
'        If Not adors.EOF Then
'            ss = ss & IIf(IsNull(adors(0)), "", adors(0))
'        End If
'    End If
'    ss = ss & ")"
'
    'Coloca Responsable en su caso
    LlenaCombo ComboResponsable, "select id, descripción from usuariossistema where responsable<>0 and baja=0", "", True
    
'    If adors.State > 0 Then adors.Close
'    If miResponsable > 0 Then
'        ComboResponsable.ListIndex = BuscaCombo(ComboResponsable, miResponsable, True)
'    Else
'        adors.Open "select responsable from usuariossistema where id=" & giUsuario, gConSql, adOpenStatic, adLockReadOnly
'        If Not adors.EOF Then
'            If adors(0) > 0 Then
'                ComboResponsable.ListIndex = BuscaCombo(ComboResponsable, giUsuario, True)
'            End If
'        Else
'            If adors.State > 0 Then adors.Close
'            adors.Open "select idres from usuariossistema where id=" & giUsuario, gConSql, adOpenStatic, adLockReadOnly
'            If Not adors.EOF Then
'                If adors(0) > 0 Then
'                    ComboResponsable.ListIndex = BuscaCombo(ComboResponsable, adors(0), True)
'                End If
'            End If
'        End If
'    End If
    'Obtiene datos para las Tareas
    CargaTareas

    'Desenlaces
    'Call CargaDesenlaces(miTarea)
    'Actividades por programar
    Call CargaActsProg(miTarea)
    'Documentos
    Call CargaDocumentos(miTarea)

Exit Sub
ErrorActualiza:
If Err.Number = 35601 Then Resume Next
yError = MsgBox("Error: " + Err.Description, vbAbortRetryIgnore + vbCritical, "Error no esperado (" + Str(Err.Number) + ")")
If yError = vbCancel Then
    Exit Sub
ElseIf yError = vbRetry Then
    Resume
ElseIf yError = vbIgnore Then
    Resume Next
End If
End Sub

'Carga datos en el árbol3 Tareas
Private Sub CargaTareas()
Dim adors As New ADODB.Recordset, s As String, i As Byte, node1 As Node
adors.Open "select f_seguimiento_query(1," & mlAnálisis & "," & miActividad & " ) from dual", gConSql, adOpenStatic, adLockReadOnly
TreeView3.Nodes.Clear
If Not adors.EOF Then
    s = adors(0)
    i = Val(Mid(s, InStrRev(s, "|") + 1))
    s = Mid(s, 1, InStrRev(s, "|") - 1)
    Call CargaDatosArbolVariosNiveles(TreeView3, s, i, False, True)
End If
End Sub

'Carga datos en el árbol2 Documentos
Private Sub CargaDocumentos(iTarea As Integer)
Dim adors As New ADODB.Recordset, s As String, i As Byte
adors.Open "select f_seguimiento_query(4," & mlAnálisis & "," & iTarea & ") from dual", gConSql, adOpenStatic, adLockReadOnly
TreeView2.Nodes.Clear
If Not adors.EOF Then
    s = adors(0)
    i = Val(Mid(s, InStrRev(s, "|") + 1))
    s = Mid(s, 1, InStrRev(s, "|") - 1)
    Call CargaDatosArbolVariosNiveles(TreeView2, s, i, False, True)
End If
End Sub

'Carga desenlaces en el combo
Private Sub CargaDesenlaces(iTarea As Integer)
Dim adors As New ADODB.Recordset, s As String, i As Byte
adors.Open "select f_seguimiento_query(3," & mlAnálisis & "," & iTarea & ") from dual", gConSql, adOpenStatic, adLockReadOnly
ComboDesenlaces.Clear
If Not adors.EOF Then
    s = adors(0)
    i = Val(Mid(s, InStrRev(s, "|") + 1))
    s = Mid(s, 1, InStrRev(s, "|") - 1)
    Call LlenaCombo(ComboDesenlaces, s, "", True)
End If
End Sub

'Carga datos en el árbol1 Acts por programar
Private Sub CargaActsProg(iTarea As Integer)
Dim adors As New ADODB.Recordset, s As String, i As Byte, s1 As String, Y As Integer
Dim d As Date
adors.Open "select f_seguimiento_query(2," & mlAnálisis & "," & iTarea & ") from dual", gConSql, adOpenStatic, adLockReadOnly
TreeView1.Nodes.Clear
If Not adors.EOF Then
    s = adors(0)
    i = Val(Mid(s, InStrRev(s, "|") + 1))
    s = Mid(s, 1, InStrRev(s, "|") - 1)
        Call CargaDatosArbolVariosNiveles(TreeView1, s, i, False, True)
End If
'Programa Actividades Automáticamente
If adors.State Then adors.Close
adors.Open "select idact, diasprog, diash from relacióntareaactividad where idtar=" & iTarea & " and progaut <> 0", gConSql, adOpenForwardOnly, adLockReadOnly
Do While Not adors.EOF
    For Y = 1 To TreeView1.Nodes.Count
        If Val(Right(TreeView1.Nodes(Y).Key, 4)) = adors(0) And ComboResponsable.ListIndex >= 0 And IsDate(txtcampo(1).Text) Then 'Busca la actividad en el arbol de acts Prog
            d = CDate(txtcampo(1).Text)
            If adors(2) <> 0 Then
                d = DíasHábiles(d, adors(1))
            Else
                d = d + adors(1)
            End If
            ComboResponsable.ListIndex = ComboResponsable.ListIndex
            ComboResponsable.Refresh
            s = TreeView1.Nodes(Y).Text & "(" & Format(d, "dd/mm/yyyy hh:mm") & " Resp.: " & ComboResponsable.Text & ")"
            TreeView1.Nodes(Y).Text = s
            TreeView1.Nodes(Y).Checked = True
            'msActsProg = msActsProg & Right(TreeView1.Nodes(y).Key, 4) & "|" & s & "|" & Trim(Str(adors!ID)) & "|"
        End If
    Next
    adors.MoveNext
Loop
End Sub


'Carga Acuerdo en caso de que exista la propiedad de obtener Acuerdo en la tabla  RelaciónActividadTarea
Private Sub CargaAcuerdo(iTarea As Integer)
Dim adors As New ADODB.Recordset, s As String, i As Integer
adors.Open "select count(*) from relaciónactividadtarea where idact=" & miActividad & " and idtar=" & iTarea & " and idotr=1", gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    If Not txtAcuerdo.Visible Then
        txtAcuerdo.Visible = True
        EtiAcuerdo.Visible = True
        chkacuerdo.Visible = True
    End If
    If adors.State Then adors.Close
    adors.Open "select to_char(sysdate,'yyyy') from dual", gConSql, adOpenStatic, adLockReadOnly
    If Not IsNull(adors(0)) Then
        i = adors(0)
    Else
        i = Year(Date)
    End If
    If Len(Trim(txtAcuerdo.Text)) = 0 And yTipoOperación <> 0 Then
        'txtAcuerdo.Text = "ACUERDO/DAS/" & i & "/"
        txtAcuerdo.Text = ""
        'txtAcuerdo.Text = "AUTOMÁTICO"
    End If
    Exit Sub
End If
If txtAcuerdo.Visible Then
    txtAcuerdo.Visible = False
    EtiAcuerdo.Visible = False
    chkacuerdo.Visible = False
End If
End Sub

'Carga datos de la sanción
Private Sub CargaSanción(iTarea As Integer)
Dim adors As New ADODB.Recordset, s As String, i As Byte
adors.Open "select count(*) from relaciónactividadtarea where idact=" & miActividad & " and idtar=" & iTarea & " and idotr=2", gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    If Not cmdSanción.Visible Then
        cmdSanción.Visible = True
    End If
    Exit Sub
End If
If cmdSanción.Visible Then
    cmdSanción.Visible = False
End If
End Sub

'Carga datos de la Condonación
Private Sub CargaCondonación(iTarea As Integer)
Dim adors As New ADODB.Recordset, s As String, i As Byte
adors.Open "select count(*) from relaciónactividadtarea where idact=" & miActividad & " and idtar=" & iTarea & " and idotr=6", gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    If Not cmdCondonación.Visible Then
        cmdCondonación.Visible = True
    End If
    Exit Sub
End If
If cmdCondonación.Visible Then
    cmdCondonación.Visible = False
End If
End Sub

'Carga información del archivo por subir vía FTP a estrados electrónicos
Private Sub CargaSubirArchivoFTP(iTarea As Integer)
Dim adors As New ADODB.Recordset, s As String, i As Byte
adors.Open "select count(*) from relaciónactividadtarea where idact=" & miActividad & " and idtar=" & iTarea & " and idotr=3", gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    If Not cmdSubirDocto.Visible Then
        cmdSubirDocto.Visible = True
    End If
    cmdVerificaDocto.Visible = True
    sArchivoFTP = ""
    yVerificaDocto = 0
    Exit Sub
End If
If cmdSubirDocto.Visible Then
    cmdSubirDocto.Visible = False
    cmdVerificaDocto.Visible = False
End If
End Sub

'Carga información del archivo por subir vía FTP a estrados electrónicos
Private Sub CargaVerificaArchivoFTP(iTarea As Integer)
Dim adors As New ADODB.Recordset, s As String, i As Byte
adors.Open "select count(*) from relaciónactividadtarea where idact=" & miActividad & " and idtar=" & iTarea & " and idotr in (3,4)", gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    If Not cmdVerificaDocto.Visible Then
        cmdVerificaDocto.Visible = True
    End If
    If adors.State Then adors.Close
    adors.Open "select idpro from tareas where id=" & miTarea, gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        If adors(0) = 4 Then 'Sanciones Busca el oficio de sanción
            If adors.State Then adors.Close
            adors.Open "select f_sancion_oficio(" & mlAnálisis & "," & mlAnt & ") from dual", gConSql, adOpenStatic, adLockReadOnly
        Else 'Emplazamiento Busca el oficio en análisis
            If adors.State Then adors.Close
            adors.Open "select f_analisis_oficio(" & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
        End If
        If Not adors.EOF Then
            If Not IsNull(adors(0)) Then
                sArchivoFTP = Replace(adors(0), "/", "_") & ".pdf"
            End If
        End If
    End If
    yVerificaDocto = 0
    Exit Sub
End If
If cmdVerificaDocto.Visible Then
    cmdVerificaDocto.Visible = False
End If
End Sub

'Carga Memo y fecha Memo cuando en la tabla RelaciónActividadTarea idotr=5
Private Sub CargaMemo(iTarea As Integer)
Dim adors As New ADODB.Recordset, s As String, i As Integer
adors.Open "select count(*) from relaciónactividadtarea where idact=" & miActividad & " and idtar=" & iTarea & " and idotr=5", gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    If Not txtAcuerdo.Visible Then
        txtcampo(3).Visible = True
        etiTexto(3).Visible = True
        txtcampo(4).Visible = True
        etiTexto(4).Visible = True
    End If
    If adors.State Then adors.Close
    adors.Open "select to_char(sysdate,'yyyy') from dual", gConSql, adOpenStatic, adLockReadOnly
    If Not IsNull(adors(0)) Then
        i = adors(0)
    Else
        i = Year(Date)
    End If
    If Len(Trim(txtAcuerdo.Text)) = 0 And yTipoOperación <> 0 Then
        txtAcuerdo.Text = "MEMORANDO/DAS/" & i & "/"
    End If
    Exit Sub
End If
If txtcampo(3).Visible Then
    txtcampo(3).Visible = False
    etiTexto(3).Visible = False
    txtcampo(4).Visible = False
    etiTexto(4).Visible = False
End If
End Sub


'Verifica actividades programadas aut.
Private Sub VerificaActividadProgAut(iTarea As Integer)
Dim adors As New ADODB.Recordset, s As String, i As Byte
adors.Open "select count(*) from relaciónactividadtarea where idact=" & miActividad & " and idtar=" & iTarea & " and idotr=2", gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    If Not cmdSanción.Visible Then
        cmdSanción.Visible = True
        Exit Sub
    End If
End If
If cmdSanción.Visible Then
    cmdSanción.Visible = False
End If
End Sub


Sub Inicia()
Dim adors As ADODB.Recordset, i As Integer
mlAnálisis = -1
mlSeguimiento = -1
yTipoOperación = 0
bAceptar = False
Call LlenaCombo(ComboResponsable, "usuariossistema", "baja=0")
End Sub


Function nodo(ByRef Tv As TreeView, sNodo) As Integer
Dim i As Long
For i = 1 To Tv.Nodes.Count
    If Tv.Nodes(i).Key = sNodo Then
        nodo = i
        Exit Function
    End If
Next
nodo = 0
End Function

Sub AbortaProceso(sLinea As String)
'gwsUnico.Rollback
End Sub


Function SubeFTP(sArchivo As String) As Boolean
Dim db As DAO.Database, yIntentosHost As Byte
Dim Y As Byte
Dim strURL As String      ' URL string
Dim bData() As Byte      ' Data variable
Dim intFile As Integer   ' FreeFile variable
Dim f
Dim l As Long
Dim s1 As String
Dim s As String, ss As String, s2 As String
On Error GoTo ERRORCOMUNICACIÓN:

ss = sArchivo
s2 = "c:\" & Mid(ss, InStrRev(ss, "\") + 1)
s2 = Replace(s2, " ", "")
If Len(Dir(sArchivo)) Then 'Mueve archivo a raiz de c:\
    If Len(Dir(s2)) Then Kill s2
    Name ss As s2
End If
If Len(Dir(s2)) = 0 Then
    MsgBox "No se logró mover archivo a c:\ (" & s2 & ")", vbOKOnly + vbInformation, ""
    Exit Function
End If

If InStr(s2, "\") Then
    s = Mid(s2, InStrRev(s2, "\") + 1)
Else
    s = s2
End If
'Set Inet1 = frm.Inet1
sHostRemoto = "ftp://sioenvio:510sio@192.168.10.170"
'sHostRemoto = "148.235.190.170"
'Inet1.Execute "ftp://sioenvio:510sio@" & sHostRemoto, "SEND " & sArchivo & " " & s

'''Inet1.Execute sHostRemoto, "SEND " & s2 & " " & s
'''Do While Inet1.StillExecuting
'''    DoEvents
'''Loop

'Inet1.Execute "ftp://sioenvio:510sio@192.168.10.170", "SIZE " & s
's1 = Inet1.ResponseInfo


'''Inet1.Execute sHostRemoto, "DIR " & s
'''s1 = Inet1.ResponseInfo
If InStr(LCase(s1), "no hay") = 0 And Len(s1) > 0 Then
    'OK
    SubeFTP = True
Else
    MsgBox "No está cargando el Archivo: " & s & " en el sitio ftp", vbCritical, ""
    SubeFTP = False
End If

Name s2 As ss
If Len(Dir(ss)) = 0 Then
    MsgBox "No se logró regresar archivo a su lugar de origen (" & ss & ")", vbOKOnly + vbInformation, ""
    Exit Function
End If

Exit Function
ERRORCOMUNICACIÓN:
If (Err.Number = 35754 Or Err.Number = 35761) And yIntentosHost < 2 Then  'intenta con el otro ip
    If yIntentosHost = 0 Then
        'sHostRemoto = "148.235.190.170"
        sHostRemoto = "192.168.10.170"
    Else
        sHostRemoto = "central.condusef.gob.mx"
    End If
    yIntentosHost = yIntentosHost + 1
    Resume
End If
'ErrorBase:
Y = MsgBox(Err.Description, vbRetryCancel, "")
If Y = vbRetry Then
    Resume
ElseIf Y = vbIgnore Then
    Resume Next
End If
Error (Err.Number & ":" & Err.Description)

End Function

'Function SubeFTP2(sArchivo As String) As Boolean
'Dim db As DAO.Database, yIntentosHost As Byte
'Dim Y As Byte
'Dim strURL As String      ' URL string
'Dim bData() As Byte      ' Data variable
'Dim intFile As Integer   ' FreeFile variable
'Dim f
'Dim l As Long
'Dim s1 As String
'Dim s As String, ss As String, s2 As String
'
''Dim sftp As New ChilkatSFtp
''  Any string automatically begins a fully-functional 30-day trial.
'Dim success As Long
'success = sftp.UnlockComponent("Anything for 30-day trial")
'If (success <> 1) Then
'    MsgBox sftp.LastErrorText
'    Exit Function
'End If
'
''  Set some timeouts, in milliseconds:
'sftp.ConnectTimeoutMs = 5000
'sftp.IdleTimeoutMs = 10000
'
''  Connect to the SSH server.
''  The standard SSH port = 22
''  The hostname may be a hostname or IP address.
'Dim port As Long
'Dim hostname As String
'hostname = "192.168.10.12"
'port = 22
'success = sftp.Connect(hostname, port)
'If (success <> 1) Then
'    MsgBox sftp.LastErrorText
'    Exit Function
'End If
'
''  Authenticate with the SSH server.  Chilkat SFTP supports
''  both password-based authenication as well as public-key
''  authentication.  This example uses password authenication.
'success = sftp.AuthenticatePw("estrados", "3str4d0s")
'If (success <> 1) Then
'    MsgBox sftp.LastErrorText
'    Exit Function
'End If
'
''  After authenticating, the SFTP subsystem must be initialized:
'success = sftp.InitializeSftp()
'If (success <> 1) Then
'    MsgBox sftp.LastErrorText
'    Exit Function
'End If
'
'
'On Error GoTo ERRORCOMUNICACIÓN:
'
'ss = sArchivo
's2 = "c:\" & Mid(ss, InStrRev(ss, "\") + 1)
's2 = Replace(s2, " ", "")
'If Len(Dir(sArchivo)) Then 'Mueve archivo a raiz de c:\
'    If Len(Dir(s2)) Then Kill s2
'    Name ss As s2
'End If
'If Len(Dir(s2)) = 0 Then
'    MsgBox "No se logró mover archivo a c:\ (" & s2 & ")", vbOKOnly + vbInformation, ""
'    Exit Function
'End If
'
'If InStr(s2, "\") Then
'    s = Mid(s2, InStrRev(s2, "\") + 1)
'Else
'    s = s2
'End If
'
''  Open a file for writing on the SSH server.
''  If the file already exists, it is overwritten.
''  (Specify "createNew" instead of "createTruncate" to
''  prevent overwriting existing files.)
'Dim handle As String
'handle = sftp.OpenFile(s, "writeOnly", "createTruncate")
'If (handle = vbNullString) Then
'    MsgBox sftp.LastErrorText
'    Exit Function
'End If
'
'
'
''  Upload from the local file to the SSH server.
'success = sftp.UploadFile(handle, s2)
'If (success <> 1) Then
'    MsgBox sftp.LastErrorText
'    Exit Function
'End If
'
''  Close the file.
'success = sftp.CloseHandle(handle)
'If (success <> 1) Then
'    MsgBox sftp.LastErrorText
'    Exit Function
'End If
'
'
'Name s2 As ss
'If Len(Dir(ss)) = 0 Then
'    MsgBox "No se logró regresar archivo a su lugar de origen (" & ss & ")", vbOKOnly + vbInformation, ""
'    Exit Function
'End If
'sArchivoFTP = s
'SubeFTP2 = True
'
'
'Exit Function
'ERRORCOMUNICACIÓN:
'If (Err.Number = 35754 Or Err.Number = 35761) And yIntentosHost < 2 Then  'intenta con el otro ip
'    If yIntentosHost = 0 Then
'        'sHostRemoto = "148.235.190.170"
'        sHostRemoto = "192.168.10.170"
'    Else
'        sHostRemoto = "central.condusef.gob.mx"
'    End If
'    yIntentosHost = yIntentosHost + 1
'    Resume
'End If
''ErrorBase:
'Y = MsgBox(Err.Description, vbRetryCancel, "")
'If Y = vbRetry Then
'    Resume
'ElseIf Y = vbIgnore Then
'    Resume Next
'End If
'Error (Err.Number & ":" & Err.Description)
'
'End Function

''Utilizando nueva librería
'Function SubeFTP3(sArchivo As String) As Boolean
'Dim db As DAO.Database, yIntentosHost As Byte
'Dim Y As Byte
'Dim strURL As String      ' URL string
'Dim bData() As Byte      ' Data variable
'Dim intFile As Integer   ' FreeFile variable
'Dim f
'Dim l As Long
'Dim s1 As String
'Dim s As String, ss As String, s2 As String
'
''  Any string automatically begins a fully-functional 30-day trial.
'
''  Connect to the SSH server.
''  The standard SSH port = 22
''  The hostname may be a hostname or IP address.
'Dim port As Long
'Dim hostname As String, usr As String, pwd As String
'hostname = "192.168.10.12"
'port = 22
'usr = "estrados"
'pwd = "3str4d0s"
'
'ConnectSFTP hostname, port, usr, pwd
'
'
'On Error GoTo ERRORCOMUNICACIÓN:
'
'ss = sArchivo
's2 = "c:\" & Mid(ss, InStrRev(ss, "\") + 1)
's2 = Replace(s2, " ", "")
'If Len(Dir(sArchivo)) Then 'Mueve archivo a raiz de c:\
'    If Len(Dir(s2)) Then Kill s2
'    Name ss As s2
'End If
'If Len(Dir(s2)) = 0 Then
'    MsgBox "No se logró mover archivo a c:\ (" & s2 & ")", vbOKOnly + vbInformation, ""
'    Exit Function
'End If
'
'If InStr(s2, "\") Then
'    s = Mid(s2, InStrRev(s2, "\") + 1)
'Else
'    s = s2
'End If
'
''Sube el documento
'
'Call Upload(s2, s)
'
'
'Name s2 As ss
'If Len(Dir(ss)) = 0 Then
'    MsgBox "No se logró regresar archivo a su lugar de origen (" & ss & ")", vbOKOnly + vbInformation, ""
'    Exit Function
'End If
'sArchivoFTP = s
'SubeFTP3 = True
'
'
'Exit Function
'ERRORCOMUNICACIÓN:
'If (Err.Number = 35754 Or Err.Number = 35761) And yIntentosHost < 2 Then  'intenta con el otro ip
'    If yIntentosHost = 0 Then
'        'sHostRemoto = "148.235.190.170"
'        sHostRemoto = "192.168.10.170"
'    Else
'        sHostRemoto = "central.condusef.gob.mx"
'    End If
'    yIntentosHost = yIntentosHost + 1
'    Resume
'End If
''ErrorBase:
'Y = MsgBox(Err.Description, vbRetryCancel, "")
'If Y = vbRetry Then
'    Resume
'ElseIf Y = vbIgnore Then
'    Resume Next
'End If
'Error (Err.Number & ":" & Err.Description)
'
'End Function
'
'Private Sub ConnectSFTP(hostname As String, port As Long, usr As String, pwd As String)
'  'frmConnProps.Show vbModal, Me
'  'If frmConnProps.Executed Then
'    SftpClient.UserName = usr
'    SftpClient.Password = pwd
'    SftpClient.EnableAuthenticationType SSH_AUTH_TYPE_PASSWORD
'    SftpClient.EnableAuthenticationType SSH_AUTH_TYPE_KEYBOARD
'
'    ' Optionally load the private key
'    SftpClient.KeyStorage = KeyStorage.object
'    Dim Key As IElSSHKeyX
'    Dim nPrivateKeyLoadingError As Integer
'    Call KeyStorage.Clear
'    Set Key = CreateObject("SSHBBoxCli7.ElSSHKeyX")
'    nPrivateKeyLoadingError = -1
'
'    ' load the key from the file, if the file has been specified
'    'If frmConnProps.editPrivateKeyFile.Text <> "" Then
'    '  On Error GoTo LoadFailed
'    '  Key.LoadPrivateKey frmConnProps.editPrivateKeyFile.Text, frmConnProps.edtKeyPassword.Text
'    '  nPrivateKeyLoadingError = 0
'LoadFailed:
'    'End If
'
'    If nPrivateKeyLoadingError = 0 Then
'      Call KeyStorage.Add(Key)
'      SftpClient.EnableAuthenticationType SSH_AUTH_TYPE_PUBLICKEY
'    Else
'      SftpClient.DisableAuthenticationType SSH_AUTH_TYPE_PUBLICKEY
'    End If
'    Set Key = Nothing
'
'    ' Initiate connection
'    lvLog.ListItems.Clear
'    Log "Connecting to " & hostname, False
'     SftpClient.Address = hostname
'     SftpClient.port = port
'     On Error GoTo HandleErr
'     SftpClient.Open
'     If SftpClient.Active Then
'       'RefreshData
'     End If
'  'End If
'     Exit Sub
'HandleErr:
'    Call Log("Error: " & Err.Description, True)
'    Call Log("If you have ensured that all connection parameters are correct and you still can't connect,", True)
'    Call Log("please contact EldoS support as described on http://www.eldos.com/sbb/support.php", True)
'    Call Log("Remember to provide details about the error that happened.", True)
'    If Len(SftpClient.ServerSoftwareName) > 0 Then
'        Call Log("Server software identified itself as: " + SftpClient.ServerSoftwareName, True)
'    End If
'End Sub
'
'Private Sub Upload(sOrigen As String, sDestino As String)
'Dim shortName As String
'Dim Size As Long
'Dim i As Integer
'
''Dim dlgProgress As New frmProgress
'  If SftpClient.Active Then
'
''    CommonDialog.DialogTitle = "Upload"
''    CommonDialog.FileName = ""
''    On Error Resume Next
''    CommonDialog.ShowOpen
''    i = Err
''    Err.Clear
''    On Error GoTo 0
''    If i = 32755 Then
''      Exit Sub
''    End If
''
''    On Error GoTo HandleErr
''    Call Log("Uploading file " & CommonDialog.FileName, False)
''    shortName = ExtractFileName(CommonDialog.FileName)
''
''    Dim RemoteName As String
'
'
'
'    Canceled = False
'
'    'dlgProgress.Canceled = False
'    'dlgProgress.Caption = "Transferencia"
'    'frmProgress.Show
'    lvLog.Visible = False
'    frProgress.Visible = True
'    pbProgress.Value = 0
'    lSourceFilename.Caption = sOrigen
'    lDestFilename.Caption = sDestino
'    lProgress.Caption = "0 / 0"
'    'dlgProgress.Show vbModal
'    lvLog.Refresh
'    frProgress.Refresh
'
'    TransferOperationActive = True
'    Call SftpClient.UploadFile(sOrigen, sDestino)
'    TransferOperationActive = False
'    frProgress.Visible = False
'
'    ' Adjust attributes for the remote file
'    Dim Attrs As New ElSftpFileAttributesX
'    Attrs.CTime = Date
'    Attrs.ATime = Attrs.CTime
'    Attrs.MTime = Attrs.CTime
'    Attrs.IncludeAttribute SB_SFTP_ATTR_ATIME
'    Attrs.IncludeAttribute SB_SFTP_ATTR_CTIME
'    Attrs.IncludeAttribute SB_SFTP_ATTR_MTIME
'
'    Call SftpClient.SetAttributes(sDestino, Attrs)
'
'    Call Log("Upload finished", False)
'    'Call RefreshData
'  End If
'
'  Exit Sub
'HandleErr:
'  TransferOperationActive = False
'  If frProgress.Visible Then
'     frProgress.Visible = False
'  End If
'  Call Log("Error: " & Err.Description, True)
'End Sub
'
'Private Sub Log(s As String, IsError As Boolean)
'Dim Item As Object
'  If Not lvLog.Visible Then
'    lvLog.Visible = True
'    Me.Height = 9135
'    Me.Refresh
'  End If
'  Set Item = lvLog.ListItems.Add
'  Item.Text = Time
'  Let Item.SubItems(1) = s
'  'If Not IsError Then
'  '  item.SmallIcon = 1
'  'Else
'  '  item.SmallIcon = 2
'  'End If
'End Sub
'
'Private Sub Disconnect()
'  Log "Disconnecting", False
'  If SftpClient.Active Then
'    SftpClient.Close
'  End If
'End Sub
'

