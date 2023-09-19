VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Publicaciones 
   Caption         =   "Publicaciones"
   ClientHeight    =   8304
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   12732
   LinkTopic       =   "Form1"
   ScaleHeight     =   8304
   ScaleWidth      =   12732
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4248
      Left            =   0
      TabIndex        =   13
      Top             =   4068
      Width           =   12660
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         Height          =   3690
         Left            =   11340
         TabIndex        =   14
         Top             =   420
         Width           =   1320
         Begin VB.CommandButton cmdProceso 
            Caption         =   "&Agregar Oficio"
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   48
            TabIndex        =   18
            Top             =   192
            Width           =   1200
         End
         Begin VB.CommandButton cmdProceso 
            Caption         =   "&Preparar Publicación"
            Enabled         =   0   'False
            Height          =   1164
            Index           =   2
            Left            =   48
            Picture         =   "PublicacionesBK.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   1452
            Width           =   1200
         End
         Begin VB.CommandButton cmdProceso 
            Caption         =   "&Quitar Oficio"
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   48
            TabIndex        =   15
            Top             =   816
            Width           =   1200
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3720
         Left            =   72
         TabIndex        =   17
         Top             =   396
         Width           =   11244
         _ExtentX        =   19833
         _ExtentY        =   6562
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "IF"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Oficio"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Ley / Causa"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Monto"
            Object.Width           =   11466
         EndProperty
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Actividades Realizadas:"
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   3
         Left            =   72
         TabIndex        =   19
         Top             =   180
         Width           =   1692
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   2628
      Left            =   1665
      TabIndex        =   8
      Top             =   1404
      Width           =   11040
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "Nombre"
         ForeColor       =   &H00808080&
         Height          =   288
         Index           =   4
         Left            =   108
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Tag             =   "c"
         ToolTipText     =   "Datos capturados del Oficio en la etapa de Análisis"
         Top             =   2280
         Width           =   2376
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "Nombre"
         ForeColor       =   &H00808080&
         Height          =   288
         Index           =   3
         Left            =   108
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Tag             =   "c"
         ToolTipText     =   "Datos capturados del Oficio en la etapa de Análisis"
         Top             =   1812
         Width           =   2376
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "Nombre"
         ForeColor       =   &H00808080&
         Height          =   690
         Index           =   5
         Left            =   3780
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Tag             =   "c"
         ToolTipText     =   "Datos capturados del Oficio en la etapa de Análisis"
         Top             =   1812
         Width           =   7200
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "Nombre"
         ForeColor       =   &H00808080&
         Height          =   1095
         Index           =   2
         Left            =   7785
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Tag             =   "c"
         ToolTipText     =   "Datos del documento de Solicitud"
         Top             =   495
         Width           =   3200
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "Nombre"
         ForeColor       =   &H00808080&
         Height          =   1095
         Index           =   1
         Left            =   3780
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Tag             =   "c"
         ToolTipText     =   "Nombre de la Institución y del Usuario"
         Top             =   495
         Width           =   4000
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "Nombre"
         ForeColor       =   &H00808080&
         Height          =   1095
         Index           =   0
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Tag             =   "c"
         ToolTipText     =   "Datos del origen de la Solicitud"
         Top             =   495
         Width           =   3700
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Monto total de la Sanción:"
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   5
         Left            =   108
         TabIndex        =   25
         Top             =   2088
         Width           =   1836
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha de la Sanción:"
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   4
         Left            =   144
         TabIndex        =   23
         Top             =   1620
         Width           =   1512
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ley / Causa / Monto:"
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   3
         Left            =   3816
         TabIndex        =   21
         Top             =   1620
         Width           =   1452
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Documento de la Solicitud:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   7830
         TabIndex        =   11
         Top             =   270
         Width           =   1905
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Institución / Nombre(s):"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   3780
         TabIndex        =   10
         Top             =   270
         Width           =   1650
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Origen de la Solicitud:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   270
         Width           =   1545
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   1464
      Left            =   1665
      TabIndex        =   5
      Top             =   -45
      Width           =   11040
      Begin VB.TextBox txtOficio 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   2580
         MaxLength       =   80
         TabIndex        =   12
         Tag             =   "c"
         ToolTipText     =   "No. de Oficio a realizar seguimiento"
         Top             =   945
         Width           =   3552
      End
      Begin VB.CommandButton cmdActualpen 
         Caption         =   "Nueva &Consulta"
         Height          =   420
         Left            =   8388
         TabIndex        =   0
         Top             =   828
         Width           =   1590
      End
      Begin VB.CommandButton cmdContinuar 
         BackColor       =   &H00008000&
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6480
         Picture         =   "PublicacionesBK.frx":05E6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   864
         Width           =   1500
      End
      Begin ComctlLib.ImageList Imagenes 
         Left            =   7920
         Top             =   180
         _ExtentX        =   995
         _ExtentY        =   995
         BackColor       =   -2147483643
         ImageWidth      =   103
         ImageHeight     =   104
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   2
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PublicacionesBK.frx":1155
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PublicacionesBK.frx":9067
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Número de Oficio de Sanción:"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   324
         TabIndex        =   7
         Top             =   972
         Width           =   2172
      End
      Begin VB.Label Eti 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Módulo de Publicaciones"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   14.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   348
         Index           =   2
         Left            =   336
         TabIndex        =   6
         Top             =   180
         Width           =   8052
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Image Image1 
      Height          =   3972
      Left            =   0
      Picture         =   "PublicacionesBK.frx":11969
      Stretch         =   -1  'True
      Top             =   48
      Width           =   1668
   End
End
Attribute VB_Name = "Publicaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mlSeguimiento As Long 'Contiene el id del Seguimiento del oficio buscado
Dim mlAsunto As Long 'Contiene el id del Asunto registrado en registro
Dim msLeyes As String 'Contine los id de las leyes
Dim msCausas As String 'Contine los id de las causas
Dim mlAsuxIF As Long 'Contiene id de la ifregxif selecionado
Dim mlAnálisis As Long 'Contiene id del análisis
Dim mlAnálisisImp As Long 'Contiene id del análisis
Dim msLeyesImp As String 'Contine los id de las leyes
Dim msCausasIMP As String 'Contine los id de las causas
Dim msMotivosImp As String 'Contine los id de los motivos de improcedencia
Dim msTodasCausasImp As String
Dim mbLimpiaExp As Boolean 'indicador para limpiar el campo o lista de exp Pendientes
Dim msActs As String


Private Sub cmbCampo_Click(Index As Integer)
Dim i As Long, adors As New ADODB.Recordset
If Index = 3 Then 'SINE
    If ListView1.SelectedItem.Index <= 0 Then
        cmdProceso(3).Visible = False
        Exit Sub
    End If
    If adors.State Then adors.Close
    adors.Open "select f_sine_ifactiva(f_analisis_idins(" & mlAnálisis & "),f_act_idpro(f_seguimiento_idact(" & ListView1.ListItems(ListView1.SelectedItem.Index).Tag & "))) from dual", gConSql, adOpenStatic, adLockReadOnly
    If adors(0) <= 0 Then
        Call MsgBox("La IF de este asunto no tiene convenio de notificaiones electrónicas", vbOKOnly + vbInformation, "Validación")
        cmdProceso(3).Visible = False
        Exit Sub
    End If
    If adors.State Then adors.Close
    adors.Open "select f_url_conexion(1," & giUsuario & "," & ListView1.ListItems(ListView1.SelectedItem.Index).Tag & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    If Len(adors(0)) > 0 Then
        gsWWW = adors(0)
    Else
        MsgBox "La cadena de conexión al SINE viene vacia Favor de reportarlo al administrador del Sistema  (Miguel Ext.6032)"
        Exit Sub
    End If
    gsWWW = gsWWW & "&perm=0"
    With Browser
        .yÚnicavez = 0
        .Show vbModal
    End With

ElseIf Index = 0 And cmbCampo(Index).ListIndex >= 0 Then 'Refresca datos de combo Oficios
    mlAsuxIF = cmbCampo(Index).ItemData(cmbCampo(Index).ListIndex)
'    If adors.State Then adors.Close
'    adors.Open "{call p_seguimiento_datosoficio()}", gConSql, adOpenStatic, adLockReadOnly
    If adors.State Then adors.Close
    adors.Open "select id, oficio||case when procedente=1 then '' else ' (Improc.)' end as descripción from análisis where idregxif=" & mlAsuxIF, gConSql, adOpenStatic, adLockReadOnly
    cmbCampo(1).Clear
    Do While Not adors.EOF
        cmbCampo(1).AddItem adors(1)
        cmbCampo(1).ItemData(cmbCampo(1).NewIndex) = adors(0)
        adors.MoveNext
    Loop
    LimpiaControlesInfConsulta 1
    If cmbCampo(1).ListCount = 1 Then
        cmbCampo(1).ListIndex = 0
    End If
ElseIf Index = 1 Then  'filtra pendientes
    If cmbCampo(Index).ListIndex >= 0 Then
        mlAnálisis = cmbCampo(Index).ItemData(cmbCampo(Index).ListIndex)
        If adors.State Then adors.Close
        adors.Open "select oficio||' ('||to_char(fecha,'dd/mon/yyyy')||')'||chr(13)||chr(10)||case when procedente=1 then f_causas(id) else f_causasimp(id) end as descripción from análisis where id=" & mlAnálisis, gConSql, adOpenForwardOnly, adLockReadOnly
        If Not adors.EOF Then
            txtCampo(3).Text = IIf(IsNull(adors(0)), "", adors(0))
        End If

        
        If adors.State Then adors.Close
        adors.Open "{call p_seguimientoActspen(" & mlAnálisis & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
        cmbCampo(2).Clear
        If adors.EOF Then 'No hay pendientes
            cmbCampo(2).AddItem "NINGUNO"
            cmbCampo(2).ItemData(cmbCampo(2).NewIndex) = -1
            cmbCampo(2).ListIndex = 0
        Else
            msActs = ""
            Do While Not adors.EOF
                cmbCampo(2).AddItem adors(3)
                cmbCampo(2).ItemData(cmbCampo(2).NewIndex) = adors(0)
                msActs = msActs & adors(1) & ","
                adors.MoveNext
            Loop
        End If
        'Actualiza Actividades realizadas
        ListView1.ListItems.Clear
        If adors.State Then adors.Close
        adors.Open "{call p_seguimientoActsRealizadas(" & mlAnálisis & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
        i = 1
        Do While Not adors.EOF
            ListView1.ListItems.Add i, , adors(1) 'Actividad
            ListView1.ListItems(i).SubItems(1) = IIf(IsNull(adors(2)), "", adors(2)) 'Tarea
            ListView1.ListItems(i).SubItems(2) = IIf(IsNull(adors(3)), "", adors(3)) 'Fecha
            ListView1.ListItems(i).SubItems(3) = IIf(IsNull(adors(4)), "", adors(4)) 'Responsable
            ListView1.ListItems(i).SubItems(4) = IIf(IsNull(adors(5)), "", adors(5)) 'Observaciones
            ListView1.ListItems(i).Tag = adors(0) 'Guarda el id
            adors.MoveNext
            i = i + 1
        Loop
        
    Else
        
        cmbCampo(2).Clear
    End If
ElseIf Index = 2 And cmbCampo(Index).ListIndex >= 0 Then 'Ejecuta módulo de captura de actividades
    If cmbCampo(Index).ItemData(cmbCampo(Index).ListIndex) < 0 Then 'No hace caso (No hay pendientes)
        Exit Sub
    End If
    If InStr(cmbCampo(Index).Text, ") (") Then 'verifica que el usuario sea el responsable quien debe ejecutar esta actividad programada
        If adors.State Then adors.Close
        adors.Open "select idusi from seguimientoprog where idant=" & cmbCampo(Index).ItemData(cmbCampo(Index).ListIndex) & " and idact=" & F_Obten_Act(msActs, cmbCampo(Index).ListIndex), gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            If adors(0) <> giUsuario Then
                Call MsgBox("El responsable a quien se asigno la actividad es la única persona quien puede dar seguimiento", vbOKOnly + vbInformation, "Validación de Seguimiento")
                Exit Sub
            End If
        End If
    End If
    With Actividades
        .miTarea = 0
        .mlAnálisis = mlAnálisis
        .mlAnt = cmbCampo(Index).ItemData(cmbCampo(Index).ListIndex)
        .miActividad = F_Obten_Act(msActs, cmbCampo(Index).ListIndex)
        .mlSeguimiento = 0
        .msObservaciones = ""
        .yTipoOperación = 1 'Agregar
        .mdFecha = CDate("01/01/1900")
        If InStr(cmbCampo(Index).Text, ") (") Then
            .msProgResp = Mid(cmbCampo(Index).Text, InStrRev(Mid(cmbCampo(Index).Text, 1, InStr(cmbCampo(Index).Text, ") (") - 3), "("))
        Else
            .msProgResp = ""
        End If
        gs = "no iniciar var"
        .Show vbModal
        If gs = "cancelar" Then
            cmbCampo(Index).ListIndex = -1
        Else
            cmbCampo_Click 1
        End If
    End With
End If
End Sub

'Obtiene el valor de la Actividad correspondiente al lugar iLugar
Private Function F_Obten_Act(ByVal sActs As String, iLugarCombo As Integer) As Integer
Dim i As Integer
For i = 1 To iLugarCombo
    If InStr(sActs, ",") = 0 Then
        F_Obten_Act = 0
        Exit Function
    End If
    sActs = Mid(sActs, InStr(sActs, ",") + 1)
Next
F_Obten_Act = Val(sActs)
End Function

Private Sub cmbPendientes_Click()
If cmbPendientes.ListIndex >= 0 And mbLimpiaExp Then
    mbLimpiaExp = False
    txtExpediente.Text = ""
    txtOficio.Text = ""
End If
End Sub

Private Sub cmbPendientes_GotFocus()
mbLimpiaExp = True
End Sub

Private Sub cmbPendientes_LostFocus()
mbLimpiaExp = False
End Sub

Private Sub cmdActualpen_Click()
Dim i As Integer
ActualizaPendientes
LimpiaControlesInfConsulta 0
End Sub


'Busca y en en caso de encontrar obtiene datos de este folio
Private Sub cmdContinuar_Click()
Dim adors As New ADODB.Recordset
Dim l As Long
If Len(Trim(txtOficio.Text)) > 0 Then
    If adors.State Then adors.Close
    adors.Open "select a.id,a.idregxif,ri.idreg,ss.idseg,sp.status from seguimientosanción ss, seguimiento s, análisis a, registroxif ri,seguimientosanpub sp  where ss.oficio='" & Replace(txtOficio.Text, "'", "''") & "' and ss.idseg=s.id and s.idana=a.id and a.idregxif=ri.id and s.id=sp.idseg(+)", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        If IsNull(adors!Status) Then ' No hay problema no ha sido seleccionado
        ElseIf adors!Status = 1 Then
            MsgBox "El Oficio ya fue seleccionado para publicación", vbOKOnly + vbInformation, ""
            Exit Sub
        ElseIf adors!publicar > 1 Then
            MsgBox "El Oficio ya fue publicado", vbOKOnly + vbInformation, ""
            Exit Sub
        End If
        
        mlAnálisis = adors(0)
        mlAsuxIF = adors(1)
        mlAsunto = adors(2)
        mlSeguimiento = adors(3)
        txtOficio.Enabled = False
        cmdContinuar.Enabled = False
        RefrescaDatos
    Else
        MsgBox "No se encontró asunto alguno con ese No. de Oficio", vbOKOnly + vbInformation, ""
    End If
Else
    MsgBox "Debe capturar el número de Oficio de Sanción", vbOKOnly + vbInformation, ""
End If
Exit Sub
End Sub

Private Sub RefrescaDatos()
Dim adors As ADODB.Recordset, i As Integer
Set adors = New ADODB.Recordset
adors.Open "{call p_analisis_datosregistro(" & mlAsunto & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
If Not adors.EOF Then
    For i = 0 To 2
        txtCampo(i).Text = adors(i)
    Next
    Set adors = New ADODB.Recordset
    adors.Open "select fecha,f_seguimiento_san_monto(idseg),f_seguimiento_causas(idseg) from seguimientosanxanacau where idseg=" & mlSeguimiento, gConSql, adOpenForwardOnly, adLockReadOnly
    If Not adors.EOF Then
        txtCampo(3).Text = Format(adors(0), "dd/mmm/yyyy")
        txtCampo(4).Text = Format(adors(1), "###,###,###.00")
        txtCampo(5).Text = adors(2)
        txtCampo(3).Tag = mlSeguimiento
    End If
    cmdProceso(0).Enabled = True
Else
    For i = 0 To 5
        txtCampo(i).Text = ""
    Next
    For i = 0 To cmbCampo.UBound
        cmbCampo(i).Clear
    Next
End If
End Sub

'Acciones Nuevo, editar , agrega, borra...
Private Sub cmdProceso_Click(Index As Integer)
Dim adors  As New ADODB.Recordset

On Error GoTo ErrorGuardaDatos:
If Index = 0 Then 'Agrega Sanción
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Tag = mlSeguimiento Then
            MsgBox "El oficio de sanción ya fue selecionado", vbOKOnly + vbInformation, ""
            Exit Sub
        End If
    Next
    If mlSeguimiento > 0 Then
        i = ListView1.ListItems.Count + 1
        ListView1.ListItems.Add i, , txtCampo(1).Text 'Institución
        ListView1.ListItems(i).SubItems(1) = txtCampo(0).Text 'Oficio
        ListView1.ListItems(i).SubItems(2) = txtCampo(3).Text 'Fecha
        ListView1.ListItems(i).SubItems(3) = txtCampo(4).Text 'Ley/Causa
        ListView1.ListItems(i).SubItems(4) = txtCampo(5).Text 'Monto
        ListView1.ListItems(i).Tag = txtCampo(3).Tag 'Guarda el idseg
    End If
    

'ElseIf Index = 1 Then 'Consultar
'    With Actividades
'        .mlAnálisis = mlAnálisis
'        .mlSeguimiento = ListView1.SelectedItem.Tag
'        .yTipoOperación = 0
'        .ySoloConsulta = 1
'        gs = "no iniciar var"
'        .Show vbModal
'    End With
ElseIf Index = 1 Then 'Quitar
    If ListView1.SelectedItem.Index > 0 Then
        If MsgBox("Está seguro de quitar el Oficio de sanción seleccionado", vbYesNo + vbQuestion, "") = vbYes Then
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
            ListView1.Refresh
        End If
    End If
ElseIf Index = 2 Then 'Publicar
    If ListView1.ListItems.Count <= 0 Then
        cmdProceso(3).Enabled = False
        Exit Sub
    End If
    If MsgBox("Está seguro de Preparar la publicación de los (" & ListView1.ListItems.Count & ") Oficios Seleccionados", vbYesNo + vbQuestion, "Confirmación") = vbNo Then
        Exit Sub
    End If
    If ListView1.ListItems.Count > 0 Then
        s = ""
        For i = 1 To ListView1.ListItems.Count
            s = s & ListView1.ListItems(i).Tag & ","
        Next
        If Len(s) > 0 Then
            s = Mid(s, 1, Len(s) - 1)
            gConSql.Execute "insert into seguimientosanpub select idseg,sysdate as fecha," & giUsuario & " as idusi,1 as status from seguimientosanción where idseg in (" & s & ")", iRows
            If iRows > 0 Then
                MsgBox "Se prepararon para publicación " & iRows & " Oficios", vbOKOnly + vbInformation, ""
                ListView1.ListItems.Clear
                cmdActualpen_Click
            End If
        End If
    End If
End If
Exit Sub
ErrorGuardaDatos:
If gConSql.Errors.Count > 0 Then
    yError = MsgBox("AVISO: " + gConSql.Errors(0).Description, vbAbortRetryIgnore + vbInformation, "Excepción (" + Str(gConSql.Errors(0).Number) + ")")
Else
    yError = MsgBox("AVISO: " + Err.Description, vbAbortRetryIgnore + vbInformation, "Excepción (" + Str(Err.Number) + ")")
End If


If yError = vbRetry Then
    Resume
ElseIf yError = vbIgnore Then
    Resume Next
End If


End Sub

Private Sub Form_Load()
ActualizaPendientes
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If ListView1.ListItems.Count > 0 Then
    If MsgBox("Está seguro de salir sin preparar la publicación de los oficios seleccionados", vbQuestion + vbYesNo, "Confirmación") = vbNo Then
        Cancel = True
        Exit Sub
    End If
End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim adors As New ADODB.Recordset
If Not cmdProceso(1).Enabled Then cmdProceso(1).Enabled = True
If Not cmdProceso(2).Enabled Then cmdProceso(2).Enabled = True
End Sub

Private Sub txtCampo_Change(Index As Integer)
If Index >= 5 And Index <= 6 Then
    'Image2.Picture = Imagenes.ListImages(2).Picture
End If
End Sub

Private Sub txtCampo_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 1 And InStr("-", Chr(KeyAscii)) > 0 Then Exit Sub
KeyAscii = TeclaOprimida(txtCampo(Index), KeyAscii, txtCampo(Index).Tag, False)
'MsgBox "asd"
End Sub

Private Sub txtCampo_LostFocus(Index As Integer)
Dim adors As New ADODB.Recordset
If Mid(txtCampo(Index).Tag, 1, 1) = "f" Then
    If IsDate(txtCampo(Index).Text) Then
        d = CDate(txtCampo(Index).Text)
        txtCampo(Index).Text = Format(d, gsFormatoFecha)
        adors.Open "select sysdate from dual", gConSql, adOpenStatic, adLockReadOnly
        If Int(adors(0)) - Int(d) < 0 Then
            Call MsgBox("Fecha no válida. No se permite ingresar fecha mayor a la fecha actual (" & Format(adors(0), gsFormatoFecha) & ")", vbOKOnly + vbInformation, "")
            txtCampo(Index) = ""
            Exit Sub
        End If
    Else
        If Len(txtCampo(Index).Text) > 0 Then
            Call MsgBox("Fecha no válida. Verificar", vbOKOnly + vbInformation, "")
            txtCampo(Index) = ""
        End If
    End If
End If
End Sub

Sub ActualizaPendientes()
Dim adors As New ADODB.Recordset
End Sub

Private Function F_ObtieneAct(ByVal sActs As String, ByVal iPos As Integer) As Integer
Dim i As Integer
Do While i < iPos
    If InStr(sActs, ",") = 0 Then
        F_ObtieneAct = -99
        Return
    End If
    sActs = Mid(sActs, InStr(sActs, ",") + 1)
Loop
If i = iPos Then
    F_ObtieneAct = Val(sActs)
End If
End Function

Private Sub LimpiaControlesInfConsulta(iTipo As Byte)
If iTipo = 0 Then 'Limpia todo y prepara un nuevo seguimiento
    txtOficio.Enabled = True
    txtOficio.Text = ""
    cmdContinuar.Enabled = True
    'ListView1.ListItems.Clear
    For i = 0 To cmdProceso.UBound
        cmdProceso(i).Enabled = False
    Next
    For i = 0 To txtCampo.UBound
        txtCampo(i).Text = ""
    Next
ElseIf iTipo = 1 Then 'Limpia una nueva institución
    txtCampo(3).Text = ""
    ListView1.ListItems.Clear
End If
End Sub

