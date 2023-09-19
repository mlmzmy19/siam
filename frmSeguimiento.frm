VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "fm20.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form Seguimiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Análisis"
   ClientHeight    =   10035
   ClientLeft      =   4905
   ClientTop       =   2595
   ClientWidth     =   14145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   14145
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1050
      Left            =   45
      TabIndex        =   31
      Top             =   3690
      Width           =   14010
      Begin VB.ComboBox cmbCampo 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   0
         ItemData        =   "frmSeguimiento.frx":0000
         Left            =   2250
         List            =   "frmSeguimiento.frx":0002
         TabIndex        =   33
         ToolTipText     =   "Institución"
         Top             =   225
         Width           =   11640
      End
      Begin VB.ComboBox cmbCampo 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   1
         ItemData        =   "frmSeguimiento.frx":0004
         Left            =   2250
         List            =   "frmSeguimiento.frx":0006
         TabIndex        =   32
         ToolTipText     =   "Oficio de Emplazamiento o Acuerdo de Improcedencia"
         Top             =   630
         Width           =   11640
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Institución:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   1395
         TabIndex        =   35
         Top             =   270
         Width           =   765
      End
      Begin VB.Label etiCombo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Acuerdo improc. / Ofi. emplazamiento (Ofi. sanción) :"
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   1
         Left            =   90
         TabIndex        =   34
         Top             =   540
         Width           =   2190
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5325
      Left            =   45
      TabIndex        =   20
      Top             =   4635
      Width           =   14010
      Begin VB.ComboBox comboMulta 
         Height          =   315
         Left            =   135
         TabIndex        =   39
         Top             =   1350
         Visible         =   0   'False
         Width           =   7755
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "Nombre"
         ForeColor       =   &H00808080&
         Height          =   690
         Index           =   3
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Tag             =   "c"
         ToolTipText     =   "Datos capturados del Oficio en la etapa de Análisis"
         Top             =   360
         Width           =   13755
      End
      Begin VB.ComboBox cmbCampo 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   2
         ItemData        =   "frmSeguimiento.frx":0008
         Left            =   8010
         List            =   "frmSeguimiento.frx":000A
         TabIndex        =   26
         ToolTipText     =   "Actividades pendientes en seguimiento del Oficio"
         Top             =   1350
         Width           =   5880
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         Height          =   3375
         Left            =   12060
         TabIndex        =   21
         Top             =   1890
         Width           =   1830
         Begin VB.ComboBox cmbSINE 
            Height          =   315
            Left            =   135
            TabIndex        =   42
            Top             =   2340
            Width           =   1590
         End
         Begin VB.CommandButton cmdProceso 
            Caption         =   "&SINE"
            Height          =   375
            Index           =   3
            Left            =   450
            TabIndex        =   36
            Top             =   2835
            Width           =   945
         End
         Begin VB.CommandButton cmdProceso 
            Caption         =   "&Modificar"
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   495
            TabIndex        =   25
            Top             =   225
            Width           =   945
         End
         Begin VB.CommandButton cmdProceso 
            Caption         =   "&Borrar"
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   495
            TabIndex        =   23
            Top             =   1404
            Width           =   945
         End
         Begin VB.CommandButton cmdProceso 
            Caption         =   "&Consultar"
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   495
            TabIndex        =   22
            Top             =   828
            Width           =   945
         End
         Begin VB.Label etiCombo 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Notificación Especial SINE:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   480
            Index           =   4
            Left            =   288
            TabIndex        =   41
            Top             =   1908
            Width           =   1272
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3315
         Left            =   90
         TabIndex        =   24
         Top             =   1935
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   5847
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Actividad"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tarea"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Responsable"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Usuario SIAM"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Observaciones"
            Object.Width           =   11465
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sanción (Multa) --> Seguimiento:"
         Height          =   240
         Left            =   135
         TabIndex        =   38
         Top             =   1080
         Visible         =   0   'False
         Width           =   5325
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Datos del oficio:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   30
         Top             =   135
         Width           =   1140
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Actividades realizadas:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   28
         Top             =   1710
         Width           =   1620
      End
      Begin VB.Label etiCombo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Actividades pendientes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   2
         Left            =   8010
         TabIndex        =   27
         Top             =   1035
         Width           =   2595
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1725
      Left            =   45
      TabIndex        =   14
      Top             =   2025
      Width           =   14010
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "Nombre"
         ForeColor       =   &H00808080&
         Height          =   1095
         Index           =   2
         Left            =   9285
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Tag             =   "c"
         ToolTipText     =   "Datos del documento de Solicitud"
         Top             =   360
         Width           =   4545
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "Nombre"
         ForeColor       =   &H00808080&
         Height          =   1095
         Index           =   1
         Left            =   4695
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Tag             =   "c"
         ToolTipText     =   "Nombre de la Institución y del Usuario"
         Top             =   360
         Width           =   4545
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
         TabIndex        =   8
         Tag             =   "c"
         ToolTipText     =   "Datos del origen de la Solicitud"
         Top             =   360
         Width           =   4545
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Documento de la solicitud:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   9315
         TabIndex        =   17
         Top             =   135
         Width           =   1875
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Institución / Nombre(s):"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   4770
         TabIndex        =   16
         Top             =   135
         Width           =   1650
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Origen de la solicitud:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   15
         Top             =   135
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00B9D6F2&
      Height          =   2040
      Left            =   2160
      TabIndex        =   11
      Top             =   -45
      Width           =   11985
      Begin VB.CommandButton Command1 
         Caption         =   "Actualiza Lista Pendientes"
         Height          =   444
         Left            =   8172
         TabIndex        =   4
         Top             =   1476
         Width           =   1452
      End
      Begin VB.CommandButton cmdDif 
         BackColor       =   &H00C0C0FF&
         Caption         =   "No empata con MSS (Ver diferencias)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   9828
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1440
         Visible         =   0   'False
         Width           =   1995
      End
      Begin MSComctlLib.ImageList Imagenes 
         Left            =   8910
         Top             =   900
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   103
         ImageHeight     =   104
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSeguimiento.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSeguimiento.frx":7F1E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtOficioS 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   5730
         MaxLength       =   80
         TabIndex        =   3
         Tag             =   "c"
         ToolTipText     =   "No. de Oficio a realizar seguimiento"
         Top             =   1125
         Width           =   3300
      End
      Begin VB.CommandButton cmdDescto 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Consultar Descuento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.TextBox txtOficioE 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   5730
         MaxLength       =   80
         TabIndex        =   2
         Tag             =   "c"
         ToolTipText     =   "No. de Oficio a realizar seguimiento"
         Top             =   630
         Width           =   3300
      End
      Begin VB.CommandButton cmdActualpen 
         Caption         =   "Nuevo seguimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   216
         TabIndex        =   6
         Top             =   180
         Width           =   2244
      End
      Begin VB.ComboBox cmbPendientes 
         BackColor       =   &H8000000F&
         Height          =   288
         Left            =   192
         TabIndex        =   1
         ToolTipText     =   "Oficios pendientes de realizar seguimiento"
         Top             =   1575
         Width           =   7896
      End
      Begin VB.TextBox txtExpediente 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   192
         MaxLength       =   100
         TabIndex        =   0
         Tag             =   "c"
         ToolTipText     =   "No de expediente a realizar seguimiento"
         Top             =   936
         Width           =   3990
      End
      Begin VB.CommandButton cmdContinuar 
         BackColor       =   &H00008000&
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10044
         Picture         =   "frmSeguimiento.frx":10820
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   864
         Width           =   1635
      End
      Begin MSForms.CommandButton cmdIrAna 
         Height          =   555
         Left            =   8280
         TabIndex        =   43
         Top             =   180
         Width           =   1725
         BackColor       =   14735199
         Caption         =   "<< Regresa Análisis"
         Size            =   "3043;979"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label2 
         BackColor       =   &H00B9D6F2&
         Caption         =   "Oficio de sanción:"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   4320
         TabIndex        =   37
         Top             =   1170
         Width           =   1410
      End
      Begin VB.Label Label2 
         BackColor       =   &H00B9D6F2&
         Caption         =   "No. Expediente:"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   216
         TabIndex        =   19
         Top             =   648
         Width           =   1632
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00B9D6F2&
         Caption         =   "Oficios pendientes:"
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   7
         Left            =   180
         TabIndex        =   18
         Top             =   1332
         Width           =   1356
      End
      Begin VB.Label Label2 
         BackColor       =   &H00B9D6F2&
         Caption         =   "Ofi. emplazamiento:"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   4320
         TabIndex        =   13
         Top             =   675
         Width           =   1410
      End
      Begin VB.Label Eti 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00B9D6F2&
         Caption         =   "Módulo de Seguimiento"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   2
         Left            =   1035
         TabIndex        =   12
         Top             =   180
         Width           =   7740
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Image Image1 
      Height          =   1875
      Left            =   0
      Picture         =   "frmSeguimiento.frx":1138F
      Stretch         =   -1  'True
      Top             =   90
      Width           =   2115
   End
End
Attribute VB_Name = "Seguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim miSINE 'Indica si fue invocado las notificaciones especiales de SINE que se pueden realizar desde el seguimiento
Dim mbLimpiaExp As Boolean 'indicador para limpiar el campo o lista de exp Pendientes
Dim msActs As String
Dim msExpediente As String


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
        MsgBox "La cadena de conexión al SINE viene vacia Favor de reportarlo al administrador del Sistema (Miguel Ext.6032)"
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
    adors.Open "select id, oficio||' (Sanción: '||f_analisis_san_oficio(id)||') '||case when procedente=1 then '' else ' (Improc.)' end as descripción from análisis where idregxif=" & mlAsuxIF, gConSql, adOpenStatic, adLockReadOnly
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
ElseIf Index = 1 Then  'filtra pendientes salvo existan avances con idanacau not nulo
    If cmbCampo(Index).ListIndex >= 0 Then
        mlAnálisis = cmbCampo(Index).ItemData(cmbCampo(Index).ListIndex)
        If adors.State Then adors.Close
        adors.Open "select oficio||' ('||to_char(fecha,'dd/mon/yyyy')||')'||chr(13)||chr(10)||case when procedente=1 then f_causas(id) else f_causasimp(id) end as descripción from análisis where id=" & mlAnálisis, gConSql, adOpenForwardOnly, adLockReadOnly
        If Not adors.EOF Then
            txtcampo(3).Text = IIf(IsNull(adors(0)), "", adors(0))
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
        adors.Open "{call p_seguimientoActsRealizadas2(" & mlAnálisis & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
        i = 1
        Do While Not adors.EOF
            ListView1.ListItems.Add i, , adors(1) 'Actividad
            ListView1.ListItems(i).SubItems(1) = IIf(IsNull(adors(2)), "", adors(2)) 'Tarea
            ListView1.ListItems(i).SubItems(2) = IIf(IsNull(adors(3)), "", adors(3)) 'Fecha
            ListView1.ListItems(i).SubItems(3) = IIf(IsNull(adors(4)), "", adors(4)) 'Responsable
            ListView1.ListItems(i).SubItems(4) = IIf(IsNull(adors(5)), "", adors(5)) 'UsuarioSIAM
            ListView1.ListItems(i).SubItems(5) = IIf(IsNull(adors(6)), "", adors(6)) 'Observaciones
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
            'If adors(0) <> giUsuario Then
            '    Call MsgBox("El responsable a quien se asigno la actividad es la única persona quien puede dar seguimiento", vbOKOnly + vbInformation, "Validación de Seguimiento")
            '    Exit Sub
            'End If
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
    txtOficioE.Text = ""
    txtOficioS.Text = ""
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
'ActualizaPendientes
LimpiaControlesInfConsulta 0
End Sub


'Busca y en en caso de encontrar obtiene datos de este folio
Private Sub cmdContinuar_Click()
Dim adors As New ADODB.Recordset
Dim l As Long, n1 As Integer, n2 As Integer, n3 As Integer, n4 As Integer
If Len(Trim(txtExpediente.Text)) = 0 And Len(Trim(txtOficioE.Text)) = 0 And Len(Trim(txtOficioS.Text)) = 0 And cmbPendientes.ListIndex < 0 Then
    MsgBox "Debe capturar el No. de expediente, oficio o seleccionar de la lista el expediente pendiente", vbOKOnly + vbInformation, ""
    Exit Sub
End If
If Len(Trim(txtOficioS.Text)) > 0 Then 'Busca Por oficio sanción
    If adors.State Then adors.Close
    adors.Open "select f_seguimiento_idana(idseg) from seguimientosanción where oficio='" & Replace(txtOficioS.Text, "'", "''") & "'", gConSql, adOpenStatic, adLockReadOnly
    If adors.EOF Then
        MsgBox "No se encontró asunto alguno con ese No. de oficio de sanción", vbOKOnly + vbInformation, ""
    ElseIf adors(0) > 0 Then
        l = adors(0) 'id análisis
    Else 'No encontro datos
        MsgBox "No se encontró asunto alguno con ese No. de oficio de sanción", vbOKOnly + vbInformation, ""
    End If
    If adors.State Then adors.Close
    adors.Open "select f_analisis_idregxif(" & l & ") as regxif,f_analisis_idreg(" & l & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        mlAnálisis = l
        mlAsuxIF = adors(0)
        mlAsunto = adors(1)
        If f_verificaModNoPend > 0 Then
            'MsgBox "No puede realizar seguimiento hasta completar las causas pendientes de asociar turnadas desde MÓDULOS", vbOKOnly + vbInformation, "Validación"
            'Exit Sub
            cmdDif.Visible = True
        End If
        txtExpediente.Enabled = False
        txtOficioS.Enabled = False
        txtOficioE.Enabled = False
        cmbPendientes.Enabled = False
        cmdContinuar.Enabled = False
        RefrescaDatos
    Else
        MsgBox "No se encontró asunto alguno con ese No. de Oficio", vbOKOnly + vbInformation, ""
    End If
ElseIf Len(Trim(txtOficioE.Text)) > 0 Then 'Busca Por oficio emplazamiento
    If adors.State Then adors.Close
    adors.Open "select count(*) as cont, max(id) as idana from análisis where oficio='" & Replace(txtOficioE.Text, "'", "''") & "'", gConSql, adOpenStatic, adLockReadOnly
    If adors(0) = 0 Then
        MsgBox "No se encontró asunto alguno con ese No. de oficio de emplazamiento", vbOKOnly + vbInformation, ""
    ElseIf adors(1) > 0 Then
        l = adors(1) 'id análisis
        mlAnálisis = l
    End If
    If adors.State Then adors.Close
    adors.Open "select a.id,a.idregxif,ri.idreg from análisis a, registroxif ri where oficio='" & Replace(txtOficioE.Text, "'", "''") & "' and a.idregxif=ri.id", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then '''Cambio hoy
        mlAnálisis = adors(0)
        mlAsuxIF = adors(1)
        mlAsunto = adors(2)
        If f_verificaModNoPend > 0 Then
            MsgBox "No puede realizar seguimiento hasta completar las causas pendientes de asociar turnadas desde MÓDULOS", vbOKOnly + vbInformation, "Validación"
            Exit Sub
        End If
        txtExpediente.Enabled = False
        txtOficioS.Enabled = False
        txtOficioE.Enabled = False
        cmbPendientes.Enabled = False
        cmdContinuar.Enabled = False
        RefrescaDatos
    Else
        MsgBox "No se encontró asunto alguno con ese No. de oficio de emplazamiento", vbOKOnly + vbInformation, ""
    End If
ElseIf Len(Trim(txtExpediente.Text)) > 0 Then 'Busca por expediente
    msExpediente = txtExpediente.Text
    If adors.State Then adors.Close
    adors.Open "select f_asuntoxfolio('" & txtExpediente.Text & "') from dual", gConSql, adOpenStatic, adLockReadOnly
    If adors(0) > 0 Then
        mlAsunto = adors(0)
        If adors.State Then adors.Close
        adors.Open "select count(*) from análisis where idregxif in (select id from registroxif where idreg=" & mlAsunto & ")", gConSql, adOpenStatic, adLockReadOnly
        If adors(0) = 0 Then
            MsgBox "No se encontró registro de oficio de emplazamiento para este expediente. Favor de registrar primero en el módulo de Análisis ", vbOKOnly + vbInformation, ""
            Exit Sub
        End If
        If f_verificaModNoPend > 0 Then
            'MsgBox "No puede realzar seguimiento hasta completar las causas pendientes de asociar turnadas desde MÓDULOS", vbOKOnly + vbInformation, "Validación"
            'Exit Sub
            cmdDif.Visible = True
        End If
        txtExpediente.Enabled = False
        txtOficioS.Enabled = False
        txtOficioE.Enabled = False
        cmbPendientes.Enabled = False
        cmdContinuar.Enabled = False
        RefrescaDatos
    Else
        MsgBox "No se encontró asunto alguno con ese No. de Expediente.", vbOKOnly + vbInformation, ""
    End If
ElseIf cmbPendientes.ListIndex >= 0 Then 'Por pendiente
    'gi = cmbPendientes.ItemData(cmbPendientes.ListIndex)
    If adors.State Then adors.Close
    adors.Open "select a.id,a.idregxif,ri.idreg from análisis a, registroxif ri where a.id=" & cmbPendientes.ItemData(cmbPendientes.ListIndex) & " and a.idregxif=ri.id", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        If adors(0) > 0 Then
            mlAnálisis = adors(0)
            mlAsuxIF = adors(1)
            mlAsunto = adors(2)
            If f_verificaModNoPend > 0 Then
                MsgBox "No puede realzar seguimiento hasta completar las causas pendientes de asociar turnadas desde MÓDULOS", vbOKOnly + vbInformation, "Validación"
                Exit Sub
            End If
            txtExpediente.Enabled = False
            txtOficioS.Enabled = False
            txtOficioE.Enabled = False
            cmbPendientes.Enabled = False
            cmdContinuar.Enabled = False
            RefrescaDatos
        End If
    End If
    'gi = 0
End If
If Not cmdDescto.Visible Then cmdDescto.Visible = True

cmbSINE.Clear
If adors.State Then adors.Close
adors.Open "select f_sine_ifactiva(f_analisis_idins(" & mlAnálisis & "),2) from dual", gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then 'Tiene convenio SINE
    cmdProceso(3).Enabled = True
    cmbSINE.Enabled = True
    'mete los oficios y procesos por considerar
    'MsgBox "1"
    If adors.State Then adors.Close
    adors.Open "select sum(case when f_idtar_idpro(idtar)=2 then 1 else 0 end) as Pro1,sum(case when f_idtar_idpro(idtar)=3 then 1 else 0 end) as Pro2,sum(case when f_idtar_idpro(idtar)=4 then 1 else 0 end) as Pro3,sum(case when f_idtar_idpro(idtar)=9 then 1 else 0 end)as Pro4 from seguimiento where idana=" & mlAnálisis, gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        n1 = IIf(IsNull(adors(0)), 0, adors(0))
        n2 = IIf(IsNull(adors(1)), 0, adors(1))
        n3 = IIf(IsNull(adors(2)), 0, adors(2))
        n4 = IIf(IsNull(adors(3)), 0, adors(3))
    End If
    'Agrega Consulta
    cmbSINE.AddItem "Consulta"
    cmbSINE.ItemData(cmbSINE.ListCount - 1) = 0
    'MsgBox "2"
    If n1 > 0 Then 'Existes actividades de emplazamiento
        If adors.State Then adors.Close
        adors.Open "select f_analisis_oficio(" & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
        cmbSINE.AddItem "Emplazamiento" & IIf(IsNull(adors(0)), "", "(" & adors(0) & ")")
        cmbSINE.ItemData(cmbSINE.ListCount - 1) = 2
    End If
    'MsgBox "3"
    If n2 > 0 Then 'Existes actividades de sanción
        If adors.State Then adors.Close
        adors.Open "select f_analisis_san_oficio(" & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
        cmbSINE.AddItem "Instrucción" & IIf(IsNull(adors(0)), "", "(" & adors(0) & ")")
        cmbSINE.ItemData(cmbSINE.ListCount - 1) = 3
    End If
    'MsgBox "3"
    If n3 > 0 Then 'Existes actividades de sanción
        If adors.State Then adors.Close
        adors.Open "select f_analisis_san_oficio(" & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
        cmbSINE.AddItem "Sanción" & IIf(IsNull(adors(0)), "", "(" & adors(0) & ")")
        cmbSINE.ItemData(cmbSINE.ListCount - 1) = 4
    End If
    'MsgBox "4"
    If n4 > 0 Then 'Existes actividades de condonación
        If adors.State Then adors.Close
        adors.Open "select f_analisis_cond_oficio(" & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
        cmbSINE.AddItem "Condonación" & IIf(IsNull(adors(0)), "", "(" & adors(0) & ")")
        cmbSINE.ItemData(cmbSINE.ListCount - 1) = 9
    End If
    If cmbSINE.ListCount >= 1 Then
        cmbSINE.ListIndex = 0
    End If
'ElseIf adors(0) = 4 Then 'Sanción
'    If adors.State Then adors.Close
'    adors.Open "select f_analisis_san_oficio(" & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
'ElseIf adors(0) = 9 Then 'Condonación
'    If adors.State Then adors.Close
'    adors.Open "select f_analisis_cond_oficio(" & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
Else
    cmbSINE.Enabled = False
    cmdProceso(3).Enabled = False
End If


Exit Sub
End Sub

Function f_verificaModNoPend() As Integer
Dim adors As New ADODB.Recordset
adors.Open "select f_mod_CausasIncompletas(" & mlAsunto & ") from dual", gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    f_verificaModNoPend = 1
End If
End Function




Private Sub RefrescaDatos()
Dim adors As ADODB.Recordset, i As Integer
Set adors = New ADODB.Recordset
adors.Open "{call p_analisis_datosregistro(" & mlAsunto & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
If Not adors.EOF Then
    For i = 0 To 2
        txtcampo(i).Text = adors(i)
    Next
    txtExpediente.Text = IIf(IsNull(adors(4)), "", adors(4))
    LlenaCombo cmbCampo(0), "select id,f_institucionh(idins) as descrip from registroxif where idreg=" & mlAsunto, "", True
    If cmbCampo(0).ListCount = 1 Then
        cmbCampo(0).ListIndex = 0
    Else
        If mlAsuxIF > 0 Then
            i = BuscaCombo(cmbCampo(0), mlAsuxIF, True, False)
            If i >= 0 Then
                cmbCampo(0).ListIndex = i
            End If
            DoEvents 'llena el combo de oficios
        End If
    End If
    If mlAnálisis > 0 Then
        If mlAsuxIF > 0 Then
            i = BuscaCombo(cmbCampo(1), mlAnálisis, True, False)
            If i >= 0 Then
                cmbCampo(1).ListIndex = i
                If Len(Trim(txtOficioE.Text)) = 0 And InStr(cmbCampo(1), "(") > 0 Then
                    txtOficioE.Text = Mid(cmbCampo(1).Text, 1, InStrRev(cmbCampo(1), "(") - 1)
                Else
                    txtOficioE.Text = cmbCampo(1).Text
                End If
            End If
        End If
    End If
Else
    For i = 0 To 2
        txtcampo(i).Text = ""
    Next
    For i = 0 To cmbCampo.UBound
        cmbCampo(i).Clear
    Next
End If
If gi > 0 Then
    i = BuscaCombo(cmbCampo(0), gi, True, False)
    If i >= 0 Then
        cmbCampo(0).ListIndex = i
    End If
End If
End Sub

Private Sub cmdDescto_Click()
Dim adors As New ADODB.Recordset
adors.Open "select f_seguimiento_tarea_real(" & mlAnálisis & ", 37, 1),f_seguimiento_tarea_real(" & mlAnálisis & ", 41, 1) from dual", gConSql, adOpenStatic, adLockReadOnly
If adors(0) <= 0 Or adors(1) <= 0 Then
    Call MsgBox("Este asunto no tiene oficio de sanción realizada o no ha sido notificada aun", vbOKOnly + vbInformation, "Informativo")
    Exit Sub
End If
'With Finanzas_Desctos
'    .piSeg = adors(0)
'    .Show vbModal
'End With
End Sub

Private Sub cmdIrAna_Click()
Dim frm As Form
If mlAsunto > 0 Then
    gs = "<<"
    If Len(Trim(msExpediente)) > 0 Then
        gs1 = Trim(msExpediente)
    End If
    If cmbCampo(0).Visible And cmbCampo(0).ListIndex >= 0 Then
        gi1 = cmbCampo(0).ItemData(cmbCampo(0).ListIndex)
    End If
    Set frm = Análisis
    With frm
        .Show
    End With
End If
End Sub

'Acciones Nuevo, editar , agrega, borra...
Private Sub cmdProceso_Click(Index As Integer)
Dim adors  As New ADODB.Recordset
Dim bUnico As Boolean, s As String
Dim iPro As Integer
Dim strFic As String
Dim web1 As Object
Dim EnvString As String, i As Integer, PathLen As Integer
On Error GoTo ErrorGuardaDatos:
If Index = 3 Then 'SINE
    If mlAnálisis <= 0 Then
        cmdProceso(3).Enabled = False
        Exit Sub
    End If
    If ListView1.SelectedItem.Index <= 0 Then
        cmdProceso(3).Visible = False
        Exit Sub
    End If
    If cmbSINE.ListIndex < 0 Then
        MsgBox "Debe eleguir una opción: COnsulta o Notificación Especial", vbInformation + vbOKOnly, ""
        If Not cmbSINE.Visible Or Not cmbSINE.Enabled Then
            cmbSINE.Visible = True
            cmbSINE.Enabled = True
            cmbSINE.SetFocus
            Exit Sub
        End If
    End If
    iPro = cmbSINE.ItemData(cmbSINE.ListIndex)
    
    'If adors.State Then adors.Close
    'adors.Open "select idusi from seguimiento where id=" & ListView1.ListItems(ListView1.SelectedItem.Index).Tag, gConSql, adOpenStatic, adLockReadOnly
    'If Not adors.EOF Then
    '    If adors(0) <> giUsuario Then
    '        Call MsgBox("El responsable a quien realizó la actividad es la única persona quien puede visualizar las notificaciones del SINE", vbOKOnly + vbInformation, "Validación de Seguimiento")
    '        Exit Sub
    '    End If
    'End If

'    If adors.State Then adors.Close
'    adors.Open "select f_sine_ifactiva(f_analisis_idins(" & mlAnálisis & "),f_act_idpro(f_seguimiento_idact(" & ListView1.ListItems(ListView1.SelectedItem.Index).Tag & "))) from dual", gConSql, adOpenStatic, adLockReadOnly
'    If adors(0) <= 0 Then
'        If MsgBox("La IF de este asunto no tiene convenio de notificaciones electrónicas. Desea consultar notificaciones pasadas en caso de haber tenido SINE en el pasado", vbYesNo + vbInformation, "Validación") = vbNo Then
'            cmdProceso(3).Visible = False
'            Exit Sub
'        End If
'    End If

    If adors.State Then adors.Close
    adors.Open "select f_url_conexion(13," & giUsuario & "," & mlAnálisis & "," & iPro & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    If Len(adors(0)) > 0 Then
        gsWWW = adors(0)
    Else
        MsgBox "La conexión al SINE no aplica para el proceso " & adors(1) & ". Si tiene alguna duda por favor de aclararlo con el personal de la UDEPO.", vbOKOnly + vbInformation
        Exit Sub
    End If
    i = 1
    Do
        EnvString = Environ(i)    ' Get environment
                    ' variable.
        If Left(UCase(EnvString), 5) = "PATH=" Then ' Se cuenta con chrome
            'Call MsgBox(EnvString)
            If InStr(UCase(EnvString), "CHROME") > 0 Then
                'Call MsgBox("OK")
                i = 4000
                Exit Do
            End If
            i = i + 1
        Else
            i = i + 1    ' Not PATH entry,
        End If    ' so increment.
    Loop Until EnvString = ""
    If i = 4000 Then ' Se tiene chrome
        strFic = "CHROME.exe """ & gsWWW & """"
        'Call MsgBox(strFic, vbOKOnly, "")
        Shell strFic, vbMaximizedFocus
    Else
        strFic = "C:\Program Files\Internet Explorer\iexplore.exe"
        If Len(Dir(strFic, vbArchive)) > 0 Then
            Shell strFic & " " & gsWWW, vbMaximizedFocus
            'abrir$ = ruta$ & "\" & inicio.TxtTituloWeb & ".htm"
            'Set web1 = CreateObject("InternetExplorer.Application")
            'web1.Navigate (gsWWW)
            'web1.Visible = True
        Else
            gsWWW = gsWWW '& "&perm=0"
            With Browser
                .yÚnicavez = 0
                .Caption = "Notificaciones Electrónicas"
                .Show vbModal
            End With
        End If
    End If
ElseIf Index = 0 Then 'Modificar
    If ListView1.ListItems.Count = 0 Then
        Exit Sub
    End If
    If adors.State Then adors.Close
    adors.Open "select motivo from tareasnoborrar where idtar=(select idtar from seguimiento where id=" & ListView1.ListItems(ListView1.SelectedItem.Index).Tag & ")", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        If adors(0) <> giUsuario Then
            Call MsgBox("La actividad/tarea del seguimiento no puede ser modificado debido a: " & adors(0), vbOKOnly + vbInformation, "Validación de Seguimiento")
            Exit Sub
        End If
    End If
    ''Verifica si es propietario de la actividad para poder modificar
    'If adors.State Then adors.Close
    'adors.Open "select idusi from seguimiento where id=" & ListView1.ListItems(ListView1.SelectedItem.Index).Tag, gConSql, adOpenStatic, adLockReadOnly
    'If Not adors.EOF Then
    '    If adors(0) <> giUsuario Then
    '        Call MsgBox("El responsable a quien realizó la actividad es la única persona quien puede modificarla", vbOKOnly + vbInformation, "Validación de Seguimiento")
    '        Exit Sub
    '    End If
    'End If
    With Actividades
        .mlAnálisis = mlAnálisis
        .mlSeguimiento = ListView1.SelectedItem.Tag
        .yTipoOperación = 2
        '.ySoloConsulta = 0
        gs = "no iniciar var"
        .Show vbModal
        cmbCampo_Click 1
    End With
ElseIf Index = 1 Then 'Consultar
    With Actividades
        .mlAnálisis = mlAnálisis
        .mlSeguimiento = ListView1.SelectedItem.Tag
        .yTipoOperación = 0
        '.ySoloConsulta = 1
        gs = "no iniciar var"
        .Show vbModal
    End With
ElseIf Index = 2 Then 'borrar
    Dim iRows As Integer
    If adors.State Then adors.Close
    adors.Open "select motivo from tareasnoborrar where idtar=(select idtar from seguimiento where id=" & ListView1.ListItems(ListView1.SelectedItem.Index).Tag & ")", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        If adors(0) <> giUsuario Then
            Call MsgBox("La actividad/tarea del seguimiento no puede ser borrado debido a: " & adors(0), vbOKOnly + vbInformation, "Validación de Seguimiento")
            Exit Sub
        End If
    End If
    
    'Verifica si es única actividad para proceder a borrar oficio de emplazamiento
    s = "Está seguro de borrar el registro seleccionado"
    If adors.State Then adors.Close
    adors.Open "select count(*) from seguimiento where idana=" & mlAnálisis, gConSql, adOpenStatic, adLockReadOnly
    If adors(0) = 1 Then
        bUnico = True
        s = "Borrar la actividad única eliminará el oficio de emplazamieto también. Está seguro de borrar el registro seleccionado"
    End If
    
    'Verifica quien dio de alta el seguimiento para proceder o no
    'If adors.State Then adors.Close
    'adors.Open "select idusi from seguimiento where id=" & ListView1.ListItems(ListView1.SelectedItem.Index).Tag, gConSql, adOpenStatic, adLockReadOnly
    'If Not adors.EOF Then
    '    If adors(0) <> giUsuario Then
    '        Call MsgBox("El responsable a quien realizó la actividad es la única persona quien puede borrarla", vbOKOnly + vbInformation, "Validación de Seguimiento")
    '        Exit Sub
    '    End If
    'End If
    If MsgBox(s, vbYesNo + vbQuestion, "") = vbYes Then
        If adors.State Then adors.Close
        If Not bUnico Then
            adors.Open "{call p_seguimiento_borrareg(" & ListView1.ListItems(ListView1.SelectedItem.Index).Tag & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
        Else
            adors.Open "{call p_analisis_borrareg(" & mlAnálisis & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
        End If
        If adors(0) > 0 Then
            MsgBox "Se borró el registro seleccionado", vbOKOnly, ""
            cmbCampo_Click 1
        Else
            MsgBox "No se borró el registro seleccionado", vbOKOnly + vbInformation, ""
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

Private Sub Command1_Click()
ActualizaPendientes
End Sub

Private Sub Form_Activate()
Dim i As Integer
If gs = ">>" Then
    If Len(gs1) > 0 Then
        cmdActualpen_Click
        If cmdContinuar.Enabled Then
            txtExpediente.Text = gs1
            cmdContinuar_Click
            If cmbCampo(0).ListIndex < 0 And cmbCampo(0).ListCount > 1 And gi1 > 0 Then
                i = BuscaCombo(cmbCampo(0), gi1, True)
                If i >= 0 Then
                    cmbCampo(0).ListIndex = i
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub Form_Load()
'ActualizaPendientes
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
gs = ""
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim adors As New ADODB.Recordset
mlSeguimiento = Val(Item.Tag)
If adors.State Then adors.Close
adors.Open "select count(*) from seguimiento s, (select * from seguimiento where idana=" & mlAnálisis & " and id<>idant) s1 where s.id=" & mlSeguimiento & " and s.id=s1.idant", gConSql, adOpenStatic, adLockReadOnly
cmdProceso(0).Enabled = (adors(0) = 0) 'modificar
cmdProceso(1).Enabled = True 'consultar
cmdProceso(2).Enabled = (adors(0) = 0) 'borrar
If adors.State Then adors.Close
adors.Open "select f_sine_ifactiva(f_analisis_idins(" & mlAnálisis & "),f_act_idpro(f_seguimiento_idact(" & Item.Tag & "))) from dual", gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    If Not cmdProceso(3).Visible Then cmdProceso(3).Visible = True
    If Not cmdProceso(3).Enabled Then cmdProceso(3).Enabled = True
Else
    If cmdProceso(3).Visible Then cmdProceso(3).Visible = False
End If
End Sub

Private Sub txtCampo_Change(Index As Integer)
If Index >= 5 And Index <= 6 Then
    Image2.Picture = Imagenes.ListImages(2).Picture
End If
End Sub

Private Sub txtCampo_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 1 And InStr("-", Chr(KeyAscii)) > 0 Then Exit Sub
KeyAscii = TeclaOprimida(txtcampo(Index), KeyAscii, txtcampo(Index).Tag, False)
'MsgBox "asd"
End Sub

Private Sub txtCampo_LostFocus(Index As Integer)
Dim adors As New ADODB.Recordset
If Mid(txtcampo(Index).Tag, 1, 1) = "f" Then
    If IsDate(txtcampo(Index).Text) Then
        d = CDate(txtcampo(Index).Text)
        txtcampo(Index).Text = Format(d, gsFormatoFecha)
        adors.Open "select sysdate from dual", gConSql, adOpenStatic, adLockReadOnly
        If Int(adors(0)) - Int(d) < 0 Then
            Call MsgBox("Fecha no válida. No se permite ingresar fecha mayor a la fecha actual (" & Format(adors(0), gsFormatoFecha) & ")", vbOKOnly + vbInformation, "")
            txtcampo(Index) = ""
            Exit Sub
        End If
    Else
        If Len(txtcampo(Index).Text) > 0 Then
            Call MsgBox("Fecha no válida. Verificar", vbOKOnly + vbInformation, "")
            txtcampo(Index) = ""
        End If
    End If
End If
End Sub

Sub ActualizaPendientes()
Dim adors As New ADODB.Recordset
adors.Open "{call p_seguimientopendientes}", gConSql, adOpenForwardOnly, adLockReadOnly
cmbPendientes.Clear
If adors.EOF Then
    cmbPendientes.AddItem "NO HAY PENDIENTES"
    cmbPendientes.ItemData(cmbPendientes.NewIndex) = -1
Else
    Do While Not adors.EOF
        cmbPendientes.AddItem adors(1)
        cmbPendientes.ItemData(cmbPendientes.NewIndex) = adors(0)
        adors.MoveNext
    Loop
End If
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
If cmdDescto.Visible Then cmdDescto.Visible = False
If iTipo = 0 Then 'Limpia todo y prepara un nuevo seguimiento
    txtOficioS.Enabled = True
    txtOficioE.Enabled = True
    txtExpediente.Text = ""
    txtOficioS.Text = ""
    txtOficioE.Text = ""
    cmbPendientes.Enabled = True
    cmdContinuar.Enabled = True
    txtExpediente.Enabled = True
    ListView1.ListItems.Clear
    For i = 0 To cmdProceso.UBound
        cmdProceso(i).Enabled = False
    Next
    For i = 0 To txtcampo.UBound
        txtcampo(i).Text = ""
    Next
    For i = 0 To cmbCampo.UBound
        cmbCampo(i).Clear
    Next
    cmdDif.Visible = False
ElseIf iTipo = 1 Then 'Limpia una nueva institución
    txtcampo(3).Text = ""
    cmbCampo(2).Clear
    cmbCampo(2).ListIndex = -1
    ListView1.ListItems.Clear
End If
End Sub

Private Sub txtExpediente_Change()
If Len(txtExpediente) > 0 Then
    If Len(txtOficioE.Text) > 0 Then
        txtOficioE.Text = ""
    End If
    If Len(txtOficioS.Text) > 0 Then
        txtOficioS.Text = ""
    End If
    If cmbPendientes.ListIndex >= 0 Then
        cmbPendientes.ListIndex = -1
    End If
End If
End Sub

Private Sub txtOficioE_Change()
If Len(txtOficioE.Text) > 0 Then
    If Len(txtExpediente.Text) > 0 Then
        txtExpediente.Text = ""
    End If
    If Len(txtOficioS.Text) > 0 Then
        txtOficioS.Text = ""
    End If
    If cmbPendientes.ListIndex >= 0 Then
        cmbPendientes.ListIndex = -1
    End If
End If

End Sub

Private Sub txtOficioS_Change()
If Len(txtOficioS.Text) > 0 Then
    If Len(txtExpediente.Text) > 0 Then
        txtExpediente.Text = ""
    End If
    If Len(txtOficioE.Text) > 0 Then
        txtOficioE.Text = ""
    End If
    If cmbPendientes.ListIndex >= 0 Then
        cmbPendientes.ListIndex = -1
    End If
End If

End Sub
