VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form ConsultaMasiva 
   Caption         =   "Consulta por datos seguimiento (relizados o en status)"
   ClientHeight    =   10710
   ClientLeft      =   3630
   ClientTop       =   5175
   ClientWidth     =   18480
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10710
   ScaleWidth      =   18480
   Begin VB.Frame Frame9 
      BackColor       =   &H00DBF5F9&
      Caption         =   "Area"
      Height          =   1095
      Left            =   9045
      TabIndex        =   69
      Top             =   1350
      Width           =   2355
      Begin VB.OptionButton opcArea 
         BackColor       =   &H00DBF5F9&
         Caption         =   "Grupo 2 (DSEF)"
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   8
         Top             =   810
         Width           =   1590
      End
      Begin VB.OptionButton opcArea 
         BackColor       =   &H00DBF5F9&
         Caption         =   "Grupo 1 (DSIF)"
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   540
         Width           =   1590
      End
      Begin VB.OptionButton opcArea 
         BackColor       =   &H00DBF5F9&
         Caption         =   "Ambos"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   270
         Value           =   -1  'True
         Width           =   1590
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Resultados de la busqueda:"
      Height          =   6180
      Left            =   0
      TabIndex        =   62
      Top             =   4500
      Width           =   18465
      Begin VB.CommandButton cmdResumen 
         Caption         =   ".."
         Height          =   420
         Left            =   17910
         TabIndex        =   48
         Top             =   5625
         Width           =   330
      End
      Begin VB.Frame FrameCaja 
         BackColor       =   &H80000005&
         Caption         =   "Operación de la caja"
         Enabled         =   0   'False
         Height          =   870
         Left            =   135
         TabIndex        =   63
         Top             =   5265
         Width           =   15405
         Begin VB.TextBox txtCajaDescrip 
            Height          =   525
            Left            =   7695
            MaxLength       =   1000
            MultiLine       =   -1  'True
            TabIndex        =   45
            Tag             =   "f"
            Top             =   315
            Width           =   5820
         End
         Begin VB.TextBox txtCajaNombre 
            Height          =   300
            Left            =   4275
            MaxLength       =   50
            TabIndex        =   44
            Tag             =   "f"
            Top             =   495
            Width           =   3390
         End
         Begin VB.ComboBox cmbCajaxAsig 
            DataField       =   "idoridir"
            Height          =   315
            ItemData        =   "ConsultaMasiva.frx":0000
            Left            =   270
            List            =   "ConsultaMasiva.frx":0002
            TabIndex        =   43
            ToolTipText     =   "Dirección General de Origen"
            Top             =   495
            Width           =   3945
         End
         Begin VB.CommandButton cmdCaja 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Asignar caja"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   13725
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   225
            Width           =   1500
         End
         Begin VB.Label Label6 
            BackColor       =   &H80000005&
            Caption         =   "Descripción de la caja:"
            Height          =   240
            Index           =   2
            Left            =   7695
            TabIndex        =   66
            Top             =   135
            Width           =   2985
         End
         Begin VB.Label Label6 
            BackColor       =   &H80000005&
            Caption         =   "Nombre de la caja:"
            Height          =   240
            Index           =   1
            Left            =   4275
            TabIndex        =   65
            Top             =   225
            Width           =   2985
         End
         Begin VB.Label Label6 
            BackColor       =   &H80000005&
            Caption         =   "Cajas Disponible en el procceso:"
            Height          =   240
            Index           =   0
            Left            =   270
            TabIndex        =   64
            Top             =   225
            Width           =   2985
         End
      End
      Begin VB.CommandButton cmdSeg 
         BackColor       =   &H0080C0FF&
         Caption         =   "Seguimiento Masivo"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   15705
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   5445
         Width           =   2040
      End
      Begin VB.ListBox listEncontrados 
         Height          =   645
         Left            =   135
         TabIndex        =   36
         Top             =   225
         Width           =   8295
      End
      Begin VB.TextBox txtResumen 
         Height          =   780
         Left            =   9990
         TabIndex        =   40
         Text            =   "Resumen"
         Top             =   135
         Width           =   8295
      End
      Begin VB.CommandButton cmdver 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Exportar a &Excel"
         Height          =   555
         Left            =   8505
         Picture         =   "ConsultaMasiva.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   225
         Width           =   1395
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4350
         Left            =   135
         TabIndex        =   42
         Top             =   945
         Width           =   18195
         _ExtentX        =   32094
         _ExtentY        =   7673
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   17
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "idAna"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Expediente"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Memorando"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Emplazamiento"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fecha"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Unidad"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Clase(Sector)"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Key             =   "Institución"
            Text            =   "Institución"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Proceso"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Actividad"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Tarea"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Responsable"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Status_Proceso"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Caja"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Status Tarea"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Tipo_Notif"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "idact_pen"
            Object.Width           =   882
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBF5F9&
      Caption         =   "Especifique los datos obligatorios sobre el Seguimiento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   54
      Top             =   1125
      Width           =   18510
      Begin VB.Frame Frame8 
         BackColor       =   &H00DBF5F9&
         Height          =   915
         Left            =   180
         TabIndex        =   67
         Top             =   315
         Width           =   1590
         Begin MSForms.OptionButton opcTipo 
            Height          =   330
            Index           =   0
            Left            =   90
            TabIndex        =   1
            Top             =   135
            Width           =   1230
            BackColor       =   14415353
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "2170;582"
            Value           =   "1"
            Caption         =   "Realizado"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.OptionButton opcTipo 
            Height          =   330
            Index           =   1
            Left            =   90
            TabIndex        =   2
            Top             =   540
            Width           =   1410
            BackColor       =   14415353
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "2487;582"
            Value           =   "0"
            Caption         =   "Estacionado"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
      End
      Begin VB.ComboBox cmbCaja 
         BackColor       =   &H8000000F&
         DataField       =   "idoridir"
         Height          =   315
         ItemData        =   "ConsultaMasiva.frx":0446
         Left            =   5310
         List            =   "ConsultaMasiva.frx":0448
         TabIndex        =   4
         ToolTipText     =   "Dirección General de Origen"
         Top             =   495
         Width           =   3540
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00DBF5F9&
         Caption         =   "Fecha de la tarea"
         Height          =   1410
         Left            =   11565
         TabIndex        =   57
         Top             =   45
         Width           =   6945
         Begin VB.OptionButton opcRango 
            BackColor       =   &H00DBF5F9&
            Caption         =   "Omitir"
            Height          =   285
            Index           =   5
            Left            =   5715
            TabIndex        =   68
            Top             =   270
            Value           =   -1  'True
            Width           =   1050
         End
         Begin VB.OptionButton opcRango 
            BackColor       =   &H00DBF5F9&
            Caption         =   "Anual"
            Height          =   285
            Index           =   0
            Left            =   315
            TabIndex        =   9
            Top             =   270
            Width           =   780
         End
         Begin VB.OptionButton opcRango 
            BackColor       =   &H00DBF5F9&
            Caption         =   "Bimestral"
            Height          =   285
            Index           =   1
            Left            =   1305
            TabIndex        =   10
            Top             =   270
            Width           =   1050
         End
         Begin VB.OptionButton opcRango 
            BackColor       =   &H00DBF5F9&
            Caption         =   "Mensual"
            Height          =   285
            Index           =   2
            Left            =   2475
            TabIndex        =   58
            Top             =   270
            Width           =   1005
         End
         Begin VB.OptionButton opcRango 
            BackColor       =   &H00DBF5F9&
            Caption         =   "Semanal"
            Height          =   285
            Index           =   3
            Left            =   3645
            TabIndex        =   11
            Top             =   270
            Width           =   1095
         End
         Begin VB.OptionButton opcRango 
            BackColor       =   &H00DBF5F9&
            Caption         =   "Otro"
            Height          =   285
            Index           =   4
            Left            =   4815
            TabIndex        =   12
            Top             =   270
            Width           =   600
         End
         Begin VB.TextBox txtFin 
            Height          =   300
            Left            =   1305
            TabIndex        =   15
            Tag             =   "f"
            Top             =   1065
            Width           =   4110
         End
         Begin VB.TextBox txtIni 
            Height          =   300
            Left            =   1320
            TabIndex        =   14
            Tag             =   "f"
            Top             =   705
            Width           =   4110
         End
         Begin VB.CommandButton cmdAntSig 
            Caption         =   "Sig >>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Index           =   1
            Left            =   5850
            TabIndex        =   16
            Top             =   585
            Width           =   855
         End
         Begin VB.CommandButton cmdAntSig 
            Caption         =   "<< Ant"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Index           =   0
            Left            =   5850
            TabIndex        =   17
            Top             =   945
            Width           =   855
         End
         Begin MSForms.Label Label5 
            Height          =   195
            Left            =   360
            TabIndex        =   60
            Top             =   1125
            Width           =   870
            BackColor       =   14415353
            Caption         =   "Al:"
            Size            =   "1535;344"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label4 
            Height          =   195
            Left            =   360
            TabIndex        =   59
            Top             =   720
            Width           =   870
            BackColor       =   14415353
            Caption         =   "Del:"
            Size            =   "1535;344"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.ComboBox cmbTarea 
         BackColor       =   &H8000000F&
         DataField       =   "idoridir"
         Height          =   315
         ItemData        =   "ConsultaMasiva.frx":044A
         Left            =   2070
         List            =   "ConsultaMasiva.frx":044C
         TabIndex        =   5
         ToolTipText     =   "Dirección General de Origen"
         Top             =   1035
         Width           =   5655
      End
      Begin VB.ComboBox cmbProceso 
         DataField       =   "idoridir"
         Height          =   315
         ItemData        =   "ConsultaMasiva.frx":044E
         Left            =   2070
         List            =   "ConsultaMasiva.frx":0450
         TabIndex        =   3
         ToolTipText     =   "Dirección General de Origen"
         Top             =   495
         Width           =   2730
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBF5F9&
         Caption         =   "Caja:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   5310
         TabIndex        =   61
         Top             =   270
         Width           =   360
      End
      Begin MSForms.Label Label3 
         Height          =   195
         Left            =   2025
         TabIndex        =   56
         Top             =   810
         Width           =   870
         BackColor       =   14415353
         Caption         =   "Tarea:"
         Size            =   "1535;344"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Left            =   2070
         TabIndex        =   55
         Top             =   270
         Width           =   870
         BackColor       =   14415353
         Caption         =   "Proceso:"
         Size            =   "1535;344"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBF5F9&
      Caption         =   "Especifique aqui criterios de búsqueda opcionales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   0
      TabIndex        =   13
      Top             =   2565
      Width           =   18465
      Begin VB.CommandButton cmdLimpiarRegs 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Limpiar Registros Encontrados"
         Height          =   600
         Left            =   14940
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1215
         Width           =   1590
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Limpia datos opcionales para Nueva búsqueda"
         Height          =   600
         Left            =   9675
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1215
         Width           =   2355
      End
      Begin VB.TextBox txtExp 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1305
         MaxLength       =   35
         TabIndex        =   18
         Tag             =   "c"
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox txtMemo 
         BackColor       =   &H8000000F&
         DataField       =   "n_cvepersona"
         Height          =   285
         Left            =   1305
         MaxLength       =   20
         TabIndex        =   19
         Tag             =   "c"
         ToolTipText     =   """Numero consecutivo de registro"""
         Top             =   675
         Width           =   3645
      End
      Begin VB.CommandButton cmdContinuar 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Buscar y Agregar"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   16785
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1215
         Width           =   1410
      End
      Begin VB.ComboBox cmbLey 
         BackColor       =   &H8000000F&
         DataField       =   "idoriuni"
         Height          =   315
         ItemData        =   "ConsultaMasiva.frx":0452
         Left            =   13500
         List            =   "ConsultaMasiva.frx":0468
         TabIndex        =   22
         ToolTipText     =   "Unidad de Origen"
         Top             =   270
         Width           =   4665
      End
      Begin VB.ComboBox cmbCausa 
         BackColor       =   &H8000000F&
         DataField       =   "idmat"
         Height          =   315
         ItemData        =   "ConsultaMasiva.frx":04B6
         Left            =   7065
         List            =   "ConsultaMasiva.frx":04B8
         TabIndex        =   21
         ToolTipText     =   "Materia de la Sanción"
         Top             =   765
         Width           =   5565
      End
      Begin VB.ComboBox cmbUnidad 
         BackColor       =   &H8000000F&
         DataField       =   "idoridir"
         Height          =   315
         ItemData        =   "ConsultaMasiva.frx":04BA
         Left            =   7065
         List            =   "ConsultaMasiva.frx":04BC
         TabIndex        =   20
         ToolTipText     =   "Dirección General de Origen"
         Top             =   315
         Width           =   5565
      End
      Begin VB.ComboBox cmbClase 
         BackColor       =   &H8000000F&
         Height          =   315
         ItemData        =   "ConsultaMasiva.frx":04BE
         Left            =   13500
         List            =   "ConsultaMasiva.frx":04C0
         TabIndex        =   23
         ToolTipText     =   "Clase de Institución"
         Top             =   720
         Width           =   4680
      End
      Begin VB.ComboBox cmbInst 
         BackColor       =   &H8000000F&
         Height          =   315
         ItemData        =   "ConsultaMasiva.frx":04C2
         Left            =   225
         List            =   "ConsultaMasiva.frx":04C4
         TabIndex        =   29
         ToolTipText     =   "Institución"
         Top             =   1575
         Width           =   9120
      End
      Begin VB.TextBox txtBuscarIF 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4725
         TabIndex        =   27
         Top             =   1215
         Width           =   3750
      End
      Begin VB.CommandButton cmdBusIF 
         Caption         =   "Sig"
         Height          =   330
         Left            =   8640
         TabIndex        =   28
         Top             =   1170
         Width           =   645
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00DBF5F9&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1350
         TabIndex        =   24
         Top             =   1125
         Width           =   2175
         Begin VB.OptionButton opcIF 
            BackColor       =   &H00DBF5F9&
            Caption         =   "Todas"
            Height          =   285
            Index           =   1
            Left            =   1215
            TabIndex        =   26
            Top             =   90
            Width           =   1005
         End
         Begin VB.OptionButton opcIF 
            BackColor       =   &H00DBF5F9&
            Caption         =   "Vigentes"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   25
            Top             =   90
            Value           =   -1  'True
            Width           =   1005
         End
      End
      Begin MSForms.CheckBox chkRegistrados 
         Height          =   375
         Left            =   12285
         TabIndex        =   31
         Top             =   1305
         Visible         =   0   'False
         Width           =   2625
         BackColor       =   14415353
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4630;661"
         Value           =   "0"
         Caption         =   "Registrados (en SIAM)"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2 
         BackColor       =   &H00DBF5F9&
         Caption         =   "Expediente:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   180
         TabIndex        =   51
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBF5F9&
         Caption         =   "Memorando:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   50
         Top             =   765
         Width           =   885
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBF5F9&
         Caption         =   "Cuasa de la sanción:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   5310
         TabIndex        =   49
         Top             =   765
         Width           =   1485
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBF5F9&
         Caption         =   "Ley:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   12825
         TabIndex        =   41
         Top             =   405
         Width           =   300
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBF5F9&
         Caption         =   "Área de origen:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   5400
         TabIndex        =   39
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label etiCombo 
         BackColor       =   &H00DBF5F9&
         Caption         =   "Sector (Clase Inst.):"
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   3
         Left            =   12690
         TabIndex        =   37
         Top             =   810
         Width           =   900
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBF5F9&
         Caption         =   "Institución:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   225
         TabIndex        =   35
         Top             =   1215
         Width           =   765
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBF5F9&
         Caption         =   "Busca IF:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   4005
         TabIndex        =   33
         Top             =   1260
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      Height          =   5280
      Left            =   0
      TabIndex        =   0
      Top             =   5400
      Width           =   18465
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00B1E8EF&
      Height          =   1155
      Left            =   0
      TabIndex        =   52
      Top             =   180
      Width           =   18465
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8676
         Top             =   144
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   17
         ImageHeight     =   17
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ConsultaMasiva.frx":04C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ConsultaMasiva.frx":088C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ConsultaMasiva.frx":0C52
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ConsultaMasiva.frx":1118
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ConsultaMasiva.frx":14DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ConsultaMasiva.frx":18A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ConsultaMasiva.frx":1C6A
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ConsultaMasiva.frx":2030
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ConsultaMasiva.frx":23F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ConsultaMasiva.frx":27BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ConsultaMasiva.frx":2B82
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   915
         Left            =   135
         Picture         =   "ConsultaMasiva.frx":2F48
         Stretch         =   -1  'True
         Top             =   135
         Width           =   1020
      End
      Begin VB.Label Eti 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00B1E8EF&
         Caption         =   "Módulo de Consulta por datos de Seguimiento (Status: Pendiente /  Realizado)"
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
         Height          =   690
         Index           =   4
         Left            =   4230
         TabIndex        =   53
         Top             =   360
         Width           =   11790
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "ConsultaMasiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msSol As String
Dim msAnt As String 'Contiene los idAnt para guardado masivo
Dim miColOrden As Long
Dim msClases(200) As String
Dim miCerrados As Long
Dim miAbiertos As Long
Dim msAna As String 'Contiene los idana de los registros seleccionados
Dim miPro As Integer 'id del proceso
Dim miCol As Integer 'Columna ordenada



Private Sub cmbCajaxAsig_Click()
Dim i As Integer, adors As ADODB.Recordset
If cmbCajaxAsig.ListIndex >= 0 Then
    If cmbCajaxAsig.ItemData(cmbCajaxAsig.ListIndex) = 0 Then
        txtCajaNombre.Text = ""
        txtCajaDescrip.Text = ""
        Exit Sub
    End If
    Set adors = New ADODB.Recordset
    adors.Open "select * from cajas where idcaja=" & cmbCajaxAsig.ItemData(cmbCajaxAsig.ListIndex), gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        txtCajaNombre.Text = IIf(IsNull(adors!nombre), "", adors!nombre)
        txtCajaDescrip.Text = IIf(IsNull(adors!descrip), "", adors!descrip)
    Else
        txtCajaNombre.Text = ""
        txtCajaDescrip.Text = ""
    End If
Else
    txtCajaNombre.Text = ""
    txtCajaDescrip.Text = ""
End If
End Sub

Private Sub cmbClase_Click()
ActualizaComboIF
End Sub


Private Sub cmbProceso_Click()
Dim adors As New ADODB.Recordset
If cmbProceso.ListIndex > 0 Then
    miPro = cmbProceso.ItemData(cmbProceso.ListIndex)
    adors.Open "{call paq_registro.tarea(" & miPro & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
    Call LlenaComboCursor(cmbTarea, adors)
    If adors.State > 0 Then adors.Close
    adors.Open "{call paq_registro.caja(" & miPro & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
    Call LlenaComboCursor(cmbCaja, adors)
Else
    'adors.Open "{call paq_registro.tarea(0)}", gConSql, adOpenForwardOnly, adLockReadOnly
    
    cmbTarea.Clear
    cmbCaja.Clear
End If
End Sub

Private Sub cmdAntSig_Click(Index As Integer)
Dim d As Date, i As Integer

If Not IsDate(txtIni.Text) Then
    txtIni.Text = Format(Now, "dd/mm/yyyy")
End If
If Not IsDate(txtFin.Text) Then
    txtFin.Text = Format(Now, "dd/mm/yyyy")
End If

If Index = 0 Then
    For i = 0 To opcRango.UBound
        If opcRango(i).Value Then Exit For
    Next
    d = CDate(txtIni.Text)
    Select Case i
    Case 0
        d = CDate("01/01/" & (Year(d) - 1))
        txtIni.Text = Format(d, "dd/mm/yyyy")
        d = DateAdd("yyyy", 1, d) - 1
        txtFin.Text = Format(d, "dd/mm/yyyy")
    Case 1
        d = DateAdd("m", -2, d)
        d = d - Day(d) + 1
        txtIni.Text = Format(d, "dd/mm/yyyy")
        d = DateAdd("m", 2, d)
        d = d - Day(d)
        txtFin.Text = Format(d, "dd/mm/yyyy")
    Case 2
        d = CDate(txtFin.Text)
        d = d - Day(d)
        txtFin.Text = Format(d, "dd/mm/yyyy")
        d = d - Day(d) + 1
        txtIni.Text = Format(d, "dd/mm/yyyy")
    Case 3
        d = CDate(txtFin.Text)
        d = d - 7
        txtFin.Text = Format(d, "dd/mm/yyyy")
        d = d - 4
        txtIni.Text = Format(d, "dd/mm/yyyy")
    Case 4
        d = d - 1
        txtIni.Text = Format(d, "dd/mm/yyyy")
        d = CDate(txtFin.Text) - 1
        txtFin.Text = Format(d, "dd/mm/yyyy")
    End Select
Else
    For i = 0 To opcRango.UBound
        If opcRango(i).Value Then Exit For
    Next
    d = CDate(txtIni.Text)
    Select Case i
    Case 0
        d = CDate("01/01/" & (Year(d) + 1))
        txtIni.Text = Format(d, "dd/mm/yyyy")
        d = DateAdd("yyyy", 1, d) - 1
        txtFin.Text = Format(d, "dd/mm/yyyy")
    Case 1
        d = DateAdd("m", 2, d)
        d = d - Day(d) + 1
        txtIni.Text = Format(d, "dd/mm/yyyy")
        d = DateAdd("m", 2, d)
        d = d - Day(d)
        txtFin.Text = Format(d, "dd/mm/yyyy")
    Case 2
        d = CDate(txtFin.Text)
        d = DateAdd("m", 2, d)
        d = d - Day(d)
        txtFin.Text = Format(d, "dd/mm/yyyy")
        d = d - Day(d) + 1
        txtIni.Text = Format(d, "dd/mm/yyyy")
    Case 3
        d = CDate(txtFin.Text)
        d = d + 7
        txtFin.Text = Format(d, "dd/mm/yyyy")
        d = d - 4
        txtIni.Text = Format(d, "dd/mm/yyyy")
    Case 4
        d = d + 1
        txtIni.Text = Format(d, "dd/mm/yyyy")
        d = CDate(txtFin.Text) + 1
        txtFin.Text = Format(d, "dd/mm/yyyy")
    End Select
End If
End Sub

Private Sub cmdBusIF_Click()
Dim i As Long, iPos As Long
If Len(txtBuscarIF.Text) > 0 And cmbInst.ListCount > 0 Then
    iPos = cmbInst.ListIndex
    If iPos = cmbInst.ListCount - 1 Then
        i = -1
    Else
        i = BuscaCombo(cmbInst, txtBuscarIF.Text, 0, True, 0, iPos + 1)
    End If
    If i >= 0 Then
        cmbInst.ListIndex = i
    ElseIf iPos >= 0 Then
        cmbInst.ListIndex = -1
    End If
End If
End Sub

Private Sub cmdCaja_Click()
Dim adors As ADODB.Recordset, sReg As String
Dim sNombre As String, sDescrip As String

On Error GoTo salir:

If cmbCajaxAsig.ListIndex < 0 Then
    Call MsgBox("Debe elejir una opción de caja disponible", vbOKOnly + vbInformation, "Validación de datos")
    Exit Sub
End If
If Len(Trim(txtCajaNombre.Text)) = 0 Then
    Call MsgBox("El nombre de la caja es un dato requerido", vbOKOnly + vbInformation, "Validación de datos")
    Exit Sub
End If
Set adors = New ADODB.Recordset
If cmbCajaxAsig.ItemData(cmbCajaxAsig.ListIndex) = 0 Then 'Valida que sea nuevo el nombre de la caja
    sNombre = Trim(Replace(txtCajaNombre.Text, "'", "''"))
    adors.Open "select count(*) from cajas where nombre='" & sNombre & "'", gConSql, adOpenStatic, adLockReadOnly
    If adors(0) > 0 Then
        Call MsgBox("El nombre de la caja ya existe. Favor de verifcar", vbOKOnly + vbInformation, "Validación de datos")
        Exit Sub
    End If
    If adors.State Then adors.Close
End If
If Len(msAna) <= 1 Then
    cmdCaja.Enabled = False
    FrameCaja.Enabled = False
    Exit Sub
End If
If MsgBox("Esta seguro de asignar la " & IIf(cmbCajaxAsig.ItemData(cmbCajaxAsig.ListIndex) = 0, "nueva ", "") & "caja a los expedientes seleccionados", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmación") = vbNo Then
    Exit Sub
End If

If cmbCajaxAsig.ItemData(cmbCajaxAsig.ListIndex) > 0 Then
    sNombre = cmbCajaxAsig.List(cmbCajaxAsig.ListIndex)
End If

adors.Open "select f_analisis_idreg(id) as idreg from análisis where id in (" & msAna & ")", gConSql, adOpenStatic, adLockReadOnly
Do While Not adors.EOF
    sReg = sReg & adors(0) & ","
    adors.MoveNext
Loop
If InStr(sReg, ",") Then
    sReg = Mid(sReg, 1, Len(sReg) - 1)
End If
If adors.State Then adors.Close
adors.Open "{call paq_analisis.creacaja('" & sReg & "','" & sNombre & "','" & Replace(txtCajaDescrip.Text, "'", "''") & "'," & miPro & "," & giUsuario & ") }", gConSql, adOpenForwardOnly, adLockReadOnly
If adors(0) <= 0 Then
    Call MsgBox("No se realizó la asignación de la caja" & adors(1), vbOKOnly + vbInformation, "Aviso")
    Exit Sub
Else
    Call MsgBox("Se realizó la asignación correctamente " & adors(1), vbOKOnly + vbInformation, "Aviso")
End If
For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(i).Checked Then 'Actualiza la caja en la columna 12
        ListView1.ListItems(i).SubItems(13) = txtCajaNombre.Text
        ListView1.ListItems(i).Checked = False
    End If
Next
Exit Sub

salir: 'Atrapa excepción

Me.MousePointer = 0
Dim yError As Long
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

Private Sub cmdContinuar_Click()
Dim adors As New ADODB.Recordset
Dim iTip As Integer, iPro As Integer, iTar As Integer, sIni As String, sFin As String, iCaja As Integer
Dim iCla As Long, iIns As Long, iLey As Long, iSan As Long, iUni As Long
Dim i As Long, n As Long
Dim iGrupo As Integer
On Error GoTo salir:
If opcTipo(0).Value Then
    iTip = 1
Else
    iTip = 2
End If
If cmbProceso.ListIndex >= 0 Then
    iPro = cmbProceso.ItemData(cmbProceso.ListIndex)
End If
If cmbTarea.ListIndex >= 0 Then
    iTar = cmbTarea.ItemData(cmbTarea.ListIndex)
End If
If IsDate(txtIni.Text) Then
    sIni = txtIni.Text
End If
If IsDate(txtFin.Text) Then
    sFin = txtFin.Text
End If
If cmbClase.ListIndex >= 0 Then
    iCla = cmbClase.ItemData(cmbClase.ListIndex)
End If
If cmbInst.ListIndex >= 0 Then
    iIns = cmbInst.ItemData(cmbInst.ListIndex)
End If
If cmbCaja.ListIndex >= 0 Then
    iCaja = cmbCaja.ItemData(cmbCaja.ListIndex)
End If
If cmbLey.ListIndex >= 0 Then
    iLey = cmbLey.ItemData(cmbLey.ListIndex)
End If
If cmbCausa.ListIndex >= 0 Then
    iSan = cmbCausa.ItemData(cmbCausa.ListIndex)
End If
If cmbUnidad.ListIndex >= 0 Then
    iUni = cmbUnidad.ItemData(cmbUnidad.ListIndex)
End If
'Validación
If Not opcTipo(0).Value And Not opcTipo(1).Value Then
    MsgBox "El Staus de la tarea es un dato requerido (Estacionado  / Relaizado)", vbInformation + vbOKOnly, "Validación"
    Exit Sub
End If
If iPro = 0 Then
    MsgBox "El Proceso de la tarea es un dato requerido", vbInformation + vbOKOnly, "Validación"
    Exit Sub
End If
If opcRango(5).Value And iTip = 1 Then
    MsgBox "La consulta de tareas realizadas no puede omitir el rango de fechas", vbInformation + vbOKOnly, "Validación"
    Exit Sub
End If
If opcArea(1).Value Then
    iGrupo = 1
ElseIf opcArea(2).Value Then
    iGrupo = 2
End If
'ListView1.ListItems.Clear
Me.MousePointer = 11
If adors.State Then adors.Close
'If chkRegistrados.Value Then
'    adors.Open "{call paq_registro.busca_datosReg('" & txtExp.Text & "','" & txtMemo.Text & "'," & iUni & "," & iCla & "," & iIns & "," & iInc & "," & iSan & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
'    If Not adors.EOF And ListView1.ListItems.Count = 0 Then
'        'If Frame3.Enabled Then Frame3.Enabled = False
'    End If
'Else
    adors.Open "{call paq_seguimiento.busca_datosXSeg(" & iTip & "," & iGrupo & "," & iPro & "," & iTar & ",'" & sIni & "','" & sFin & "','" & txtExp.Text & "','" & txtMemo.Text & "'," & iCaja & "," & iUni & "," & iLey & "," & iSan & "," & iCla & "," & iIns & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
    If Not adors.EOF Then
        If Not Frame3.Enabled Then Frame3.Enabled = True
    End If
'End If

If ListView1.Sorted Then  ' VERIFICA SI ESTÁ ORDENADO PARA QUITAR EL ORDENAMIENTO
    ListView1.Sorted = False
End If

i = ListView1.ListItems.Count + 1
Do While Not adors.EOF
    If InStr(msSol, "|" & adors(0) & "|") = 0 Then
        msSol = msSol & adors(0) & "|"
        ListView1.ListItems.Add i, , adors(0) 'idana
        ListView1.ListItems(i).SubItems(1) = IIf(IsNull(adors(1)), "", adors(1)) 'Expediente
        ListView1.ListItems(i).SubItems(2) = IIf(IsNull(adors(2)), "", adors(2)) 'Oficio Emplazamiento
        ListView1.ListItems(i).SubItems(3) = IIf(IsNull(adors(3)), "", adors(3)) 'Memorando
        ListView1.ListItems(i).SubItems(4) = IIf(IsNull(adors(4)), "", adors(4)) 'Fecha Tarea
        ListView1.ListItems(i).SubItems(5) = IIf(IsNull(adors(5)), "", adors(5)) 'Unidad
        ListView1.ListItems(i).SubItems(6) = IIf(IsNull(adors(6)), "", adors(6)) 'Clase
        ListView1.ListItems(i).SubItems(7) = IIf(IsNull(adors(7)), "", adors(7)) 'Institución
        ListView1.ListItems(i).SubItems(8) = IIf(IsNull(adors(8)), "", adors(8)) 'Proceso
        ListView1.ListItems(i).SubItems(9) = IIf(IsNull(adors(9)), "", adors(9)) 'Actividad
        ListView1.ListItems(i).SubItems(10) = IIf(IsNull(adors(10)), "", adors(10)) 'Tarea
        ListView1.ListItems(i).SubItems(11) = IIf(IsNull(adors(18)), "", adors(18)) 'Responsable
        ListView1.ListItems(i).SubItems(12) = IIf(IsNull(adors(11)), "", adors(11)) 'Status_proceso
        ListView1.ListItems(i).SubItems(13) = IIf(IsNull(adors(12)), "", adors(12)) 'Caja
        ListView1.ListItems(i).SubItems(14) = IIf(IsNull(adors(13)), "", adors(13)) 'Status Tarea
        ListView1.ListItems(i).SubItems(15) = IIf(IsNull(adors(14)), "", adors(14)) 'TipoNotif
        ListView1.ListItems(i).SubItems(16) = IIf(IsNull(adors(15)), "", adors(15) & "|" & adors(16)) 'idant & idact pend
        ListView1.ListItems(i).Tag = adors(17) 'Guarda el id del seguimiento del registro
        If InStr(ListView1.ListItems(i).SubItems(12), "Pend") > 0 Then
            miAbiertos = miAbiertos + 1
        Else
            miCerrados = miCerrados + 1
        End If
        i = i + 1
        n = n + 1
    End If
    adors.MoveNext
Loop
Me.MousePointer = 0
If n > 0 Then 'Se agregaron n registros se obtiene la descripción de la consulta y se agrega en el combo
    If adors.State Then adors.Close
    adors.Open "select paq_seguimiento.busca_datosXSeg(" & iTip & "," & iGrupo & "," & iPro & "," & iTar & ",'" & sIni & "','" & sFin & "','" & txtExp.Text & "','" & txtMemo.Text & "'," & iCaja & "," & iUni & "," & iLey & "," & iSan & "," & iCla & "," & iIns & ") from dual", gConSql, adOpenForwardOnly, adLockReadOnly
    listEncontrados.AddItem adors(0), listEncontrados.ListCount
    listEncontrados.ItemData(listEncontrados.ListCount - 1) = n
End If
ActualizaResumen
MsgBox n & " registro(s) agregado(s)", vbOKOnly + vbInformation, ""
Exit Sub
salir:
Me.MousePointer = 0
Dim yError As Long
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

Sub ActualizaResumen()
If miAbiertos + miCerrados > 0 Then
    txtResumen.Text = "Asuntos listados: " & (miAbiertos + miCerrados) & " .- Cerrados: " & miCerrados & " (" & Format((miCerrados * 100 / (miCerrados + miAbiertos)), "##.##") & " %)" & " Pendientes: " & miAbiertos & " (" & Format((miAbiertos * 100 / (miCerrados + miAbiertos)), "##.##") & " %)"
Else
    txtResumen.Text = "Asuntos listados: Ninguno"
End If
End Sub

Private Sub cmdLimpiar_Click()
txtExp.Text = ""
txtMemo.Text = ""
cmbUnidad.Text = ""
cmbClase.Text = ""
cmbInst.Text = ""
cmbCausa.Text = ""
cmbLey.Text = ""
cmbCaja.Text = ""
cmbTarea.Text = ""
If cmbUnidad.ListCount >= 0 Then
    cmbUnidad.ListIndex = -1
End If
If cmbLey.ListCount >= 0 Then
    cmbLey.ListIndex = -1
End If
If cmbCausa.ListCount >= 0 Then
    cmbCausa.ListIndex = -1
End If
If cmbClase.ListCount >= 0 Then
    cmbClase.ListIndex = -1
End If
If cmbInst.ListCount > 0 Then
    cmbInst.ListIndex = -1
End If
If cmbCaja.ListCount > 0 Then
    cmbCaja.ListIndex = -1
End If
If cmbTarea.ListCount > 0 Then
    cmbTarea.ListIndex = -1
End If
End Sub

Private Sub cmdLimpiarRegs_Click()
If ListView1.ListItems.Count > 0 Then
    If MsgBox("Estás seguro de quitar todos los registros encontrados en previas búsquedas", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmación") = vbNo Then
        Exit Sub
    End If
    ListView1.ListItems.Clear
    msSol = "|"
    listEncontrados.Clear
    txtResumen.Text = ""
    miAbiertos = 0
    miCerrados = 0
    cmdSeg.Enabled = False
    cmdCaja.Enabled = False
End If
txtCajaNombre.Text = ""
txtCajaDescrip.Text = ""
End Sub


Function Busca_Bin(iDato As Long, sDatos() As String) As Long
Dim i As Long, iMax As Long, iMin As Long
iMax = UBound(sDatos)
iAnt = iMax
Busca_Bin = -1
Do While iMin <> iMax
    If iMax - iMin = 1 Then
        If i = iMin Then
            i = iMax
        Else
            i = iMin
        End If
    Else
        i = iMin + Round((iMax - iMin + 0.0001) / 2, 0)
    End If
    If i > iMax Or i < iMin Then
        Exit Do
    End If
    If Val(sDatos(i)) = iDato Then
        Busca_Bin = i
        Exit Function
    End If
    If Val(sDatos(i)) < iDato Then
        iMin = i
    Else
        iMax = i
    End If
Loop
End Function


Private Sub cmdResumen_Click()
Dim i As Integer, s As String, sM As String
s = "-898989"
If miCol > 1 Then
    For i = 1 To ListView1.ListItems.Count
        If s <> ListView1.ListItems(i).SubItems(miCol - 1) Then 'Asigna en mensaje
            If m > 0 Then
                sM = sM & s & " : " & m & " (" & Format((m / ListView1.ListItems.Count), "###.## %") & ")" & Chr(13) & Chr(10)
                m = 0
            End If
            s = ListView1.ListItems(i).SubItems(miCol - 1)
        End If
        m = m + 1
    Next
    If m > 0 Then
        sM = sM & s & " : " & m & " (" & Format((m / ListView1.ListItems.Count), "###.## %") & ")"
    End If
    Call MsgBox(sM, vbOKOnly, "RESUMEN")
End If
End Sub

Private Sub cmdSeg_Click()
Dim i As Integer, sSeg As String, iAct As Integer, iAcc As Integer, s As String
Dim lAna As Long, n As Integer, sExp As String
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            If InStr("," & sSeg, "," & Val(ListView1.ListItems(i).Tag) & ",") = 0 Then
                s = ListView1.ListItems(i).SubItems(16)
                
                sExp = sExp & ListView1.ListItems(i).SubItems(1) & ", "

                sSeg = sSeg & Val(s) & ","
                If iAct = 0 Then
                    iAct = Val(Mid(s, InStr(s, "|") + 1))
                End If
                n = n + 1
            End If
            If lAna = 0 Then
                lAna = Val(ListView1.ListItems(i).Text)
            End If
        End If
    Next
    If InStr(sExp, ",") Then
        sExp = Mid(sExp, 1, Len(sExp) - 2)
    End If
    If iAct > 0 And InStr(sSeg, ",") > 0 And lAna > 0 And n > 0 Then
        sSeg = Mid(sSeg, 1, Len(sSeg) - 1)
        With Actividades
            .miTarea = 0
            gs = "no iniciar var"
            .miMasivo = 1
            .msSeg = sSeg
            .mlAnálisis = lAna
            '.mlAnt = lSeg
            .miActividad = iAct
            .mlSeguimiento = 0
            .msObservaciones = ""
            .yTipoOperación = 1 'Agregar
            .TreeView1.Enabled = False
            .mdFecha = CDate("01/01/1900")
            .Caption = "Registro MASIVO de Actividades"
            '.msProgResp = sProg
            .txtExp.Text = "Varios (" & n & "): " & sExp
            .Show vbModal
            If gs = "OK" Then
                MsgBox "Se quitaran de la consulta los registros afectados", vbOKOnly + vbInformation, "Aviso"
                n = 0
                For i = ListView1.ListItems.Count To 1 Step -1
                    If ListView1.ListItems(i).Checked Then
                        ListView1.ListItems.Remove (i)
                        n = n + 1
                    End If
                Next
                miAbiertos = miAbiertos - n
                ActualizaResumen
                cmdSeg.Enabled = False
            End If
        End With
    Else
        cmdSeg.Enabled = False
    End If

End Sub

Private Sub cmdver_Click()
Dim Hoja As Excel.Worksheet
Dim LibroExcel As Excel.Workbook
Dim ApExcel As Excel.Application
Dim l As Long, i As Integer, yErr As Byte
On Error GoTo ErrArchivo:
If ListView1.ListItems.Count > 0 Then
    Set ApExcel = CreateObject("Excel.Application") 'Método CreateObject y Application
    Set LibroExcel = ApExcel.Workbooks.Add   'con .Add añadimso Libros de trabajo de la aplicacion
    Set Hoja = LibroExcel.Worksheets(1)      'referenciado la primera hoja del libro de trabajo
    Hoja.Activate    'Activando la hoja
    ApExcel.Visible = True  'Hacemos visible la aplicación
    Hoja.Cells(1, 1).Value = "Criterios de la búsqueda"
    For Y = 0 To listEncontrados.ListCount - 1
        Hoja.Cells(1, Y + 2).Value = listEncontrados.List(Y) & " : " & listEncontrados.ItemData(Y)
    Next
    For Y = 1 To ListView1.ColumnHeaders.Count
        Hoja.Cells(2, Y).Value = ListView1.ColumnHeaders(Y).Text
    Next
    For i = 1 To ListView1.ListItems.Count
        Hoja.Cells(i + 2, 1).Value = ListView1.ListItems(i).Text
        For Y = 1 To ListView1.ColumnHeaders.Count - 1
            Hoja.Cells(i + 2, Y + 1).Value = ListView1.ListItems(i).SubItems(Y)
        Next
    Next
Else
    MsgBox "No se encontraron registros...", vbOKOnly + vbInformation, "Información"
End If
Exit Sub
ErrArchivo:
bExportar = False
If Err.Number = 384 Then
    Resume Next
End If
yErr = MsgBox("Error: " + Err.Description, vbAbortRetryIgnore + vbCritical, "Error no esperado (" + Str(Err.Number) + ")")
If yErr = vbCancel Then
    Exit Sub
ElseIf yErr = vbRetry Then
    Resume
ElseIf yErr = vbIgnore Then
    Resume Next
End If
End Sub

Private Sub Command1_Click()
End Sub

Private Sub Form_Load()
Dim adors As New ADODB.Recordset
adors.Open "{call paq_registro.proceso()}", gConSql, adOpenForwardOnly, adLockReadOnly
Call LlenaComboCursor(cmbProceso, adors)
If adors.State Then adors.Close
adors.Open "{call paq_registro.unidad(0)}", gConSql, adOpenForwardOnly, adLockReadOnly
Call LlenaComboCursor(cmbUnidad, adors)
If adors.State Then adors.Close
adors.Open "{call paq_registro.clase(0)}", gConSql, adOpenForwardOnly, adLockReadOnly
Call LlenaComboCursor(cmbClase, adors)
If adors.State Then adors.Close
adors.Open "{call paq_registro.ley()}", gConSql, adOpenForwardOnly, adLockReadOnly
Call LlenaComboCursor(cmbLey, adors)
msSol = "|"
cmdLimpiar_Click
If adors.State Then adors.Close
adors.Open "select id,descripción from claseinstitución order by 1", gConSql, adOpenForwardOnly, adLockReadOnly
Do While Not adors.EOF 'Agrega las clases en el arreglo
    If adors(0) < 200 And adors(0) > 0 Then
        msClases(adors(0)) = adors(1)
    End If
    adors.MoveNext
Loop
cmbLey.Text = ""
cmbLey.ListIndex = -1
cmbUnidad.Text = ""
cmbUnidad.ListIndex = -1
cmbClase.Text = ""
cmbClase.ListIndex = -1
cmbLey.Text = ""
cmbLey.ListIndex = -1
If adors.State Then adors.Close
adors.Open "select f_usuariosis_idgpo(" & giUsuario & ") from dual", gConSql, adOpenForwardOnly, adLockReadOnly
If adors(0) = 1 Then
    opcArea(1).Value = True
ElseIf adors(0) = 2 Then
    opcArea(2).Value = True
End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ListView1.SortKey = ColumnHeader.Index - 1 'Indico al ListView que ordene según los datos de la columna 1 (esta propiedad utiliza un valor que es igual al Indice de la columna - 1)
miCol = ColumnHeader.Index
If miColOrden = ColumnHeader.Index Then
    ListView1.SortOrder = lvwDescending ' Ordena en forma descendente
    miColOrden = 0
Else
    ListView1.SortOrder = lvwAscending ' Ordena en forma ascendente
    miColOrden = ColumnHeader.Index
End If
ListView1.Sorted = True ' con esto se ordena la lista.
End Sub

'Registro de actividad pendiente del asunto seleccionado
Private Sub ListView1_DblClick()
Dim i As Long, lSeg As Long, adors As New ADODB.Recordset, lAna As Long, sProg As String, iAct As Integer
Dim s As String
i = ListView1.SelectedItem.Index
If InStr(ListView1.ListItems(i).SubItems(12), "Pendiente") > 0 And Val(ListView1.ListItems(i).SubItems(16)) > 0 Then
    lAna = Val(ListView1.ListItems(i).Text)
    s = ListView1.ListItems(i).SubItems(16)
    lSeg = Val(s)
    If InStr(s, "|") > 0 Then
        iAct = Val(Mid(s, InStr(s, "|") + 1))
    End If
    'adors.Open "select s.idana,'('||to_char(sp.fecha,'dd/mm/yyyy hh24:mi')||') ('||paq_conceptos.responsable(sp.idusi)||')' as prog from seguimiento s left join seguimientoprog sp on s.id=sp.idant where s.id=" & lSeg, gConSql, adOpenForwardOnly, adLockReadOnly
    'If Not adors.EOF Then
    '    lAna = adors(0)
    '    If InStr(adors(1), "/") > 10 Then
    '        sProg = adors(1)
    '    End If
    'Else
    '    Exit Sub
    'End If
    If lAna > 0 And iAct > 0 And lSeg > 0 Then
        With Actividades
            .miTarea = 0
            .mlAnálisis = lAna
            .mlAnt = lSeg
            .miActividad = iAct
            .mlSeguimiento = 0
            .msObservaciones = ""
            .yTipoOperación = 1 'Agregar
            .mdFecha = CDate("01/01/1900")
            '.msProgResp = sProg
            gs = "no iniciar var"
            .miMasivo = 0
            .Show vbModal
            If gs = "OK" Then
                ListView1.ListItems.Remove (i)
                miAbiertos = miAbiertos - 1
                ActualizaResumen
            End If
        End With
    End If
End If
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim bChek As Boolean, i As Integer, bMas As Boolean, iAct As Integer, s As String, iA As Integer
Dim adors As New ADODB.Recordset, bCaja As Boolean, iCaja As Integer, sCajas  As String
Dim sNotif As String
'Replica el chek a los registros de la lista seleccionados
If Item.Checked Then
    bChek = True
End If
bMas = ListView1.ListItems.Count > 1
bCaja = True
msAnt = ""
msAna = ""
For i = 1 To ListView1.ListItems.Count 'Hace el barrido de todos los registros
    If ListView1.ListItems(i).Selected Then 'Está seleccionado replica el check
        If ListView1.ListItems(i).Checked <> bChek Then
            ListView1.ListItems(i).Checked = bChek
        End If
    End If
    If ListView1.ListItems(i).Checked Then
        s = ListView1.ListItems(i).SubItems(16) 'idant | idact Pendiente
        msAnt = msAnt & Val(s) & ","
        msAna = msAna & Val(ListView1.ListItems(i).Text) & ","
        If Len(sNotif) = 0 Then
            sNotif = ListView1.ListItems(i).SubItems(15)
        End If
        If bMas Then 'Validación para seguimiento masivo
            If InStr(s, "|") > 0 Then
                iA = Val(Mid(s, InStr(s, "|") + 1))
                If iAct = 0 Then
                    iAct = iA
                    adors.Open "select count(*) from seg_accion where idact=" & iAct, gConSql, adOpenStatic, adLockReadOnly
                    If adors(0) = 0 Then
                        bMas = False
                    End If
                End If
                If iAct <> iA Or InStr("|5|107|109|112|113|114|116|93|106|111|115|", "|" & iAct & "|") > 0 And InStr("sine,estrados electrónicos", LCase(sNotif)) > 0 Then 'Valida no sea notificación y no sea dif la actividad
                    bMas = False
                End If
            Else
                bMas = False
            End If
        End If
    End If
    If ListView1.ListItems(i).Checked And bCaja Then 'Validación para asignación de caja
        If bCaja And InStr(ListView1.ListItems(i).SubItems(12), "Pendient") = 0 Then
            bCaja = False
        ElseIf bCaja Then
            If InStr(sCajas, "|" & ListView1.ListItems(i).SubItems(13)) = 0 Then
                sCajas = "|" & sCajas & ListView1.ListItems(i).SubItems(13)
            End If
        End If
    End If
Next
If InStr(msAna, ",") Then
    msAna = Mid(msAna, 1, Len(msAna) - 1)
End If
If InStr(msAnt, ",") Then
    msAnt = Mid(msAnt, 1, Len(msAnt) - 1)
End If
cmdSeg.Enabled = bMas
If bCaja Then
    FrameCaja.Enabled = True
    cmdCaja.Enabled = True
    If InStrRev(sCajas, "|") > 1 Then 'Varias cajas
        cmbCajaxAsig.ListIndex = -1
        cmbCajaxAsig.Text = ""
        txtCajaNombre = "VARIAS CAJAS"
        txtCajaDescrip = "VARIAS CAJAS"
    ElseIf InStrRev(sCajas, "|") = 1 Then 'Una caja
        cmbCajaxAsig.Clear
        cmbCajaxAsig.AddItem "*** NUEVA CAJA ***", 0
        cmbCajaxAsig.ItemData(0) = 0
        LlenaCombo cmbCajaxAsig, "select idcaja, nombre from cajas where idpro=" & miPro & " and cerrado=0", "", True, True
        If adors.State Then adors.Close
        adors.Open "select * from cajas where nombre='" & Mid(sCajas, 2) & "'", gConSql, adOpenStatic, adLockReadOnly
        i = BuscaCombo(cmbCajaxAsig, Mid(sCajas, 2), False)
        If i > 0 Then
            cmbCajaxAsig.ListIndex = i
        Else
            cmbCajaxAsig.ListIndex = -1
            cmbCajaxAsig_Click
        End If
    Else 'ninguna caja
    End If
Else
    FrameCaja.Enabled = False
    cmdCaja.Enabled = False
End If
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long
Dim n As Long
'If KeyCode = 46 Then 'Borra registro seleccionados
'    For i = ListView1.ListItems.Count To 1 Step -1
'        If ListView1.ListItems(i).Selected Then
'            msSol = Replace(msSol, "|" & ListView1.ListItems(i).Text & "|", "|")
'            ListView1.ListItems.Remove (i)
'            n = n + 1
'        End If
'    Next
'    MsgBox "Se eliminaron " & n & " Registros", vbOKOnly, ""
'End If
End Sub



Private Sub opcIF_Click(Index As Integer)
ActualizaComboIF
End Sub


Sub ActualizaComboIF()
Dim adors As New ADODB.Recordset
If cmbClase.ListIndex > 0 Then
    adors.Open "{call paq_registro.institucion(" & IIf(opcIF(0).Value, 0, 1) & "," & cmbClase.ItemData(cmbClase.ListIndex) & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
Else
    adors.Open "{call paq_registro.institucion(" & IIf(opcIF(0).Value, 0, 1) & ",0)}", gConSql, adOpenForwardOnly, adLockReadOnly
End If
Call LlenaComboCursor(cmbInst, adors)
End Sub


Private Sub opcRango_Click(Index As Integer)
Dim d As Date
Select Case Index
Case 0
    d = CDate("01/01/" & (Year(Date) - 1))
    txtIni.Text = Format(d, "dd/mm/yyyy")
    d = DateAdd("yyyy", 1, d)
    d = d - 1
    txtFin.Text = Format(d, "dd/mm/yyyy")
Case 1
    d = DateAdd("m", IIf(Month(Date) Mod 2 = 0, -3, -2), Date)
    d = d - Day(d) + 1
    txtIni.Text = Format(d, "dd/mm/yyyy")
    d = DateAdd("m", 2, d)
    d = d - Day(d)
    txtFin.Text = Format(d, "dd/mm/yyyy")
Case 2
    d = Date - Day(Date)
    txtFin.Text = Format(d, "dd/mm/yyyy")
    d = d - Day(d) + 1
    txtIni.Text = Format(d, "dd/mm/yyyy")
Case 3
    d = Date - Weekday(Date, vbSaturday)
    txtFin.Text = Format(d, "dd/mm/yyyy")
    d = d - 4
    txtIni.Text = Format(d, "dd/mm/yyyy")
Case 4
    If Not txtIni.Enabled Then
        txtIni.Enabled = True
        txtFin.Enabled = True
    End If
    If txtIni.Text = "" Then
        txtIni = Format(Now, "dd/mm/yyyy")
    End If
    If txtFin.Text = "" Then
        txtFin = Format(Now, "dd/mm/yyyy")
    End If
Case 5
    txtIni.Text = ""
    txtFin.Text = ""
End Select
If Index < 4 And txtIni.Enabled Then
    txtIni.Enabled = False
    txtFin.Enabled = False
End If
End Sub

Sub ValidaSeleccionados()
Dim i As Integer
For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(i).Checked Then
        If ListView1.ListItems(i).SubItems(14) <> "Estacionado" Then
            cmdSeg.Enabled = False
            Exit Sub
        End If
    End If
Next
End Sub


