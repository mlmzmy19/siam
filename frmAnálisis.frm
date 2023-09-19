VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form Análisis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Análisis"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   15750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   15750
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1995
      Left            =   90
      TabIndex        =   51
      Top             =   1260
      Width           =   15630
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "Nombre"
         ForeColor       =   &H00808080&
         Height          =   1095
         Index           =   2
         Left            =   10500
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Tag             =   "c"
         ToolTipText     =   "Datos de la Documento de la Solicitud"
         Top             =   495
         Width           =   5000
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "Nombre"
         ForeColor       =   &H00808080&
         Height          =   1095
         Index           =   1
         Left            =   5280
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Tag             =   "c"
         ToolTipText     =   "Institución y nombre del Usuario"
         Top             =   495
         Width           =   5000
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "Nombre"
         ForeColor       =   &H00808080&
         Height          =   1095
         Index           =   0
         Left            =   45
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Tag             =   "c"
         ToolTipText     =   "Datos del origen de la Solicitud"
         Top             =   495
         Width           =   5000
      End
      Begin VB.ComboBox cmbCampo 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   0
         ItemData        =   "frmAnálisis.frx":0000
         Left            =   1080
         List            =   "frmAnálisis.frx":0002
         TabIndex        =   52
         ToolTipText     =   "Institución"
         Top             =   1620
         Width           =   14385
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Documento de la solicitud:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   10530
         TabIndex        =   56
         Top             =   270
         Width           =   1875
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Institución / Nombre(s):"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   5310
         TabIndex        =   55
         Top             =   270
         Width           =   3465
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Origen de la solicitud:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   54
         Top             =   270
         Width           =   1515
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Institución:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   53
         Top             =   1710
         Width           =   765
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6585
      Left            =   45
      TabIndex        =   7
      Top             =   3240
      Width           =   15675
      _ExtentX        =   27649
      _ExtentY        =   11615
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "PROCEDENTES"
      TabPicture(0)   =   "frmAnálisis.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "IMPROCEDENTES"
      TabPicture(1)   =   "frmAnálisis.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   5460
         Left            =   -74880
         TabIndex        =   60
         Top             =   405
         Width           =   15225
         Begin VB.TextBox txtCampo 
            BackColor       =   &H8000000F&
            DataField       =   "Nombre"
            Enabled         =   0   'False
            Height          =   465
            Index           =   8
            Left            =   1755
            MaxLength       =   250
            TabIndex        =   33
            Tag             =   "c"
            ToolTipText     =   "Fecha de Emplazamiento"
            Top             =   720
            Width           =   10065
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFFFFF&
            Height          =   5130
            Left            =   13365
            TabIndex        =   69
            Top             =   225
            Width           =   1300
            Begin VB.CommandButton cmdProcesoImp 
               Caption         =   "&Borra oficio"
               Enabled         =   0   'False
               Height          =   375
               Index           =   5
               Left            =   45
               TabIndex        =   46
               Top             =   4680
               Width           =   1200
            End
            Begin VB.CommandButton cmdProcesoImp 
               Caption         =   "Ac&tualiza oficio"
               Enabled         =   0   'False
               Height          =   780
               Index           =   3
               Left            =   45
               Picture         =   "frmAnálisis.frx":003C
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   2565
               Width           =   1200
            End
            Begin VB.CommandButton cmdProcesoImp 
               Caption         =   "&Nuevo oficio"
               Enabled         =   0   'False
               Height          =   375
               Index           =   0
               Left            =   45
               TabIndex        =   41
               Top             =   270
               Width           =   1200
            End
            Begin VB.CommandButton cmdProcesoImp 
               Caption         =   "&Agrega oficio"
               Enabled         =   0   'False
               Height          =   825
               Index           =   1
               Left            =   45
               Picture         =   "frmAnálisis.frx":0DC6
               Style           =   1  'Graphical
               TabIndex        =   42
               Top             =   855
               Width           =   1200
            End
            Begin VB.CommandButton cmdProcesoImp 
               Caption         =   "&Edita oficio"
               Enabled         =   0   'False
               Height          =   375
               Index           =   2
               Left            =   45
               TabIndex        =   43
               Top             =   2025
               Width           =   1200
            End
            Begin VB.CommandButton cmdProcesoImp 
               Caption         =   "Des&hace oficio"
               Enabled         =   0   'False
               Height          =   780
               Index           =   4
               Left            =   45
               Picture         =   "frmAnálisis.frx":1B50
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   3510
               Width           =   1200
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   3660
            Left            =   360
            TabIndex        =   63
            Top             =   1350
            Width           =   12255
            Begin VB.TextBox txtCampo 
               BackColor       =   &H8000000F&
               DataField       =   "Nombre"
               Height          =   285
               Index           =   12
               Left            =   1395
               MaxLength       =   20
               TabIndex        =   86
               Tag             =   "f"
               ToolTipText     =   "Fecha del acuerdo de improcedencia"
               Top             =   1845
               Width           =   3030
            End
            Begin VB.CommandButton cmdAgregaCausa 
               Caption         =   "Selecciona Turnadas"
               Enabled         =   0   'False
               Height          =   330
               Index           =   5
               Left            =   10170
               TabIndex        =   39
               Top             =   2250
               Width           =   1860
            End
            Begin VB.CommandButton cmdAgregaCausa 
               Caption         =   "Quita causa"
               Height          =   465
               Index           =   1
               Left            =   10665
               TabIndex        =   38
               Top             =   1170
               Width           =   1185
            End
            Begin VB.CommandButton cmdAgregaCausa 
               Caption         =   "Agrega causa"
               Height          =   465
               Index           =   0
               Left            =   10665
               TabIndex        =   37
               Top             =   585
               Width           =   1185
            End
            Begin VB.ListBox ListCausaLeyImp 
               Height          =   645
               Left            =   270
               TabIndex        =   40
               Top             =   2655
               Width           =   11760
            End
            Begin VB.ComboBox cmbCampo 
               BackColor       =   &H8000000F&
               Height          =   315
               Index           =   3
               ItemData        =   "frmAnálisis.frx":2A56
               Left            =   1395
               List            =   "frmAnálisis.frx":2A58
               TabIndex        =   36
               ToolTipText     =   "motivo"
               Top             =   1260
               Width           =   8500
            End
            Begin VB.ComboBox cmbCampo 
               BackColor       =   &H8000000F&
               Height          =   315
               Index           =   1
               ItemData        =   "frmAnálisis.frx":2A5A
               Left            =   1395
               List            =   "frmAnálisis.frx":2A5C
               TabIndex        =   34
               ToolTipText     =   "ley"
               Top             =   360
               Width           =   8500
            End
            Begin VB.ComboBox cmbCampo 
               BackColor       =   &H8000000F&
               Height          =   315
               Index           =   2
               ItemData        =   "frmAnálisis.frx":2A5E
               Left            =   1395
               List            =   "frmAnálisis.frx":2A60
               TabIndex        =   35
               ToolTipText     =   "Causa"
               Top             =   810
               Width           =   8500
            End
            Begin VB.Label etiTexto 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Fecha de Infracción:"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   12
               Left            =   225
               TabIndex        =   87
               Top             =   1800
               Width           =   1110
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Causa (Ley):"
               BeginProperty Font 
                  Name            =   "Constantia"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   270
               TabIndex        =   67
               Top             =   2340
               Width           =   1140
            End
            Begin VB.Label etiCombo 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Motivo improcedencia:"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   3
               Left            =   270
               TabIndex        =   66
               Top             =   1215
               Width           =   1110
            End
            Begin VB.Label etiCombo 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Ley:"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   270
               TabIndex        =   65
               Top             =   450
               Width           =   300
            End
            Begin VB.Label etiCombo 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Causa:"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   2
               Left            =   270
               TabIndex        =   64
               Top             =   900
               Width           =   495
            End
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H8000000F&
            DataField       =   "Nombre"
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   2610
            MaxLength       =   50
            TabIndex        =   31
            Tag             =   "c"
            ToolTipText     =   "No. de Acuerdo de improcedencia"
            Top             =   360
            Width           =   3345
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H8000000F&
            DataField       =   "Nombre"
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   8685
            MaxLength       =   20
            TabIndex        =   32
            Tag             =   "f"
            ToolTipText     =   "Fecha del acuerdo de improcedencia"
            Top             =   360
            Width           =   2715
         End
         Begin VB.Label etiTexto 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Observaciones:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   540
            TabIndex        =   75
            Top             =   720
            Width           =   1110
         End
         Begin VB.Image Image2 
            Height          =   570
            Left            =   12015
            Picture         =   "frmAnálisis.frx":2A62
            Stretch         =   -1  'True
            Top             =   720
            Width           =   600
         End
         Begin VB.Label etiTexto 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "No.Acuerdo Adm. de Improc.:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   315
            TabIndex        =   62
            Top             =   360
            Width           =   2100
         End
         Begin VB.Label etiTexto 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha de Acuerdo de Improc.:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   6345
            TabIndex        =   61
            Top             =   405
            Width           =   2160
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   6090
         Left            =   0
         TabIndex        =   49
         Top             =   405
         Width           =   15540
         Begin VB.TextBox txtCampo 
            BackColor       =   &H8000000F&
            DataField       =   "Nombre"
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   9180
            MaxLength       =   4
            TabIndex        =   10
            Tag             =   "n"
            ToolTipText     =   "Fecha de Emplazamiento"
            Top             =   180
            Width           =   570
         End
         Begin VB.CheckBox chkPruebas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Captura de fecha de Alegatos y Pruebas"
            Height          =   240
            Left            =   10080
            TabIndex        =   11
            Top             =   135
            Width           =   3570
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00FFFFFF&
            Height          =   870
            Left            =   9900
            TabIndex        =   82
            Top             =   315
            Visible         =   0   'False
            Width           =   3660
            Begin VB.TextBox txtCampo 
               BackColor       =   &H8000000F&
               DataField       =   "Nombre"
               Enabled         =   0   'False
               Height          =   285
               Index           =   10
               Left            =   1800
               MaxLength       =   20
               TabIndex        =   16
               Tag             =   "f"
               ToolTipText     =   "Fecha de Emplazamiento"
               Top             =   540
               Width           =   1560
            End
            Begin VB.TextBox txtCampo 
               BackColor       =   &H8000000F&
               DataField       =   "Nombre"
               Enabled         =   0   'False
               Height          =   285
               Index           =   9
               Left            =   1800
               MaxLength       =   20
               TabIndex        =   15
               Tag             =   "f"
               ToolTipText     =   "Fecha de Emplazamiento"
               Top             =   180
               Width           =   1560
            End
            Begin VB.Label etiTexto 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Fecha pruebas:"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   11
               Left            =   270
               TabIndex        =   85
               Top             =   585
               Width           =   1110
            End
            Begin VB.Label etiTexto 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Fecha alegatos:"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   10
               Left            =   225
               TabIndex        =   84
               Top             =   225
               Width           =   1140
            End
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H8000000F&
            DataField       =   "Nombre"
            Enabled         =   0   'False
            Height          =   465
            Index           =   7
            Left            =   1320
            MaxLength       =   250
            TabIndex        =   12
            Tag             =   "c"
            ToolTipText     =   "Fecha de Emplazamiento"
            Top             =   585
            Width           =   8490
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   3435
            Left            =   135
            TabIndex        =   70
            Top             =   1125
            Width           =   13830
            Begin VB.TextBox txtCampo 
               BackColor       =   &H8000000F&
               DataField       =   "Nombre"
               Enabled         =   0   'False
               Height          =   285
               Index           =   14
               Left            =   6480
               MaxLength       =   20
               TabIndex        =   91
               Tag             =   "n"
               ToolTipText     =   "Fecha de Emplazamiento"
               Top             =   648
               Width           =   2064
            End
            Begin VB.TextBox txtCampo 
               BackColor       =   &H8000000F&
               DataField       =   "Nombre"
               Enabled         =   0   'False
               Height          =   285
               Index           =   13
               Left            =   3204
               MaxLength       =   20
               TabIndex        =   89
               Tag             =   "f"
               ToolTipText     =   "Fecha de Emplazamiento"
               Top             =   648
               Width           =   1560
            End
            Begin VB.CommandButton cmdAgregaCausa 
               Caption         =   "Sel.Turnadas"
               Enabled         =   0   'False
               Height          =   600
               Index           =   4
               Left            =   12195
               TabIndex        =   19
               Top             =   270
               Width           =   1275
            End
            Begin VB.CommandButton cmdAgregaCausa 
               Caption         =   "Quita causa"
               Height          =   375
               Index           =   3
               Left            =   10350
               TabIndex        =   18
               Top             =   504
               Width           =   1320
            End
            Begin VB.CommandButton cmdAgregaCausa 
               Caption         =   "Agrega causa"
               Height          =   375
               Index           =   2
               Left            =   10350
               TabIndex        =   17
               Top             =   135
               Width           =   1320
            End
            Begin VB.ComboBox cmbCampo 
               BackColor       =   &H8000000F&
               Height          =   288
               Index           =   5
               ItemData        =   "frmAnálisis.frx":B354
               Left            =   1908
               List            =   "frmAnálisis.frx":B356
               TabIndex        =   14
               ToolTipText     =   "Causa"
               Top             =   288
               Width           =   7812
            End
            Begin VB.ComboBox cmbCampo 
               BackColor       =   &H8000000F&
               Height          =   288
               Index           =   4
               ItemData        =   "frmAnálisis.frx":B358
               Left            =   288
               List            =   "frmAnálisis.frx":B35A
               TabIndex        =   13
               ToolTipText     =   "Ley"
               Top             =   288
               Width           =   1548
            End
            Begin VB.ListBox ListCausaLey 
               Height          =   645
               Left            =   48
               TabIndex        =   20
               Top             =   1128
               Width           =   13656
            End
            Begin VB.ListBox ListIncAcep 
               Height          =   645
               ItemData        =   "frmAnálisis.frx":B35C
               Left            =   90
               List            =   "frmAnálisis.frx":B35E
               TabIndex        =   21
               Top             =   2295
               Width           =   6795
            End
            Begin VB.ListBox ListIncNoAcep 
               Height          =   645
               ItemData        =   "frmAnálisis.frx":B360
               Left            =   6930
               List            =   "frmAnálisis.frx":B362
               TabIndex        =   77
               Top             =   2295
               Width           =   6795
            End
            Begin VB.CommandButton cmdAgregaIncum 
               Caption         =   "Agregar Incump"
               Height          =   240
               Left            =   5175
               TabIndex        =   79
               Top             =   2025
               Width           =   1500
            End
            Begin VB.Label etiCombo 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Monto Multa (Aprox.):"
               ForeColor       =   &H00000000&
               Height          =   192
               Index           =   7
               Left            =   4932
               TabIndex        =   90
               Top             =   720
               Width           =   1488
            End
            Begin VB.Label etiCombo 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "F. Infracción:"
               ForeColor       =   &H00000000&
               Height          =   192
               Index           =   6
               Left            =   2232
               TabIndex        =   88
               Top             =   684
               Width           =   888
            End
            Begin VB.Label Label5 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Incumplimientos No Aceptados (Doble Clic para Agregar)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   6930
               TabIndex        =   78
               Top             =   2115
               Width           =   5415
            End
            Begin VB.Label Label4 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Incumplimientos Aceptados (Doble Clic para Quitar)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   90
               TabIndex        =   76
               Top             =   2115
               Width           =   4425
            End
            Begin VB.Label etiCombo 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Causa:"
               ForeColor       =   &H00000000&
               Height          =   192
               Index           =   5
               Left            =   1908
               TabIndex        =   73
               Top             =   108
               Width           =   492
            End
            Begin VB.Label etiCombo 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Ley:"
               ForeColor       =   &H00000000&
               Height          =   192
               Index           =   4
               Left            =   276
               TabIndex        =   72
               Top             =   108
               Width           =   336
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Causa (Ley):"
               BeginProperty Font 
                  Name            =   "Constantia"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   228
               Left            =   96
               TabIndex        =   71
               Top             =   864
               Width           =   1140
            End
         End
         Begin VB.CommandButton cmdAgregarCausa1 
            BackColor       =   &H000080FF&
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
            Height          =   375
            Left            =   1620
            Picture         =   "frmAnálisis.frx":B364
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   1035
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H8000000F&
            DataField       =   "Nombre"
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   6300
            MaxLength       =   20
            TabIndex        =   9
            Tag             =   "f"
            ToolTipText     =   "Fecha de Emplazamiento"
            Top             =   180
            Width           =   1560
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H8000000F&
            DataField       =   "Nombre"
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   1755
            MaxLength       =   50
            TabIndex        =   8
            Tag             =   "c"
            ToolTipText     =   "No. Oficio de emplazamiento"
            Top             =   225
            Width           =   3090
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            Height          =   5535
            Left            =   14085
            TabIndex        =   50
            Top             =   270
            Width           =   1290
            Begin VB.CommandButton cmdProceso 
               Caption         =   "Des&hace oficio"
               Enabled         =   0   'False
               Height          =   780
               Index           =   4
               Left            =   90
               Picture         =   "frmAnálisis.frx":D396
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   3645
               Width           =   1200
            End
            Begin VB.CommandButton cmdProceso 
               Caption         =   "&Edita oficio"
               Enabled         =   0   'False
               Height          =   375
               Index           =   2
               Left            =   45
               TabIndex        =   26
               Top             =   2160
               Width           =   1200
            End
            Begin VB.CommandButton cmdProceso 
               Caption         =   "&Agrega oficio"
               Enabled         =   0   'False
               Height          =   825
               Index           =   1
               Left            =   45
               Picture         =   "frmAnálisis.frx":E29C
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   945
               Width           =   1200
            End
            Begin VB.CommandButton cmdProceso 
               Caption         =   "&Nuevo oficio"
               Enabled         =   0   'False
               Height          =   375
               Index           =   0
               Left            =   45
               TabIndex        =   24
               Top             =   405
               Width           =   1200
            End
            Begin VB.CommandButton cmdProceso 
               Caption         =   "Ac&tualiza oficio"
               Enabled         =   0   'False
               Height          =   780
               Index           =   3
               Left            =   45
               Picture         =   "frmAnálisis.frx":F026
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   2700
               Width           =   1200
            End
            Begin VB.CommandButton cmdProceso 
               Caption         =   "&Borra oficio"
               Enabled         =   0   'False
               Height          =   375
               Index           =   5
               Left            =   90
               TabIndex        =   29
               Top             =   4815
               Width           =   1200
            End
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1470
            Left            =   180
            TabIndex        =   30
            Top             =   4545
            Width           =   13740
            _ExtentX        =   24236
            _ExtentY        =   2593
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Oficio"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Emplazamiento"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Responsable"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Causa(s) (Ley(es))"
               Object.Width           =   10936
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Observaciones"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Días Otorgados"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Fecha Alegatos"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Facha Pruebas"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Días Otorgados"
            Height          =   285
            Left            =   7965
            TabIndex        =   83
            Top             =   225
            Width           =   1185
         End
         Begin VB.Label etiTexto 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Observaciones:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   135
            TabIndex        =   74
            Top             =   585
            Width           =   1110
         End
         Begin VB.Image Image3 
            Height          =   390
            Left            =   13590
            Picture         =   "frmAnálisis.frx":FDB0
            Stretch         =   -1  'True
            Top             =   270
            Width           =   420
         End
         Begin VB.Label Etilist 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Causa (Ley):"
            BeginProperty Font 
               Name            =   "Constantia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   225
            Left            =   315
            TabIndex        =   59
            Top             =   990
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label etiTexto 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha oficio emp.:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   4950
            TabIndex        =   58
            Top             =   225
            Width           =   1305
         End
         Begin VB.Label etiTexto 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Oficio emplazamiento:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   135
            TabIndex        =   57
            Top             =   270
            Width           =   1545
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0D75F&
      Height          =   1275
      Left            =   1710
      TabIndex        =   23
      Top             =   0
      Width           =   14010
      Begin MSComctlLib.ImageList Imagenes 
         Left            =   1935
         Top             =   180
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
               Picture         =   "frmAnálisis.frx":186A2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAnálisis.frx":205B4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdActualpen 
         Caption         =   "Nuevo análisis"
         Height          =   420
         Left            =   11925
         TabIndex        =   2
         Top             =   225
         Width           =   1680
      End
      Begin VB.ComboBox cmbPendientes 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   4185
         TabIndex        =   1
         ToolTipText     =   "Expedientes pendientes de análisis"
         Top             =   900
         Width           =   7395
      End
      Begin VB.TextBox txtNuevoExp 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   90
         MaxLength       =   80
         TabIndex        =   0
         Tag             =   "c"
         ToolTipText     =   "Número de expediente a realizar análisis"
         Top             =   900
         Width           =   3975
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
         Left            =   11988
         Picture         =   "frmAnálisis.frx":28EB6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   756
         Width           =   1500
      End
      Begin MSForms.CommandButton cmdIrReg 
         Height          =   555
         Left            =   8190
         TabIndex        =   81
         Top             =   225
         Width           =   1680
         BackColor       =   11854537
         Caption         =   "<< Regresa Registro"
         Size            =   "2963;979"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdIrSeg 
         Height          =   555
         Left            =   9990
         TabIndex        =   80
         Top             =   225
         Width           =   1590
         BackColor       =   12179186
         Caption         =   "Ir a Seguimiento >>"
         Size            =   "2805;979"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0D75F&
         Caption         =   "Expedientes pendientes de análisis:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   4185
         TabIndex        =   68
         Top             =   630
         Width           =   2520
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0D75F&
         Caption         =   "No. Expediente:"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   48
         Top             =   630
         Width           =   1635
      End
      Begin VB.Label Eti 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0D75F&
         Caption         =   "Módulo de Análisis"
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
         Left            =   1980
         TabIndex        =   47
         Top             =   180
         Width           =   7740
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   90
      Picture         =   "frmAnálisis.frx":29A25
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1530
   End
End
Attribute VB_Name = "Análisis"
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
Dim miUni As Integer 'Contiene el id de la unidad de origen
Dim miClase As Integer 'Contiene el id de la unidad de origen
Dim mbCambioImp As Boolean 'indica si hay cambio en datos de improcedencia
Dim mbLimpiaExp As Boolean 'indicador para limpiar el campo o lista de exp Pendientes
Dim sFolioImpEsperado As String 'Contiene el formato del folio improcedente esperado
Dim sFolioProEsperado As String 'Contiene el formato del folio procedente esperado
Dim miModulo As Long 'id de módulos en caso que exista tabla s_sanciones
Dim msModXCau As String 'contiene los ids Causa Ley de módulos en caso que exista tabla R_san_ins_rcl
Dim msModXCauXAso As String 'contiene los ids Causa Ley de módulos Por asociar
Dim miModPro As Integer 'Indicador del Aceptación y Análisis procedente de todas las causas capturadas en Módulos
Dim msModInc As String 'Contiene de ids de incumplimientos asociados a acada causa
Dim msModIncD As String 'Contiene de las descripciónes de los incumplimientos para todas las causas
Dim micau As Integer 'Causa de la sanción selecionada
Dim msExpediente As String 'Contiene el núm Expediente
Dim miPro As Integer 'Contiene el valor del proceso cuando está asociado a DGEV
Dim csUniDGEV As String '"|1011|1012|1013|1018|" 'id de unidad correspondientes a la DGEV
Dim bTurMod As Boolean 'Indica si fue turnado desde módulos

Private Sub chkPruebas_Click()
If chkPruebas.Value > 0 Then
    If Not Frame9.Visible Then Frame9.Visible = True
Else
    If Frame9.Visible Then Frame9.Visible = False
End If
End Sub

'Dim miUnidad As Integer 'Contine el id de la unidad de origen para conocer el origen de la solicitud cuando provenga de MSS


Private Sub cmbCampo_Click(Index As Integer)
Dim i As Long, adors As New ADODB.Recordset
Dim s As String
'Dim iCau As Integer, iley As Integer
Dim iIns As Long 'id de la IF
If Index = 0 And cmbCampo(Index).ListIndex >= 0 Then 'Refresca datos de Análisis de esta institución
    iIns = cmbCampo(Index).ItemData(cmbCampo(Index).ListIndex)
    If adors.State Then adors.Close
    adors.Open "select idins from registroxif where id=" & iIns, gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        iIns = adors(0)
    End If
    
    If adors.State Then adors.Close
    adors.Open "select count(*) from registromodulos where idreg=f_expediente_idreg('" & txtNuevoExp.Text & "')", gConSql, adOpenStatic, adLockReadOnly
    
    If adors(0) > 0 Then 'viene de módulos
        'Verifica si es procedente o no desde módulos
        If adors.State Then adors.Close
        adors.Open "select idmod,procede from registromodulos where idreg=f_expediente_idreg('" & txtNuevoExp.Text & "')", gConSql, adOpenStatic, adLockReadOnly
        miModulo = adors(0)
        If adors(1) = 1 Then 'Es procedente
            miModPro = 1
            'SSTab1.TabEnabled(1) = False
            'SSTab1.Tab = 0
        ElseIf adors(1) = 2 Then 'Es improcedente
            miModPro = 2
            'SSTab1.TabEnabled(0) = False
            'SSTab1.Tab = 1
        Else
            MsgBox "El asunto está ligado a módulos de Socilitud de Sanción sin embargo no está definido la Aceptación o Rechazo. Favor de Informar a Miguel esta situación para dar seguimiento", vbOKOnly, ""
            Exit Sub
        End If
    End If
    
    
    
    
    
    
    'PROCEDENTES
    mlAsuxIF = cmbCampo(Index).ItemData(cmbCampo(Index).ListIndex)
    
    If adors.State Then adors.Close
    adors.Open "select count(*) from análisis where idregxif=" & mlAsuxIF & " and idusi=" & giUsuario, gConSql, adOpenStatic, adLockReadOnly
    
    ''Verifica si el usuario debería darle el análisis
    'If adors(0) <= 0 Then ' El usuario no realizó el análisis por lo que no puede modificarlo o terminar de realizar el análisis
    '    If adors.State Then adors.Close
    '    adors.Open "select instr(f_registro_idres_turnados(" & mlAsunto & "),'|'||" & giResponsable & "||'|') from dual", gConSql, adOpenStatic, adLockReadOnly
    '    If adors(0) <= 0 Then ' El usuario no realizó el análisis y tampoco se le turno el asunto por tanto no realiza ningún análisis
    '        MsgBox "El Asunto fue turnado a otro responsable. Únicamente el responsable a quien fue turnado puede realizar el análisis", vbOKOnly + vbInformation, "Validación"
    '        Exit Sub
    '    End If
    'End If
    
    
    'Actualiza botones
    BloqueaControles 2
    cmdProceso(0).Enabled = True 'nuevo
    cmdProceso(1).Enabled = False 'agregar
    cmdProceso(2).Enabled = False 'editar
    cmdProceso(3).Enabled = False 'actualizar
    cmdProceso(4).Enabled = False 'deshacer
    cmdProceso(5).Enabled = False 'borrar
    
    ListView1.ListItems.Clear
    If adors.State Then adors.Close
    adors.Open "select a.oficio,a.fecha,us.descripción,f_causas(a.id),ao.observaciones,a.otorgados,a.f_alegatos,a.f_pruebas,a.id from análisis a, usuariossistema us, análisisobs ao where a.idregxif=" & mlAsuxIF & "  and a.procedente<>0 and a.idusi=us.id(+) and a.id=ao.idana(+)", gConSql, adOpenStatic, adLockReadOnly
    i = 1
    cmbCampo(4).ListIndex = -1
    cmbCampo(5).ListIndex = -1
    If adors.EOF Then
        txtCampo(3).Text = ""
        txtCampo(4).Text = ""
        ListCausaLey.Clear
    End If
    Do While Not adors.EOF
        ListView1.ListItems.Add i, , adors(0)
        ListView1.ListItems(i).SubItems(1) = IIf(IsNull(adors(1)), "", adors(1))
        ListView1.ListItems(i).SubItems(2) = IIf(IsNull(adors(2)), "", adors(2))
        ListView1.ListItems(i).SubItems(3) = IIf(IsNull(adors(3)), "", adors(3))
        ListView1.ListItems(i).SubItems(4) = IIf(IsNull(adors(4)), "", adors(4))
        ListView1.ListItems(i).SubItems(5) = IIf(IsNull(adors(5)), "", adors(5))
        ListView1.ListItems(i).SubItems(6) = IIf(IsNull(adors(6)), "", Format(adors(6), gsFormatoFecha))
        ListView1.ListItems(i).SubItems(7) = IIf(IsNull(adors(7)), "", Format(adors(7), gsFormatoFecha))
        ListView1.ListItems(i).Tag = adors(8) 'Guarda el id
        adors.MoveNext
        i = i + 1
    Loop
    If ListView1.ListItems.Count > 0 Then
        ListView1.ListItems(1).Selected = True
        Call ListView1_ItemClick(ListView1.ListItems(1))
    End If
    If adors.RecordCount > 0 Then
        Image3.Picture = Imagenes.ListImages(1).Picture
    Else
        Image3.Picture = Imagenes.ListImages(2).Picture
    End If
    If Not cmdProceso(0).Enabled And Not cmdProceso(1).Enabled Then
        cmdProceso(0).Enabled = True
    End If
    Frame7.Enabled = False
    'IMPROCEDENTES
    'Actualiza botones
    cmdProcesoImp(0).Enabled = False 'nuevo
    cmdProcesoImp(1).Enabled = False 'agregar
    cmdProcesoImp(2).Enabled = False 'editar
    cmdProcesoImp(3).Enabled = False 'actualizar
    cmdProcesoImp(4).Enabled = False 'deshacer
    cmdProcesoImp(5).Enabled = False 'borrar
    ListCausaLeyImp.Clear
    If adors.State Then adors.Close
    adors.Open "select a.oficio,a.fecha,ao.observaciones,a.id from análisis a, usuariossistema us, análisisobs ao where a.idregxif=" & mlAsuxIF & " and a.procedente=0 and a.idusi=us.id(+) and a.id=ao.idana(+)", gConSql, adOpenStatic, adLockReadOnly
    msCausasIMP = ""
    'txtCampo(5).Enabled = False
    If Not adors.EOF Then
        txtCampo(5).Text = adors(0)
        txtCampo(6).Text = Format(adors(1), gsFormatoFecha)
        txtCampo(8).Text = IIf(IsNull(adors(2)), "", adors(2))
        mlAnálisisImp = adors(3)
        mlAnálisis = adors(3)
        If adors.State Then adors.Close
        adors.Open "select r.idcau,c.descripción ||' ('||l.descripción||') ('||m.descripción||') ('||to_char(r.infraccion,'dd/mm/yyyy')||')', r.idley, r.idmotimp, to_char(r.infraccion,'dd/mm/yyyy') from análisiscausasimp r, leyes l, causas c, motivosimp m where r.idana=" & mlAnálisisImp & " and r.idcau=c.id(+) and r.idley=l.id(+) and r.idmotimp=m.id(+)", gConSql, adOpenStatic, adLockReadOnly
        cmbCampo(1).ListIndex = -1
        cmbCampo(2).ListIndex = -1
        cmbCampo(3).ListIndex = -1
        Do While Not adors.EOF
            ListCausaLeyImp.AddItem adors(1)
            'ley * un millón +  causa*1000 + motivo
            ListCausaLeyImp.ItemData(ListCausaLeyImp.NewIndex) = adors(2) * 1000000 + adors(0) * 1000 + adors(3)
            msCausasIMP = msCausasIMP & adors(2) & "," & adors(0) & "|" & adors(3) & "|" & adors(4) & "|"
            adors.MoveNext
        Loop
        i = 200
        cmdProcesoImp(2).Enabled = True
        cmdProcesoImp(5).Enabled = True
    Else
        txtCampo(5).Text = ""
        txtCampo(6).Text = ""
        txtCampo(8).Text = ""
        txtCampo(12).Text = ""
        cmdProcesoImp(0).Enabled = True
        mlAnálisisImp = 0
        i = 0
    End If
    
    mbCambioImp = False
    If i = 200 Then
        Image2.Picture = Imagenes.ListImages(1).Picture
    Else
        Image2.Picture = Imagenes.ListImages(2).Picture
    End If
    'filtra los combos de Leyes y Causas según la unidad de origen  y Clase
    If adors.State Then adors.Close
    adors.Open "select f_registroxif_iduni(" & mlAsuxIF & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        miUnidad = IIf(IsNull(adors(0)), 0, adors(0))
    Else
        miUnidad = 0
    End If
    If adors.State Then adors.Close
    adors.Open "select f_registroxif_idcla(" & mlAsuxIF & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        miClase = IIf(IsNull(adors(0)), 0, adors(0))
    Else
        miClase = 0
    End If
    LlenaCombo cmbCampo(1), "select l.id,l.descripción from relaciónleyunidad rlu, leyes l where rlu.iduni=" & miUnidad & " and rlu.idley=l.id and l.fechabaja is null and f_relacionuniley_concau( " & miUnidad & ", rlu.idley, " & miClase & ")>0 order by 2", "", True
    LlenaCombo cmbCampo(4), "select l.id,l.descripción from relaciónleyunidad rlu, leyes l where rlu.iduni=" & miUnidad & " and rlu.idley=l.id and l.fechabaja is null and f_relacionuniley_concau( " & miUnidad & ", rlu.idley, " & miClase & ")>0 order by 2", "", True
    'LlenaCombo cmbCampo(2), "select c.id,c.descripción from relaciónunidadcausa ruc, causas c where ruc.iduni=" & miUnidad & " and ruc.idcau=c.id and c.fechabaja is null order by 2", "", True
    'LlenaCombo cmbCampo(5), "select c.id,c.descripción from relaciónunidadcausa ruc, causas c where ruc.iduni=" & miUnidad & " and ruc.idcau=c.id and c.fechabaja is null order by 2", "", True
    LlenaCombo cmbCampo(3), "select m.id,m.descripción from motivosimp m where m.fechabaja is null order by 2", "", True
    
ElseIf Index = 1 Then  'filtra causas
    If adors.State Then adors.Close
    adors.Open "select f_registroxif_idcla(" & mlAsuxIF & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        i = IIf(IsNull(adors(0)), 0, adors(0))
    Else
        i = 0
    End If
    If cmbCampo(Index).ListIndex >= 0 Then
        'LlenaCombo cmbCampo(2), "select c.id,c.descripción from relaciónunidadcausa ruc, causas c where ruc.iduni=" & miUnidad & " and ruc.idley=" & cmbCampo(Index).ItemData(cmbCampo(Index).ListIndex) & " and ruc.idcau=c.id and c.fechabaja is null and f_relacionleycau_clase(ruc.idcau," & i & ")>0 order by 2", "", True
        'LlenaCombo cmbCampo(2), "select c.id,c.descripción from relaciónunidadcausa ruc, causas c where ruc.iduni=" & miUnidad & " and ruc.idcau=c.id and c.fechabaja is null and f_relacionleycau_clase(ruc.idcau," & i & ")>0 order by 2", "", True
        LlenaCombo cmbCampo(2), "select c.id,c.descripción from relaciónleycausa rlc, causas c where rlc.idley=" & cmbCampo(Index).ItemData(cmbCampo(Index).ListIndex) & " and rlc.idcau=c.id order by 2", "", True
    Else
        cmbCampo(2).Clear
    End If
ElseIf Index = 3 And cmbCampo(Index).ListIndex >= 0 Then 'Cambia motivo de improcedencia cuando se tiene causa ley desde Módulos

Dim iLey As Long, iCau As Long, iMot As Long
    If Len(msModXCauXAso) > 0 And cmbCampo(1).ListIndex >= 0 And cmbCampo(2).ListIndex >= 0 Then 'Causas por asociar dede Módulos
        If ListCausaLeyImp.ListIndex >= 0 Then 'Está seleccionada una causa ley
            iLey = Round(ListCausaLeyImp.ItemData(ListCausaLeyImp.ListIndex) / 1000000, 0)
            iCau = Round(ListCausaLeyImp.ItemData(ListCausaLeyImp.ListIndex) / 1000, 0) Mod 1000
            iMot = ListCausaLeyImp.ItemData(ListCausaLeyImp.ListIndex) Mod 1000
            If iLey = cmbCampo(1).ItemData(cmbCampo(1).ListIndex) And iCau = cmbCampo(2).ItemData(cmbCampo(2).ListIndex) And iMot <> cmbCampo(Index).ItemData(cmbCampo(Index).ListIndex) Then ' pregunta si se desea actualizar el motivo de la causa
                If MsgBox("Desea cambia el Motivo de Causa Ley " & ListCausaLeyImp.Text, vbYesNo + vbQuestion, "Confirmación") = vbNo Then
                    Exit Sub
                End If
                i = cmbCampo(Index).ItemData(cmbCampo(Index).ListIndex)
                msCausasIMP = Replace(msCausasIMP, iLey & "," & iCau & "|" & iMot & "|", iLey & "," & iCau & "|" & i & "|")
                ListCausaLeyImp.List(ListCausaLeyImp.ListIndex) = Mid(ListCausaLeyImp.List(ListCausaLeyImp.ListIndex), 1, InStrRev(ListCausaLeyImp.List(ListCausaLeyImp.ListIndex), "(")) & cmbCampo(3).Text & ")"
                ListCausaLeyImp.ItemData(ListCausaLeyImp.ListIndex) = iLey * 1000000 + iCau * 1000 + i
                ListCausaLeyImp_Click
            End If
        End If
    End If
ElseIf Index = 4 Then  'filtra causas
    If adors.State Then adors.Close
    adors.Open "select f_registroxif_idcla(" & mlAsuxIF & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        i = IIf(IsNull(adors(0)), 0, adors(0))
    Else
        i = 0
    End If
    If cmbCampo(Index).ListIndex >= 0 Then
        If miModulo > 0 Then 'Provienen de módulo, obtiene las causas según la unidad de origen a traves del procedimiento
            LlenaCombo cmbCampo(5), "{call p_unidad_ori_causas(" & miUni & ")}", "", True
            For i = 0 To cmbCampo(5).ListCount - 1
                s = s & cmbCampo(5).ItemData(i) & ","
            Next
            If Len(s) > 1 Then
                s = Mid(s, 1, Len(s) - 1)
            Else
                s = "-1"
            End If
            LlenaCombo cmbCampo(5), "select c.id,c.descripción from relaciónleycausa rlc, causas c where rlc.idley=" & cmbCampo(Index).ItemData(cmbCampo(Index).ListIndex) & " and rlc.idcau=c.id and c.id not in (" & s & " ) order by 2", "", True, True
        Else
            LlenaCombo cmbCampo(5), "select c.id,c.descripción from relaciónleycausa rlc, causas c where rlc.idley=" & cmbCampo(Index).ItemData(cmbCampo(Index).ListIndex) & " and rlc.idcau=c.id order by 2", "", True
        End If
    Else
        cmbCampo(5).Clear
    End If
ElseIf Index = 5 Then  'limpia infracción y Monto cuando no existe la nueva combinación en mscausas
    If cmbCampo(4).ListIndex >= 0 And cmbCampo(5).ListIndex >= 0 Then
        iLey = cmbCampo(4).ItemData(cmbCampo(4).ListIndex)
        iCau = cmbCampo(5).ItemData(cmbCampo(5).ListIndex)
        If InStr("|" & msCausas, "|" & iLey & "," & iCau & ",") = 0 Then
            txtCampo(13).Text = ""
            txtCampo(14).Text = ""
        End If
    End If
End If
If Index > 1 Then
    Image2.Picture = Imagenes.ListImages(2).Picture
    mbCambioImp = True
End If

End Sub

Private Sub cmbPendientes_Click()
If cmbPendientes.ListIndex >= 0 And mbLimpiaExp Then
    mbLimpiaExp = False
    txtNuevoExp.Text = ""
End If
End Sub

Private Sub cmbPendientes_GotFocus()
mbLimpiaExp = True
End Sub

Private Sub cmbPendientes_LostFocus()
mbLimpiaExp = False
End Sub

Private Sub cmdActualpen_Click()
If cmdProceso(1).Enabled Or cmdProceso(3).Enabled Then
    If MsgBox("Desea ignorar los cambios realizados en Procedencia", vbYesNo + vbQuestion + vbDefaultButton2, "") = vbNo Then
        Exit Sub
    End If
End If
If mbCambioImp And txtCampo(4).Enabled Then
    If MsgBox("Desea ignorar los cambios realizados en Improcedencia", vbYesNo + vbQuestion + vbDefaultButton2, "") = vbNo Then
        Exit Sub
    End If
End If
ActualizaPendientes
txtNuevoExp.Enabled = True
cmbPendientes.Enabled = True
cmdContinuar.Enabled = True
For i = 0 To txtCampo.UBound
    txtCampo(i).Text = ""
Next
cmbCampo(0).Clear
cmbCampo(1).Clear
cmbCampo(2).Clear
cmbCampo(3).Clear
cmbCampo(4).Clear
cmbCampo(5).Clear
cmbCampo(0).ListIndex = -1
cmbCampo(1).ListIndex = -1
cmbCampo(2).ListIndex = -1
cmbCampo(3).ListIndex = -1
cmbCampo(4).ListIndex = -1
cmbCampo(5).ListIndex = -1
For i = 0 To cmdProceso.UBound
    cmdProceso(i).Enabled = False
    cmdProcesoImp(i).Enabled = False
Next
ListView1.ListItems.Clear
ListCausaLey.Clear
ListCausaLeyImp.Clear
txtNuevoExp.Text = ""
For i = 0 To cmdProceso.UBound
    cmdProceso(i).Enabled = False
Next
miModulo = 0
msModXCau = ""
msModXCauXAso = ""
miModPro = 0
SSTab1.TabEnabled(0) = True
SSTab1.TabEnabled(1) = True
End Sub



'Agrega causa(s) seleccionadas del Arbol
Private Sub cmdAgregaCausa1_Click()
Dim s As String, adors As New ADODB.Recordset, Y As Integer
gs = "ArbolVarios-->SELECT rlc.idley,rlc.idcau,l.descripción,c.descripción FROM relaciónleycausa rlc, leyes l, causas c where rlc.idley=l.id and rlc.idcau=c.id and l.fechabaja is null and c.fechabaja is null ORDER BY l.descripción,c.descripción"
'gs2 = Str(iInsPrincipal) + "," + Str(iClaPrincipal)
gs2 = Val(msCausas) & "," & Val(msLeyes)
If InStr(msCausas, ",") > 0 Then
    gs1 = Val(msCausas) & Mid(msCausas, InStr(msCausas, ","))
Else
    gs1 = msCausas
End If
If InStr(msLeyes, ",") > 0 Then
    gs3 = Val(msLeyes) & Mid(msLeyes, InStr(msLeyes, ","))
Else
    gs3 = msLeyes
End If
With SelProceso
    .Caption = "Causas por Ley"
    .TreeView1.CheckBoxes = True
    .Show vbModal
    If Len(gs) > 0 Then
        If InStr(gs, ",") > 0 Then
            s = gs
            Y = 1
            If adors.State > 0 Then adors.Close
            adors.Open "select rlc.idley,rlc.idcau,l.descripción as ley,c.descripción as causa from causas c, relaciónleycausa rlc, leyes l where c.id in (" + Mid(gs, 1, Len(gs) - 1) + ") and c.id=rlc.idcau and rlc.idley=l.id", gConSql, adOpenStatic, adLockReadOnly
            ListCausaLey.Clear
            Do While Not adors.EOF
                ListCausaLey.AddItem adors!causa
                ListCausaLey.ItemData(ListCausaLey.NewIndex) = adors(1)
                adors.MoveNext
            Loop
            msCausas = gs
            msLeyes = gs3
        Else
            ListCausaLey.Clear
            msCausas = ""
            msLeyes = ""
        End If
    End If
End With
End Sub

Private Sub cmdAgregaIncum_Click()
Dim iCau As Integer, adors As New ADODB.Recordset, s As String, i As Integer, s1 As String, s2 As String
If InStr(csUniDGEV, "|" & miUni & "|") > 0 Then 'DGEV
    If cmbCampo(5).ListIndex < 0 Then
        MsgBox "Debe seleccionar una causas de sanción para poder elegir los incumplimientos", vbInformation + vbOKOnly, ""
        Exit Sub
    End If
    iCau = cmbCampo(5).ItemData(cmbCampo(5).ListIndex)
    i = 0
    Do While i < ListIncAcep.ListCount
        s = s & ListIncAcep.ItemData(i) & ","
        i = i + 1
    Loop
    i = 0
    Do While i < ListIncNoAcep.ListCount
        s = s & ListIncNoAcep.ItemData(i) & ","
        i = i + 1
    Loop
    If adors.State Then adors.Close
    adors.Open "{call p_idcau_incumps(" & iCau & ",'" & s & "')}", gConSql, adOpenForwardOnly, adLockReadOnly
    If Not adors.EOF Then
        gs = "{call p_idcau_incumps(" & iCau & ",'" & s & "')}"
        MsgBox "Favor de seleccionar los incumplimientos por agregar"
        SelProceso.TreeView1.CheckBoxes = True
        SelProceso.Show vbModal
        If Len(gs) > 0 Then
            If Mid(gs, Len(gs), 1) = "," Then
                gs = Mid(gs, 1, Len(gs) - 1)
            End If
            If adors.State Then adors.Close
            adors.Open "select * from mssdgev.incumplimientos where id in (" & gs & ")"
            Do While Not adors.EOF
                ListIncAcep.AddItem adors(1)
                ListIncAcep.ItemData(ListIncAcep.NewIndex) = adors(0)
                adors.MoveNext
            Loop
            'Actualiza datos de la cuasa y sus incumplimientos
            'If InStr(msModXCau, iLey & "," & iCau & ";") > 0 Then
            '    s = Mid(msModXCau, InStr(msModXCau, iLey & "," & iCau & ";"))
            '    s = Mid(s, 1, InStr(s, "|"))
            '    msModXCau = Replace(msModXCau, s, "")
            'End If
            'msModXCau = msModXCau & iLey & "," & iCau & ";"
            If Mid(msModInc, 1, 1) <> "|" And Len(msModInc) > 0 Then
                msModInc = "|" & msModInc
            End If
            Do While InStr(msModInc, "|" & iCau & ":") > 0
                If Len(s) = 1 Then
                    Exit Do
                End If
                s = Mid(msModInc, InStr(msModInc, "|" & iCau & ":"))
                s = Mid(s, 1, InStr(Mid(s, 2), "|") + 1)
                msModInc = Replace(msModInc, s, "|")
            Loop
            msModInc = msModInc & iCau & ":1_"
            For i = 0 To ListIncAcep.ListCount - 1
                msModInc = msModInc & ListIncAcep.ItemData(i) & "_1," 'se asigna el incumplimiento que tiene la lista
                Call Agrega_incump(ListIncAcep.List(i) & "¬", ListIncAcep.ItemData(i) & ",")
            Next
            msModInc = msModInc & "|"
            ActualizaListasIncump iCau
            
        End If
    Else
        Call MsgBox("No existen más incumplimientos asociados a la causa: " & cmbCampo(5).ItemData(cmbCampo(5).ListIndex), vbInformation + vbOKOnly)
    End If
End If
End Sub

'Busca y en en caso de encontrar obtiene datos de este folio
Private Sub cmdContinuar_Click()
Dim adors As New ADODB.Recordset
Dim l As Long
If Len(Trim(txtNuevoExp.Text)) = 0 And cmbPendientes.ListIndex < 0 Then
    MsgBox "Debe capturar el No. de expediente o seleccionar de la lista el expediente pendiente", vbOKOnly + vbInformation, ""
    Exit Sub
End If
If Len(Trim(txtNuevoExp.Text)) > 0 Then
    If adors.State Then adors.Close
    adors.Open "select f_asuntoxfolio('" & txtNuevoExp.Text & "') from dual", gConSql, adOpenStatic, adLockReadOnly
    If adors(0) > 0 Then
        mlAsunto = adors(0)
        RefrescaDatos
        txtNuevoExp.Enabled = False
        cmbPendientes.Enabled = False
        cmdContinuar.Enabled = False
    Else
        MsgBox "No se encontró asunto alguno con ese No. de Expediente.", vbOKOnly + vbInformation, ""
    End If
ElseIf cmbPendientes.ListIndex >= 0 Then
    If adors.State Then adors.Close
    gi = cmbPendientes.ItemData(cmbPendientes.ListIndex)
    adors.Open "select r.id from registroxif ri, registro r where ri.id=" & gi & " and ri.idreg=r.id", gConSql, adOpenStatic, adLockReadOnly
    If adors(0) > 0 Then
        mlAsunto = adors(0)
        RefrescaDatos
        txtNuevoExp.Enabled = False
        cmbPendientes.Enabled = False
        cmdContinuar.Enabled = False
    Else
        MsgBox "No se encontró asunto alguno con ese No. de Expediente.", vbOKOnly + vbInformation, ""
    End If
    gi = 0
End If
If Not ListView1.Enabled Then ListView1.Enabled = True
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
    'txtExpediente.Text = adors(4)
    txtNuevoExp.Text = IIf(IsNull(adors(4)), "", adors(4))
    msExpediente = Trim(txtNuevoExp.Text)
    LlenaCombo cmbCampo(0), "select id,f_institucionh(idins) from registroxif where idreg=" & mlAsunto, "", True
    If cmbCampo(0).ListCount = 1 Then
        cmbCampo(0).ListIndex = 0
    End If
Else
    For i = 0 To 2
        txtCampo(i).Text = ""
    Next
End If
If gi > 0 Then
    i = BuscaCombo(cmbCampo(0), gi, True, False)
    If i >= 0 Then
        cmbCampo(0).ListIndex = i
    End If
End If
End Sub

Private Sub cmdGuaraImp_Click(Index As Integer)
Dim adors As New ADODB.Recordset
On Error GoTo ErrorGuardaDatosImp:
If Index = 0 Then
    If Len(Trim(txtCampo(5).Text)) = 0 Then
        MsgBox "El No. Oficio en un dato Requerido", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
    If Not IsDate(txtCampo(6).Text) Then
        MsgBox "la fecha del acuerdo de improcedencia es requerido", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
    If Len(msCausasIMP) < 3 Then
        MsgBox "La(s) caus(as) de improcedencia es(son) requerida(s)", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
    
    
    If mlAnálisisImp <= 0 Then 'agrega
        adors.Open "{call p_AnalisisGuardaDatos(0," & mlAsuxIF & ",0,'" & Replace(txtCampo(5), "'", "''") & "','" & Format(CDate(txtCampo(6).Text), "dd/mm/yyyy") & "','" & Replace(msCausasIMP, ",", "|") & "'," & giUsuario & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
    Else 'Actualizar
        adors.Open "{call p_AnalisisGuardaDatos(" & mlAnálisisImp & "," & mlAsuxIF & ",0,'" & Replace(txtCampo(5), "'", "''") & "','" & Format(CDate(txtCampo(6).Text), "dd/mm/yyyy") & "','" & Replace(msCausasIMP, ",", "|") & "'," & giUsuario & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
    End If
    If adors(0) > 0 Then 'Se guardo Correcto
        MsgBox IIf(mlAnálisisImp = 1, "Se ingreso un nuevo", "Se actualizo el") & " registro.", vbOKOnly + vbInformation, ""
        Call cmbCampo_Click(0)
        Exit Sub
    Else
        MsgBox IIf(mlAnálisisImp = 1, "No se ingreso el", "No se actualizo el") & " registro", vbOKOnly + vbCritical, ""
        Exit Sub
    End If
ElseIf Index = 1 Then
    Call cmbCampo_Click(0)
End If
Exit Sub
ErrorGuardaDatosImp:
If gConSql.Errors.Count > 0 Then
    yError = MsgBox("Error: " + gConSql.Errors(0).Description, vbAbortRetryIgnore + vbCritical, "Error no esperado (" + Str(gConSql.Errors(0).Number) + ")")
Else
    yError = MsgBox("Error: " + Err.Description, vbAbortRetryIgnore + vbCritical, "Error no esperado (" + Str(Err.Number) + ")")
End If


If yError = vbRetry Then
    Resume
ElseIf yError = vbIgnore Then
    Resume Next
End If

End Sub

Private Sub cmdAgregaCausa_Click(Index As Integer)
Dim adors As New ADODB.Recordset, sInfraccion As String
Dim iLey As Long, iCau As Long, iMot As Long, s As String
If Index = 0 Then 'agrega improcedentes
    For i = 1 To 3
        If cmbCampo(i).ListIndex < 0 Then
            MsgBox "Debe capturar " & etiCombo(i).Caption, vbOKOnly + vbInformation, ""
            Exit Sub
        End If
    Next
    If Not IsDate(txtCampo(12).Text) Then
        MsgBox "Debe capturar la fecha de Infracción", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
    iLey = cmbCampo(1).ItemData(cmbCampo(1).ListIndex)
    iCau = cmbCampo(2).ItemData(cmbCampo(2).ListIndex)
    iMot = cmbCampo(3).ItemData(cmbCampo(3).ListIndex)
    sInfraccion = Format(CDate(txtCampo(12).Text), gsFormatoFecha)
    If InStr("|" & msCausasIMP, "|" & iLey & "," & iCau & "|") > 0 Then
        MsgBox "La ley/causa ya existe", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
    msCausasIMP = msCausasIMP & iLey & "," & iCau & "|" & iMot & "|" & sInfraccion & "|"
    ListCausaLeyImp.AddItem cmbCampo(2).Text & " (" & cmbCampo(1).Text & ") (" & cmbCampo(3).Text & ") (" & txtCampo(12).Text & ")"
    ListCausaLeyImp.ItemData(ListCausaLeyImp.NewIndex) = iLey * 1000000 + iCau * 1000 + iMot
ElseIf Index = 1 Then 'Quita improcedentes
    If ListCausaLeyImp.ListIndex < 0 Then
        MsgBox "Debe seleccionar una causa", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
    If MsgBox("Está seguro de quitar la causa " & ListCausaLeyImp.List(ListCausaLeyImp.ListIndex), vbYesNo, "Confirmación") = vbNo Then
        Exit Sub
    End If
    Dim s1 As String
    Dim s2 As String
    iLey = Round(ListCausaLeyImp.ItemData(ListCausaLeyImp.ListIndex) / 1000000, 0)
    iCau = Round(ListCausaLeyImp.ItemData(ListCausaLeyImp.ListIndex) / 1000, 0) Mod 1000
    iMot = ListCausaLeyImp.ItemData(ListCausaLeyImp.ListIndex) Mod 1000
    If Len(msModXCauXAso) > 0 Then
        's = msModXCauXAso
        's1 = Mid(s, InStr(s, "|") + 1)
        's2 = Mid(s1, InStr(s1, "|") + 1)
        's1 = Mid(s1, 1, InStr(s1, "|") - 1)
        
        'If InStr("," & s1, "," & iLey & ",") > 0 Then
        '    Do While Val(s1) <> iLey
        '        s1 = Mid(s1, InStr(s1, ",") + 1)
        '        s2 = Mid(s2, InStr(s2, ",") + 1)
        '    Loop
        '    If Val(s1) = iLey And Val(s2) = iCau Then
        '        Call MsgBox("No es posible quitar causa turnada de las áreas, asociada previamente", vbOKOnly + vbInformation, "")
        '        Exit Sub
        '    End If
        'End If
        'If InStr("|" & msModXCauXAso, "|" & iLey & "," & iCau & ";") Then
        '    Call MsgBox("No es posible quitar causa turnada de las áreas, asociada previamente", vbOKOnly + vbInformation, "")
        '    Exit Sub
        'End If
    End If
    sInfraccion = Mid(ListCausaLeyImp.Text, InStrRev(ListCausaLeyImp.Text, "(") + 1)
    sInfraccion = Mid(sInfraccion, 1, Len(sInfraccion) - 1)
    If Not IsDate(sInfraccion) Then
        sInfraccion = ""
    End If
    If InStr(msCausasIMP, iLey & "," & iCau & "|" & iMot & "|" & sInfraccion & "|") > 0 Then
        ListCausaLeyImp.RemoveItem ListCausaLeyImp.ListIndex
        msCausasIMP = Replace(msCausasIMP, iLey & "," & iCau & "|" & iMot & "|" & sInfraccion & "|", "")
    Else
        'MsgBox "Falla el proceso de quitar Causa", vbOKOnly, ""
    End If
ElseIf Index = 2 Then 'agrega en Procedentes
    For i = 4 To 5
        If cmbCampo(i).ListIndex < 0 Then
            MsgBox "Debe capturar " & etiCombo(i).Caption, vbOKOnly + vbInformation, ""
            Exit Sub
        End If
    Next
    iLey = cmbCampo(4).ItemData(cmbCampo(4).ListIndex)
    iCau = cmbCampo(5).ItemData(cmbCampo(5).ListIndex)
    If InStr("|" & msCausas, "|" & iLey & "," & iCau & ",") > 0 Then
        MsgBox "La ley/causa ya existe", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
    If Not IsDate(txtCampo(13).Text) Then 'Valida Fecha de infracción
        MsgBox "Debe capturar la fecha de infracción correspondiente a la causa", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
    If Val(txtCampo(14).Text) <= 0 Then 'Valida Monto de la multa correspondiente a la causa
        MsgBox "Debe capturar el monto aproximado de la multa (Debe ser mayor a cero)", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
    If miModulo > 0 And InStr(csUniDGEV, "|" & miUni & "|") > 0 Then 'Es de la DGEV
        If adors.State Then adors.Close
        adors.Open "Select f_mod_cauxaso2(" & miModulo & ", f_registro_iduni(" & mlAsunto & ")) FROM DUAL", gConSql, adOpenStatic, adLockReadOnly
        If InStr("|" & adors(0), "|" & iLey & "," & iCau & ";") > 0 Then 'Se permite la entrada
            If adors.State Then adors.Close
            adors.Open "select f_obtensubcadena(mssdgev.f_solsan_incump(" & miModulo & "," & iCau & "),1) as Incumplimientos,f_obtensubcadena(mssdgev.f_solsan_incump(" & miModulo & "," & iCau & "),2) as Incump_Ids" & _
                       " from dual", gConSql, adOpenStatic, adLockReadOnly
            msModInc = msModInc & iCau & ":1_" & Replace(adors(1), ",", "_1,") & "|" 'se asignan todas
            Call Agrega_incump(adors(0), adors(1))
                        
        Else ' En caso contrario anula el ingreso
            'Verifica que se tenga un solo incumplimiento para poder ingresarlo.
            iPro = 0
            If InStr(msExpediente, "-I/") Or InStr(msExpediente, "-S/") > 0 Or InStr(msExpediente, "-T/") > 0 Or InStr(msExpediente, "-E/") Then
                miPro = IIf(InStr(msExpediente, "-I/"), 4, IIf(InStr(msExpediente, "-S/"), 3, IIf(InStr(msExpediente, "-T/"), 2, 1)))
            End If
            If adors.State Then adors.Close
            adors.Open "Select f_obtensubcadena(mssdgev.f_iduniprocau_idsinc(" & miUni & "," & miPro & "," & iCau & "),1),f_obtensubcadena(mssdgev.f_iduniprocau_idsinc(" & miUni & "," & miPro & "," & iCau & "),2) FROM DUAL", gConSql, adOpenStatic, adLockReadOnly
            If Len(adors(0)) > 0 Then
                msModInc = msModInc & iCau & ":0_" & Replace(adors(1), ",", "_0,") & "|" 'no se asigna ninguna
                Call Agrega_incump(adors(0), adors(1))
                ListCausaLey.Height = 1100
            Else
                'MsgBox "Solo es permitido agregar causa de la sanción que la DGEV cuando hay incumplimientos asociados en MSS de la DGVE", vbInformation + vbOKOnly, "Validación"
                'Exit Sub
            End If
        End If
        
    End If
    msCausas = msCausas & iLey & "," & iCau & "," & Format(CDate(txtCampo(13).Text), gsFormatoFecha) & "," & Replace(txtCampo(14).Text, ",", "") & "|"
    ListCausaLey.AddItem cmbCampo(5).Text & " (" & cmbCampo(4).Text & ")" & " (" & txtCampo(13).Text & ")" & " (" & txtCampo(14).Text & ")"
    ListCausaLey.ItemData(ListCausaLey.NewIndex) = iLey * 1000 + iCau
ElseIf Index = 3 Then 'Quita en Procedentes
    If ListCausaLey.ListIndex < 0 Then
        MsgBox "Debe seleccionar una causa", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
    i = ListCausaLey.ListIndex
    If MsgBox("Está seguro de quitar la causa " & ListCausaLey.List(ListCausaLey.ListIndex), vbYesNo, "Confirmación") = vbNo Then
        Exit Sub
    End If
    iLey = Int(ListCausaLey.ItemData(ListCausaLey.ListIndex) / 1000)
    iCau = ListCausaLey.ItemData(ListCausaLey.ListIndex) Mod 1000
    If InStr("|" & msCausas, "|" & iLey & "," & iCau & ",") > 0 Then
        ListCausaLey.RemoveItem ListCausaLey.ListIndex
        s = Mid(msCausas, InStr("|" & msCausas, "|" & iLey & "," & iCau & ","))
        s = Mid(s, 1, InStr(s, "|"))
        msCausas = Replace(msCausas, s, "")
        If InStr(msModXCau, iLey & "," & iCau & ";") > 0 Then
            s = Mid(msModXCau, InStr(msModXCau, iLey & "," & iCau & ";"))
            s = Mid(s, 1, InStr(s, "|"))
            msModXCau = Replace(msModXCau, s, "")
        End If
        If InStr(msModInc, "|" & iCau & ":") > 0 Then
            s = Mid(msModInc, InStr(msModInc, "|" & iCau & ":") + 1)
            s = "|" & Mid(s, 1, InStr(s, "|"))
            msModInc = Replace(msModInc, s, "|")
        End If
        ListIncAcep.Clear
        ListIncNoAcep.Clear
        If i - 1 >= 0 And i - 1 <= ListCausaLey.ListCount Then
            ListCausaLey.ListIndex = i - 1
        ElseIf i >= 0 And i < ListCausaLey.ListCount Then
            ListCausaLey.ListIndex = i
        End If
    Else
        MsgBox "Falla el proceso de quitar Causa", vbOKOnly, ""
    End If
ElseIf Index = 5 Then 'Selecciona causa - ley Turnados desde Módulos Improcedente
    If adors.State Then adors.Close
    adors.Open "{call P_mod_cauxaso2(" & miModulo & "," & miUni & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
    Do While Not adors.EOF
        iLey = Round(adors(0) / 1000000, 0)
        iCau = Round(adors(0) / 1000, 0) Mod 1000
        iMot = adors(0) Mod 1000
        msCausasIMP = msCausasIMP & iLey & "," & iCau & "|" & iMot & "||"
        ListCausaLeyImp.AddItem adors(1)
        ListCausaLeyImp.ItemData(ListCausaLeyImp.NewIndex) = adors(0)
        adors.MoveNext
    Loop
    cmdAgregaCausa(5).Enabled = False
ElseIf Index = 4 Then 'Selecciona causa - ley Turnados desde Módulos Procedente
    'gs = "{call P_mod_cauxaso(" & miModulo & ")}"
    'g1 = "1"
    'SelProceso.Show vbModal
    If adors.State Then adors.Close
    adors.Open "{call P_mod_cauxaso2(" & miModulo & "," & miUni & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
    msModInc = "|"
    msModIncD = "|"
    Do While Not adors.EOF
        iLey = Round(adors(0) / 1000000, 0)
        iCau = Round(adors(0) / 1000, 0) Mod 1000
        iMot = adors(0) Mod 1000
        If InStr("|" & msCausas, "|" & iLey & "," & iCau & ",") = 0 Then
            msCausas = msCausas & iLey & "," & iCau & ",,0|" ' idLey, idCua, infraccion, monto |
            ListCausaLey.AddItem adors(1)
            ListCausaLey.ItemData(ListCausaLey.NewIndex) = iLey * 1000 + iCau
        End If
        If InStr(csUniDGEV, "|" & miUni & "|") > 0 Then 'DGEV Obtiene incumplimientos descripciónes ids y aceptación por defecto 1
            msModInc = msModInc & iCau & ":1_" & Replace(adors(3), ",", "_1,") & "|" 'se asignan todas
            Call Agrega_incump(adors(2), adors(3))
        End If
        adors.MoveNext
    Loop
    'Prepara los objetos según se trate o no de la DGEV
    If InStr(csUniDGEV, "|" & miUni & "|") > 0 Then
        ListCausaLey.Height = 1100
        'ActualizaListasIncump
        If ListCausaLey.ListCount > 0 Then
            ListCausaLey.ListIndex = 0
        End If
    Else
        ListCausaLey.Height = 1680
    End If
    msModXCau = msModXCauXAso
    cmdAgregaCausa(4).Enabled = False
End If
End Sub

Sub Agrega_incump(sDesc As String, sIds As String)
Dim s1 As String, s2 As String
s1 = sDesc
s2 = sIds
Do While InStr(s2, ",") > 0
    If InStr("|" & msModIncD, "|" & Val(s2) & ":") = 0 Then
        msModIncD = msModIncD & Val(s2) & ":" & Mid(s1, 1, InStr(s1, "¬") - 1) & "|"
    End If
    s2 = Mid(s2, InStr(s2, ",") + 1)
    s1 = Mid(s1, InStr(s1, "¬") + 1)
Loop
End Sub

Function Busca_incump(iInc As Integer) As String
Dim s As String, i As Integer
i = InStr("|" & msModIncD, "|" & iInc & ":")
If i > 0 Then
    s = Mid(msModIncD, i)
    s = Mid(s, InStr(s, ":") + 1)
    Busca_incump = Mid(s, 1, InStr(s, "|") - 1)
End If
End Function

Sub ActualizaListasIncump(iCau As Integer)
Dim s As String, i As Integer
i = InStr("|" & msModInc, "|" & iCau & ":")
ListIncAcep.Clear
ListIncNoAcep.Clear
If i = 0 Then
    Exit Sub
End If
s = Mid(msModInc, i)
s = Mid(s, InStr(s, ":") + 3) 'Salta ":0_" o ":1_"
s = Mid(s, 1, InStr(s, "|") - 1)
Do While InStr(s, ",") > 0
    i = Val(s)
    If Val(Mid(s, InStr(s, "_") + 1)) > 0 Then 'Aceptada
        ListIncAcep.AddItem Busca_incump(i)
        ListIncAcep.ItemData(ListIncAcep.NewIndex) = i
    Else 'No aceptada
        ListIncNoAcep.AddItem Busca_incump(i)
        ListIncNoAcep.ItemData(ListIncNoAcep.NewIndex) = i
    End If
    s = Mid(s, InStr(s, ",") + 1)
Loop

End Sub

Private Sub cmdIrReg_Click()
Dim frm As Form
If Len(Trim(txtNuevoExp.Text)) > 0 Then
    gs = "<<"
    gs1 = Trim(txtNuevoExp.Text)
    If cmbCampo(0).Visible And cmbCampo(0).ListIndex >= 0 Then
        gi1 = cmbCampo(0).ItemData(cmbCampo(0).ListIndex)
    End If
    Set frm = Registro
    With frm
        .Show
    End With
End If
End Sub

Private Sub cmdIrSeg_Click()
Dim frm As Form
If Len(Trim(txtNuevoExp.Text)) > 0 Then
    gs = ">>"
    gs1 = Trim(txtNuevoExp.Text)
    If cmbCampo(0).Visible And cmbCampo(0).ListIndex >= 0 Then
        gi1 = cmbCampo(0).ItemData(cmbCampo(0).ListIndex)
    End If
    Set frm = Seguimiento
    With frm
        .Show
    End With
End If
End Sub

'Acciones Nuevo, editar , agrega, borra...
Private Sub cmdProceso_Click(Index As Integer)
Dim adors  As New ADODB.Recordset
Dim i As Long, sPruebas As String, sAlegatos As String, iOtorgados As Integer
On Error GoTo ErrorGuardaDatos:
If Index = 0 Then 'Nuevo
    cmdProceso(0).Enabled = False 'nuevo
    cmdProceso(1).Enabled = True 'agregar
    cmdProceso(2).Enabled = False 'editar
    cmdProceso(3).Enabled = False 'actualizar
    cmdProceso(4).Enabled = True 'deshacer
    cmdProceso(5).Enabled = False 'borrar
    ListCausaLey.Clear
    txtCampo(3).Enabled = True
    If adors.State Then adors.Close
    adors.Open "select f_nuevofolio(2,0,0) from dual", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        txtCampo(3).Text = IIf(IsNull(adors(0)), "CONDUSEF/VJ/DGSL/DAS/" & "/" & Format(AhoraServidor, "yyyy"), adors(0))
    Else
        txtCampo(3).Text = "CONDUSEF/VJ/DGSL/DAS/" & "/" & Format(AhoraServidor, "yyyy")
    End If
    sFolioProEsperado = txtCampo(3).Text
    txtCampo(4).Text = ""
    txtCampo(4).Enabled = True
    txtCampo(7).Enabled = True
    txtCampo(9).Enabled = True
    txtCampo(10).Enabled = True
    txtCampo(11).Enabled = True
    txtCampo(13).Enabled = True
    txtCampo(14).Enabled = True
    txtCampo(7).Text = ""
    txtCampo(9).Text = ""
    txtCampo(10).Text = ""
    txtCampo(11).Text = ""
    txtCampo(13).Text = ""
    txtCampo(14).Text = ""
    txtCampo(3).SetFocus
    ListCausaLey.Enabled = True
    Frame7.Enabled = True
    msLeyes = ""
    msCausas = ""
    If ListIncAcep.ListCount > 0 Then ListIncAcep.Clear
    If ListIncNoAcep.ListCount > 0 Then ListIncNoAcep.Clear
    If miModulo > 0 Then
        If adors.State Then adors.Close
        adors.Open "Select f_mod_cauxaso2(" & miModulo & ", f_registro_iduni(" & mlAsunto & ")),f_registro_iduni(" & mlAsunto & ") FROM DUAL", gConSql, adOpenStatic, adLockReadOnly
        miUni = adors(1)
        If Len(adors(0)) > 0 Then 'Existen Causas por asociar
            cmdAgregaCausa(4).Enabled = True
            msModXCauXAso = adors(0)
            msModXCau = ""
            If InStr(csUniDGEV, "|" & miUni & "|") > 0 Then 'El origen corresponde a DGEV
                ListCausaLey.Height = 1100
                If ListCausaLey.ListCount > 0 Then
                    ListCausaLey.Clear
                    msModInc = ""
                    msModIncD = ""
                End If
            Else
                ListCausaLey.Height = 2400
            End If
        End If
    End If
    If ListView1.Enabled Then ListView1.Enabled = False
    chkPruebas.Value = 0
ElseIf Index = 1 Or Index = 3 Then 'agregar / actualizar
    If Len(Trim(txtCampo(3).Text)) = 0 Or Len(Trim(txtCampo(4).Text)) = 0 Or InStr(msCausas, ",") = 0 Then
        MsgBox "El No. y fecha del oficio y causa(s) son datos requeridos. Favor de capturarlos.", vbOKOnl´ + vbInformation, ""
        Exit Sub
    End If
    iOtorgados = Val(txtCampo(11).Text)
    If chkPruebas.Value > 0 Then
        If chkPruebas.Value And (Not IsDate(txtCampo(9).Text) Or Not IsDate(txtCampo(10).Text)) Then
            MsgBox "Las fechas de Elegatos y Pruebas son datos requeridos cuando especifica su captura. Favor de capturarlos.", vbOKOnl´ + vbInformation, ""
            Exit Sub
        End If
        sAlegatos = Format(CDate(txtCampo(9).Text), gsFormatoFecha)
        sPruebas = Format(CDate(txtCampo(10).Text), gsFormatoFecha)
'    If Len(Trim(txtCampo(4).Text)) = 0 Or InStr(msCausas, ",") = 0 Then
'        MsgBox "La fecha del oficio y causa(s) son datos requeridos. Favor de capturarlos.", vbOKOnly + vbInformation, ""
'        Exit Sub
'    End If
    End If
    If adors.State Then adors.Close
    adors.Open "select count(*) from análisis where oficio='" & Replace(txtCampo(3).Text, "'", "''") & "'" & IIf(Index = 1, "", " and id<>" & mlAnálisis), gConSql, adOpenStatic, adLockReadOnly
    If adors(0) > 0 Then
        MsgBox "El número de Oficio '" & Replace(txtCampo(3).Text, "'", "''") & "' ya existe en la base de datos, no puede repetirse. Favor de verificar el No. de Oficio", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
    s = Replace(msCausas, ",", "|")
'    s1 = msLeyes
'    s2 = msCausas
'    Do While InStr(s1, ",")
'        s = s & Val(s1) & "|" & Val(s2) & "|"
'        s1 = Mid(s1, InStr(s1, ",") + 1)
'        s2 = Mid(s2, InStr(s2, ",") + 1)
'    Loop

'    If adors.State Then adors.Close
'    adors.Open "select f_nuevofolio(2,0," & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
'    If Not adors.EOF Then
'        s1 = adors(0)
'        If InStr(s1, "???") Then
'            i = F_PreguntaConsecutivo(2, s1)
'            If i < 0 Then 'Se Ejecutó cancelar
'                Exit Sub
'            End If
'        End If
'    End If
    
    i = ValidaFolio(txtCampo(3).Text, sFolioProEsperado)
    If i <= 0 Then
        Exit Sub
    ElseIf i = 1 Then 'Pregunta confirmación
        i = MsgBox("Está seguro de " & IIf(Index = 1, "generar", "actualizar") & " el Folio: " & txtCampo(3).Text, vbYesNo + vbInformation + vbDefaultButton1, "Validación")
        If i = vbNo Then
            Exit Sub
        End If
    End If
    
    
    If adors.State Then adors.Close
'    If Index = 1 Then 'agregar
'        adors.Open "{call p_AnalisisGuardaDatos(0," & mlAsuxIF & ",1," & i & ",'" & Format(CDate(txtCampo(4).Text), "dd/mm/yyyy") & "','" & s & "'," & giUsuario & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
'    Else 'Actualizar
'        adors.Open "{call p_AnalisisGuardaDatos(" & mlAnálisis & "," & mlAsuxIF & ",1," & i & ",'" & Format(CDate(txtCampo(4).Text), "dd/mm/yyyy") & "','" & s & "'," & giUsuario & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
'    End If
    If Index = 1 Then 'agregar cambio en nov2019
        'adors.Open "{call p_AnalisisGuardaDatos_3(0," & mlAsuxIF & ",1,'" & txtCampo(3).Text & "','" & Format(CDate(txtCampo(4).Text), "dd/mm/yyyy") & "','" & s & "'," & giUsuario & ",'" & Replace(txtCampo(7).Text, "'", "''") & "','" & msModXCau & "'," & miModulo & ",'" & msModInc & "')}", gConSql, adOpenForwardOnly, adLockReadOnly
        adors.Open "{call p_AnalisisGuardaDatosDA2(0," & mlAsuxIF & ",1,'" & txtCampo(3).Text & "','" & Format(CDate(txtCampo(4).Text), "dd/mm/yyyy") & "','" & s & "'," & iOtorgados & ",'" & sAlegatos & "','" & sPruebas & "'," & giUsuario & ",'" & Replace(txtCampo(7).Text, "'", "''") & "','" & msModXCau & "'," & miModulo & ",'" & msModInc & "')}", gConSql, adOpenForwardOnly, adLockReadOnly
    Else 'Actualizar
        'adors.Open "{call p_AnalisisGuardaDatos_3(" & mlAnálisis & "," & mlAsuxIF & ",1,'" & txtCampo(3).Text & "','" & Format(CDate(txtCampo(4).Text), "dd/mm/yyyy") & "','" & s & "'," & giUsuario & ",'" & Replace(txtCampo(7).Text, "'", "''") & "','" & msModXCau & "'," & miModulo & ",'" & msModInc & "')}", gConSql, adOpenForwardOnly, adLockReadOnly
        adors.Open "{call p_AnalisisGuardaDatosDA2(" & mlAnálisis & "," & mlAsuxIF & ",1,'" & txtCampo(3).Text & "','" & Format(CDate(txtCampo(4).Text), "dd/mm/yyyy") & "','" & s & "'," & iOtorgados & ",'" & sAlegatos & "','" & sPruebas & "'," & giUsuario & ",'" & Replace(txtCampo(7).Text, "'", "''") & "','" & msModXCau & "'," & miModulo & ",'" & msModInc & "')}", gConSql, adOpenForwardOnly, adLockReadOnly
    End If
    
    If adors(0) > 0 Then 'Se guardo Correcto
        MsgBox IIf(Index = 1, "Se ingreso un nuevo", "Se actualizo el") & " oficio: " & adors(1), vbOKOnly + vbInformation, ""
'        If Index = 1 Then
'            Call cmbCampo_Click(0)
'            cmdProceso_Click (0)
'        Else
            Call cmbCampo_Click(0) 'Actualiza la lista de oficios
'        End If
        If Not ListView1.Enabled Then ListView1.Enabled = True
        Exit Sub
    Else 'hubo problemas
        MsgBox IIf(Index = 0, "No se ingreso el", "No se actualizo el") & " oficio", vbOKOnly + vbCritical, ""
        Exit Sub
    End If
    Image3.Picture = Imagenes.ListImages(2).Picture
    If ListCausaLey.Height <> 2400 Then ListCausaLey.Height = 2400
    msModInc = ""
    msModIncD = ""
    If Not ListView1.Enabled Then ListView1.Enabled = True
ElseIf Index = 2 Then 'Editar
    cmdProceso(0).Enabled = False 'nuevo
    cmdProceso(1).Enabled = False 'agregar
    cmdProceso(2).Enabled = False 'editar
    cmdProceso(3).Enabled = True 'actualizar
    cmdProceso(4).Enabled = True 'deshacer
    cmdProceso(5).Enabled = False 'borrar
    txtCampo(3).Enabled = True
    txtCampo(4).Enabled = True
    txtCampo(7).Enabled = True
    txtCampo(9).Enabled = True
    txtCampo(10).Enabled = True
    txtCampo(11).Enabled = True
    txtCampo(13).Enabled = True
    txtCampo(14).Enabled = True
    txtCampo(13).Text = ""
    txtCampo(14).Text = ""
    If adors.State Then adors.Close
    adors.Open "select f_nuevofolio(2,0,0) from dual", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        sFolioProEsperado = IIf(IsNull(adors(0)), "CONDUSEF/VJ/DGSL/DAS/" & "/" & Format(AhoraServidor, "yyyy"), adors(0))
    Else
        sFolioProEsperado = "CONDUSEF/VJ/DGSL/DAS/" & "/" & Format(AhoraServidor, "yyyy")
    End If
    
    ListCausaLey.Enabled = True
    Image3.Picture = Imagenes.ListImages(2).Picture
    Frame7.Enabled = True
    msModInc = ""
    msModIncD = ""
    If ListView1.Enabled Then ListView1.Enabled = False
    'Obtiene daots del incumplimiento en caso de tratarse de la DGEV
    If miModulo > 0 And InStr(csUniDGEV, "|" & miUni & "|") > 0 Then
        If adors.State Then adors.Close
        adors.Open "{call p_mod_cauaso_DGEV(" & miModulo & "," & mlAnálisis & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
        If Not adors.EOF Then
            msModInc = IIf(IsNull(adors(0)), "", adors(0))
            msModXCau = IIf(IsNull(adors(1)), "", adors(1))
            msModIncD = IIf(IsNull(adors(2)), "", adors(2))
        End If
    End If
ElseIf Index = 4 Then 'Deshacer
    If MsgBox("¿Está seguro de ignorar los cambios?", vbQuestion + vbYesNo + vbDefaultButton2, "") = vbNo Then
        Exit Sub
    End If
    Call cmbCampo_Click(0)
    ListCausaLey.Clear
    If ListCausaLey.Height <> 2400 Then ListCausaLey.Height = 2400
    msModInc = ""
    msModIncD = ""
    If Not ListView1.Enabled Then ListView1.Enabled = True
ElseIf Index = 5 Then 'borrar
    Dim iRows As Integer
    If MsgBox("Está seguro de borrar el registro seleccionado", vbYesNo + vbQuestion + vbDefaultButton2, "") = vbYes Then
        If adors.State Then adors.Close
        adors.Open "{call p_analisis_borrareg(" & mlAnálisis & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
        If Not adors.EOF Then
            MsgBox "Se borró el registro seleccionado", vbOKOnly, ""
            Call cmbCampo_Click(0)
        Else
            MsgBox "No se borró el registro seleccionado", vbOKOnly + vbInformation, ""
        End If
    End If
    If ListCausaLey.Height <> 2400 Then ListCausaLey.Height = 2400
    msModInc = ""
    msModIncD = ""
    If Not ListView1.Enabled Then ListView1.Enabled = True
End If
Exit Sub
ErrorGuardaDatos:
If gConSql.Errors.Count > 0 Then
    yError = MsgBox("Error: " + gConSql.Errors(0).Description, vbAbortRetryIgnore + vbCritical, "Error no esperado (" + Str(gConSql.Errors(0).Number) + ")")
Else
    yError = MsgBox("Error: " + Err.Description, vbAbortRetryIgnore + vbCritical, "Error no esperado (" + Str(Err.Number) + ")")
End If


If yError = vbRetry Then
    Resume
ElseIf yError = vbIgnore Then
    Resume Next
End If


End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdProcesoImp_Click(Index As Integer)
Dim adors  As New ADODB.Recordset
Dim i As Long
On Error GoTo ErrorGuardaDatos:
If Index = 0 Then 'Nuevo
    cmdProcesoImp(0).Enabled = False 'nuevo
    cmdProcesoImp(1).Enabled = True 'agregar
    cmdProcesoImp(2).Enabled = False 'editar
    cmdProcesoImp(3).Enabled = False 'actualizar
    cmdProcesoImp(4).Enabled = True 'deshacer
    cmdProcesoImp(5).Enabled = False 'borrar
    ListCausaLey.Clear
    'txtCampo(5).Enabled = True
    'txtcampo(5).Text = "Automático" '"ACUERDO/DAS//2010"
    If adors.State Then adors.Close
    adors.Open "select f_nuevofolio(3,0," & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        txtCampo(5).Text = adors(0)
        sFolioImpEsperado = adors(0)
    Else
        txtCampo(5).Text = "ACUERDO/DAS//2010"
        sFolioImpEsperado = txtCampo(5).Text
    End If
    txtCampo(5).Enabled = True
    txtCampo(6).Text = ""
    txtCampo(6).Enabled = True
    txtCampo(8).Text = ""
    txtCampo(8).Enabled = True
    txtCampo(12).Text = ""
    txtCampo(12).Enabled = True
    Frame4.Enabled = True
    txtCampo(6).SetFocus
    msLeyesImp = ""
    msCausasIMP = ""
    If miModulo > 0 Then
        If adors.State Then adors.Close
        adors.Open "Select f_mod_cauxaso(" & miModulo & ") FROM DUAL", gConSql, adOpenStatic, adLockReadOnly
        If Len(adors(0)) > 0 Then 'Existen Causas por asociar
            cmdAgregaCausa(5).Enabled = True
            msModXCauXAso = adors(0)
        End If
    End If
ElseIf Index = 1 Or Index = 3 Then 'agregar / actualizar
    If Len(Trim(txtCampo(12).Text)) = 0 Or Len(Trim(txtCampo(6).Text)) = 0 Or InStr(msCausasIMP, ",") = 0 Then
        MsgBox "La fecha del acuerdo, la fecha de infracción y causa(s) de improcedencia son datos requeridos. Favor de capturarlos.", vbOKOnl´ + vbInformation, ""
        Exit Sub
    End If
    If Index = 1 Then
        If adors.State Then adors.Close
        adors.Open "select count(*) from análisis where oficio='" & Replace(txtCampo(5).Text, "'", "''") & "'" & IIf(Index = 1, "", " and id<>" & mlAnálisis), gConSql, adOpenStatic, adLockReadOnly
        If adors(0) > 0 Then
            MsgBox "El número de Oficio ya existe en la base de datos, no puede repetirse.(" & txtCampo(5).Text & ") Favor de verificar el No. de Oficio", vbOKOnly + vbInformation, ""
            Exit Sub
        End If
    End If
    
    s = Replace(msCausasIMP, ",", "|")

'    s1 = msLeyesImp
'    s2 = msCausasIMP
'    Do While InStr(s1, ",")
'        s = s & Val(s1) & "|" & Val(s2) & "|"
'        s1 = Mid(s1, InStr(s1, ",") + 1)
'        s2 = Mid(s2, InStr(s2, ",") + 1)
'    Loop
    
    
''    If adors.State Then adors.Close
''    adors.Open "select f_nuevofolio(3,0," & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
''    If Not adors.EOF Then
''        s1 = adors(0)
''        If InStr(s1, "???") Then
''            i = F_PreguntaConsecutivo(3, s1)
''        End If
''    End If
''    If Not ValidaFolioAcuerdo(txtCampo(5).Text) Then
''        MsgBox "El número de Oficio no es correcto debe contener el siguiente Formato (ACUERDO/DAS/Cosecutivo/Año)  Favor de corregirlo", vbOKOnly + vbInformation, ""
''        Exit Sub
''    End If
    If adors.State Then adors.Close
    adors.Open "select f_nuevofolio(3,0," & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    If Not IsNull(adors(0)) Then
        sFolioImpEsperado = adors(0)
    End If
    
    i = ValidaFolio(txtCampo(5).Text, sFolioImpEsperado)
    If i = 0 Then
        Exit Sub
    ElseIf i = 1 Then 'Pregunta confirmación
        i = MsgBox("Está seguro de " & IIf(Index = 1, "generar", "actualizar") & " el Folio: " & txtCampo(5).Text, vbYesNo + vbInformation + vbDefaultButton1, "Validación")
        If i = vbNo Then
            Exit Sub
        End If
    End If
    
    
    
    If adors.State Then adors.Close
    If Index = 1 Then 'agregar
        adors.Open "{call p_AnalisisGuardaDatosimp(0," & mlAsuxIF & ",0,'" & txtCampo(5).Text & "','" & Format(CDate(txtCampo(6).Text), "dd/mm/yyyy") & "','" & s & "'," & giUsuario & ",'" & Replace(txtCampo(8).Text, "'", "''") & "','" & msModXCauXAso & "'," & miModulo & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
    Else 'Actualizar
        adors.Open "{call p_AnalisisGuardaDatosimp(" & mlAnálisisImp & "," & mlAsuxIF & ",0,'" & txtCampo(5).Text & "','" & Format(CDate(txtCampo(6).Text), "dd/mm/yyyy") & "','" & s & "'," & giUsuario & ",'" & Replace(txtCampo(8).Text, "'", "''") & "','" & msModXCauXAso & "'," & miModulo & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
    End If
    
'    If adors.State Then adors.Close
'    If Index = 1 Then 'agregar
'        adors.Open "{call p_AnalisisGuardaDatos(0," & mlAsuxIF & ",0," & i & ",'" & Format(CDate(txtcampo(6).Text), "dd/mm/yyyy") & "','" & s & "'," & giUsuario & ",'" & Replace(txtcampo(8).Text, "'", "''") & "')}", gConSql, adOpenForwardOnly, adLockReadOnly
'    Else 'Actualizar
'        adors.Open "{call p_AnalisisGuardaDatos(" & mlAnálisisImp & "," & mlAsuxIF & ",0," & i & ",'" & Format(CDate(txtcampo(6).Text), "dd/mm/yyyy") & "','" & s & "'," & giUsuario & ",'" & Replace(txtcampo(8).Text, "'", "''") & "')}", gConSql, adOpenForwardOnly, adLockReadOnly
'    End If
    
    If adors(0) > 0 Then 'Se guardo Correcto
        'MsgBox IIf(Index = 1, "Se ingreso un nuevo", "Se actualizo el") & " registro.", vbOKOnly + vbInformation, ""
        MsgBox IIf(Index = 1, "Se ingreso un nuevo", "Se actualizo el") & " acuerdo: " & adors(1), vbOKOnly + vbInformation, ""
        Call cmbCampo_Click(0)
        Exit Sub
    Else 'hubo problemas
        MsgBox IIf(Index = 0, "No se ingreso el", "No se actualizo el") & " registro", vbOKOnly + vbCritical, ""
        Exit Sub
    End If
    Image3.Picture = Imagenes.ListImages(2).Picture
    If ListCausaLey.Height <> 2400 Then ListCausaLey.Height = 2400
ElseIf Index = 2 Then 'Editar
    cmdProcesoImp(0).Enabled = False 'nuevo
    cmdProcesoImp(1).Enabled = False 'agregar
    cmdProcesoImp(2).Enabled = False 'editar
    cmdProcesoImp(3).Enabled = True 'actualizar
    cmdProcesoImp(4).Enabled = True 'deshacer
    cmdProcesoImp(5).Enabled = False 'borrar
    txtCampo(5).Enabled = True
    txtCampo(6).Enabled = True
    txtCampo(8).Enabled = True
    Image2.Picture = Imagenes.ListImages(2).Picture
    Frame4.Enabled = True
ElseIf Index = 4 Then 'Deshacer
    If adors.State Then adors.Close
    adors.Open "select * from análisis where idregxif=" & mlAsuxIF & " and procedente=0", gConSql, adOpenStatic, adLockReadOnly
    If adors.EOF Then 'nuevo
        cmdProcesoImp(0).Enabled = True 'nuevo
        cmdProcesoImp(1).Enabled = False 'agregar
        cmdProcesoImp(2).Enabled = False 'editar
        cmdProcesoImp(3).Enabled = False 'actualizar
        cmdProcesoImp(4).Enabled = False 'deshacer
        cmdProcesoImp(5).Enabled = False 'borrar
        txtCampo(5).Enabled = False
        txtCampo(6).Enabled = False
        txtCampo(8).Enabled = False
        txtCampo(5).Text = ""
        txtCampo(6).Text = ""
        Image2.Picture = Imagenes.ListImages(1).Picture
        Frame4.Enabled = False
        cmbCampo(1).ListIndex = -1
        cmbCampo(2).ListIndex = -1
        cmbCampo(3).ListIndex = -1
    Else 'Solo editar y borrar
        cmdProcesoImp(0).Enabled = False 'nuevo
        cmdProcesoImp(1).Enabled = False 'agregar
        cmdProcesoImp(2).Enabled = True 'editar
        cmdProcesoImp(3).Enabled = False 'actualizar
        cmdProcesoImp(4).Enabled = False 'deshacer
        cmdProcesoImp(5).Enabled = True 'borrar
        cmbCampo(1).ListIndex = -1
        cmbCampo(2).ListIndex = -1
        cmbCampo(3).ListIndex = -1
        Frame4.Enabled = False
        txtCampo(5).Enabled = False
        txtCampo(6).Enabled = False
        txtCampo(8).Enabled = False
        txtCampo(5).Text = adors!oficio
        txtCampo(6).Text = adors!FECHA
    End If
    ListCausaLeyImp.Clear
ElseIf Index = 5 Then 'borrar
    Dim iRows As Integer
    If MsgBox("Está seguro de borrar el registro seleccionado", vbYesNo + vbQuestion, "") = vbYes Then
        
        If adors.State Then adors.Close
        adors.Open "{call p_analisis_borrareg(" & mlAnálisisImp & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
        If Not adors.EOF Then
            If adors(0) > 0 Then
                MsgBox "Se borró el registro seleccionado", vbOKOnly, ""
                Call cmbCampo_Click(0)
                Exit Sub
            End If
        End If
        MsgBox "No se borró el registro seleccionado", vbOKOnly + vbInformation, ""
    End If
End If
Exit Sub
ErrorGuardaDatos:
If gConSql.Errors.Count > 0 Then
    yError = MsgBox("Error: " + gConSql.Errors(0).Description, vbAbortRetryIgnore + vbCritical, "Error no esperado (" + Str(gConSql.Errors(0).Number) + ")")
Else
    yError = MsgBox("Error: " + Err.Description, vbAbortRetryIgnore + vbCritical, "Error no esperado (" + Str(Err.Number) + ")")
End If


If yError = vbRetry Then
    Resume
ElseIf yError = vbIgnore Then
    Resume Next
End If
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub Form_Activate()
Dim i As Integer
If gs = ">>" Or gs = "<<" Then
    If Len(gs1) > 0 Then
        cmdActualpen_Click
        If cmdContinuar.Enabled Then
            txtNuevoExp.Text = gs1
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
Dim adors As New ADODB.Recordset
LlenaCombo cmbCampo(3), "select id,descripción from motivosimp where fechabaja is null", "", True
ActualizaPendientes
If adors.State Then adors.Close
adors.Open "select f_dgev_iduni from dual", gConSql, adOpenStatic, adLockReadOnly
If Not adors.EOF Then
     csUniDGEV = adors(0)
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
gs = ""
End Sub

Private Sub ListCausaLey_Click()
Dim iLey As Integer, iCau As Integer, iMot As Integer
Dim s As String
Dim adors As New ADODB.Recordset
If ListCausaLey.ListIndex >= 0 Then
    iLey = Round(ListCausaLey.ItemData(ListCausaLey.ListIndex) / 1000, 0)
    iCau = ListCausaLey.ItemData(ListCausaLey.ListIndex) Mod 1000
    micau = iCau
    i = BuscaCombo(cmbCampo(4), iLey, True, False)
    If i >= 0 Then
        cmbCampo(4).ListIndex = i
    End If
    i = BuscaCombo(cmbCampo(5), iCau, True, False)
    If i >= 0 Then
        cmbCampo(5).ListIndex = i
    End If
    If InStr("|" & msCausas, "|" & iLey & "," & iCau & ",") > 0 Then
        s = Mid(msCausas, InStr("|" & msCausas, "|" & iLey & "," & iCau & ","))
        s = Mid(s, InStr(s, ",") + 1) 'Quita idley
        s = Mid(s, InStr(s, ",") + 1)  'Quita idcau
        txtCampo(13).Text = Mid(s, 1, InStr(s, ",") - 1) 'Obtiene infracción
        s = Mid(s, InStr(s, ",") + 1)  'Quita infracción
        txtCampo(14).Text = Mid(s, 1, InStr(s, "|") - 1) 'Obtiene monto
    End If
    If InStr(csUniDGEV, "|" & miUni & "|") > 0 Then
        If ListCausaLey.Height <> 1100 Then ListCausaLey.Height = 1100
        If Len(msModInc) = 0 Then
            If adors.State Then adors.Close
            adors.Open "select f_anamod_cauinc1(" & mlAnálisis & "),f_anamod_incdesc(" & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
            If Not adors.EOF Then
                msModInc = IIf(IsNull(adors(0)), "", adors(0))
                msModIncD = IIf(IsNull(adors(1)), "", adors(1))
            End If
        End If
        ActualizaListasIncump micau
    End If
End If
End Sub

'Actualiza valores de los combo cuando se selecciona una causa de la lista de casuas de improcedencia
Private Sub ListCausaLeyImp_Click()
Dim iLey As Integer, iCau As Integer, iMot As Integer
If ListCausaLeyImp.ListIndex >= 0 Then
    iLey = Round(ListCausaLeyImp.ItemData(ListCausaLeyImp.ListIndex) / 1000000, 0)
    iCau = Round(ListCausaLeyImp.ItemData(ListCausaLeyImp.ListIndex) / 1000, 0) Mod 1000
    iMot = ListCausaLeyImp.ItemData(ListCausaLeyImp.ListIndex) Mod 1000
    sInfraccion = Mid(ListCausaLeyImp.Text, InStrRev(ListCausaLeyImp.Text, "(") + 1)
    sInfraccion = Mid(sInfraccion, 1, Len(sInfraccion) - 1)
    If IsDate(sInfraccion) Then
        txtCampo(12).Text = sInfraccion
    Else
        txtCampo(12).Text = ""
    End If
    i = BuscaCombo(cmbCampo(1), iLey, True, False)
    If i >= 0 Then
        cmbCampo(1).ListIndex = i
    End If
    i = BuscaCombo(cmbCampo(2), iCau, True, False)
    If i >= 0 Then
        cmbCampo(2).ListIndex = i
    End If
    i = BuscaCombo(cmbCampo(3), iMot, True, False)
    If i >= 0 Then
        cmbCampo(3).ListIndex = i
    End If
End If
End Sub


Private Sub ListIncAcep_DblClick()
Dim i As Integer, n As Integer, s As String, iInc As Integer
Dim ss As String
If ListIncAcep.ListIndex >= 0 Then
'    If ListIncAcep.ListCount = 1 Then
'        MsgBox "Debe existir por lo menos un incumplimiento aceptado. No es permitido quitar todos", vbOKOnly + vbInformation, "Validación"
'        Exit Sub
'    End If
    iInc = ListIncAcep.ItemData(ListIncAcep.ListIndex)
    i = InStr("|" & msModInc, "|" & micau & ":") 'Localiza la posición de la causa
    If i > 0 Then 'Ubica la cadena con sus incumplimientos y reemplaza el '_1' por '_0'
        s = Mid("|" & msModInc, i + 1)
        s = Mid(s, 1, 1 + InStr(Mid(s, 2), "|"))
        ss = Replace(Replace(s, "_" & iInc & "_1", "_" & iInc & "_0"), "," & iInc & "_1", "," & iInc & "_0")
        msModInc = Replace(msModInc, s, ss)
        ActualizaListasIncump micau
    End If

End If
End Sub

Private Sub ListIncNoAcep_DblClick()
Dim i As Integer, n As Integer, s As String, iInc As Integer
Dim ss As String
If ListIncNoAcep.ListIndex >= 0 Then
    iInc = ListIncNoAcep.ItemData(ListIncNoAcep.ListIndex)
    i = InStr("|" & msModInc, "|" & micau & ":") 'Localiza la posición de la causa
    If i > 0 Then 'Ubica la cadena con sus incumplimientos y reemplaza el '_0' por '_1'
        s = Mid("|" & msModInc, i + 1)
        s = Mid(s, 1, 1 + InStr(Mid(s, 2), "|"))
        ss = Replace(Replace(s, "_" & iInc & "_0", "_" & iInc & "_1"), "," & iInc & "_0", "," & iInc & "_1")
        msModInc = Replace(msModInc, s, ss)
        ActualizaListasIncump micau
    End If
    
End If
End Sub


Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim adors As New ADODB.Recordset
If ListCausaLey.Height <> 2400 Then ListCausaLey.Height = 2400
If Not cmdProceso(1).Enabled And Not cmdProceso(3).Enabled Then
    mlAnálisis = Val(Item.Tag)
    txtCampo(3).Text = Item.Text
    txtCampo(4).Text = Item.SubItems(1)
    txtCampo(7).Text = Item.SubItems(4)
    txtCampo(9).Text = Item.SubItems(6)
    txtCampo(10).Text = Item.SubItems(7)
    txtCampo(11).Text = Item.SubItems(5)
    If IsDate(txtCampo(9).Text) Or IsDate(txtCampo(10).Text) Then
        chkPruebas.Value = 1
    Else
        chkPruebas.Value = 0
    End If
    mlAnálisis = Val(Item.Tag)
    RefrescaCausas mlAnálisis
    cmdProceso(0).Enabled = True 'nuevo
    cmdProceso(1).Enabled = False 'agregar
    cmdProceso(2).Enabled = True 'editar
    cmdProceso(3).Enabled = False 'actualizar
    cmdProceso(4).Enabled = False 'deshacer
    cmdProceso(5).Enabled = True 'borrar
    txtCampo(3).Enabled = False
    txtCampo(4).Enabled = False
    txtCampo(7).Enabled = False
    txtCampo(9).Enabled = False
    txtCampo(10).Enabled = False
    txtCampo(11).Enabled = False
    txtCampo(13).Enabled = False
    txtCampo(14).Enabled = False
    ListCausaLey.Enabled = False
    If adors.State Then adors.Close
    adors.Open "select f_registro_iduni(f_analisis_idreg(" & mlAnálisis & ")) from dual", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        miUni = adors(0)
    End If
End If
End Sub

Private Sub txtCampo_Change(Index As Integer)
If Index >= 5 And Index <= 6 Then
    Image2.Picture = Imagenes.ListImages(2).Picture
    mbCambioImp = True
End If
End Sub

Private Sub txtCampo_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 1 And InStr("-", Chr(KeyAscii)) > 0 Then Exit Sub
KeyAscii = TeclaOprimida(txtCampo(Index), KeyAscii, txtCampo(Index).Tag, False)
'MsgBox "asd"
End Sub

Private Sub RefrescaCausas(iAna As Long)
Dim adors As ADODB.Recordset
Set adors = New ADODB.Recordset
adors.Open "SELECT r.idley,r.idcau,nvl(l.descripción,''),nvl(c.descripción,''),infraccion,monto FROM análisiscausas r, leyes l, causas c where R.IDANA=" & iAna & " and r.idley=l.id and r.idcau=c.id ORDER BY c.descripción,l.descripción", gConSql, adOpenStatic, adLockReadOnly
ListCausaLey.Clear
msCausas = ""
'msLeyes = ""
Do While Not adors.EOF
    'msLeyes = msLeyes & adors(0) & ","
    'msCausas = msCausas & adors(1) & ","
    msCausas = msCausas & adors(0) & "," & adors(1) & "," & Format(adors(4), gsFormatoFecha) & "," & adors(5) & "|"
    ListCausaLey.AddItem adors(3) & " (" & adors(2) & ") (" & Format(adors(4), gsFormatoFecha) & ") (" & adors(5) & ")"
    ListCausaLey.ItemData(ListCausaLey.NewIndex) = (adors(0) * 1000 + adors(1))
    adors.MoveNext
Loop
End Sub

Private Sub txtCampo_LostFocus(Index As Integer)
Dim adors As New ADODB.Recordset
If Mid(txtCampo(Index).Tag, 1, 1) = "f" Then
    If IsDate(txtCampo(Index).Text) Then
        d = CDate(txtCampo(Index).Text)
        txtCampo(Index).Text = Format(d, gsFormatoFecha)
        adors.Open "select sysdate from dual", gConSql, adOpenStatic, adLockReadOnly
        If Int(adors(0)) + 60 - Int(d) < 0 Then
            Call MsgBox("Fecha no válida. No se permite ingresar fecha mayor a 60 posteriores a la fecha actual (" & Format(adors(0), gsFormatoFecha) & ")", vbOKOnly + vbInformation, "")
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
adors.Open "{call p_analisispendientes2(" & giResponsable & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
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

Private Sub txtNuevoExp_Change()
If mbLimpiaExp Then
    mbLimpiaExp = False
    cmbPendientes.ListIndex = -1
End If
End Sub

Private Sub txtNuevoExp_GotFocus()
mbLimpiaExp = True
End Sub

Private Sub txtNuevoExp_LostFocus()
mbLimpiaExp = False
End Sub

'Bloquea o desbloque controles según la variable bDesBloquea en la sección según iProcedente
'0: Improcedentes; 1: Procedentes; 2: Ambos
Private Sub BloqueaControles(iProcedente As Byte, Optional bDesBloquea As Boolean)
If iProcedente = 1 Or iProcedente = 2 Then
    Frame7.Enabled = bDesBloquea
    txtCampo(4).Enabled = bDesBloquea
    txtCampo(7).Enabled = bDesBloquea
End If
If iProcedente <> 1 Or iProcedente = 2 Then
    Frame4.Enabled = bDesBloquea
    txtCampo(5).Enabled = bDesBloquea
    txtCampo(6).Enabled = bDesBloquea
    txtCampo(8).Enabled = bDesBloquea
End If
End Sub


'Private Sub Redondear_Botón(Botón As Command, Radio As Long)
'
'Dim Region As Long
'Dim Ret As Long
'Dim Ancho As Long
'Dim Alto As Long
'Dim old_Scale As Integer
'
'    ' guardar la escala
'    old_Scale = Botón.ScaleMode
'
'    ' cambiar la escala a pixeles
'    Botón.ScaleMode = vbPixels
'
'    'Obtenemos el ancho y alto de la region del Form
'    Ancho = Botón.ScaleWidth
'    Alto = Botón.ScaleHeight
'
'    'Pasar el ancho alto del formualrio y el valor de redondeo .. es decir el radio
'    Region = CreateRoundRectRgn(0, 0, Ancho, Alto, Radio, Radio)
'
'    ' Aplica la región al formulario
'    Ret = SetWindowRgn(Botón.hwnd, Region, True)
'
'    ' restaurar la escala
'    Botón.ScaleMode = old_Scale
'
'End Sub

Function ValidaFolioAcuerdo(s As String) As Boolean
Dim ss As String
Dim adors As New ADODB.Recordset
Dim s1 As String
adors.Open "select f_nuevofolio(3,0," & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
If Len(Trim(adors(0))) > 0 Then
    s1 = adors(0)
End If
If InStr(s, "ACUERDO/DAS/") = 0 Then 'No tiene la primera parte
    Exit Function
End If
ss = Mid(s, InStr(s, "ACUERDO/DAS/") + 12)
If Val(ss) = 0 Then 'EL cONSECUTIVO DEBE SER MAYOR A CERO
    Exit Function
End If
If InStr(ss, "/") = 0 Then 'No tiene la última parte (AÑO)
    Exit Function
End If
If Val(Mid(ss, InStr(ss, "/") + 1)) = 0 Then 'EL AÑO debe se mayor a cero
    Exit Function
End If
ValidaFolioAcuerdo = True
End Function
