VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form Registro 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro"
   ClientHeight    =   9960
   ClientLeft      =   6120
   ClientTop       =   2130
   ClientWidth     =   12315
   Icon            =   "Registro.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9960
   ScaleWidth      =   12315
   Begin VB.VScrollBar VScroll 
      Height          =   9468
      LargeChange     =   50
      Left            =   11970
      Max             =   100
      SmallChange     =   10
      TabIndex        =   72
      Top             =   405
      Visible         =   0   'False
      Width           =   315
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   767
      ButtonWidth     =   635
      ButtonHeight    =   609
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   34
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nuevo Registro"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Limpiar"
            Object.ToolTipText     =   "Limpiar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Guardar"
            Object.ToolTipText     =   "Guardar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Deshacer"
            Object.ToolTipText     =   "Deshacer"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ir_a"
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Primero"
            Object.ToolTipText     =   "Primer Registro"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Anterior"
            Object.ToolTipText     =   "Anterior Reg"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Siguiente Reg"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Último"
            Object.ToolTipText     =   "Último Reg"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3645
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   45
         Visible         =   0   'False
         Width           =   912
      End
      Begin VB.TextBox txtNoReg 
         Height          =   375
         Left            =   3465
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   0
         Width           =   1188
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00B4E2C9&
      Height          =   1692
      Left            =   1980
      TabIndex        =   31
      Top             =   396
      Width           =   10005
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7380
         Top             =   135
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
               Picture         =   "Registro.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Registro.frx":03D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Registro.frx":0798
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Registro.frx":0C5E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Registro.frx":1024
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Registro.frx":13EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Registro.frx":17B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Registro.frx":1B76
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Registro.frx":1F3C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Registro.frx":2302
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Registro.frx":26C8
               Key             =   ""
            EndProperty
         EndProperty
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
         Height          =   372
         Left            =   7965
         Picture         =   "Registro.frx":2A8E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1176
         Width           =   1455
      End
      Begin VB.OptionButton OpcPAU 
         BackColor       =   &H00B4E2C9&
         Caption         =   "No"
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   1
         Left            =   3645
         TabIndex        =   1
         Top             =   1104
         Width           =   555
      End
      Begin VB.OptionButton OpcPAU 
         BackColor       =   &H00B4E2C9&
         Caption         =   "Sí"
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   0
         Left            =   3645
         TabIndex        =   0
         Top             =   792
         Width           =   510
      End
      Begin VB.TextBox txtExpediente 
         BackColor       =   &H8000000F&
         DataField       =   "n_cvepersona"
         Height          =   285
         Left            =   5310
         MaxLength       =   20
         TabIndex        =   2
         Tag             =   "c"
         ToolTipText     =   """Numero consecutivo de registro"""
         Top             =   888
         Width           =   1800
      End
      Begin VB.TextBox txtNuevoExp 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   5310
         MaxLength       =   35
         TabIndex        =   3
         Tag             =   "c"
         Top             =   1248
         Width           =   2490
      End
      Begin MSForms.CommandButton cmdIrAna 
         Height          =   555
         Left            =   8010
         TabIndex        =   4
         Top             =   270
         Width           =   1500
         BackColor       =   14735199
         Caption         =   "Ir a Análisis >>"
         Size            =   "2646;979"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Eti 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00B4E2C9&
         Caption         =   "Módulo de Registro"
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
         Index           =   4
         Left            =   1530
         TabIndex        =   45
         Top             =   270
         Width           =   6765
         WordWrap        =   -1  'True
      End
      Begin VB.Label Eti 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00B4E2C9&
         Caption         =   "Asunto derivado del Proceso de Atención a Usuarios"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   576
         Index           =   2
         Left            =   180
         TabIndex        =   39
         Top             =   840
         Width           =   3372
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00B4E2C9&
         Caption         =   "Folio SIO:"
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   0
         Left            =   4548
         TabIndex        =   33
         Top             =   888
         Width           =   696
      End
      Begin VB.Label Label2 
         BackColor       =   &H00B4E2C9&
         Caption         =   "Expediente:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4365
         TabIndex        =   32
         Top             =   1260
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Height          =   7860
      Left            =   90
      TabIndex        =   23
      Top             =   2025
      Width           =   11895
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   5130
         TabIndex        =   74
         Top             =   3465
         Width           =   2400
         Begin VB.OptionButton opcIF 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Vigentes"
            Height          =   285
            Index           =   0
            Left            =   45
            TabIndex        =   76
            Top             =   90
            Value           =   -1  'True
            Width           =   1005
         End
         Begin VB.OptionButton opcIF 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Todas"
            Height          =   285
            Index           =   1
            Left            =   1170
            TabIndex        =   75
            Top             =   90
            Width           =   1005
         End
      End
      Begin VB.CommandButton cmdBusIF 
         Caption         =   "Sig"
         Height          =   330
         Left            =   11040
         TabIndex        =   12
         Top             =   3510
         Width           =   510
      End
      Begin VB.TextBox txtBuscarIF 
         Height          =   285
         Left            =   8505
         TabIndex        =   11
         Top             =   3555
         Width           =   2490
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Buscar Solicitudes"
         Height          =   4290
         Left            =   90
         TabIndex        =   66
         Top             =   360
         Width           =   2400
         Begin VB.CommandButton cmdModulos 
            BackColor       =   &H0080C0FF&
            Caption         =   "Busca Asunto(s) Turnado(s) de MSS según criterios Especificados"
            Height          =   1050
            Left            =   270
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Obtiene información desde módulos que turnan a Sanción"
            Top             =   3150
            Width           =   1860
         End
         Begin VB.Frame Frame3 
            Caption         =   "Área Origen"
            Height          =   1725
            Left            =   225
            TabIndex        =   70
            Top             =   675
            Width           =   1950
            Begin VB.OptionButton opcAreOri 
               Caption         =   "Otro"
               Height          =   240
               Index           =   3
               Left            =   225
               TabIndex        =   65
               Top             =   1350
               Width           =   1680
            End
            Begin VB.OptionButton opcAreOri 
               Caption         =   "DGEV"
               Height          =   240
               Index           =   2
               Left            =   225
               TabIndex        =   63
               Top             =   990
               Width           =   1590
            End
            Begin VB.OptionButton opcAreOri 
               Caption         =   "PAU (SIO)"
               Height          =   240
               Index           =   1
               Left            =   225
               TabIndex        =   61
               Top             =   630
               Value           =   -1  'True
               Width           =   1545
            End
            Begin VB.OptionButton opcAreOri 
               Caption         =   "Todas"
               Height          =   240
               Index           =   0
               Left            =   225
               TabIndex        =   59
               Top             =   270
               Width           =   1680
            End
         End
         Begin VB.TextBox txtExpMod 
            Height          =   285
            Left            =   270
            TabIndex        =   67
            Top             =   2700
            Width           =   1815
         End
         Begin VB.Label Eti 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Origen de la solicitud:"
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
            Index           =   3
            Left            =   225
            TabIndex        =   69
            Top             =   360
            Width           =   2025
         End
         Begin VB.Label Label1 
            Caption         =   "Expediente MSS"
            Height          =   240
            Left            =   315
            TabIndex        =   68
            Top             =   2475
            Width           =   1590
         End
      End
      Begin VB.CheckBox chkSinCau 
         Caption         =   "Prod. sin causa"
         Height          =   285
         Left            =   7380
         TabIndex        =   64
         Top             =   4275
         Width           =   1455
      End
      Begin VB.OptionButton opcSIO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RECA"
         Height          =   285
         Index           =   2
         Left            =   6120
         TabIndex        =   57
         Top             =   4320
         Width           =   780
      End
      Begin VB.OptionButton opcSIO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SIO"
         Height          =   285
         Index           =   1
         Left            =   5265
         TabIndex        =   56
         Top             =   4320
         Width           =   690
      End
      Begin VB.OptionButton opcSIO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No Aplica"
         Height          =   285
         Index           =   0
         Left            =   4005
         TabIndex        =   55
         Top             =   4320
         Width           =   1005
      End
      Begin VB.CommandButton cmdPro 
         Caption         =   "..."
         Height          =   330
         Left            =   10935
         TabIndex        =   15
         Top             =   4590
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtResObs 
         BackColor       =   &H8000000A&
         Height          =   780
         Left            =   7020
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Top             =   7065
         Visible         =   0   'False
         Width           =   4400
      End
      Begin VB.TextBox txtCausa 
         BackColor       =   &H00C0C0C0&
         DataField       =   "Nombre"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3330
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   16
         Tag             =   "c"
         ToolTipText     =   "Nombre"
         Top             =   4950
         Width           =   8190
      End
      Begin VB.TextBox txtProducto 
         BackColor       =   &H00C0C0C0&
         DataField       =   "Nombre"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3330
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   14
         Tag             =   "c"
         ToolTipText     =   "Nombre"
         Top             =   4635
         Width           =   7470
      End
      Begin VB.TextBox txtObs 
         BackColor       =   &H8000000F&
         Height          =   735
         Left            =   2565
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         Tag             =   "f"
         ToolTipText     =   "Observaciones"
         Top             =   7065
         Width           =   8880
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "fecha_memorando"
         Height          =   285
         Index           =   2
         Left            =   5805
         MaxLength       =   20
         TabIndex        =   19
         Tag             =   "f"
         ToolTipText     =   "Fecha del Memo de Envío"
         Top             =   5895
         Width           =   2805
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   630
         Picture         =   "Registro.frx":35FD
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   6795
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdAgregarIF 
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
         Height          =   375
         Left            =   10035
         Picture         =   "Registro.frx":4387
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5265
         Width           =   1455
      End
      Begin VB.ComboBox cmbCampo 
         BackColor       =   &H8000000F&
         DataField       =   "idrestur"
         Height          =   315
         Index           =   5
         ItemData        =   "Registro.frx":5210
         Left            =   2550
         List            =   "Registro.frx":5212
         TabIndex        =   21
         ToolTipText     =   "Responsable a quien se turna el expediente"
         Top             =   6465
         Width           =   8895
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "recepción"
         Height          =   285
         Index           =   3
         Left            =   8910
         MaxLength       =   20
         TabIndex        =   20
         Tag             =   "f"
         ToolTipText     =   "Recepción del Expediente en el área de Sanciones"
         Top             =   5895
         Width           =   2580
      End
      Begin VB.ComboBox cmbCampo 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   4
         ItemData        =   "Registro.frx":5214
         Left            =   2565
         List            =   "Registro.frx":5216
         TabIndex        =   13
         ToolTipText     =   "Institución"
         Top             =   3870
         Width           =   8985
      End
      Begin VB.ComboBox cmbCampo 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   3
         ItemData        =   "Registro.frx":5218
         Left            =   2610
         List            =   "Registro.frx":521A
         TabIndex        =   10
         ToolTipText     =   "Clase de Institución"
         Top             =   2970
         Width           =   8895
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "memorando"
         Height          =   285
         Index           =   1
         Left            =   2610
         MaxLength       =   60
         TabIndex        =   18
         Tag             =   "c"
         ToolTipText     =   "Número de memo de envío"
         Top             =   5895
         Width           =   3120
      End
      Begin VB.ComboBox cmbCampo 
         BackColor       =   &H8000000F&
         DataField       =   "idoridir"
         Height          =   315
         Index           =   0
         ItemData        =   "Registro.frx":521C
         Left            =   2610
         List            =   "Registro.frx":521E
         TabIndex        =   6
         ToolTipText     =   "Dirección General de Origen"
         Top             =   630
         Width           =   8895
      End
      Begin VB.ComboBox cmbCampo 
         BackColor       =   &H8000000F&
         DataField       =   "idmat"
         Height          =   315
         Index           =   2
         ItemData        =   "Registro.frx":5220
         Left            =   2610
         List            =   "Registro.frx":5222
         TabIndex        =   8
         ToolTipText     =   "Materia de la Sanción"
         Top             =   1800
         Width           =   8895
      End
      Begin VB.ComboBox cmbCampo 
         BackColor       =   &H8000000F&
         DataField       =   "idoriuni"
         Height          =   315
         Index           =   1
         ItemData        =   "Registro.frx":5224
         Left            =   2610
         List            =   "Registro.frx":523A
         TabIndex        =   7
         ToolTipText     =   "Unidad de Origen"
         Top             =   1215
         Width           =   8895
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "Nombre"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   2610
         MaxLength       =   80
         TabIndex        =   9
         Tag             =   "c"
         ToolTipText     =   "Nombre"
         Top             =   2430
         Width           =   8895
      End
      Begin VB.ListBox ListClaseIns 
         Height          =   255
         Left            =   2610
         TabIndex        =   41
         Top             =   2988
         Width           =   8895
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   9000
         TabIndex        =   58
         Top             =   4275
         Width           =   2310
         Begin VB.OptionButton opcVig 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Todos"
            Height          =   240
            Index           =   1
            Left            =   1305
            TabIndex        =   62
            Top             =   0
            Width           =   915
         End
         Begin VB.OptionButton opcVig 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Vigentes"
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   60
            Top             =   0
            Value           =   -1  'True
            Width           =   960
         End
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Busca IF:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   7740
         TabIndex        =   73
         Top             =   3600
         Width           =   675
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Origen Producto:"
         Height          =   240
         Left            =   2610
         TabIndex        =   54
         Top             =   4365
         Width           =   1365
      End
      Begin VB.Label lblResObs 
         Caption         =   "Responsable de la solicitiud / Observaciones:"
         Height          =   240
         Left            =   7110
         TabIndex        =   50
         Top             =   6840
         Visible         =   0   'False
         Width           =   3795
      End
      Begin VB.Label lblTipoExp 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   4230
         TabIndex        =   49
         Top             =   5310
         Width           =   5685
      End
      Begin VB.Label lblExpMod 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   4770
         TabIndex        =   48
         Top             =   315
         Width           =   6675
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Causa:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   2580
         TabIndex        =   47
         Top             =   4950
         Width           =   495
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Producto:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   2580
         TabIndex        =   46
         Top             =   4695
         Width           =   690
      End
      Begin VB.Label etiObs 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Observaciones:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   2565
         TabIndex        =   43
         Top             =   6855
         Width           =   1110
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Turnar expediente a:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   2565
         TabIndex        =   38
         Top             =   6180
         Width           =   1470
      End
      Begin VB.Label Eti 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Seguimiento a la solicitud:"
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
         Index           =   1
         Left            =   90
         TabIndex        =   37
         Top             =   6165
         Width           =   2415
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha de recepción del área de sanciones:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   8370
         TabIndex        =   36
         Top             =   5685
         Width           =   3075
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha de memorando de envío:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   5805
         TabIndex        =   35
         Top             =   5685
         Width           =   2280
      End
      Begin VB.Label Eti 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Documento de la solicitud:"
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
         Index           =   0
         Left            =   90
         TabIndex        =   34
         Top             =   5715
         Width           =   2460
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Institución:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   2610
         TabIndex        =   30
         Top             =   3690
         Width           =   765
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sector (Clase Institución):"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   2610
         TabIndex        =   29
         Top             =   2745
         Width           =   1800
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "No. de memorando de envío:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   2610
         TabIndex        =   28
         Top             =   5685
         Width           =   2085
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dirección General de origen:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   2610
         TabIndex        =   27
         Top             =   360
         Width           =   2025
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Unidad de origen:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   2610
         TabIndex        =   26
         Top             =   945
         Width           =   1260
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Materia de la sanción:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   2610
         TabIndex        =   25
         Top             =   1530
         Width           =   1560
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nombre(s) del (los) usuario(s) reclamante(s):"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   2610
         TabIndex        =   24
         Top             =   2205
         Width           =   3060
      End
      Begin VB.Label Etilist 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clases e Instituciones:"
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
         Left            =   2610
         TabIndex        =   40
         Top             =   2745
         Visible         =   0   'False
         Width           =   2625
      End
   End
   Begin VB.PictureBox CReport 
      Height          =   480
      Left            =   3240
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   44
      Top             =   360
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   1632
      Left            =   108
      Picture         =   "Registro.frx":5288
      Stretch         =   -1  'True
      Top             =   384
      Width           =   1824
   End
End
Attribute VB_Name = "Registro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents AdorsPrin As ADODB.Recordset, adorsCata As ADODB.Recordset
Attribute AdorsPrin.VB_VarHelpID = -1
Dim msConsulta As String, msOrden As String, msPrin As String, msConsultaP As String
Dim mbLimpia As Boolean, mbCambio As Boolean, mbRefresca As Boolean, mbInicio  As Boolean, mbNoPreg As Boolean, mlAnt As Long
Dim mbNuevo As Boolean 'indicador de asunto nuevo
Dim mbRefrescaDatos As Boolean 'Indicador que está refrescando Datos
Dim mlAsunto As Long, mlBuscaAsu As Long  'clave principal de la persona
Dim mlRegxIf 'Clave de RegxIf seleccionado cuando hya más de una IF
Dim mlAsuSIO As Long, miPr1 As Integer, miPr2 As Integer, miPr3 As Integer, micau As Integer 'Variable con el id del asunto del SIO obtenido
Dim myPermiso As Byte, myTabla As Byte, myPermisoRep As Byte, msFolio As String, miAnio As Integer, miDel As Integer, mlCon As Long
Dim msInstituciones As String 'Información de las varias instituciones
Dim msClaseIns As String 'Información de las varias clases de institución
Dim mlModulo As Long 'id del módulo en caso
Dim miModPro As Integer 'Indica si es procedente o rechazado
Dim miUni As Integer 'Id de la Unidad de origen
Dim miCla As Integer 'id de la clase
Dim miIns As Integer 'id de la IF
Dim msProSel As String 'Guarda información de los Productos seleccionados
Dim miProCambio As Integer 'Indica si los datos del producto están sufriendo cambio para ignorar el evento click de la opción opcSIO
Dim miSIO_SoloPro As Integer 'Indica si solo debe considerar productos y no causas del SIO
Dim msClaIns As String 'Datos de las Clases e Instituciones como las devuelve la selección múltiple, para enviarselas de la misma forma
Const csUniDGEV = "|1011|1012|1013|1018|" 'id de unidad correspondientes a la DGEV


'Actualiza datos por movimiento de registro
Private Sub AdorsPrin_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'On Error GoTo salir:
Static iCom
If iCom = 1 Then
    Exit Sub
    iCom = 0
End If
iCom = 1
If AdorsPrin.RecordCount > 0 Then
    If AdorsPrin.Bookmark > 0 Then
        RefrescaDatos
    End If
End If
iCom = 0
Exit Sub
'salir:
iCom = 0
End Sub


Private Sub chkSinCau_Click()
If chkSinCau.Value Then
    If miSIO_SoloPro = 0 Then miSIO_SoloPro = 2
Else
    If miSIO_SoloPro > 0 Then miSIO_SoloPro = 0
End If
End Sub


'Selección del combo, filtro de instituciones y unidades
Private Sub cmbCampo_Click(Index As Integer)
If Index = 0 And cmbCampo(Index).ListIndex >= 0 Then 'unidades
    LlenaCombo cmbCampo(1), "select u.id,u.descripción from relacióndirecciónunidad rdu, unidades u where rdu.iddir=" & cmbCampo(0).ItemData(cmbCampo(0).ListIndex) & " and rdu.iduni=u.id order by 2", "", True
    If cmbCampo(1).ListCount = 1 Then
        cmbCampo(2).Clear
        cmbCampo(2).Text = ""
        cmbCampo(1).ListIndex = 0
    End If
End If
If Index = 1 And cmbCampo(Index).ListIndex >= 0 Then 'materia de sanción
    LlenaCombo cmbCampo(2), "select ms.id,ms.descripción from relaciónunidadmateria rum, materiasanción ms where rum.iduni=" & cmbCampo(Index).ItemData(cmbCampo(Index).ListIndex) & " and rum.idmat=ms.id order by 2", "", True
    If cmbCampo(2).ListCount = 1 Then
        cmbCampo(2).ListIndex = 0
    End If
End If
If Index = 3 And cmbCampo(Index).ListIndex >= 0 Then 'instituciones
    ActualizaComboIF
End If
If Not mbCambio And Not mbRefresca Then mbCambio = True
If Index = 3 Then 'Clase
    'mipro = 0
    'micau = 0
    'txtProducto.Text = ""
    'txtCausa.Text = ""
End If
End Sub

Sub ActualizaComboIF()
If cmbCampo(3).ListIndex < 0 Then
    Exit Sub
End If

If opcIF(0).Value Then
    LlenaCombo cmbCampo(4), "select i.id,i.descripción from relaciónclaseinstitución rci, instituciones i where rci.idcla=" & cmbCampo(3).ItemData(cmbCampo(3).ListIndex) & " and rci.idins=i.id and i.status=1 order by 2", "", True
Else
    LlenaCombo cmbCampo(4), "select i.id,i.descripción from relaciónclaseinstitución rci, instituciones i where rci.idcla=" & cmbCampo(3).ItemData(cmbCampo(3).ListIndex) & " and rci.idins=i.id order by 2", "", True
End If
If cmbCampo(4).ListCount = 1 Then
    cmbCampo(4).ListIndex = 0
End If
End Sub

Private Sub cmbCampo_KeyPress(Index As Integer, KeyAscii As Integer)
Dim i As Long, i1 As Long
If KeyAscii = 13 Then
    'Manda el cursor al obejeto siguiente en el formulario
    i1 = cmbCampo(Index).TabIndex
    For i = Controls.Count - 1 To 0 Step -1
        If InStr("txt,cmb,cmd,chk,opc", Mid(LCase(Controls(i).Name), 1, 3)) > 0 Then
            'Debug.Print Controls(i).TabIndex
            If Controls(i).TabIndex = i1 + 1 Then
                Controls(i).SetFocus
                Exit Sub
            End If
        End If
    Next
End If
End Sub


Private Sub cmdAgregarIF_Click()
Dim s As String, adors As New ADODB.Recordset, Y As Integer, s1 As String
    'gs = "ArbolVarios-->SELECT rci.idcla,rci.idins,ci.descripción,i.descripción FROM relaciónclaseinstitución rci, claseinstitución ci, instituciones i where rci.idcla=ci.id and rci.idins=i.id and (i.c_status is null or i.c_status<>3) and i.baja=0 ORDER BY ci.descripción,i.descripción"
    gs = "{call P_ClaseInsts(0)}"
    'gs2 = Str(iInsPrincipal) + "," + Str(iClaPrincipal)
    If cmbCampo(3).Visible Then 'una sola IF
        'If cmbCampo(3).ListIndex >= 0 And cmbCampo(4).ListIndex >= 0 Then
        '    gs2 = cmbCampo(4).ItemData(cmbCampo(4).ListIndex) & "," & cmbCampo(3).ItemData(cmbCampo(3).ListIndex)
        'Else
        '    gs2 = Str(iInsPrincipal) + "," + Str(iClaPrincipal)
        'End If
'        If cmbCampo(3).ListIndex >= 0 Then
'            gs3 = cmbCampo(3).ItemData(cmbCampo(3).ListIndex) & ","
'            gs2 = cmbCampo(3).ItemData(cmbCampo(3).ListIndex)
'        Else
'            gs3 = ""
'            gs2 = "0"
'        End If
'        If cmbCampo(4).ListIndex >= 0 Then
'            gs1 = cmbCampo(4).ItemData(cmbCampo(4).ListIndex) & ","
'            gs2 = cmbCampo(4).ItemData(cmbCampo(4).ListIndex) & "," & gs2
'        Else
'            gs1 = ""
'            gs2 = "0," & gs2
'        End If
        If cmbCampo(3).ListIndex >= 0 And cmbCampo(4).ListIndex >= 0 Then 'Obtiene la clase e Inst para pasarla a la selección
            gs2 = "r" & Right("000000000" & cmbCampo(4).ItemData(cmbCampo(4).ListIndex), 10) & Right("000000000" & cmbCampo(3).ItemData(cmbCampo(3).ListIndex), 10) & "|"
        Else
            gs2 = ""
        End If
    Else
'        gs2 = Val(msInstituciones) & "," & Val(msClaseIns)
'        If InStr(msInstituciones, ",") > 0 Then
'            gs1 = Val(msInstituciones) & Mid(msInstituciones, InStr(msInstituciones, ","))
'        Else
'            gs1 = msInstituciones
'        End If
'        If InStr(msClaseIns, ",") > 0 Then
'            gs3 = Val(msClaseIns) & Mid(msClaseIns, InStr(msClaseIns, ","))
'        Else
'            gs3 = msClaseIns
'        End If
        s = msClaseIns
        s1 = msInstituciones
        gs2 = "-->"
        Do While InStr(s, ",") > 0 And InStr(s1, ",") > 0
            gs2 = gs2 & Right("000000000" & Val(s), 10) & Right("000000000" & Val(s1), 10) & "|"
            s = Mid(s, InStr(s, ",") + 1)
            s1 = Mid(s1, InStr(s1, ",") + 1)
        Loop
    End If
    With SelProceso
        .Caption = "Instituciones por clase"
        .TreeView1.CheckBoxes = True
        .Show vbModal
        If Len(gs) > 0 Then
            If InStr(gs1, "|") > 0 Then
                msClaIns = gs1
                If InStr(gs1, "|") <> InStrRev(gs1, "|") Then 'Se tiene más de una selección
                    f_UnaSolaIns False
                    If adors.State > 0 Then adors.Close
                    adors.Open "select lpad(ci.id,10,'0')||lpad(i.id,10,'0') as ClaIns,i.id as idins,ci.descripción as clase,i.descripción as institución from instituciones i, relaciónclaseinstitución rci, claseinstitución ci where lpad(ci.id,10,'0')||lpad(i.id,10,'0') in ('" & Replace(Mid(gs1, 1, Len(gs1) - 1), "|", "','") & "') and i.id=rci.idins and rci.idcla=ci.id", gConSql, adOpenStatic, adLockReadOnly
                    ListClaseIns.Clear
                    s = gs1
                    Do While InStr(s, "|")
                        s1 = Mid(s, 1, InStr(s, "|") - 1)
                        adors.MoveFirst
                        Do While Not adors.EOF
                            If adors(0) = s1 Then
                                ListClaseIns.AddItem adors!institución
                                ListClaseIns.ItemData(ListClaseIns.NewIndex) = adors(1)
                                Exit Do
                            End If
                            adors.MoveNext
                        Loop
                        s = Mid(s, InStr(s, "|") + 1)
                    Loop
                    msInstituciones = gs
                    msClaseIns = gs3
                    msProSel = ""
                    MsgBox "Favor de especificar y/o verificar los Productos asociados a cada Institución"
                    If ListClaseIns.ListIndex >= 0 Then
                        ListClaseIns_Click
                    Else
                        ListClaseIns.ListIndex = 0
                    End If
                Else
                    f_UnaSolaIns True
                End If
'                If InStr(Mid(gs, 1, Len(gs) - 1), ",") > 0 Then
'                    f_UnaSolaIns False
'                    s = gs
'                    Y = 1
'                    If adors.State > 0 Then adors.Close
'                    adors.Open "select rci.idcla,rci.idins,ci.descripción as clase,i.descripción as institución from instituciones i, relaciónclaseinstitución rci, claseinstitución ci where i.id in (" + Mid(gs, 1, Len(gs) - 1) + ") and i.id=rci.idins and rci.idcla=ci.id", gConSql, adOpenStatic, adLockReadOnly
'                    ListClaseIns.Clear
'                    Do While Not adors.EOF
'                        ListClaseIns.AddItem adors!institución
'                        ListClaseIns.ItemData(ListClaseIns.NewIndex) = adors(1)
'                        adors.MoveNext
'                    Loop
'
'                Else
'                    f_UnaSolaIns True
'                    If adors.State > 0 Then adors.Close
'                    adors.Open "select id,idcla,descripción from instituciones i, relaciónclaseinstitución rci where i.id in (" + Mid(gs, 1, Len(gs) - 1) + ") and i.id=rci.idins ", gConSql, adOpenStatic, adLockReadOnly '**************mmm
'                    If cmbCampo(3).ListIndex >= 0 Then
'                        If cmbCampo(3).ItemData(cmbCampo(3).ListIndex) <> adors(1) Then
'                            cmbCampo(3).ListIndex = BuscaCombo(cmbCampo(3), Str(adors(1)), True)
'                        End If
'                    Else
'                        cmbCampo(3).ListIndex = BuscaCombo(cmbCampo(3), Str(adors(1)), True)
'                    End If
'                    cmbCampo(4).ListIndex = BuscaCombo(cmbCampo(4), Str(Val(gs)), True)
'                End If
'                msInstituciones = gs
'                msClaseIns = gs3
            Else
                f_UnaSolaIns True
                msInstituciones = ""
                msClaseIns = ""
                cmbCampo(4).ListIndex = -1
            End If
        End If
    End With
End Sub

'Muestra y oculta objetos según existan más de una Institución
Private Sub f_UnaSolaIns(bActiva As Boolean)
etiCombo(3).Visible = bActiva
etiCombo(4).Visible = bActiva
cmbCampo(3).Visible = bActiva
cmbCampo(4).Visible = bActiva
Etilist.Visible = Not bActiva
ListClaseIns.Visible = Not bActiva
End Sub


Private Sub cmdBusIF_Click()
Dim i As Integer, iPos As Integer
If Len(txtBuscarIF.Text) > 0 And cmbCampo(4).ListCount > 0 Then
    iPos = cmbCampo(4).ListIndex
    If iPos = cmbCampo(4).ListCount - 1 Then
        i = -1
    Else
        i = BuscaCombo(cmbCampo(4), txtBuscarIF.Text, 0, True, 0, iPos + 1)
    End If
    If i >= 0 Then
        cmbCampo(4).ListIndex = i
    ElseIf iPos >= 0 Then
        cmbCampo(4).ListIndex = -1
    End If
End If
End Sub

'Obtiene información del SIO y la coloca en los campos del Formulario
Private Sub cmdContinuar_Click()
Dim conn As New ADODB.Connection, adors As New ADODB.Recordset
Dim s As String, ycon As Integer, ss As String, s2 As String, s3 As String
Dim sExpediente As String, i As Integer
On Error GoTo salir:
If Len(Trim(txtExpediente)) = 0 Then
    MsgBox "Debe capturar el número de folio del SIO.", vbInformation + vbOKOnly, ""
    Exit Sub
End If
If adors.State Then adors.Close
adors.Open "select f_asuntoxfolioSIO('" & txtExpediente.Text & "') from dual", gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    MsgBox "El asunto ya existe en el SIAM. No puede agregar nuevamente el mismo FOLIO", vbOKOnly + vbInformation, "Validación"
    Exit Sub
End If
'MsgBox "pasa 1", vbOKOnly, ""
If adors.State Then adors.Close
adors.Open "{call p_modulosdatosXExp('" & txtExpediente.Text & "')}", gConSql, adOpenForwardOnly, adLockReadOnly
If Not adors.EOF Then
    If adors!idreg > 0 Then
        Call MsgBox("El expediente: " & txtExpMod.Text & " se encuentra ya asignado", vbOKOnly + vbInformation, "")
        Exit Sub
    End If
    If MsgBox("Está seguro de obtener datos del asunto con no. expediente: " & adors!expediente, vbYesNo + vbQuestion, "Confirmación") = vbNo Then
        Exit Sub
    End If
    i = 200
    sExpediente = txtExpediente.Text
    miUni = adors!iduni
    ContinuaCarga adors, i, sExpediente
End If
'MsgBox "pasa 2", vbOKOnly, ""
If cmbCampo(0).ListIndex >= 0 Then
    If MsgBox("Los datos capturados se actualizarán a la información obtenida del SIO. ¿Está seguro de continuar?", vbYesNo) = vbNo Then
        Exit Sub
    End If
End If
If Not Frame1.Enabled Then Frame1.Enabled = True
's = "FILEDSN=c:\archivos de programa\archivos comunes\odbc\data sources\dsnora.dsn;pwd=siodesa"
s = Replace(Replace(Replace(Replace(Replace(gsConexión, "siam_desa", "siodesa"), "DSN=SIAM", "DSN=SIO"), "siam.dsn", "dsnora.dsn"), "uid=siamdesa", "uid=sio"), "pwd=siamdesa", "pwd=siodesa")
's = "dsn=sio;pwd=siodesa"
Call EstableceConexiónServidor(s, conn)
ycon = 1 'estableció conexión con el SIO
If adors.State Then adors.Close
adors.Open "select f_asuntoxfolio('" & txtExpediente.Text & "') from dual", conn, adOpenStatic, adLockReadOnly
'MsgBox "pasa 3", vbOKOnly, ""
If adors(0) > 0 Then
    mlAsuSIO = adors(0)
Else
    Call MsgBox("El folio no existe", vbOKOnly + vbInformation, "")
    Exit Sub
End If
If adors.State Then adors.Close
'adors.Open "select idcla||'|'||idins, min(av.fecha) as fecha from asuntoinstitución ai, avances av where ai.idasu=" & mlAsuSIO & " and ai.id=av.idasuins and av.idtar in (select id from actividades where clase=2 and lower(descripción) like '%sanci_n%') group by idcla||'|'||idins", conn, adOpenStatic, adLockReadOnly
adors.Open "{call P_SIAM_Registro_BuscaIF(" & mlAsuSIO & ")}", conn, adOpenForwardOnly, adLockReadOnly
'MsgBox "pasa 4", vbOKOnly, ""

If Not adors.EOF Then
    s = ""
    ss = ""
    Do While Not adors.EOF
        s = s & adors(0) & ";"
        ss = ss & adors(1) & ","
        adors.MoveNext
    Loop
    If adors.State Then adors.Close
    'adors.Open "select idcla||'|'||idins, min(av.fecha) as fecha, f_asunto_idprocau(" & mlAsuSIO & "), f_asunto_productos(" & mlAsuSIO & ") as Producto, f_asunto_causa(" & mlAsuSIO & ") as causa from asuntoinstitución ai, avances av where ai.idasu=" & mlAsuSIO & " and ai.id=av.idasuins group by idcla||'|'||idins", conn, adOpenStatic, adLockReadOnly
    adors.Open "{call P_SIAM_Registro_BuscaIF2(" & mlAsuSIO & ")}", conn, adOpenForwardOnly, adLockReadOnly
    'MsgBox "pasa 5 ", vbOKOnly, ""
    s = ""
    ss = ""
    micau = 0
    miPr1 = 0
    miPr2 = 0
    miPr3 = 0
    Do While Not adors.EOF
        If micau = 0 And Not IsNull(adors(2)) Then
            s2 = adors(2)
            If InStr(s2, "|") > 0 Then 'obtiene los valores PN1, PN2, PN3, Cau de la cadena 3er campo
                miPr1 = Val(s2)
                s2 = Mid(s2, InStr(s2, "|") + 1)
                miPr2 = Val(s2)
                s2 = Mid(s2, InStr(s2, "|") + 1)
                miPr3 = Val(s2)
                s2 = Mid(s2, InStr(s2, "|") + 1)
                micau = Val(s2)
                txtProducto.Text = adors(3)
                txtCausa.Text = adors(4)
            End If
        End If
        s = s & adors(0) & ";"
        ss = ss & adors(1) & ","
        adors.MoveNext
    Loop
Else
    If MsgBox("Este asunto no tiene procedimiento de sanción en el SIO. ¿Está seguro de continuar?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirmación") = vbNo Then
        Exit Sub
    End If
    If adors.State Then adors.Close
    adors.Open "{call P_SIAM_Registro_BuscaIF2(" & mlAsuSIO & ")}", conn, adOpenForwardOnly, adLockReadOnly
    'adors.Open "select idcla||'|'||idins, min(av.fecha) as fecha, f_asunto_idprocau(" & mlAsuSIO & "), f_asunto_productos(" & mlAsuSIO & ") as Producto, f_asunto_causa(" & mlAsuSIO & ") as causa from asuntoinstitución ai, avances av where ai.idasu=" & mlAsuSIO & " and ai.id=av.idasuins group by idcla||'|'||idins", conn, adOpenStatic, adLockReadOnly
    s = ""
    ss = ""
    micau = 0
    miPr1 = 0
    miPr2 = 0
    miPr3 = 0
    Do While Not adors.EOF
        If micau = 0 And Not IsNull(adors(2)) Then
            s2 = adors(2)
            If InStr(s2, "|") > 0 Then 'obtiene los valores PN1, PN2, PN3, Cau de la cadena 3er campo
                miPr1 = Val(s2)
                s2 = Mid(s2, InStr(s2, "|") + 1)
                miPr2 = Val(s2)
                s2 = Mid(s2, InStr(s2, "|") + 1)
                miPr3 = Val(s2)
                s2 = Mid(s2, InStr(s2, "|") + 1)
                micau = Val(s2)
                txtProducto.Text = adors(3)
                txtCausa.Text = adors(4)
            End If
        End If
        s = s & adors(0) & ";"
        ss = ss & adors(1) & ","
        adors.MoveNext
    Loop
End If
If miPr1 > 0 Then
    msProSel = "000000000"
    msProSel = Right(msProSel & miPr1, 10) & Right(msProSel & miPr2, 10) & Right(msProSel & miPr3, 10) & Right(msProSel & micau, 10)
End If
'Obtiene datos del SIO y los coloca en los campos de Registro de SIAM
If adors.State Then adors.Close

's3 = "select n.iddel,ai.idcla,ai.idins,r.director,f_1erusuario(ai.idasu) as usuario,n.año,n.consecutivo,F_AsuIns_FechaRecep(ai.id) as recepción,a.idpr1,a.idpr2,a.idpr3,a.idcau from asuntoinstitución ai, nominales n, delegaciones d, regiones r, asuntos a where ai.idasu=a.id and d.idreg=r.id(+) and n.iddel=d.id and ai.idasu=n.idasu and instr('" & s & "',ai.idcla||'|'||ai.idins||';')>0 and ai.idasu=" & mlAsuSIO

'MsgBox "pasa 5: " & s3, vbOKOnly, ""

adors.Open "{call P_SIAM_Registro_BuscaIF3(" & mlAsuSIO & ",'" & s & "')}", conn, adOpenForwardOnly, adLockReadOnly
If Not adors.EOF Then
    cmbCampo(0).ListIndex = BuscaCombo(cmbCampo(0), (IIf(IsNull(adors(3)), 0, adors(3)) + 1), True)
    If cmbCampo(0).ListIndex >= 0 Then
        cmbCampo(1).ListIndex = BuscaCombo(cmbCampo(1), adors(0), True)
    End If
    cmbCampo(2).ListIndex = BuscaCombo(cmbCampo(2), "1", True)
    miAnio = IIf(IsNull(adors(5)), 0, adors(5))
    miDel = IIf(IsNull(adors(0)), 0, adors(0))
    mlCon = IIf(IsNull(adors(6)), 0, adors(6))
    txtCampo(0).Text = IIf(IsNull(adors(4)), "", adors(4))
    InhibeCampoPrinc True, 1
    miSIO_SoloPro = 0
    msProSel = Right("000000000" & adors(8), 10) & Right("000000000" & adors(9), 10) & Right("000000000" & adors(10), 10) & Right("000000000" & adors(11), 10) & "|"
    'If IsDate(Mid(ss, 1, InStr(ss, ",") - 1)) Then
    '    txtCampo(1).Text = IIf(IsNull(adors(4)), "", adors(4))
    'End If
    'txtCampo(3).Text = IIf(IsNull(adors(7)), "", Format(adors(7), gsFormatoFecha))
    If adors.RecordCount = 1 Then 'una sola IF
        cmbCampo(3).ListIndex = BuscaCombo(cmbCampo(3), adors(1), True)
        If cmbCampo(3).ListIndex >= 0 Then
            cmbCampo(4).ListIndex = BuscaCombo(cmbCampo(4), adors(2), True)
        End If
    Else 'Varias IF
        msInstituciones = ""
        msClaseIns = ""
        msClaIns = ""
        i = 0
        s = msProSel
        Do While Not adors.EOF
            msInstituciones = msInstituciones & adors(2) & ","
            msClaseIns = msClaseIns & adors(1) & ","
            msClaIns = msClaIns & Right("000000000" & adors(1), 10) & Right("000000000" & adors(2), 10) & "|"
            If i > 0 Then
                msProSel = msProSel & s
            End If
            i = i + 1
            adors.MoveNext
        Loop
        f_UnaSolaIns False
        If adors.State > 0 Then adors.Close
        adors.Open "select rci.idcla,rci.idins,ci.descripción as clase,i.descripción as institución from instituciones i, relaciónclaseinstitución rci, claseinstitución ci where i.id in (" + Mid(msInstituciones, 1, Len(msInstituciones) - 1) + ") and i.id=rci.idins and rci.idcla=ci.id", gConSql, adOpenStatic, adLockReadOnly
        ListClaseIns.Clear
        Do While Not adors.EOF
            ListClaseIns.AddItem adors!institución
            ListClaseIns.ItemData(ListClaseIns.NewIndex) = adors(1)
            adors.MoveNext
        Loop
    End If
End If

Set adors = Nothing
If conn.State Then conn.Close
Set conn = Nothing

Exit Sub
salir:
If ycon > 0 Then
    If adors.State Then adors.Close
    Set adors = Nothing
    If conn.State Then conn.Close
    Set conn = Nothing
    Call MsgBox("Error no esperado: " & Err.Description & ".", vbOKOnly + vbInformation)
    'Resume
    Exit Sub
End If
yErr = MsgBox("Error no esperado: " & Err.Description & ".", vbAbortRetryIgnore)
If yErr = vbRetry Then
    Resume
ElseIf yErr = vbIgnore Then
    Resume Next
End If
End Sub

Private Sub cmdIrAna_Click()
Dim frm As Form
If mlAsunto > 0 Then
    gs = ">>"
    If Len(Trim(txtExpediente.Text)) > 0 Then
        gs1 = Trim(txtExpediente.Text)
    Else
        gs1 = Trim(txtNuevoExp.Text)
    End If
    If cmbCampo(4).Visible And cmbCampo(4).ListIndex >= 0 Then
        gi1 = cmbCampo(4).ItemData(cmbCampo(4).ListIndex)
    End If
    Set frm = Análisis
    With frm
        .Show
    End With
End If

End Sub

Private Sub cmdModulos_Click()
'Obtiene consulta para obtener información de asuntos turnados desde módulos SIAM
Dim s As String, i As Integer
Dim adors As New ADODB.Recordset
Dim adors1 As New ADODB.Recordset
Dim sInsCla As String
Dim iIns As Integer
Dim sExpediente As String
If Len(Trim(txtExpMod.Text)) > 0 Then
    adors.Open "{call p_modulosdatosXExp('" & txtExpMod.Text & "')}", gConSql, adOpenForwardOnly, adLockReadOnly
    If adors.EOF Then
        Call MsgBox("No se encuentra información con el expediente: " & txtExpMod.Text, vbOKOnly + vbInformation, "")
        Exit Sub
    Else
        If adors!idreg > 0 Then
            Call MsgBox("El expediente: " & txtExpMod.Text & " se encuentra ya asignado", vbOKOnly + vbInformation, "")
            Exit Sub
        End If
        If MsgBox("Está seguro de obtener datos del asunto con no. expediente: " & adors!expediente, vbYesNo + vbQuestion, "Confirmación") = vbNo Then
            Exit Sub
        End If
    End If
    i = 200
    sExpediente = txtExpMod.Text
    miUni = adors!iduni
Else
    If opcAreOri(0).Value Then
        s = "0,"
    ElseIf opcAreOri(1).Value Then
        s = "3,"
    ElseIf opcAreOri(2).Value Then
        s = "2,"
    ElseIf opcAreOri(3).Value Then
        s = "1,"
    End If
    For i = 0 To 1
        If cmbCampo(i).ListIndex >= 0 Then
            s = s & cmbCampo(i).ItemData(cmbCampo(i).ListIndex) & ","
        Else
            s = s & "0,"
        End If
    Next
    s = Mid(s, 1, Len(s) - 1)
    gs1 = 4
    gs = "{call P_Modulos_TurNoIni1(" & s & ")}"
    gs4 = 200 'Para que no se habra el arbol
    SelProceso.Caption = "Seleccione un Expediente"
    SelProceso.Show vbModal
    If Val(gs) > 0 Then
        miUni = Int(Val(gs) / 1000000)
        If miUni > 0 Then
            gs = Val(gs) Mod 1000000
            adors.Open "{call p_modulosdatos(" & Val(gs) & "," & miUni & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
            If Not adors.EOF Then
                sExpediente = adors!expediente
                If MsgBox("Está seguro de obtener datos del asunto con no. expediente: " & adors!expediente, vbYesNo + vbQuestion, "Confirmación") = vbNo Then
                    Exit Sub
                End If
                i = 200
            End If
        End If
    End If
End If
ContinuaCarga adors, i, sExpediente
End Sub

Sub ContinuaCarga(ByRef adors As ADODB.Recordset, ByRef i As Integer, ByRef sExpediente As String)
Dim adors1 As New ADODB.Recordset
Dim sInsCla As String
Dim iIns As Integer
If cmdAgregarIF.Enabled Then cmdAgregarIF.Enabled = False
If i = 200 Then
    mlModulo = adors!idmod
    'Busca si hay asociados mas de un producto impedir la aceptación hasta que de origen del MSS dejen un solo producto
    If InStr(csUniDGEV, miUni & "|") > 0 Then 'Se trata de DGEV
        If adors1.State Then adors1.Close
        adors1.Open "select mssdgev.f_Sol_ProductosxExp('" & sExpediente & "'), instr(mssdgev.f_Sol_ProductosxExp('" & sExpediente & "'),'|') as pipe1,instr(mssdgev.f_Sol_ProductosxExp('" & sExpediente & "'),'|',-1) as pipe2 from dual", gConSql, adOpenStatic, adLockReadOnly
        If adors1(1) <> adors1(2) Then
            MsgBox "El asunto contiene más de un producto asociado por lo que se requiere se corriga desde origen para poder recibirlo", vbOKOnly + vbInformation, "Inconsistencia en el asunto"
                        Dim Botón As Object
            Set Botón = Me.Toolbar.Buttons(4)
            gi = 199 'indica que no debe preguntar nuevamente
            Call Toolbar_ButtonClick(Botón)
            Exit Sub
        End If
        msProSel = IIf(IsNull(adors1(0)), 0, adors1(0))
        mbRefrescaDatos = True
        If Len(msProSel) > 35 Then 'Se trata del Producto del SIO
            opcSIO(1).Value = True
            If adors1.State Then adors1.Close
            adors1.Open "select paq_conceptos.sio_PN1(" & Mid(msProSel, 1, 10) & ") as PN1, paq_conceptos.sio_PN2(" & Mid(msProSel, 11, 10) & ") as PN2,paq_conceptos.sio_PN3(" & Mid(msProSel, 21, 10) & ") as PN3 from dual", gConSql, adOpenStatic, adLockReadOnly
            txtProducto.Text = adors1(0) & " / " & adors1(1) & " / " & adors1(2)
            miPr1 = Val(Mid(msProSel, 1, 10))
            miPr2 = Val(Mid(msProSel, 11, 10))
            miPr3 = Val(Mid(msProSel, 21, 10))
            micau = 0
            miSIO_SoloPro = 2
            txtCausa.Text = "No Aplica"
        ElseIf Len(msProSel) >= 30 Then ' Se trata de Producto de RECA
            opcSIO(2).Value = True
            If adors1.State Then adors1.Close
            adors1.Open "select paq_conceptos.reca_operacion(" & Mid(msProSel, 1, 10) & ") as Operacion, paq_conceptos.reca_Producto(" & Mid(msProSel, 11, 10) & ") as Producto,paq_conceptos.reca_Subproducto(" & Mid(msProSel, 21, 10) & ") as Subproducto from dual", gConSql, adOpenStatic, adLockReadOnly
            txtProducto.Text = adors1(0) & " / " & adors1(1)
            miPr1 = Val(Mid(msProSel, 1, 10))
            miPr2 = Val(Mid(msProSel, 11, 10))
            miPr3 = Val(Mid(msProSel, 21, 10))
            txtCausa.Text = adors1(2)
        End If
        opcVig(1).Value = True
        opcVig(0).Enabled = False
        opcVig(1).Enabled = False
        opcSIO(0).Enabled = False
        opcSIO(1).Enabled = False
        opcSIO(2).Enabled = False
        chkSinCau.Value = 1
        chkSinCau.Enabled = True
        mbRefrescaDatos = False
    ElseIf miUni <= 450 Then 'MSS_SIO
        mbRefrescaDatos = True
        opcSIO(1).Value = True
        If adors1.State Then adors1.Close
        adors1.Open "{call P_idExp_ProCau(" & mlModulo & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
        If Not adors1.EOF Then
            miPr1 = adors1(0)
            miPr2 = adors1(1)
            miPr3 = adors1(2)
            micau = adors1(3)
            txtProducto.Text = adors1(4)
            msProSel = adors1(6)
            txtCausa.Text = adors1(5)
        Else
            miPr1 = 0
            miPr2 = 0
            miPr3 = 0
            micau = 0
            txtProducto.Text = ""
            msProSel = ""
            txtCausa.Text = ""
        End If
        txtCampo(0).Text = IIf(IsNull(adors(12)), "", adors(12))
        miSIO_SoloPro = 0
        opcVig(1).Value = True
        opcVig(0).Enabled = False
        opcVig(1).Enabled = False
        opcSIO(0).Enabled = False
        opcSIO(1).Enabled = False
        opcSIO(2).Enabled = False
        chkSinCau.Value = 0
        chkSinCau.Enabled = True
        mbRefrescaDatos = False
    End If
    cmbCampo(0).ListIndex = BuscaCombo(cmbCampo(0), IIf(IsNull(adors!iddir), 0, adors!iddir), True)
    If cmbCampo(0).ListIndex >= 0 Then
        cmbCampo(0).Locked = True
    End If
    cmbCampo(1).ListIndex = BuscaCombo(cmbCampo(1), IIf(IsNull(adors!iduni), 0, adors!iduni), True)
    If cmbCampo(1).ListIndex >= 0 Then
        cmbCampo(1).Locked = True
    End If
    cmbCampo(2).ListIndex = BuscaCombo(cmbCampo(2), IIf(IsNull(adors!idmat), 0, adors!idmat), True)
    If cmbCampo(2).ListIndex >= 0 Then
        cmbCampo(2).Locked = True
    End If
    msFolio = adors!expediente
    
    lblExpMod.Caption = "Expediente a generar: " & msFolio
    lblTipoExp.Caption = "Tipo de Expediente: " & IIf(IsNull(adors!TipoExp), 0, adors!TipoExp)
    txtResObs.Text = IIf(IsNull(adors!Responsable_obs), "", adors!Responsable_obs)
    txtResObs.Visible = True
    lblResObs.Visible = True
    txtObs.Width = 4440

    txtCampo(1).Text = IIf(IsNull(adors!no_memorando), "", adors!no_memorando)
    txtCampo(1).Locked = True
    txtCampo(2).Text = IIf(IsNull(adors!f_memorando), "", Format(adors!f_memorando, "dd/mm/yyyy"))
    txtCampo(2).Locked = True
    
    sInsCla = IIf(IsNull(adors!clasesif), "", adors!clasesif)
    If InStr(sInsCla, ",") = 0 Then
        msInstituciones = ""
        msClaseIns = ""
        If InStr(sInsCla, "|") > 0 Then
            f_UnaSolaIns (True)
            cmbCampo(3).ListIndex = BuscaCombo(cmbCampo(3), Val(Mid(sInsCla, InStr(sInsCla, "|") + 1, 10)), True)
            If cmbCampo(3).ListIndex >= 0 Then
                cmbCampo(3).Locked = True
            End If
            cmbCampo(4).ListIndex = BuscaCombo(cmbCampo(4), Val(sInsCla), True)
            iIns = Val(sInsCla)
            If cmbCampo(4).ListIndex >= 0 Then
                cmbCampo(4).Locked = True
            End If
        End If
    Else
        f_UnaSolaIns (False)
        msInstituciones = Mid(sInsCla, 1, InStr(sInsCla, "|") - 1) & ","
        msClaseIns = Mid(sInsCla, InStr(sInsCla, "|") + 1, 50) & ","
        If adors.State Then adors.Close
        adors.Open "select f_idsecvm_idcla(f_idins_idsec(id)) as idcla, id as idins, f_clase(f_idsecvm_idcla(f_idins_idsec(id))) as clase, descripción as if from instituciones where id in (" & Mid(msInstituciones, 1, Len(msInstituciones) - 1) & ")", gConSql, adOpenStatic, adLockReadOnly
        ListClaseIns.Clear
        If Not adors.EOF Then
            If adors.RecordCount > 1 Then 'Varias Instituciones
                Do While Not adors.EOF
                    ListClaseIns.AddItem adors(3) & " (" & adors(2) & ")"
                    ListClaseIns.ItemData(ListClaseIns.NewIndex) = adors(1)
                    adors.MoveNext
                Loop
            End If
        End If
    End If
    
    'Call MsgBox("El asunto fue turnado desde otra área por medio de MSS. favor de analizar, aceptar o rechazar las causas", vbInformation + vbOKOnly, "")
    gs = sExpediente
    gi1 = mlModulo
    gs2 = txtCampo(1).Text & "  (" & txtCampo(2).Text & ")"
    gs3 = txtResObs.Text
    gs4 = miUni
    If cmbCampo(4).ListIndex >= 0 Then
        gi = cmbCampo(4).ItemData(cmbCampo(4).ListIndex)
    Else
        gi = iIns
    End If
    AceptaCausas.Show vbModal
    If gs = "ok" Then 'Fue Aceptada
        miModPro = 1
        If cmbCampo(5).ListIndex > 0 Then
             If cmbCampo(5).ItemData(cmbCampo(5).ListIndex) = 0 Then
                cmbCampo(5).ListIndex = -1
             End If
        End If
        If cmbCampo(5).Locked Then cmbCampo(5).Locked = False
        'SSTab1.TabEnabled(1) = False
        'SSTab1.Tab = 0
    Else 'Fue rechazada
        miModPro = 2
        'asigna a usuario esp para regresar al área
        cmbCampo(5).ListIndex = BuscaCombo(cmbCampo(5), 999, True)
        If cmbCampo(5).ListIndex >= 0 Then
            cmbCampo(5).Locked = True
        End If
        'SSTab1.TabEnabled(0) = False
        'SSTab1.Tab = 1
    End If
End If
End Sub


Private Sub cmdPro_Click()
Dim adors As New ADODB.Recordset, iCla As Integer, iIns As Long
Dim sPro As String
If cmbCampo(3).Visible Then 'Se tiene una if
    If cmbCampo(3).ListIndex < 0 Then
        MsgBox "debe elegir el sector antes de seleccionar el producto", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
    If opcSIO(1).Value Then
        gs = "{call P_Productos(" & IIf(miSIO_SoloPro > 0, 2, IIf(opcSIO(1).Value, 1, 0)) & "," & cmbCampo(3).ItemData(cmbCampo(3).ListIndex) & "," & IIf(opcVig(0).Value = 1, 1, 0) & ")}"
        If miSIO_SoloPro > 0 Then
            gs1 = miPr3
        Else
            gs1 = micau
        End If
    Else
        gs = "{call P_Productos(0," & cmbCampo(3).ItemData(cmbCampo(3).ListIndex) & "," & IIf(opcVig(0).Value = 1, 1, 0) & ")}"
        gs1 = miPr3
    End If
Else 'Considera varias if
    If ListClaseIns.ListCount <= 0 Then
        MsgBox "debe elegir las Instituciones involucradas antes de seleccionar el producto", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
    If ListClaseIns.ListIndex < 0 Then
        MsgBox "Debe elegir las IF asociada al producto", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
    If opcSIO(1).Value Then
        'iCla = 1 'f_obtenclains(msSel, 1 + ListClaseIns.ListIndex)
        gs = "{call P_Productos(" & IIf(miSIO_SoloPro > 0, 2, IIf(opcSIO(1).Value, 1, 0)) & "," & miCla & "," & IIf(opcVig(0).Value = 1, 1, 0) & ")}"
        If miSIO_SoloPro > 0 Then
            gs1 = miPr3
        Else
            gs1 = micau
        End If
    Else
        gs = "{call P_Productos(" & IIf(opcSIO(1).Value, 1, 0) & "," & miCla & "," & IIf(opcVig(0).Value = 1, 1, 0) & ")}"
        gs1 = miPr3
    End If
End If
'gs4 = 200 'Para que no se habra el arbol
SelProceso.Caption = "Seleccione el Producto" & IIf(opcSIO(1).Value, " y causa", "")
SelProceso.piValida = 1
If opcSIO(1).Value Then
    If miSIO_SoloPro > 0 Then
        If miPr1 > 0 And miPr2 > 0 And miPr3 > 0 Then
            gs2 = "-->" & Right("000000000" & miPr1, 10) & Right("000000000" & miPr2, 10) & Right("000000000" & miPr3, 10)
        Else
            gs2 = ""
        End If
    Else
        If miPr1 > 0 And miPr2 > 0 And miPr3 > 0 And micau > 0 Then
            gs2 = "-->" & Right("000000000" & miPr1, 10) & Right("000000000" & miPr2, 10) & Right("000000000" & miPr3, 10) & Right("000000000" & micau, 10)
        Else
            gs2 = ""
        End If
    End If
Else
    If miPr1 > 0 And miPr2 > 0 And miPr3 > 0 Then
        gs2 = "-->" & Right("000000000" & miPr1, 10) & Right("000000000" & miPr2, 10) & Right("000000000" & miPr3, 10)
    Else
        gs2 = ""
    End If
End If
SelProceso.Show vbModal
If Val(gs) > 0 Then
    sPro = gs1
    If miSIO_SoloPro > 0 And Len(sPro) < 40 Then
        sPro = sPro & "0000000000"
    End If
    AsignaNvoPro sPro
End If

End Sub

Sub AsignaNvoPro(sPro)
Dim s As String, i As Integer, ss As String, sSel As String
If cmbCampo(3).Visible Then 'Una sola IF
    msProSel = sPro
Else
    If ListClaseIns.ListIndex < 0 Then
        Call MsgBox("No es posible asignar el Producto seleccionado ya que no hay IF seleccionada", vbInformation + vbOKOnly, "")
        Exit Sub
    End If
    s = msProSel
    i = 1
    Do While i <= ListClaseIns.ListIndex
        If InStr(s, "|") > 0 Then
            ss = Mid(s, 1, InStr(s, "|"))
            sSel = sSel & ss
            s = Mid(s, InStr(s, "|") + 1)
        Else
            sSel = sSel & "|"
        End If
        i = i + 1
    Loop
    sSel = sSel & sPro & "|"
    If InStr(s, "|") Then
        s = Mid(s, InStr(s, "|") + 1)
        sSel = sSel & s
    Else
        If ListClaseIns.ListCount - i > 0 Then
            sSel = sSel & Mid("|||||||||||||||||||||", 1, ListClaseIns.ListCount - i)
        End If
    End If
    msProSel = sSel
End If
ActualizaDatosPro (sPro)
End Sub

Private Sub ActualizaDatosPro(sPro As String)
Dim i As Integer, adors As New ADODB.Recordset
On Error GoTo salir:
miProCambio = 1
If Len(sPro) >= 40 Or miSIO_SoloPro > 0 Then 'Del SIO
    opcSIO(1).Value = True
    miPr1 = Val(Mid(sPro, 1, 10))
    miPr2 = Val(Mid(sPro, 11, 10))
    miPr3 = Val(Mid(sPro, 21, 10))
    If miSIO_SoloPro <= 0 Then
        micau = Val(Mid(sPro, 31, 10))
    End If
ElseIf Len(sPro) >= 30 Then 'Del RECA
    opcSIO(2).Value = True
    miPr1 = Val(Mid(sPro, 1, 10))
    miPr2 = Val(Mid(sPro, 11, 10))
    miPr3 = Val(Mid(sPro, 21, 10))
Else
    opcSIO(0).Value = True
    miPr1 = 0
    miPr2 = 0
    miPr3 = 0
    micau = 0
End If
If adors.State Then adors.Close
adors.Open "select f_reg_prodnivel(" & IIf(opcSIO(1).Value, 1, 0) & ",1," & IIf(opcSIO(1).Value, miPr3, miPr2) & "),f_reg_prodnivel(" & IIf(opcSIO(1).Value, 1, 0) & ",2," & IIf(opcSIO(1).Value, micau, miPr3) & ") from dual", gConSql, adOpenStatic, adLockReadOnly
txtProducto.Text = adors(0)
If miSIO_SoloPro <= 0 Then
    txtCausa.Text = adors(1)
Else
    txtCausa.Text = "No Aplica"
End If
If opcSIO(1).Value Then
    etiTexto(5).Caption = "Causa"
Else
    etiTexto(5).Caption = "Subproducto"
End If
miProCambio = 0
Exit Sub
salir:
miProCambio = 0
End Sub

Private Function f_ObtenProductos(sPro As String, iPos As Integer) As String
Dim s As String, i As Integer
i = 1
s = sPro
Do While i < iPos
    s = Mid(s, InStr(s, "|") + 1)
    i = i + 1
Loop
If InStr(s, "|") > 0 Then
    f_ObtenProductos = Mid(s, 1, InStr(s, "|") - 1)
End If
End Function


Private Sub Command2_Click()
Dim Botón As Object
Set Botón = Me.Toolbar.Buttons(4)
Call Toolbar_ButtonClick(Botón)
End Sub

Private Sub Form_Activate()
Dim s As String, i As Integer
Dim Botón As Object
s = "persona-->"
If gs Like s & "*" Then
    gs = LCase(gs)
    If Val(Mid(gs, InStr(gs, s) + Len(s))) > 0 Then
        mlBuscaPer = Val(Mid(gs, InStr(gs, s) + Len(s)))
       
        Set Botón = Me.Toolbar.Buttons(6)
        Call Toolbar_ButtonClick(Botón)
        
    End If
ElseIf gs = "<<" Then
    If Len(gs1) > 0 Then
        Set Botón = Me.Toolbar.Buttons(2)
        Call Toolbar_ButtonClick(Botón)
        If Len(Trim(txtNuevoExp.Text)) = 0 And Len(Trim(txtExpediente.Text)) = 0 Then
            txtNuevoExp.Text = gs1
            Set Botón = Me.Toolbar.Buttons(3)
            Call Toolbar_ButtonClick(Botón)
        End If
    End If
End If
End Sub

'Inicio del formulario al cargarse
Private Sub Form_Load()
Dim sCampos As String, i As Integer, sOrden As String
myPermiso = Val(Mid(gsPermisos, 4, 1))
myPermisoRep = Val(Mid(gsPermisosRep, 4, 1))
myTabla = 4
Set AdorsPrin = New ADODB.Recordset
Set adorsCata = New ADODB.Recordset
For i = txtCampo.LBound To txtCampo.UBound
    sCampos = sCampos & txtCampo(i).DataField & ","
    'If i = 0 Then msPrin = txtCampo(i).DataField
    If i = 1 Or i = 2 Or i = 3 Then sOrden = sOrden & txtCampo(i).DataField & ","
Next
sOrden = Mid(sOrden, 1, Len(sOrden) - 1)
For i = cmbCampo.LBound To cmbCampo.UBound
    If Len(cmbCampo(i).DataField) > 0 Then
        sCampos = sCampos & cmbCampo(i).DataField & ","
    End If
Next
'otros campos tipo opción y casilla
msConsulta = "select " & Mid(sCampos, 1, Len(sCampos) - 1) & " from registro"
msOrden = " order by año,consecutivo"
adorsCata.Open msConsulta & msOrden, gConSql, adOpenStatic, adLockReadOnly
msConsultaP = "select id from registro where rownum<2"
LlenaCombo cmbCampo(0), "select id,descripción from direccióngeneral where baja=0 and id<=2 order by 2", "", True
'LlenaCombo cmbCampo(1), "select id,descripción from unidades order by 2", "", True
LlenaCombo cmbCampo(2), "select id,descripción from materiasanción where baja=0 order by 2", "", True
LlenaCombo cmbCampo(3), "select id,descripción from claseinstitución where baja=0 order by 2", "", True
'LlenaCombo cmbCampo(4), "select id,descripción from direccióngeneral where baja=0 order by 2", "", True
LlenaCombo cmbCampo(5), "select id,descripción from usuariossistema where baja=0 and responsable<>0 order by 2", "", True
AdorsPrin.Open msConsultaP, gConSql, adOpenStatic, adLockReadOnly
For i = txtCampo.LBound + 1 To txtCampo.UBound
    Debug.Print adorsCata.Fields(i).Name
    For j = 0 To txtCampo.UBound
        If LCase(txtCampo(j).DataField) = LCase(adorsCata.Fields(i).Name) Then
            txtCampo(j).MaxLength = IIf(adorsCata.Fields(i).DefinedSize > 500, 500, adorsCata.Fields(i).DefinedSize)
            Exit For
        End If
    Next
Next
'RefrescaDatos
mbInicio = True
Dim Botón As Object
Set Botón = Me.Toolbar.Buttons(1)
Call Toolbar_ButtonClick(Botón)
'Call ActualizaBotones(Me, 2, myPermiso)
mbCambio = False
mbLimpia = False

End Sub

'Antes de cerrar el formulario valida lo pendiente
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If mbLimpia Or mbCambio Then
    If MsgBox("¿Desea salir del módulo de " & Me.Caption & ", sin guardarlos cambios?", vbYesNo + vbQuestion) = vbNo Then
        Cancel = 1
    End If
End If
gs = ""
End Sub

Private Sub ListClaseIns_Click()
Dim s As String, i As Integer
Dim sPro As String
On Error GoTo salir:
s = msClaIns
If ListClaseIns.ListIndex >= 0 Then
    Do While i < ListClaseIns.ListIndex
        s = Mid(s, InStr(s, "|") + 1)
        i = i + 1
    Loop
    s = Mid(s, 1, InStr(s, "|") - 1)
    If Len(s) >= 20 Then
        miCla = Val(Mid(s, 1, 10))
        miIns = Val(Mid(s, 11, 10))
    Else
        miCla = 0
        miIns = 0
    End If
End If
sPro = f_ObtenProductos(msProSel, i + 1)
'Actualiza Productos segun la IF
miProCambio = 1
ActualizaDatosPro sPro
miProCambio = 0
Exit Sub
salir:
miProCambio = 0
End Sub



Private Sub opcIF_Click(Index As Integer)
ActualizaComboIF
End Sub

Private Sub OpcPAU_Click(Index As Integer)
cmdContinuar.Enabled = OpcPAU(0).Value
If Index = 0 Then
    'txtExpediente.Locked = False
    'txtNuevoExp.Locked = True
    txtExpediente.Enabled = True
    txtNuevoExp.Enabled = False
    txtNuevoExp.Visible = Not mbNuevo
    Label2.Visible = Not mbNuevo
    LlenaCombo cmbCampo(0), "select id,descripción from direccióngeneral where baja=0 and id in (2,3) order by 2", "", True
    Frame1.Enabled = False
    If mbNuevo Then
        opcSIO(1).Value = True
    End If
Else
    txtNuevoExp.Visible = Not mbNuevo
    Label2.Visible = Not mbNuevo
        
    txtExpediente.Enabled = False
    txtNuevoExp.Enabled = True
    LlenaCombo cmbCampo(0), "select id,descripción from direccióngeneral where baja=0 order by 2", "", True
    Frame1.Enabled = True
    If mbNuevo Then
        opcSIO(0).Value = True
    End If
        
'    txtExpediente.Locked = True
'    txtNuevoExp.Locked = False
End If
End Sub


Private Sub opcSIO_Click(Index As Integer)
If mbLimpia Then Exit Sub
cmdPro.Visible = (Not opcSIO(0).Value)
opcVig(0).Visible = (Not opcSIO(0).Value)
opcVig(1).Visible = (Not opcSIO(0).Value)
chkSinCau.Visible = (opcSIO(1).Value)
If miProCambio > 0 Then
    Exit Sub
End If
If miPr1 > 0 And Not mbRefrescaDatos Then
    If MsgBox("Tiene producto seleccionado, está seguro de cambiar los datos del producto", vbYesNo + vbInformation, "") = vbNo Then
        ActOpcionIF
        Exit Sub
    End If
    AsignaNvoPro ("")
End If
End Sub

Private Sub ActOpcionIF()
Dim sPro As String
miProCambio = 1
On Error GoTo salir:
If cmbCampo(3).Visible Then 'una sola IF
    If Len(msProSel) >= 40 Then
        opcSIO(1).Value = True
    ElseIf Len(msProSel) >= 30 Then
        opcSIO(2).Value = True
    Else
        opcSIO(0).Value = True
    End If
Else
    If ListClaseIns.ListIndex >= 0 Then
        sPro = f_ObtenProductos(msProSel, ListClaseIns.ListIndex + 1)
    Else
        sPro = ""
    End If
    ActualizaDatosPro sPro
End If
miProCambio = 0
Exit Sub
salir:
miProCambio = 0
End Sub

Private Sub opcVig_Click(Index As Integer)
ActualizaComboIF
End Sub

'Acciones de los botones de la barra de herramientas
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim iIr_a As Long, i As Integer, sCondición As String, yError As Integer, l As Long
Dim adors As New ADODB.Recordset, s As String, s1 As String, s2 As String
Dim iAnio As Integer, iCon As Long
Dim sVal As String, sCam As String
Dim ii As Integer, iRows As Integer
Dim bBeginTrans As Boolean 'indicador si inición begin transacction
Dim adorsBloqueo As ADODB.Recordset, bBloqueo As Boolean
Dim sFolio As String, sCam2 As String
On Error GoTo ErrorAcción:

Select Case Button.Key
Case "Primero", "Anterior", "Siguiente", "Último", "Ir_a"
    If mbCambio Then
        
        If MsgBox("¿Desea guardar los cambios realizados?", vbYesNo + vbQuestion, "") = vbYes Then
            Dim Botón As Object
            Set Botón = Me.Toolbar.Buttons(4)
            gi = 199 'indica que no debe preguntar nuevamente
            Call Toolbar_ButtonClick(Botón)
        Else
            RefrescaDatos
        End If
    End If
    MoverCursor Me, Button.Key, AdorsPrin, Val(txtNoReg.Text)
    mbLimpia = False
    mbCambio = False
    If Val(txtNoReg.Text) > 0 Then
        If txtNoReg.Locked Then txtNoReg.Locked = False
    End If
Case "Guardar"
    If mlAsunto <= 0 Then 'Nuevo
        mbNuevo = True
    Else
        mbNuevo = False
        If Len(Trim(txtExpediente.Text)) = 0 And Len(Trim(txtNuevoExp.Text)) = 0 Then
            MsgBox "Se requiere el número de expediente. Favor de capturar", vbOKOnly + vbInformation, ""
            Exit Sub
        End If
        If Len(Trim(txtExpediente.Text)) = 0 Then
            s = txtNuevoExp.Text
        Else
            s = txtExpediente.Text
        End If
        If adors.State Then adors.Close
        adors.Open "select count(*) from registro where id<>" & mlAsunto & " and expediente='" & Replace(s, "'", "''") & "'", gConSql, adOpenStatic, adLockReadOnly
        If adors(0) > 0 Then
            MsgBox "El número de expediente: " & s & " ya existe. Verifique la información capturada", vbInformation + vbOKOnly, ""
            Exit Sub
        End If
    End If
        'If Len(Trim(txtExpediente.Text)) = 0 And Len(Trim(txtNuevoExp.Text)) = 0 Then
        '    MsgBox "Se requiere el número de expediente. Favor de capturar", vbOKOnly + vbInformation, ""
        '    Exit Sub
        'End If
        If OpcPAU(0).Value Then
            sFolio = UCase(txtExpediente.Text)
            If Len(sFolio) = 0 Then
                MsgBox "El número de expediente no puede estar vacio. Capture el número de expediente", vbInformation + vbOKOnly, ""
                Exit Sub
            End If
            If adors.State Then adors.Close
            adors.Open "select count(*) from registro where id<>" & mlAsunto & " and expediente='" & Replace(sFolio, "'", "''") & "'", gConSql, adOpenStatic, adLockReadOnly
            If adors(0) > 0 Then
                MsgBox "El número de expediente: " & sFolio & " ya existe. Verifique la información capturada", vbInformation + vbOKOnly, ""
                Exit Sub
            End If
        End If
        For i = 0 To txtCampo.UBound
            If Len(Trim(txtCampo(i).Text)) = 0 Then
                If i = 0 Then
                    txtCampo(0) = "--" 'En caso de nombre sin dato asigna --
                Else
                    Exit For
                End If
            End If
        Next
        If i <= txtCampo.UBound Then
            MsgBox "Falta dato requerido: " & txtCampo(i).DataField, vbInformation + vbOKOnly, "validación"
            Exit Sub
        End If
        For i = 1 To cmbCampo.UBound
            If cmbCampo(i).ListIndex < 0 And cmbCampo(i).Visible Then Exit For
        Next
        If i <= cmbCampo.UBound Then
            MsgBox "Falta dato requerido: " & etiCombo(i).Caption, vbInformation + vbOKOnly, "validación"
            Exit Sub
        End If
            
        'txtCampo(0).Text = mlAsunto
'        For i = 0 To Controls.Count - 1
'            If LCase(Controls(i).Name) = "txtcampo" Then
'                If Len(Controls(i).DataField) > 0 Then
'                    s = ArmaCadenaCampo(Controls(i).DataField, Controls(i).Text, Controls(i).Tag, 2)
'                    sVal = sVal & s
'                    sCam = sCam & Controls(i).DataField & ","
'                End If
'            ElseIf Mid(Controls(i).Name, 1, 3) = "cmb" Then
'                If Len(Controls(i).DataField) > 0 Then
'                    If Controls(i).ListIndex >= 0 Then
'                        s = Controls(i).ItemData(Controls(i).ListIndex)
'                    Else
'                        s = ""
'                    End If
'                    s = ArmaCadenaCampo(Controls(i).DataField, s, "n", 2)
'                    sVal = sVal & s
'                    sCam = sCam & Controls(i).DataField & ","
'                End If
'            ElseIf Mid(Controls(i).Name, 1, 3) = "chk" Then
'                If Len(Controls(i).DataField) > 0 Then
'                    sVal = sVal & IIf(Controls(i).Value = 1, 1, 0) & ","
'                    sCam = sCam & Controls(i).DataField & ","
'                End If
'            End If
'        Next
        
        'Verifica cada uno de los datos requeridos
        'Dirección General
        sCam = cmbCampo(0).ItemData(cmbCampo(0).ListIndex) & ","
        'Unidad de Origen
        sCam = sCam & cmbCampo(1).ItemData(cmbCampo(1).ListIndex) & ","
        'Materia de la Sanción
        sCam = sCam & cmbCampo(2).ItemData(cmbCampo(2).ListIndex) & ",'" & Replace(txtCampo(0).Text, "'", "''") & "'"
        If cmbCampo(3).Visible Then 'una sola IF
            msClaIns = Right("000000000" & cmbCampo(3).ItemData(cmbCampo(3).ListIndex), 10) & Right("000000000" & cmbCampo(4).ItemData(cmbCampo(4).ListIndex), 10)
        Else 'Varias if
            If InStr(msClaIns, "|") = 0 Then
                Call MsgBox("Se requiere información del Sector e Institución Financiera", vbOKOnly + vbInformation, "Validación")
                Exit Sub
            End If
        End If
        'Memorando
        sCam2 = "'" & Replace(txtCampo(1).Text, "'", "''") & "','"
        'Fecha del Memo
        sCam2 = sCam2 & Format(txtCampo(2).Text, gsFormatoFecha) & "','"
        'Fecha de recepción
        sCam2 = sCam2 & Format(txtCampo(3).Text, gsFormatoFecha) & "',"
        'Responsable quien se le turna, Obs, usuario, año, consecutivo, expediente, PAU
        sCam2 = sCam2 & cmbCampo(5).ItemData(cmbCampo(5).ListIndex) & ",'" & Replace(Trim(txtObs.Text), "'", "''") & "'"
        
        If mbNuevo Then
            If OpcPAU(1).Value And mlModulo = 0 Then 'no viene del SIO
                i = cmbCampo(2).ItemData(cmbCampo(2).ListIndex)
                If adors.State Then adors.Close
                'Call MsgBox("1:", vbOKOnly, "")
                adors.Open "select f_nuevofolio(1," & i & ",0) from dual", gConSql, adOpenStatic, adLockReadOnly
                'Call MsgBox("2:", vbOKOnly, "")
                If IsNull(adors(0)) Then
                    Call MsgBox("No se pudo generar el folio.", vbOKOnly + vbQuestion, "")
                    Exit Sub
                End If
                sFolio = adors(0)
    '            If InStr(sFolio, "???") Then
    '                i = F_PreguntaConsecutivo(1, sFolio)
    '                If i > 0 Then
    '                    sFolio = Replace(sFolio, "???", i)
    '                Else
    '                    Exit Sub
    '                End If
    '            End If
                    
                    sFolio = F_PreguntaFolio(1, sFolio)
                    If sFolio = "salir" Then
                        Exit Sub
                    End If
                
                iAnio = Val(sFolio)
                iCon = Val(Mid(sFolio, InStrRev(sFolio, "/") + 1))
                
            Else 'Viene del SIO
                'con el propósito de tener más de un asunto del SIO del mismo año de diferente delegación
                'El año estará conformado por (año-2000)*1000+iddel
                If mlModulo > 0 Then 'Caso de Módulos
                    iAnio = 0
                    iCon = 0
                    sFolio = msFolio
                    
                    ''miModPro = 0
                    
                    'Verifica si el asunto fue turnado desde módulos
                    
                    'If adors.State Then adors.Close
                    'adors.Open "select F_ana_verifexp_Mod(f_expediente_idreg('" & txtNuevoExp.Text & "')) from dual", gConSql, adOpenStatic, adLockReadOnly
                    'If adors(0) > 0 Then 'Si fue turnado desde módulos y por tanto realiza la disyuntiva Acepta/Rechaza o bloquea la pestaña según haya sido aceptado o rechazado
                        'Debe señalar si es Aceptado o rechazado de acuerdo a las causas enviadas
                        'If adors.State Then adors.Close
                        'adors.Open "select idmod,procede from registromodulos where idreg=f_expediente_idreg('" & txtNuevoExp.Text & "')", gConSql, adOpenStatic, adLockReadOnly
                        'miModulo = adors(0)
                        'If adors(1) = 1 Then 'Es procedente
                        '    miModPro = 1
                        '    SSTab1.TabEnabled(1) = False
                        '    SSTab1.Tab = 0
                        'ElseIf adors(1) = 2 Then 'Es improcedente
                        '    miModPro = 2
                        '    SSTab1.TabEnabled(0) = False
                        '    SSTab1.Tab = 1
                        'Else 'No se ha definido Si procede o no
                            'Obtiene las causa desde módulos para mostrarlas y decidan si proceden o no
                        
                        
                        
                        'End If
                    'End If
                    If miModPro = 2 And Len(Trim(txtObs.Text)) < 5 Then
                        MsgBox "Debe explicar la razón del rechazo en el campo de Observaciones", vbOKOnly + vbInformation, "Falta Campo Requerido"
                        Exit Sub
                    End If
                Else
                    iAnio = (Val(sFolio) - 2000) * 1000 + Val(Mid(sFolio, InStr(sFolio, "/") + 1))
                    iCon = Val(Mid(sFolio, InStrRev(sFolio, "/") + 1))
                End If
            End If
        Else
            If Len(Trim(txtNuevoExp.Text)) <= 1 Then
                sFolio = txtExpediente.Text
            Else
                sFolio = txtNuevoExp.Text
            End If
            If OpcPAU(1).Value And mlModulo = 0 Then 'no viene del SIO
                iAnio = Val(sFolio)
                iCon = Val(Mid(sFolio, InStrRev(sFolio, "/") + 1))
            Else
                If mlModulo > 0 Then 'Caso de Módulos
                    iAnio = 0
                    iCon = 0
                    sFolio = msFolio
                Else
                    iAnio = (Val(sFolio) - 2000) * 1000 + Val(Mid(sFolio, InStr(sFolio, "/") + 1))
                    iCon = Val(Mid(sFolio, InStrRev(sFolio, "/") + 1))
                End If
            End If
        End If
        
        sCam2 = sCam2 & "," & giUsuario & "," & IIf(iAnio = 0, "null", iAnio) & "," & IIf(miDel = 0, "null", miDel) & "," & IIf(iCon = 0, "null", iCon) & ",'" & UCase(sFolio) & "'," & IIf(OpcPAU(0).Value, 1, 0)
        
        If MsgBox("Se " & IIf(mbNuevo, "agregará un nuevo", "actualizará el") & " registro. ¿Está seguro de la operación?", vbYesNo + vbQuestion, "") = vbNo Then
            Exit Sub
        End If
        
'        'inicia bloqueo
'        gConSql.Execute "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
'
'        Set adorsBloqueo = New ADODB.Recordset
'        'bloquea el último id de registroxif
'        adorsBloqueo.Open "select * from registroxif where idreg=(select max(id) from registro)", gConSql, adOpenDynamic, adLockPessimistic
'        adorsBloqueo.AddNew
'        bBloqueo = True
'
'        Set adors = New ADODB.Recordset
'        adors.Open "select max(id) from registro", gConSql, adOpenStatic, adLockReadOnly
'        If adors(0) > 0 Then
'            mlAsunto = adors(0) + 1
'        Else
'            mlAsunto = 1
'        End If
'        gConSql.BeginTrans
'        bBeginTrans = True
'        gConSql.Execute "insert into registro (id," & sCam & "idusi,año,consecutivo,expediente,PAU) values (" & mlAsunto & "," & sVal & giUsuario & "," & IIf(iAnio = 0, "null", iAnio) & "," & IIf(iCon = 0, "null", iCon) & ",'" & UCase(sFolio) & "'," & IIf(OpcPAU(0).Value, 1, 0) & ")", iRows
'        If iRows > 0 Then
'            If cmbCampo(3).Visible Then 'Una sola Institución
'                If adors.State Then adors.Close
'                adors.Open "select max(id) from registroxif", gConSql, adOpenStatic, adLockReadOnly
'                If Not IsNull(adors(0)) Then
'                    l = adors(0) + 1
'                Else
'                    l = 1
'                End If
'                s = "insert into registroxif (id,idreg,idcla,idins,registro) values (" & l & "," & mlAsunto & "," & cmbCampo(3).ItemData(cmbCampo(3).ListIndex) & "," & cmbCampo(4).ItemData(cmbCampo(4).ListIndex) & ",sysdate)"
'                gConSql.Execute s, iRows
'                If iRows <= 0 Then
'                    If gConSql.Errors.Count > 0 Then
'                        MsgBox "No se realizó el alta de la institución." & gConSql.Errors(0).Description, vbCritical + vbOKOnly, ""
'                    Else
'                        MsgBox "No se realizó el alta de la institución", vbInformation + vbOKOnly, ""
'                    End If
'                    gConSql.RollbackTrans
'                    bBeginTrans = False
'                    Exit Sub
'                End If
'            Else
'                'Varias Instituciones
'                If adors.State Then adors.Close
'                adors.Open "select max(id) from registroxif", gConSql, adOpenStatic, adLockReadOnly
'                If Not IsNull(adors(0)) Then
'                    l = adors(0) + 1
'                Else
'                    l = 1
'                End If
'                s1 = msClaseIns
'                s2 = msInstituciones
'                Do While InStr(s1, ",")
'                    s = "insert into registroxif (id,idreg,idcla,idins,registro) values (" & l & "," & mlAsunto & "," & Val(s1) & "," & Val(s2) & ",sysdate)"
'                    gConSql.Execute s, i
'                    iRows = iRows + i
'                    s1 = Mid(s1, InStr(s1, ",") + 1)
'                    s2 = Mid(s2, InStr(s2, ",") + 1)
'                    l = l + 1
'                Loop
'                If iRows <= 0 Then
'                    If gConSql.Errors.Count > 0 Then
'                        MsgBox "No se realizó el alta de la institución." & gConSql.Errors(0).Description, vbCritical + vbOKOnly, ""
'                    Else
'                        MsgBox "No se realizó el alta de la institución", vbInformation + vbOKOnly, ""
'                    End If
'                    gConSql.RollbackTrans
'                    bBeginTrans = False
'                    Exit Sub
'                End If
'            End If
'            'Ya no se hace así se guarda en la misma tabla el No. de Expediente sea o no del PAU
''            If miAnio > 0 And miDel And mlCon > 0 Then
''                gConSql.Execute "insert into registroasuntosio (idreg,idasusio,año,iddel,consecutivo,registro) values (" & mlAsunto & "," & mlAsunto & "," & miAnio & "," & miDel & "," & mlCon & ",sysdate)", iRows
''            End If
'            If Len(txtObs.Text) > 5 Then
'                gConSql.Execute "insert into registroobs (idreg,observaciones) values(" & mlAsunto & ",'" & Replace(txtObs.Text, "'", "''") & "')"
'            End If
'
'            'Actualiza datos del SIO Producto Causa
'            If OpcPAU(0).Value Then 'viene del SIO
'                gConSql.Execute "insert into registrodatossio (idreg,idasu,idpr1,idpr2,idpr3,idcau,registro) values (" & mlAsunto & "," & mlAsuSIO & "," & miPr1 & "," & miPr2 & "," & miPr3 & "," & miCau & ",sysdate)"
'            End If
'
'            'Actualiza datos de Módulos si es que proviene de ahí
'            If mlModulo > 0 Then 'viene de Módulos
'                gConSql.Execute "insert into registromodulos (idreg,idmod,registro,procede) values (" & mlAsunto & "," & mlModulo & ",sysdate," & miModPro & " )"
'            End If
'            'Actualiza datos de Productos en su caso
'            If mlModulo > 0 Then 'viene de Módulos
'                gConSql.Execute "insert into registromodulos (idreg,idmod,registro,procede) values (" & mlAsunto & "," & mlModulo & ",sysdate," & miModPro & " )"
'            End If
'
'            gConSql.CommitTrans
            
            'Call MsgBox("3:" & "{call p_registroguardadatos(" & mlAsunto & "," & sCam & ",'" & msClaIns & "','" & msProSel & "'," & sCam2 & "," & mlAsuSIO & "," & mlModulo & "," & miModPro & ")}", vbOKOnly, "")

            If adors.State Then adors.Close
            adors.Open "{call p_registroguardadatos(" & mlAsunto & "," & sCam & ",'" & msClaIns & "','" & msProSel & "'," & sCam2 & "," & mlAsuSIO & "," & mlModulo & "," & miModPro & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
            
            'Call MsgBox("4:", vbOKOnly, "")

            If adors(0) > 0 Then
                If mbNuevo Then
                    miModPro = 0
                    MsgBox "Se dio de alta el expediente: " & adors(1), vbOKOnly + vbInformation, ""
                    mbCambio = False
                    Set Botón = Me.Toolbar.Buttons(1)
                    Call Toolbar_ButtonClick(Botón)
                Else 'Actualización
                    MsgBox "Se actualizó información del expediente: " & adors(1), vbOKOnly + vbInformation, ""
                End If
            Else
                MsgBox "No se realizó el alta del asunto. ", vbCritical + vbOKOnly, ""
            End If
            'miModPro = 0
            'Set adors = New ADODB.Recordset
            'adors.Open "select expediente from registro where id=" & mlAsunto, gConSql, adOpenStatic, adLockReadOnly
            'If adors(0) > 0 Then
            '    MsgBox "Se dio de alta el expediente: " & adors(0), vbOKOnly + vbInformation, ""
            'End If
            'mbCambio = False
            'Set Botón = Me.Toolbar.Buttons(1)
            'Call Toolbar_ButtonClick(Botón)
'        Else
'            gConSql.RollbackTrans
'            bBeginTrans = False
'            If gConSql.Errors.Count > 0 Then
'                MsgBox "No se realizó el alta del asunto. " & gConSql.Errors(0).Description, vbCritical + vbOKOnly, ""
'            Else
'                MsgBox "No se realizó el alta del asunto. ", vbCritical + vbOKOnly, ""
'            End If
'            Exit Sub
'        End If
'    Else 'Actualiza
'        For i = 0 To txtCampo.UBound - 1
'            If Len(Trim(txtCampo(i).Text)) = 0 Then
'                If i = 0 Then 'En caso de nombre sin dato asigna --
'                    txtCampo(0) = "--"
'                Else
'                    Exit For
'                End If
'            End If
'        Next
'        If i < txtCampo.UBound Then
'            MsgBox "Falta dato requerido: " & txtCampo(i).DataField, vbInformation + vbOKOnly, "validación"
'            Exit Sub
'        End If
'        For i = 1 To cmbCampo.UBound
'            If cmbCampo(i).ListIndex < 0 And cmbCampo(i).Visible Then Exit For
'        Next
'        If i <= cmbCampo.UBound Then
'            MsgBox "Falta dato requerido: " & Eti(i).DataField, vbInformation + vbOKOnly, "validación"
'            Exit Sub
'        End If
'        If Not cmbCampo(3).Visible Then 'varias IF
'            If InStr(msInstituciones, ",") = 0 Then
'                MsgBox "Falta capturar institución", vbInformation + vbOKOnly, ""
'                Exit Sub
'            End If
'        End If
'
'        sVal = ""
'        For i = 0 To Controls.Count - 1
'            If n > 0 Then n = 0
'            If LCase(Controls(i).Name) = "txtcampo" Then
'                If LCase(Controls(i).DataField) <> LCase(msPrin) Then
'                    If InStr(Controls(i).Tag, "|") > 0 Then
'                        If Mid(Controls(i).Tag, InStr(Controls(i).Tag, "|") + 1) <> Controls(i).Text Then
'                            n = 1
'                        End If
'                    Else
'                        n = 1
'                    End If
'                    If n > 0 Then
'                        s = ArmaCadenaCampo(Controls(i).DataField, Controls(i).Text, Controls(i).Tag, 0)
'                        sVal = sVal & s
'                    End If
'                End If
'            ElseIf Mid(Controls(i).Name, 1, 3) = "cmb" Then
'                If Len(Trim(Controls(i).DataField)) > 0 Then
'                    If InStr(Controls(i).Tag, "|") > 0 Then
'                        If Val(Mid(Controls(i).Tag, InStr(Controls(i).Tag, "|") + 1)) <> Controls(i).ListIndex Then
'                            n = 1
'                        End If
'                    Else
'                        n = 1
'                    End If
'                    If n > 0 Then
'                        If Controls(i).ListIndex >= 0 Then
'                            s = Controls(i).ItemData(Controls(i).ListIndex)
'                        Else
'                            s = ""
'                        End If
'                        s = ArmaCadenaCampo(Controls(i).DataField, s, "n", 0)
'                        sVal = sVal & s
'                    End If
'                End If
'            ElseIf Mid(Controls(i).Name, 1, 3) = "chk" Then
'                If InStr(Controls(i).Tag, "|") > 0 Then
'                    If Val(Mid(Controls(i).Tag, InStr(Controls(i).Tag, "|") + 1)) <> IIf(Controls(i).Value = 1, 1, 0) Then
'                        n = 1
'                    End If
'                End If
'
'                If n > 0 And Len(Controls(i).DataField) > 0 Then
'                    sVal = sVal & Controls(i).DataField & "=" & IIf(Controls(i).Value = 1, 1, 0) & ","
'                End If
'            End If
'        Next
'        'mlAsunto = Val(txtCampo(0).Text)
'        If gi = 199 Then 'indica que no debe realizar confirmación
'            gi = 0
'        Else
'            If MsgBox("Se actualizarán datos del asunto. ¿Está seguro de la operación?", vbYesNo + vbQuestion, "") = vbNo Then
'                Exit Sub
'            End If
'        End If
'
'        gConSql.BeginTrans
'        bBeginTrans = True
'
'        If OpcPAU(0).Value Then
'            's = txtExpediente.Text
'            'sVal = sVal & "pau=1, expediente=" & txtExpediente.Text & ","
'        Else
'            's = txtNuevoExp.Text
'            'sVal = sVal & "pau=0, expediente=" & txtExpediente.Text & ","
'        End If
'        If Len(sVal) > 0 Then
'            If Right(sVal, 1) = "," Then
'                sVal = Mid(sVal, 1, Len(sVal) - 1)
'            End If
'            gConSql.Execute "update registro set " & sVal & " where id=" & mlAsunto, iRows
'            If iRows <= 0 Then
'                gConSql.RollbackTrans
'                bBeginTrans = False
'                If gConSql.Errors.Count > 0 Then
'                    MsgBox "No se actualizó la información. " & Err.Description, vbOKOnly + vbCritical, ""
'                Else
'                    MsgBox "No se actualizó la información. " & gConSql.Errors(0).Description, vbOKOnly + vbCritical, ""
'                End If
'                Exit Sub
'            End If
'            'RefrescaDatos
'        End If
'        If cmbCampo(3).Visible Then 'Una sola Ins
'            If adors.State Then adors.Close
'            adors.Open "select min(id) from registroxif where idreg=" & mlAsunto, gConSql, adOpenStatic, adLockReadOnly
'            If adors(0) > 0 Then
'                gConSql.Execute "Update registroxif set idcla=" & cmbCampo(3).ItemData(cmbCampo(3).ListIndex) & ", idins=" & cmbCampo(4).ItemData(cmbCampo(4).ListIndex) & " where id=" & adors(0)
'            ElseIf adors(0) = 0 Then
'                If adors.State Then adors.Close
'                adors.Open "select max(id) from registroxif ", gConSql, adOpenStatic, adLockReadOnly
'                If adors(0) > 0 Then
'                    l = adors(0) + 1
'                Else
'                    l = 1
'                End If
'                gConSql.Execute "insert into registroxif (id,idreg,idcla,idins) values (" & l & "," & mlAsunto & "," & cmbCampo(3).ItemData(cmbCampo(3).ListIndex) & "," & cmbCampo(4).ItemData(cmbCampo(4).ListIndex) & ")"
'            End If
'        Else 'varias inst
'            If adors.State Then adors.Close
'            adors.Open "select max(id) from registroxif ", gConSql, adOpenStatic, adLockReadOnly
'            If adors(0) > 0 Then
'                l = adors(0) + 1
'            Else
'                l = 1
'            End If
'            'gConSql.Execute "delete from registroxif where idreg=" & mlAsunto, iRows
''            If adors.State Then adors.Close
''            adors.Open "select id from registroxif where idreg=" & mlAsunto, gConSql, adOpenStatic, adLockReadOnly
'            s1 = msClaseIns
'            s2 = msInstituciones
'            s = ""
'            Do While InStr(s1, ",") > 0
'                If adors.State Then adors.Close
'                adors.Open "select count(*) from registroxif where idreg=" & mlAsunto & " and idcla=" & Val(s1) & " and idins=" & Val(s2), gConSql, adOpenStatic, adLockReadOnly
'                If adors(0) > 0 Then 'Actualiza
'                    'gConSql.Execute "Update registroxif set idcla=" & Val(s1) & ", idins=" & Val(s2) & " where id=" & adors(0)
'                    'adors.MoveNext
'                Else 'nuevo registroxif
'                    gConSql.Execute "insert into registroxif (id,idreg,idcla,idins) values (" & l & "," & mlAsunto & "," & Val(s1) & ", " & Val(s2) & ")"
'                    l = l + 1
'                End If
'                s = s & Val(s1) & "|" & Val(s2) & "|"
'                s1 = Mid(s1, InStr(s1, ",") + 1)
'                s2 = Mid(s2, InStr(s2, ",") + 1)
'            Loop
'            gConSql.Execute "delete from registroxif where idreg=" & mlAsunto & " and instr('" & s & "', idcla||'|'||idins||'|')=0", iRows
'        End If
'        If adors.State Then adors.Close
'        adors.Open "select count(*) from registroobs where idreg=" & mlAsunto, gConSql, adOpenStatic, adLockReadOnly
'        If adors(0) > 0 Then
'            i = 1
'        Else
'            i = 0
'        End If
'        If Len(txtObs.Text) > 5 Then
'            If i = 1 Then 'Existe por tanto update
'                gConSql.Execute "update registroobs set observaciones='" & Replace(txtObs.Text, "'", "''") & "' where idreg=" & mlAsunto
'            Else ' no existe por tanto insert
'                gConSql.Execute "insert into registroobs (idreg,observaciones) values(" & mlAsunto & ",'" & Replace(txtObs.Text, "'", "''") & "')"
'            End If
'        Else
'
'        End If
'        gConSql.CommitTrans
'        'If adors(0) > 0 Then
'            MsgBox "Se actualizó el expediente", vbOKOnly + vbInformation, ""
'        'End If
'        'Call ActualizaBotones(Me, 2, myPermiso)
'        Set Botón = Me.Toolbar.Buttons(3)
'        Call Toolbar_ButtonClick(Botón)
'    End If
'    If AdorsPrin.State > 0 Then AdorsPrin.Close
'    AdorsPrin.Open Mid(msConsultaP, 1, InStr(msConsultaP, " where ") - 1) & " where id=" & mlAsunto, gConSql, adOpenStatic, adLockReadOnly
'    'AdorsCata.Requery
'    'AdorsPrin.AbsolutePosition = ii
'    opcReg(0).Value = False
'    opcReg(1).Value = False
    'RefrescaDatos

Case "Eliminar"
    If mlAsunto > 0 Then
        If MsgBox("¿Está Seguro de borrar el registro seleccionado?", vbYesNo + vbQuestion, "Confirmación") = vbYes Then
            i = AdorsPrin.Bookmark
            'GuardaBitácora gs_usuario, myTabla, Val(txtCampo(0).Text), 5
            adors.Open "{call P_Reg_borraReg(" & mlAsunto & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
            If Not adors.EOF Then
                If adors(0) > 0 Then
                    MsgBox "El registro se eliminó correctamente", vbOKOnly, ""
                Else
                    MsgBox "Ocurrió un problema al intentar eliminar el asunto: " & adors(1), vbOKOnly, ""
                End If
            End If
            Set Botón = Me.Toolbar.Buttons(2)
            Call Toolbar_ButtonClick(Botón)
            If i > AdorsPrin.RecordCount Then
                i = AdorsPrin.RecordCount
            End If
        End If
    End If
Case "Deshacer"
    If MsgBox("Está seguro de deshacer los cambios", vbYesNo + vbQuestion, "") = vbNo Then
        Exit Sub
    End If
    If mlAsunto > 0 Then
        RefrescaDatos
        mbCambio = False
    Else
        Set Botón = Me.Toolbar.Buttons(1)
        Call Toolbar_ButtonClick(Botón)
    End If
Case "Nuevo"
    If mbCambio And Toolbar.Buttons(4).Enabled Then
        If MsgBox("¿Desea ignorar los cambios realizados?", vbYesNo + vbQuestion, "") = vbNo Then
            Exit Sub
        End If
    End If
    mbNuevo = True
    mbLimpia = True
    
    For i = txtCampo.LBound To txtCampo.UBound
        txtCampo(i).Text = ""
        If txtCampo(i).Locked Then txtCampo(i).Locked = False
    Next
    For i = cmbCampo.LBound To cmbCampo.UBound
        cmbCampo(i).ListIndex = -1
        If cmbCampo(i).Locked Then cmbCampo(i).Locked = False
        'cmbCampo(i).Text = ""
    Next
    txtNoReg.Locked = True
    Call ActualizaBotones(Me, 1, myPermiso)
    mlAsunto = 0
    If Not cmbCampo(3).Visible Then
        f_UnaSolaIns True
    End If
    ListClaseIns.Clear
    msInstituciones = ""
    msClaseIns = ""
    txtObs = ""
    txtExpediente.Text = ""
    txtNuevoExp.Text = ""
    txtCausa.Text = ""
    txtProducto.Text = ""
    mlAsuSIO = 0
    msFolio = ""
    miAnio = 0
    miDel = 0
    mlCon = 0
    mbCambio = False
    mbNuevo = True
    'cmbResp.Text = ""
    'Desinhibe los campos principales
    InhibeCampoPrinc False
    cmdContinuar.Enabled = True
    If OpcPAU(1).Value Then
        OpcPAU_Click 1
    End If
    cmdModulos.Enabled = True
    txtExpMod.Text = ""
    txtExpMod.Enabled = True
    mlModulo = 0
    lblExpMod.Caption = ""
    lblTipoExp.Caption = ""
    txtResObs.Text = ""
    txtResObs.Visible = False
    lblResObs.Visible = False
    txtObs.Width = 8880
    If Not cmdAgregarIF.Enabled Then cmdAgregarIF.Enabled = True
    msProSel = ""
    miSIO_SoloPro = 0
    opcVig(1).Value = False
    opcVig(0).Enabled = True
    opcVig(1).Enabled = True
    opcSIO(0).Enabled = True
    opcSIO(1).Enabled = True
    opcSIO(2).Enabled = True
    chkSinCau.Value = 0
    chkSinCau.Enabled = False
    mbLimpia = False
Case "Limpiar"
    If mbCambio Then
        If MsgBox("¿Desea ignorar los cambios realizados?", vbYesNo + vbQuestion, "") = vbNo Then
            Exit Sub
        End If
    End If
    mbLimpia = True
    mbNuevo = False
    For i = txtCampo.LBound To txtCampo.UBound
        txtCampo(i).Text = ""
        If txtCampo(i).Locked Then txtCampo(i).Locked = False
    Next
    For i = cmbCampo.LBound To cmbCampo.UBound
        cmbCampo(i).ListIndex = -1
        'cmbCampo(i).Text = ""
        If cmbCampo(i).Locked Then cmbCampo(i).Locked = False
    Next
    txtNoReg.Locked = True
    Call ActualizaBotones(Me, 2, myPermiso)
    mlAsunto = 0
    If Not cmbCampo(3).Visible Then
        f_UnaSolaIns True
    End If
    ListClaseIns.Clear
    msInstituciones = ""
    msClaseIns = ""
    txtExpediente.Text = ""
    txtNuevoExp.Text = ""
    txtObs = ""
    txtCausa.Text = ""
    txtProducto.Text = ""
    mlAsuSIO = 0
    msFolio = ""
    miAnio = 0
    miDel = 0
    mlCon = 0
    mbCambio = False
    If Not txtNuevoExp.Visible Then
        txtNuevoExp.Visible = True
        Label2.Visible = True
    End If
    'Desinhibe los campos principales
    InhibeCampoPrinc False
    cmdContinuar.Enabled = False
    cmdModulos.Enabled = False
    txtExpMod.Text = ""
    txtExpMod.Enabled = flase
    'cmbResp.Text = ""
    mlModulo = 0
    lblExpMod.Caption = ""
    lblTipoExp.Caption = ""
    txtResObs.Text = ""
    txtResObs.Visible = False
    lblResObs.Visible = False
    txtObs.Width = 8880
    If Not cmdAgregarIF.Enabled Then cmdAgregarIF.Enabled = True
    msProSel = ""
    chkSinCau.Value = 0
    miSIO_SoloPro = 0
    opcVig(1).Value = False
    opcVig(0).Enabled = True
    opcVig(1).Enabled = True
    opcSIO(0).Enabled = True
    opcSIO(1).Enabled = True
    opcSIO(2).Enabled = True
    chkSinCau.Value = 0
    chkSinCau.Enabled = False
Case "Buscar"
    If Len(Trim(txtExpediente.Text)) > 0 Or Len(Trim(txtNuevoExp.Text)) > 0 Then
        If Len(Trim(txtExpediente)) > 0 Then
            If adors.State Then adors.Close
            adors.Open "select f_asuntoxfolioSIO('" & txtExpediente.Text & "') from dual", gConSql, adOpenStatic, adLockReadOnly
            If adors(0) > 0 Then
                l = adors(0)
                If AdorsPrin.State > 0 Then AdorsPrin.Close
                AdorsPrin.Open Mid(msConsultaP, 1, InStr(msConsultaP, " where ") - 1) & " where id=" & l, gConSql, adOpenStatic, adLockReadOnly
                RefrescaDatos
            Else
                MsgBox "No se encontró asunto alguno con ese Folio del SIO.", vbOKOnly + vbInformation, ""
                Exit Sub
            End If
        Else
            If adors.State Then adors.Close
            adors.Open "select f_asuntoxfolio('" & txtNuevoExp.Text & "') from dual", gConSql, adOpenStatic, adLockReadOnly
            If adors(0) > 0 Then
                l = adors(0)
                If AdorsPrin.State > 0 Then AdorsPrin.Close
                AdorsPrin.Open Mid(msConsultaP, 1, InStr(msConsultaP, " where ") - 1) & " where id=" & l, gConSql, adOpenStatic, adLockReadOnly
                RefrescaDatos
            Else
                MsgBox "No se encontró asunto alguno con ese Número de Expediente.", vbOKOnly + vbInformation, ""
                Exit Sub
            End If
        End If
        Call ActualizaBotones(Me, 3, myPermiso)
        Exit Sub
    Else
        sCondición = ""
        For i = 0 To Controls.Count - 1
            If LCase(Controls(i).Name) = "txtcampo" Then
                If Len(Trim(Controls(i).DataField)) > 0 And Len(Trim(Controls(i).Text)) > 0 Then
                    s = ArmaCadenaCampo(Controls(i).DataField, Controls(i).Text, Controls(i).Tag, 1)
                    sCondición = sCondición & s
                End If
            ElseIf Mid(Controls(i).Name, 1, 3) = "cmb" Then
                If Controls(i).ListIndex >= 0 Then
                    If Len(Trim(Controls(i).DataField)) > 0 Then
                        s = Controls(i).ItemData(Controls(i).ListIndex)
                        s = ArmaCadenaCampo(Controls(i).DataField, s, "n", 1)
                        sCondición = sCondición & s
                    End If
                End If
            ElseIf Mid(Controls(i).Name, 1, 3) = "chk" Then
                If Controls(i).Value < 2 And Len(Controls(i).DataField) > 0 Then
                    sCondición = sCondición & Controls(i).DataField & "=" & IIf(Controls(i).Value = 1, 1, 0) & " and "
                End If
            End If
        Next
        If Len(txtObs.Text) > 0 Then
            If txtObs.Text = "*" Or txtObs.Text = "-*" Then
                sCondición = sCondición & " id " & IIf(txtObs.Text = "*", "", "not ") & "in (select idreg from registroobs) and "
            Else
                s1 = ArmaCadenaCampo("observaciones", txtObs.Text, "c", 1)
                sCondición = sCondición & " id in (select idreg from registroobs where " & Mid(s1, 1, Len(s1) - 5) & ") and "
            End If
        End If
    End If
    If txtCampo(3).Visible Then 'una sola Inst
        If cmbCampo(3).ListIndex >= 0 And cmbCampo(4).ListIndex >= 0 Then
            sCondición = sCondición & "id in (select idreg from registroxif where idcla||'|'||idins='" & cmbCampo(3).ItemData(cmbCampo(3).ListIndex) & "|" & cmbCampo(3).ItemData(cmbCampo(3).ListIndex) & "') and "
        ElseIf cmbCampo(3).ListIndex >= 0 Then
            sCondición = sCondición & "id in (select idreg from registroxif where idcla=" & cmbCampo(3).ItemData(cmbCampo(3).ListIndex) & ") and "
        ElseIf cmbCampo(4).ListIndex >= 0 Then
            sCondición = sCondición & "id in (select idreg from registroxif where idins=" & cmbCampo(4).ItemData(cmbCampo(4).ListIndex) & ") and "
        End If
    Else
        sCondición = sCondición & "id in (select idreg from registroxif where idcla||'|'||idins='" & Val(msClaseIns) & "|" & Val(msInstituciones) & "') and "
    End If
    
    If Len(sCondición) > 3 Then
        sCondición = " where " & Mid(sCondición, 1, Len(sCondición) - 5)
        If AdorsPrin.State > 0 Then AdorsPrin.Close
        AdorsPrin.Open Mid(msConsultaP, 1, InStr(msConsultaP, " where ") - 1) & sCondición, gConSql, adOpenStatic, adLockReadOnly
        If AdorsPrin.RecordCount > 0 Then
            RefrescaDatos
            Call ActualizaBotones(Me, 3, myPermiso)
        Else
            Call MsgBox("No existen registros que cumplan el criterio de búsqueda", vbOKOnly + vbInformation, "")
        End If
    Else
        opcReg_Click (1)
    End If
        
Case "Imprimir"
'    i = MsgBox("¿Desea emitir informe del personal (SI) o Gafete(NO)?", vbYesNoCancel, "")
'    If i = vbCancel Then Exit Sub
'    CReport.ReportFileName = gsDirReportes + IIf(vbYes = i, "personal.rpt", "gafete.rpt")
'    CReport.ParameterFields(0) = "@ipersona;" & mlAsunto & "); true"
'
'    s = LCase(gConSql.ConnectionString)
'    s = Mid(s, InStr(s, ";pwd=") + 1)
'    s = Mid(s, 1, InStr(Mid(s, 2), ";"))
'    CReport.Connect = "Filedsn=rh.dsn;" & s
'
'    CReport.Action = 1
'    gi = mlAsunto
'    RHReportePersonal.Show vbModal
Dim S_ServerExt, S_BaseDatosExt, S_PassExt, S_LogExt
    If mlAsunto = 0 Then
        Exit Sub
    End If
    If myPermisoRep = 0 Then
        MsgBox "No cuenta con privilegios para emitir informes de este módulo", vbOKOnly + vbInformation, "Validación"
        Exit Sub
    End If
'    CReport.ReportFileName = gsDirReportes + "\Personal.rpt"
'    CReport.ParameterFields(0) = "iPersona;" & mlAsunto & "; true"
'    S_LogExt = Fu_LeeDatosArchConfig(1, "Central")
'    S_PassExt = Fu_LeeDatosArchConfig(2, "Central")
'    S_BaseDatosExt = Fu_LeeDatosArchConfig(3, "Central")
'    S_ServerExt = Fu_LeeDatosArchConfig(4, "Central")
'    CReport.Connect = "Provider=SQLOLEDB;SERVER=" & S_ServerExt & ";DATABASE=" & S_BaseDatosExt & "; UID=" & S_LogExt & "; PWD=" & S_PassExt
'    CReport.Action = 1
    
    
    
'    gn_OpcionReporte = 4
'    If mlAsunto > 0 Then
'        gi2 = mlAsunto
'    End If
'    gi1 = 102
'    Load RHReportes
'    RHReportes.Caption = "Reportes: " & Trim(MDI_Prin.mnuRepRH(gn_OpcionReporte).Caption)
'    RHReportes.Show vbModal
'    gi1 = 0
'
Case "Salir"

    Unload Me

End Select
Exit Sub

ErrorAcción:

If mbNuevo Then
    mlAsunto = 0
End If
If bBeginTrans Then
    gConSql.RollbackTrans
    gConSql.Execute "SET TRANSACTION ISOLATION LEVEL SERIALIZABLE"
    Set adorsBloqueo = Nothing
    If gConSql.Errors.Count > 0 Then
        Call MsgBox("Error: " + gConSql.Errors(0).Description, vbOKOnly + vbCritical, "Error no esperado (" + Str(gConSql.Errors(0).Number) + ")")
    Else
        Call MsgBox("Error: " + Err.Description, vbOKOnly + vbCritical, "Error no esperado (" + Str(Err.Number) + ")")
    End If
    'Resume
    Exit Sub
End If

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

'Actualiza datos en el formulario originados por el movimiento de registro del cursor
Sub RefrescaDatos()
Dim i As Long, l As Long
Dim ii As Integer, yError As Byte
Dim adors As New ADODB.Recordset
On Error GoTo ErrorRefresca:
mbRefrescaDatos = True
If adorsCata.State Then adorsCata.Close
If AdorsPrin.EOF Then
    Exit Sub
End If
mlAsuSIO = 0
micau = 0
miPr1 = 0
miPr2 = 0
miPr3 = 0
msFolio = ""
mlAsunto = AdorsPrin!ID

adorsCata.Open "select expediente,pau," & Mid(msConsulta, 8) & " where id=" & mlAsunto, gConSql, adOpenStatic, adLockReadOnly

If mlAsunto > 0 Then 'Obtiene datos de la IF y Productos
    If adors.State Then adors.Close
    adors.Open "select f_registro_clainspro(" & mlAsunto & ",1) as ClaseIns, f_registro_clainspro(" & mlAsunto & ",2) as Prods from dual", gConSql, adOpenStatic, adLockReadOnly
    msClaIns = adors(0)
    msProSel = adors(1)
End If

If adorsCata!PAU <> 0 Then
    OpcPAU(0).Value = True
Else
    OpcPAU(1).Value = True
End If
For i = 0 To txtCampo.UBound
    txtCampo(i).Text = IIf(IsNull(adorsCata(txtCampo(i).DataField)), "", adorsCata(txtCampo(i).DataField))
    If InStr(txtCampo(i).Tag, "|") > 0 Then
        txtCampo(i).Tag = Mid(txtCampo(i).Tag, 1, InStr(txtCampo(i).Tag, "|")) & adorsCata(txtCampo(i).DataField)
    Else
        txtCampo(i).Tag = txtCampo(i).Tag & "|" & adorsCata(txtCampo(i).DataField)
    End If
    If txtCampo(i).Tag = "f" Or InStr(txtCampo(i).Tag, "f|") > 0 Then
        ii = i
        Call txtCampo_LostFocus(ii)
    End If
Next
For i = 0 To cmbCampo.UBound
    If Len(cmbCampo(i).DataField) > 0 Then
        l = IIf(IsNull(adorsCata(cmbCampo(i).DataField)), -1, adorsCata(cmbCampo(i).DataField))
        If InStr(cmbCampo(i).Tag, "|") > 0 Then
            cmbCampo(i).Tag = Mid(cmbCampo(i).Tag, 1, InStr(cmbCampo(i).Tag, "|")) & l
        Else
            cmbCampo(i).Tag = "|" & l
        End If
        If l > 0 Then
            l = BuscaCombo(cmbCampo(i), l, True)
            If l >= 0 Then
                cmbCampo(i).ListIndex = l
            Else
                If i = 0 Then
                    s = "select descripción from direccióngeneral where id=" & l
                ElseIf i = 1 Then
                    s = "select descripción from unidades where id=" & l
                ElseIf i = 2 Then
                    s = "select descripción from materiasanción where id=" & l
                ElseIf i = 3 Then
                    s = "select descripción from claseinstitución where id=" & l
                ElseIf i = 4 Then
                    s = "select descripción from instituciones where id=" & l
                ElseIf i = 5 Then
                    s = "select descripción from usuariossistema where id=" & l
                End If
                If adors.State Then adors.Close
                adors.Open s, gConSql, adOpenStatic, adLockReadOnly
                If Not adors.EOF Then
                    cmbCampo(i).Text = adors(0)
                Else
                    cmbCampo(i).Text = ""
                End If
            End If
        Else
            cmbCampo(i).ListIndex = -1
        End If
    ElseIf i = 3 Then 'clase e institución (3,4)
        If adors.State Then adors.Close
        adors.Open "select r.idcla,r.idins,ci.descripción, i.descripción,r.id from registroxif r, claseinstitución ci, instituciones i where r.idreg=" & mlAsunto & " and r.idcla=ci.id(+) and r.idins=i.id(+)", gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            mlRegxIf = adors!ID
            If adors.RecordCount > 1 Then 'Varias Instituciones
                f_UnaSolaIns False
                msInstituciones = ""
                msClaseIns = ""
                ListClaseIns.Clear
                Do While Not adors.EOF
                    ListClaseIns.AddItem adors(3) & " (" & adors(2) & ")"
                    ListClaseIns.ItemData(ListClaseIns.NewIndex) = adors(1)
                    msInstituciones = msInstituciones & adors(1) & ","
                    msClaseIns = msClaseIns & adors(0) & ","
                    adors.MoveNext
                Loop
            Else
                f_UnaSolaIns True
                msInstituciones = ""
                msClaseIns = ""
                l = BuscaCombo(cmbCampo(3), adors(0), True)
                If l >= 0 Then
                    cmbCampo(3).ListIndex = l
                Else
                    cmbCampo(3).Text = adors(2)
                End If
                l = BuscaCombo(cmbCampo(4), adors(1), True)
                If l >= 0 Then
                    cmbCampo(4).ListIndex = l
                Else
                    cmbCampo(4).Text = adors(3)
                End If
            End If
        Else
            cmbCampo(i).Text = ""
        End If
        
    End If
Next
'Coloca los datos de los folio SIAM y SIo en su caso
If adorsCata!PAU <> 0 Then
    OpcPAU(0).Value = True
    txtExpediente = adorsCata!expediente
    txtNuevoExp = ""
Else
    OpcPAU(1).Value = True
    txtExpediente = ""
    txtNuevoExp = adorsCata!expediente
End If
'If adors.State Then adors.Close
'adors.Open "select f_folio(" & mlAsunto & ") from dual", gConSql, adOpenStatic, adLockReadOnly
'txtNuevoExp.Text = IIf(IsNull(adors(0)), "", adors(0))
'If adors.State Then adors.Close
'adors.Open "select f_foliosiamsio(" & mlAsunto & ") from dual", gConSql, adOpenStatic, adLockReadOnly
'txtExpediente.Text = IIf(IsNull(adors(0)), "", adors(0))
If Not Frame1.Enabled Then Frame1.Enabled = True
ActualizaCampoEsp
'inhibe los campos y datos iniciales, ya no se pueden modificar
InhibeCampoPrinc True


If AdorsPrin.RecordCount = 0 Then
    MsgBox "No Existen Registros", vbOKOnly + vbInformation, ""
Else
    'muestra observaciones
    If adors.State Then adors.Close
    adors.Open "select observaciones from registroobs where idreg=" & mlAsunto, gConSql, adOpenStatic, adLockReadOnly
    If adors.EOF Then
        txtObs.Text = ""
    Else
        txtObs.Text = IIf(IsNull(adors(0)), "", adors(0))
    End If
    If AdorsPrin.Bookmark > 0 Then
        l = AdorsPrin(0)
        txtNoReg.Text = (AdorsPrin.Bookmark) & " / " & AdorsPrin.RecordCount
    Else
        txtNoReg.Text = "??? / " & AdorsPrin.RecordCount
    End If
    txtNoReg.Refresh
End If
mbLimpia = False
mbCambio = False
'Actualiza datos de Producto y causa del SIO
If mlAsunto > 0 Then
    If adors.State Then adors.Close
    adors.Open "select f_registroxif_prodcau(" & mlRegxIf & ",1),f_registroxif_prodcau(" & mlRegxIf & ",0) from dual", gConSql, adOpenStatic, adLockReadOnly
    txtProducto.Text = adors(0)
    txtCausa.Text = adors(1)
    If Len(msProSel) >= 40 Then
        miPr1 = Val(Mid(msProSel, 1, 10))
        miPr2 = Val(Mid(msProSel, 11, 10))
        miPr3 = Val(Mid(msProSel, 21, 10))
        micau = Val(Mid(msProSel, 31, 10))
        opcSIO(0).Enabled = False
        opcSIO(1).Enabled = False
        opcSIO(2).Enabled = False
        opcVig(0).Enabled = False
        opcVig(1).Enabled = False
        chkSinCau.Enabled = True
        opcSIO(1).Value = True
        opcVig(1).Value = True
        If micau = 0 Then
            miSIO_SoloPro = 2
        End If
        cmdAgregarIF.Enabled = False
    ElseIf Len(msProSel) >= 30 Then
        miPr1 = Val(Mid(msProSel, 1, 10))
        miPr2 = Val(Mid(msProSel, 11, 10))
        miPr3 = Val(Mid(msProSel, 21, 10))
        opcSIO(0).Enabled = False
        opcSIO(1).Enabled = False
        opcSIO(2).Enabled = False
        opcVig(0).Enabled = False
        opcVig(1).Enabled = False
        chkSinCau.Enabled = True
        opcSIO(2).Value = True
        opcVig(1).Value = True
        micau = 0
        cmdAgregarIF.Enabled = False
    Else
        If Not cmdAgregarIF.Enabled Then cmdAgregarIF.Enabled = True
    End If
End If
cmdModulos.Enabled = False
txtExpMod.Text = ""
txtExpMod.Enabled = False
mlModulo = 0
lblExpMod.Caption = ""
lblTipoExp.Caption = ""
txtResObs.Text = ""
txtResObs.Visible = False
lblResObs.Visible = False
txtObs.Width = 8880
mbRefrescaDatos = False
Exit Sub
ErrorRefresca:

yError = MsgBox("Error: " + Err.Description, vbAbortRetryIgnore + vbCritical, "Error no esperado (" + Str(Err.Number) + ")")


If yError = vbRetry Then
    Resume
ElseIf yError = vbIgnore Then
    Resume Next
End If
mbRefrescaDatos = False
End Sub

'Actualiza combo de municipios según valor del estado y otros (sexo, edad, ...)
Sub ActualizaCampoEsp()
Static iEdo As Integer
Dim d As Date, b As Boolean, adors As New ADODB.Recordset
End Sub


'habilita bandera de cambio
Private Sub txtCampo_Change(Index As Integer)
If Not mbCambio And Not mbRefresca Then mbCambio = True
End Sub

Private Sub txtCampo_DblClick(Index As Integer)
If Mid(txtCampo(Index).Tag, 1, 1) = "f" And Len(Trim(txtCampo(Index).Text)) = 0 Then
    txtCampo(Index).Text = Format(Now, "dd/mm/yyyy")
End If
End Sub

Private Sub txtCampo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Mid(txtCampo(Index).Tag, 1, 1) = "f" And KeyCode = 27 Then txtCampo(Index) = ""
End Sub

'Valida caracteres de entrada según el tipo de campo
Private Sub txtCampo_KeyPress(Index As Integer, KeyAscii As Integer)
Dim i As Long, i1 As Long
If KeyAscii = 13 Then
    i1 = txtCampo(Index).TabIndex
    For i = Controls.Count - 1 To 0 Step -1
        If InStr("txt,cmb,cmd,chk,opc", Mid(LCase(Controls(i).Name), 1, 3)) > 0 Then
            'Debug.Print Controls(i).TabIndex
            If Controls(i).TabIndex = i1 + 1 Then
                Controls(i).SetFocus
                Exit Sub
            End If
        End If
    Next
End If
If Index = 1 And InStr("-", Chr(KeyAscii)) > 0 Then Exit Sub
KeyAscii = TeclaOprimida(txtCampo(Index), KeyAscii, txtCampo(Index).Tag, Me.Toolbar.Buttons(3).Enabled)
'If Index = 1 And txtCampo(Index).SelStart = 0 And InStr("1234567890", Chr(KeyAscii)) = 0 Then
'    KeyAscii = 0
'End If

End Sub

Private Sub txtCampo_LostFocus(Index As Integer)
Dim d As Date, Y As Integer, y2 As Integer, adors As New ADODB.Recordset
If Mid(txtCampo(Index).Tag, 1, 1) = "f" Then
    If IsDate(txtCampo(Index).Text) Then
        d = CDate(txtCampo(Index).Text)
        txtCampo(Index).Text = Format(d, gsFormatoFecha)
        adors.Open "select sysdate from dual", gConSql, adOpenStatic, adLockReadOnly
        If Index = 2 Then
            If Int(adors(0)) + 7 - Int(d) < 0 Then
                Call MsgBox("Fecha no válida. No se permite ingresar fecha mayor a la fecha (" & Format(adors(0) + 7, gsFormatoFecha) & ")", vbOKOnly + vbInformation, "")
                txtCampo(Index) = ""
                Exit Sub
            End If
        Else
            If Int(adors(0)) - Int(d) < 0 Then
                Call MsgBox("Fecha no válida. No se permite ingresar fecha mayor a la fecha actual (" & Format(adors(0), gsFormatoFecha) & ")", vbOKOnly + vbInformation, "")
                txtCampo(Index) = ""
                Exit Sub
            End If
        End If
    Else
        If Len(txtCampo(Index).Text) > 0 Then
            Call MsgBox("Fecha no válida. Verificar", vbOKOnly + vbInformation, "")
            txtCampo(Index) = ""
        End If
    End If
End If
End Sub


'Verifica y realiza acción del campo de No. Registro
Private Sub txtNoReg_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Botón As Object
If KeyCode = 13 Then
    If Val(txtNoReg) > 0 Then
        'mdi.ActiveForm.iIr_a = Val(txtReg)
        Set Botón = Me.Toolbar.Buttons(8)
        Call Toolbar_ButtonClick(Botón)
    End If
End If
End Sub

'Refrescadatos del historial de actividades
Sub ActualizaLista()
Dim i As Long
Dim adors As New ADODB.Recordset
If AdorsPrin.Bookmark <= 0 Then
    Exit Sub
End If
If adors.State > 0 Then adors.Close
adors.Open "select ps.*,r.s_evento,s.s_seguimiento,o.s_observaciones,res.s_nombre+rtrim(' '+res.s_paterno)+rtrim(' '+res.s_materno) as responsable,u.s_nombre from t_rhpersonalseg ps left join c_rhperseg s on ps.n_cveperseg=s.n_cveperseg left join c_rheventos r on ps.n_cveevento=r.n_cveevento left join t_rhpersonalsegobs o on ps.n_cveseguimiento=o.n_cveseguimiento left join t_rhpersonal res on ps.n_cveresponsable=res.n_cvepersona left join c_segusuarios u on ps.n_cveusuario=u.n_cveusuario where ps.n_cvepersona=" & IIf(IsNull(mlAsunto), -1, mlAsunto) & " and f_fecha is not null order by f_fecha", gConSql, adOpenStatic, adLockReadOnly
i = 1
ListView1.ListItems.Clear
Do While Not adors.EOF
    ListView1.ListItems.Add i, , IIf(IsNull(adors![s_seguimiento]), "", adors![s_seguimiento])
    ListView1.ListItems(i).Tag = adors!n_cveseguimiento
    ListView1.ListItems(i).SubItems(1) = IIf(IsNull(adors![s_evento]), "", adors![s_evento])
    ListView1.ListItems(i).SubItems(2) = IIf(IsNull(adors![n_resultado]), "", IIf(adors![n_resultado] = 1, "Continua", "Concluye"))
    If IsNull(adors!f_programado) Then
        ListView1.ListItems(i).SubItems(3) = ""
    Else
        ListView1.ListItems(i).SubItems(3) = Format(adors!f_programado, gsFormatoFechaHora)
    End If
    ListView1.ListItems(i).SubItems(4) = Format(adors!f_Fecha, gsFormatoFechaHora)
    ListView1.ListItems(i).SubItems(5) = IIf(IsNull(adors![s_observaciones]), "", adors![s_observaciones])
    ListView1.ListItems(i).SubItems(6) = IIf(IsNull(adors![responsable]), "", adors![responsable])
    ListView1.ListItems(i).SubItems(7) = IIf(IsNull(adors![s_Nombre]), "", adors![s_Nombre])
    adors.MoveNext
    i = i + 1
Loop
If i > 1 Then
    ListView1.ListItems(i - 1).Tag = ListView1.ListItems(i - 1).Tag & "|"
End If
If Not ListView1.Visible Then
    ListView1.Visible = True
    ListView1.Left = 80
End If
End Sub

'valida si es posible borrar la actividad
Function ValidaBorradoAct(lSeguimiento As Long) As Boolean
Dim adors As New ADODB.Recordset, s As String
adors.Open "select count(*) from t_rhpersonalseg where n_cveant=" & lSeguimiento & " and n_cveseguimiento<>n_cveant and f_fecha is not null", gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    Call MsgBox("La Actividad no es posible borrar ya que esta no es la última del proceso", vbOKOnly + vbInformation)
    Exit Function
End If
adors.Close
adors.Open "select count(*) from t_rhpersonalseg where n_cveseguimiento=" & lSeguimiento & " and f_actualización<convert(datetime,'" & Format(gdAhora - 30, gsFormatoFecha) & "',105)", gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    Call MsgBox("No se permite realizar borrado de información después de 30 días", vbOKOnly + vbInformation)
    Exit Function
End If
's = ListView1.SelectedItem.ListSubItems(4).Text
'If ListView1.SelectedItem > 0 Then
'    If InStr(s, "-") > 0 Then
'        If CDate(s) < Date - 30 Then
'            Call MsgBox("No se permite realizar borrado de información después de 30 días", vbOKOnly + vbInformation)
'        End If
'    End If
'End If
ValidaBorradoAct = True
End Function

'Cambio de opción de registros mostrados
Private Sub opcReg_Click(Index As Integer)
If Index = 0 And Not mbNoPreg Then
    If MsgBox("¿Está seguro de mostrar los últimos 10 registros del Personal?", vbYesNo + vbQuestion, "") = vbYes Then
        If AdorsPrin.State Then AdorsPrin.Close
        AdorsPrin.Open msConsultaP & msOrden, gConSql, adOpenStatic, adLockReadOnly
        Call ActualizaBotones(Me, 3, myPermiso)
        RefrescaDatos
    End If
ElseIf Index = 1 And Not mbNoPreg Then
    'If MsgBox("¿Está seguro de mostrar todo el Personal?", vbYesNo + vbQuestion, "") = vbYes Then
        If AdorsPrin.State Then AdorsPrin.Close
        AdorsPrin.Open Mid(msConsultaP, 1, InStr(msConsultaP, " where ") - 1) & " order by id", gConSql, adOpenStatic, adLockReadOnly
        Call ActualizaBotones(Me, 3, myPermiso)
        RefrescaDatos
    'End If
End If
End Sub

Sub InhibeCampoPrinc(bInhibe As Boolean, Optional yExcepción As Byte)
'inhibe los campos y datos iniciales, ya no se pueden modificar
OpcPAU(0).Enabled = Not bInhibe
OpcPAU(1).Enabled = Not bInhibe
txtExpediente.Locked = bInhibe
txtNuevoExp.Locked = bInhibe
cmdContinuar.Enabled = Not bInhibe
If yExcepción > 0 Then
    cmbCampo(0).Locked = False
    cmbCampo(1).Locked = False
    cmbCampo(2).Locked = False
Else
    cmbCampo(0).Locked = bInhibe
    cmbCampo(1).Locked = bInhibe
    cmbCampo(2).Locked = bInhibe
End If
End Sub

Private Sub VScroll_Change()
Dim l As Long
l = -VScroll.Value + 180
txtCampo(0).Top = l
txtCampo(1).Top = l
txtNombre.Top = l
txtExp.Top = l
comboLigas.ToolTipText = l
etiTexto(0).Top = l - 180
etiTexto(1).Top = l - 180
EtiNombre.Top = l - 180
lblExp.Top = l - 180
lblLigas.Top = l - 180
SSTab1.Top = l + 320
End Sub

