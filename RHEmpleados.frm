VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RHEmpleados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empleados"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12375
   Icon            =   "RHEmpleados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   12375
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageUsuarios"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   37
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Guardar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Borrar"
            Object.ToolTipText     =   "Borrar"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Actualizar"
            Object.ToolTipText     =   "Actualizar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Limpiar"
            Object.ToolTipText     =   "Limpiar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "Ir_a"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "Todos_Reg"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Primero"
            Object.ToolTipText     =   "Primer Registro"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Anterior"
            Object.ToolTipText     =   "Registro Anterior"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button23 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button24 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button25 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button26 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Siguiente Registro "
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button27 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Último"
            Object.ToolTipText     =   "Último Registro"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button28 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button29 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button30 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button31 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button32 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Emite informe"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button33 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Ayuda"
            Object.ToolTipText     =   "Ayuda en línea"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button34 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button35 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button36 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button37 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.TextBox txtNoReg 
         Height          =   375
         Left            =   3000
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   0
         Width           =   1425
      End
   End
   Begin VB.Frame Frame3 
      Height          =   510
      Left            =   7965
      TabIndex        =   36
      Top             =   495
      Visible         =   0   'False
      Width           =   2580
      Begin VB.OptionButton opcReg 
         Caption         =   "TODOS"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   4
         Top             =   225
         Width           =   900
      End
      Begin VB.OptionButton opcReg 
         Caption         =   "Últimos 10 Reg"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   225
         Width           =   1485
      End
   End
   Begin VB.TextBox txtcompleto 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   1485
      Locked          =   -1  'True
      TabIndex        =   2
      Tag             =   "c"
      Top             =   675
      Width           =   5460
   End
   Begin VB.TextBox txtCampo 
      DataField       =   "n_cveempleado"
      Height          =   285
      Index           =   0
      Left            =   90
      MaxLength       =   8
      TabIndex        =   0
      Tag             =   "n"
      ToolTipText     =   """Numero consecutivo de registro"""
      Top             =   675
      Width           =   1260
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   3240
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin TabDlg.SSTab SSTDatos 
      Height          =   6930
      Left            =   90
      TabIndex        =   33
      Top             =   990
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   12224
      _Version        =   393216
      Tabs            =   6
      TabHeight       =   520
      TabCaption(0)   =   "Datos Personales"
      TabPicture(0)   =   "RHEmpleados.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos Adicionales"
      TabPicture(1)   =   "RHEmpleados.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Movimientos Administrativos"
      TabPicture(2)   =   "RHEmpleados.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Evaluaciones"
      TabPicture(3)   =   "RHEmpleados.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Puestos Ocupados"
      TabPicture(4)   =   "RHEmpleados.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ListView3"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Eventos Grupales"
      TabPicture(5)   =   "RHEmpleados.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "ListView4"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6120
         Left            =   -74910
         TabIndex        =   105
         Top             =   630
         Width           =   11715
         Begin MSComctlLib.ListView ListView2 
            Height          =   5190
            Left            =   135
            TabIndex        =   61
            Top             =   225
            Width           =   11400
            _ExtentX        =   20108
            _ExtentY        =   9155
            View            =   3
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
               Text            =   "TipoEvaluación"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Resultado"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Fecha"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Responsable"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Observaciones"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Usuario Sistema"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Frame Frame9 
            Height          =   600
            Left            =   135
            TabIndex        =   106
            Top             =   5445
            Width           =   11370
            Begin VB.CommandButton cmdProceso 
               Caption         =   "&Consultar"
               Height          =   330
               Index           =   7
               Left            =   10125
               TabIndex        =   65
               Top             =   180
               Width           =   945
            End
            Begin VB.CommandButton cmdProceso 
               Caption         =   "B&orrar"
               Height          =   330
               Index           =   6
               Left            =   6810
               TabIndex        =   64
               Top             =   180
               Width           =   945
            End
            Begin VB.CommandButton cmdProceso 
               Caption         =   "&Modificar"
               Height          =   330
               Index           =   5
               Left            =   3495
               TabIndex        =   63
               Top             =   180
               Width           =   945
            End
            Begin VB.CommandButton cmdProceso 
               Caption         =   "&Agregar"
               Height          =   330
               Index           =   4
               Left            =   180
               TabIndex        =   62
               Top             =   180
               Width           =   945
            End
         End
      End
      Begin VB.Frame Frame5 
         Height          =   5850
         Left            =   -74910
         TabIndex        =   95
         Top             =   720
         Width           =   11775
         Begin VB.TextBox txtCampo 
            DataField       =   "Emp_monto1"
            Height          =   285
            Index           =   27
            Left            =   4725
            MaxLength       =   20
            TabIndex        =   50
            Tag             =   "m"
            ToolTipText     =   "Registro federal de contribuyentes"
            Top             =   5310
            Width           =   1725
         End
         Begin VB.TextBox txtCampo 
            DataField       =   "Emp_monto2"
            Height          =   285
            Index           =   28
            Left            =   6630
            MaxLength       =   20
            TabIndex        =   51
            Tag             =   "m"
            ToolTipText     =   "Clave unica de registro poblacional"
            Top             =   5310
            Width           =   1740
         End
         Begin VB.TextBox txtCampo 
            DataField       =   "Emp_fecha2"
            Height          =   285
            Index           =   26
            Left            =   2565
            MaxLength       =   20
            TabIndex        =   49
            Tag             =   "f"
            ToolTipText     =   "Fecha de Nacimiento"
            Top             =   5310
            Width           =   1905
         End
         Begin VB.TextBox txtCampo 
            DataField       =   "Emp_fecha1"
            Height          =   285
            Index           =   25
            Left            =   240
            MaxLength       =   20
            TabIndex        =   48
            Tag             =   "f"
            ToolTipText     =   "A.Materno"
            Top             =   5310
            Width           =   2145
         End
         Begin VB.TextBox txtCampo 
            DataField       =   "Emp_texto2"
            Height          =   285
            Index           =   24
            Left            =   6270
            MaxLength       =   20
            TabIndex        =   47
            Tag             =   "c"
            ToolTipText     =   "A.Paterno"
            Top             =   4500
            Width           =   5340
         End
         Begin VB.TextBox txtCampo 
            DataField       =   "Emp_texto1"
            Height          =   285
            Index           =   23
            Left            =   240
            MaxLength       =   30
            TabIndex        =   46
            Tag             =   "c"
            ToolTipText     =   "Nombre"
            Top             =   4500
            Width           =   5565
         End
         Begin VB.TextBox txtCampo 
            DataField       =   "Emp_entero1"
            Height          =   285
            Index           =   29
            Left            =   8700
            MaxLength       =   15
            TabIndex        =   52
            Tag             =   "n"
            ToolTipText     =   "Registro federal de contribuyentes"
            Top             =   5310
            Width           =   1545
         End
         Begin VB.TextBox txtCampo 
            DataField       =   "Emp_entero2"
            Height          =   285
            Index           =   30
            Left            =   10275
            MaxLength       =   15
            TabIndex        =   53
            Tag             =   "n"
            ToolTipText     =   "Clave unica de registro poblacional"
            Top             =   5310
            Width           =   1335
         End
         Begin VB.ComboBox cmbCampo 
            DataField       =   "n_cveRespBaja"
            Height          =   315
            Index           =   9
            ItemData        =   "RHEmpleados.frx":04EA
            Left            =   2790
            List            =   "RHEmpleados.frx":04EC
            Style           =   2  'Dropdown List
            TabIndex        =   45
            ToolTipText     =   "Responsable de Baja del Empleado"
            Top             =   3420
            Width           =   5415
         End
         Begin VB.TextBox txtCampo 
            DataField       =   "s_CLABE"
            Height          =   285
            Index           =   19
            Left            =   8235
            MaxLength       =   20
            TabIndex        =   41
            Tag             =   "c"
            ToolTipText     =   "CLABE"
            Top             =   990
            Width           =   3315
         End
         Begin VB.TextBox txtCampo 
            DataField       =   "f_Baja"
            Height          =   285
            Index           =   22
            Left            =   225
            MaxLength       =   30
            TabIndex        =   44
            Tag             =   "fh"
            ToolTipText     =   "Fecha de Baja"
            Top             =   3465
            Width           =   2190
         End
         Begin VB.TextBox txtCampo 
            DataField       =   "f_alta"
            Height          =   285
            Index           =   17
            Left            =   225
            MaxLength       =   30
            TabIndex        =   37
            Tag             =   "fh"
            ToolTipText     =   "Fecha de Alta del Empleado"
            Top             =   495
            Width           =   2640
         End
         Begin VB.TextBox txtCampo 
            DataField       =   "s_CtaBancaria"
            Height          =   285
            Index           =   18
            Left            =   4725
            MaxLength       =   20
            TabIndex        =   40
            Tag             =   "c"
            ToolTipText     =   "Número de Cuenta  Bancaria"
            Top             =   990
            Width           =   3315
         End
         Begin VB.TextBox txtCampo 
            DataField       =   "s_Referencias"
            Height          =   465
            Index           =   20
            Left            =   225
            MaxLength       =   20
            TabIndex        =   42
            Tag             =   "c"
            ToolTipText     =   "Referencias"
            Top             =   1620
            Width           =   11325
         End
         Begin VB.ComboBox cmbCampo 
            DataField       =   "n_cveBanco"
            Height          =   315
            Index           =   8
            ItemData        =   "RHEmpleados.frx":04EE
            Left            =   225
            List            =   "RHEmpleados.frx":04F0
            Style           =   2  'Dropdown List
            TabIndex        =   39
            ToolTipText     =   "Banco"
            Top             =   990
            Width           =   3765
         End
         Begin VB.ComboBox cmbCampo 
            DataField       =   "n_cveRespAlta"
            Height          =   315
            Index           =   7
            ItemData        =   "RHEmpleados.frx":04F2
            Left            =   3645
            List            =   "RHEmpleados.frx":04F4
            Style           =   2  'Dropdown List
            TabIndex        =   38
            ToolTipText     =   "Responsable de Alta del Empleado"
            Top             =   450
            Width           =   4830
         End
         Begin VB.TextBox txtCampo 
            DataField       =   "s_observacionesEmp"
            Height          =   645
            Index           =   21
            Left            =   225
            MaxLength       =   60
            TabIndex        =   43
            Tag             =   "c"
            ToolTipText     =   "Observaciones, comentarios o datos adicionales"
            Top             =   2430
            Width           =   11370
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Monto1:"
            Height          =   195
            Index           =   27
            Left            =   4785
            TabIndex        =   116
            Top             =   5085
            Width           =   585
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Monto2 :"
            Height          =   195
            Index           =   28
            Left            =   6660
            TabIndex        =   115
            Top             =   5085
            Width           =   630
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Fecha2:"
            Height          =   195
            Index           =   29
            Left            =   2520
            TabIndex        =   114
            Top             =   5040
            Width           =   585
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Fecha1:"
            Height          =   195
            Index           =   24
            Left            =   240
            TabIndex        =   113
            Top             =   5040
            Width           =   585
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Texto2:"
            Height          =   195
            Index           =   23
            Left            =   6255
            TabIndex        =   112
            Top             =   4140
            Width           =   540
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Texto1:"
            Height          =   195
            Index           =   22
            Left            =   270
            TabIndex        =   111
            Top             =   4140
            Width           =   540
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Entero1:"
            Height          =   195
            Index           =   21
            Left            =   8775
            TabIndex        =   110
            Top             =   5085
            Width           =   600
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Entero2 :"
            Height          =   195
            Index           =   20
            Left            =   10275
            TabIndex        =   109
            Top             =   5085
            Width           =   645
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   0
            X2              =   11745
            Y1              =   3960
            Y2              =   3960
         End
         Begin VB.Label etiCampo 
            AutoSize        =   -1  'True
            Caption         =   "Responsable Baja:"
            Height          =   195
            Index           =   9
            Left            =   2790
            TabIndex        =   104
            Top             =   3150
            Width           =   5385
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "CLABE:"
            Height          =   195
            Index           =   19
            Left            =   8235
            TabIndex        =   103
            Top             =   765
            Width           =   555
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Baja:"
            Height          =   195
            Index           =   6
            Left            =   255
            TabIndex        =   102
            Top             =   3195
            Width           =   1080
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Alta:"
            Height          =   195
            Index           =   32
            Left            =   225
            TabIndex        =   101
            Top             =   225
            Width           =   1035
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Cta.Bancaria:"
            Height          =   195
            Index           =   31
            Left            =   4725
            TabIndex        =   100
            Top             =   765
            Width           =   960
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Referencias:"
            Height          =   195
            Index           =   30
            Left            =   225
            TabIndex        =   99
            Top             =   1350
            Width           =   900
         End
         Begin VB.Label etiCampo 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Index           =   8
            Left            =   225
            TabIndex        =   98
            Top             =   765
            Width           =   3705
         End
         Begin VB.Label etiCampo 
            AutoSize        =   -1  'True
            Caption         =   "Responsable Alta:"
            Height          =   195
            Index           =   7
            Left            =   3645
            TabIndex        =   97
            Top             =   225
            Width           =   4800
         End
         Begin VB.Label Label 
            Caption         =   "Observaciones:"
            Height          =   195
            Index           =   18
            Left            =   225
            TabIndex        =   96
            Top             =   2160
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6285
         Left            =   45
         TabIndex        =   68
         Top             =   540
         Width           =   11820
         Begin VB.CommandButton cmdPersonal 
            Caption         =   "Actualizar Datos Personales"
            Height          =   420
            Index           =   1
            Left            =   3690
            TabIndex        =   35
            Top             =   5760
            Width           =   2265
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H80000016&
            DataField       =   "s_nombre"
            Height          =   285
            Index           =   1
            Left            =   225
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   5
            Tag             =   "c"
            ToolTipText     =   "Nombre"
            Top             =   450
            Width           =   2370
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H80000016&
            DataField       =   "s_paterno"
            Height          =   285
            Index           =   2
            Left            =   2745
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   6
            Tag             =   "c"
            ToolTipText     =   "A.Paterno"
            Top             =   450
            Width           =   2235
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H80000016&
            DataField       =   "s_materno"
            Height          =   285
            Index           =   3
            Left            =   5130
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   7
            Tag             =   "c"
            ToolTipText     =   "A.Materno"
            Top             =   450
            Width           =   2055
         End
         Begin VB.ComboBox cmbCampo 
            BackColor       =   &H80000016&
            DataField       =   "n_cveedocivil"
            Height          =   315
            Index           =   1
            ItemData        =   "RHEmpleados.frx":04F6
            Left            =   225
            List            =   "RHEmpleados.frx":050C
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   16
            ToolTipText     =   "Estado civil al momento del registro"
            Top             =   1710
            Width           =   4020
         End
         Begin VB.ComboBox cmbCampo 
            BackColor       =   &H80000016&
            DataField       =   "n_cveestudio"
            Height          =   315
            Index           =   2
            ItemData        =   "RHEmpleados.frx":055A
            Left            =   225
            List            =   "RHEmpleados.frx":055C
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   19
            ToolTipText     =   """Grado de Estudios"""
            Top             =   2340
            Width           =   3495
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H80000016&
            DataField       =   "f_fecha_nac"
            Height          =   285
            Index           =   4
            Left            =   7335
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "f"
            ToolTipText     =   "Fecha de Nacimiento"
            Top             =   450
            Width           =   1500
         End
         Begin VB.ComboBox cmbCampo 
            BackColor       =   &H80000016&
            DataField       =   "n_cveEstadoNac"
            Height          =   315
            Index           =   0
            ItemData        =   "RHEmpleados.frx":055E
            Left            =   2160
            List            =   "RHEmpleados.frx":0560
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   11
            ToolTipText     =   "Estado de Nacimiento"
            Top             =   1035
            Width           =   3210
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H80000016&
            DataField       =   "s_curp"
            Height          =   285
            Index           =   6
            Left            =   4590
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   17
            Tag             =   "c"
            ToolTipText     =   "Clave unica de registro poblacional"
            Top             =   1710
            Width           =   3000
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H80000016&
            DataField       =   "s_rfc"
            Height          =   285
            Index           =   5
            Left            =   9045
            Locked          =   -1  'True
            MaxLength       =   13
            TabIndex        =   14
            Tag             =   "c"
            ToolTipText     =   "Registro federal de contribuyentes"
            Top             =   1080
            Width           =   2445
         End
         Begin VB.TextBox txtedad 
            BackColor       =   &H80000016&
            Height          =   285
            Left            =   9180
            Locked          =   -1  'True
            TabIndex        =   9
            ToolTipText     =   "Edad: Años Meses"
            Top             =   450
            Width           =   2310
         End
         Begin VB.Frame Frame2 
            Caption         =   "Nacionalidad"
            Enabled         =   0   'False
            Height          =   555
            Left            =   180
            TabIndex        =   70
            Top             =   855
            Width           =   1680
            Begin VB.CheckBox chkNacionalidad 
               BackColor       =   &H80000016&
               Caption         =   "Mexicana"
               DataField       =   "n_nacionalidad"
               Enabled         =   0   'False
               Height          =   285
               Left            =   135
               TabIndex        =   10
               ToolTipText     =   "Nacionalidad Mexicana/Extranjero"
               Top             =   225
               Width           =   1095
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Sexo"
            Enabled         =   0   'False
            Height          =   510
            Left            =   5670
            TabIndex        =   69
            Top             =   855
            Width           =   2985
            Begin VB.OptionButton OpcSexo 
               BackColor       =   &H80000016&
               Caption         =   "Masculino"
               Enabled         =   0   'False
               Height          =   195
               Index           =   0
               Left            =   405
               TabIndex        =   12
               Top             =   225
               Width           =   1140
            End
            Begin VB.OptionButton OpcSexo 
               BackColor       =   &H80000016&
               Caption         =   "Femenino"
               Enabled         =   0   'False
               Height          =   195
               Index           =   1
               Left            =   1665
               TabIndex        =   13
               Top             =   225
               Width           =   1050
            End
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H80000016&
            DataField       =   "s_imss"
            Height          =   285
            Index           =   7
            Left            =   7695
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   18
            Tag             =   "n"
            ToolTipText     =   "Número de Seguridad Socialdel IMSS"
            Top             =   1710
            Width           =   3810
         End
         Begin VB.ComboBox cmbCampo 
            BackColor       =   &H80000016&
            DataField       =   "n_cveestado"
            Height          =   315
            Index           =   5
            Left            =   225
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   25
            ToolTipText     =   """Estado"""
            Top             =   3645
            Width           =   2175
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H80000016&
            DataField       =   "s_codigo_postal"
            Height          =   285
            Index           =   11
            Left            =   5760
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   27
            Tag             =   "c"
            ToolTipText     =   "Código Postal"
            Top             =   3645
            Width           =   1245
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H80000016&
            DataField       =   "s_colonia"
            Height          =   285
            Index           =   9
            Left            =   4095
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   23
            Tag             =   "c"
            ToolTipText     =   "Colonia"
            Top             =   3060
            Width           =   3135
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H80000016&
            DataField       =   "s_domicilio"
            Height          =   285
            Index           =   8
            Left            =   225
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   22
            Tag             =   "c"
            ToolTipText     =   "Dirección: Calle y Número"
            Top             =   3060
            Width           =   3750
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H80000016&
            DataField       =   "s_telefono"
            Height          =   285
            Index           =   12
            Left            =   7335
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   28
            Tag             =   "c"
            ToolTipText     =   "Teléfono particular"
            Top             =   3645
            Width           =   2385
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H80000016&
            DataField       =   "s_celular"
            Height          =   285
            Index           =   13
            Left            =   225
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   29
            Tag             =   "c"
            ToolTipText     =   "Teléfono Celular"
            Top             =   4275
            Width           =   2625
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H80000016&
            DataField       =   "s_correo"
            Height          =   285
            Index           =   14
            Left            =   3645
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   30
            Tag             =   "c"
            ToolTipText     =   "Correo electrónico (E-Mail)"
            Top             =   4275
            Width           =   3990
         End
         Begin VB.ComboBox cmbCampo 
            BackColor       =   &H80000016&
            DataField       =   "n_cvemunicipio"
            Height          =   315
            Index           =   6
            Left            =   2565
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   26
            ToolTipText     =   """Estado"""
            Top             =   3645
            Width           =   3120
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H80000016&
            DataField       =   "f_fecha_baja"
            Height          =   285
            Index           =   15
            Left            =   7875
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   31
            Tag             =   "f"
            ToolTipText     =   "Fecha de Baja"
            Top             =   4275
            Width           =   1785
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H80000016&
            DataField       =   "s_observaciones"
            Height          =   780
            Index           =   16
            Left            =   225
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   32
            Tag             =   "c"
            ToolTipText     =   "Observaciones, comentarios o datos adicionales"
            Top             =   4905
            Width           =   9525
         End
         Begin VB.ComboBox cmbCampo 
            BackColor       =   &H80000016&
            DataField       =   "n_cvereclutamiento"
            Height          =   315
            Index           =   3
            ItemData        =   "RHEmpleados.frx":0562
            Left            =   4050
            List            =   "RHEmpleados.frx":0564
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   20
            ToolTipText     =   """Grado de Estudios"""
            Top             =   2340
            Width           =   3540
         End
         Begin VB.ComboBox cmbCampo 
            BackColor       =   &H80000016&
            DataField       =   "n_cveTipoPuesto"
            Height          =   315
            Index           =   4
            ItemData        =   "RHEmpleados.frx":0566
            Left            =   7965
            List            =   "RHEmpleados.frx":0568
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   21
            ToolTipText     =   """Grado de Estudios"""
            Top             =   2340
            Width           =   3540
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H80000016&
            DataField       =   "s_entre_calles"
            Height          =   285
            Index           =   10
            Left            =   7380
            Locked          =   -1  'True
            TabIndex        =   24
            Tag             =   "c"
            ToolTipText     =   "Entre las calles de .. y .."
            Top             =   3060
            Width           =   4245
         End
         Begin VB.CommandButton cmdPersonal 
            Caption         =   "Asignar Persona Contratada"
            Height          =   375
            Index           =   0
            Left            =   450
            TabIndex        =   34
            Top             =   5760
            Width           =   2265
         End
         Begin VB.Image PicFoto 
            Height          =   2130
            Left            =   9900
            Stretch         =   -1  'True
            Top             =   3645
            Width           =   1770
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Nombre(s):"
            Height          =   195
            Index           =   1
            Left            =   210
            TabIndex        =   94
            Top             =   180
            Width           =   765
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Apellido Paterno:"
            Height          =   195
            Index           =   2
            Left            =   2790
            TabIndex        =   93
            Top             =   180
            Width           =   1200
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Apellido Materno:"
            Height          =   195
            Index           =   3
            Left            =   5175
            TabIndex        =   92
            Top             =   180
            Width           =   1230
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "F.Nacimiento:"
            Height          =   195
            Index           =   4
            Left            =   7335
            TabIndex        =   91
            Top             =   180
            Width           =   975
         End
         Begin VB.Label etiCampo 
            AutoSize        =   -1  'True
            Caption         =   "Estudios:"
            Height          =   195
            Index           =   2
            Left            =   225
            TabIndex        =   90
            Top             =   2115
            Width           =   3525
         End
         Begin VB.Label etiCampo 
            AutoSize        =   -1  'True
            Caption         =   "Estado Civil:"
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   89
            Top             =   1440
            Width           =   3975
         End
         Begin VB.Label etiCampo 
            AutoSize        =   -1  'True
            Caption         =   "Lugar de Nacimiento(ESTADO):"
            Height          =   195
            Index           =   0
            Left            =   2160
            TabIndex        =   88
            Top             =   810
            Width           =   2265
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "CURP :"
            Height          =   195
            Index           =   7
            Left            =   4590
            TabIndex        =   87
            Top             =   1440
            Width           =   540
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "R.F.C. :"
            Height          =   195
            Index           =   5
            Left            =   9045
            TabIndex        =   86
            Top             =   855
            Width           =   540
         End
         Begin VB.Label Label1 
            Caption         =   "Edad:"
            Height          =   240
            Left            =   9180
            TabIndex        =   85
            Top             =   180
            Width           =   780
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "IMSS :"
            Height          =   195
            Index           =   8
            Left            =   7695
            TabIndex        =   84
            Top             =   1440
            Width           =   480
         End
         Begin VB.Label Label 
            Caption         =   "C.P. :"
            Height          =   255
            Index           =   12
            Left            =   5760
            TabIndex        =   83
            Top             =   3420
            Width           =   495
         End
         Begin VB.Label etiCampo 
            Caption         =   "Estado:"
            Height          =   255
            Index           =   5
            Left            =   225
            TabIndex        =   82
            Top             =   3420
            Width           =   2160
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Colonia:"
            Height          =   195
            Index           =   10
            Left            =   4095
            TabIndex        =   81
            Top             =   2835
            Width           =   570
         End
         Begin VB.Label Label 
            Caption         =   "Dirección:"
            Height          =   255
            Index           =   9
            Left            =   225
            TabIndex        =   80
            Top             =   2835
            Width           =   855
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono particular:"
            Height          =   195
            Index           =   13
            Left            =   7335
            TabIndex        =   79
            Top             =   3420
            Width           =   1365
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Correo electronico:"
            Height          =   195
            Index           =   15
            Left            =   3645
            TabIndex        =   78
            Top             =   4095
            Width           =   1335
         End
         Begin VB.Label etiCampo 
            Caption         =   "Del./Mpio:"
            Height          =   255
            Index           =   6
            Left            =   2565
            TabIndex        =   77
            Top             =   3420
            Width           =   3105
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Celular:"
            Height          =   195
            Index           =   14
            Left            =   225
            TabIndex        =   76
            Top             =   4095
            Width           =   525
         End
         Begin VB.Label Label 
            Caption         =   "Fecha Baja:"
            Height          =   195
            Index           =   16
            Left            =   7875
            TabIndex        =   75
            Top             =   4095
            Width           =   1455
         End
         Begin VB.Label Label 
            Caption         =   "Observaciones:"
            Height          =   195
            Index           =   17
            Left            =   225
            TabIndex        =   74
            Top             =   4680
            Width           =   1455
         End
         Begin VB.Label etiCampo 
            AutoSize        =   -1  'True
            Caption         =   "Reclutamiento:"
            Height          =   195
            Index           =   3
            Left            =   4050
            TabIndex        =   73
            Top             =   2115
            Width           =   3540
         End
         Begin VB.Label etiCampo 
            AutoSize        =   -1  'True
            Caption         =   "Puesto Solicitado:"
            Height          =   195
            Index           =   4
            Left            =   7965
            TabIndex        =   72
            Top             =   2115
            Width           =   3570
         End
         Begin VB.Label Label 
            Caption         =   "Entre Calles:"
            Height          =   195
            Index           =   11
            Left            =   7380
            TabIndex        =   71
            Top             =   2835
            Width           =   4245
         End
      End
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5970
         Left            =   -74910
         TabIndex        =   66
         Top             =   675
         Width           =   11715
         Begin VB.Frame Frame8 
            Height          =   600
            Left            =   180
            TabIndex        =   67
            Top             =   5130
            Width           =   11415
            Begin VB.CommandButton cmdProceso 
               Caption         =   "&Agregar"
               Height          =   330
               Index           =   0
               Left            =   180
               TabIndex        =   55
               Top             =   180
               Width           =   945
            End
            Begin VB.CommandButton cmdProceso 
               Caption         =   "&Modificar"
               Height          =   330
               Index           =   1
               Left            =   3570
               TabIndex        =   56
               Top             =   180
               Width           =   945
            End
            Begin VB.CommandButton cmdProceso 
               Caption         =   "B&orrar"
               Height          =   330
               Index           =   2
               Left            =   6960
               TabIndex        =   57
               Top             =   180
               Width           =   945
            End
            Begin VB.CommandButton cmdProceso 
               Caption         =   "&Consultar"
               Height          =   330
               Index           =   3
               Left            =   10350
               TabIndex        =   58
               Top             =   180
               Width           =   945
            End
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   4785
            Left            =   90
            TabIndex        =   54
            Top             =   270
            Visible         =   0   'False
            Width           =   11400
            _ExtentX        =   20108
            _ExtentY        =   8440
            View            =   3
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
               Text            =   "Movimiento"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Fecha"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Responsable"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Observaciones"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Usuario Sistema"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   5775
         Left            =   -74820
         TabIndex        =   107
         Top             =   900
         Width           =   11670
         _ExtentX        =   20585
         _ExtentY        =   10186
         View            =   3
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
            Text            =   "Area"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Departamento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tipo Puesto"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fecha"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Observaciones"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   5865
         Left            =   -74820
         TabIndex        =   108
         Top             =   810
         Width           =   11670
         _ExtentX        =   20585
         _ExtentY        =   10345
         View            =   3
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Evento"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Lugar"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Expositor(es)"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Inicio"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Término"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Observaciones"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Empleado:"
      Height          =   240
      Index           =   0
      Left            =   1530
      TabIndex        =   15
      Top             =   480
      Width           =   735
   End
   Begin ComctlLib.ImageList ImageUsuarios 
      Left            =   3840
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RHEmpleados.frx":056A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RHEmpleados.frx":0674
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RHEmpleados.frx":077E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RHEmpleados.frx":0888
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RHEmpleados.frx":0992
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RHEmpleados.frx":0A9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RHEmpleados.frx":0BA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RHEmpleados.frx":0CB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RHEmpleados.frx":0DBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RHEmpleados.frx":0EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RHEmpleados.frx":0FCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RHEmpleados.frx":10D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RHEmpleados.frx":1A8A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Id Empleado:"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   480
      Width           =   930
   End
End
Attribute VB_Name = "RHEmpleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents AdorsPrin As ADODB.Recordset, adorsCata As ADODB.Recordset
Attribute AdorsPrin.VB_VarHelpID = -1
Dim msConsulta As String, msOrden As String, msPrin As String, msConsultaP As String
Dim mbLimpia As Boolean, mbCambio As Boolean, mbRefresca As Boolean, mbInicio  As Boolean, mbNoPreg As Boolean, mlAnt As Long
Dim mlEmpleado As Long  'clave principal del Empleado
Dim mlMovimiento As Long  'clave de la seguimiento de la persona que originó la contratación
Dim mlPersona As Long  'clave de la persona
Dim mlBuscaEmp As Long   'Busca un empleado especifica cuando se   invoca desde otro formulario
Dim myPermiso As Byte, myTabla As Byte, myPermisoRep As Byte

'Actualiza datos pormovimiento deregistro
Private Sub AdorsPrin_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo Salir:
Static iCom
If iCom = 1 Then
    iCom = 0
    Exit Sub
End If
iCom = 1
If AdorsPrin.RecordCount > 0 Then
    If AdorsPrin.Bookmark > 0 Then
        RefrescaDatos
    End If
End If
iCom = 0
Exit Sub
Salir:
iCom = 0
End Sub

Private Sub chkNacionalidad_KeyPress(KeyAscii As Integer)
Dim i As Long, i1 As Long
If KeyAscii = 13 Then
    i1 = chkNacionalidad.TabIndex
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

'Selección del combo, filtro de los municipios
Private Sub cmbCampo_Click(Index As Integer)
If Index = 5 And cmbCampo(Index).ListIndex >= 0 Then
    LlenaCombo cmbCampo(6), "select n_cvemunicipio,s_municipio from c_RHmunicipios where n_cveestado=" & cmbCampo(Index).ItemData(cmbCampo(Index).ListIndex), "", True
End If
If Index <= 6 Then Exit Sub
If Not mbCambio And Not mbRefresca Then mbCambio = True
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

Private Sub cmdPersonal_Click(Index As Integer)
Dim yPermiso As Byte
yPermiso = Val(Mid(gsPermisos, 8, 1))  'Permiso de Asignar
If yPermiso <= 1 Then
    MsgBox "Solo tiene permisos de consulta", vbInformation + vbOKOnly, "Validación"
    cmdPersonal(0).Enabled = False
    cmdPersonal(1).Enabled = False
    Exit Sub
End If
If Index = 1 Then
    If mlPersona <= 0 Then
        MsgBox "Debe selecionar un registro que tenga asignación de persona", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
    gs = "persona-->" & mlPersona
    RHPersonal.Show vbModal
    Exit Sub
End If

If mlPersona > 0 Then
'    If InStr("2468", yPermiso) = 0 Then
'        MsgBox "No cuenta con permiso para actualizar en este módulo", vbInformation + vbOKOnly, "Validación"
'        cmdPersonal(Index).Enabled = False
'        Exit Sub
'    End If
    Call MsgBox(" No se permite asignar a otra persona diferente", vbOKOnly + vbInformation, "Validación")
    Exit Sub
Else
    If Index = 0 Then  'agregar
        If InStr("2578", yPermiso) = 0 Then
            MsgBox "No cuenta con permiso para agregar en este módulo", vbInformation + vbOKOnly, "Validación"
            cmdPersonal(Index).Enabled = False
            Exit Sub
        End If
    End If
End If
gs = "ArbolVarios-->select case when r.n_cvereclutamiento is null then -1 else r.n_cvereclutamiento end,ps.n_cveseguimiento,case when r.s_reclutamiento is null then 'Fuente Reclutamiento no Especifoicado' else r.s_reclutamiento end,p.s_nombre+rtrim(' '+p.s_paterno)+rtrim(' '+p.s_materno) as Personal " + _
     "from t_rhpersonalseg ps " + _
     "inner join t_rhpersonal p on ps.n_cvepersona=p.n_cvepersona " + _
     "left join c_rhreclutamiento r on p.n_cvereclutamiento=r.n_cvereclutamiento " + _
     "left join t_rhempleados e on ps.n_cveseguimiento=e.n_cveseguimiento " + _
     "Where ps.n_cveperseg = 4 And ps.n_cveevento = 1 And ps.n_resultado = 1 And e.n_cveempleado Is Null order by 3,4"
With RHSelProceso
    .Caption = "Seleccione la persona reclutado ya contratado"
    .Show vbModal
End With
If Val(gs) > 0 Then
    If mlEmpleado > 0 Then
        mlMovimiento = Val(gs)
        
        '''Nueva información
        If F_Transacción("update t_rhEmpleados set n_cveseguimiento=" & mlMovimiento & " where n_cveempleado=" & mlEmpleado) Then
            RefrescaDatosAsignación
            If Len(Trim(txtCampo(0).Text)) = 0 Then
                txtCampo(0).Text = "Nuevo"
            End If
        Else
            If gConSql.Errors.Count > 0 Then
                MsgBox "No se realizó el cambio. Descripción del problema: " & gConSql.Errors(0).Description, vbOKOnly + vbInformation, "Problemas al guardar la información"
            End If
        End If
    Else
        mlMovimiento = Val(gs)
        RefrescaDatosAsignación
        If Len(Trim(txtCampo(0).Text)) = 0 Then
            txtCampo(0).Text = "Nuevo"
        End If
    End If
End If
End Sub

'Botón para consulta, actualización o agregar movimientos administrativos
Private Sub cmdProceso_Click(Index As Integer)
Dim iRows As Integer, i As Integer
Dim adors As New ADODB.Recordset
If mlEmpleado = 0 Then Exit Sub
i = Val(Mid(gsPermisos, 3, 1))
If Index < 2 Or Index = 3 Then
    gi1 = mlEmpleado
    gi = IIf(Index = 0, 3, IIf(Index = 1, 2, 1))
    If Index = 0 Then  'agregar
        If InStr("2578", i) = 0 Then
            MsgBox "No cuenta con permiso para agregar en este módulo", vbInformation + vbOKOnly, "Validación"
            cmdProceso(Index).Enabled = False
            Exit Sub
        End If
    ElseIf Index = 1 Then  'actualizar
        If InStr("2468", i) = 0 Then
            MsgBox "No cuenta con permiso para actualizar en este módulo", vbInformation + vbOKOnly, "Validación"
            cmdProceso(Index).Enabled = False
            Exit Sub
        End If
    End If
    If gi = 3 Then
        'adors.Open "select min(n_cvemovimiento) from t_rhEmpleadosMov where n_cveempleado=" & mlEmpleado , gConSql, adOpenStatic, adLockReadOnly
        'If adors(0) > 0 Then
        '    gi2 = adors(0)
        'Else
            gi2 = 0
        'End If
    Else
        If ListView1.ListItems.Count = 0 Then Exit Sub
        If ListView1.SelectedItem.Index = 0 Then Exit Sub
        gi2 = Val(ListView1.ListItems(ListView1.SelectedItem.Index).Tag)
    End If
    gi3 = mlAnt
    RHEmpleadosMov.Show vbModal
    If gi = -99 Then  'Significa Ok.. se realizó cambio hay que refrescar datos especiales
        ActualizaCampoEsp
    End If
ElseIf Index = 2 Then
    If ListView1.ListItems.Count <= 0 Then
        cmdProceso(2).Enabled = False
        Exit Sub
    End If
    If InStr("2367", i) = 0 Then
        MsgBox "No cuenta con permiso para borrar en este módulo", vbInformation + vbOKOnly, "Validación"
        cmdProceso(2).Enabled = False
        Exit Sub
    End If
    If MsgBox("¿Está seguro de borrar la actividad?", vbYesNo + vbQuestion, "Confirmación") = vbYes Then
        If ValidaBorradoAct(Val(ListView1.SelectedItem.Tag)) Then
            gConSql.Execute "delete from t_rhEmpleadosMov where n_cvemovimiento=" & Val(ListView1.SelectedItem.Tag), iRows
            If iRows > 0 Then
                Call MsgBox("Se eliminó correctamente el registro", vbInformation + vbOKOnly, "")
                ActualizaCampoEsp
            End If
        End If
    End If
ElseIf Index > 3 And Index < 6 Or Index = 7 Then
    gi1 = mlEmpleado
    gi = IIf(Index = 4, 3, IIf(Index = 5, 2, 1))
    If Index = 0 Then  'agregar
        If InStr("2578", i) = 0 Then
            MsgBox "No cuenta con permiso para agregar en este módulo", vbInformation + vbOKOnly, "Validación"
            cmdProceso(Index).Enabled = False
            Exit Sub
        End If
    ElseIf Index = 1 Then  'actualizar
        If InStr("2468", i) = 0 Then
            MsgBox "No cuenta con permiso para actualizar en este módulo", vbInformation + vbOKOnly, "Validación"
            cmdProceso(Index).Enabled = False
            Exit Sub
        End If
    End If
    If gi = 3 Then
        'adors.Open "select min(n_cvemovimiento) from t_rhEmpleadosMov where n_cveempleado=" & mlEmpleado , gConSql, adOpenStatic, adLockReadOnly
        'If adors(0) > 0 Then
        '    gi2 = adors(0)
        'Else
            gi2 = 0
        'End If
    Else
        If ListView2.ListItems.Count = 0 And Index <> 4 Then Exit Sub
        If ListView2.SelectedItem.Index = 0 Then Exit Sub
        gi2 = Val(ListView2.ListItems(ListView2.SelectedItem.Index).Tag)
    End If
    gi3 = mlAnt
    RHEmpleadosEva.Show vbModal
    If gi = -99 Then  'Significa Ok.. se realizó cambio hay que refrescar datos especiales
        ActualizaCampoEsp
    End If
ElseIf Index = 6 Then
    If ListView1.ListItems.Count <= 0 Then
        cmdProceso(6).Enabled = False
        Exit Sub
    End If
    If InStr("2367", i) = 0 Then
        MsgBox "No cuenta con permiso para borrar en este módulo", vbInformation + vbOKOnly, "Validación"
        cmdProceso(2).Enabled = False
        Exit Sub
    End If
    If MsgBox("¿Está seguro de borrar la evaluación?", vbYesNo + vbQuestion, "Confirmación") = vbYes Then
        If ValidaBorradoAct(Val(ListView2.SelectedItem.Tag)) Then
            gConSql.Execute "delete from t_rhEmpleadosEva where n_cveEvaluacion=" & Val(ListView2.SelectedItem.Tag), iRows
            If iRows > 0 Then
                Call MsgBox("Se eliminó correctamente el registro", vbInformation + vbOKOnly, "")
                ActualizaCampoEsp
            Else
                If gConSql.Errors.Count > 0 Then
                    Call MsgBox("No pudo ser eliminado el Registro: " & gConSql.Errors(0).Description, vbOKOnly + vbInformation, "")
                Else
                    Call MsgBox("No pudo ser eliminado el Registro ", vbOKOnly + vbInformation, "")
                End If
            End If
        End If
    End If
End If
End Sub


Private Sub cmdProceso_KeyPress(Index As Integer, KeyAscii As Integer)
Dim i As Long, i1 As Long
If KeyAscii = 13 Then
    i1 = cmdProceso(Index).TabIndex
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

Private Sub Form_Activate()
Dim s As String
s = "empleado-->"
If gs Like s & "*" Then
    gs = LCase(gs)
    If Val(Mid(gs, InStr(gs, s) + Len(s))) > 0 Then
        mlBuscaEmp = Val(Mid(gs, InStr(gs, s) + Len(s)))
        Dim botón As Object
        Set botón = Me.Toolbar.Buttons(5)
        Call Toolbar_ButtonClick(botón)
    End If
End If
If mlMovimiento > 0 Then
    RefrescaDatosAsignación
End If
End Sub

'Inicio del formulario al cargarse
Private Sub Form_Load()
Dim sCampos As String, i As Integer, sOrden As String
myPermiso = Val(Mid(gsPermisos, 5, 1))
myPermisoRep = Val(Mid(gsPermisosRep, 5, 1))
myTabla = 6
Set AdorsPrin = New ADODB.Recordset
Set adorsCata = New ADODB.Recordset
For i = txtCampo.LBound To txtCampo.UBound
    sCampos = sCampos & txtCampo(i).DataField & ","
    If i = 0 Then msPrin = txtCampo(i).DataField
    If i = 1 Or i = 2 Or i = 3 Then sOrden = sOrden & txtCampo(i).DataField & ","
Next
sOrden = Mid(sOrden, 1, Len(sOrden) - 1)
For i = cmbCampo.LBound To cmbCampo.UBound
    sCampos = sCampos & cmbCampo(i).DataField & ","
Next
'otros campos tipo opción y casilla
sCampos = sCampos & "n_sexo,n_nacionalidad"
msConsulta = "select " & sCampos & ",ps.n_cveseguimiento,ps.n_cvepersona from t_rhEmpleados e left join t_rhPersonalSeg ps on e.n_cveseguimiento=ps.n_cveseguimiento left join t_rhpersonal p on ps.n_cvepersona=p.n_cvepersona"
msOrden = " order by n_cveempleado"
adorsCata.Open msConsulta & msOrden, gConSql, adOpenStatic, adLockReadOnly
msConsultaP = "select n_cveempleado from t_rhEmpleados where n_cveempleado>(select max(n_cveempleado) from t_rhEmpleados)-10"
LlenaCombo cmbCampo(0), "select n_cveestado,s_estado from c_RHestados order by 2", "", True
LlenaCombo cmbCampo(1), "select n_cveedocivil,s_estado_civil from c_RHEstadoCivil order by 2", "", True
LlenaCombo cmbCampo(2), "select n_cveestudio,s_estudio from c_RHEstudios order by 2", "", True
LlenaCombo cmbCampo(3), "select n_cvereclutamiento,s_reclutamiento from c_RHreclutamiento order by 2", "", True
LlenaCombo cmbCampo(4), "select n_cveTipopuesto,s_Tipopuesto from c_RHTipopuestos order by 2", "", True
LlenaCombo cmbCampo(5), "select n_cveestado,s_estado from c_RHestados order by 2", "", True
LlenaCombo cmbCampo(7), "select dbo.f_empleado_idper(n_cveempleado) as id ,dbo.f_responsable(dbo.f_empleado_idper(n_cveempleado)) as nombre from t_rhpuestos where n_cvedepartamento=115 and n_cveempleado is not null", "", True
LlenaCombo cmbCampo(8), "select n_cvebanco,s_banco from c_RHbancos order by 2", "", True
LlenaCombo cmbCampo(9), "select dbo.f_empleado_idper(n_cveempleado) as id ,dbo.f_responsable(dbo.f_empleado_idper(n_cveempleado)) as nombre from t_rhpuestos where n_cvedepartamento=115 and n_cveempleado is not null", "", True
cmbCampo(0).Tag = "select n_cveestado,s_estado from c_RHestados where n_cveestado="
cmbCampo(1).Tag = "select n_cveedocivil,s_estado_civil from c_RHEstadoCivil where n_cveedocivil="
cmbCampo(2).Tag = "select n_cveestudio,s_estudio from c_RHEstudios where n_cveestudio="
cmbCampo(3).Tag = "select n_cvereclutamiento,s_reclutamiento from c_RHreclutamiento where n_cvereclutamiento="
cmbCampo(4).Tag = "select n_cveTipopuesto,s_Tipopuesto from c_RHTipopuestos where n_cveTipopuesto="
cmbCampo(5).Tag = "select n_cveestado,s_estado from c_RHestados order by 2"
cmbCampo(7).Tag = "select dbo.f_empleado_idper(n_cveempleado) as id,dbo.f_responsable(dbo.f_empleado_idper(n_cveempleado)) as nombre from t_rhpuestos where dbo.f_empleado_idper(n_cveempleado)="
cmbCampo(8).Tag = "select n_cvebanco,s_banco from c_RHbancos where n_cvebanco="
cmbCampo(9).Tag = "select dbo.f_empleado_idper(n_cveempleado) as id,dbo.f_responsable(dbo.f_empleado_idper(n_cveempleado)) as nombre from t_rhpuestos where dbo.f_empleado_idper(n_cveempleado)="
AdorsPrin.Open msConsultaP, gConSql, adOpenStatic, adLockReadOnly
For i = txtCampo.LBound + 1 To txtCampo.UBound
    txtCampo(i).MaxLength = IIf(adorsCata.Fields(i).DefinedSize > 500, 500, adorsCata.Fields(i).DefinedSize)
Next
mbInicio = True
Dim botón As Object
Set botón = Me.Toolbar.Buttons(4)
Call Toolbar_ButtonClick(botón)
'Call ActualizaBotones(Me, 2, myPermiso)
GuardaBitácora gs_usuario, myTabla, -1, 4
mbCambio = False
mbLimpia = False
'indica que debe limpiar cuando se activa la ventana

End Sub

'Antes de cerrar el formulario valida lo pendiente
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If mbLimpia Or mbCambio Then
    If MsgBox("¿Desea salir del módulo de " & Me.Caption & ", sin guardarlos cambios?", vbYesNo + vbQuestion) = vbNo Then
        Cancel = 1
    End If
End If
End Sub

'Al hacer clic en la lista del historial de actividades ejecutadas
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
'If Not cmdProceso(1).Enabled Then
'    cmdProceso(3).Enabled = True
'End If
'cmdProceso(2).Enabled = (InStr(Item.Tag, "|") > 0)
'cmdProceso(1).Enabled = cmdProceso(2).Enabled
cmdProceso(0).Enabled = True
cmdProceso(1).Enabled = True
cmdProceso(2).Enabled = True
cmdProceso(3).Enabled = True
End Sub

Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
'If Not cmdProceso(5).Enabled Then
'    cmdProceso(7).Enabled = True
'End If
cmdProceso(6).Enabled = True  '(InStr(Item.Tag, "|") > 0)
cmdProceso(5).Enabled = True  'cmdProceso(5).Enabled
cmdProceso(4).Enabled = True  'cmdProceso(5).Enabled
cmdProceso(7).Enabled = True  'cmdProceso(5).Enabled
End Sub

Private Sub OpcSexo_KeyPress(Index As Integer, KeyAscii As Integer)
Dim i As Long, i1 As Long
If KeyAscii = 13 Then
    i1 = OpcSexo(Index).TabIndex
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

'Cambio de opción de registros mostrados
Private Sub opcReg_Click(Index As Integer)
If Index = 0 And Not mbNoPreg Then
    If MsgBox("¿Está seguro de mostrar los últimos 10 Empleados Registrados?", vbYesNo + vbQuestion, "") = vbYes Then
        If AdorsPrin.State Then AdorsPrin.Close
        AdorsPrin.Open msConsultaP & msOrden, gConSql, adOpenStatic, adLockReadOnly
        Call ActualizaBotones(Me, 2, myPermiso)
        RefrescaDatos
    End If
ElseIf Index = 1 And Not mbNoPreg Then
    'If MsgBox("¿Está seguro de mostrar todo Los Empleados?", vbYesNo + vbQuestion, "") = vbYes Then
        If AdorsPrin.State Then AdorsPrin.Close
        AdorsPrin.Open Mid(msConsultaP, 1, InStr(msConsultaP, " where ") - 1) & " order by 1", gConSql, adOpenStatic, adLockReadOnly
        Call ActualizaBotones(Me, 2, myPermiso)
        RefrescaDatos
    'End If
End If
End Sub

Private Sub opcReg_KeyPress(Index As Integer, KeyAscii As Integer)
Dim i As Long, i1 As Long
If KeyAscii = 13 Then
    i1 = opcReg(Index).TabIndex
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

'Acciones de los botones de la barra de herramientas
Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
Dim iIr_a As Long, i As Integer, sCondición As String, yError As Integer, l As Long
Dim adors As ADODB.Recordset, s As String, s1 As String
Dim sVal As String, sCam As String
Dim ii As Integer

On Error GoTo ErrorAcción:

Select Case Button.Key
Case "Primero", "Anterior", "Siguiente", "Último", "Ir_a"
    If mbCambio Then
        If MsgBox("¿Desea guardar los cambios realizados?", vbYesNo + vbQuestion, "") = vbYes Then
            Dim botón As Object
            Set botón = Me.Toolbar.Buttons(3)
            Call Toolbar_ButtonClick(botón)
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
    If mlMovimiento = 0 Then
        Call MsgBox("Debe asignarse una persona antes de guardar los datos del nuevo Empleado", vbOKOnly + vbInformation, "")
        Exit Sub
    End If
    If mlMovimiento = 0 Then
        MsgBox "Debe Asignar a la persona reclutada (Contratada): ", vbInformation + vbOKOnly, "validación"
        Exit Sub
    End If
    For i = 17 To 17
        If Len(Trim(txtCampo(i).Text)) = 0 Then Exit For
    Next
    If i < 17 Then
        MsgBox "Falta dato requerido: " & txtCampo(i).DataField, vbInformation + vbOKOnly, "validación"
        Exit Sub
    End If
    For i = 7 To 7
        If cmbCampo(i).ListIndex < 0 Then Exit For
    Next
    If i < 7 Then
        MsgBox "Falta dato requerido: " & cmbCampo(i).DataField, vbInformation + vbOKOnly, "validación"
        Exit Sub
    End If
        
    If MsgBox("Se agregará un nuevo registro. ¿Está seguro de la operación?", vbYesNo + vbQuestion, "") = vbNo Then
        Exit Sub
    End If
    Set adors = New ADODB.Recordset
    adors.Open "select max(n_cveempleado) from t_rhEmpleados", gConSql, adOpenStatic, adLockReadOnly
    If adors(0) > 0 Then
        mlEmpleado = adors(0) + 1
    Else
        mlEmpleado = 1
    End If
    txtCampo(0).Text = mlEmpleado
    For i = 17 To 30  'Solo los elementos de la tabla de empleados
        s = ArmaCadenaCampo(txtCampo(i).DataField, txtCampo(i).Text, txtCampo(i).Tag, 2)
        sVal = sVal & s
        sCam = sCam & txtCampo(i).DataField & ","
    Next
    For i = 7 To 9  'Solo los elementos de la tabla de empleados
        If cmbCampo(i).ListIndex >= 0 Then
            s = cmbCampo(i).ItemData(cmbCampo(i).ListIndex)
        Else
            s = ""
        End If
        s = ArmaCadenaCampo(cmbCampo(i).DataField, s, "n", 2)
        sVal = sVal & s
        sCam = sCam & cmbCampo(i).DataField & ","
    Next
    sVal = Mid(sVal, 1, Len(sVal) - 1)
    sCam = Mid(sCam, 1, Len(sCam) - 1)
    GuardaBitácora gs_usuario, myTabla, mlEmpleado, 2
    If F_Transacción("insert into t_rhEmpleados (n_cveempleado,n_cveseguimiento," & sCam & ",n_cveusuario,f_registro) values (" & mlEmpleado & "," & mlMovimiento & "," & sVal & "," & gs_usuario & ", getdate())") Then
        Set botón = Me.Toolbar.Buttons(5)
        Call Toolbar_ButtonClick(botón)
    End If
    
    If AdorsPrin.State > 0 Then AdorsPrin.Close
    AdorsPrin.Open Mid(msConsultaP, 1, InStr(msConsultaP, " where ") - 1) & " where n_cveempleado=" & mlEmpleado, gConSql, adOpenStatic, adLockReadOnly
    'AdorsCata.Requery
    'AdorsPrin.AbsolutePosition = ii
    opcReg(0).Value = False
    opcReg(1).Value = False
    RefrescaDatos

Case "Borrar"
    If Val(txtCampo(0).Text) > 0 And Val(txtNoReg.Text) > 0 Then
        If MsgBox("¿Está Seguro de borrar el registro seleccionado?", vbYesNo + vbQuestion, "Confirmación") = vbYes Then
            i = AdorsPrin.Bookmark
            GuardaBitácora gs_usuario, myTabla, Val(txtCampo(0).Text), 5
            F_Transacción ("delete from t_rhEmpleados where n_cveempleado=" & Val(txtCampo(0).Text))
            Set botón = Me.Toolbar.Buttons(8)
            Call Toolbar_ButtonClick(botón)
            If i > AdorsPrin.RecordCount Then
                i = AdorsPrin.RecordCount
            End If
            If i > 0 Then
                txtNoReg.Text = i
                Set botón = Me.Toolbar.Buttons(7)
                Call Toolbar_ButtonClick(botón)
            End If
        End If
    End If
Case "Actualizar"
    If mlMovimiento = 0 Then
        Call MsgBox("Es necesario asignar una persona antes de actualizar los datos del Empleado", vbOKOnly + vbInformation, "")
        Exit Sub
    End If
    For i = 17 To 17
        If Len(Trim(txtCampo(i).Text)) = 0 Then Exit For
    Next
    If i < 17 Then
        MsgBox "Falta dato requerido: " & txtCampo(i).DataField, vbInformation + vbOKOnly, "validación"
        Exit Sub
    End If
    For i = 7 To 7
        If cmbCampo(i).ListIndex < 0 Then Exit For
    Next
    If i < 7 Then
        MsgBox "Falta dato requerido: " & cmbCampo(i).DataField, vbInformation + vbOKOnly, "validación"
        Exit Sub
    End If
    If Val(txtCampo(0).Text) = 0 Then Exit Sub
    sVal = ""
    For i = 17 To 30  'Solo los elementos de la tabla de empleados
        n = 0
        If InStr(txtCampo(i).Tag, "|") > 0 Then
            If Mid(txtCampo(i).Tag, InStr(txtCampo(i).Tag, "|") + 1) <> txtCampo(i).Text Then
                n = 1
            End If
        End If
        If n > 0 Then
            s = ArmaCadenaCampo(txtCampo(i).DataField, txtCampo(i).Text, txtCampo(i).Tag, 0)
            sVal = sVal & s
        End If
    Next
    For i = 7 To 9  'Solo los elementos de la tabla de empleados
        n = 0
        If InStr(cmbCampo(i).Tag, "|") > 0 Then
            If Val(Mid(cmbCampo(i).Tag, InStr(cmbCampo(i).Tag, "|") + 1)) <> cmbCampo(i).ListIndex Then
                n = 1
            End If
        End If
        If n > 0 Then
            If cmbCampo(i).ListIndex >= 0 Then
                s = cmbCampo(i).ItemData(cmbCampo(i).ListIndex)
            Else
                s = ""
            End If
            s = ArmaCadenaCampo(cmbCampo(i).DataField, s, "n", 0)
            sVal = sVal & s
        End If
    Next
    mlEmpleado = Val(txtCampo(0).Text)
    If Len(sVal) > 0 Then
        If Mid(sVal, Len(sVal), 1) = "," Then
            sVal = Mid(sVal, 1, Len(sVal) - 1)
        End If
        If MsgBox("Se actualizarán datos del registro actual. ¿Está seguro de la operación?", vbYesNo + vbQuestion, "") = vbNo Then
            Exit Sub
        End If
        GuardaBitácora gs_usuario, myTabla, mlEmpleado, 3
        Call F_Transacción("update t_rhEmpleados set " & sVal & " where n_cveempleado=" & mlEmpleado)
    Else
        MsgBox "No se realizó ningún cambio", vbInformation + vbOKOnly, ""
        Exit Sub
    End If
    If AdorsPrin.State > 0 Then AdorsPrin.Close
    AdorsPrin.Open Mid(msConsultaP, 1, InStr(msConsultaP, " where ") - 1) & " where n_cveempleado=" & mlEmpleado, gConSql, adOpenStatic, adLockReadOnly
    opcReg(0).Value = False
    opcReg(1).Value = False
    RefrescaDatos
Case "Limpiar"
    mbLimpia = True
    opcReg(0).Value = False
    opcReg(1).Value = False
    For i = txtCampo.LBound To txtCampo.UBound
        txtCampo(i).Text = ""
        If txtCampo(i).Locked Then txtCampo(i).Locked = False
    Next
    For i = cmbCampo.LBound To cmbCampo.UBound
        cmbCampo(i).ListIndex = -1
        'cmbCampo(i).Text = ""
    Next
    mlPersona = 0
    mlMovimiento = 0
    mlEmpleado = 0
    txtNoReg.Locked = True
    txtcompleto.Text = ""
    txtedad.Text = ""
    chkNacionalidad.Value = 2
    OpcSexo(0).Value = False
    OpcSexo(1).Value = False
    Call ActualizaBotones(Me, 1, myPermiso)
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    ListView3.ListItems.Clear
    ListView4.ListItems.Clear
    Set PicFoto.Picture = Nothing
    
Case "Buscar"
    sCondición = ""
    If Val(txtCampo(0).Text) > 0 Or mlBuscaEmp > 0 Then
        If mlBuscaEmp > 0 Then
            sCondición = " n_cveEmpleado=" & mlBuscaEmp & " and "
            mlBuscaEmp = 0
        Else
            sCondición = " n_cveempleado=" & Val(txtCampo(0).Text) & " and "
        End If
    Else
        For i = 1 To Controls.Count - 1
            If LCase(Controls(i).Name) = "txtcampo" Then
                If LCase(Controls(i).DataField) <> LCase(msPrin) And Len(Trim(Controls(i).Text)) > 0 Then
                    s = ArmaCadenaCampo(Controls(i).DataField, Controls(i).Text, Controls(i).Tag, 1)
                    sCondición = sCondición & s
                End If
            ElseIf Mid(Controls(i).Name, 1, 3) = "cmb" Then
                If Controls(i).ListIndex >= 0 Then
                    s = Controls(i).ItemData(Controls(i).ListIndex)
                    s = ArmaCadenaCampo(Controls(i).DataField, s, "n", 1)
                    sCondición = sCondición & s
                End If
            ElseIf Mid(Controls(i).Name, 1, 3) = "chk" Then
                If Controls(i).Value < 2 Then
                    sCondición = sCondición & Controls(i).DataField & "=" & IIf(Controls(i).Value = 1, 1, 0) & " and "
                End If
            End If
        Next
        If OpcSexo(0).Value Then
            sCondición = sCondición & "n_sexo=0 and "
        ElseIf OpcSexo(1).Value Then
            sCondición = sCondición & "n_sexo=1 and "
        End If
    End If
    If Len(sCondición) > 3 Then
        sCondición = " where " & Mid(sCondición, 1, Len(sCondición) - 5)
        If AdorsPrin.State > 0 Then AdorsPrin.Close
        If InStr(msConsulta, " where ") > 0 Then
            AdorsPrin.Open Mid(msConsulta, 1, InStr(msConsulta, " where ") - 1) & sCondición, gConSql, adOpenStatic, adLockReadOnly
        Else
            AdorsPrin.Open msConsulta & sCondición, gConSql, adOpenStatic, adLockReadOnly
        End If
        If AdorsPrin.RecordCount > 0 Then
            RefrescaDatos
            Call ActualizaBotones(Me, 2, myPermiso)
            opcReg(0).Value = False
            opcReg(1).Value = False
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
'    CReport.ParameterFields(0) = "@ipersona;" & mlEmpleado & "); true"
'
'    s = LCase(gConSql.ConnectionString)
'    s = Mid(s, InStr(s, ";pwd=") + 1)
'    s = Mid(s, 1, InStr(Mid(s, 2), ";"))
'    CReport.Connect = "Filedsn=rh.dsn;" & s
'
'    CReport.Action = 1
'    gi = mlEmpleado
'    ReporteEmpleado.Show vbModal
Dim S_ServerExt, S_BaseDatosExt, S_PassExt, S_LogExt
    If mlEmpleado = 0 Then
        Exit Sub
    End If
    If myPermisoRep = 0 Then
        MsgBox "No cuenta con privilegios para emitir informes de este módulo", vbOKOnly + vbInformation, "Validación"
        Exit Sub
    End If
'    CReport.ReportFileName = gsDirReportes + "\Empleados Movimientos.rpt"
'    CReport.ParameterFields(0) = "iEmpleado;" & mlEmpleado & "; true"
'    S_LogExt = Fu_LeeDatosArchConfig(1, "Central")
'    S_PassExt = Fu_LeeDatosArchConfig(2, "Central")
'    S_BaseDatosExt = Fu_LeeDatosArchConfig(3, "Central")
'    S_ServerExt = Fu_LeeDatosArchConfig(4, "Central")
'    CReport.Connect = "Provider=SQLOLEDB;SERVER=" & S_ServerExt & ";DATABASE=" & S_BaseDatosExt & "; UID=" & S_LogExt & "; PWD=" & S_PassExt
'    CReport.Action = 1


    If myPermisoRep = 0 Then
        MsgBox "No cuenta con privilegios para emitir informes de este módulo", vbOKOnly + vbInformation, "Validación"
        Exit Sub
    End If

    gn_OpcionReporte = 5
    If mlEmpleado > 0 Then
        gi2 = mlEmpleado
    End If
    gi1 = 105
    Load RHReportes
    RHReportes.Caption = "Reportes: " & Trim(MDI_Prin.mnuRepRH(gn_OpcionReporte).Caption)
    RHReportes.Show vbModal
    gi1 = 0

Case "Salir"

    Unload Me

End Select
Exit Sub

ErrorAcción:

yError = MsgBox("Error: " + Err.Description, vbAbortRetryIgnore + vbCritical, "Error no esperado (" + Str(Err.Number) + ")")


If yError = vbRetry Then
    Resume
ElseIf yError = vbIgnore Then
    Resume Next
End If

End Sub

'Actualiza datos en el formulario originados por el movimiento de registro del cursor
Sub RefrescaDatos()
Dim i As Long, adors As New ADODB.Recordset, ii As Integer
If adorsCata.State Then adorsCata.Close
If AdorsPrin.EOF Then Exit Sub
mlEmpleado = AdorsPrin!n_cveempleado
adorsCata.Open msConsulta & " where n_cveempleado=" & mlEmpleado, gConSql, adOpenStatic, adLockReadOnly
For i = 0 To cmdProceso.UBound
    cmdProceso(i).Enabled = True
Next
mlMovimiento = adorsCata!n_cveseguimiento
mlPersona = adorsCata!n_cvepersona
For i = 0 To etiCampo.UBound
    If InStr(etiCampo(i).Caption, ":") Then
        If Len(etiCampo(i).Caption) > InStr(etiCampo(i).Caption, ":") Then
            etiCampo(i).Caption = Mid(etiCampo(i).Caption, 1, InStr(etiCampo(i).Caption, ":"))
        End If
    End If
Next
For i = 0 To Controls.Count - 1
    If LCase(Controls(i).Name) = "txtcampo" Then
        Controls(i).Text = IIf(IsNull(adorsCata(Controls(i).DataField)), "", adorsCata(Controls(i).DataField))
        If InStr(Controls(i).Tag, "|") Then
            Controls(i).Tag = Mid(Controls(i).Tag, 1, InStr(Controls(i).Tag, "|") - 1)
        End If
        Controls(i).Tag = Controls(i).Tag & "|" & Controls(i).Text
    ElseIf Mid(Controls(i).Name, 1, 3) = "cmb" Then
        If IsNull(adorsCata(Controls(i).DataField)) Then
            Controls(i).ListIndex = -1
            'Controls(i).Text = ""
        ElseIf Controls(i).ListCount > 0 Then
            Controls(i).ListIndex = BuscaCombo(Controls(i), adorsCata(Controls(i).DataField), True)
            DoEvents
            If Controls(i).ListIndex < 0 Then
                If InStr(Controls(i).Tag, "select") > 0 Then
                    If adors.State Then adors.Close
                    If InStr(Controls(i).Tag, "|") Then
                        adors.Open Mid(Controls(i).Tag, 1, InStr(Controls(i).Tag, "|") - 1) & adorsCata(Controls(i).DataField), gConSql, adOpenStatic, adLockReadOnly
                    Else
                        adors.Open Controls(i).Tag & adorsCata(Controls(i).DataField), gConSql, adOpenStatic, adLockReadOnly
                    End If
                    If Not adors.EOF Then
                        ii = Controls(i).Index
                        If InStr(etiCampo(ii).Caption, ":") > 0 Then
                            etiCampo(ii).Caption = Mid(etiCampo(ii).Caption, 1, InStr(etiCampo(ii).Caption, ":")) & adors(1)
                        Else
                            etiCampo(ii).Caption = etiCampo(ii).Caption & ":" & adors(1)
                        End If
                    End If
                Else
                    MsgBox "No se localizó el valor en el catálogo del campo: " & Controls(i).DataField, vbOKOnly + vbCritical
                End If
            End If
        End If
        If InStr(Controls(i).Tag, "|") Then
            Controls(i).Tag = Mid(Controls(i).Tag, 1, InStr(Controls(i).Tag, "|") - 1)
        End If
        Controls(i).Tag = Controls(i).Tag & "|" & Controls(i).ListIndex
    ElseIf Mid(Controls(i).Name, 1, 3) = "chk" Then
        If IsNull(adorsCata(Controls(i).DataField)) Then
            Controls(i).Value = 0
        Else
            Controls(i).Value = IIf(adorsCata(Controls(i).DataField), 1, 0)
        End If
        If InStr(Controls(i).Tag, "|") Then
            Controls(i).Tag = Mid(Controls(i).Tag, 1, InStr(Controls(i).Tag, "|") - 1)
        End If
        Controls(i).Tag = Controls(i).Tag & "|" & Controls(i).Value
    End If
Next
ActualizaCampoEsp
If AdorsPrin.RecordCount = 0 Then
    MsgBox "No Existen Registros", vbOKOnly + vbInformation, ""
Else
    If AdorsPrin.Bookmark > 0 Then
        txtNoReg.Text = (AdorsPrin.Bookmark) & " / " & AdorsPrin.RecordCount
    Else
        txtNoReg.Text = "??? / " & AdorsPrin.RecordCount
    End If
    txtNoReg.Refresh
End If
mbLimpia = False
mbCambio = False
RefrescaDatosPuestos
RefrescaDatosEventosGpo
End Sub

'Actualiza datos en el formulario originados por la asignación de puestos al empleado
Sub RefrescaDatosPuestos()
Dim i As Long
Dim s As String
Dim adors As New ADODB.Recordset
'Busca el empleado en los puestos asignados
If adors.State Then adors.Close
adors.Open "select * from v_RHEmpleadoPuestos where n_cveempleado=" & mlEmpleado, gConSql, adOpenStatic, adLockReadOnly
i = 1
ListView3.ListItems.Clear
Do While Not adors.EOF
    ListView3.ListItems.Add i, , IIf(IsNull(adors![s_area]), "", adors![s_area])
    ListView3.ListItems(i).Tag = adors!n_cvePuestoAsig
    ListView3.ListItems(i).SubItems(1) = IIf(IsNull(adors![s_departamento]), "", adors![s_departamento])
    ListView3.ListItems(i).SubItems(2) = IIf(IsNull(adors![s_TipoPuesto]), "", adors![s_TipoPuesto])
    ListView3.ListItems(i).SubItems(3) = IIf(IsNull(adors![s_clave]), "", adors![s_clave])
    ListView3.ListItems(i).SubItems(4) = IIf(IsNull(adors![f_Fecha]), "", Format(adors![f_Fecha], gsFormatoFecha))
    ListView3.ListItems(i).SubItems(5) = IIf(IsNull(adors![s_observaciones]), "", adors![s_observaciones])
    adors.MoveNext
    i = i + 1
Loop
End Sub

Sub RefrescaDatosEventosGpo()
Dim i As Long
Dim s As String
Dim adors As New ADODB.Recordset
'Busca el empleado en los eventos grupales
If adors.State Then adors.Close
adors.Open "select * from v_RHEmpleadoEventosGpo where n_cveempleado=" & mlEmpleado, gConSql, adOpenStatic, adLockReadOnly
i = 1
ListView4.ListItems.Clear
Do While Not adors.EOF
    ListView4.ListItems.Add i, , IIf(IsNull(adors![s_tipoeventogpo]), "", adors![s_tipoeventogpo])
    ListView4.ListItems(i).Tag = adors!n_cveeventogpo
    ListView4.ListItems(i).SubItems(1) = IIf(IsNull(adors![s_clave]), "", adors![s_clave])
    ListView4.ListItems(i).SubItems(2) = IIf(IsNull(adors![s_Lugar]), "", adors![s_Lugar])
    ListView4.ListItems(i).SubItems(3) = IIf(IsNull(adors![s_expositor]), "", adors![s_expositor])
    ListView4.ListItems(i).SubItems(4) = IIf(IsNull(adors![f_inicio]), "", Format(adors![f_inicio], gsFormatoFecha))
    ListView4.ListItems(i).SubItems(5) = IIf(IsNull(adors![f_termino]), "", Format(adors![f_termino], gsFormatoFecha))
    ListView4.ListItems(i).SubItems(6) = IIf(IsNull(adors![s_observaciones]), "", adors![s_observaciones])
    adors.MoveNext
    i = i + 1
Loop
End Sub

'Actualiza datos en el formulario originados por el movimiento de registro del cursor
Sub RefrescaDatosAsignación()
Dim i As Long
Dim s As String
'Busca la persona que
If adorsCata.State Then adorsCata.Close
adorsCata.Open "select p.* from t_rhpersonalseg ps inner join t_rhpersonal p on ps.n_cvepersona=p.n_cvepersona where ps.n_cveseguimiento=" & mlMovimiento, gConSql, adOpenStatic, adLockReadOnly
If adorsCata.EOF Then Exit Sub
mlPersona = adorsCata!n_cvepersona
'For i = 0 To cmdProceso.UBound
'    cmdProceso(i).Enabled = False
'Next
For i = 1 To 16
    If InStr(txtCampo(i).DataField, ".") Then
        s = Mid(txtCampo(i).DataField, InStr(txtCampo(i).DataField, ".") + 1)
    Else
        s = txtCampo(i).DataField
    End If
    txtCampo(i).Text = IIf(IsNull(adorsCata(s)), "", adorsCata(s))
    If InStr(txtCampo(i).Tag, "|") Then
        txtCampo(i).Tag = Mid(txtCampo(i).Tag, 1, InStr(txtCampo(i).Tag, "|") - 1)
    End If
    txtCampo(i).Tag = txtCampo(i).Tag & "|" & txtCampo(i).Text
Next
For i = 0 To 6
    If IsNull(adorsCata(cmbCampo(i).DataField)) Then
        cmbCampo(i).ListIndex = -1
        'cmbCampo(i).Text = ""
    ElseIf cmbCampo(i).ListCount > 0 Then
        cmbCampo(i).ListIndex = BuscaCombo(cmbCampo(i), adorsCata(cmbCampo(i).DataField), True)
        DoEvents
        If cmbCampo(i).ListIndex < 0 Then
            MsgBox "No se localizó el valor en el catálogo del campo: " & cmbCampo(i).DataField, vbOKOnly + vbCritical
        End If
    End If
    If InStr(cmbCampo(i).Tag, "|") Then
        cmbCampo(i).Tag = Mid(cmbCampo(i).Tag, 1, InStr(cmbCampo(i).Tag, "|") - 1)
    End If
    cmbCampo(i).Tag = cmbCampo(i).Tag & "|" & cmbCampo(i).ListIndex
Next
If IsNull(adorsCata(chkNacionalidad.DataField)) Then
    chkNacionalidad.Value = 0
Else
    chkNacionalidad.Value = IIf(adorsCata(chkNacionalidad.DataField), 1, 0)
End If
If InStr(chkNacionalidad.Tag, "|") Then
    chkNacionalidad.Tag = Mid(chkNacionalidad.Tag, 1, InStr(chkNacionalidad.Tag, "|") - 1)
End If
chkNacionalidad.Tag = chkNacionalidad.Tag & "|" & chkNacionalidad.Value
ActualizaCampoEsp False
End Sub

'Actualiza combo de municipios según valor del estado y otros (sexo, edad, ...)
Sub ActualizaCampoEsp(Optional bNoActLista As Boolean)
Static iEdo As Integer
Dim d As Date, b As Boolean, adors As New ADODB.Recordset
If mbInicio Then
    mbInicio = False
    iEdo = -1
End If
If adorsCata.Bookmark > 0 Then
    If Not IsNull(adorsCata!f_fecha_nac) Then
        txtedad.Text = sEdad(Format(adorsCata!f_fecha_nac, gsFormatoFecha))
    End If
    If Not IsNull(adorsCata!n_sexo) Then
        OpcSexo(IIf(adorsCata!n_sexo = 0, 0, 1)).Value = True
    Else
        OpcSexo(0).Value = False
        OpcSexo(1).Value = False
    End If
    
    'ActualizaLista
    If Not bNoActLista Then ActualizaLista
    
    If adors.State Then adors.Close
    adors.Open "select * from t_rhPersonalfoto where n_cvepersona=" & mlPersona, gConSql, adOpenDynamic, adLockOptimistic
    If Not adors.EOF Then
        If Not IsNull(adors(1)) Then
            On Error GoTo 0
            LeerBinary adors(1), PicFoto, mlPersona

            'Set PicFoto.Picture = adors(0)
        Else
            Set PicFoto.Picture = Nothing
        End If
    Else
        Set PicFoto.Picture = Nothing
    End If
    txtcompleto.Text = LTrim(txtCampo(1).Text & " ") & LTrim(txtCampo(2).Text & " ") & LTrim(txtCampo(3).Text & " ") & "(" & LTrim(txtCampo(5).Text) & ")"
End If
End Sub

'habilita bandera de cambio
Private Sub txtCampo_Change(Index As Integer)
If i < 16 Then Exit Sub
If Not mbCambio And Not mbRefresca Then mbCambio = True
'HABILITA
End Sub

Private Sub txtCampo_GotFocus(Index As Integer)
gs1 = txtCampo(Index).Text
End Sub

Private Sub txtCampo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Mid(txtCampo(Index).Tag, 1, 1) = "f" And KeyCode = 27 Then txtCampo(Index) = ""
End Sub

'Valida caracteres de entrada según el tipo de campo
Private Sub txtCampo_KeyPress(Index As Integer, KeyAscii As Integer)
Dim i As Long, i1 As Long
If Index >= 1 And Index <= 16 Then
    KeyAscii = 0
    Exit Sub
End If
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
KeyAscii = TeclaOprimida(txtCampo(Index), KeyAscii, txtCampo(Index).Tag, Me.Toolbar.Buttons(5).Enabled)
'If Index = 1 And txtCampo(Index).SelStart = 0 And InStr("1234567890", Chr(KeyAscii)) = 0 Then
'    KeyAscii = 0
'End If

End Sub

Private Sub txtCampo_LostFocus(Index As Integer)
Dim d As Date, adors As New ADODB.Recordset
If txtCampo(Index).Text = gs1 Then Exit Sub
If Mid(txtCampo(Index).Tag, 1, 1) = "f" Then
    If IsDate(txtCampo(Index).Text) Then
        d = CDate(txtCampo(Index).Text)
        adors.Open "select getdate()", gConSql, adOpenStatic, adLockReadOnly
        If Int(adors(0)) - Int(d) < 0 Then
            Call MsgBox("Fecha no válida. No se permite ingresar fecha mayor a la fecha actual (" & Format(adors(0), gsFormatoFecha) & ")", vbOKOnly + vbInformation, "")
            txtCampo(Index).Text = ""
            Exit Sub
        End If
        If Index = 17 Then  'Valida Alta
            If IsDate(txtCampo(22).Text) Then
                If CDate(txtCampo(22).Text) > d Then
                    Call MsgBox("Fecha no válida. La fecha de alta no debe ser menor a la fecha de baja (" & Format(CDate(txtCampo(22).Text), gsFormatoFecha) & ")", vbOKOnly + vbInformation, "")
                    txtCampo(Index).Text = ""
                    Exit Sub
                End If
            End If
        ElseIf Index = 22 Then 'valida baja
            If IsDate(txtCampo(17).Text) Then
                If CDate(txtCampo(17).Text) > d Then
                    Call MsgBox("Fecha no válida. La fecha de baja no debe ser mayor a la fecha de alta (" & Format(CDate(txtCampo(17).Text), gsFormatoFecha) & ")", vbOKOnly + vbInformation, "")
                    txtCampo(Index).Text = ""
                    Exit Sub
                End If
            End If
        End If
        txtCampo(Index).Text = Format(d, gsFormatoFecha)
    Else
        If Len(txtCampo(Index).Text) > 0 Then
            Call MsgBox("Fecha no válida. Verificar", vbOKOnly + vbInformation, "")
            txtCampo(Index) = ""
        End If
        If Index <= 5 Then
            txtedad.Text = ""
        End If
    End If
End If
End Sub

Private Sub txtcompleto_KeyPress(KeyAscii As Integer)
Dim i As Long, i1 As Long
If KeyAscii = 13 Then
    i1 = txtcompleto.TabIndex
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

Private Sub txtedad_KeyPress(KeyAscii As Integer)
Dim i As Long, i1 As Long
If KeyAscii = 13 Then
    i1 = txtedad.TabIndex
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

'Verifica y realiza acción del campo de No. Registro
Private Sub txtNoReg_KeyDown(KeyCode As Integer, Shift As Integer)
Dim botón As Object
If KeyCode = 13 Then
    If Val(txtNoReg) > 0 Then
        'mdi.ActiveForm.iIr_a = Val(txtReg)
        Set botón = Me.Toolbar.Buttons(7)
        Call Toolbar_ButtonClick(botón)
    End If
End If
End Sub

'Refresca datos del movimientos administrativos del empleado
Sub ActualizaLista()
Dim i As Long
Dim adors As New ADODB.Recordset
If AdorsPrin.EOF Then
    Exit Sub
End If
If adors.State > 0 Then adors.Close
adors.Open "select em.n_cveMovimiento,mov.s_tipomovimiento,em.f_fecha,res.s_nombre+rtrim(' '+res.s_paterno)+rtrim(' '+res.s_materno) as responsable,em.s_Observaciones,u.s_nombre from t_rhEmpleadosMov em left join c_rhTipoMov mov on em.n_cvetipomov=mov.n_cvetipomov left join t_rhPersonal res on em.n_cveresponsable=res.n_cvepersona left join c_segusuarios u on em.n_cveusuario=u.n_cveusuario where em.n_cveempleado=" & IIf(IsNull(mlEmpleado), -1, mlEmpleado) & " order by f_fecha", gConSql, adOpenStatic, adLockReadOnly
i = 1
ListView1.ListItems.Clear
Do While Not adors.EOF
    ListView1.ListItems.Add i, , IIf(IsNull(adors(1)), "", adors(1))   '1er dato TipoMovimiento
    ListView1.ListItems(i).Tag = adors!n_cvemovimiento   'id del Movimiento
    ListView1.ListItems(i).SubItems(1) = IIf(IsNull(adors(2)), "", adors(2))  'fecha
    ListView1.ListItems(i).SubItems(2) = IIf(IsNull(adors(3)), "", adors(3))  'responsable
    ListView1.ListItems(i).SubItems(3) = IIf(IsNull(adors(4)), "", adors(4))  'observaciones
    ListView1.ListItems(i).SubItems(4) = IIf(IsNull(adors(5)), "", adors(5))  'usuario sistema
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
ActualizaLista2
End Sub

'Refresca datos de evaluaciones del Empleado
Sub ActualizaLista2()
Dim i As Long
Dim adors As New ADODB.Recordset
If AdorsPrin.EOF Then
    Exit Sub
End If
If adors.State > 0 Then adors.Close
adors.Open "select n_cveEvaluacion,dbo.f_tipoevaluacion(n_cvetipoevaluacion) as tipo,dbo.f_resultadoeva(n_cveresultadoeva) as resultado,f_fecha,dbo.f_responsable(n_cveresponsable) as responsable,s_Observaciones,dbo.f_Usuario(n_cveusuario) from t_rhEmpleadosEva where n_cveempleado=" & IIf(IsNull(mlEmpleado), -1, mlEmpleado) & " order by f_fecha", gConSql, adOpenStatic, adLockReadOnly
i = 1
ListView2.ListItems.Clear
Do While Not adors.EOF
    ListView2.ListItems.Add i, , IIf(IsNull(adors(1)), "", adors(1))  '1er dato TipoMovimiento
    ListView2.ListItems(i).Tag = adors(0)  'id de la Evaluación
    ListView2.ListItems(i).SubItems(1) = IIf(IsNull(adors(2)), "", adors(2))  'Resultado
    ListView2.ListItems(i).SubItems(2) = IIf(IsNull(adors(3)), "", adors(3))  'fecha
    ListView2.ListItems(i).SubItems(3) = IIf(IsNull(adors(4)), "", adors(4))  'responsable
    ListView2.ListItems(i).SubItems(4) = IIf(IsNull(adors(5)), "", adors(5))  'observaciones
    ListView2.ListItems(i).SubItems(5) = IIf(IsNull(adors(6)), "", adors(6))  'usuario sistema
    adors.MoveNext
    i = i + 1
Loop
If i > 1 Then
    ListView2.ListItems(i - 1).Tag = ListView2.ListItems(i - 1).Tag & "|"
End If
If Not ListView2.Visible Then
    ListView2.Visible = True
    ListView2.Left = 80
End If

End Sub

'valida si es posible borrar la actividad
Function ValidaBorradoAct(lMovimiento As Long) As Boolean
Dim adors As New ADODB.Recordset, s As String
Dim dAhora As Date
'adors.Open "select count(*) from t_rhEmpleadosMov where n_cveant=" & lMovimiento & " and f_fecha is not null", gConSql, adOpenStatic, adLockReadOnly
'If adors(0) > 0 Then
'    Call MsgBox("La Actividad no es posible borrar ya que esta no es la última del proceso", vbOKOnly + vbInformation)
'    Exit Function
'End If
adors.Open "select top 1 getdate() from c_rhestados", gConSql, adOpenStatic, adLockReadOnly
dAhora = adors(0)
If adors.State Then adors.Close
adors.Open "select count(*) from t_rhEmpleadosMov where n_cvemovimiento=" & lMovimiento & " and f_registro<convert(datetime,'" & Format(gdAhora - 30, gsFormatoFecha) & "',105)", gConSql, adOpenStatic, adLockReadOnly
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

Function ValidaBorradoEva(lEvaluacion As Long) As Boolean
Dim adors As New ADODB.Recordset, s As String
Dim dAhora As Date
adors.Open "select getdate()", gConSql, adOpenStatic, adLockReadOnly
dAhora = adors(0)
If adors.State Then adors.Close
adors.Open "select count(*) from t_rhEmpleadosEva where n_cveevaluacion=" & lEvaluacion & " and f_registro<convert(datetime,'" & Format(gdAhora - 30, gsFormatoFecha) & "',105)", gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    Call MsgBox("No se permite realizar borrado de información después de 30 días", vbOKOnly + vbInformation)
    Exit Function
End If
ValidaBorradoEva = True
End Function

Private Sub txtPen_KeyPress(KeyAscii As Integer)
Dim i As Long, i1 As Long
If KeyAscii = 13 Then
    i1 = txtPen.TabIndex
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
