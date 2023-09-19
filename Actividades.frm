VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "fm20.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form Actividades 
   Caption         =   "Registro de actividades"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   12645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9405
      Left            =   30
      TabIndex        =   33
      Top             =   0
      Width           =   12525
      Begin VB.TextBox txtcampo 
         Height          =   285
         Index           =   5
         Left            =   10440
         TabIndex        =   79
         Tag             =   "f"
         Top             =   6105
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtFAcuerdo2 
         Height          =   285
         Left            =   3720
         TabIndex        =   77
         Tag             =   "f"
         Top             =   6105
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Frame Frame5 
         Caption         =   "Opciones de actividades siguientes:"
         Height          =   735
         Left            =   120
         TabIndex        =   73
         Top             =   8040
         Visible         =   0   'False
         Width           =   12255
         Begin MSForms.OptionButton opcAct 
            Height          =   400
            Index           =   2
            Left            =   7920
            TabIndex        =   76
            Top             =   240
            Width           =   4095
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "7223;706"
            Value           =   "0"
            Caption         =   "Actividad1"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.OptionButton opcAct 
            Height          =   400
            Index           =   1
            Left            =   3960
            TabIndex        =   75
            Top             =   240
            Width           =   3735
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "6588;706"
            Value           =   "0"
            Caption         =   "Actividad1"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.OptionButton opcAct 
            Height          =   400
            Index           =   0
            Left            =   120
            TabIndex        =   74
            Top             =   240
            Width           =   3735
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "6588;706"
            Value           =   "0"
            Caption         =   "Actividad1"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.CommandButton cmdSig 
         BackColor       =   &H00008000&
         Caption         =   "Siguiente >>"
         Height          =   375
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   7560
         Width           =   1335
      End
      Begin VB.CommandButton cmdAnt 
         BackColor       =   &H000040C0&
         Caption         =   "<< Anterior"
         Height          =   375
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   7560
         Width           =   1335
      End
      Begin VB.TextBox txtExp 
         Height          =   315
         Left            =   9660
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   2820
      End
      Begin VB.Frame frProgress 
         Height          =   1620
         Left            =   1530
         TabIndex        =   47
         Top             =   7965
         Visible         =   0   'False
         Width           =   7530
         Begin MSComctlLib.ProgressBar pbProgress 
            Height          =   330
            Left            =   180
            TabIndex        =   56
            Top             =   855
            Width           =   5910
            _ExtentX        =   10425
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.CommandButton btnCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   6210
            TabIndex        =   54
            Top             =   765
            Width           =   1110
         End
         Begin VB.Label lSource 
            Caption         =   "Archivo Origen:"
            Height          =   255
            Left            =   225
            TabIndex        =   53
            Top             =   225
            Width           =   1200
            WordWrap        =   -1  'True
         End
         Begin VB.Label lDest 
            Caption         =   "Archivo Destino:"
            Height          =   255
            Left            =   225
            TabIndex        =   52
            Top             =   540
            Width           =   1215
            WordWrap        =   -1  'True
         End
         Begin VB.Label lProcessed 
            Caption         =   "Procesado:"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   1215
            Width           =   975
         End
         Begin VB.Label lSourceFilename 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   1560
            TabIndex        =   50
            Top             =   360
            Width           =   45
         End
         Begin VB.Label lDestFilename 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   1560
            TabIndex        =   49
            Top             =   720
            Width           =   45
         End
         Begin VB.Label lProgress 
            Height          =   255
            Left            =   1620
            TabIndex        =   48
            Top             =   1215
            Width           =   4425
         End
      End
      Begin MSComctlLib.ListView lvLog 
         Height          =   1275
         Left            =   90
         TabIndex        =   57
         Top             =   8055
         Width           =   10590
         _ExtentX        =   18680
         _ExtentY        =   2249
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CheckBox chkacuerdo 
         Caption         =   "No Emitir Número de Acuerdo"
         Height          =   240
         Left            =   6960
         TabIndex        =   11
         Top             =   5640
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.CommandButton cmdCondonación 
         Caption         =   "Condonación"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   7560
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.TextBox txtcampo 
         Height          =   285
         Index           =   4
         Left            =   8640
         TabIndex        =   13
         Tag             =   "f"
         Top             =   6120
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtcampo 
         Height          =   285
         Index           =   3
         Left            =   5610
         TabIndex        =   12
         Top             =   6120
         Visible         =   0   'False
         Width           =   2955
      End
      Begin VB.CommandButton cmdVerificaDocto 
         Caption         =   "Verificar Documento"
         Height          =   375
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   7560
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.PictureBox Inet1 
         Height          =   480
         Left            =   10680
         ScaleHeight     =   420
         ScaleWidth      =   1740
         TabIndex        =   55
         Top             =   1035
         Width           =   1800
      End
      Begin VB.CommandButton cmdSubirDocto 
         Caption         =   "Subir Documento"
         Height          =   375
         Left            =   1455
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   7560
         Width           =   1545
      End
      Begin VB.CommandButton cmdSanción 
         Caption         =   "Sanción"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   7575
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.TextBox txtAcuerdo 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   6120
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.TextBox txtEtiqueta 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Height          =   465
         Left            =   1530
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   43
         Text            =   "Actividades.frx":0000
         Top             =   90
         Visible         =   0   'False
         Width           =   2500
      End
      Begin VB.CommandButton cmdBotón 
         Caption         =   "&Integración de expediente"
         Height          =   435
         Index           =   2
         Left            =   10215
         Picture         =   "Actividades.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   7545
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.CommandButton cmdBotón 
         Caption         =   "&Aceptar"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   4560
         TabIndex        =   25
         Top             =   7560
         Width           =   1035
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   1965
         Top             =   120
      End
      Begin VB.CommandButton cmdBotón 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5880
         TabIndex        =   31
         Top             =   7575
         Width           =   1050
      End
      Begin VB.CommandButton cmdNavega 
         Height          =   345
         Index           =   3
         Left            =   2970
         Picture         =   "Actividades.frx":0466
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Ir al último registro"
         Top             =   180
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton cmdNavega 
         Height          =   345
         Index           =   2
         Left            =   2610
         Picture         =   "Actividades.frx":091C
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Siguiente registro"
         Top             =   135
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton cmdNavega 
         Height          =   345
         Index           =   0
         Left            =   3240
         Picture         =   "Actividades.frx":0DD2
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Ir al primer registro"
         Top             =   135
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdNavega 
         Height          =   345
         Index           =   1
         Left            =   3630
         Picture         =   "Actividades.frx":1288
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Anterior registro"
         Top             =   135
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtRegistro 
         Height          =   345
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "Reg./Total"
         ToolTipText     =   "No.de registro"
         Top             =   135
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.TextBox txtFechaProgramada 
         Height          =   315
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   3765
      End
      Begin VB.ComboBox ComboDesenlaces 
         Height          =   315
         Left            =   135
         TabIndex        =   8
         Top             =   5535
         Visible         =   0   'False
         Width           =   6555
      End
      Begin VB.TextBox txtcampo 
         DataField       =   "fecha"
         DataSource      =   "datEncuestas"
         Height          =   315
         Index           =   0
         Left            =   165
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "Automática"
         ToolTipText     =   "Actividad"
         Top             =   510
         Width           =   5610
      End
      Begin VB.ComboBox ComboResponsable 
         DataField       =   "idres"
         Height          =   315
         Left            =   135
         TabIndex        =   3
         ToolTipText     =   "Responsable"
         Top             =   1080
         Width           =   6720
      End
      Begin VB.TextBox txtcampo 
         BackColor       =   &H00E0E0E0&
         Height          =   2115
         Index           =   2
         Left            =   6750
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         ToolTipText     =   "Observaciones"
         Top             =   3480
         Width           =   5640
      End
      Begin VB.TextBox txtcampo 
         DataField       =   "fecha"
         DataSource      =   "datEncuestas"
         Height          =   315
         Index           =   1
         Left            =   6885
         TabIndex        =   4
         Tag             =   "f"
         Text            =   "Automática"
         ToolTipText     =   "Fecha de inicio"
         Top             =   1080
         Width           =   3765
      End
      Begin MSComctlLib.TreeView TreeView2 
         Height          =   1740
         Left            =   120
         TabIndex        =   7
         Top             =   3540
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   3069
         _Version        =   393217
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   1710
         Left            =   6750
         TabIndex        =   6
         Top             =   1620
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3016
         _Version        =   393217
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView TreeView3 
         Height          =   1695
         Left            =   135
         TabIndex        =   5
         Top             =   1620
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   2990
         _Version        =   393217
         Sorted          =   -1  'True
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos Adicionales"
         Height          =   1065
         Left            =   45
         TabIndex        =   64
         Top             =   6390
         Visible         =   0   'False
         Width           =   12330
         Begin VB.TextBox txtFAcuerdo 
            Height          =   330
            Left            =   4905
            MaxLength       =   20
            TabIndex        =   18
            Tag             =   "f"
            Top             =   540
            Width           =   1635
         End
         Begin VB.TextBox txtAcuCierre 
            Height          =   330
            Left            =   180
            MaxLength       =   80
            TabIndex        =   17
            Top             =   540
            Width           =   4725
         End
         Begin VB.TextBox txtFMemo 
            Height          =   330
            Left            =   10545
            MaxLength       =   20
            TabIndex        =   20
            Tag             =   "f"
            Top             =   540
            Width           =   1635
         End
         Begin VB.TextBox txtMemorando 
            Height          =   330
            Left            =   6540
            MaxLength       =   80
            TabIndex        =   19
            Top             =   540
            Width           =   4005
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha Acuerdo:"
            Height          =   285
            Left            =   4905
            TabIndex        =   68
            Top             =   270
            Width           =   1365
         End
         Begin VB.Label lblAdoCierre 
            AutoSize        =   -1  'True
            Caption         =   "No. Acuerdo de cierre:"
            Height          =   195
            Left            =   180
            TabIndex        =   67
            Top             =   270
            Width           =   1605
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha Memo de cierre:"
            Height          =   285
            Left            =   10560
            TabIndex        =   66
            Top             =   240
            Width           =   1650
         End
         Begin VB.Label Label8 
            Caption         =   "Memorando de cierre:"
            Height          =   285
            Left            =   6540
            TabIndex        =   65
            Top             =   270
            Width           =   2085
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos Adicionales"
         Height          =   720
         Left            =   90
         TabIndex        =   59
         Top             =   6480
         Visible         =   0   'False
         Width           =   10500
         Begin VB.ComboBox cmbNotificador 
            Height          =   315
            Left            =   1080
            TabIndex        =   14
            Top             =   240
            Width           =   5340
         End
         Begin VB.Label Label2 
            Caption         =   "Notificador:"
            Height          =   165
            Left            =   180
            TabIndex        =   60
            Top             =   270
            Width           =   1365
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos Adicionales"
         Height          =   885
         Left            =   45
         TabIndex        =   61
         Top             =   6465
         Visible         =   0   'False
         Width           =   12330
         Begin VB.TextBox txtOficio 
            Height          =   330
            Left            =   180
            TabIndex        =   15
            Top             =   420
            Width           =   6885
         End
         Begin VB.TextBox txtDOtorgados 
            Height          =   330
            Left            =   7350
            TabIndex        =   16
            Tag             =   "n"
            Top             =   420
            Width           =   735
         End
         Begin VB.Label lblNvoOficio 
            AutoSize        =   -1  'True
            Caption         =   "Nuevo Oficio:"
            Height          =   195
            Left            =   180
            TabIndex        =   63
            Top             =   150
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Días Otorgados:"
            Height          =   285
            Left            =   7350
            TabIndex        =   62
            Top             =   150
            Width           =   1365
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Motivo de No Sanción"
         Height          =   690
         Left            =   135
         TabIndex        =   69
         Top             =   6720
         Visible         =   0   'False
         Width           =   12165
         Begin VB.ComboBox cmbMotivoNoSan 
            Height          =   315
            Left            =   90
            TabIndex        =   70
            Top             =   270
            Width           =   11895
         End
      End
      Begin VB.Label etiTexto 
         Caption         =   "11. Fecha Entrega"
         Height          =   195
         Index           =   5
         Left            =   10485
         TabIndex        =   80
         Top             =   5880
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label etiFAcuerdo2 
         Caption         =   "8. Fecha Acuerdo"
         Height          =   195
         Left            =   3765
         TabIndex        =   78
         Top             =   5880
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label Label1 
         Caption         =   "Expediente::"
         Height          =   285
         Left            =   9705
         TabIndex        =   58
         Top             =   225
         Width           =   2070
      End
      Begin VB.Label etiTexto 
         Caption         =   "10. Fecha Memorando"
         Height          =   195
         Index           =   4
         Left            =   8640
         TabIndex        =   46
         Top             =   5895
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label etiTexto 
         Caption         =   "9. No. Memorando"
         Height          =   195
         Index           =   3
         Left            =   5655
         TabIndex        =   45
         Top             =   5880
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label EtiAcuerdo 
         Caption         =   "7. No. Acuerdo"
         Height          =   195
         Left            =   180
         TabIndex        =   44
         Top             =   5895
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label etiArbol3 
         Caption         =   "3. Tareas:"
         Height          =   225
         Left            =   180
         TabIndex        =   42
         Top             =   1380
         Width           =   2295
      End
      Begin VB.Label etiTexto1 
         Caption         =   "Fecha programada (Responsable):"
         Height          =   285
         Left            =   5970
         TabIndex        =   41
         Top             =   270
         Width           =   3465
      End
      Begin VB.Label etiCombo 
         Caption         =   "6. Desenlaces:"
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   40
         Top             =   5310
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.Label etiArbol2 
         Caption         =   "5. Documentos:"
         Height          =   225
         Left            =   180
         TabIndex        =   39
         Top             =   3330
         Width           =   2295
      End
      Begin VB.Label etiArbol1 
         Caption         =   "4. Programar siguiente(s) actividad(es):"
         Height          =   225
         Left            =   5715
         TabIndex        =   38
         Top             =   1395
         Width           =   2895
      End
      Begin VB.Label etiTexto 
         Caption         =   "6. Observaciones:"
         Height          =   285
         Index           =   2
         Left            =   5715
         TabIndex        =   37
         Top             =   3330
         Width           =   1635
      End
      Begin VB.Label etiTexto 
         Caption         =   "2. Fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   1
         Left            =   6915
         TabIndex        =   36
         Top             =   840
         Width           =   4920
      End
      Begin VB.Label etiCombo 
         Caption         =   "1. Responsable:"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   35
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label etiTexto 
         Caption         =   "Actividad:"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   34
         Top             =   270
         Width           =   3585
      End
   End
End
Attribute VB_Name = "Actividades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Public msFAcuerdo As String 'Fecha acuerdo
Public msFecha2 As String 'Fecha entrega/devolución
Public msMemo As String 'Contiene información del No. de memorando asosiado al avance
Public msSanción As String 'Contiene información de los datos de la sanción
Public msCondonación As String 'Contiene información de los datos de la condonación
Public miOtro As Integer 'Indicador que dato adicional hay que solicitar.

Public msSeg As String 'Contiene los id de seguimiento que pueden generarse masivamente
Public miAccion ' contiene el valor de la acción para solicitar los datos adicionales cuando es masivo

Public yTipoOperación As Byte '1:Agrega/Seguimiento; 2:Modifica; 0:Consulta

Public sFechaMín As String
Public sFechaMáx As String
Public bAceptar As Boolean
'Public ySoloConsulta As Byte
'Public yUltimo As Byte
Public yFormaRecepción As Byte

Public Canceled As Boolean

Public miMasivo As Integer 'Seguimiento a varios expedientes

Private dlgProgress As frmProgress

Dim yActEstrados As Byte
Dim ySubirArchivo As Byte
Dim yVerificaArchivo As Byte

Dim TransferOperationActive As Boolean

Dim sArchivoFTP As String 'Contiene el nombre del archivo subido  a FTP para publicar en el portal de estrados
Dim sHostRemoto As String 'Contiene el nombre del servidor FTP remoto
Dim yVerificaDocto As Byte 'indicador si ya se trató visualizar el Docto.
Dim ySubirDocto As Byte '1: Extrados; 2: SINE

Dim iActApo As Integer
Dim bRR As Boolean
Dim bOprimióTecla As Boolean
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
Dim sArchivos As String 'Contiene el nombre de los archivos que se generan por los documentos
Dim bNodoSeleccionado As Boolean 'indica que paso el procedimiento de selección del nodo seleccionado
Dim adorsBloqueoAva As New ADODB.Recordset 'Recorset para bloqueo ADO
Dim rsBloqueoAva As DAO.Recordset 'Recorset para bloqueo DAO
Dim miDatosAdi As Integer
Dim sGestiónDoctos As String
Dim bGestiónDoctos As Boolean 'Indica si debe guardar documentos seleccionados al ejecutarse la actividad de UNES
Dim miMotNoSan As Integer 'Id Motívo de no sanción



Private Sub btnCancel_Click()
  Canceled = True
End Sub


Private Sub chkacuerdo_Click()
bOprimióTecla = True
HabilitaAceptar False
End Sub



Private Sub cmdAnt_Click()
If Accion(0) Then 'Guarda datos o acepta la consulta
    yUnico = 0 'Actualiza variable de actulización por única vez
    'Call MsgBox(" actual / ant :" & mlSeguimiento & " / " & mlAnt, vbOKOnly, "")
    mlSeguimiento = mlAnt 'Actualiza el ID a consultar
    yTipoOperación = 0 'Solo consulta dado que se va a la historia de actividades anteriores
    ActFormulario
End If

End Sub

Private Sub cmdBotón_Click(Index As Integer)
If Accion(Index) Then
    Unload Me
End If
End Sub

Function Accion(Index As Integer)
Dim i As Integer, s As String, yy As Integer, Y As Long, adors As New ADODB.Recordset, yErr As Byte, d As Date
Dim l As Long, iNot As Integer, iOtor As Long, iAcc As Integer, iTipNot As Integer
Dim iValor As Single, iNoSan As Integer
bGuarda = False

On Error GoTo ErrorBloqueo:
If Index = 1 Then 'Cancelar
    Accion = True
    Exit Function
End If
If yTipoOperación = 0 Then
    Accion = True
    Exit Function
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
    Exit Function
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
                Exit Function
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
            Exit Function
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
            Exit Function
        End If
    Else
        If InStr(msCondonación, "|") > 0 Then
            msCondonación = ""
        End If
    End If
    If cmdSubirDocto.Visible Then 'Es obligatorio subir Docto
        'If ySubirDocto = 2 Then
            If cmdSubirDocto.BackColor <> cnVerde Then
                MsgBox "Debe subir el documento de notificación. Favor de subirlo", vbOKOnly + vbInformation, "Validación"
                Exit Function
            End If
        'Else
        '    If Len(sArchivoFTP) = 0 Then 'debe subir el documento a FTP
        '        MsgBox "Debe subir el documento de notificación. Favor de subirlo", vbOKOnly + vbInformation, "Validación"
        '        Exit Sub
        '    End If
            'If yVerificaDocto = 0 Then 'Debe verifica el documento a FTP
            '    MsgBox "Debe verificar si el documento se encuentra correctamente en estrados electrónicos. Favor de verificar", vbOKOnly + vbInformation, "Validación"
            '    Exit Sub
            'End If
        'End If
    Else
        If cmdVerificaDocto.Visible And (miTarea = 71 Or miTarea = 222) Then 'Debe verifica el documento a FTP
            If yVerificaDocto <= 0 Then 'Debe verificar el documento a FTP
                MsgBox "Debe verificar si el documento se encuentra correctamente en estrados electrónicos. Favor de verificar", vbOKOnly + vbInformation, "Validación"
                Exit Function
            End If
            d = DíasHábiles(Int(AhoraServidor), 1)
            If MsgBox("¿Está seguro que se subió correctamente el documento que se publicará el día hábil siguiente?: " & Format(d, gsFormatoFecha), vbYesNo + vbQuestion + vbDefaultButton2, "Confirmación") = vbNo Then
                Exit Function
            End If
        Else
            If Len(sArchivoFTP) > 0 Then
                sArchivoFTP = ""
            End If
        End If
    End If
    If txtAcuerdo.Visible Then 'Valida Auerdo u oficio de estrados y fecha de acuerdo
        If Len(Trim(txtAcuerdo.Text)) = 0 And chkacuerdo.Value = 0 Then
            MsgBox Mid(EtiAcuerdo.Caption, InStr(EtiAcuerdo.Caption, " ") + 1) & " es requerido", vbOKOnly + vbInformation, "Validación"
            Exit Function
        End If
        If miOtro = 11 Then 'Valida fecha acuerdo
            If Not IsDate(txtFAcuerdo2.Text) Then
                MsgBox Mid(etiFAcuerdo2.Caption, InStr(etiFAcuerdo2.Caption, " ") + 1) & " es requerido", vbOKOnly + vbInformation, "Validación"
                Exit Function
            End If
            msFAcuerdo = Format(CDate(txtFAcuerdo2.Text), gsFormatoFecha)
        End If
        msAcuerdo = txtAcuerdo.Text
        If adors.State Then adors.Close
        If InStr(EtiAcuerdo.Caption, "Oficio") > 0 Then 'Se trata de Oficio de Estrados Electrónicos, portanto verifica el No. de Oficio en Estrados
            adors.Open "select f_ee_verif_oficio('" & Replace(Trim(txtAcuerdo.Text), "'", "''") & "') from dual", gConSql, adOpenStatic, adLockReadOnly
            If adors(0) > 0 Then
                Call MsgBox("Existe ya el numero de oficio en estrados. Favor de verificar y cambiar por otro", vbOKOnly + vbInformation, "Validación de Número de Oficio")
                Exit Function
            End If
        Else 'Verifica el acuerdo
            adors.Open "select f_analisis_oficio(s.idana) from seguimientoacuerdos sa, seguimiento s where sa.acuerdo='" & Replace(msAcuerdo, "'", "''") & "' and sa.idseg=s.id and s.id<>" & mlSeguimiento, gConSql, adOpenStatic, adLockReadOnly
            If Not adors.EOF Then
                If MsgBox("Existe ya un acuerdo con ese número (oficio: " & adors(0) & "). ¿Está seguro de asignar el mismo número?", vbYesNo + vbQuestion + vbDefaultButton2, "Validación de Número de Acuerdo") = vbNo Then
                    Exit Function
                End If
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
    
    If txtcampo(5).Visible Then 'Valida fecha2 entraga/devolución
        If Len(txtcampo(5).Text) > 0 Then
            If Not IsDate(txtcampo(5).Text) Then
                MsgBox Mid(etiTexto(5).Caption, InStr(etiTexto(5).Caption, " ") + 1) & " es incorrecta", vbOKOnly + vbInformation, "Validación"
                Exit Function
            End If
            msFecha2 = Format(CDate(txtcampo(5).Text), gsFormatoFecha)
        Else
            msFecha2 = ""
        End If
    Else
        msFecha2 = ""
    End If
    If msFecha2 = "" And txtcampo(5).Visible Then 'Valida fecha 2
        MsgBox Mid(etiTexto(5).Caption, InStr(etiTexto(5).Caption, " ") + 1) & " es requerida", vbOKOnly + vbInformation, "Validación"
        Exit Function
    End If
    If txtcampo(3).Visible Then 'valida la captura de memo
        If Len(Trim(txtcampo(3).Text)) = 0 Then
            MsgBox "El número de memorando es requerido", vbOKOnly + vbInformation, "Validación"
            Exit Function
        End If
        msMemo = txtcampo(3).Text
        If adors.State Then adors.Close
        adors.Open "select f_analisis_oficio(s.idana) from seguimientomemos sm, seguimiento s where sm.memorando='" & Replace(msMemo, "'", "''") & "' and sm.idseg=s.id and s.id<>" & mlSeguimiento, gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            If MsgBox("Existe ya un memorando con ese número (oficio: " & adors(0) & "). ¿Está seguro de asignar el mismo número?", vbYesNo + vbQuestion + vbDefaultButton2, "Validación de Número de Memorando") = vbNo Then
                Exit Function
            End If
        End If
    End If
    If txtcampo(4).Visible Then 'valida fecha memo/oficio
        If Not IsDate(txtcampo(4).Text) Then
            MsgBox "La fecha del " & IIf(InStr(EtiAcuerdo.Caption, "Oficio") > 0, "Oficio", "Memorando") & " es requerida", vbOKOnly + vbInformation, "Validación"
            Exit Function
        End If
        If Len(msMemo) > 2 Then
            msMemo = msMemo & "|" & Format(txtcampo(4).Text, "dd/mm/yyyy") & "|"
        End If
        If InStr(EtiAcuerdo.Caption, "Oficio") > 0 Then 'Se trata de oficio de Estrados agrega la fecha en la variable ms acuerdo
            msAcuerdo = msAcuerdo & "||" & Format(txtcampo(4).Text, "dd/mmm/yyyy")
        End If
    Else
        msMemo = ""
    End If
    
    If Frame1.Visible Then 'Notificador
        If cmbNotificador.ListIndex < 0 Then
            Call MsgBox("Debe especificar el notificador", vbInformation + vbOKOnly, "Validación")
            Exit Function
        End If
        iNot = cmbNotificador.ItemData(cmbNotificador.ListIndex)
    End If
    
    If Frame2.Visible Then 'Nuevo Oficio y días otorgados
        If miMasivo = 0 Then
            If txtOficio.Visible And Len(txtOficio.Text) < 5 Then
                Call MsgBox("Debe especificar correctamente el nuevo oficio", vbInformation + vbOKOnly, "Validación")
                Exit Function
            End If
        Else
            If txtOficio.Visible And Val(txtOficio.Text) <= 0 Then
                Call MsgBox("Debe especificar correctamente el consecutivo para los nuevos oficios", vbInformation + vbOKOnly, "Validación")
                Exit Function
            End If
        End If
        If Val(txtDOtorgados.Text) = 0 Then
            Call MsgBox("Debe especificar correctamente los días otorgados", vbInformation + vbOKOnly, "Validación")
            Exit Function
        End If
        iOtor = Val(txtDOtorgados.Text)
    End If
    
    If Frame3.Visible Then 'Acuerdo y memorando de cierre
        If miMasivo = 0 Then
            If Len(txtAcuCierre.Text) < 5 Then
                Call MsgBox("Debe especificar correctamente el Acuerdo de Cierre oficio", vbInformation + vbOKOnly, "Validación")
                Exit Function
            End If
        Else
            If txtOficio.Visible And Val(txtOficio.Text) <= 0 Then
                Call MsgBox("Debe especificar correctamente el consecutivo para los Acuerdos de cierre", vbInformation + vbOKOnly, "Validación")
                Exit Function
            End If
        End If
        If Not IsDate(txtFAcuerdo.Text) Then
            Call MsgBox("La fecha del Acuerdo de cierre es un dato requerido", vbInformation + vbOKOnly, "Validación")
            Exit Function
        End If
        If Len(txtMemorando.Text) = 0 Then
            Call MsgBox("Debe especificar correctamente el memo de cierre", vbInformation + vbOKOnly, "Validación")
            Exit Function
        End If
        If Not IsDate(txtFMemo.Text) Then
            Call MsgBox("La fecha del Memorando de cierre es un dato requerido", vbInformation + vbOKOnly, "Validación")
            Exit Function
        End If
        iOtor = Val(txtDOtorgados.Text)
    End If
    
    If Frame4.Visible Then 'No sanción
        If cmbMotivoNoSan.ListIndex < 0 Then
            Call MsgBox("Se requiere el motivo de no Sanción, favor de capturarlo", vbInformation + vbOKOnly, "Validación")
            Exit Function
        End If
        miMotNoSan = cmbMotivoNoSan.ItemData(cmbMotivoNoSan.ListIndex)
    End If

    If adors.State Then adors.Close
    If miMasivo > 0 Then 'Registro masivo
        adors.Open "select max(valor) as idacc, max(tipo_not) as tipo_not from seg_accion where idact=" & miActividad & " and idtar=" & miTarea, gConSql, adOpenStatic, adLockReadOnly
        iValor = IIf(IsNull(adors(0)), 0, adors(0))
        iTipNot = IIf(IsNull(adors(1)), 0, adors(1))
        If iValor > 0 Then 'Se encontró la acción segun actividad y tarea
            If iTipNot > 0 Then 'Verifica si la acción debe guardarse según el tipo_notificación en caso que haya varios tipo_not
                If adors.State Then adors.Close
                adors.Open "select sum(case when tipo_not in (2,3,4) then 1 else 0 end) as suma, sum(case when tipo_not=" & iTipNot & " then 1 else 0 end) as sumtipnot from seg_accion where idact=" & miActividad & " and valor between " & Int(iValor) & " and " & (Int(iValor) + 1 - 0.1), gConSql, adOpenStatic, adLockReadOnly
            End If
            If adors(0) > 1 Then
                If adors(1) <= 1 Then
                    iValor = Int(iValor)
                End If
            End If
        Else
            MsgBox "Verificar con el administrador de sistema Excepción: 2020 - no hay acción Masiva relacionada con la tarea", vbOKOnly + vbCritical, "Validación"
            Unload Me
            Exit Function
        End If
        If adors(0) > 0 Then
            If MsgBox("Está seguro de realizar el registro masivo de esta actividad", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmación") = vbNo Then
                Exit Function
            End If
            If adors.State Then adors.Close
            If miDatosAdi > 0 Then
                If Frame1.Visible Then
                    adors.Open "{call P_GuardaDatosSEG(" & iValor & ",'" & msSeg & "','" & Format(CDate(txtcampo(1).Text), "dd/mm/yyyy hh:mm:ss") & "','" & Replace(txtcampo(2).Text, "'", "''") & "',1," & iNot & ",0,0,0,'','',''," & giUsuario & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
                ElseIf Frame2.Visible Then
                    adors.Open "{call P_GuardaDatosSEG(" & iValor & ",'" & msSeg & "','" & Format(CDate(txtcampo(1).Text), "dd/mm/yyyy hh:mm:ss") & "','" & Replace(txtcampo(2).Text, "'", "''") & "',2,0," & Val(txtOficio.Text) & "," & Val(txtDOtorgados.Text) & ",0,'','',''," & giUsuario & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
                ElseIf Frame3.Visible Then
                    adors.Open "{call P_GuardaDatosSEG(" & iValor & ",'" & msSeg & "','" & Format(CDate(txtcampo(1).Text), "dd/mm/yyyy hh:mm:ss") & "','" & Replace(txtcampo(2).Text, "'", "''") & "',3,0,0,0," & Val(txtAcuCierre.Text) & ",'" & txtFAcuerdo.Text & "','" & txtMemorando.Text & "','" & txtFMemo.Text & "'," & giUsuario & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
                End If
            Else
                adors.Open "{call P_GuardaDatosSEG(" & iValor & ",'" & msSeg & "','" & Format(CDate(txtcampo(1).Text), "dd/mm/yyyy hh:mm:ss") & "','" & Replace(txtcampo(2).Text, "'", "''") & "',0,0,0,0,0,'','',''," & giUsuario & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
            End If
            If adors(0) > 0 Then
                Call MsgBox("Se ingresó correctamente la actividad. " & adors(1), vbOKOnly + vbInformation, "")
            Else
                Call MsgBox("No se registro ninguna actividad. " & adors(1), vbOKOnly + vbInformation, "")
                Exit Function
            End If
                
        End If
    Else 'Registro ordinario con un solo asunto
        'adors.Open "{call p_seguimientoguardadatos(" & mlSeguimiento & "," & mlAnt & "," & mlAnálisis & ",'" & Format(CDate(txtCampo(1).Text), "dd/mm/yyyy hh:mm:ss") & "'," & miActividad & "," & miTarea & "," & miResponsable & "," & giUsuario & "," & miDesenlace & ",'" & Replace(txtCampo(2).Text, "'", "''") & "','" & msDoctos & "','" & msActsProg & "'," & l & ",'" & msSanción & "','" & msCondonación & "','" & msMemo & "')}", gConSql, adOpenForwardOnly, adLockReadOnly
        If miDatosAdi > 0 Then
            adors.Open "{call p_seguimientoguardadatosDA(" & mlSeguimiento & "," & mlAnt & "," & mlAnálisis & ",'" & Format(CDate(txtcampo(1).Text), "dd/mm/yyyy hh:mm:ss") & "'," & miActividad & "," & miTarea & "," & miResponsable & "," & giUsuario & "," & miDesenlace & ",'" & Replace(txtcampo(2).Text, "'", "''") & "','" & msDoctos & "','" & msActsProg & "'," & miDatosAdi & "," & iNot & ",'" & txtOficio.Text & "'," & iOtor & ",'" & Replace(txtAcuCierre.Text, "'", "''") & "','" & Replace(txtFAcuerdo.Text, "'", "''") & "','" & txtMemorando.Text & "','" & txtFMemo.Text & "' )}", gConSql, adOpenForwardOnly, adLockReadOnly
        Else
            'MsgBox ("MSFECHA2: " & msFecha2)
            adors.Open "{call p_seguimientoguardadatos4(" & mlSeguimiento & "," & mlAnt & "," & mlAnálisis & ",'" & Format(CDate(txtcampo(1).Text), "dd/mm/yyyy hh:mm:ss") & "'," & miActividad & "," & miTarea & "," & miResponsable & "," & giUsuario & "," & miDesenlace & ",'" & Replace(txtcampo(2).Text, "'", "''") & "','" & msDoctos & "','" & msActsProg & "','" & msAcuerdo & "','" & msFAcuerdo & "','" & msSanción & "','" & msCondonación & "','" & msMemo & msFecha2 & "|'," & miMotNoSan & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
        End If
    End If
    If adors(0) < 0 Then
        MsgBox "No se realizó el Alta del avance." & adors(1), vbOKOnly + vbInformation, ""
        Exit Function
    Else
        mlSeguimiento = adors(0)
        If miTarea = 71 Then 'Estrados Electrónicos autorizado
            l = adors(0)
            If adors.State Then adors.Close
            adors.Open "select f_ee_obtieneinf(" & l & ") from dual", gConSql, adOpenStatic, adLockReadOnly
            If Not adors.EOF Then
                Call MsgBox("Información del estrado electrónico a publicar:" & Chr(13) & Chr(10) & adors(0), vbOKOnly + vbInformation, "")
            Else
                Call MsgBox("No se localizó el Estrado Electrónico por publicar", vbOKOnly + vbInformation, "")
            End If
        End If
    End If
    gs = "OK"
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
'Me.Hide
Accion = True
Exit Function
ErrorBloqueo:
If gConSql.Errors.Count > 0 Then
    yErr = MsgBox("Error: " + Err.Description + ". vuelva a intentar", vbOKOnly + vbCritical, "Error no esperado (" + Str(IIf(Err.Number < 0, gConSql.Errors(gConSql.Errors.Count - 1).Number, Err.Number)) + ")")
Else
    yErr = MsgBox("Error: " & Err.Description & ". vuelva a intentar", vbOKOnly + vbCritical, "Error no esperado (" & IIf(Err.Number > 0, Err.Number, "???") & ")")
End If
If yErr = vbCancel Or yErr = vbAbort Then
    Exit Function
ElseIf yErr = vbRetry Then
    Resume
ElseIf yErr = vbIgnore Then
    Resume Next
End If
End Function

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
    'If InStr(cmdSanción.Caption, "No ") > 0 Then
    '    .Check1.Enabled = False
    '    .Check1.Value = 0
    'End If
    .Show vbModal
    If gs <> "cancelar" Then
        msSanción = gs
    End If
    bOprimióTecla = True
    HabilitaAceptar
End With
End Sub

Private Sub cmdSig_Click()
Dim i As Integer
If Accion(0) Then 'Guarda datos o aceta la consulta
    yUnico = 0 'Actualiza variable de actulización por única vez
    For i = 0 To 2
        If opcAct(i).Value Then
            mlSeguimiento = Val(opcAct(i).Tag)
        End If
    Next
    yTipoOperación = 0 'Solo consulta dado que se va a la historia de actividades anteriores
    ActFormulario
End If

End Sub

Private Sub cmdSubirDocto_Click()
Dim sArchivo As String, adors As New ADODB.Recordset, sArchivoOK As String
Dim s As String, i As Integer, i2 As Integer, EnvString As String
Dim yErr As Byte
Dim strFic As String
On Error GoTo salir:
If ySubirDocto = 2 Then 'SINE
    If adors.State Then adors.Close
    adors.Open "select f_url_conexion(1," & giUsuario & "," & mlAnt & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    If Len(adors(0)) > 0 Then
        gsWWW = adors(0)
    Else
        MsgBox "La cadena de conexión al SINE viene vacia Favor de reportarlo al administrador del Sistema  (Miguel Ext.6032)"
        Exit Sub
    End If
    If adors.State Then adors.Close
    adors.Open "select f_seguimiento_idpro(" & mlAnt & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    If adors(0) = 2 Then 'Emplazamiento
        If adors.State Then adors.Close
        adors.Open "select f_analisis_oficio(" & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    ElseIf adors(0) = 3 Then 'Instrucción
        If adors.State Then adors.Close
        adors.Open "select f_analisis_oficio(" & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    ElseIf adors(0) = 4 Then 'Sanción
        If adors.State Then adors.Close
        adors.Open "select f_analisis_san_oficio(" & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    ElseIf adors(0) = 9 Then 'Condonación
        If adors.State Then adors.Close
        adors.Open "select f_analisis_cond_oficio(" & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    Else
        i = 200
    End If
    If Len(adors(0)) > 0 And i <> 200 Then
        gsWWW = gsWWW & "&oficio='" & adors(0) & "'"
    End If
    'Verifica si puede abrir con chrome que es dónde funciona correctamente
    i = 1
    Do 'revisa en variables de entorno está path para executar chrome
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
    If adors.State Then adors.Close
    adors.Open "select f_sine_verif_docto(" & mlAnt & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    If adors(0) > 0 Then
        cmdSubirDocto.BackColor = cnVerde
    Else
        cmdSubirDocto.BackColor = cnRojo
    End If

Else 'Estrados
    If adors.State Then adors.Close
    adors.Open "select f_url_conexion(2," & giUsuario & "," & mlAnt & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    If Len(adors(0)) > 0 Then
        gsWWW = adors(0)
    Else
        MsgBox "La cadena de conexión a Estrados Electrónicos viene vacia Favor de reportarlo al administrador del Sistema  (Miguel Ext.6032)"
        Exit Sub
    End If
    With Browser
        .yÚnicavez = 0
        .Caption = "Estrados Electrónicos"
        .Show vbModal
    End With
    If adors.State Then adors.Close
    adors.Open "select f_ee_verif_docto(" & mlAnt & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    If Len(adors(0)) > 0 Then
        cmdSubirDocto.BackColor = cnVerde
    Else
        cmdSubirDocto.BackColor = cnRojo
    End If
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
Dim i As Integer
yVerificaDocto = 1
If yActEstrados Then 'Verifica documento de estrados electrónicos
    Dim adors As New ADODB.Recordset
    If adors.State Then adors.Close
    adors.Open "select f_url_conexion(5,0," & mlAnt & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    If Len(adors(0)) > 0 And i <> 200 Then
        gsWWW = adors(0)
    End If
    With Browser
        .yÚnicavez = 0
        .Caption = "Documento por validar de ESTRADOS ELECTRÓNICOS"
        .Show vbModal
    End With
    Exit Sub
End If
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

Private Sub Combo1_Change()

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
ActFormulario
End Sub

Sub ActFormulario()
Dim Y As Byte, s As String, adors As New ADODB.Recordset, i As Integer, ss As String, yy As Byte
Dim i1 As Integer
Dim s2 As String
miOtro = 0
ySubirArchivo = 0
'LicenseManager.SetLicenseKey ("7C19F6E1416C480FD3CBB509133177EE9F9F5113722D31910A08BFCE618A52807D2793EA40EF65D2BE5FD6A3D2F555C6523AB270F56522235C68936422CCE57B9A71F3A22A8B34E40B1FD1E173D6634BC954F9EB6DBAF523F064F4FD75B910C749155FE6E74C1343E621FF459619D6C8C9008580C91EB5799498C718050C9AC1EC657031FDB6A9FF573EE6DB9BF0C7CBB672EC305696ABBC8D682DB762236DF711950AB629FF589FA86B0759EEBE9F1155DD1302AB276F879C9074B0F20D5EC3444608772A9A845A2A7CBEDAC15DA1A0050DC7A687788CF1E1ECC040D50AAEC6957F385BAA7C177307D5DB091711D87B7E678459AD3FABF39E089073DE762EC3")
bOprimióTecla = True
If yUnico = 0 Then
    'Bloquea controles en caso de consulta
    If miMasivo = 0 Then
        TreeView1.Enabled = True
        cmdAnt.Visible = True
        cmdSig.Visible = True
        Frame5.Visible = True
    Else
        cmdAnt.Visible = False
        cmdSig.Visible = False
        Frame5.Visible = False
    End If
    If yTipoOperación = 0 Then 'Consulta
        TreeView1.Enabled = False
        TreeView2.Enabled = False
        TreeView3.Enabled = False
        txtcampo(1).Locked = True
        txtcampo(2).Locked = True
        cmdBotón(0).Enabled = True
    Else ' los habilita
        TreeView2.Enabled = True
        TreeView3.Enabled = True
        txtcampo(1).Locked = False
        txtcampo(2).Locked = False
        'cmdBotón(0).Enabled = True
        If yTipoOperación = 1 Then 'inicializavar
            msSanción = ""
            msCondonación = ""
            msAcuerdo = ""
            msFAcuerdo = ""
            msFecha2 = ""
        End If
    End If
    If mlSeguimiento > 0 Then
        'valores del seguimiento
        If adors.State Then adors.Close
        adors.Open "select s.*,f_analisis_expediente(idana) as Exp, paq_conceptos.actividad(idact) as activ from seguimiento s where id=" & mlSeguimiento, gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            mlAnt = adors!idant
            mlAnálisis = adors!idana
            mdFecha = adors!FECHA
            miResponsable = adors!idres
            miActividad = adors!idact
            miTarea = adors!idtar
            miDesenlace = IIf(IsNull(adors!iddes), 0, adors!iddes)
            If miMasivo = 0 Then
                txtExp.Text = adors!Exp
            End If
            If Not IsNull(adors!activ) Then
                txtcampo(0).Text = adors!activ
            End If
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
        adors.Open "select acuerdo,fecha from seguimientoacuerdos where idseg=" & mlSeguimiento, gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            msAcuerdo = adors(0)
            msFAcuerdo = IIf(IsNull(adors(1)), "", Format(adors(1), gsFormatoFecha))
            'txtAcuerdo.Text = msAcuerdo
        Else
            msAcuerdo = ""
            msFAcuerdo = ""
        End If
        'No y fecha de memorando
        If adors.State Then adors.Close
        adors.Open "select memorando,fecha,fecha2 from seguimientomemos where idseg=" & mlSeguimiento, gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            msMemo = adors(0) & "|" & Format(adors(1), "dd/mm/yyyy") & "|"
            If Not IsNull(adors(2)) Then
                msFecha2 = Format(adors(2), "dd/mm/yyyy")
            End If
        Else
            msMemo = ""
            msFecha2 = ""
        End If
        'Datos Sanción y Condonación
        If adors.State Then adors.Close
        adors.Open "select f_sancion_sanxcau(" & mlSeguimiento & "),f_cond_condxcau(" & mlSeguimiento & ") from dual", gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            msSanción = adors(0)
            msCondonación = adors(1)
        Else
            msSanción = ""
            msCondonación = ""
        End If
'        'Datos Condonación
'        If adors.State Then adors.Close
'        adors.Open "select f_cond_condxcau(" & mlSeguimiento & ") from dual", gConSql, adOpenStatic, adLockReadOnly
'        If Not adors.EOF Then
'            'msCondonación = adors!OFICIO & "|" & Format(adors!FECHA, "dd/mm/yyyy") & "|" & adors!Porcentaje & "|"
'            msCondonación = adors(0)
'        Else
'            msCondonación = ""
'        End If

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
            adors.Open "select f_actividad(" & miActividad & "),f_analisis_expediente(" & mlAnálisis & ") as Exp from dual", gConSql, adOpenStatic, adLockReadOnly
            If adors.EOF Then
                s = ""
            Else
                If IsNull(adors(0)) Then
                    s = ""
                Else
                    s = adors(0)
                End If
                If miMasivo = 0 Then
                    txtExp.Text = adors(1)
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
    If txtcampo(5).Visible Then
        txtcampo(5).Text = msFecha2
    End If
    yUnico = 100
    If miMasivo = 0 Then 'Actualiza botones de Act Anterior y siguiente
        cmdAnt.Enabled = True
        If mlAnt = mlSeguimiento Then
            cmdAnt.Enabled = False
        End If
        If adors.State Then adors.Close
        adors.Open "select id,paq_conceptos.actividad(idact) from seguimiento where idant=" & mlSeguimiento, gConSql, adOpenStatic, adLockReadOnly
        i = 0
        If adors.EOF Then
            If TreeView1.Nodes.Count = 0 Or Not TreeView1.Visible Then 'Se verifica que no haya actividades por programar dado que ya está al final del proceso
                'No se cuenta con actividades posteriores
                Frame5.Visible = False
                cmdSig.Enabled = False
            Else
                ActOpcActsSig_Prog
            End If
        Else
            Frame5.Visible = True
            cmdSig.Enabled = True
            Do While Not adors.EOF
                opcAct(i).Tag = adors(0)
                opcAct(i).Visible = True
                opcAct(i).Caption = adors(1)
                i = i + 1
                adors.MoveNext
            Loop
            Do While i < 3
                opcAct(i).Visible = False
                i = i + 1
            Loop
            opcAct(0).Value = True
        End If
    ElseIf cmdAnt.Visible Then
        cmdAnt.Visible = False
        cmdSig.Visible = False
    End If
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

'Actualiza opciones de actividad siguiente según la actividad programada o opciones de actividad
Sub ActOpcActsSig_Prog()
Dim i As Integer
For i = 0 To TreeView1.Nodes.Count - 1
    If TreeView1.Nodes(i).Checked Then
        Exit For
    End If
Next
If i = 0 Then 'No se tiene actividades por programar.
    Frame5.Visible = False
    cmdSig.Enabled = False
    Exit Sub
End If
'Verifica si se encontró una opción seleccionada de la programación
If i >= TreeView1.Nodes.Count Then 'Ni uno Mete todas las opciones
    For i = 0 To TreeView1.Nodes.Count - 1
        opcAct(i).Caption = TreeView1.Nodes(i).Text
        opcAct(i).Tag = "0"
    Next
Else 'Solo opciones selecionadas
    For i = 0 To TreeView1.Nodes.Count - 1
        If TreeView1.Nodes(i).Checked Then
            opcAct(i).Caption = TreeView1.Nodes(i).Text
            opcAct(i).Tag = "0"
        End If
    Next
End If
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
If miTarea = 198 Then
    If InStr(LCase(txtcampo(2).Text), "oficio: ") = 0 Then
        cmdBotón(0).Enabled = False
        Exit Sub
    End If
End If
cmdBotón(0).Enabled = (ComboDesenlaces.ListCount = 0 Or ComboDesenlaces.ListIndex >= 0) And (i <= TreeView3.Nodes.Count Or TreeView3.Nodes.Count = 0)
cmdAnt.Enabled = cmdBotón(0).Enabled
cmdSig.Enabled = cmdBotón(0).Enabled
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
    'ySoloConsulta = 0
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

Private Sub Frame12_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame6_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Text1_Change()

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
miDatosAdi = -1
'Verifica datos adicionales
If adors.State Then adors.Close
adors.Open "select max(datosadi) from seg_accion where idact=" & miActividad & " and idtar=" & miTarea, gConSql, adOpenStatic, adLockReadOnly
If adors(0) = 1 Then
    miDatosAdi = 1
    Frame1.Visible = True
    If Frame2.Visible Then Frame2.Visible = False
    If Frame3.Visible Then Frame3.Visible = False
    LlenaCombo cmbNotificador, "select idusinot as id, paq_conceptos.responsable(idusinot) as notificador from notificadores where fecha_baja is null", "", True
    Call CargaDatosAdi(1)
ElseIf adors(0) = 2 Or adors(0) = 3 Then
    miDatosAdi = adors(0)
    If Frame1.Visible Then Frame1.Visible = False
    Frame2.Visible = True
    If Frame3.Visible Then Frame3.Visible = False
    Call CargaDatosAdi(2)
ElseIf adors(0) = 4 Then
    miDatosAdi = 4
    Frame3.Visible = True
    If Frame2.Visible Then Frame2.Visible = False
    If Frame1.Visible Then Frame1.Visible = False
    Call CargaDatosAdi(4)
Else
    If Frame1.Visible Then Frame1.Visible = False
    If Frame2.Visible Then Frame2.Visible = False
    If Frame3.Visible Then Frame3.Visible = False
End If
If miTarea = 33 Then 'No sanción
    Frame4.Visible = True
    Call LlenaCombo(cmbMotivoNoSan, "select idmotnosan,motivo from motivosnosan", "", True)
    If mlSeguimiento > 0 Then
        If adors.State Then adors.Close
        adors.Open "select max(idmotnosan) from seguimientomotnosan where idseg=" & mlSeguimiento, gConSql, adOpenStatic, adLockReadOnly
        If adors(0) > 0 Then
            i = BuscaCombo(cmbMotivoNoSan, adors(0), True)
            If i >= 0 Then
                cmbMotivoNoSan.ListIndex = i
            End If
        End If
    End If
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

Private Sub txtFAcuerdo2_Change()
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

Private Sub txtFAcuerdo_KeyPress(KeyAscii As Integer)
KeyAscii = TeclaOprimida(txtFAcuerdo, KeyAscii, txtFAcuerdo.Tag, False)
bOprimióTecla = True
bAceptar = False
HabilitaAceptar
yHabilita = 200
End Sub

Private Sub txtFAcuerdo_LostFocus()
If IsDate(txtFAcuerdo.Text) Then
    txtFAcuerdo.Text = Format(CDate(txtFAcuerdo.Text), gsFormatoFecha)
End If
End Sub

Private Sub txtFAcuerdo2_KeyPress(KeyAscii As Integer)
KeyAscii = TeclaOprimida(txtFAcuerdo2, KeyAscii, txtFAcuerdo2.Tag, False)
bOprimióTecla = True
bAceptar = False
HabilitaAceptar
yHabilita = 200
End Sub

Private Sub txtFAcuerdo2_LostFocus()
If IsDate(txtFAcuerdo2.Text) Then
    txtFAcuerdo2.Text = Format(CDate(txtFAcuerdo2.Text), gsFormatoFecha)
End If
End Sub

Private Sub txtFMemo_KeyPress(KeyAscii As Integer)
KeyAscii = TeclaOprimida(txtFMemo, KeyAscii, txtFMemo.Tag, False)
bOprimióTecla = True
bAceptar = False
HabilitaAceptar
yHabilita = 200
End Sub

Private Sub txtFMemo_LostFocus()
If IsDate(txtFMemo.Text) Then
    txtFMemo.Text = Format(CDate(txtFMemo.Text), gsFormatoFecha)
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
If giUsuEsp > 0 Then
    adors.Open "select f_seguimiento_query2(1," & mlAnálisis & "," & miActividad & " ) from dual", gConSql, adOpenStatic, adLockReadOnly
Else
    adors.Open "select f_seguimiento_query(1," & mlAnálisis & "," & miActividad & " ) from dual", gConSql, adOpenStatic, adLockReadOnly
End If
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
adors.Open "select f_seguimiento_query2(4," & mlAnálisis & "," & iTarea & ") from dual", gConSql, adOpenStatic, adLockReadOnly
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
adors.Open "select f_seguimiento_query2(3," & mlAnálisis & "," & iTarea & ") from dual", gConSql, adOpenStatic, adLockReadOnly
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
adors.Open "select f_seguimiento_query2(2," & mlAnálisis & "," & iTarea & "," & mlAnt & ",'" & Format(txtcampo(1).Text, "dd/mm/yyyy") & "') from dual", gConSql, adOpenStatic, adLockReadOnly
TreeView1.Nodes.Clear
If Not adors.EOF Then
    s = adors(0)
    i = Val(Mid(s, InStrRev(s, "|") + 1))
    s = Mid(s, 1, InStrRev(s, "|") - 1)
    Call CargaDatosArbolVariosNiveles(TreeView1, s, i, False, True, True)
    For i = 1 To TreeView1.Nodes.Count
        If InStr(TreeView1.Nodes(i).Text, "Resp.: ") > InStr(TreeView1.Nodes(i).Text, " (") Then
            TreeView1.Nodes(i).Checked = True
        End If
    Next
End If
If TreeView1.Nodes.Count > 0 Then 'Hay posibilidad de generar más actividades posteriormente
    If Not cmdSig.Enabled Then cmdSig.Enabled = True
    If mlSeguimiento <= 0 Then
    '''falta programación
    End If
Else
    If cmdSig.Enabled Then cmdSig.Enabled = False
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
Dim adors As New ADODB.Recordset, s As String, i As Integer, iPro As Integer
Dim l As Long
Dim frm As New SelProceso
If iTarea = 198 Then 'Corresponde a Oficio por inicio del proceso de estrados electrónicos cuando se autoriza
    'etiTexto(3).Visible = False
    'etiTexto(4).Visible = True
    'txtcampo(4).Visible = True
    'EtiAcuerdo.Visible = True
    'txtAcuerdo.Visible = True
    'EtiAcuerdo.Caption = "7. No. Oficio"
    'etiTexto(4).Caption = "8. Fecha del Oficio"
    'txtAcuerdo.Locked = True
    'txtcampo(4).Locked = True
    adors.Open "select f_act_idpro(f_seguimiento_idact(f_ee_idseg_inipro(" & mlAnt & "))) from dual", gConSql, adOpenStatic, adLockReadOnly
    iPro = adors(0)
    adors.Close
    adors.Open "select f_ee_verif_expediente(f_analisis_expediente(" & mlAnálisis & "),1) as Num,f_analisis_expediente(" & mlAnálisis & ") as Exp,f_ee_verif_expediente(f_analisis_expediente(" & mlAnálisis & "),2) as ID from dual", gConSql, adOpenStatic, adLockReadOnly
    If adors(0) <= 0 Then
        MsgBox "El expediente " & adors(1) & " no existe en el módulo de estrados. Debe darse de alta antes de especificar esta actividad", vbOKOnly, "Validación de estrados electrónicos"
        Exit Sub
    End If
    If adors(0) = 1 Then
        l = adors(2)
    Else
        gs = "{call P_ee_XExp(f_analisis_expediente(" & mlAnálisis & "))}"
        frm.Caption = "Favor de seleccionar el Oficio correspondiente"
        frm.Show vbModal
        If Val(gs) > 0 Then
            l = Val(gs)
        End If
    End If
    If l > 0 Then
        If adors.State Then adors.Close
        adors.Open "select no_oficio,to_char(fecha_oficio,'dd/mon/yyyy') from estradosesp where id=" & l, gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            txtcampo(2).Text = "Oficio: " & adors(0) & " Fecha_Oficio: " & adors(1)
        Else
            l = -1
        End If
    Else
        l = -1
    End If
    If l < 0 Then 'No se ha elegido Oficio
        Call MsgBox("Debe especificar un oficio para poder continuar.", vbOKOnly, "validación")
        cmdBotón(0).Enabled = False
        Exit Sub
    Else
        If Not cmdBotón(0).Enabled Then cmdBotón(0).Enabled = True
    End If
    
    'i = InStr(adors(0), "|")
    'If i > 0 Then 'obtiene Oficio y fecha del oficio
    '    txtAcuerdo.Text = Mid(adors(0), 1, i - 1)
    '    txtcampo(4).Text = Mid(adors(0), i + 1)
    'End If
    cmdSubirDocto.Visible = False
    Exit Sub
End If
adors.Open "select idotr from relaciónactividadtarea where idact=" & miActividad & " and idtar=" & iTarea, gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    miOtro = adors(0)
Else
    miOtro = 0
End If
If miOtro = 11 Then 'Acuerdo y Fecha acuerdo
    If Not txtAcuerdo.Visible Then
        txtAcuerdo.Visible = True
        EtiAcuerdo.Visible = True
        txtFAcuerdo2.Visible = True
        etiFAcuerdo2.Visible = True
        chkacuerdo.Visible = True
    End If
'    If adors.State Then adors.Close
'    adors.Open "select to_char(sysdate,'yyyy') from dual", gConSql, adOpenStatic, adLockReadOnly
'    If Not IsNull(adors(0)) Then
'        i = adors(0)
'    Else
'        i = Year(Date)
'    End If
'    If Len(Trim(txtAcuerdo.Text)) = 0 And yTipoOperación <> 0 Then
'        'txtAcuerdo.Text = "ACUERDO/DAS/" & i & "/"
'        txtAcuerdo.Text = ""
'        'txtAcuerdo.Text = "AUTOMÁTICO"
'    End If
    txtAcuerdo.Text = ""
    Exit Sub
ElseIf miOtro = 1 Then
    If Not txtAcuerdo.Visible Then
        txtAcuerdo.Visible = True
        EtiAcuerdo.Visible = True
        txtFAcuerdo2.Visible = True
        etiFAcuerdo2.Visible = True
        chkacuerdo.Visible = True
    End If
'    If adors.State Then adors.Close
'    adors.Open "select to_char(sysdate,'yyyy') from dual", gConSql, adOpenStatic, adLockReadOnly
'    If Not IsNull(adors(0)) Then
'        i = adors(0)
'    Else
'        i = Year(Date)
'    End If
'    If Len(Trim(txtAcuerdo.Text)) = 0 And yTipoOperación <> 0 Then
'        'txtAcuerdo.Text = "ACUERDO/DAS/" & i & "/"
'        txtAcuerdo.Text = ""
'        'txtAcuerdo.Text = "AUTOMÁTICO"
'    End If
    txtAcuerdo.Text = ""
    txtFAcuerdo2.Text = ""
    Exit Sub
End If
If txtAcuerdo.Visible Then
    txtAcuerdo.Visible = False
    EtiAcuerdo.Visible = False
    txtFAcuerdo2.Visible = False
    etiFAcuerdo2.Visible = False
    chkacuerdo.Visible = False
End If
End Sub

'Carga datos de la sanción
Private Sub CargaSanción(iTarea As Integer)
Dim adors As New ADODB.Recordset, s As String, i As Byte
adors.Open "select count(*),max(idotr) from relaciónactividadtarea where idact=" & miActividad & " and idtar=" & iTarea & " and idotr in (2,8)", gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    If Not cmdSanción.Visible Then
        cmdSanción.Visible = True
        If adors(1) = 8 Then
            cmdSanción.Caption = "No Sanción"
        End If
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
adors.Open "select f_Tarea_subirDocto(" & miActividad & "," & iTarea & "),f_Tarea_verificaDocto(" & miActividad & "," & iTarea & "),f_Actividad_estrados(" & miActividad & ") from dual", gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    ySubirArchivo = 1
End If
If adors(1) > 0 Then
    yVerificaArchivo = 1
End If
If adors(2) > 0 Then
    yActEstrados = 1
End If
If ySubirArchivo > 0 Then 'Subir documento para estados electrónicos o Subir documento para SINE
    If Not cmdSubirDocto.Visible Then
        cmdSubirDocto.Visible = True
    End If
    If yActEstrados > 0 Then 'Estrados
        If yVerificaArchivo > 0 Then
            cmdVerificaDocto.Visible = True
        Else
            cmdVerificaDocto.Visible = False
        End If
        sArchivoFTP = ""
        'yVerificaDocto = 0
        ySubirDocto = 1
    Else 'SINE
        ySubirDocto = 2
    End If
    Exit Sub
Else
    If yVerificaArchivo > 0 Then
        cmdVerificaDocto.Visible = True
    End If
End If
If cmdSubirDocto.Visible Then
    cmdSubirDocto.Visible = False
    If yVerificaArchivo > 0 Then
        cmdVerificaDocto.Visible = False
    End If
End If
End Sub

'Carga información del archivo por subir vía FTP a estrados electrónicos
Private Sub CargaVerificaArchivoFTP(iTarea As Integer)
Dim adors As New ADODB.Recordset, s As String, i As Byte
adors.Open "select count(*) from relaciónactividadtarea where idact=" & miActividad & " and idtar=" & iTarea & " and idotr in (3,4)", gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    'If Not cmdVerificaDocto.Visible Then
        'CmdVerificaDocto.Visible = True
    'End If
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
    If yActEstrados Then
        cmdVerificaDocto.Visible = False
    End If
End If
End Sub

'Carga Memo y fecha Memo cuando en la tabla RelaciónActividadTarea idotr=5
Private Sub CargaMemo(iTarea As Integer)
Dim adors As New ADODB.Recordset, s As String, i As Integer
'adors.Open "select count(*) from relaciónactividadtarea where idact=" & miActividad & " and idtar=" & iTarea & " and idotr=5", gConSql, adOpenStatic, adLockReadOnly
If miOtro = 5 Or miOtro = 12 Or miOtro = 13 Then 'tres opciones
    'If Not txtAcuerdo.Visible Then
        txtcampo(3).Visible = True
        etiTexto(3).Visible = True
        txtcampo(4).Visible = True
        etiTexto(4).Visible = True
    'End If
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
Else
    If txtcampo(3).Visible Then
        txtcampo(3).Visible = False
        etiTexto(3).Visible = False
        txtcampo(4).Visible = False
        etiTexto(4).Visible = False
    End If
End If
If miOtro = 12 Or miOtro = 13 Then
    If miOtro = 12 Then 'Fecha entrega / fecha devolución
        etiTexto(5).Caption = "Fecha entrega"
    Else
        etiTexto(5).Caption = "Fecha devolución"
    End If
    If Not txtcampo(5).Visible Then
        txtcampo(5).Visible = True
        etiTexto(5).Visible = True
    End If
Else
    txtcampo(5).Visible = False
    etiTexto(5).Visible = False
End If
End Sub

'Carga Datos Adicionales
Private Sub CargaDatosAdi(iTipDato As Integer)
Dim adors As New ADODB.Recordset, s As String, i As Integer
If cmbNotificador.ListIndex < 0 And iTipDato = 1 Then
    adors.Open "select max(idusinot) from registroxif_notif where idregxif=f_analisis_idregxif(" & mlAnálisis & ")", gConSql, adOpenStatic, adLockReadOnly
    If adors(0) > 0 Then
        i = BuscaCombo(cmbNotificador, adors(0), True)
        If i > 0 Then
            cmbNotificador.ListIndex = i
        End If
    End If
    Exit Sub
ElseIf iTipDato = 2 Then
    If adors.State > 0 Then adors.Close
    adors.Open "select max(oficio), max(otorgados) from análisis where id=" & mlAnálisis, gConSql, adOpenStatic, adLockReadOnly
    If Len(adors(0)) > 0 Then
        If miMasivo > 0 And miAccion = 0 Then
            txtOficio.Text = "10" '10 por defecto
            txtDOtorgados.Tag = "n" 'Acepta puros números
            lblNvoOficio.Caption = "Consecutivo inicial Nvos Oficios: "
            txtOficio.Text = ""
        Else
            txtOficio.Text = adors(0)
            txtDOtorgados.Text = IIf(IsNull(adors(1)), "", adors(1))
        End If
    Else
        adors.Open "select max(ofi_emp) from grupousuarios where id=f_usuariosis_idgpo(" & giUsuario & ")", gConSql, adOpenStatic, adLockReadOnly
        If miMasivo > 0 And miAccion = 0 Then
            txtOficio.Text = "10" '10 por defecto
            txtDOtorgados.Tag = "n" 'Acepta puros números
            lblNvoOficio.Caption = "Consecutivo inicial Nvos Oficios: "
            txtOficio.Text = ""
        Else
            If Len(adors(0)) > 0 Then
                txtOficio.Text = adors(0)
            End If
        End If
    End If
ElseIf iTipDato = 4 Then
    If adors.State > 0 Then adors.Close
    adors.Open "select max(a.acuerdo), max(to_char(a.fecha_acuerdo,'dd/mm/yyyy')), max(m.memorando), max(to_char(m.fecha,'dd/mm/yyyy')) from seguimientoadocie a inner join emp_memo_cierre m on a.idmem=m.idmem where a.idseg=" & mlSeguimiento, gConSql, adOpenStatic, adLockReadOnly
    If Len(adors(0)) > 0 Then
        If miMasivo > 0 And miAccion = 0 Then
            txtFAcuerdo.Text = ""
            txtFMemo.Text = ""
            txtAcuCierre.Tag = "n" 'Acepta puros números
            lblAdoCierre.Caption = "Consecutivo inicial Acuerdos Cierre: "
            txtMemorando.Text = ""
        Else
            txtAcuCierre.Text = IIf(IsNull(adors(0)), "", adors(0))
            txtFAcuerdo.Text = IIf(IsNull(adors(1)), "", adors(1))
            txtMemorando.Text = IIf(IsNull(adors(2)), "", adors(2))
            txtFMemo.Text = IIf(IsNull(adors(3)), "", adors(3))
        End If
    Else
        If adors.State > 0 Then adors.Close
        adors.Open "select max(ado_cie), max(memo_cie) from grupousuarios where id=f_usuariosis_idgpo(" & giUsuario & ")", gConSql, adOpenStatic, adLockReadOnly
        If miMasivo > 0 And miAccion = 0 Then
            txtFAcuerdo.Text = ""
            txtFMemo.Text = ""
            txtAcuCierre.Tag = "n" 'Acepta puros números
            lblAdoCierre.Caption = "Consecutivo inicial Acuerdos Cierre: "
            txtMemorando.Text = IIf(IsNull(adors(1)), "", adors(1))
        Else
            If Len(adors(0)) > 0 Then
                txtAcuCierre.Text = adors(0)
                txtMemorando.Text = adors(1)
            End If
        End If
    End If
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
