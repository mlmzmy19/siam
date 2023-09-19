VERSION 5.00
Begin VB.Form Informes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes"
   ClientHeight    =   4872
   ClientLeft      =   9132
   ClientTop       =   1836
   ClientWidth     =   4728
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4872
   ScaleWidth      =   4728
   Begin VB.Frame Frame1 
      Height          =   4728
      Left            =   0
      TabIndex        =   0
      Top             =   72
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "Consulta Tabular en Excel"
         Height          =   624
         Index           =   4
         Left            =   684
         Picture         =   "Informes.frx":0000
         TabIndex        =   5
         Top             =   3888
         Width           =   3500
      End
      Begin VB.CommandButton Command1 
         Enabled         =   0   'False
         Height          =   510
         Index           =   3
         Left            =   680
         Picture         =   "Informes.frx":563E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2160
         Width           =   3500
      End
      Begin VB.CommandButton Command1 
         Height          =   510
         Index           =   2
         Left            =   680
         Picture         =   "Informes.frx":622E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3060
         Width           =   3500
      End
      Begin VB.CommandButton Command1 
         Height          =   510
         Index           =   1
         Left            =   680
         Picture         =   "Informes.frx":AA6A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1350
         Width           =   3500
      End
      Begin VB.CommandButton Command1 
         Height          =   624
         Index           =   0
         Left            =   680
         Picture         =   "Informes.frx":102B4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   372
         Width           =   3500
      End
   End
End
Attribute VB_Name = "Informes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
Dim s As String, adors As New ADODB.Recordset
If Index = 0 Then 'reportes de supervisión
    gi = 1 'tipo de reporte
    Informe.Show vbModal
ElseIf Index = 1 Then 'reportes de estadísticos
    gi = 2
    Informe.Show vbModal
ElseIf Index = 2 Then 'reportes de referencias Cruzadas
    frmRefCruz.Show vbModal
ElseIf Index = 3 Then 'reportes de MSS
    InformeModulos.Show vbModal
ElseIf Index = 4 Then 'reportes de MSS
    ConsultaTabular_Excel.Show vbModal
End If
End Sub

Private Sub Form_Load()
Dim adors As New ADODB.Recordset
'    If adors.State Then adors.Close
'    adors.Open "select count(*) from usuesp where idusi=" & giUsuario & " and tipo=2 and baja=0", gConSql, adOpenStatic, adLockReadOnly
'    If adors(0) > 0 Then
'        Command1(5).Enabled = True
'    End If
End Sub
