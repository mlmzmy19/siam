VERSION 5.00
Begin VB.Form Principal 
   BackColor       =   &H0000C000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Menú Principal"
   ClientHeight    =   9144
   ClientLeft      =   2268
   ClientTop       =   2028
   ClientWidth     =   15180
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "Principal.frx":0000
   ScaleHeight     =   9144
   ScaleWidth      =   15180
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "Sanciones impuestas (Buró)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   12780
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5130
      Width           =   2295
   End
   Begin VB.CommandButton cmdGuia 
      Caption         =   "Guia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   13320
      Picture         =   "Principal.frx":1DB48
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Infracc. Turnadas SIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12780
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4590
      Width           =   2295
   End
   Begin VB.CommandButton cmdBotón 
      BackColor       =   &H00FFFFFF&
      Height          =   1950
      Index           =   4
      Left            =   12960
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Principal.frx":1F74A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5715
      UseMaskColor    =   -1  'True
      Width           =   1950
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   375
      Left            =   13080
      Picture         =   "Principal.frx":22941
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdBotón 
      BackColor       =   &H00FFFFFF&
      Height          =   1950
      Index           =   3
      Left            =   9966
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Principal.frx":232C7
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   1950
   End
   Begin VB.CommandButton cmdBotón 
      BackColor       =   &H00FFFFFF&
      Height          =   1995
      Index           =   2
      Left            =   6974
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Principal.frx":26814
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   1950
   End
   Begin VB.CommandButton cmdBotón 
      BackColor       =   &H00FFFFFF&
      Height          =   1995
      Index           =   1
      Left            =   3996
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Principal.frx":28109
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   1950
   End
   Begin VB.CommandButton cmdBotón 
      BackColor       =   &H00FFFFFF&
      Height          =   1995
      Index           =   0
      Left            =   1044
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Principal.frx":29628
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   1950
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Consultas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   10215
      TabIndex        =   7
      Top             =   6750
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Seguimiento"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   6660
      TabIndex        =   6
      Top             =   6750
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Análisis"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   5
      Top             =   6795
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Registro"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   1035
      TabIndex        =   4
      Top             =   6795
      Visible         =   0   'False
      Width           =   1725
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_Click()
    Dim adors As New ADODB.Recordset
    If adors.State Then adors.Close
    adors.Open "select f_url_conexion(3,0,0) from dual", gConSql, adOpenStatic, adLockReadOnly
    If Len(adors(0)) > 0 And i <> 200 Then
        gsWWW = adors(0)
    End If
    With Browser
        .yÚnicavez = 0
        .Caption = "Infracciones turnadas a Sanciones por el SIO"
        .Height = 3000
        .Width = 10000
        .cmd1.Visible = True
        .cmd2.Visible = True
        .Show vbModal
    End With
End Sub

Private Sub cmdBotón_Click(Index As Integer)
If Index = 0 Then
    Registro.Show 'vbModal
ElseIf Index = 1 Then
    Análisis.Show 'vbModal
ElseIf Index = 2 Then
    Seguimiento.Show 'vbModal
ElseIf Index = 3 Then
    Informes.Show 'vbModal
ElseIf Index = 4 Then
    Utilerias.Show 'vbModal
End If
End Sub

Private Sub cmdGuia_Click()
    Dim adors As New ADODB.Recordset
    If adors.State Then adors.Close
    adors.Open "select f_url_conexion(6,0,0) from dual", gConSql, adOpenStatic, adLockReadOnly
    If Len(adors(0)) > 0 And i <> 200 Then
        gsWWW = adors(0)
    End If
    With Browser
        .yÚnicavez = 0
        .Caption = "GUIA SIAM"
        .Show vbModal
    End With
End Sub

Private Sub cmdSalir_Click()
End
Unload Me
End Sub

Private Sub Command1_Click()
    Dim adors As New ADODB.Recordset
    If adors.State Then adors.Close
    adors.Open "select f_url_conexion(9,0,0) from dual", gConSql, adOpenStatic, adLockReadOnly
    If Len(adors(0)) > 0 And i <> 200 Then
        gsWWW = adors(0)
    End If
    With Browser
        .yÚnicavez = 0
        .Caption = "Sanciones Impuestas (Buró)"
        .Height = 3000
        .Width = 10000
        .cmd1.Visible = True
        .cmd2.Visible = True
        .Show vbModal
    End With
End Sub

Private Sub Form_Load()
Dim adors As New ADODB.Recordset
'Deshabilita botones dacceso a los módulos que no tiene permiso
If adors.State Then adors.Close
adors.Open "select * from usuariossistema where id=" & giUsuario, gConSql, adOpenStatic, adLockReadOnly
If Not adors.EOF Then
    If adors!Registro = 0 Then
        cmdBotón(0).Enabled = False
    End If
    If adors!Análisis = 0 Then
        cmdBotón(1).Enabled = False
    End If
    If adors!Seguimiento = 0 Then
        cmdBotón(2).Enabled = False
    End If
    If adors!Reportes = 0 Then
        cmdBotón(3).Enabled = False
    End If
End If
End Sub
