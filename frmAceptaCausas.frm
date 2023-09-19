VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form AceptaCausas 
   Caption         =   "Detalle del asunto por ley y causa"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtObs 
      Height          =   600
      Left            =   2565
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1395
      Width           =   7350
   End
   Begin VB.TextBox txtMemorando 
      Height          =   285
      Left            =   2565
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1080
      Width           =   7350
   End
   Begin VB.TextBox txtExp 
      Height          =   285
      Left            =   225
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   495
      Width           =   2310
   End
   Begin VB.TextBox txtIF 
      Height          =   600
      Left            =   2565
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   450
      Width           =   7350
   End
   Begin MSComctlLib.ListView lvCausasTurn 
      Height          =   1770
      Left            =   180
      TabIndex        =   2
      Top             =   2115
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   3122
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Ley"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Causa"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Responsable"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "F.Infracción"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Rechazar / No Procede"
      Height          =   420
      Index           =   1
      Left            =   5130
      TabIndex        =   5
      Top             =   4095
      Width           =   1995
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar / Procede"
      Height          =   420
      Index           =   0
      Left            =   3015
      TabIndex        =   3
      Top             =   4095
      Width           =   1590
   End
   Begin VB.Label Label4 
      Caption         =   "Observaciones:"
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   1440
      Width           =   1950
   End
   Begin VB.Label Label3 
      Caption         =   "Número y Fecha Memorando:"
      Height          =   285
      Left            =   270
      TabIndex        =   8
      Top             =   1080
      Width           =   2265
   End
   Begin VB.Label Label2 
      Caption         =   "Expediente:"
      Height          =   285
      Left            =   270
      TabIndex        =   6
      Top             =   180
      Width           =   1725
   End
   Begin VB.Label Label1 
      Caption         =   "Institución:"
      Height          =   195
      Left            =   2610
      TabIndex        =   4
      Top             =   180
      Width           =   1950
   End
End
Attribute VB_Name = "AceptaCausas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAceptar_Click(Index As Integer)
If Index = 0 Then
    gs = "ok"
End If
Unload Me
End Sub

Private Sub Form_Load()
Dim adors As New ADODB.Recordset, i As Integer
If Len(gs) > 0 And gi1 > 0 And gi > 0 Then 'Obtiene las causas del asunto
    txtMemorando.Text = gs2
    txtObs.Text = gs3
    adors.Open "{call P_ModulosDatosXLeyCausa(" & gi1 & "," & gi & "," & gs4 & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
    lvCausasTurn.ListItems.Clear
    
    Do While Not adors.EOF
        If i = 0 Then
            i = 1
            txtExp.Text = adors(1)
            txtIF.Text = adors(2)
        End If
        lvCausasTurn.ListItems.Add i, , adors(3)
        lvCausasTurn.ListItems(i).SubItems(1) = IIf(IsNull(adors(4)), "", adors(4))
        lvCausasTurn.ListItems(i).SubItems(2) = IIf(IsNull(adors(5)), "", adors(5))
        lvCausasTurn.ListItems(i).SubItems(3) = IIf(IsNull(adors(6)), "", adors(6))
        lvCausasTurn.ListItems(i).Tag = adors(0) 'Guarda el id
        adors.MoveNext
        i = i + 1
    Loop
End If
End Sub


