VERSION 5.00
Begin VB.Form frmNuevaContrase�a 
   Caption         =   "Nueva Contrase�a "
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frame 
      Caption         =   "Nueva Contrase�a de Acceso al Sistema"
      Height          =   2505
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   4065
      Begin VB.CommandButton cmdBot�n 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1860
         Width           =   1365
      End
      Begin VB.CommandButton cmdBot�n 
         Caption         =   "&Aceptar"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1860
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2010
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1200
         Width           =   1635
      End
      Begin VB.TextBox Text1 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   2010
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   540
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Confirmar Contrase�a"
         Height          =   315
         Index           =   1
         Left            =   270
         TabIndex        =   4
         Top             =   1230
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Nueva Contrase�a"
         Height          =   315
         Index           =   0
         Left            =   270
         TabIndex        =   2
         Top             =   570
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmNuevaContrase�a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBot�n_Click(Index As Integer)
Dim s As String, i As Integer
If Index = 1 Then
    MsgBox "La contrase�a Seguir� siendo la misma", vbOKOnly, ""
    Unload Me
    Exit Sub
End If
If Len(Text1(0).Text) = 0 Then
    MsgBox "La contrase�a no puede estar vacia", vbOKOnly, ""
    Exit Sub
ElseIf UCase(Text1(0).Text) <> UCase(Text1(1).Text) Then
    MsgBox "La contrase�a no ha sido confirmada correctamente", vbOKOnly + vbExclamation, ""
    Exit Sub
End If
    gConSql.Execute "update usuariossistema set contrase�a='" + UCase(Text1(0).Text) + "' where id=" & giUsuario, i
If i > 0 Then
    MsgBox "La contrase�a se cambi� exit�samente", vbOKOnly + vbInformation, ""
    Unload Me
Else
    MsgBox "El cambio no se realiz�. Vuelva a intentar", vbOKOnly + vbCritical, ""
End If
End Sub

Private Sub Form_Activate()
'ActualizaColorFormulario Me

End Sub

Private Sub Text1_Change(Index As Integer)
If Index = 1 Then
    cmdBot�n(0).Enabled = UCase(Text1(0).Text) = UCase(Text1(1).Text)
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 0 And KeyAscii = 13 Then
    Text1(1).SetFocus
ElseIf Index = 1 And KeyAscii = 13 Then
    cmdBot�n_Click (0)
End If
End Sub
