VERSION 5.00
Begin VB.Form Frm_Seguridad 
   Caption         =   "Seguridad"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4230
   Icon            =   "frm_Acceso.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bot_Seg 
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton bot_Seg 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame Frm_Seg 
      Height          =   1335
      Left            =   1200
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.TextBox Txt_SegCon 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Txt_SegCta 
         Height          =   285
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Lbl_SegCon 
         Caption         =   "Contraseña:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Lbl_SegCta 
         Caption         =   "Cuenta:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "Frm_Seguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bot_Seg_Click(Index As Integer)
Dim s_condiseg As String, S_NombreUsu As String
Select Case Index
    Case 0  'Aceptar
        If Len(Trim(Frm_Seguridad.Txt_SegCta)) = 0 And Len(Trim(Frm_Seguridad.Txt_SegCon)) = 0 Then
            Exit Sub
        Else
            If FU_ValidaInforSeg("Seguridad") Then
                If Not FU_ValidaDatosSeguridad Then
                    Exit Sub
                Else
                    '--------Paso correctamente--------
                    s_condiseg = "n_cveusuario = " & Trim(Frm_Seguridad.Txt_SegCta)
                    gs_usuario = Val(Trim(Frm_Seguridad.Txt_SegCta))
                    'Set Frm_Seguridad = Nothing
                    Unload Frm_Seguridad
                    S_NombreUsu = FU_ExtraeNombreUsuario(gs_usuario)
                    Load MDI_Prin
                    MDI_Prin.Txt_UsuarioAct = S_NombreUsu
                    MDI_Prin.Show
                    
                    Call PR_GrabaBitacora(0, 1, "Acceso", Str(gs_usuario))
                End If
            End If
        End If
        
    Case 1  'Cancelar
        End
End Select
End Sub
Private Sub Form_Load()
''Frm_Seguridad.Txt_SegCta = "10001"
''Frm_Seguridad.Txt_SegCon = "PATO1234"

S_Path = Trim(App.Path)
If FU_DatosServerExt() Then
    'MsgBox "Si hay conexión"
Else
    MsgBox "NO hay conexión con la base de datos", 0 + 16, "Verificar archivo de Conf."
    Frm_Seguridad.bot_Seg(0).Enabled = False
End If
End Sub

Private Sub Txt_SegCon_GotFocus()
If Len(Trim(Frm_Seguridad.Txt_SegCon)) > 0 Then
   Frm_Seguridad.Txt_SegCon.SelStart = 0
   Frm_Seguridad.Txt_SegCon.SelLength = Len(Trim(Frm_Seguridad.Txt_SegCon))
End If
End Sub

Private Sub Txt_SegCon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Frm_Seguridad.bot_Seg(0).SetFocus
End Sub

Private Sub Txt_SegCta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Frm_Seguridad.Txt_SegCon.SetFocus
KeyAscii = FU_vte_ValTecla(KeyAscii, NUMEROS, " ")
End Sub
