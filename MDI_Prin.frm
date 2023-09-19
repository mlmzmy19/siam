VERSION 5.00
Begin VB.MDIForm MDI_Prin 
   AutoShowChildren=   0   'False
   BackColor       =   &H00C0C0C0&
   Caption         =   "SIAM"
   ClientHeight    =   6540
   ClientLeft      =   135
   ClientTop       =   -1890
   ClientWidth     =   16080
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDI_Prin.frx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Tb_General 
      Align           =   1  'Align Top
      Height          =   420
      HelpContextID   =   1
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   16020
      TabIndex        =   0
      Top             =   0
      Width           =   16080
      Begin VB.TextBox Txt_UsuarioAct 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6930
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "USUARIO"
         Top             =   30
         Width           =   8235
      End
   End
   Begin VB.Menu mnuModulos 
      Caption         =   "&Registro"
      Index           =   0
   End
   Begin VB.Menu mnuModulos 
      Caption         =   "Registro &Masivo"
      Index           =   1
   End
   Begin VB.Menu mnuModulos 
      Caption         =   "&Análisis"
      Index           =   2
   End
   Begin VB.Menu mnuModulos 
      Caption         =   "&Seguimiento"
      Index           =   3
   End
   Begin VB.Menu mnuModulos 
      Caption         =   "Re&portes"
      Index           =   4
   End
   Begin VB.Menu mnuModulos 
      Caption         =   "&Consulta"
      Index           =   5
   End
   Begin VB.Menu mnuModulos 
      Caption         =   "&Utilerías"
      Index           =   6
   End
   Begin VB.Menu mnuVentana 
      Caption         =   "&Ventana"
      Index           =   1
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "A&yuda"
      Begin VB.Menu mnuGuia 
         Caption         =   "&Guia"
      End
      Begin VB.Menu mnuInfracciones 
         Caption         =   "&Infracciones Turnadas SIO"
      End
      Begin VB.Menu mnuBURO 
         Caption         =   "Sanciones Impuestas &BURO"
      End
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "Sa&lir"
   End
End
Attribute VB_Name = "MDI_Prin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Activate()
  'Actualiza valor de variable global ocupada en reportes
Dim adors As New ADODB.Recordset
  gi2 = 0
'Principal.Left = (MDI.Width - Principal.Width) / 2
'Principal.Top = (MDI.Height - Principal.Height) / 2
'Principal.Show
adors.Open "select descripción from usuariossistema where id=" & giUsuario, gConSql, adOpenStatic, adLockReadOnly
If Not adors.EOF Then
    Txt_UsuarioAct.Text = "USUARIO: " & adors(0) & ". Versión: " & gsVersión
    If gidesa > 0 Then
        Txt_UsuarioAct.Text = Txt_UsuarioAct.Text & " (DESARROLLO). Versión: " & gsVersión
    End If
End If
'Principal.WindowState = 2
End Sub

Private Sub MDIForm_Load()
Dim S_condibus As String, N_Cveperfil As Integer
Dim adors As New ADODB.Recordset

End Sub

Private Sub Mnu_ArchivoSal_Click()
PR_Desconecta
End

End Sub

Private Sub Mnu_Ayudaacerca_Click()
Load Frm_Acerca
Frm_Acerca.Show vbModal
End Sub

Private Sub Mnu_CapturaPro_Click()

If gn_Miperfil = 1 Or gn_Miperfil = 3 Or gn_Miperfil = 4 Then
    gs_ProcCuestionario = 2  'Cuestionario de Producción
    
    Load Frm_Cuestionario
    Frm_Cuestionario.Caption = "Definición del encabezado del Cuestionario (PRODUCCIÓN)"
    Frm_Cuestionario.Show vbModal
Else
    MsgBox "Su perfil no esta autorizado para entrar a este módulo.", 0 + 48, "Consultar con el Administrador del Sistema"
    Exit Sub
End If

'gs_ProcCuestionario = 2 'Cuestionario de Producción
'If gn_Miperfil = 1 Or gn_Miperfil = 3 Or gn_Miperfil = 4 Then
'    Load Frm_PEC
'    Frm_PEC.Lbl_PECTitulo.Caption = "Cuestionario de Producción"
'    Frm_PEC.Caption = "Cuestionario de Producción"
'    Frm_PEC.Bot_PEC(1).Enabled = True
'    Frm_PEC.Show vbModal
'Else
'    MsgBox "Su perfil no esta autorizado para entrar a este módulo.", 0 + 48, "Consultar con el Administrador del Sistema"
'    Exit Sub
'End If
End Sub

Private Sub mnuSanTurnadas_Click()

End Sub

Private Sub mnuAnalisis_Click()
Análisis.Show
End Sub

Private Sub mnuBURO_Click()
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

Private Sub mnuGuia_Click()
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

Private Sub mnuInfracciones_Click()
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

Private Sub mnuRegistro_Click()
End Sub

Private Sub mnuReportes_Click()
Informes.Show
End Sub

Private Sub mnuModulos_Click(Index As Integer)
Dim frm As Form
On Error Resume Next
If Index = 0 Then
    Set frm = Registro
ElseIf Index = 1 Then
    Set frm = RegMasivo
    frm.Width = 16380
    frm.Height = 11175
ElseIf Index = 2 Then
    Set frm = Análisis
ElseIf Index = 3 Then
    Set frm = Seguimiento
ElseIf Index = 4 Then
    Set frm = Informes
ElseIf Index = 5 Then
    Set frm = ConsultaMasiva
    frm.Width = 18600
    frm.Height = 11130
ElseIf Index = 6 Then
    Set frm = Utilerias
End If
With frm
    .Left = (Me.Width - .Width - 50) / 2
    .Top = (Me.Height - .Height - 1000) / 2
    'If Index = 0 Or Index = 2 Or Index = 3 Then
    '    .Show vbModal
    'Else
        .Show
    'End If
End With
End Sub

Private Sub mnuSalir_Click()
Dim objForm As Form
For Each objForm In Forms
    If InStr("|Registro|Análisis|Seguimiento|", "|" & objForm.Name & "|") > 0 Then
        If MsgBox("Se encuentra abierto el Formulario " & objForm.Name & ". Está seguro de salir?", vbYesNo + vbQuestion + vbDefaultButton2, "Validación") = vbNo Then
            Exit Sub
        End If
    End If
Next
End
Unload Me
End Sub

Private Sub mnuSeguimiento_Click()
Seguimiento.Show
End Sub

Private Sub mnuUtilerias_Click()
Utilerias.Show
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

