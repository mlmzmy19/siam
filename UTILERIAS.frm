VERSION 5.00
Begin VB.Form Utilerias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utilerías"
   ClientHeight    =   6705
   ClientLeft      =   9735
   ClientTop       =   3630
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   4785
   Begin VB.Frame Frame1 
      Height          =   6450
      Left            =   36
      TabIndex        =   0
      Top             =   36
      Width           =   4695
      Begin VB.CommandButton cmdMasivo 
         Caption         =   "MODULO DE EMPLAZAMIENTO MASIVO (ASUNTOS TRIMESTRALES)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   405
         TabIndex        =   8
         Top             =   3735
         Width           =   3840
      End
      Begin VB.CommandButton cmdReturnar 
         Caption         =   "Re-turnar a Análisis Nuevo Responsable"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   396
         TabIndex        =   7
         Top             =   2844
         Width           =   3840
      End
      Begin VB.CommandButton Command1 
         Height          =   465
         Index           =   2
         Left            =   405
         Picture         =   "UTILERIAS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   3810
      End
      Begin VB.CommandButton Command1 
         Height          =   510
         Index           =   3
         Left            =   405
         Picture         =   "UTILERIAS.frx":1924
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1158
         Width           =   3855
      End
      Begin VB.CommandButton Command1 
         Height          =   510
         Index           =   4
         Left            =   405
         Picture         =   "UTILERIAS.frx":31ED
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2001
         Width           =   3855
      End
      Begin VB.CommandButton Command1 
         Enabled         =   0   'False
         Height          =   510
         Index           =   5
         Left            =   372
         Picture         =   "UTILERIAS.frx":7D73
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2844
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.CommandButton Command1 
         Height          =   555
         Index           =   1
         Left            =   1224
         Picture         =   "UTILERIAS.frx":90BE
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5625
         Width           =   2310
      End
      Begin VB.CommandButton Command1 
         Height          =   600
         Index           =   0
         Left            =   1116
         Picture         =   "UTILERIAS.frx":E5D4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4830
         Width           =   2535
      End
      Begin VB.Line Line1 
         X1              =   45
         X2              =   4635
         Y1              =   4620
         Y2              =   4620
      End
   End
End
Attribute VB_Name = "Utilerias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdMasivo_Click()
Dim sAppName As String, sAppPath As String, adors As New ADODB.Recordset
Dim iUsi As Integer

sAppName = "Google Chorme"
iUsi = giUsuario

If adors.State Then adors.Close
adors.Open "select f_url_conexion(16," & iUsi & ",0) from dual", gConSql, adOpenStatic, adLockReadOnly

If Not adors.EOF Then
    sAppPath = adors(0)
End If
If Not Len(Dir(sAppPath)) > 0 Then
    If adors.State Then adors.Close
    iUsi = 0
    adors.Open "select f_url_conexion(16," & iUsi & ",0) from dual", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        sAppPath = adors(0)
    End If
End If
If Not Len(Dir(sAppPath)) > 0 Then
    Call MsgBox("Favor de contartar a Administrador de SIAM ya queno se ubica las aplicaciones de CHROME o INTERNET EXPLORER", vbOKOnly + vbInformation, "Validación")
    Exit Sub
End If


If adors.State Then adors.Close
adors.Open "select f_url_conexion(15," & giUsuario & ",0) from dual", gConSql, adOpenStatic, adLockReadOnly
If Len(adors(0)) > 0 And i <> 200 Then
    If iUsi = 1 Then
        sAppPath = sAppPath & " -url " & adors(0)
    Else
        sAppPath = sAppPath & " " & adors(0)
    End If
End If


Shell sAppPath, vbMinimizedFocus


'    Dim adors As New ADODB.Recordset
'    If adors.State Then adors.Close
'    adors.Open "select f_url_conexion(15," & giUsuario & ",0) from dual", gConSql, adOpenStatic, adLockReadOnly
'    If Len(adors(0)) > 0 And i <> 200 Then
'        gsWWW = adors(0)
'    End If
'    With Browser
'        .yÚnicavez = 0
'        .Caption = "MODULO DE REGISTRO MASIVO ASUNTOS TRIMESTRALES"
'        '.Height = 3000
'        '.Width = 10000
'        .Show
'    End With
End Sub

Private Sub cmdReturnar_Click()
'Dim frm As New TurnarExp
'TurnarExp.Show vbModal
End Sub

Private Sub Command1_Click(Index As Integer)
Dim s As String, adors As New ADODB.Recordset
Dim i As Integer
Me.MousePointer = 11
If Index = 0 Then 'Cambio de contraseña
    gi1 = giUsuario
    gi = 28
    If adors.State Then adors.Close
    adors.Open "select descripción,contraseña from usuariossistema where id=" & gi1, gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        s = adors(0)
        gs = adors(1)
    End If
    If Len(s) > 0 Then
        giUsuario = -1
        With frmAcceso
            .Caption = "Acceso a Cambio de Contraseña. Epecifique su CONTRASEÑA ACTUAL"
            '.Label5(2).Caption = "Favor de especificar la contraseña del administrador"
            '.Label5(0).Caption = ""
            .txtUsuario.Enabled = False
            .txtUsuario.Text = s
            .ComboUsuarios.Visible = False
            .txtUsuario.Visible = True
            .cmdBotón(2).Visible = False
            .Show vbModal
        End With
    End If
    If giUsuario = gi1 Then
        frmNuevaContraseña.Show vbModal
    Else
        giUsuario = gi1
    End If
ElseIf Index = 1 Then 'Acerca del sistema
    Frm_Acerca.Show vbModal
ElseIf Index = 2 Then 'Instituciones Activas en SINE
    gs = "ArbolVarios-->select ci.id as idcla, i.id as idins, ci.descripción as clase, i.descripción|| case when i.id<>f_idinsh(i.id) then ' ('||substr(paq_conceptos.institucion(f_idinsh(i.id)),1,25)||case when length(paq_conceptos.institucion(f_idinsh(i.id)))>25 then '...' else '' end||')' else '' end  as Institucion " & _
         "from instituciones i left join claseinstitución ci on i.idsec=ci.idsec_sipres " & _
         "where f_sine_ifactiva(i.id, 2)>0 order by 3,4"
'         "select rci.idcla,ia.idins,paq_sio_conceptos.clase@sioprod(rci.idcla) as Clase, i.descripción || ' (' || to_char(min(trunc(registro)),'dd/mon/yyyy') || ')' as Institución " & _
'         "from sine.n_ifactivas@sioprod ia, instituciones i, relaciónclaseinstitución rci " & _
'         "Where i.ID = rci.idins And ia.idins = i.ID And ia.Status = 1 " & _
'         "group by rci.idcla,ia.idins,paq_sio_conceptos.clase@sioprod(rci.idcla), i.descripción " & _
'         "order by 3,4"
    gs3 = "expandir todos niveles"
    With SelProceso
        .TreeView1.CheckBoxes = False
        .Caption = "Instituciones Activas en el sistema de Notificaciones Electrónicas"
        .cmdAcción(1).Visible = False
        .cmdAcción(0).Caption = "Cerrar"
        .Frame1.Left = (.Width + 1900 - .Frame1.Width) / 2
        .cmdAcción(0).Left = (.Frame1.Width - .cmdAcción(0).Width) / 2
        .Width = .Width + 2000
        .TreeView1.Width = .TreeView1.Width + 2000
        .Show vbModal
    End With
ElseIf Index = 3 Then 'Habilita Instituciones no Activas para registro (Solo para usuarios especiales)
    gs = "ArbolVarios-->select rci.idcla, i.id as idins, paq_conceptos.clase(rci.idcla) as clase, i.descripción  as Inst " & _
         "from instituciones i, relaciónclaseinstitución rci " & _
         "Where i.ID = rci.idins And i.baja = 1 And i.ID > 0 " & _
         "order by 3,4"
    gs3 = "expandir todos niveles"
    With SelProceso
        .TreeView1.CheckBoxes = True
        .Caption = "Seleccione las Instituciones las cuales desea activar para Registro"
        '.cmdAcción(1).Visible = False
        '.cmdAcción(0).Caption = "Cerrar"
        .Frame1.Left = .Frame1.Left + 1000
        '.cmdAcción(0).Left = (.Frame1.Width - .cmdAcción(0).Width) / 2
        .Width = .Width + 2000
        .TreeView1.Width = .TreeView1.Width + 2000
        .Show vbModal
    End With
    If Len(gs) > 2 Then
        If adors.State Then adors.Close
        adors.Open "{call P_IF_ActivaIFs('" & gs & "')}", gConSql, adOpenForwardOnly, adLockReadOnly
        If adors.EOF Then
            i = 200
            s = ""
        Else
            s = adors(1)
        End If
        If i = 200 Then 'No activo las Ifs.
            Call MsgBox("Fallo el proceso de activación de IFs. " & s, vbOKOnly + vbInformation, "")
        Else
            Call MsgBox("Se activaron las IFs. Después de la primer asunto que registre con la Institución, está sera desactivada nuevamente", vbOKOnly + vbInformation, "")
        End If
        
    End If
    
ElseIf Index = 4 Then 'Publicación Sanciones
    Publicaciones.Show vbModal
ElseIf Index = 5 Then 'Publicación Estrados
    Estrados.Show vbModal
End If
Me.MousePointer = 0
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
Dim adors As New ADODB.Recordset
    If adors.State Then adors.Close
    adors.Open "select count(*) from usuesp where idusi=" & giUsuario & " and tipo=2 and baja=0", gConSql, adOpenStatic, adLockReadOnly
    If adors(0) > 0 Then
        Command1(5).Enabled = True
        cmdReturnar.Enabled = True
    End If
End Sub
