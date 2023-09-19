VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form Publicaciones 
   Caption         =   "Publicaciones"
   ClientHeight    =   9048
   ClientLeft      =   0
   ClientTop       =   348
   ClientWidth     =   14628
   LinkTopic       =   "Form1"
   ScaleHeight     =   9048
   ScaleWidth      =   14628
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   90
      TabIndex        =   29
      Top             =   4545
      Width           =   14505
      Begin VB.ComboBox comboTipo 
         Height          =   315
         Left            =   8910
         TabIndex        =   39
         Top             =   270
         Width           =   5460
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   2
         Left            =   6165
         MaxLength       =   2
         TabIndex        =   36
         Tag             =   "n"
         Top             =   360
         Width           =   555
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   1
         Left            =   5220
         MaxLength       =   2
         TabIndex        =   34
         Tag             =   "n"
         Top             =   360
         Width           =   645
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   3870
         MaxLength       =   4
         TabIndex        =   31
         Tag             =   "n"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Día:"
         Height          =   195
         Left            =   6255
         TabIndex        =   37
         Top             =   135
         Width           =   510
      End
      Begin VB.Label Label5 
         Caption         =   "Mes:"
         Height          =   195
         Left            =   5355
         TabIndex        =   35
         Top             =   135
         Width           =   510
      End
      Begin VB.Label Label4 
         Caption         =   "Año:"
         Height          =   195
         Left            =   3915
         TabIndex        =   33
         Top             =   135
         Width           =   510
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Firmeza:"
         Height          =   195
         Left            =   7335
         TabIndex        =   32
         Top             =   315
         Width           =   1590
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha cuándo queda firme la sanción (Especificar Año / Año y Mes / Año, Mes y Dia):"
         Height          =   420
         Left            =   225
         TabIndex        =   30
         Top             =   180
         Width           =   3480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3750
      Left            =   45
      TabIndex        =   15
      Top             =   5265
      Width           =   14550
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         Height          =   3600
         Left            =   12285
         TabIndex        =   16
         Top             =   90
         Width           =   2220
         Begin VB.Frame Frame4 
            Height          =   1455
            Left            =   45
            TabIndex        =   40
            Top             =   2160
            Width           =   2175
            Begin VB.CommandButton cmdBusca 
               Caption         =   "Busca"
               Height          =   285
               Left            =   405
               TabIndex        =   44
               Top             =   1080
               Width           =   1500
            End
            Begin VB.CheckBox chkSel 
               Caption         =   "Selecciona"
               Height          =   195
               Left            =   540
               TabIndex        =   43
               Top             =   765
               Width           =   1185
            End
            Begin VB.TextBox txtOficioB 
               Height          =   285
               Left            =   180
               TabIndex        =   41
               Top             =   405
               Width           =   1860
            End
            Begin VB.Label Label7 
               Caption         =   "Oficio / Expediente"
               Height          =   240
               Left            =   315
               TabIndex        =   42
               Top             =   180
               Width           =   1635
            End
         End
         Begin VB.CommandButton cmdProceso 
            Caption         =   "&Agregar Oficio"
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   450
            TabIndex        =   20
            Top             =   192
            Width           =   1200
         End
         Begin VB.CommandButton cmdProceso 
            Caption         =   "&Preparar Publicación"
            Enabled         =   0   'False
            Height          =   1020
            Index           =   2
            Left            =   450
            Picture         =   "Publicaciones.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1125
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.CommandButton cmdProceso 
            Caption         =   "&Quitar Oficio"
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   450
            TabIndex        =   17
            Top             =   675
            Width           =   1200
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3315
         Left            =   90
         TabIndex        =   19
         Top             =   360
         Width           =   12195
         _ExtentX        =   21505
         _ExtentY        =   5842
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IF"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Oficio"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Expediente"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fecha"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Ley / Causa"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Monto"
            Object.Width           =   2647
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Responsable"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Fecha Firmeza"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Tipo Firmeza"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Actividades realizadas:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   21
         Top             =   135
         Width           =   1620
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   2490
      Left            =   45
      TabIndex        =   11
      Top             =   2025
      Width           =   14550
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "Nombre"
         ForeColor       =   &H00808080&
         Height          =   288
         Index           =   4
         Left            =   108
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Tag             =   "c"
         ToolTipText     =   "Datos capturados del Oficio en la etapa de Análisis"
         Top             =   2145
         Width           =   2800
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "Nombre"
         ForeColor       =   &H00808080&
         Height          =   288
         Index           =   3
         Left            =   108
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Tag             =   "c"
         ToolTipText     =   "Datos capturados del Oficio en la etapa de Análisis"
         Top             =   1680
         Width           =   2800
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "Nombre"
         ForeColor       =   &H00808080&
         Height          =   735
         Index           =   5
         Left            =   3780
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Tag             =   "c"
         ToolTipText     =   "Datos capturados del Oficio en la etapa de Análisis"
         Top             =   1665
         Width           =   10635
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "Nombre"
         ForeColor       =   &H00808080&
         Height          =   1095
         Index           =   2
         Left            =   9990
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Tag             =   "c"
         ToolTipText     =   "Datos del documento de Solicitud"
         Top             =   360
         Width           =   4410
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "Nombre"
         ForeColor       =   &H00808080&
         Height          =   1095
         Index           =   1
         Left            =   5370
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Tag             =   "c"
         ToolTipText     =   "Nombre de la Institución y del Usuario"
         Top             =   360
         Width           =   4590
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H8000000F&
         DataField       =   "Nombre"
         ForeColor       =   &H00808080&
         Height          =   1095
         Index           =   0
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Tag             =   "c"
         ToolTipText     =   "Datos del origen de la Solicitud"
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Monto total de la sanción:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   27
         Top             =   1950
         Width           =   1830
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha de la sanción:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   25
         Top             =   1485
         Width           =   1485
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ley / Causa:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   3810
         TabIndex        =   23
         Top             =   1485
         Width           =   915
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Documento de la solicitud:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   10005
         TabIndex        =   14
         Top             =   135
         Width           =   1875
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Institución / Nombre(s):"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   5370
         TabIndex        =   13
         Top             =   135
         Width           =   1650
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Origen de la solicitud:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   12
         Top             =   135
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   2115
      TabIndex        =   8
      Top             =   -45
      Width           =   12480
      Begin VB.TextBox txtOficio 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   5310
         MaxLength       =   80
         TabIndex        =   1
         Tag             =   "c"
         ToolTipText     =   "No. de Oficio a realizar seguimiento"
         Top             =   900
         Width           =   5130
      End
      Begin VB.ComboBox comboOficios 
         Height          =   315
         Left            =   450
         TabIndex        =   2
         Top             =   1530
         Width           =   10050
      End
      Begin VB.TextBox txtExpediente 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   495
         MaxLength       =   80
         TabIndex        =   0
         Tag             =   "c"
         ToolTipText     =   "No. de Oficio a realizar seguimiento"
         Top             =   855
         Width           =   4500
      End
      Begin MSComctlLib.ImageList Imagenes 
         Left            =   9090
         Top             =   225
         _ExtentX        =   995
         _ExtentY        =   995
         BackColor       =   -2147483643
         ImageWidth      =   103
         ImageHeight     =   104
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Publicaciones.frx":05E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Publicaciones.frx":84F8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdActualpen 
         Caption         =   "Nueva &consulta"
         Height          =   420
         Left            =   10710
         TabIndex        =   3
         Top             =   720
         Width           =   1590
      End
      Begin VB.CommandButton cmdContinuar 
         BackColor       =   &H00008000&
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10845
         Picture         =   "Publicaciones.frx":10DFA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1260
         Width           =   1500
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Oficio de Sanción:"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   5355
         TabIndex        =   38
         Top             =   630
         Width           =   1860
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Número de expediente:"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   495
         TabIndex        =   28
         Top             =   585
         Width           =   1860
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Oficio de sanción (Causa Fecha de Infracción -- Status)"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   450
         TabIndex        =   10
         Top             =   1305
         Width           =   5820
      End
      Begin VB.Label Eti 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Módulo de Publicaciones"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   15.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   2
         Left            =   315
         TabIndex        =   9
         Top             =   180
         Width           =   8055
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Image Image1 
      Height          =   2040
      Left            =   135
      Picture         =   "Publicaciones.frx":11969
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2025
   End
End
Attribute VB_Name = "Publicaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mlMulta As Long 'Contiene el id de la multa seleccionada
Dim mlAsunto As Long 'Contiene el id del Asunto registrado en registro
Dim msLeyes As String 'Contine los id de las leyes
Dim msCausas As String 'Contine los id de las causas
Dim mlAsuxIF As Long 'Contiene id de la ifregxif selecionado
Dim mlAnálisis As Long 'Contiene id del análisis
Dim mlAnálisisImp As Long 'Contiene id del análisis
Dim msLeyesImp As String 'Contine los id de las leyes
Dim msCausasIMP As String 'Contine los id de las causas
Dim msMotivosImp As String 'Contine los id de los motivos de improcedencia
Dim msTodasCausasImp As String
Dim mlSeg As LogEventTypeConstants 'Contiene el idseg
Dim mbLimpiaExp As Boolean 'indicador para limpiar el campo o lista de exp Pendientes
Dim msActs As String




'Obtiene el valor de la Actividad correspondiente al lugar iLugar
Private Function F_Obten_Act(ByVal sActs As String, iLugarCombo As Integer) As Integer
Dim i As Integer
For i = 1 To iLugarCombo
    If InStr(sActs, ",") = 0 Then
        F_Obten_Act = 0
        Exit Function
    End If
    sActs = Mid(sActs, InStr(sActs, ",") + 1)
Next
F_Obten_Act = Val(sActs)
End Function

Private Sub cmbPendientes_Click()
If cmbPendientes.ListIndex >= 0 And mbLimpiaExp Then
    mbLimpiaExp = False
    txtExpediente.Text = ""
    txtOficio.Text = ""
End If
End Sub

Private Sub cmbPendientes_GotFocus()
mbLimpiaExp = True
End Sub

Private Sub cmbPendientes_LostFocus()
mbLimpiaExp = False
End Sub

Private Sub cmdActualpen_Click()
Dim i As Integer
ActualizaPendientes
LimpiaControlesInfConsulta 0
End Sub


Private Sub cmdBusca_Click()
Dim i As Integer
If ListView1.ListItems.Count = 0 Then
    cmdBusca.Enabled = False
    Exit Sub
End If
'Call MsgBox(ListView1.ListItems(1).SubItems(2) & " / " & ListView1.ListItems(1).SubItems(3), vbOKOnly, "")
For i = 1 To ListView1.ListItems.Count
    If LCase(ListView1.ListItems(i).SubItems(2)) = LCase(txtOficioB.Text) Or LCase(ListView1.ListItems(i).SubItems(3)) = LCase(txtOficioB.Text) Then
        If chkSel.Value Then
            ListView1.ListItems(i).Checked = True
        End If
        ListView1.ListItems(i).Selected = True
        Exit For
    End If
Next
If i >= ListView1.ListItems.Count Then
    Call MsgBox("No se localizó el expediente u Oficio: " & txtOficioB.Text, vbOKOnly, "")
Else
    If chkSel.Value Then
        Call MsgBox("Se seleccionó renglón: " & i, vbOKOnly, "")
    Else
        Call MsgBox("Se encontró renglón: " & i, vbOKOnly, "")
    End If
    verificaSel
End If

End Sub

'Busca y en en caso de encontrar obtiene datos de este folio
Private Sub cmdContinuar_Click()
Dim adors As New ADODB.Recordset
Dim l As Long, sOfi As String
'En caso de capturar el expediente sin especificar el oficio
If Len(Trim(txtExpediente.Text)) > 0 And Len(Trim(comboOficios.Text)) = 0 Then
    If adors.State Then adors.Close
    adors.Open "{call paq_publicacion.p_pub_exp_busoficios('" & txtExpediente.Text & "')}", gConSql, adOpenForwardOnly, adLockReadOnly
    If adors.EOF Then 'No hay registros con ese no de expediente
        Call MsgBox("No existen oficios de sanción asociados al Expediente o no existe el Expediente", vbOKOnly + vbInformation)
        Exit Sub
    End If
    comboOficios.Clear
    l = 0
    Do While Not adors.EOF
        comboOficios.AddItem adors(1), l
        comboOficios.ItemData(l) = adors(0)
        adors.MoveNext
    Loop
    If comboOficios.ListCount = 1 Then
        comboOficios.ListIndex = 0
    Else
        Call MsgBox("Seleccione el oficio", vbOKOnly + vbInformation, "")
        If comboOficios.Enabled Then comboOficios.SetFocus
        Exit Sub
    End If
End If
If Len(Trim(txtOficio.Text)) > 0 And Len(Trim(comboOficios.Text)) = 0 Then
    If adors.State Then adors.Close
    adors.Open "{call paq_publicacion.p_pub_san_busoficios('" & txtOficio.Text & "')}", gConSql, adOpenForwardOnly, adLockReadOnly
    If adors.EOF Then 'No hay registros con ese no de expediente
        Call MsgBox("No existen oficios de sanción asociados al Expediente o no existe el Expediente", vbOKOnly + vbInformation)
        Exit Sub
    End If
    comboOficios.Clear
    l = 0
    Do While Not adors.EOF
        comboOficios.AddItem adors(1), l
        comboOficios.ItemData(l) = adors(0)
        adors.MoveNext
    Loop
    If comboOficios.ListCount = 1 Then
        comboOficios.ListIndex = 0
    Else
        Call MsgBox("Seleccione el oficio", vbOKOnly + vbInformation, "")
        If comboOficios.Enabled Then comboOficios.SetFocus
        Exit Sub
    End If
End If
If comboOficios.ListIndex >= 0 Then 'En caso de especificar el oficio
    If adors.State Then adors.Close
    'adors.Open "select a.id,a.idregxif,ri.idreg,ss.idseg,sp.status from seguimientosanción ss, seguimiento s, análisis a, registroxif ri,seguimientosanpub sp  where ss.oficio='" & Replace(txtOficio.Text, "'", "''") & "' and ss.idseg=s.id and s.idana=a.id and a.idregxif=ri.id and s.id=sp.idseg(+)", gConSql, adOpenStatic, adLockReadOnly
    adors.Open "{call paq_publicacion.p_pub_buscaoficio(" & comboOficios.ItemData(comboOficios.ListIndex) & ")}", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        If IsNull(adors!Status) Then ' No hay problema no ha sido seleccionado
        ElseIf adors!Status = 1 Then
            MsgBox "El Oficio ya fue seleccionado para publicación", vbOKOnly + vbInformation, ""
            Exit Sub
        ElseIf adors!Status > 1 Then
            MsgBox "El Oficio ya fue publicado", vbOKOnly + vbInformation, ""
            Exit Sub
        End If
        
        txtFecha(0).Locked = False
        txtFecha(1).Locked = False
        txtFecha(2).Locked = False
        mlAnálisis = adors(0)
        mlAsuxIF = adors(1)
        mlAsunto = adors(2)
        mlSeg = adors(3)
        mlMulta = adors(5)
        If adors(6) > 0 Then
            comboTipo.ListIndex = BuscaCombo(comboTipo, adors(6), True)
        Else
            comboTipo.ListIndex = -1
            comboTipo.Text = ""
        End If
        comboOficios.Enabled = False
        txtExpediente.Enabled = False
        cmdContinuar.Enabled = False
        RefrescaDatos
    Else
        MsgBox "No se encontró asunto alguno con ese No. de Oficio", vbOKOnly + vbInformation, ""
    End If
Else
    MsgBox "Debe capturar el número de Oficio de Sanción", vbOKOnly + vbInformation, ""
End If
Exit Sub
End Sub

Private Sub RefrescaDatos()
Dim adors As ADODB.Recordset, i As Integer
Set adors = New ADODB.Recordset
adors.Open "{call p_analisis_datosregistro(" & mlAsunto & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
If Not adors.EOF Then
    For i = 0 To 2
        txtCampo(i).Text = adors(i)
    Next
    Set adors = New ADODB.Recordset
    adors.Open "select fecha,importepesos,paq_conceptos.leycausa(idanacau),aniof,mesf,diaf from seguimientosanxanacau where id=" & mlMulta, gConSql, adOpenForwardOnly, adLockReadOnly
    If Not adors.EOF Then
        txtCampo(3).Text = Format(adors(0), "dd/mmm/yyyy")
        txtCampo(4).Text = Format(adors(1), "###,###,###.00")
        txtCampo(5).Text = adors(2)
        txtCampo(3).Tag = mlMulta
        If IsNull(adors(3)) Then
            txtFecha(0).Text = ""
        Else
            txtFecha(0).Text = adors(3)
        End If
        If IsNull(adors(4)) Then
            txtFecha(1).Text = ""
        Else
            txtFecha(1).Text = adors(4)
        End If
        If IsNull(adors(5)) Then
            txtFecha(2).Text = ""
        Else
            txtFecha(2).Text = adors(5)
        End If
    End If
    cmdProceso(0).Enabled = True
Else
    For i = 0 To 5
        txtCampo(i).Text = ""
    Next
    For i = 0 To cmbCampo.UBound
        cmbCampo(i).Clear
    Next
End If
End Sub

'Acciones Nuevo, editar , agrega, borra...
Private Sub cmdProceso_Click(Index As Integer)
Dim adors  As New ADODB.Recordset
Dim n As Integer
On Error GoTo ErrorGuardaDatos:
If Index = 0 Then 'Agrega Sanción
    If Val(txtFecha(0).Text) = 0 Then
        Call MsgBox("Debe especificar previamente (fecha firmeza): Ano / Año y mes / Año, mes y día", vbOKOnly + vbInformation, "Validación")
        Exit Sub
    End If
    If i > comboTipo.ListIndex < 0 Then
        Call MsgBox("Debe especificar el tipo de firmeza", vbOKOnly + vbInformation, "Validación")
        Exit Sub
    End If
    If adors.State Then adors.Close
    adors.Open "select count(*) from seguimientosanpub where idmul=" & mlMulta, gConSql, adOpenStatic, adLockReadOnly
    If adors(0) > 0 Then
        MsgBox "El oficio de sanción ya fue elegido o publicado", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
    If adors.State Then adors.Close
    adors.Open "select idusi,paq_conceptos.responsable(idusi) as Responsable from seguimiento where id=" & mlSeg, gConSql, adOpenStatic, adLockReadOnly
    If adors(0) <> giUsuario Then
        If MsgBox("El Oficio de sanción corresponde a: " & adors(1) & ". Está seguro de agragar el Oficio Especificado", vbYesNo + vbQuestion, "Confirmación") = vbNo Then
            Exit Sub
        End If
    Else
        If MsgBox("Está seguro de agragar el Oficio Especificado", vbYesNo + vbQuestion, "Confirmación") = vbNo Then
            Exit Sub
        End If
    End If
    
    's = Mid(s, 1, Len(s) - 1)
    If adors.State Then adors.Close
    adors.Open "{call paq_publicacion.p_pub_agregamulta(" & mlMulta & "," & giUsuario & "," & Val(txtFecha(0).Text) & "," & Val(txtFecha(1).Text) & "," & Val(txtFecha(2).Text) & "," & comboTipo.ItemData(comboTipo.ListIndex) & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
    If adors(0) <= 0 Then
        Call MsgBox("No se realizó la operación favor de intentar nuevamente, si persiste el problema favor de reportarlo", vbInformation + vbOKOnly, "")
    End If
    'gConSql.Execute "insert into seguimientosanpub (idseg,fecha,idusi,status,numero,idins) values(" & mlSeguimiento & ",sysdate," & giUsuario & ",1,case when (select count(*) from seguimientosanpub where idusi=" & giUsuario & " and status=1)>0 then 1+(select max(nvl(numero,0)) from seguimientosanpub where idusi=" & giUsuario & " and status=1) else 1 end, f_analisis_idins(f_seguimiento_idana(" & mlSeguimiento & ")))", iRows
    'If iRows > 0 Then
    '    MsgBox "Se agrego el oficio", vbOKOnly + vbInformation, ""
    'End If
        
    ActualizaPendientes
End If
'ElseIf Index = 1 Then 'Consultar
'    With Actividades
'        .mlAnálisis = mlAnálisis
'        .mlSeguimiento = ListView1.SelectedItem.Tag
'        .yTipoOperación = 0
'        .ySoloConsulta = 1
'        gs = "no iniciar var"
'        .Show vbModal
'    End With
If Index = 1 Then 'Quitar
    
    If MsgBox("Está seguro de quitar el(los) Oficio(s) de sanción seleccionado(s)", vbYesNo + vbQuestion, "") = vbYes Then
        For i = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(i).Checked Then
                gConSql.Execute "delete from seguimientosanpub where idmul=" & ListView1.ListItems(i).Tag & " and status=1", iRows
                If iRows > 0 Then
                    n = n + 1
                End If
            End If
        Next
        If n > 0 Then
            Call MsgBox("Se quitaron " & n & " registros", vbOKOnly + vbInformation, "")
            ActualizaPendientes
        Else
            MsgBox "No fue posible realizar la operación", vbInformation + vbOKOnly, ""
            ActualizaPendientes
        End If
        'ListView1.ListItems.Remove ListView1.SelectedItem.Index
        'ListView1.Refresh
    End If
ElseIf Index = 2 Then 'Publicar
    n = 0
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            n = n + 1
        End If
    Next
    If n <= 0 Then
        cmdProceso(3).Enabled = False
        Exit Sub
    End If
    If MsgBox("Está seguro de Preparar la publicación de los (" & n & ") Oficios Seleccionados", vbYesNo + vbQuestion, "Confirmación") = vbNo Then
        Exit Sub
    End If
    
    If n > 0 Then
        s = ""
        For i = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(i).Checked Then
                s = s & ListView1.ListItems(i).Tag & ","
            End If
        Next
        If Len(s) > 0 Then
            s = Mid(s, 1, Len(s) - 1)
            If adors.State Then adors.Close
            adors.Open "{call paq_publicacion.p_pub_PublicarSel('" & s & "'," & giUsuario & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
            If adors.State Then adors.Close
            adors.Open "select count(*),f_pub_periodo(sysdate) from seguimientosanpub where idmul in (" & s & ") and idusipub=" & giUsuario, gConSql, adOpenForwardOnly, adLockReadOnly
            
            If adors(0) > 0 Then
                MsgBox "Se prepararon para publicación " & adors(0) & " Oficios, vigentes en el periodo: " & adors(1), vbOKOnly + vbInformation, ""
                ListView1.ListItems.Clear
                cmdActualpen_Click
            End If
'            s = Mid(s, 1, Len(s) - 1)
'            gConSql.Execute "update seguimientosanpub status=2,fechaini=to_date('" & Format(d1, "dd/mm/yyyy") & "','dd/mm/yyyy'),fechafin=to_date('" & Format(d1 + 30, "dd/mm/yyyy") & "','dd/mm/yyyy'), from seguimientosanpub where idseg in (" & s & ") and status=1", iRows
'            If iRows > 0 Then
'                MsgBox "Se prepararon para publicación " & iRows & " Oficios", vbOKOnly + vbInformation, ""
'                ListView1.ListItems.Clear
'                cmdActualpen_Click
'            End If
        End If
    End If
End If
Exit Sub
ErrorGuardaDatos:
If gConSql.Errors.Count > 0 Then
    yError = MsgBox("AVISO: " + gConSql.Errors(0).Description, vbAbortRetryIgnore + vbInformation, "Excepción (" + Str(gConSql.Errors(0).Number) + ")")
Else
    yError = MsgBox("AVISO: " + Err.Description, vbAbortRetryIgnore + vbInformation, "Excepción (" + Str(Err.Number) + ")")
End If


If yError = vbRetry Then
    Resume
ElseIf yError = vbIgnore Then
    Resume Next
End If


End Sub

Private Sub comboOficios_Change()
If comboOficios.ListIndex > 0 Then
    cmdContinuar_Click
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Call LlenaCombo(comboTipo, "tipofirmeza", "baja=0", False)
ActualizaPendientes
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If ListView1.ListItems.Count > 0 Then
    'If MsgBox("Está seguro de salir sin preparar la publicación de los oficios seleccionados", vbQuestion + vbYesNo, "Confirmación") = vbNo Then
    '    Cancel = True
    '    Exit Sub
    'End If
End If
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
verificaSel
End Sub

Sub verificaSel()
Dim i As Integer
For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(i).Checked Then
        Exit For
    End If
Next
cmdProceso(1).Enabled = (i <= ListView1.ListItems.Count)
cmdProceso(2).Enabled = cmdProceso(1).Enabled
End Sub

Private Sub txtAnio_Change()

End Sub

Private Sub txtAnio_LostFocus()

End Sub

Private Sub txtCampo_Change(Index As Integer)
If Index >= 5 And Index <= 6 Then
    'Image2.Picture = Imagenes.ListImages(2).Picture
End If
End Sub

Private Sub txtCampo_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 1 And InStr("-", Chr(KeyAscii)) > 0 Then Exit Sub
KeyAscii = TeclaOprimida(txtCampo(Index), KeyAscii, txtCampo(Index).Tag, False)
'MsgBox "asd"
End Sub

Private Sub txtCampo_LostFocus(Index As Integer)
Dim adors As New ADODB.Recordset
If Mid(txtCampo(Index).Tag, 1, 1) = "f" Then
    If IsDate(txtCampo(Index).Text) Then
        d = CDate(txtCampo(Index).Text)
        txtCampo(Index).Text = Format(d, gsFormatoFecha)
        If adors.State Then adors.Close
        adors.Open "select sysdate from dual", gConSql, adOpenStatic, adLockReadOnly
        If Int(adors(0)) - Int(d) < 0 Then
            Call MsgBox("Fecha no válida. No se permite ingresar fecha mayor a la fecha actual (" & Format(adors(0), gsFormatoFecha) & ")", vbOKOnly + vbInformation, "")
            txtCampo(Index) = ""
            Exit Sub
        End If
    Else
        If Len(txtCampo(Index).Text) > 0 Then
            Call MsgBox("Fecha no válida. Verificar", vbOKOnly + vbInformation, "")
            txtCampo(Index) = ""
        End If
    End If
End If
End Sub

Sub ActualizaPendientes()
Dim adors As New ADODB.Recordset
adors.Open "select paq_publicacion.f_pub_usu_priv(" & giUsuario & ") from dual", gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then    'Usuario con privilegios para realizar Publicación de los Oficios seleccionados
    cmdProceso(2).Visible = True
End If
If adors.State Then adors.Close
adors.Open "{call paq_publicacion.p_pub_preparados(" & giUsuario & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
ListView1.ListItems.Clear
Do While Not adors.EOF
    i = ListView1.ListItems.Count + 1
    ListView1.ListItems.Add i, , IIf(IsNull(adors(0)), 0, adors(0)) 'No
    ListView1.ListItems(i).SubItems(1) = adors(1) 'Institución
    ListView1.ListItems(i).SubItems(2) = adors(2) 'Oficio
    ListView1.ListItems(i).SubItems(3) = adors(3) 'Expediente
    ListView1.ListItems(i).SubItems(4) = adors(4) 'Fecha
    ListView1.ListItems(i).SubItems(5) = adors(5) 'Ley/Causa
    ListView1.ListItems(i).SubItems(6) = adors(6) 'Monto
    ListView1.ListItems(i).SubItems(7) = adors(7) 'Responsable
    ListView1.ListItems(i).SubItems(8) = adors(8) 'Fecha Firmeza
    ListView1.ListItems(i).SubItems(9) = adors(9) 'Tipo Firmeza
    ListView1.ListItems(i).Tag = adors(10) 'Guarda el idmul
    adors.MoveNext
Loop
End Sub

Private Function F_ObtieneAct(ByVal sActs As String, ByVal iPos As Integer) As Integer
Dim i As Integer
Do While i < iPos
    If InStr(sActs, ",") = 0 Then
        F_ObtieneAct = -99
        Return
    End If
    sActs = Mid(sActs, InStr(sActs, ",") + 1)
Loop
If i = iPos Then
    F_ObtieneAct = Val(sActs)
End If
End Function

Private Sub LimpiaControlesInfConsulta(iTipo As Byte)
If iTipo = 0 Then 'Limpia todo y prepara un nuevo seguimiento
    comboOficios.Enabled = True
    txtExpediente.Enabled = True
    comboOficios.Text = ""
    cmdContinuar.Enabled = True
    txtFecha(0).Text = ""
    txtFecha(1).Text = ""
    txtFecha(2).Text = ""
    txtFecha(0).Locked = True
    txtFecha(1).Locked = True
    txtFecha(2).Locked = True
    comboTipo.ListIndex = -1
    comboTipo.Text = ""
    'ListView1.ListItems.Clear
    For i = 0 To cmdProceso.UBound
        cmdProceso(i).Enabled = False
    Next
    For i = 0 To txtCampo.UBound
        txtCampo(i).Text = ""
    Next
ElseIf iTipo = 1 Then 'Limpia una nueva institución
    txtCampo(3).Text = ""
    ListView1.ListItems.Clear
End If
End Sub

Private Sub txtExpediente_Change()
comboOficios.Clear
comboOficios.Text = ""
txtOficio.Text = ""
End Sub

Private Sub txtExpediente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdContinuar_Click
    Exit Sub
End If
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
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
KeyAscii = TeclaOprimida(txtCampo(Index), KeyAscii, txtCampo(Index).Tag, False)

End Sub

'Validación de los campos Año/mes/día
Private Sub txtFecha_LostFocus(Index As Integer)
If Index = 0 And Len(txtFecha(0).Text) > 0 Then 'Año
    If Val(txtFecha(0).Text) < 2010 Or Val(txtFecha(0).Text) > Year(Now) Then  'el Año debe ser mayor 2009  hasta el presente
        Call MsgBox("Año incorrecto, favor de verificar el año", vbOKOnly + vbInformation, "Validación")
        txtCampo(0).SetFocus
        Exit Sub
    End If
ElseIf Index = 1 And Len(txtFecha(1).Text) > 0 Then 'Mes
    If Val(txtFecha(0).Text) = 0 Then
        txtFecha(1).Text = ""
        Call MsgBox("Debe especificar primeramente el año", vbOKOnly + vbInformation, "Validación")
        If txtFecha(0).Enabled Then txtCampo(0).SetFocus
        Exit Sub
    End If
    If Val(txtFecha(1).Text) <= 0 Or Val(txtFecha(1).Text) > 12 Then 'el Mes debe estar 1-12
        Call MsgBox("Mes incorrecto, favor de verificar", vbOKOnly + vbInformation, "Validación")
        txtFecha(1).SetFocus
        Exit Sub
    End If
    
ElseIf Index = 2 And Len(txtFecha(2).Text) > 0 Then 'Dia
    If Val(txtFecha(1).Text) = 0 Then
        txtFecha(2).Text = ""
        Call MsgBox("Debe especificar primeramente el mes", vbInformation + vbOKOnly, "Validación")
        If txtFecha(1).Enabled Then txtFecha(1).SetFocus
        Exit Sub
    End If
    If Not IsDate(Val(txtFecha(0)) & "/" & Right("0" & Val(txtFecha(1).Text), 2) & "/" & Right("0" & Val(txtFecha(2).Text), 2)) Then
        Call MsgBox("Fecha incorrecta, favor de correguirla: ", vbOKOnly, "Validación")
        txtFecha(2).Text = ""
        txtFecha(2).SetFocus
    End If
End If
End Sub

Private Sub txtOficio_Change()
comboOficios.Clear
comboOficios.Text = ""
txtExpediente.Text = ""
End Sub

Private Sub txtOficio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdContinuar_Click
    Exit Sub
End If
End Sub
