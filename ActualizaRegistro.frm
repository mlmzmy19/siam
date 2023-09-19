VERSION 5.00
Begin VB.Form ActualizaRegistro 
   Caption         =   "Actualización de datos de Registro"
   ClientHeight    =   8025
   ClientLeft      =   8490
   ClientTop       =   4065
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   12015
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Height          =   6240
      Left            =   0
      TabIndex        =   18
      Top             =   1665
      Width           =   11895
      Begin VB.CommandButton cmdCancela 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Cancelar"
         Height          =   465
         Left            =   10665
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   5445
         Width           =   1140
      End
      Begin VB.ComboBox cmbProductos 
         BackColor       =   &H8000000F&
         Height          =   315
         ItemData        =   "ActualizaRegistro.frx":0000
         Left            =   360
         List            =   "ActualizaRegistro.frx":0002
         TabIndex        =   9
         ToolTipText     =   "Institución"
         Top             =   2790
         Visible         =   0   'False
         Width           =   11100
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   9000
         TabIndex        =   20
         Top             =   4275
         Width           =   2310
      End
      Begin VB.ComboBox cmbUnidades 
         BackColor       =   &H8000000F&
         DataField       =   "idoriuni"
         Height          =   315
         ItemData        =   "ActualizaRegistro.frx":0004
         Left            =   360
         List            =   "ActualizaRegistro.frx":001A
         TabIndex        =   2
         ToolTipText     =   "Unidad de Origen"
         Top             =   630
         Width           =   11100
      End
      Begin VB.TextBox txtMemo 
         BackColor       =   &H8000000F&
         DataField       =   "memorando"
         Height          =   285
         Left            =   405
         MaxLength       =   60
         TabIndex        =   10
         Tag             =   "c"
         ToolTipText     =   "Número de memo de envío"
         Top             =   3420
         Width           =   3435
      End
      Begin VB.ComboBox cmbClases 
         BackColor       =   &H8000000F&
         Height          =   315
         ItemData        =   "ActualizaRegistro.frx":0068
         Left            =   360
         List            =   "ActualizaRegistro.frx":006A
         TabIndex        =   3
         ToolTipText     =   "Clase de Institución"
         Top             =   1260
         Width           =   11100
      End
      Begin VB.ComboBox cmbInst 
         BackColor       =   &H8000000F&
         Height          =   315
         ItemData        =   "ActualizaRegistro.frx":006C
         Left            =   315
         List            =   "ActualizaRegistro.frx":006E
         TabIndex        =   8
         ToolTipText     =   "Institución"
         Top             =   2295
         Width           =   11145
      End
      Begin VB.TextBox txtRecepcion 
         BackColor       =   &H8000000F&
         DataField       =   "recepción"
         Height          =   285
         Left            =   8370
         MaxLength       =   20
         TabIndex        =   12
         Tag             =   "f"
         ToolTipText     =   "Recepción del Expediente en el área de Sanciones"
         Top             =   3465
         Width           =   3075
      End
      Begin VB.ComboBox cmbTurnar 
         BackColor       =   &H8000000F&
         DataField       =   "idrestur"
         Height          =   315
         ItemData        =   "ActualizaRegistro.frx":0070
         Left            =   345
         List            =   "ActualizaRegistro.frx":0072
         TabIndex        =   13
         ToolTipText     =   "Responsable a quien se turna el expediente"
         Top             =   3990
         Width           =   11055
      End
      Begin VB.CommandButton cmdActualiza 
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   10890
         Picture         =   "ActualizaRegistro.frx":0074
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4590
         Width           =   615
      End
      Begin VB.TextBox txtFechaMemo 
         BackColor       =   &H8000000F&
         DataField       =   "fecha_memorando"
         Height          =   285
         Left            =   4500
         MaxLength       =   20
         TabIndex        =   11
         Tag             =   "f"
         ToolTipText     =   "Fecha del Memo de Envío"
         Top             =   3465
         Width           =   3435
      End
      Begin VB.TextBox txtObs 
         BackColor       =   &H8000000F&
         Height          =   1410
         Left            =   315
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Tag             =   "f"
         ToolTipText     =   "Observaciones"
         Top             =   4635
         Width           =   10140
      End
      Begin VB.TextBox txtBuscarIF 
         Height          =   285
         Left            =   6255
         TabIndex        =   6
         Top             =   1845
         Width           =   4605
      End
      Begin VB.CommandButton cmdBusIF 
         Caption         =   "Sig"
         Height          =   330
         Left            =   10935
         TabIndex        =   7
         Top             =   1845
         Width           =   510
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2880
         TabIndex        =   19
         Top             =   1755
         Width           =   2400
         Begin VB.OptionButton opcIF 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Todas"
            Height          =   285
            Index           =   1
            Left            =   1215
            TabIndex        =   5
            Top             =   90
            Value           =   -1  'True
            Width           =   1005
         End
         Begin VB.OptionButton opcIF 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Vigentes"
            Height          =   285
            Index           =   0
            Left            =   45
            TabIndex        =   4
            Top             =   90
            Width           =   1005
         End
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Unidad de origen:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   30
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "No. de memorando de envío:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   405
         TabIndex        =   29
         Top             =   3165
         Width           =   2085
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sector (Clase Institución):"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   28
         Top             =   1035
         Width           =   1800
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Institución:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   27
         Top             =   1980
         Width           =   765
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha de memorando de envío:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   4500
         TabIndex        =   26
         Top             =   3195
         Width           =   2280
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha de recepción del área de sanciones:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   8325
         TabIndex        =   25
         Top             =   3210
         Width           =   3075
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Turnar expediente a:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   24
         Top             =   3705
         Width           =   1470
      End
      Begin VB.Label etiObs 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Observaciones:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   23
         Top             =   4380
         Width           =   1110
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Producto:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   315
         TabIndex        =   22
         Top             =   2565
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Busca IF:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   5490
         TabIndex        =   21
         Top             =   1890
         Width           =   675
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   1692
      Left            =   1890
      TabIndex        =   0
      Top             =   12
      Width           =   10005
      Begin VB.TextBox txtNuevoExp 
         BackColor       =   &H8000000F&
         Height          =   870
         Left            =   270
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         Tag             =   "c"
         Top             =   765
         Width           =   9510
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Expediente(s) a actualizar:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   17
         Top             =   540
         Width           =   1860
      End
      Begin VB.Label Eti 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Actualización de Registro"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   4
         Left            =   1320
         TabIndex        =   16
         Top             =   285
         Width           =   7215
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Image Image1 
      Height          =   1632
      Left            =   18
      Picture         =   "ActualizaRegistro.frx":0DFE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1824
   End
End
Attribute VB_Name = "ActualizaRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public msRegs As String 'Registros a Actualizar
Public msExp As String
Dim miIns As Integer 'Guarda el valor inicial despues de la consulta antes de la actualización
Dim miUnidad As Integer
Dim miClase As Integer
Dim miInstitucion As Integer
Dim miTurnar As Integer
Dim msMemo As String
Dim mdMemo As Date
Dim mdRecepcion As Date
Dim msObs As String
Dim mbCambio As Boolean


Private Sub cmbClases_Click()
ActualizaComboIF
If cmbClases.ListIndex >= 0 Then
    If cmbClases.ItemData(cmbClases.ListIndex) <> miClase Then
        mbCambio = True
    End If
Else
    If miClase <> 0 Then
        mbCambio = True
    End If
End If
End Sub

Private Sub cmbInst_Click()
If cmbInst.ListIndex >= 0 Then
    If cmbInst.ItemData(cmbInst.ListIndex) <> miInstitucion Then
        mbCambio = True
    End If
Else
    If miInstitucion <> 0 Then
        mbCambio = True
    End If
End If
End Sub

Private Sub cmbTurnar_Click()
If cmbTurnar.ListIndex >= 0 Then
    If cmbTurnar.ItemData(cmbTurnar.ListIndex) <> miTurnar Then
        mbCambio = True
    End If
Else
    If miTurnar <> 0 Then
        mbCambio = True
    End If
End If
End Sub

Private Sub cmbUnidades_KeyPress(KeyAscii As Integer)
If cmbUnidades.ListIndex >= 0 Then
    If cmbUnidades.ItemData(cmbUnidades.ListIndex) <> miUnidad Then
        mbCambio = True
    End If
Else
    If miUnidad <> 0 Then
        mbCambio = True
    End If
End If
End Sub

Private Sub cmdActualiza_Click()
Dim iUni As Integer, iCla As Integer, iIns As Integer, iTur As Integer, sObs As String, sMemo As String, sFMemo As String, sRecep As String, sCamposXCambiar As String
Dim bCont As Boolean
Dim adors As New ADODB.Recordset
On Error GoTo salir:
If cmbUnidades.ListIndex >= 0 Then
    iUni = cmbUnidades.ItemData(cmbUnidades.ListIndex)
    If iUni = miUnidad Then
        iUni = 0
    Else
        bCont = True
        sCamposXCambiar = sCamposXCambiar & "Unidad de origen, "
    End If
End If
If cmbClases.ListIndex >= 0 Then
    iCla = cmbClases.ItemData(cmbClases.ListIndex)
    If iCla = miClase Then
        iCla = 0
    Else
        bCont = True
        sCamposXCambiar = sCamposXCambiar & "Sector, "
    End If
End If
If cmbInst.ListIndex >= 0 Then
    iIns = cmbInst.ItemData(cmbInst.ListIndex)
    If iIns = miInstitucion Then
        iIns = 0
    Else
        bCont = True
        sCamposXCambiar = sCamposXCambiar & "Institución, "
    End If
End If
If cmbTurnar.ListIndex >= 0 Then
    iTur = cmbTurnar.ItemData(cmbTurnar.ListIndex)
    If iTur = miTurnar Then
        iTur = 0
    Else
        bCont = True
        sCamposXCambiar = sCamposXCambiar & "Turnar a, "
    End If
End If
If Len(txtMemo.Text) > 0 Then
    sMemo = txtMemo.Text
    If sMemo <> msMemo Then
        bCont = True
        sCamposXCambiar = sCamposXCambiar & "Memorando, "
    End If
End If
If IsDate(txtFechaMemo.Text) Then
    sFMemo = Format(CDate(txtFechaMemo.Text), "dd/mm/yyyy")
    If CDate(sFMemo) <> mdMemo Then
        bCont = True
        sCamposXCambiar = sCamposXCambiar & "Fecha del Memorando, "
    Else
        sFMemo = "VARIOS"
    End If
End If
If IsDate(txtRecepcion.Text) Then
    sRecep = CDate(txtRecepcion.Text)
    If CDate(sRecep) <> mdRecepcion Then
        bCont = True
        sCamposXCambiar = sCamposXCambiar & "Fecha de Recepción, "
    Else
        sRecep = "VARIOS"
    End If
End If
If Len(txtObs.Text) > 0 Then
    sObs = txtObs.Text
    If sObs <> msObs Then
        bCont = True
        sCamposXCambiar = sCamposXCambiar & "Observaciones, "
    End If
End If
If Not bCont Then
    MsgBox "No hay cambios que realizar, favor de especificarlos", vbInformation + vbOKOnly, "Validación"
    Exit Sub
End If
If MsgBox("Está seguro de realizar la actualización con los datos especificados en: " & sCamposXCambiar, vbYesNo + vbQuestion + vbDefaultButton2, "Validación") = vbNo Then
    Exit Sub
End If

    

adors.Open "{call paq_registro.actualizadatos('" & msRegs & "'," & iUni & "," & iCla & "," & iIns & ",'" & sMemo & "','" & sFMemo & "','" & sRecep & "'," & iTur & ",'" & sObs & "'," & giUsuario & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
If adors(0) > 0 Then
    MsgBox "La actualización se realizó exitósamente " & adors(1), vbOKOnly, "Aviso"
    gs1 = "OK"
    Unload Me
Else
    MsgBox "La actualización no se realizó exitósamente " & adors(1), vbOKOnly, "Aviso"
End If

Exit Sub
salir:
Dim yError As Integer
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


Private Sub cmdBusIF_Click()
Dim i As Integer, iPos As Integer
If Len(txtBuscarIF.Text) > 0 And cmbInst.ListCount > 0 Then
    iPos = cmbInst.ListIndex
    If iPos = cmbInst.ListCount - 1 Then
        i = -1
    Else
        i = BuscaCombo(cmbInst, txtBuscarIF.Text, 0, True, 0, iPos + 1)
    End If
    If i >= 0 Then
        cmbInst.ListIndex = i
    ElseIf iPos >= 0 Then
        cmbInst.ListIndex = -1
    End If
End If
End Sub


Private Sub cmdLimpiarRegs_Click()

End Sub

Private Sub cmdCancela_Click()
If mbCambio Then
    If MsgBox("Está seguro de salir sin guardar los cambios", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmación") = vbNo Then
        Exit Sub
    End If
    gs1 = "Cancelar"
End If
Unload Me
End Sub

Private Sub Form_Activate()
Dim adors As New ADODB.Recordset, i As Integer
If Len(msRegs) <= 2 Then
    MsgBox "no hay registros que actualizar", vbInformation + vbOKOnly, ""
    Unload Me
    Exit Sub
End If
adors.Open "{call paq_registro.busca_datosxact('" & msRegs & "')}", gConSql, adOpenForwardOnly, adLockReadOnly
If Not adors.EOF Then
    If adors(0) > 0 Then 'Unidad de origen
        i = BuscaCombo(cmbUnidades, adors(0), True)
        If i >= 0 Then
            cmbUnidades.ListIndex = i
        End If
        miUnidad = adors(0)
    Else
        cmbUnidades.ListIndex = -1
        cmbUnidades.Text = "VARIOS"
        miUnidad = 0
    End If
    If adors(1) > 0 Then 'CLASE
        i = BuscaCombo(cmbClases, adors(1), True)
        If i >= 0 Then
            cmbClases.ListIndex = i
        End If
        miClase = adors(1)
    Else
        cmbClases.ListIndex = -1
        cmbClases.Text = "VARIOS"
        miClase = 0
    End If
    If adors(2) <> -9999 Then 'Institución
        i = BuscaCombo(cmbInst, adors(2), True)
        If i >= 0 Then
            cmbInst.ListIndex = i
        End If
        miInstitucion = adors(2)
    Else
        cmbInst.ListIndex = -1
        cmbInst.Text = "VARIOS"
        miInstitucion = 0
    End If
    If adors(3) <> "-9999" Then 'Memorando
        txtMemo.Text = adors(3)
        msMemo = adors(3)
    Else
        txtMemo.Text = "VARIOS"
        msMemo = "VARIOS"
    End If
    If adors(4) <> CDate("01/01/1901") Then 'Fecha Memorando
        txtFechaMemo.Text = Format(adors(4), "dd/mm/yyyy")
        mdMemo = adors(4)
    Else
        txtFechaMemo.Text = "VARIOS"
        mdMemo = CDate("01/01/1901")
    End If
    If adors(5) <> CDate("01/01/1901") Then 'Fecha Recepción
        txtRecepcion.Text = Format(adors(5), "dd/mm/yyyy")
        mdRecepcion = adors(5)
    Else
        txtRecepcion.Text = "VARIOS"
        mdRecepcion = CDate("01/01/1901")
    End If
    If adors(6) <> -9999 Then 'Turnar
        i = BuscaCombo(cmbTurnar, adors(6), True)
        If i >= 0 Then
            cmbTurnar.ListIndex = i
        End If
        miTurnar = adors(6)
    Else
        cmbTurnar.ListIndex = -1
        cmbTurnar.Text = "VARIOS"
        miTurnar = 0
    End If
    If adors(7) <> "-9999" Then 'Observaciones
        txtObs.Text = adors(7)
        msObs = adors(7)
    Else
        txtObs.Text = "VARIOS"
        msObs = "VARIOS"
    End If
End If

End Sub

Private Sub Form_Load()
LlenaCombo cmbUnidades, "select id,descripción from unidades order by 2", "", True
LlenaCombo cmbClases, "select id,descripción from claseinstitución where baja=0 order by 2", "", True
LlenaCombo cmbTurnar, "select id,descripción from usuariossistema where baja=0 and responsable<>0 order by 2", "", True
'mbInicio = True
mbCambio = False
'mbLimpia = False
txtNuevoExp.Text = msExp
End Sub

Sub ActualizaComboIF()
Dim adors As New ADODB.Recordset
If cmbClases.ListIndex >= 0 Then
    adors.Open "{call paq_registro.institucion(" & IIf(opcIF(0).Value, 0, 1) & "," & cmbClases.ItemData(cmbClases.ListIndex) & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
Else
    adors.Open "{call paq_registro.institucion(" & IIf(opcIF(0).Value, 0, 1) & ",0)}", gConSql, adOpenForwardOnly, adLockReadOnly
End If
Call LlenaComboCursor(cmbInst, adors)
End Sub

Private Sub txtFechaMemo_Click()
If txtFechaMemo.Text <> "VARIOS" And IsDate(txtFechaMemo.Text) Then
    If CDate(txtFechaMemo.Text) <> mdMemo Then
        mbCambio = True
    End If
Else
    If mdMemo <> CDate("01/01/1901") Then
        mbCambio = True
    End If
End If
End Sub

Private Sub txtMemo_LostFocus()
If txtMemo.Text <> "VARIOS" And Len(txtMemo.Text) > 0 Then
    If txtMemo.Text <> msMemo Then
        mbCambio = True
    End If
Else
    If msMemo <> "VARIOS" Then
        mbCambio = True
    End If
End If
End Sub

Private Sub txtObs_LostFocus()
If txtObs.Text <> "VARIOS" And Len(txtObs.Text) > 0 Then
    If txtObs.Text <> msObs Then
        mbCambio = True
    End If
Else
    If msObs <> "VARIOS" Then
        mbCambio = True
    End If
End If
End Sub

Private Sub txtRecepcion_LostFocus()
If txtRecepcion.Text <> "VARIOS" And IsDate(txtRecepcion.Text) Then
    If CDate(txtRecepcion.Text) <> mdRecepcion Then
        mbCambio = True
    End If
Else
    If mdRecepcion <> CDate("01/01/1901") Then
        mbCambio = True
    End If
End If

End Sub
