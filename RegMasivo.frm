VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form RegMasivo 
   Caption         =   "Registro masivo de Expedientes Trimestrales"
   ClientHeight    =   10755
   ClientLeft      =   240
   ClientTop       =   180
   ClientWidth     =   16260
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10755
   ScaleMode       =   0  'User
   ScaleWidth      =   20223.88
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      Height          =   5280
      Left            =   45
      TabIndex        =   34
      Top             =   3960
      Width           =   16170
      Begin MSComctlLib.ListView ListView1 
         Height          =   4980
         Left            =   135
         TabIndex        =   15
         Top             =   225
         Width           =   15855
         _ExtentX        =   27966
         _ExtentY        =   8784
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "idSol"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Expediente"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Memorando"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "FechaEnvío"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Unidad"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Incumplimiento"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "CausaSanción"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Key             =   "Institución"
            Text            =   "Institución"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Registro"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Recepción"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Clase"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Observaciones"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Caja"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000005&
      Caption         =   "Especifique los datos que van aplicar a todos los registros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   45
      TabIndex        =   28
      Top             =   9315
      Width           =   16170
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Registrar con datos especificados"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   13815
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   225
         Width           =   2310
      End
      Begin VB.TextBox txtRecepcion 
         BackColor       =   &H8000000F&
         DataField       =   "recepción"
         Height          =   285
         Left            =   135
         MaxLength       =   20
         TabIndex        =   16
         Tag             =   "f"
         ToolTipText     =   "Recepción del Expediente en el área de Sanciones"
         Top             =   495
         Width           =   2355
      End
      Begin VB.ComboBox cmbTurnar 
         BackColor       =   &H8000000F&
         DataField       =   "idrestur"
         Height          =   315
         ItemData        =   "RegMasivo.frx":0000
         Left            =   135
         List            =   "RegMasivo.frx":0002
         TabIndex        =   17
         ToolTipText     =   "Responsable a quien se turna el expediente"
         Top             =   1035
         Width           =   3450
      End
      Begin VB.TextBox txtObs 
         BackColor       =   &H8000000F&
         Height          =   915
         Left            =   3645
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Tag             =   "f"
         ToolTipText     =   "Observaciones"
         Top             =   405
         Width           =   10050
      End
      Begin MSForms.CommandButton cmdActualiza 
         Height          =   375
         Left            =   13995
         TabIndex        =   37
         Top             =   945
         Width           =   1995
         BackColor       =   12640511
         Caption         =   "Actualizar datos"
         Size            =   "3519;661"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label etiTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha de recepción:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   31
         Top             =   270
         Width           =   1470
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Turnar expediente a:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   180
         TabIndex        =   30
         Top             =   810
         Width           =   1470
      End
      Begin VB.Label etiObs 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Obs:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   3195
         TabIndex        =   29
         Top             =   495
         Width           =   330
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Especificar criterios de búsqueda de expedientes trimestrales para registro masivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2820
      Left            =   45
      TabIndex        =   20
      Top             =   1125
      Width           =   16170
      Begin VB.CommandButton cmdLimpiarRegs 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Limpiar Registros Encontrados"
         Height          =   510
         Left            =   12825
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2160
         Width           =   1500
      End
      Begin VB.CommandButton cmdLimpiar 
         Caption         =   "Limpiar Nva búsqueda"
         Height          =   375
         Left            =   13545
         TabIndex        =   13
         Top             =   315
         Width           =   2355
      End
      Begin VB.TextBox txtExp 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1305
         MaxLength       =   35
         TabIndex        =   1
         Tag             =   "c"
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox txtMemo 
         BackColor       =   &H8000000F&
         DataField       =   "n_cvepersona"
         Height          =   285
         Left            =   8055
         MaxLength       =   20
         TabIndex        =   2
         Tag             =   "c"
         ToolTipText     =   """Numero consecutivo de registro"""
         Top             =   315
         Width           =   3150
      End
      Begin VB.CommandButton cmdContinuar 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Buscar y Agregar"
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
         Left            =   14580
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2115
         Width           =   1410
      End
      Begin VB.ComboBox cmbIncum 
         BackColor       =   &H8000000F&
         DataField       =   "idoriuni"
         Height          =   315
         ItemData        =   "RegMasivo.frx":0004
         Left            =   7965
         List            =   "RegMasivo.frx":001A
         TabIndex        =   4
         ToolTipText     =   "Unidad de Origen"
         Top             =   1125
         Width           =   8040
      End
      Begin VB.ComboBox cmbCausa 
         BackColor       =   &H8000000F&
         DataField       =   "idmat"
         Height          =   315
         ItemData        =   "RegMasivo.frx":0068
         Left            =   225
         List            =   "RegMasivo.frx":006A
         TabIndex        =   5
         ToolTipText     =   "Materia de la Sanción"
         Top             =   1755
         Width           =   7725
      End
      Begin VB.ComboBox cmbUnidad 
         BackColor       =   &H8000000F&
         DataField       =   "idoridir"
         Height          =   315
         ItemData        =   "RegMasivo.frx":006C
         Left            =   225
         List            =   "RegMasivo.frx":006E
         TabIndex        =   3
         ToolTipText     =   "Dirección General de Origen"
         Top             =   1080
         Width           =   7680
      End
      Begin VB.ComboBox cmbClase 
         BackColor       =   &H8000000F&
         Height          =   315
         ItemData        =   "RegMasivo.frx":0070
         Left            =   8010
         List            =   "RegMasivo.frx":0072
         TabIndex        =   12
         ToolTipText     =   "Clase de Institución"
         Top             =   1710
         Width           =   8010
      End
      Begin VB.ComboBox cmbInst 
         BackColor       =   &H8000000F&
         Height          =   315
         ItemData        =   "RegMasivo.frx":0074
         Left            =   225
         List            =   "RegMasivo.frx":0076
         TabIndex        =   10
         ToolTipText     =   "Institución"
         Top             =   2430
         Width           =   7725
      End
      Begin VB.TextBox txtBuscarIF 
         Height          =   285
         Left            =   4725
         TabIndex        =   8
         Top             =   2115
         Width           =   2625
      End
      Begin VB.CommandButton cmdBusIF 
         Caption         =   "Sig"
         Height          =   330
         Left            =   7425
         TabIndex        =   9
         Top             =   2070
         Width           =   510
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1350
         TabIndex        =   21
         Top             =   2025
         Width           =   2175
         Begin VB.OptionButton opcIF 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Todas"
            Height          =   285
            Index           =   1
            Left            =   1170
            TabIndex        =   7
            Top             =   90
            Width           =   1005
         End
         Begin VB.OptionButton opcIF 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Vigentes"
            Height          =   285
            Index           =   0
            Left            =   45
            TabIndex        =   6
            Top             =   90
            Value           =   -1  'True
            Width           =   1005
         End
      End
      Begin MSForms.CheckBox chkRegistrados 
         Height          =   375
         Left            =   8325
         TabIndex        =   36
         Top             =   2205
         Width           =   2625
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4630;661"
         Value           =   "0"
         Caption         =   "Registrados (en SIAM)"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Expediente:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   180
         TabIndex        =   33
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Memorando:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   6930
         TabIndex        =   32
         Top             =   315
         Width           =   885
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cuasa de la sanción:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   27
         Top             =   1485
         Width           =   1485
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Incumplimiento:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   8010
         TabIndex        =   26
         Top             =   855
         Width           =   1095
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Área de origen:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   25
         Top             =   810
         Width           =   1080
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sector (Clase Institución):"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   8010
         TabIndex        =   24
         Top             =   1485
         Width           =   1800
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Institución:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   225
         TabIndex        =   23
         Top             =   2115
         Width           =   765
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Busca IF:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   4005
         TabIndex        =   22
         Top             =   2160
         Width           =   675
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   1110
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   16170
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8676
         Top             =   144
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   17
         ImageHeight     =   17
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RegMasivo.frx":0078
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RegMasivo.frx":043E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RegMasivo.frx":0804
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RegMasivo.frx":0CCA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RegMasivo.frx":1090
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RegMasivo.frx":1456
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RegMasivo.frx":181C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RegMasivo.frx":1BE2
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RegMasivo.frx":1FA8
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RegMasivo.frx":236E
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RegMasivo.frx":2734
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   915
         Left            =   135
         Picture         =   "RegMasivo.frx":2AFA
         Stretch         =   -1  'True
         Top             =   135
         Width           =   1020
      End
      Begin VB.Label Eti 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Módulo de Registro Masivo (EXPEDIENTES TRIMESTRALES)"
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
         Height          =   420
         Index           =   4
         Left            =   4230
         TabIndex        =   11
         Top             =   360
         Width           =   9360
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "RegMasivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msSol As String
Dim miColOrden As Long
Dim msClases(200) As String


Private Sub cmbClase_Click()
ActualizaComboIF
End Sub

Private Sub cmdActualiza_Click()
ActualizaRegistros
End Sub

Sub ActualizaRegistros()
Dim i As Long, sSol As String, sExp As String
Dim adors As New ADODB.Recordset
For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(i).Selected Then
        If Val(ListView1.ListItems(i).SubItems(8)) > 0 Then
            sSol = sSol & Val(ListView1.ListItems(i).SubItems(8)) & "|"
            sExp = sExp & ListView1.ListItems(i).SubItems(1) & ", "
        End If
    End If
Next
If InStr(sSol, "|") > 0 Then 'Habilita módulo de actualización
    With ActualizaRegistro
        .msRegs = sSol
        .msExp = Mid(sExp, 1, Len(sExp) - 2)
        .Show vbModal
        If gs1 = "OK" Then
            If ListView1.Sorted Then  ' VERIFICA SI ESTÁ ORDENADO PARA QUITAR EL ORDENAMIENTO
                ListView1.Sorted = False
            End If

            'Quita los asuntos seleccionados
            For i = ListView1.ListItems.Count To 1 Step -1
                If ListView1.ListItems(i).Selected Then
                    msSol = Replace(msSol, "|" & ListView1.ListItems(i).Text & "|", "|")
                    ListView1.ListItems.Remove (i)
                End If
            Next
            'Los ingresa nuevamente
            If adors.State > 0 Then adors.Close
            adors.Open "{call paq_registro.busca_datosRegXIDs('" & sSol & "')}", gConSql, adOpenForwardOnly, adLockReadOnly
            If Not adors.EOF Then
                If Not Frame3.Enabled Then Frame3.Enabled = True
            End If
            i = ListView1.ListItems.Count + 1
            Do While Not adors.EOF
                If InStr(msSol, "|" & adors(0) & "|") = 0 Then
                    msSol = msSol & adors(0) & "|"
                    ListView1.ListItems.Add i, , adors(0) 'idsol
                    ListView1.ListItems(i).SubItems(1) = IIf(IsNull(adors(1)), "", adors(1)) 'Expediente
                    ListView1.ListItems(i).SubItems(2) = IIf(IsNull(adors(2)), "", adors(2)) 'Memorando
                    ListView1.ListItems(i).SubItems(3) = IIf(IsNull(adors(3)), "", adors(3)) 'Folio envio
                    ListView1.ListItems(i).SubItems(4) = IIf(IsNull(adors(4)), "", adors(4)) 'Fecha envio físico
                    ListView1.ListItems(i).SubItems(5) = IIf(IsNull(adors(5)), "", adors(5)) 'Unidad
                    ListView1.ListItems(i).SubItems(6) = IIf(IsNull(adors(6)), "", adors(6)) 'Incumplimiento
                    ListView1.ListItems(i).SubItems(7) = IIf(IsNull(adors(7)), "", adors(7)) 'Sanción
                    ListView1.ListItems(i).SubItems(8) = IIf(IsNull(adors(8)), "", adors(8)) 'Institución
                    ListView1.ListItems(i).SubItems(9) = IIf(IsNull(adors(9)), "", adors(9)) 'Id Reg
                    ListView1.ListItems(i).SubItems(10) = IIf(IsNull(adors(10)), "", adors(10)) 'Recepción
                    ListView1.ListItems(i).SubItems(11) = IIf(IsNull(adors(11)), "", adors(11)) 'Clase
                    ListView1.ListItems(i).SubItems(12) = IIf(IsNull(adors(12)), "", adors(12)) 'Observaciones
                    ListView1.ListItems(i).Tag = adors(0) 'Guarda el id
                    i = i + 1
                    n = n + 1
                End If
                adors.MoveNext
            Loop
            
        End If
    End With
ElseIf ListView1.ListItems.Count > 0 Then
    Call MsgBox("Solo puede actualiza asuntos registrados... Favor de verificar", vbInformation + vbOKOnly, "Validación")
End If
Exit Sub
salir:
Dim yError As Long
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
Dim i As Long, iPos As Long
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

Private Sub cmdContinuar_Click()
Dim adors As New ADODB.Recordset
Dim iCla As Long, iIns As Long, iInc As Long, iSan As Long, iUni As Long
Dim i As Long, n As Long
On Error GoTo salir:
If cmbClase.ListIndex >= 0 Then
    iCla = cmbClase.ItemData(cmbClase.ListIndex)
End If
If cmbInst.ListIndex >= 0 Then
    iIns = cmbInst.ItemData(cmbInst.ListIndex)
End If
If cmbIncum.ListIndex >= 0 Then
    iInc = cmbIncum.ItemData(cmbIncum.ListIndex)
End If
If cmbCausa.ListIndex >= 0 Then
    iSan = cmbCausa.ItemData(cmbCausa.ListIndex)
End If
If cmbUnidad.ListIndex >= 0 Then
    iUni = cmbUnidad.ItemData(cmbUnidad.ListIndex)
End If

'ListView1.ListItems.Clear
If adors.State Then adors.Close
If chkRegistrados.Value Then
    adors.Open "{call paq_registro.busca_datosReg('" & txtExp.Text & "','" & txtMemo.Text & "'," & iUni & "," & iCla & "," & iIns & "," & iInc & "," & iSan & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
    If Not adors.EOF And ListView1.ListItems.Count = 0 Then
        'If Frame3.Enabled Then Frame3.Enabled = False
    End If
Else
    adors.Open "{call paq_registro.busca_datosXReg('" & txtExp.Text & "','" & txtMemo.Text & "'," & iUni & "," & iCla & "," & iIns & "," & iInc & "," & iSan & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
    If Not adors.EOF Then
        If Not Frame3.Enabled Then Frame3.Enabled = True
    End If
End If

If ListView1.Sorted Then  ' VERIFICA SI ESTÁ ORDENADO PARA QUITAR EL ORDENAMIENTO
    ListView1.Sorted = False
End If

i = ListView1.ListItems.Count + 1
Do While Not adors.EOF
    If InStr(msSol, "|" & adors(0) & "|") = 0 Then
        msSol = msSol & adors(0) & "|"
        ListView1.ListItems.Add i, , adors(0) 'idsol
        ListView1.ListItems(i).SubItems(1) = IIf(IsNull(adors(1)), "", adors(1)) 'Expediente
        ListView1.ListItems(i).SubItems(2) = IIf(IsNull(adors(2)), "", adors(2)) 'Memorando
        'ListView1.ListItems(i).SubItems(3) = IIf(IsNull(adors(3)), "", adors(3)) 'Folio envio
        ListView1.ListItems(i).SubItems(3) = IIf(IsNull(adors(3)), "", adors(3)) 'Fecha envio físico
        ListView1.ListItems(i).SubItems(4) = IIf(IsNull(adors(4)), "", adors(4)) 'Unidad
        ListView1.ListItems(i).SubItems(5) = IIf(IsNull(adors(5)), "", adors(5)) 'Incumplimiento
        ListView1.ListItems(i).SubItems(6) = IIf(IsNull(adors(6)), "", adors(6)) 'Sanción
        ListView1.ListItems(i).SubItems(7) = IIf(IsNull(adors(7)), "", adors(7)) 'Institución
        ListView1.ListItems(i).SubItems(8) = IIf(IsNull(adors(8)), "", adors(8)) 'Id Reg
        ListView1.ListItems(i).SubItems(9) = IIf(IsNull(adors(9)), "", adors(9)) 'Recepción
        ListView1.ListItems(i).SubItems(10) = IIf(IsNull(adors(10)), "", adors(10)) 'Clase
        ListView1.ListItems(i).SubItems(11) = IIf(IsNull(adors(11)), "", adors(11)) 'Observaciones
        ListView1.ListItems(i).SubItems(12) = IIf(IsNull(adors(12)), "", adors(12)) 'Caja
        ListView1.ListItems(i).Tag = adors(0) 'Guarda el id
        i = i + 1
        n = n + 1
    End If
    adors.MoveNext
Loop
MsgBox n & " registro(s) agregado(s)", vbOKOnly + vbInformation, ""
Exit Sub
salir:
Dim yError As Long
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

Private Sub cmdLimpiar_Click()
txtExp.Text = ""
txtMemo.Text = ""
cmbUnidad.Text = ""
If cmbUnidad.ListCount > 0 Then
    cmbUnidad.ListIndex = 0
End If
If cmbIncum.ListCount > 0 Then
    cmbIncum.ListIndex = 0
End If
If cmbCausa.ListCount > 0 Then
    cmbCausa.ListIndex = 0
End If
If cmbClase.ListCount > 0 Then
    cmbClase.ListIndex = 0
End If
If cmbInst.ListCount > 0 Then
    cmbInst.ListIndex = -1
End If
End Sub

Private Sub cmdLimpiarRegs_Click()
If ListView1.ListItems.Count > 0 Then
    If MsgBox("Estás seguro de quitar todos los registros encontrados en previas búsquedas", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmación") = vbYes Then
        ListView1.ListItems.Clear
        msSol = "|"
    End If
End If
End Sub

Private Sub Command1_Click()
Dim adors As New ADODB.Recordset, i As Long
Dim sRegs() As String, msSols() As String, sClas() As String
Dim s As String, n As Long, v As Long, sSol As String, sReg As String, sCla As String, sScript As String
On Error GoTo salir:
For i = 1 To ListView1.ListItems.Count
    If Val(ListView1.ListItems(i).SubItems(8)) = 0 Then
        sSol = sSol & Val(ListView1.ListItems(i).Text) & "|"
    End If
Next
If InStr(sSol, "|") <= 0 Then
    MsgBox "No se tienen asuntos que aplicar...Favor de verificar", vbOKOnly + vbInformation, "Validación"
    Exit Sub
End If
If Not IsDate(txtRecepcion.Text) Then
    MsgBox "Debe especificar la fecha de Recepción de los expedientes...", vbOKOnly + vbInformation, "Validación"
    Exit Sub
End If
If cmbTurnar.ListIndex < 0 Then
    MsgBox "Debe a quien se turnan los asuntos...", vbOKOnly + vbInformation, "Validación"
    Exit Sub
End If

i = cmbTurnar.ItemData(cmbTurnar.ListIndex)
sScript = "{call paq_registro.guardadatos('" & sSol & "','" & txtRecepcion & "'," & i & ",'" & txtObs.Text & "'," & giUsuario & ")}"

If Len(sScript) > 4000 Then
    MsgBox "La cadena del script (" & Len(sScript) & ") a executar supera el máximo permitido de 4000 caracteres. Favor de eligir menos asuntos", vbOKOnly + vbInformation, "Validación"
    Exit Sub
End If

If MsgBox("Está seguro de Registrar Masivamente con los datos especificados", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmación") = vbNo Then
    Exit Sub
End If


adors.Open sScript, gConSql, adOpenForwardOnly, adLockReadOnly
sSol = ""
If adors(0) > 0 Then
    MsgBox "El registro se realizó correctamente: " & "(" & adors(1) & ")", vbOKOnly + vbInformation, ""
    Do While Not adors.EOF
        sSol = sSol & adors(2)
        sReg = sReg & adors(3)
        sCla = sCla & adors(4)
        adors.MoveNext
    Loop
    msSols = Split(sSol, "|")
    sRegs = Split(sReg, "|")
    sClas = Split(sCla, "|")
    For i = 1 To ListView1.ListItems.Count
        n = Val(ListView1.ListItems(i).Text) 'Valor a buscar
        v = Busca_Bin(n, msSols)
        If v >= 0 And v <= UBound(sRegs) Then
            ListView1.ListItems(i).SubItems(8) = sRegs(v)
            ListView1.ListItems(i).SubItems(9) = txtRecepcion.Text
            If Val(sClas(v)) < 200 And Val(sClas(v)) > 0 Then
                ListView1.ListItems(i).SubItems(10) = msClases(Val(sClas(v)))
            Else
                ListView1.ListItems(i).SubItems(10) = ""
            End If
            ListView1.ListItems(i).SubItems(11) = txtObs.Text
        End If
    Next
Else
    MsgBox "El registro no se realizó correctamente verificar excepción " & "(" & adors(1) & ")", vbOKOnly + vbInformation, ""
End If
Exit Sub
salir:
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

Function Busca_Bin(iDato As Long, sDatos() As String) As Long
Dim i As Long, iMax As Long, iMin As Long
iMax = UBound(sDatos)
iAnt = iMax
Busca_Bin = -1
Do While iMin <> iMax
    If iMax - iMin = 1 Then
        If i = iMin Then
            i = iMax
        Else
            i = iMin
        End If
    Else
        i = iMin + Round((iMax - iMin + 0.0001) / 2, 0)
    End If
    If i > iMax Or i < iMin Then
        Exit Do
    End If
    If Val(sDatos(i)) = iDato Then
        Busca_Bin = i
        Exit Function
    End If
    If Val(sDatos(i)) < iDato Then
        iMin = i
    Else
        iMax = i
    End If
Loop
End Function


Private Sub Form_Load()
Dim adors As New ADODB.Recordset
adors.Open "{call paq_registro.unidad(1)}", gConSql, adOpenForwardOnly, adLockReadOnly
Call LlenaComboCursor(cmbUnidad, adors)
If adors.State Then adors.Close
adors.Open "{call paq_registro.clase(1)}", gConSql, adOpenForwardOnly, adLockReadOnly
Call LlenaComboCursor(cmbClase, adors)
If adors.State Then adors.Close
adors.Open "{call paq_registro.incumplimiento(1)}", gConSql, adOpenForwardOnly, adLockReadOnly
Call LlenaComboCursor(cmbIncum, adors)
If adors.State Then adors.Close
adors.Open "{call paq_registro.causasan(1)}", gConSql, adOpenForwardOnly, adLockReadOnly
Call LlenaComboCursor(cmbCausa, adors)
msSol = "|"
LlenaCombo cmbTurnar, "select id,descripción from usuariossistema where baja=0 and responsable<>0 order by 2", "", True
cmdLimpiar_Click
If adors.State Then adors.Close
adors.Open "select id,descripción from claseinstitución order by 1", gConSql, adOpenForwardOnly, adLockReadOnly
Do While Not adors.EOF 'Agrega las clases en el arreglo
    If adors(0) < 200 And adors(0) > 0 Then
        msClases(adors(0)) = adors(1)
    End If
    adors.MoveNext
Loop
End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ListView1.SortKey = ColumnHeader.Index - 1 'Indico al ListView que ordene según los datos de la columna 1 (esta propiedad utiliza un valor que es igual al Indice de la columna - 1)
If miColOrden = ColumnHeader.Index Then
    ListView1.SortOrder = lvwDescending ' Ordena en forma descendente
    miColOrden = 0
Else
    ListView1.SortOrder = lvwAscending ' Ordena en forma ascendente
    miColOrden = ColumnHeader.Index
End If
ListView1.Sorted = True ' con esto se ordena la lista.
End Sub

Private Sub ListView1_DblClick()
ActualizaRegistros
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long
Dim n As Long
If KeyCode = 46 Then 'Borra registro seleccionados
    For i = ListView1.ListItems.Count To 1 Step -1
        If ListView1.ListItems(i).Selected Then
            msSol = Replace(msSol, "|" & ListView1.ListItems(i).Text & "|", "|")
            ListView1.ListItems.Remove (i)
            n = n + 1
        End If
    Next
    MsgBox "Se eliminaron " & n & " Registros", vbOKOnly, ""
End If
End Sub

Private Sub opcIF_Click(Index As Integer)
ActualizaComboIF
End Sub


Sub ActualizaComboIF()
Dim adors As New ADODB.Recordset
If cmbClase.ListIndex > 0 Then
    adors.Open "{call paq_registro.institucion(" & IIf(opcIF(0).Value, 0, 1) & "," & cmbClase.ItemData(cmbClase.ListIndex) & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
Else
    adors.Open "{call paq_registro.institucion(" & IIf(opcIF(0).Value, 0, 1) & ",0)}", gConSql, adOpenForwardOnly, adLockReadOnly
End If
Call LlenaComboCursor(cmbInst, adors)
End Sub

Private Sub txtRecepcion_KeyPress(KeyAscii As Integer)
Dim i As Long, i1 As Long
If KeyAscii = 13 Then
    i1 = txtRecepcion.TabIndex
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
KeyAscii = TeclaOprimida(txtRecepcion, KeyAscii, txtRecepcion.Tag, False)
End Sub

Private Sub txtRecepcion_LostFocus()
If IsDate(txtRecepcion.Text) Then
    txtRecepcion.Text = Format(CDate(txtRecepcion.Text), "dd/mm/yyyy")
End If
End Sub
