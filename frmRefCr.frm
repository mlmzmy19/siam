VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRefCruz 
   Caption         =   "Información para generar informe de referencias cruzadas"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CReport 
      Left            =   8280
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame4 
      Height          =   1230
      Left            =   0
      TabIndex        =   23
      Top             =   36
      Width           =   8970
      Begin VB.TextBox txtFin 
         Height          =   300
         Left            =   4140
         TabIndex        =   33
         Tag             =   "f"
         Top             =   720
         Width           =   3252
      End
      Begin VB.TextBox txtIni 
         Height          =   300
         Left            =   4140
         TabIndex        =   32
         Tag             =   "f"
         Top             =   324
         Width           =   3252
      End
      Begin VB.CommandButton cmdAntSig 
         Caption         =   "Sig"
         Height          =   372
         Index           =   1
         Left            =   7560
         TabIndex        =   31
         Top             =   288
         Width           =   624
      End
      Begin VB.CommandButton cmdAntSig 
         Caption         =   "Ant"
         Height          =   372
         Index           =   0
         Left            =   7560
         TabIndex        =   30
         Top             =   684
         Width           =   624
      End
      Begin VB.OptionButton opcRango 
         Caption         =   "Anual"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   28
         Top             =   270
         Width           =   780
      End
      Begin VB.OptionButton opcRango 
         Caption         =   "Bimestral"
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   27
         Top             =   495
         Width           =   1050
      End
      Begin VB.OptionButton opcRango 
         Caption         =   "Mensual"
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   26
         Top             =   855
         Width           =   1005
      End
      Begin VB.OptionButton opcRango 
         Caption         =   "Semanal"
         Height          =   285
         Index           =   3
         Left            =   1728
         TabIndex        =   25
         Top             =   405
         Width           =   1095
      End
      Begin VB.OptionButton opcRango 
         Caption         =   "Otro"
         Height          =   285
         Index           =   4
         Left            =   1716
         TabIndex        =   24
         Top             =   765
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "Del:"
         Height          =   276
         Index           =   0
         Left            =   3588
         TabIndex        =   7
         Top             =   384
         Width           =   372
      End
      Begin VB.Label Label3 
         Caption         =   "Al:"
         Height          =   240
         Index           =   1
         Left            =   3672
         TabIndex        =   29
         Top             =   816
         Width           =   288
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3585
      Left            =   0
      TabIndex        =   3
      Top             =   1305
      Width           =   8970
      Begin VB.TextBox txtSubtítulo 
         Height          =   1005
         Left            =   210
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   4320
         Visible         =   0   'False
         Width           =   7305
      End
      Begin VB.TextBox txtTítulo 
         Height          =   375
         Left            =   210
         MaxLength       =   250
         TabIndex        =   20
         Top             =   3810
         Visible         =   0   'False
         Width           =   7275
      End
      Begin VB.Frame Frame3 
         Height          =   1755
         Left            =   7560
         TabIndex        =   17
         Top             =   1200
         Width           =   1290
         Begin VB.CommandButton cmdBotón 
            Caption         =   "&Procesa informe"
            Enabled         =   0   'False
            Height          =   555
            Index           =   0
            Left            =   150
            TabIndex        =   19
            Top             =   360
            Width           =   1000
         End
         Begin VB.CommandButton cmdBotón 
            Caption         =   "&Salir"
            Height          =   405
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   1200
            Width           =   1000
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ejemplo"
         Height          =   1425
         Left            =   270
         TabIndex        =   8
         Top             =   1980
         Width           =   7135
         Begin VB.TextBox txtContenido 
            Appearance      =   0  'Flat
            Height          =   825
            Left            =   2275
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   500
            Width           =   4740
         End
         Begin VB.TextBox txtColumna 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   3
            Left            =   5425
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   225
            Width           =   1590
         End
         Begin VB.TextBox txtColumna 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   2
            Left            =   3850
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   225
            Width           =   1590
         End
         Begin VB.TextBox txtColumna 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   1
            Left            =   2275
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   225
            Width           =   1590
         End
         Begin VB.TextBox txtColumna 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   0
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   225
            Width           =   2200
         End
         Begin VB.TextBox txtRenglón 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   1035
            Width           =   2200
         End
         Begin VB.TextBox txtRenglón 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   765
            Width           =   2200
         End
         Begin VB.TextBox txtRenglón 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   495
            Width           =   2200
         End
      End
      Begin VB.ComboBox ComboVarios 
         DataField       =   "idtip"
         DataSource      =   "datAsunto"
         Height          =   315
         Index           =   1
         ItemData        =   "frmRefCr.frx":0000
         Left            =   3960
         List            =   "frmRefCr.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "n"
         ToolTipText     =   "Valor que desea como título en las columnas"
         Top             =   495
         Width           =   3400
      End
      Begin VB.ListBox List1 
         Height          =   645
         Index           =   0
         ItemData        =   "frmRefCr.frx":0004
         Left            =   300
         List            =   "frmRefCr.frx":000E
         TabIndex        =   2
         Top             =   1110
         Width           =   3400
      End
      Begin VB.ComboBox ComboVarios 
         DataField       =   "idtip"
         DataSource      =   "datAsunto"
         Height          =   315
         Index           =   0
         ItemData        =   "frmRefCr.frx":0025
         Left            =   270
         List            =   "frmRefCr.frx":0050
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "n"
         ToolTipText     =   "Valor que desea como título en los renglones"
         Top             =   480
         Width           =   3400
      End
      Begin VB.Label Label1 
         Caption         =   "Título y subtítulo del informe:"
         Height          =   225
         Left            =   240
         TabIndex        =   22
         Top             =   3600
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.Label EtiCombo 
         Caption         =   "Valor que se desea como título de la columna:"
         Height          =   255
         Index           =   2
         Left            =   3915
         TabIndex        =   6
         Top             =   270
         Width           =   3375
      End
      Begin VB.Label EtiList 
         Caption         =   "Variable a calcular:"
         Height          =   240
         Index           =   0
         Left            =   270
         TabIndex        =   5
         Top             =   900
         Width           =   1860
      End
      Begin VB.Label EtiCombo 
         Caption         =   "Valor que se desea como título de la fila:"
         Height          =   255
         Index           =   0
         Left            =   270
         TabIndex        =   4
         Top             =   270
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmRefCruz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ciDelegación = 2
Dim sQueryPrin As String
Dim db As DAO.Database

Private Sub cmdAntSig_Click(Index As Integer)
Dim d As Date, i As Integer

If Not IsDate(txtIni.Text) Then
    txtIni.Text = Format(Now, "dd/mm/yyyy")
End If
If Not IsDate(txtFin.Text) Then
    txtFin.Text = Format(Now, "dd/mm/yyyy")
End If

If Index = 0 Then
    For i = 0 To opcRango.UBound
        If opcRango(i).Value Then Exit For
    Next
    d = CDate(txtIni.Text)
    Select Case i
    Case 0
        d = CDate("01/01/" & (Year(d) - 1))
        txtIni.Text = Format(d, "dd/mm/yyyy")
        d = DateAdd("yyyy", 1, d) - 1
        txtFin.Text = Format(d, "dd/mm/yyyy")
    Case 1
        d = DateAdd("m", -2, d)
        d = d - Day(d) + 1
        txtIni.Text = Format(d, "dd/mm/yyyy")
        d = DateAdd("m", 2, d)
        d = d - Day(d)
        txtFin.Text = Format(d, "dd/mm/yyyy")
    Case 2
        d = CDate(txtFin.Text)
        d = d - Day(d)
        txtFin.Text = Format(d, "dd/mm/yyyy")
        d = d - Day(d) + 1
        txtIni.Text = Format(d, "dd/mm/yyyy")
    Case 3
        d = CDate(txtFin.Text)
        d = d - 7
        txtFin.Text = Format(d, "dd/mm/yyyy")
        d = d - 4
        txtIni.Text = Format(d, "dd/mm/yyyy")
    Case 4
        d = d - 1
        txtIni.Text = Format(d, "dd/mm/yyyy")
        d = CDate(txtFin.Text) - 1
        txtFin.Text = Format(d, "dd/mm/yyyy")
    End Select
Else
    For i = 0 To opcRango.UBound
        If opcRango(i).Value Then Exit For
    Next
    d = CDate(txtIni.Text)
    Select Case i
    Case 0
        d = CDate("01/01/" & (Year(d) + 1))
        txtIni.Text = Format(d, "dd/mm/yyyy")
        d = DateAdd("yyyy", 1, d) - 1
        txtFin.Text = Format(d, "dd/mm/yyyy")
    Case 1
        d = DateAdd("m", 2, d)
        d = d - Day(d) + 1
        txtIni.Text = Format(d, "dd/mm/yyyy")
        d = DateAdd("m", 2, d)
        d = d - Day(d)
        txtFin.Text = Format(d, "dd/mm/yyyy")
    Case 2
        d = CDate(txtFin.Text)
        d = DateAdd("m", 2, d)
        d = d - Day(d)
        txtFin.Text = Format(d, "dd/mm/yyyy")
        d = d - Day(d) + 1
        txtIni.Text = Format(d, "dd/mm/yyyy")
    Case 3
        d = CDate(txtFin.Text)
        d = d + 7
        txtFin.Text = Format(d, "dd/mm/yyyy")
        d = d - 4
        txtIni.Text = Format(d, "dd/mm/yyyy")
    Case 4
        d = d + 1
        txtIni.Text = Format(d, "dd/mm/yyyy")
        d = CDate(txtFin.Text) + 1
        txtFin.Text = Format(d, "dd/mm/yyyy")
    End Select
End If
End Sub

Private Sub cmdBotón_Click(Index As Integer)
Dim s As String, sFrom As String, rs As Recordset, ss As String, s1 As String, s2 As String, s3 As String, yCampos As Integer, y1 As Byte, y2 As Byte, rs1 As Recordset
Dim S_s(3) As String, s_ss(3) As String, s_sfrom(3) As String, b_Otros(1) As Boolean, s_gs(3) As String
Dim adors As New ADODB.Recordset
Dim sW1 As String
Dim qdef As QueryDef
Dim Hoja As Excel.Worksheet
Dim LibroExcel As Excel.Workbook
Dim ApExcel As Excel.Application
Dim dIni As Date, dFin As Date

Me.MousePointer = 11
If IsDate(txtIni.Text) Then
    dIni = CDate(txtIni.Text)
End If
If IsDate(txtFin.Text) Then
    dFin = CDate(txtFin.Text)
End If


If Index = 1 Then 'SALIR
    Unload Me
    Exit Sub
End If

'Verifica si se manda información del Renglón y Columna para el informe de Ref Cruzadas
If ComboVarios(0).ListIndex < 0 Then
    MsgBox "Se requiere especificar la variable del Renglón para el informe", vbInformation + vbOKOnly, "Validación"
    Exit Sub
End If
If ComboVarios(1).ListIndex < 0 Then
    MsgBox "Se requiere especificar la variable de la Columna para el informe", vbInformation + vbOKOnly, "Validación"
    Exit Sub
End If
If Frame4.Visible Then
    If dIni > dFin Then
        MsgBox "El rango de fechas es incorrecto", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
End If
'Asigna nombre del informe (archivo.rpt)
If List1(0).ItemData(List1(0).ListIndex) = 1 Then
    CReport.ReportFileName = gsDirReportes & "\Reporte referencia cruzadas.rpt"
Else
    CReport.ReportFileName = gsDirReportes & "\Reporte referencia cruzadas 2.rpt"
End If
'CReport.ReportFileName = "C:\Users\Administrador\Desktop\SIAM\Reportes\Reporte referencia cruzadas.rpt"
'asigna el valor del parámetro en caso de haber
'If Frame3.Visible Then
'    CReport.ParameterFields(0) = msParámetro & ";" & ComboVarios.ItemData(ComboVarios.ListIndex) & ";true"
'Else
'    CReport.ParameterFields(0) = ""
'End If
'asigna el rango de fechas en caso de haber
If Frame4.Visible Then
    CReport.ParameterFields(1) = "psInicio;" & Format(dIni, "dd/mm/yyyy") & ";true"
    CReport.ParameterFields(2) = "psTermino;" & Format(dFin, "dd/mm/yyyy") & ";true"
Else
    CReport.ParameterFields(1) = ""
    CReport.ParameterFields(2) = ""
End If
CReport.ParameterFields(3) = "piRow;" & ComboVarios(0).ItemData(ComboVarios(0).ListIndex) & ";true"
CReport.ParameterFields(4) = "piCol;" & ComboVarios(1).ItemData(ComboVarios(1).ListIndex) & ";true"

'Asigna la conexión
CReport.Connect = gConSql.ConnectionString '& ";dsn=siam"
CReport.Connect = "FILEDSN=" & App.Path & "\siam.dsn;pwd=siam_desa"


CReport.Action = 1
Me.MousePointer = 0
Exit Sub
ErrorBorrarQuery:
    If Err.Number = 3265 Or Err.Number = 3376 Or InStr(Err.Description, "No se puede quitar") Or InStr(Err.Description, "No se puede drop vista") Or InStr(Err.Description, "table or view does not exist") Then
        Resume Next
    End If
    yErr = MsgBox(Err.Description, vbAbortRetryIgnore, "Error: " + Str(Err.Number))
    If yErr = vbRetry Then
        Resume
    ElseIf yErr = vbIgnore Then
        Resume Next
    End If
    For Y = 0 To y2 + 1
        CReport6.Formulas(Y) = ""
    Next
Me.MousePointer = 0
End Sub

Private Sub ComboVarios_Click(Index As Integer)
Dim Y As Byte
If Index = 0 Then
    cmdBotón(0).Enabled = False
    'cmdBotón(2).Enabled = False
    ComboVarios(1).Clear
    For Y = 0 To ComboVarios(0).ListCount - 1
        If Y <> ComboVarios(0).ListIndex Then
            ComboVarios(1).AddItem ComboVarios(0).List(Y)
            ComboVarios(1).ItemData(ComboVarios(1).ListCount - 1) = ComboVarios(0).ItemData(Y)
        End If
    Next
    txtColumna(0) = ComboVarios(0).Text
    For Y = 0 To 2
        txtRenglón(Y) = ComboVarios(0).Text + Str(Y + 1)
    Next
    ComboVarios(1).ListIndex = -1
Else
    For Y = 1 To 3
        txtColumna(Y) = ComboVarios(1).Text + Str(Y)
    Next
    txtTítulo = "Informe de referencias cruzadas (" + ComboVarios(0).Text + " vs " + ComboVarios(1).Text + ")"
End If
'chkOtros(Index).Visible = InStr("Causas**Institución", ComboVarios(Index).Text) Or ComboVarios(Index).Text Like "Producto Nivel*"
For Y = 0 To 1
    If ComboVarios(Y).ListIndex < 0 Then Exit For
Next
cmdBotón(0).Enabled = Y > 1
'cmdBotón(2).Enabled = Y > 1
End Sub

Private Sub Form_Activate()
Dim i As Integer
List1(0).ListIndex = 0
txtTítulo = "" '"Informe de referencias cruzadas (" + ComboVarios(0).Text + " vs " + ComboVarios(1).Text + ")"
txtSubtítulo = gsTítulo
'If InStr(gsQueryPrin, "select ") = 0 Then
'Else
'    sQueryPrin = gsQueryPrin
'End If
'If gSQLACC = cyAccess Then Set db = OpenDatabase("z:\rpt.mdb", False, False, ";uid=;pwd=837379")
'ActualizaColorFormulario Me
End Sub

Private Sub Form_Load()
'CReport6
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set db = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call QuitaMemoriaForma("frmControlReportes", 1)
End Sub

Private Sub List1_Click(Index As Integer)
'If Index = 0 Then
'    List1(1).Clear
'    List1(1).AddItem "Cuenta"
'    List1(1).ItemData(0) = 0
'    If List1(0).ListIndex > 0 Then
'        List1(1).AddItem "Suma"
''        List1(1).ItemData(1) = 1
 '       List1(1).AddItem "Promedio"
 '       List1(1).ItemData(2) = 2
 '       List1(1).AddItem "Máximo"
 '       List1(1).ItemData(3) = 3
 '       List1(1).AddItem "Mínimo"
 '       List1(1).ItemData(4) = 4
 '   End If
 '   List1(1).ListIndex = 0
 '   txtContenido = List1(1).Text + "(" + IIf(List1(0).ListIndex >= 0, List1(0).Text, "") + ")"
'ElseIf List1(1).ListIndex >= 0 Then
'    txtContenido = List1(1).Text + "(" + IIf(List1(0).ListIndex >= 0, List1(0).Text, "") + ")"
'End If
End Sub


Private Sub opcRango_Click(Index As Integer)
Dim d As Date
Select Case Index
Case 0
    d = CDate("01/01/" & (Year(Date) - 1))
    txtIni.Text = Format(d, "dd/mm/yyyy")
    d = DateAdd("yyyy", 1, d)
    d = d - 1
    txtFin.Text = Format(d, "dd/mm/yyyy")
Case 1
    d = DateAdd("m", IIf(Month(Date) Mod 2 = 0, -3, -2), Date)
    d = d - Day(d) + 1
    txtIni.Text = Format(d, "dd/mm/yyyy")
    d = DateAdd("m", 2, d)
    d = d - Day(d)
    txtFin.Text = Format(d, "dd/mm/yyyy")
Case 2
    d = Date - Day(Date)
    txtFin.Text = Format(d, "dd/mm/yyyy")
    d = d - Day(d) + 1
    txtIni.Text = Format(d, "dd/mm/yyyy")
Case 3
    d = Date - Weekday(Date, vbSaturday)
    txtFin.Text = Format(d, "dd/mm/yyyy")
    d = d - 4
    txtIni.Text = Format(d, "dd/mm/yyyy")
Case 4
    If Not txtIni.Enabled Then
        txtIni.Enabled = True
        txtFin.Enabled = True
    End If
End Select
If Index < 4 And txtIni.Enabled Then
    txtIni.Enabled = False
    txtFin.Enabled = False
End If
'Dim d As Date
'Select Case Index
'Case 0
'    DTP(0).Value = CDate("01/01/" & (Year(Date) - 1))
'    DTP(1).Value = DateAdd("yyyy", 1, DTP(0).Value)
'    DTP(1).Value = DTP(1).Value - 1
'Case 1
'    d = DateAdd("m", IIf(Month(Date) Mod 2 = 0, -3, -2), Date)
'    DTP(0).Value = d - Day(d) + 1
'    d = DateAdd("m", 2, d)
'    DTP(1).Value = d - Day(d)
'Case 2
'    DTP(1).Value = Date - Day(Date)
'    DTP(0).Value = DTP(1).Value - Day(DTP(1).Value) + 1
'Case 3
'    DTP(1).Value = Date - Weekday(Date, vbSaturday)
'    DTP(0).Value = DTP(1).Value - 4
'Case 4
'    If Not DTP(0).Enabled Then
'        DTP(0).Enabled = True
'        DTP(1).Enabled = True
'    End If
'End Select
'If Index < 4 And DTP(0).Enabled Then
'    DTP(0).Enabled = False
'    DTP(1).Enabled = False
'End If
End Sub

'Private Sub UpDown_DownClick()
'Dim d As Date, i As Integer
'For i = 0 To opcRango.UBound
'    If opcRango(i).Value Then Exit For
'Next
'Select Case i
'Case 0
'    DTP(0).Value = CDate("01/01/" & (Year(DTP(0).Value) - 1))
'    DTP(1).Value = DateAdd("yyyy", 1, DTP(0).Value) - 1
'Case 1
'    d = DateAdd("m", -2, DTP(0).Value)
'    DTP(0).Value = d - Day(d) + 1
'    d = DateAdd("m", 2, d)
'    DTP(1).Value = d - Day(d)
'Case 2
'    DTP(1).Value = DTP(1).Value - Day(DTP(1).Value)
'    DTP(0).Value = DTP(1).Value - Day(DTP(1).Value) + 1
'Case 3
'    DTP(1) = DTP(1) - 7
'    DTP(0).Value = DTP(1).Value - 4
'Case 4
'    DTP(0).Value = DTP(0).Value - 1
'    DTP(1).Value = DTP(1).Value - 1
'End Select
'End Sub
'
'Private Sub UpDown_UpClick()
'Dim d As Date, i As Integer
'For i = 0 To opcRango.UBound
'    If opcRango(i).Value Then Exit For
'Next
'Select Case i
'Case 0
'    DTP(0).Value = CDate("01/01/" & (Year(DTP(0).Value) + 1))
'    DTP(1).Value = DateAdd("yyyy", 1, DTP(0).Value) - 1
'Case 1
'    d = DateAdd("m", 2, DTP(0).Value)
'    DTP(0).Value = d - Day(d) + 1
'    d = DateAdd("m", 2, d)
'    DTP(1).Value = d - Day(d)
'Case 2
'    d = DateAdd("m", 2, DTP(1).Value)
'    DTP(1).Value = d - Day(d)
'    DTP(0).Value = DTP(1).Value - Day(DTP(1).Value) + 1
'Case 3
'    DTP(1) = DTP(1) + 7
'    DTP(0).Value = DTP(1).Value - 4
'Case 4
'    DTP(0).Value = DTP(0).Value + 1
'    DTP(1).Value = DTP(1).Value + 1
'End Select
'End Sub

Private Sub txtFin_KeyDown(KeyCode As Integer, Shift As Integer)
If Mid(txtFin.Tag, 1, 1) = "f" And KeyCode = 27 And txtFin.Enabled Then txtFin.Text = ""
End Sub

Private Sub txtFin_KeyPress(KeyAscii As Integer)
Dim i As Long, i1 As Long
If KeyAscii = 13 Then
    i1 = txtFin.TabIndex
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
KeyAscii = TeclaOprimida(txtFin, KeyAscii, txtFin.Tag, False)

End Sub

Private Sub txtFin_LostFocus()
Dim d As Date, Y As Integer, y2 As Integer
', adors As New ADODB.Recordset
If Mid(txtFin.Tag, 1, 1) = "f" Then
    If IsDate(txtFin.Text) Then
        d = CDate(txtFin.Text)
        txtFin.Text = Format(d, gsFormatoFecha)
'        adors.Open "select sysdate from dual", gConSql, adOpenStatic, adLockReadOnly
'        If Index = 2 Then
'            If Int(adors(0)) + 7 - Int(d) < 0 Then
'                Call MsgBox("Fecha no válida. No se permite ingresar fecha mayor a la fecha (" & Format(adors(0) + 7, gsFormatoFecha) & ")", vbOKOnly + vbInformation, "")
'                txtFin.Text = ""
'                Exit Sub
'            End If
'        Else
'            If Int(adors(0)) - Int(d) < 0 Then
'                Call MsgBox("Fecha no válida. No se permite ingresar fecha mayor a la fecha actual (" & Format(adors(0), gsFormatoFecha) & ")", vbOKOnly + vbInformation, "")
'                txtFin.Text = ""
'                Exit Sub
'            End If
'        End If
    Else
        If Len(txtFin.Text) > 0 Then
            Call MsgBox("Fecha no válida. Verificar", vbOKOnly + vbInformation, "")
            txtFin.Text = ""
        End If
    End If
End If
End Sub

Private Sub txtIni_KeyDown(KeyCode As Integer, Shift As Integer)
If Mid(txtIni.Tag, 1, 1) = "f" And KeyCode = 27 And txtIni.Enabled Then txtIni.Text = ""
End Sub

Private Sub txtIni_KeyPress(KeyAscii As Integer)
Dim i As Long, i1 As Long
If KeyAscii = 13 Then
    i1 = txtIni.TabIndex
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
KeyAscii = TeclaOprimida(txtIni, KeyAscii, txtIni.Tag, False)

End Sub

Private Sub txtIni_LostFocus()
Dim d As Date, Y As Integer, y2 As Integer
If Mid(txtIni.Tag, 1, 1) = "f" Then
    If IsDate(txtIni.Text) Then
        d = CDate(txtIni.Text)
        txtIni.Text = Format(d, gsFormatoFecha)
'        adors.Open "select sysdate from dual", gConSql, adOpenStatic, adLockReadOnly
'        If Index = 2 Then
'            If Int(adors(0)) + 7 - Int(d) < 0 Then
'                Call MsgBox("Fecha no válida. No se permite ingresar fecha mayor a la fecha (" & Format(adors(0) + 7, gsFormatoFecha) & ")", vbOKOnly + vbInformation, "")
'                txtIni.Text = ""
'                Exit Sub
'            End If
'        Else
'            If Int(adors(0)) - Int(d) < 0 Then
'                Call MsgBox("Fecha no válida. No se permite ingresar fecha mayor a la fecha actual (" & Format(adors(0), gsFormatoFecha) & ")", vbOKOnly + vbInformation, "")
'                txtIni.Text = ""
'                Exit Sub
'            End If
'        End If
    Else
        If Len(txtIni.Text) > 0 Then
            Call MsgBox("Fecha no válida. Verificar", vbOKOnly + vbInformation, "")
            txtIni.Text = ""
        End If
    End If
End If
End Sub

