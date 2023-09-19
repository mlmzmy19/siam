VERSION 5.00
Begin VB.Form ConsultaTabular_Excel 
   Caption         =   "Genera Consulta Tabular en Excel"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frame2 
      Caption         =   "Status"
      Height          =   600
      Left            =   90
      TabIndex        =   16
      Top             =   4635
      Width           =   6225
      Begin VB.TextBox txtStatus 
         Height          =   285
         Left            =   90
         TabIndex        =   27
         Top             =   225
         Width           =   6045
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00B4E2C9&
      Caption         =   "Seleccione la consulta Tabular que desea migrar a Excel:"
      Height          =   690
      Left            =   90
      TabIndex        =   13
      Top             =   25
      Width           =   6225
      Begin VB.ComboBox comboProcTab 
         Height          =   315
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Width           =   5955
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00B4E2C9&
      Height          =   1675
      Left            =   90
      TabIndex        =   6
      Top             =   1400
      Width           =   6225
      Begin VB.ComboBox ComboIF 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   5955
      End
      Begin VB.OptionButton opcIF 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vigentes"
         Height          =   285
         Index           =   0
         Left            =   600
         TabIndex        =   8
         Top             =   800
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton opcIF 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Todas"
         Height          =   285
         Index           =   1
         Left            =   1605
         TabIndex        =   9
         Top             =   800
         Width           =   1005
      End
      Begin VB.CommandButton cmdBusIF 
         Caption         =   "Sig"
         Height          =   330
         Left            =   5460
         TabIndex        =   11
         Top             =   800
         Width           =   510
      End
      Begin VB.TextBox txtBuscarIF 
         Height          =   285
         Left            =   3405
         TabIndex        =   10
         Top             =   800
         Width           =   2010
      End
      Begin VB.ComboBox ComboClase 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   400
         Width           =   5955
      End
      Begin VB.Label Label2 
         Caption         =   "I.F:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   495
      End
      Begin VB.Label etiCombo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Busca IF:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   2640
         TabIndex        =   21
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label1 
         BackColor       =   &H00B4E2C9&
         Caption         =   "Sector Financiero:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   160
         Width           =   2055
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00B4E2C9&
      Caption         =   "Grupo"
      Height          =   690
      Left            =   90
      TabIndex        =   5
      Top             =   700
      Width           =   6225
      Begin VB.ComboBox ComboGrupo 
         Height          =   315
         Left            =   135
         TabIndex        =   2
         Top             =   270
         Width           =   5955
      End
   End
   Begin VB.CommandButton cmdver 
      Caption         =   "&Procesa &Reporte"
      Height          =   600
      Left            =   2400
      Picture         =   "ConsultaTabular_Excel.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5280
      Width           =   1395
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00B4E2C9&
      Height          =   1590
      Left            =   90
      TabIndex        =   0
      Top             =   3060
      Width           =   6225
      Begin VB.TextBox txtFin 
         Height          =   300
         Left            =   1080
         TabIndex        =   25
         Tag             =   "f"
         Top             =   1095
         Width           =   3252
      End
      Begin VB.TextBox txtIni 
         Height          =   300
         Left            =   1080
         TabIndex        =   22
         Tag             =   "f"
         Top             =   705
         Width           =   3252
      End
      Begin VB.CommandButton cmdAntSig 
         Caption         =   "Sig"
         Height          =   372
         Index           =   1
         Left            =   4500
         TabIndex        =   24
         Top             =   660
         Width           =   624
      End
      Begin VB.CommandButton cmdAntSig 
         Caption         =   "Ant"
         Height          =   372
         Index           =   0
         Left            =   4500
         TabIndex        =   26
         Top             =   1065
         Width           =   624
      End
      Begin VB.OptionButton opcRango 
         BackColor       =   &H00B4E2C9&
         Caption         =   "Anual"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   315
         Width           =   780
      End
      Begin VB.OptionButton opcRango 
         BackColor       =   &H00B4E2C9&
         Caption         =   "Bimestral"
         Height          =   285
         Index           =   1
         Left            =   1305
         TabIndex        =   15
         Top             =   315
         Width           =   1050
      End
      Begin VB.OptionButton opcRango 
         BackColor       =   &H00B4E2C9&
         Caption         =   "Mensual"
         Height          =   285
         Index           =   2
         Left            =   2700
         TabIndex        =   17
         Top             =   315
         Width           =   1005
      End
      Begin VB.OptionButton opcRango 
         BackColor       =   &H00B4E2C9&
         Caption         =   "Semanal"
         Height          =   285
         Index           =   3
         Left            =   4050
         TabIndex        =   18
         Top             =   315
         Width           =   1095
      End
      Begin VB.OptionButton opcRango 
         BackColor       =   &H00B4E2C9&
         Caption         =   "Otro"
         Height          =   285
         Index           =   4
         Left            =   5490
         TabIndex        =   20
         Top             =   315
         Width           =   600
      End
      Begin VB.Label Label3 
         BackColor       =   &H00B4E2C9&
         Caption         =   "Del:"
         Height          =   240
         Index           =   0
         Left            =   570
         TabIndex        =   4
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label3 
         BackColor       =   &H00B4E2C9&
         Caption         =   "Al:"
         Height          =   240
         Index           =   1
         Left            =   645
         TabIndex        =   3
         Top             =   1215
         Width           =   285
      End
   End
End
Attribute VB_Name = "ConsultaTabular_Excel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sProcs As String
Dim color1 As Long
Dim color2 As Long


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

Private Sub cmdBusIF_Click()
Dim i As Integer, iPos As Integer
If Len(txtBuscarIF.Text) > 0 And ComboIF.ListCount > 0 Then
    iPos = ComboIF.ListIndex
    If iPos = ComboIF.ListCount - 1 Then
        i = -1
    Else
        i = BuscaCombo(ComboIF, txtBuscarIF.Text, 0, True, 0, iPos + 1)
    End If
    If i >= 0 Then
        ComboIF.ListIndex = i
    ElseIf iPos >= 0 Then
        ComboIF.ListIndex = -1
    End If
End If
End Sub

Private Sub cmdver_Click()
Dim Hoja As Excel.Worksheet
Dim LibroExcel As Excel.Workbook
Dim ApExcel As Excel.Application
Dim l As Long, i As Integer, yErr As Byte, iGpo As Integer, iCla As Integer, iIns As Integer
Dim s As String, df As Date, iRep As Integer
Dim adors As New ADODB.Recordset
'Dim adors As ADODB.Recordset
On Error GoTo ErrArchivo:
If comboProcTab.ListIndex < 0 Then
    MsgBox "debe seleccionar la consulta a procesar...", vbOKOnly + vbInformation, "Validación"
    Exit Sub
End If
iRep = comboProcTab.ItemData(comboProcTab.ListIndex)
If ComboGrupo.ListIndex < 0 Then
    MsgBox "debe seleccionar el grupo ...", vbOKOnly + vbInformation, "Validación"
    Exit Sub
End If
iGpo = ComboGrupo.ItemData(ComboGrupo.ListIndex)
If ComboClase.ListIndex < 0 Then
    MsgBox "debe seleccionar la Clase de la Institución Financiera...", vbOKOnly + vbInformation, "Validación"
    Exit Sub
End If
iCla = ComboClase.ItemData(ComboClase.ListIndex)
If ComboIF.ListIndex < 0 Then
    MsgBox "debe seleccionar la Institución Financiera...", vbOKOnly + vbInformation, "Validación"
    Exit Sub
End If
iIns = ComboIF.ItemData(ComboIF.ListIndex)
s = ObtieneSubCad(sProcs, comboProcTab.ListIndex + 1)
If iRep <> 27 Then
    s = Replace(s, "PARAM01", "'" & Format(CDate(txtIni.Text), "dd/mm/yyyy") & "'")
    If comboProcTab.ItemData(comboProcTab.ListIndex) = 18 Then
        s = Replace(s, "PARAM02", "'" & Format(CDate(txtFin.Text), "dd/mm/yyyy") & "',1")
    Else
        s = Replace(s, "PARAM02", "'" & Format(CDate(txtFin.Text), "dd/mm/yyyy") & "'")
    End If
End If
s = Replace(s, "(0,0,0,", "(0,0," & iIns & ",")
s = Replace(s, "(0,0,", "(0," & iCla & ",")
s = Replace(s, "(0,", "(" & iGpo & ",")
'Obtiene fecha de emisión del reporte
adors.Open "select sysdate from dual", gConSql, adOpenForwardOnly, adLockReadOnly
df = adors(0)
'MsgBox ("Proceso con parámetros: " + s)
If MsgBox("Está seguro de " & IIf(iRep = 27, "Programar reportes en la noche", " emitir reporte " & comboProcTab.Text & Chr(13) & Chr(10) + "Parámetros: " & s), vbYesNo, "Confirmación") = vbNo Then
    Exit Sub
End If
adors.Close
Me.MousePointer = 11
adors.Open "{call " & s & "}", gConSql, adOpenForwardOnly, adLockReadOnly
If Not adors.EOF Then
    If iRep = 27 Then 'Solo muestra resultado de la programación
        Me.MousePointer = 0
        If adors(0) > 0 Then
            Call MsgBox("Programación correcta: " + adors(1), vbOKOnly, "")
        Else
            Call MsgBox("Programación no correcta: " + adors(1), vbOKOnly, "")
        End If
        Exit Sub
    End If
    Frame1.BackColor = color1
    Frame3.BackColor = color1
    Frame4.BackColor = color1
    Frame5.BackColor = color1
    opcRango(0).BackColor = color1
    opcRango(1).BackColor = color1
    opcRango(2).BackColor = color1
    opcRango(3).BackColor = color1
    opcRango(4).BackColor = color1
    Frame2.BackColor = color2
    Label3(0).BackColor = color2
    Label3(1).BackColor = color2
    Set ApExcel = CreateObject("Excel.Application") 'Método CreateObject y Application
    Set LibroExcel = ApExcel.Workbooks.Add   'con .Add añadimso Libros de trabajo de la aplicacion
    Set Hoja = LibroExcel.Worksheets(1)      'referenciado la primera hoja del libro de trabajo
    Hoja.Activate    'Activando la hoja
    ApExcel.Visible = True  'Hacemos visible la aplicación
    Hoja.Cells(1, 1).Value = "Consulta: " & comboProcTab.Text
    Hoja.Cells(1, 4).Value = "Emisión: " & Format(df, "dd/mm/yyyy hh:mm:ss")
    Hoja.Cells(2, 1).Value = "Grupo: " & ComboGrupo.Text
    Hoja.Cells(3, 1).Value = "Sector: " & ComboClase.Text
    Hoja.Cells(4, 1).Value = "Institución: " & ComboIF.Text
    Hoja.Cells(5, 1).Value = "Del: " & txtIni.Text
    Hoja.Cells(5, 4).Value = "Al: " & txtFin.Text
    For Y = 0 To adors.Fields.Count - 1
        Hoja.Cells(6, Y + 1).Value = adors.Fields(Y).Name
    Next
    bExportar = True
    i = 7
    Do While Not adors.EOF
        For Y = 0 To adors.Fields.Count - 1
            Hoja.Cells(i, Y + 1).Value = Replace(Replace(IIf(IsNull(adors(Y)), "", adors(Y)), Chr(13), " "), Chr(10), " ")
        Next
        txtStatus.Text = "Registros transferidos: " & adors.AbsolutePosition & " / " & adors.RecordCount
        txtStatus.Refresh
        adors.MoveNext
        i = i + 1
    Loop
    txtStatus.Text = "Proceso terminado, registros transferidos: " & adors.RecordCount
    bExportar = False
    Frame1.BackColor = color2
    Frame3.BackColor = color2
    Frame4.BackColor = color2
    Frame5.BackColor = color2
    opcRango(0).BackColor = color2
    opcRango(1).BackColor = color2
    opcRango(2).BackColor = color2
    opcRango(3).BackColor = color2
    opcRango(4).BackColor = color2
    Label3(0).BackColor = color2
    Label3(1).BackColor = color2
    Frame2.BackColor = color1
Else
    MsgBox "No se encontraron registros...", vbOKOnly + vbInformation, "Información"
End If
Me.MousePointer = 0
Exit Sub
ErrArchivo:
bExportar = False
If Err.Number = 384 Then
    Resume Next
End If
Me.MousePointer = 0
yErr = MsgBox("Error: " + Err.Description, vbAbortRetryIgnore + vbCritical, "Error no esperado (" + Str(Err.Number) + ")")
If yErr = vbCancel Then
    Exit Sub
ElseIf yErr = vbRetry Then
    Resume
ElseIf yErr = vbIgnore Then
    Resume Next
End If
End Sub


Private Sub ComboClase_Change()
If Len(txtStatus.Text) > 0 Then
    txtStatus.Text = ""
End If
End Sub

Private Sub ComboClase_Click()
'MsgBox "entra sub ComboClase_Click"
If ComboClase.ListIndex < 0 Then
    Exit Sub
End If
ActualizaComboIF
'MsgBox "Evalua ComboClase"
End Sub

Sub ActualizaComboIF()
ComboIF.Clear
If comboProcTab.ListIndex >= 0 Then
    If comboProcTab.ItemData(comboProcTab.ListIndex) <> 27 Then
        ComboIF.AddItem "*** Todas las IF ***", 0
    End If
End If
'MsgBox "Prepara ComboIF"
If opcIF(0).Value Then
    LlenaCombo ComboIF, "select i.id,i.descripción from relaciónclaseinstitución rci, instituciones i where rci.idcla=" & ComboClase.ItemData(ComboClase.ListIndex) & " and rci.idins=i.id and i.status=1 order by 2", "", True, True
    
Else
    LlenaCombo ComboIF, "select i.id,i.descripción from relaciónclaseinstitución rci, instituciones i where rci.idcla=" & ComboClase.ItemData(ComboClase.ListIndex) & " and rci.idins=i.id order by 2", "", True, True
End If
'MsgBox "Llena combo con opciones ComboIF"
If ComboIF.ListCount = 1 Then
    ComboIF.ListIndex = 0
End If
End Sub

Private Sub ComboGrupo_Change()
If Len(txtStatus.Text) > 0 Then
    txtStatus.Text = ""
End If
End Sub

Private Sub comboProcTab_Change()
If Len(txtStatus.Text) > 0 Then
    txtStatus.Text = ""
End If
End Sub

Private Sub comboProcTab_Click()
If comboProcTab.ListIndex >= 0 Then
    If comboProcTab.ItemData(comboProcTab.ListIndex) = 27 Then
        If Frame4.Visible Then
            Frame4.Visible = False
        End If
    Else
        If Not Frame4.Visible Then
            Frame4.Visible = True
        End If
    
    End If
End If

End Sub

Private Sub Form_Load()
Dim adors As New ADODB.Recordset, i As Integer, s As String
adors.Open "select id,descripción,proc_tab from informes2 where excel=1 order by 2", gConSql, adOpenStatic, adLockReadOnly
comboProcTab.Clear
color1 = &H8000000F
color2 = &HB4E2C9
sProcs = ""
Do While Not adors.EOF
    Call comboProcTab.AddItem(adors(1), i)
    comboProcTab.ItemData(i) = adors(0)
    sProcs = sProcs & adors(2) & "|"
    adors.MoveNext
    i = i + 1
Loop
'adors.Open "call {paq_informes.p_grupos(" & giUsuario & "}", gConSql, adOpenForwardOnly, adLockReadOnly
Call LlenaCombo(ComboGrupo, "{call paq_informes.p_grupos(" & giUsuario & ")}", "", True)
Call LlenaCombo(ComboClase, "{call paq_informes.p_clases(" & giUsuario & ")}", "", True)
End Sub

Function ObtieneSubCad(ByVal ss As String, iPos As Integer)
Dim i As Integer, s As String, n As Integer
n = InStr(ss, "|")
Do While n > 0
    n = InStr(ss, "|")
    s = Mid(ss, 1, n - 1)
    ss = Mid(ss, n + 1)
    i = i + 1
    If i = iPos Then
        ObtieneSubCad = s
        Exit Function
    End If
Loop
End Function

Private Sub opcIF_Click(Index As Integer)
ActualizaComboIF
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
End Sub

Private Sub txtFin_Change()
If Len(txtStatus.Text) > 0 Then
    txtStatus.Text = ""
End If
End Sub

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

Private Sub txtIni_Change()
If Len(txtStatus.Text) > 0 Then
    txtStatus.Text = ""
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
