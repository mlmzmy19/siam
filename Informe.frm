VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form Informe 
   Caption         =   "Parámetros para el informe"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6510
   LinkTopic       =   "Form2"
   ScaleHeight     =   5445
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5400
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6465
      Begin Crystal.CrystalReport CReport 
         Left            =   5640
         Top             =   4680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CheckBox chkExcel 
         Height          =   195
         Left            =   225
         TabIndex        =   17
         Top             =   4830
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Frame Frame5 
         Caption         =   "Grupo"
         Height          =   690
         Left            =   120
         TabIndex        =   16
         Top             =   1065
         Width           =   6225
         Begin VB.ComboBox ComboGrupo 
            Height          =   315
            Left            =   135
            TabIndex        =   1
            Top             =   270
            Width           =   5955
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Informe"
         Height          =   690
         Left            =   120
         TabIndex        =   15
         Top             =   255
         Width           =   6225
         Begin VB.ComboBox comboInforme 
            Height          =   315
            Left            =   72
            TabIndex        =   0
            Top             =   270
            Width           =   6000
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Responsable"
         Height          =   690
         Left            =   120
         TabIndex        =   14
         Top             =   1860
         Width           =   6225
         Begin VB.ComboBox ComboVarios 
            Height          =   315
            Left            =   135
            TabIndex        =   2
            Top             =   240
            Width           =   5955
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1845
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   6225
         Begin VB.CommandButton cmdAntSig 
            Caption         =   "Ant"
            Height          =   372
            Index           =   0
            Left            =   4464
            TabIndex        =   21
            Top             =   1290
            Width           =   624
         End
         Begin VB.CommandButton cmdAntSig 
            Caption         =   "Sig"
            Height          =   372
            Index           =   1
            Left            =   4464
            TabIndex        =   20
            Top             =   900
            Width           =   624
         End
         Begin VB.TextBox txtIni 
            Height          =   300
            Left            =   1044
            TabIndex        =   19
            Tag             =   "f"
            Top             =   930
            Width           =   3252
         End
         Begin VB.TextBox txtFin 
            Height          =   300
            Left            =   1044
            TabIndex        =   18
            Tag             =   "f"
            Top             =   1335
            Width           =   3252
         End
         Begin VB.OptionButton opcRango 
            Caption         =   "Otro"
            Height          =   285
            Index           =   4
            Left            =   5490
            TabIndex        =   7
            Top             =   315
            Width           =   600
         End
         Begin VB.OptionButton opcRango 
            Caption         =   "Semanal"
            Height          =   285
            Index           =   3
            Left            =   4050
            TabIndex        =   6
            Top             =   315
            Width           =   1095
         End
         Begin VB.OptionButton opcRango 
            Caption         =   "Mensual"
            Height          =   285
            Index           =   2
            Left            =   2700
            TabIndex        =   5
            Top             =   315
            Width           =   1005
         End
         Begin VB.OptionButton opcRango 
            Caption         =   "Bimestral"
            Height          =   285
            Index           =   1
            Left            =   1305
            TabIndex        =   4
            Top             =   315
            Width           =   1050
         End
         Begin VB.OptionButton opcRango 
            Caption         =   "Anual"
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   3
            Top             =   315
            Width           =   780
         End
         Begin VB.Label Label3 
            Caption         =   "Al:"
            Height          =   240
            Index           =   1
            Left            =   360
            TabIndex        =   13
            Top             =   1365
            Width           =   285
         End
         Begin VB.Label Label3 
            Caption         =   "Del:"
            Height          =   240
            Index           =   0
            Left            =   270
            TabIndex        =   11
            Top             =   990
            Width           =   375
         End
      End
      Begin VB.CommandButton cmbBotón 
         Caption         =   "&Aceptar"
         Height          =   315
         Index           =   0
         Left            =   1950
         TabIndex        =   10
         Top             =   4785
         Width           =   1065
      End
      Begin VB.CommandButton cmbBotón 
         Caption         =   "&Cancelar"
         Height          =   315
         Index           =   1
         Left            =   3795
         TabIndex        =   12
         Top             =   4800
         Width           =   1065
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   540
         Picture         =   "Informe.frx":0000
         Top             =   4785
         Visible         =   0   'False
         Width           =   240
      End
   End
End
Attribute VB_Name = "Informe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msArchivo As String
Dim msParámetro As String
Dim iTipoRep As Integer

Private Sub cmbBotón_Click(Index As Integer)
Dim rs As DAO.Recordset
Dim i As Integer, Y As Integer
Dim yErr As Byte
Dim s As String, s2 As String
Dim adors As ADODB.Recordset
Dim dIni As Date, dFin As Date
'Dim adors As New ADODB.Recordset
On Error GoTo salir:
If Index = 1 Then
    Unload Me
    Exit Sub
End If
If comboInforme.ListIndex < 0 Then
    MsgBox "Necesita especificar el informe a imprimir", vbOKOnly + vbInformation, ""
    comboInforme.SetFocus
    Exit Sub
End If
If Frame3.Visible Then
    If ComboVarios.ListCount > 0 And ComboVarios.ListIndex < 0 Then
        MsgBox "Falta especificar " & Frame3.Caption, vbOKOnly + vbInformation, ""
        Exit Sub
    End If
    If ComboGrupo.ListCount > 0 And ComboGrupo.ListIndex < 0 Then
        MsgBox "Falta especificar el Grupo de Usuario propietario de la Información", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
    If ComboVarios.ListCount > 0 Then 'Verifica si es requerido el segundo reporte con el gráfico en detalle
        If ComboVarios.ItemData(ComboVarios.ListIndex) > 0 Then
            's2 = "_2" 'Buscara el segundo reporte que termina con "_2"
        'Else
            s2 = ""
        End If
    End If
End If
If Frame4.Visible Then
    If IsDate(txtIni.Text) Then
        dIni = CDate(txtIni.Text)
    End If
    If IsDate(txtFin.Text) Then
        dFin = CDate(txtFin.Text)
    End If
    If dIni > dFin Then
        MsgBox "El rango de fechas es incorrecto", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
End If
'Asigna nombre del informe (archivo.rpt)
Set adors = New ADODB.Recordset
If adors.State > 0 Then adors.Close
adors.Open "select url from informes where id=" & comboInforme.ItemData(comboInforme.ListIndex), gConSql, adOpenStatic, adLockReadOnly
If Len(adors(0)) > 0 Then
    s = adors(0)
    s = Replace(s, "Valor01", "" & ComboGrupo.ItemData(ComboGrupo.ListIndex))
    s = Replace(s, "Valor02", "" & ComboVarios.ItemData(ComboVarios.ListIndex))
    s = Replace(s, "Valor03", "" & Format(dIni, "dd/mm/yyyy"))
    s = Replace(s, "Valor04", "" & Format(dFin, "dd/mm/yyyy"))
    gsWWW = s
    With Browser
        If chkExcel.Value Then
            gsWWW = Replace(gsWWW, ".prpt/", "_excel.prpt/")
            .cmd1.Visible = True
            .cmd2.Visible = True
            .Height = 1900
            .Width = 5000
        Else
            .cmd1.Visible = False
            .cmd2.Visible = False
            .Height = 12000
            .Width = 12000
        End If
        .yÚnicavez = 0
        .Caption = "Informe de Sanciones Impuestas por Clase e Institución Homologadas y Causa de la Sanción"
        .Show vbModal
    End With
    Exit Sub
End If
CReport.ReportFileName = gsDirReportes & "\" & msArchivo & s2 & ".rpt"
'asigna el valor del parámetro en caso de haber
If Frame3.Visible Then
    CReport.ParameterFields(0) = msParámetro & ";" & ComboVarios.ItemData(ComboVarios.ListIndex) & ";true"
Else
    CReport.ParameterFields(0) = ""
End If
If Frame3.Visible Then
    CReport.ParameterFields(1) = "pigpo;" & ComboGrupo.ItemData(ComboGrupo.ListIndex) & ";true"
Else
    CReport.ParameterFields(1) = ""
End If
'asigna el rango de fechas en caso de haber
If Frame4.Visible Then
    CReport.ParameterFields(2) = "psInicio;" & Format(dIni, "dd/mm/yyyy") & ";true"
    CReport.ParameterFields(3) = "psTermino;" & Format(dFin, "dd/mm/yyyy") & ";true"
Else
    CReport.ParameterFields(2) = ""
    CReport.ParameterFields(3) = ""
End If
'Asigna la conexión
CReport.Connect = gConSql.ConnectionString '& ";dsn=siam"
CReport.Connect = "FILEDSN=" & App.Path & "\siam.dsn;pwd=siam_desa"

'Abre el informe
CReport.Action = 1
Exit Sub
salir:
If InStr(Err.Description, "nvalid file name") > 0 Then
    MsgBox "No se localizó el Archivo " & gsDirReportes & "\" & msArchivo & ".rpt" & ". Favor de verificar la omisión de este archivo en el sistema o marcar la extensión 6032 Miguel Martínez Monroy", vbOKOnly + vbInformation, "Validación"
    Exit Sub
Else
    yErr = MsgBox("Error no esperado: " & Err.Description, vbQuestion + vbAbortRetryIgnore, "")
    If yErr = vbRetry Then
        Resume
    ElseIf yErr = vbIgnore Then
        Resume Next
    End If
End If
End Sub


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

Private Sub Combo1_Change()

End Sub

Private Sub comboInforme_Click()
Dim adors As New ADODB.Recordset, s As String, i As Integer
If comboInforme.ListIndex >= 0 Then
    If adors.State Then adors.Close
    adors.Open "select i.*,ip.descripción as descpar from informes i, informeparámetros ip where i.id=" & comboInforme.ItemData(comboInforme.ListIndex) & " and i.idpar=ip.id(+)", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        msArchivo = adors!archivo
        If adors!Excel > 0 Then
            If Not chkExcel.Visible Then
                chkExcel.Visible = True
                Image1.Visible = True
            End If
        Else
            If chkExcel.Visible Then
                chkExcel.Visible = False
                Image1.Visible = False
            End If
        End If
        If adors!idpar > 0 Then
            If adors!tiempo = 0 Then
                Frame4.Visible = False
            Else
                Frame4.Visible = True
                For i = 0 To opcRango.UBound
                    If opcRango(i).Value Then Exit For
                Next
                If i > opcRango.UBound Then
                    opcRango(i - 2).Value = True
                End If
            End If
            msParámetro = IIf(IsNull(adors!nombrepar), "psParámetro", adors!nombrepar)
            i = adors!idpar
            Frame3.Visible = True
            If Not IsNull(adors!descpar) Then
                Frame3.Caption = adors!descpar
            End If
            ComboVarios.Clear
            ComboVarios.AddItem "*** TODOS ***"
            ComboVarios.ItemData(ComboVarios.NewIndex) = 0
            If adors.State Then adors.Close
            adors.Open "select * from informeparámetros where id=" & i, gConSql, adOpenStatic, adLockReadOnly
            If Not adors.EOF Then
                Call LlenaCombo(ComboVarios, adors!Script, "", True, True)
            End If
        Else
            Frame3.Visible = False
            If adors!tiempo = 0 Then
                Frame4.Visible = False
            Else
                Frame4.Visible = True
                For i = 0 To opcRango.UBound
                    If opcRango(i).Value Then Exit For
                Next
                If i > opcRango.UBound Then
                    opcRango(i - 2).Value = True
                End If
            End If
        End If
    End If
End If
End Sub
Private Sub DTP_Change(Index As Integer)
Dim d As Date
End Sub


Private Sub Form_Load()
iTipoRep = gi
Call LlenaCombo(comboInforme, "select id,descripción from informes where baja is null and tipo= " & iTipoRep & " order by 2", "", True)
ComboGrupo.AddItem "*** TODOS ***"
ComboGrupo.ItemData(ComboGrupo.NewIndex) = 0
Call LlenaCombo(ComboGrupo, "select id,descripción from grupousuarios where baja=0 order by 2", "", True, True)
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
'
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

