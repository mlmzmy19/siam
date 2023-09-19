VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Condonacion 
   Appearance      =   0  'Flat
   Caption         =   "Información de la Condonación"
   ClientHeight    =   4500
   ClientLeft      =   2025
   ClientTop       =   1995
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   2376
      TabIndex        =   8
      Top             =   3744
      Width           =   4755
      Begin VB.CommandButton cmdBotón 
         Caption         =   "&Aceptar"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   1
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton cmdBotón 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   2700
         TabIndex        =   2
         Top             =   180
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3690
      Left            =   90
      TabIndex        =   10
      Top             =   45
      Width           =   9360
      Begin MSComctlLib.ListView ListView1 
         Height          =   1188
         Left            =   144
         TabIndex        =   3
         Top             =   1008
         Width           =   9156
         _ExtentX        =   16140
         _ExtentY        =   2090
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Causa"
            Object.Width           =   5821
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Sub"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Procede"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Porcentaje"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Monto Sanción"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Monto Condonado"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Monto a Pagar"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.TextBox txtCampo 
         DataSource      =   "datAsunto"
         Height          =   285
         Index           =   0
         Left            =   90
         MaxLength       =   70
         TabIndex        =   9
         Tag             =   "c"
         Top             =   450
         Width           =   4395
      End
      Begin VB.TextBox txtCampo 
         DataSource      =   "datAsunto"
         Height          =   285
         Index           =   1
         Left            =   6390
         MaxLength       =   30
         TabIndex        =   0
         Tag             =   "f"
         Top             =   495
         Width           =   2625
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ley Causa Selecionada"
         Height          =   1140
         Left            =   135
         TabIndex        =   14
         Top             =   2340
         Width           =   9105
         Begin VB.TextBox txtCampo 
            BackColor       =   &H00E0E0E0&
            DataSource      =   "datAsunto"
            Height          =   285
            Index           =   3
            Left            =   4248
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   20
            Tag             =   "n"
            Top             =   504
            Width           =   1476
         End
         Begin VB.TextBox txtCampo 
            Height          =   288
            Index           =   2
            Left            =   2736
            MaxLength       =   5
            TabIndex        =   19
            Tag             =   "n"
            Top             =   504
            Width           =   444
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H00E0E0E0&
            DataSource      =   "datAsunto"
            Height          =   285
            Index           =   4
            Left            =   270
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   4
            Tag             =   "c"
            Top             =   540
            Width           =   390
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H00E0E0E0&
            DataSource      =   "datAsunto"
            Height          =   285
            Index           =   6
            Left            =   7245
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   7
            Tag             =   "n"
            Top             =   495
            Width           =   1380
         End
         Begin VB.TextBox txtCampo 
            BackColor       =   &H00E0E0E0&
            DataSource      =   "datAsunto"
            Height          =   285
            Index           =   5
            Left            =   5760
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   6
            Tag             =   "n"
            Top             =   504
            Width           =   1476
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Procede"
            Enabled         =   0   'False
            Height          =   285
            Left            =   828
            TabIndex        =   5
            Top             =   540
            Value           =   1  'Checked
            Width           =   1008
         End
         Begin VB.Label EtiTexto 
            Caption         =   "Monto Sanción:"
            Height          =   252
            Index           =   6
            Left            =   4248
            TabIndex        =   21
            Top             =   252
            Width           =   1212
         End
         Begin VB.Label EtiTexto 
            Caption         =   "SubInd:"
            Height          =   252
            Index           =   5
            Left            =   252
            TabIndex        =   18
            Top             =   252
            Width           =   528
         End
         Begin VB.Label EtiTexto 
            Caption         =   "Importe a Pagar:"
            Height          =   255
            Index           =   3
            Left            =   7260
            TabIndex        =   17
            Top             =   225
            Width           =   1425
         End
         Begin VB.Label EtiTexto 
            Caption         =   "Monto Condonado:"
            Height          =   252
            Index           =   2
            Left            =   5760
            TabIndex        =   16
            Top             =   252
            Width           =   1428
         End
         Begin VB.Label Label1 
            Caption         =   "Porcentaje Condonado:"
            Height          =   240
            Left            =   2136
            TabIndex        =   15
            Top             =   252
            Width           =   1992
         End
      End
      Begin VB.Label EtiTexto 
         Caption         =   "Seleccione y capture la información de cada causa asociada a la sanción:"
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   4
         Left            =   90
         TabIndex        =   13
         Top             =   810
         Width           =   5835
      End
      Begin VB.Label EtiTexto 
         Caption         =   "No. Oficio de Condonación:"
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   225
         Width           =   2385
      End
      Begin VB.Label EtiTexto 
         Caption         =   "Fecha:"
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   1
         Left            =   6480
         TabIndex        =   11
         Top             =   270
         Width           =   1740
      End
   End
End
Attribute VB_Name = "Condonacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const cLímInf = 500
Const cLímSup = 2000
Const cnGrisclaro = &HE0E0E0
Dim msSanción As String
Public mdFechaOficio 'Fecha del oficio de sanción (fecha de la actividad que invoca este formulario)
Public myAcción 'Indica que tipo de acción se realizará 0:consulta; 1:Alta; 2: Modificación
Dim bAceptar As Boolean 'Variable lógica que indica si fueron aceptadados los cambios
Dim adors As New ADODB.Recordset
Dim mlAnálisis As Long, mlSeguimiento As Long
Dim sCadena As String
Dim bValidachk1 As Boolean 'Bandera que indica si debe validar o no el check1
Dim nCausas As Integer 'Número de causas asociadas al oficio de sanción

Sub ActualizaTag()
Dim s As String
s = Mid(ListView1.ListItems(ListView1.SelectedItem.Index).Tag, 1, InStr(ListView1.ListItems(ListView1.SelectedItem.Index).Tag, "|"))
s = s & txtcampo(4).Text & "|"
ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(1).Text = txtcampo(4).Text
s = s & IIf(Check1.Value <= 1, Check1.Value, -1) & "|"
If Check1.Value = 1 Then
    ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(2).Text = "Si"
ElseIf Check1.Value = 0 Then
    ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(2).Text = "No"
Else
    ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(2).Text = "No Defenido"
End If
s = s & Replace(Replace(txtcampo(2).Text, ",", ""), "$", "") & "|"
s = s & Replace(Replace(txtcampo(3).Text, ",", ""), "$", "") & "|"
s = s & Replace(Replace(txtcampo(5).Text, ",", ""), "$", "") & "|"
s = s & Replace(Replace(txtcampo(6).Text, ",", ""), "$", "") & "|"
ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(3).Text = txtcampo(2).Text
ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(4).Text = txtcampo(3).Text
ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(5).Text = txtcampo(5).Text
ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(6).Text = txtcampo(6).Text
ListView1.ListItems(ListView1.SelectedItem.Index).Tag = s
End Sub

Private Sub cmdBotón_Click(Index As Integer)
Dim Y As Byte, adors As New ADODB.Recordset, i As Integer, s As String, s1 As String
If Index = 1 Or Index = 0 And myAcción = 0 Then
    gs = "cancelar"
    Unload Me
    Exit Sub
End If
'Validad datos
If Len(Trim(txtcampo(0).Text)) = 0 Then
    MsgBox "El número de oficio de condonación es requerido. Favor de capturarlo", vbOKOnly + vbInformation
    txtcampo(0).SetFocus
    Exit Sub
End If
adors.Open "select count(*) from seguimientocondonación where oficio='" & Replace(txtcampo(0).Text, "'", "''") & "' and idseg<>" & mlSeguimiento, gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    MsgBox "El número de oficio de condonación ya existe. Favor de verificar y cambiar el oficio", vbOKOnly + vbInformation
    txtcampo(0).SetFocus
    Exit Sub
End If
If Not IsDate(txtcampo(1).Text) Then
    MsgBox "Todo los datos son requeridos. Favor de capturar la fecha", vbOKOnly + vbInformation
    txtcampo(1).SetFocus
    Exit Sub
End If
gs = txtcampo(0).Text & "|" & Format(txtcampo(1).Text, "dd/mm/yyyy") & "|"
For i = 1 To ListView1.ListItems.Count
    s = ListView1.ListItems(i).Tag
    s = Mid(s, InStr(s, "|") + 1)
    s = Mid(s, InStr(s, "|") + 1)
    If LCase(ListView1.ListItems(i).SubItems(2)) <> "no" Then
        gs = gs & ListView1.ListItems(i).Tag
        If Val(Mid(s, 1, InStr(s, "|") - 1)) <= 0 Then
            s1 = "Falta capturar el porcentaje de la " & i & "a Causa"
            Exit For
        End If
    End If
Next
If Len(s1) > 0 Then
    MsgBox s1 & ". Favor de verificar los datos", vbOKOnly + vbInformation, "Validación de datos"
    Exit Sub
End If
bAceptar = True
Unload Me
End Sub

Private Sub ComboUnidad_Click()
HabilitaAceptar
If Val(txtcampo(2).Text) > 0 Then ValidaMonto
End Sub

Private Sub ComboUnidad_LostFocus()
ActualizaTag
End Sub

Private Sub Form_Activate()
Dim Y As Byte, s As String, s1 As String, i As Integer, n As Integer, sSub As String
Dim adors As New ADODB.Recordset
On Error GoTo salir:
mlAnálisis = Val(gs1) 'Trae el valor de id análisis
mlSeguimiento = Val(gs2) 'Trae el valor de id seguimiento
s = gs 'Contiene todos los datos de la lista detalle por causa
If InStr(s, "|") Then
    s1 = Mid(s, 1, InStr(s, "|") - 1)
    txtcampo(0).Text = s1
    s = Mid(s, InStr(s, "|") + 1)
    s1 = Mid(s, 1, InStr(s, "|") - 1)
    txtcampo(1).Text = s1
    s = Mid(s, InStr(s, "|") + 1)
    Y = 0
    ListView1.ListItems.Clear
    n = 1
    Do While InStr(InStr(InStr(InStr(InStr(InStr(InStr(s, "|") + 1, s, "|") + 1, s, "|") + 1, s, "|") + 1, s, "|") + 1, s, "|") + 1, s, "|") > 0
        s1 = Mid(s, 1, InStr(s, "|") - 1)
        If adors.State Then adors.Close
        adors.Open "select paq_conceptos.leycausa(" & s1 & ") from dual", gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            ListView1.ListItems.Add n, , adors(0)
        Else
            ListView1.ListItems.Add n, , "LeyCau No Localizado"
        End If
        'ListView1.ListItems(n).Tag = s1 & "|"
        
        'Datos completos para el Tag
        s1 = Mid(s, 1, InStr(InStr(InStr(InStr(InStr(InStr(InStr(s, "|") + 1, s, "|") + 1, s, "|") + 1, s, "|") + 1, s, "|") + 1, s, "|") + 1, s, "|"))
        ListView1.ListItems(n).Tag = s1
        
        s = Mid(s, InStr(s, "|") + 1)
        s1 = Mid(s, 1, InStr(s, "|") - 1)
        ListView1.ListItems(n).SubItems(1) = s1
        
        s = Mid(s, InStr(s, "|") + 1)
        
        If Val(s) > 0 Then
            'ListView1.ListItems.Add n, , "Procede"
            ListView1.ListItems(n).SubItems(2) = "Si"
            'ListView1.ListItems(n).Tag = ListView1.ListItems(n).Tag & "1|"
        Else
            'ListView1.ListItems.Add n, , "No Procede"
            ListView1.ListItems(n).SubItems(2) = "No"
            'ListView1.ListItems(n).Tag = ListView1.ListItems(n).Tag & "0|"
        End If
        
        s = Mid(s, InStr(s, "|") + 1)
        s1 = Mid(s, 1, InStr(s, "|") - 1)
        'i = BuscaCombo(ComboUnidad, s1, True)
        'If i >= 0 Then
        '    ListView1.ListItems(n).SubItems(3) = ComboUnidad.List(i)
        'Else
        '    ListView1.ListItems(n).SubItems(3) = ""
        'End If
        ListView1.ListItems(n).SubItems(3) = Format(Val(s1), "###.00")
        s = Mid(s, InStr(s, "|") + 1)
        s1 = Mid(s, 1, InStr(s, "|") - 1)
        ListView1.ListItems(n).SubItems(4) = Format(Val(s1), "###,###,###.00")
        s = Mid(s, InStr(s, "|") + 1)
        s1 = Mid(s, 1, InStr(s, "|") - 1)
        
        ListView1.ListItems(n).SubItems(5) = Format(Val(s1), "###,###,###.00")
        s = Mid(s, InStr(s, "|") + 1)
        s1 = Mid(s, 1, InStr(s, "|") - 1)
        
        ListView1.ListItems(n).SubItems(6) = Format(Val(s1), "###,###,###.00")
        s = Mid(s, InStr(s, "|") + 1)
        'ListView1.ListItems(i).Tag = adors(5) 'Guarda el id
        n = n + 1
    Loop
    nCausas = n - 1
    
Else
    If adors.State Then adors.Close
    adors.Open "{call P_Cond_ConsImpSanXCau(" & mlAnálisis & ")}", gConSql, adOpenForwardOnly, adLockReadOnly
    ListView1.ListItems.Clear
    n = 1
    'sSub = "abcdefghijklmnñopqrstuvwxyz"
    Do While Not adors.EOF
        ListView1.ListItems.Add n, , adors(1)
        s1 = adors(0) & "|" & adors(2) & "|" & IIf(LCase(adors(3)) = "si", 1, 0) & "|0|" & adors(4) & "|0|0|"
        ListView1.ListItems(n).Tag = s1
        'ListView1.ListItems.Add n, , adors(0)
        ListView1.ListItems(n).SubItems(1) = adors(2)
        ListView1.ListItems(n).SubItems(2) = adors(3)
        ListView1.ListItems(n).SubItems(3) = "no definido"
        ListView1.ListItems(n).SubItems(4) = Format(adors(4), "###,###,###.00")
        ListView1.ListItems(n).SubItems(5) = Format(0, "###,###,###.00")
        ListView1.ListItems(n).SubItems(6) = Format(0, "###,###,###.00")
        n = n + 1
        adors.MoveNext
    Loop
    nCausas = n - 1
End If
If Len(Trim(txtcampo(0).Text)) = 0 Then
    'Dim adors As New ADODB.Recordset
    If adors.State Then adors.Close
    adors.Open "select f_nuevofolio(5,0," & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        If Not IsNull(adors(0)) Then
            If InStr(adors(0), "???") Then
                txtcampo(0).Text = Replace(adors(0), "???", "1")
            Else
                txtcampo(0).Text = adors(0)
            End If
        End If
    End If
End If
Exit Sub
salir:
i = MsgBox("Error no esperado: " & Err.Description, vbQuestion + vbAbortRetryIgnore)
If i = vbRetry Then
    Resume
ElseIf i = vbIgnore Then
    Resume Next
End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim s As String, i As Integer, l As Single
s = Item.Tag
s = Mid(s, InStr(s, "|") + 1)
txtcampo(4).Text = Mid(s, 1, InStr(s, "|") - 1)
s = Mid(s, InStr(s, "|") + 1)
l = Val(s)
bValidachk1 = True
If l <> 0 And l <> 1 Then
    Check1.Value = 2
    ComboUnidad.ListIndex = -1
    ComboUnidad.Text = ""
    txtcampo(2).Text = ""
    txtcampo(3).Text = ""
    txtcampo(4).Text = ""
    txtcampo(5).Text = ""
    txtcampo(6).Text = ""
Else
    Check1.Value = IIf(l = 1, 1, 0)
    s = Mid(s, InStr(s, "|") + 1)
    l = Val(s)
    If l > 0 Then
        txtcampo(2).Text = Format(l, "##.00")
    Else
        txtcampo(2).Text = ""
    End If
    s = Mid(s, InStr(s, "|") + 1)
    l = Val(s)
    If l > 0 Then
        txtcampo(3).Text = Format(l, "###,###,###.00")
    Else
        txtcampo(3).Text = ""
    End If
    s = Mid(s, InStr(s, "|") + 1)
    l = Val(s)
    If l > 0 Then
        txtcampo(5).Text = Format(l, "###,###,###.00")
    Else
        txtcampo(5).Text = ""
    End If
    s = Mid(s, InStr(s, "|") + 1)
    l = Val(s)
    If l > 0 Then
        txtcampo(6).Text = Format(l, "###,###,###.00")
    Else
        txtcampo(6).Text = ""
    End If
End If
bValidachk1 = False
End Sub

Private Sub txtCampo_Change(Index As Integer)
Dim m As Currency
    If Index = 2 Then
        If Val(Replace(txtcampo(3).Text, ",", "")) > 0 Then
            m = Val(Replace(txtcampo(3).Text, ",", ""))
            txtcampo(5).Text = Format(m * (Val(txtcampo(2).Text) / 100), "###,###,###.00")
            txtcampo(6).Text = Format(m * (1 - Val(txtcampo(2).Text) / 100), "###,###,###.00")
        End If
    End If
HabilitaAceptar
End Sub

Private Sub txtCampo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Static i As Integer
If KeyCode = 27 And Index = 3 Then
    txtcampo(Index) = ""
ElseIf KeyCode = 119 And Index = 6 Then
End If

End Sub

Private Sub txtCampo_KeyPress(Index As Integer, KeyAscii As Integer)
Dim m As Currency
    KeyAscii = TeclaOprimida(txtcampo(Index), KeyAscii, txtcampo(Index).Tag, yTipoAcción = 3)
End Sub

Private Sub txtCampo_LostFocus(Index As Integer)
Static bValida As Boolean, d As Date, rs As DAO.Recordset, adors As New ADODB.Recordset
If txtcampo(Index).Tag = "f" Then
    If IsDate(txtcampo(Index)) Then
        txtcampo(Index) = Format(CDate(txtcampo(Index)), gsFormatoFecha)
        If Index = 1 And IsDate(txtcampo(1)) Then
            d = AhoraServidor
            If CDate(txtcampo(Index).Text) - d > 0 Then
                MsgBox etiTexto(1).Caption & " no puede ser mayor a la fecha de hoy ", vbCritical + vbOKOnly, "Validación de captura"
                If Not bValida Then
                    txtcampo(Index).SetFocus
                    bValida = True
                    Exit Sub
                End If
            End If
'            If mdFechaOficio < CDate(txtCampo(Index).Text) Then
'                MsgBox etiTexto(Index).Caption & " no puede ser mayor que la fecha de la sanción", vbCritical + vbOKOnly, "Validación de captura"
'                If Not bValida Then
'                    txtCampo(Index).SetFocus
'                    bValida = True
'                    Exit Sub
'                End If
'            End If
        End If
    ElseIf Len(Trim(txtcampo(Index))) > 0 Then
        MsgBox "Fecha inválida. Favor de Corregirla", 0, "Validación de captura"
        If Not bValida Then
            txtcampo(Index).SetFocus
            bValida = True
            Exit Sub
        End If
    End If
End If
If Index >= 2 Then
    ActualizaTag
End If
bValida = False
End Sub


Function ColocaConsecutivo(sTexto As String, iCon As Integer) As String
Dim i As Integer
For i = Len(sTexto) To 1 Step -1
    If InStr("0123456789", Mid(sTexto, i, 1)) > 0 Then
        sTexto = Mid(sTexto, 1, i - 1)
    Else
        Exit For
    End If
Next
ColocaConsecutivo = sTexto & iCon
End Function

Sub HabilitaAceptar()
If (IsDate(txtcampo(1).Text) And Len(Trim(txtcampo(0).Text)) > 0 And Val(Replace(txtcampo(2).Text, ",", "")) > 0) Or myAcción = 0 Then
    cmdBotón(0).Enabled = True
ElseIf cmdBotón(0).Enabled Then
    cmdBotón(0).Enabled = False
End If
End Sub

Private Sub ValidaMonto()
Dim i As Integer, adors As New ADODB.Recordset
If ComboUnidad.ListIndex >= 0 And Val(txtcampo(2).Text) Then
    i = ComboUnidad.ItemData(ComboUnidad.ListIndex)
    adors.Open "select max_sanción from unidadmonetaria where id=" & i, gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        If Not IsNull(adors(0)) Then
            If adors(0) > 0 Then
                If Val(txtcampo(2).Text) > adors(0) Then
                    MsgBox "El monto de " & ComboUnidad.Text & " no puede ser mayor a " & adors(0), vbInformation + vbOKOnly, "Validación"
                    txtcampo(2).Text = adors(0)
                    Exit Sub
                End If
            End If
        End If
    End If
End If
End Sub
