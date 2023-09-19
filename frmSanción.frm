VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Sanción 
   Appearance      =   0  'Flat
   Caption         =   "Información de la Sanción (Multa)"
   ClientHeight    =   6015
   ClientLeft      =   7080
   ClientTop       =   5280
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   11085
   Begin VB.Frame Frame3 
      Caption         =   "Ley Causa Selecionada"
      Height          =   4110
      Left            =   135
      TabIndex        =   20
      Top             =   1035
      Width           =   10860
      Begin VB.CommandButton cmdAccion 
         Caption         =   "&Deshacer"
         Height          =   450
         Index           =   2
         Left            =   4724
         TabIndex        =   11
         Top             =   1305
         Width           =   1365
      End
      Begin VB.TextBox txtCampo 
         DataSource      =   "datAsunto"
         Height          =   285
         Index           =   1
         Left            =   2925
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "f"
         Top             =   675
         Width           =   1410
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Procede"
         Height          =   285
         Left            =   4950
         TabIndex        =   5
         Top             =   720
         Value           =   1  'Checked
         Width           =   960
      End
      Begin VB.ComboBox ComboUnidad 
         Height          =   315
         Left            =   5940
         TabIndex        =   6
         Top             =   675
         Width           =   1860
      End
      Begin VB.TextBox txtCampo 
         DataSource      =   "datAsunto"
         Height          =   285
         Index           =   2
         Left            =   7875
         MaxLength       =   20
         TabIndex        =   7
         Tag             =   "n"
         Top             =   675
         Width           =   1395
      End
      Begin VB.TextBox txtCampo 
         DataSource      =   "datAsunto"
         Height          =   285
         Index           =   3
         Left            =   9315
         MaxLength       =   30
         TabIndex        =   8
         Tag             =   "n"
         Top             =   675
         Width           =   1290
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H00E0E0E0&
         DataSource      =   "datAsunto"
         Height          =   285
         Index           =   4
         Left            =   4410
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   4
         Tag             =   "c"
         Top             =   675
         Width           =   390
      End
      Begin VB.ComboBox comboCausa 
         Height          =   315
         Left            =   225
         TabIndex        =   2
         Top             =   675
         Width           =   2715
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "&Nueva Ley(Causa)"
         Height          =   450
         Index           =   0
         Left            =   810
         TabIndex        =   9
         Top             =   1305
         Width           =   1365
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "&Quitar"
         Height          =   450
         Index           =   4
         Left            =   8460
         TabIndex        =   13
         Top             =   1305
         Width           =   1185
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "Ac&tualizar"
         Height          =   450
         Index           =   3
         Left            =   6681
         TabIndex        =   12
         Top             =   1305
         Width           =   1185
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "&Agregar"
         Enabled         =   0   'False
         Height          =   450
         Index           =   1
         Left            =   2767
         TabIndex        =   10
         Top             =   1305
         Width           =   1365
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2040
         Left            =   45
         TabIndex        =   14
         Top             =   1890
         Width           =   10680
         _ExtentX        =   18838
         _ExtentY        =   3598
         View            =   3
         LabelEdit       =   1
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
            Text            =   "Ley (Causa)"
            Object.Width           =   6880
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Sub"
            Object.Width           =   441
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Procede"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Moneda"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Monto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Importe Pesos"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label EtiTexto 
         Caption         =   "Fecha Infracción:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   2925
         TabIndex        =   26
         Top             =   450
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Unidad de la Multa"
         Height          =   240
         Left            =   5985
         TabIndex        =   25
         Top             =   420
         Width           =   1635
      End
      Begin VB.Label EtiTexto 
         Caption         =   "Monto de la Multa:"
         Height          =   255
         Index           =   2
         Left            =   7875
         TabIndex        =   24
         Top             =   435
         Width           =   1320
      End
      Begin VB.Label EtiTexto 
         Caption         =   "Importe en pesos:"
         Height          =   255
         Index           =   3
         Left            =   9330
         TabIndex        =   23
         Top             =   405
         Width           =   1425
      End
      Begin VB.Label EtiTexto 
         Caption         =   "SubInd:"
         Height          =   255
         Index           =   5
         Left            =   4425
         TabIndex        =   22
         Top             =   450
         Width           =   525
      End
      Begin VB.Label Label2 
         Caption         =   "Ley (Causa):"
         Height          =   195
         Left            =   270
         TabIndex        =   21
         Top             =   405
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   3240
      TabIndex        =   0
      Top             =   5175
      Width           =   4755
      Begin VB.CommandButton cmdBotón 
         Caption         =   "A&ceptar"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   16
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton cmdBotón 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   2700
         TabIndex        =   18
         Top             =   180
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   90
      TabIndex        =   15
      Top             =   135
      Width           =   10935
      Begin VB.TextBox txtCampo 
         DataSource      =   "datAsunto"
         Height          =   285
         Index           =   0
         Left            =   225
         MaxLength       =   70
         TabIndex        =   1
         Tag             =   "c"
         Top             =   450
         Width           =   4395
      End
      Begin VB.Label EtiTexto 
         Caption         =   "Seleccione y capture la información de cada causa asociada a la sanción:"
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   19
         Top             =   3015
         Width           =   5835
      End
      Begin VB.Label EtiTexto 
         Caption         =   "No. Oficio de Sanción:"
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   0
         Left            =   255
         TabIndex        =   17
         Top             =   225
         Width           =   2385
      End
   End
End
Attribute VB_Name = "Sanción"
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
Dim msCausas As String  'Guarda las causas
Dim msSub As String 'Subíndices de cada renglón
Dim bValidachk1 As Boolean 'Bandera que indica si debe validar o no el check1
Dim nCausas As Integer 'Número de causas asociadas al oficio de sanción
Dim mbCambio As Boolean 'indica si hubo movimientos de los datos de la lista
Dim miLeyCausa As Long 'Guarda el id de Ley causa del reg actual
Dim msFecha As String 'Guarda la fecha del reg actual
Dim msMonto As String 'Guarda el monto
Dim msPesos As String 'Guarda el monto en pesos
Dim miMoneda As Integer 'Guarda el id de la moneda

Private Sub Check1_Click()
If bValidachk1 Then Exit Sub
If Check1.Value = 0 Then
    If ComboUnidad.ListIndex >= 0 Then
        If MsgBox("¿Está seguro de indicar que no procede esta ley/causa?", vbYesNo + vbQuestion, "Validación") = vbNo Then
            Check1.Value = 1
            Exit Sub
        End If
    End If
    ComboUnidad.ListIndex = -1
    ComboUnidad.Text = ""
    ComboUnidad.Locked = True
    txtCampo(2).Text = ""
    txtCampo(2).Locked = True
    txtCampo(3).Text = ""
    txtCampo(3).Locked = True
    's = ListView1.ListItems(ListView1.SelectedItem.Index).Tag
    'ListView1.ListItems(ListView1.SelectedItem.Index).Tag = Mid(s, 1, InStr(s, "|")) & "0|0|0|0|"
Else
    txtCampo(2).Locked = False
    txtCampo(3).Locked = False
    ComboUnidad.Locked = False
End If
End Sub

Sub ActualizaTag(bNuevo As Boolean)
Dim s As String, i As Integer
'CausaLey
If bNuevo Then
    If comboCausa.ListIndex < 0 Then
        ListView1.ListItems.Add , 0 & txtCampo(4).Text, 0
    Else
        ListView1.ListItems.Add , comboCausa.ItemData(comboCausa.ListIndex) & txtCampo(4).Text, comboCausa.List(comboCausa.ListIndex)
    End If
    i = ListView1.ListItems.Count
Else
    i = ListView1.SelectedItem.Index
    ListView1.ListItems(i).Text = comboCausa.List(comboCausa.ListIndex)
End If
If comboCausa.ListIndex < 0 Then
    s = "0|"
Else
    s = comboCausa.ItemData(comboCausa.ListIndex) & "|"
End If
's = Mid(ListView1.ListItems(ListView1.SelectedItem.Index).Tag, 1, InStr(ListView1.ListItems(ListView1.SelectedItem.Index).Tag, "|"))
s = s & txtCampo(1).Text & "|" 'Fecha
ListView1.ListItems(i).SubItems(1) = txtCampo(1).Text
s = s & txtCampo(4).Text & "|" 'Procede
ListView1.ListItems(i).SubItems(2) = txtCampo(4).Text
s = s & IIf(Check1.Value <= 1, Check1.Value, -1) & "|"
If Check1.Value = 1 Then
    ListView1.ListItems(i).SubItems(3) = "Procede"
ElseIf Check1.Value = 0 Then
    ListView1.ListItems(i).SubItems(3) = "No Procede"
Else
    ListView1.ListItems(i).SubItems(3) = "No Definido"
End If
If ComboUnidad.ListIndex >= 0 Then
    s = s & ComboUnidad.ItemData(ComboUnidad.ListIndex) & "|"
    ListView1.ListItems(i).SubItems(4) = ComboUnidad.List(ComboUnidad.ListIndex)
Else
    s = s & "0|"
    ListView1.ListItems(i).SubItems(4) = ""
End If
s = s & Replace(Replace(txtCampo(2).Text, ",", ""), "$", "") & "|" & Replace(Replace(txtCampo(3).Text, ",", ""), "$", "") & "|"
ListView1.ListItems(i).SubItems(5) = txtCampo(2).Text
ListView1.ListItems(i).SubItems(6) = txtCampo(3).Text
ListView1.ListItems(i).Tag = s
End Sub

'Actualiza campos a partir de la cadena psTag
Function ActualizaCampos(psTag As String)
Dim s1 As String, s As String
s = psTag

s1 = Mid(s, 1, InStr(s, "|") - 1)
'1er dato de la cadena ley_causa
comboCausa.ListIndex = BuscaCombo(comboCausa, s1, True)

'Segundo dato Fecha
s = Mid(s, InStr(s, "|") + 1)
s1 = Mid(s, 1, InStr(s, "|") - 1)
txtCampo(1).Text = s1

'3er dato subindice
s = Mid(s, InStr(s, "|") + 1)
s1 = Mid(s, 1, InStr(s, "|") - 1)
txtCampo(4).Text = s1

'4o dato Procede
s = Mid(s, InStr(s, "|") + 1)
If Val(s) > 0 Then
    Check1.Value = 1
Else
    Check1.Value = 0
End If

'5o dato Moneda
s = Mid(s, InStr(s, "|") + 1)
s1 = Mid(s, 1, InStr(s, "|") - 1)
ComboUnidad.ListIndex = BuscaCombo(ComboUnidad, s1, True)

'6o dato Monto
s = Mid(s, InStr(s, "|") + 1)
s1 = Mid(s, 1, InStr(s, "|") - 1)
txtCampo(2).Text = s1

s = Mid(s, InStr(s, "|") + 1)
s1 = Mid(s, 1, InStr(s, "|") - 1)

'7o dato Monto pesos
txtCampo(3).Text = s1

End Function

Private Sub cmdAccion_Click(Index As Integer)
Dim i As Integer
If Index = 0 Then 'Nuevo
    comboCausa.ListIndex = -1
    comboCausa.Text = ""
    ComboUnidad.ListIndex = -1
    ComboUnidad.Text = ""
    comboCausa.Text = ""
    txtCampo(1).Text = ""
    txtCampo(2).Text = ""
    txtCampo(3).Text = ""
    i = ListView1.ListItems.Count
    If i > 0 Then
        txtCampo(4).Text = Mid(msSub, i + 1, 1)
    Else
        txtCampo(4).Text = "a"
    End If
    Check1.Value = 0
    mbCambio = False
    cmdAccion(0).Enabled = False
    cmdAccion(1).Enabled = True
    cmdAccion(2).Enabled = True
    cmdAccion(3).Enabled = False
    cmdAccion(4).Enabled = False
    ListView1.Enabled = False
ElseIf Index = 1 Then 'Agregar
    If comboCausa.ListIndex < 0 Then
        Call MsgBox("Debe seleccionar la Ley(Causa)", vbOKOnly, "Validación")
        comboCausa.SetFocus
        Exit Sub
    End If
    If Check1.Value Then
        If ComboUnidad.ListIndex < 0 Then
            Call MsgBox("Debe seleccionar la Moneda", vbOKOnly, "Validación")
            ComboUnidad.SetFocus
            Exit Sub
        End If
        If Not IsDate(txtCampo(1).Text) Then
            Call MsgBox("Debe capturar la fecha", vbOKOnly, "Validación")
            txtCampo(1).SetFocus
            Exit Sub
        End If
        If Val(txtCampo(2).Text) <= 0 Then
            Call MsgBox("Debe capturar el monto", vbOKOnly, "Validación")
            txtCampo(2).SetFocus
            Exit Sub
        End If
    Else
    End If
    ActualizaTag True
    HabilitaAceptar
    If Not cmdBotón(0).Enabled Then
        cmdAccion_Click (0)
    Else
        ListView1_ItemClick ListView1.ListItems(ListView1.ListItems.Count)
        cmdAccion(0).Enabled = True
        cmdAccion(1).Enabled = False
        cmdAccion(2).Enabled = True
        cmdAccion(3).Enabled = True
        cmdAccion(4).Enabled = True
        ListView1.Enabled = True
    End If
    
ElseIf Index = 2 Then 'Deshacer
    mbCambio = False
    If ListView1.ListItems.Count > 0 Then
        If ListView1.SelectedItem.Index > 0 Then
            ActualizaCampos ListView1.ListItems(ListView1.SelectedItem.Index).Tag
        End If
        cmdAccion(0).Enabled = True
        cmdAccion(1).Enabled = False
        cmdAccion(2).Enabled = True
        cmdAccion(3).Enabled = True
        cmdAccion(4).Enabled = True
        ListView1.Enabled = True
    Else
        cmdAccion_Click (0)
    End If
ElseIf Index = 3 Then 'Actualizar
    If ListView1.SelectedItem.Index > 0 Then
        ActualizaTag False
    End If
    HabilitaAceptar
ElseIf Index = 4 Then 'Eliminar
    If ListView1.SelectedItem.Index > 0 Then
        If MsgBox("Está seguro de Eliminar el registro actual", vbQuestion + vbYesNo, "Confirmación") = vbNo Then
            Exit Sub
        End If
        i = ListView1.SelectedItem.Index
        ListView1.ListItems.Remove i
        If ListView1.ListItems.Count > 0 Then 'En caso de haber registro coloca el que ocupaba el eliminado
            If ListView1.ListItems.Count >= i Then
                ListView1_ItemClick ListView1.ListItems(i)
            Else
                If i > 1 Then
                    ListView1_ItemClick ListView1.ListItems(i - 1)
                Else
                    ListView1_ItemClick ListView1.ListItems(ListView1.ListItems.Count)
                End If
            End If
            ActualizaTag False
        Else
            cmdAccion_Click (0)
        End If
    End If
    HabilitaAceptar
End If
End Sub

Private Sub cmdBotón_Click(Index As Integer)
Dim Y As Byte, adors As New ADODB.Recordset, i As Integer, s As String, s1 As String
If Index = 1 Or Index = 0 And myAcción = 0 Then
    gs = "cancelar"
    Unload Me
    Exit Sub
End If
'Validad datos
If Len(Trim(txtCampo(0).Text)) = 0 Then
    MsgBox "El número de oficio de sanción es requerido. Favor de capturarlo", vbOKOnly + vbInformation
    txtCampo(0).SetFocus
    Exit Sub
End If
adors.Open "select count(*) from seguimientosanción where oficio='" & Replace(txtCampo(0).Text, "'", "''") & "' and idseg<>" & mlSeguimiento, gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    MsgBox "El número de oficio de sanción ya existe. Favor de verificar y cambiar el oficio", vbOKOnly + vbInformation
    txtCampo(0).SetFocus
    Exit Sub
End If
If Not IsDate(txtCampo(1).Text) Then
    MsgBox "Todo los datos son requeridos. Favor de capturar la fecha de infracción", vbOKOnly + vbInformation
    txtCampo(1).SetFocus
    Exit Sub
End If
gs = txtCampo(0).Text & "|"
For i = 1 To ListView1.ListItems.Count
    s = ListView1.ListItems(i).Tag
    s = Mid(s, InStr(s, "|") + 1)
    If Not IsDate(Mid(s, 1, InStr(s, "|") - 1)) Then
        s1 = "Falta capturar la fecha del " & i & IIf(i = 1, "er.", "o.") & " registro"
        Exit For
    End If
    s = Mid(s, InStr(s, "|") + 1)
    If Len(Mid(s, 1, InStr(s, "|") - 1)) <= 0 And nCausas > 1 Then
        s1 = "Falta capturar el subíndice de la " & i & "a Causa"
        Exit For
    End If
    s = Mid(s, InStr(s, "|") + 1)
    If Val(s) = 1 Then
        s = Mid(s, InStr(s, "|") + 1)
        If Val(s) <= 0 Then
            s1 = "Falta capturar la unidad de la " & i & "a Causa"
            Exit For
        End If
        s = Mid(s, InStr(s, "|") + 1)
        If Val(s) <= 0 Then
            s1 = "Falta capturar el monto de la " & i & "a Causa"
            Exit For
        End If
        s = Mid(s, InStr(s, "|") + 1)
        If Val(s) <= 0 Then
            s1 = "Falta capturar el monto en pesos de la " & i & "a Causa"
            Exit For
        End If
    Else
        If Val(s) <> 0 Then
            s1 = "Falta capturar el status Procede/No Procede de la " & i & "a Causa"
            Exit For
        End If
    End If
    gs = gs & ListView1.ListItems(i).Tag
Next
If Len(s1) > 0 Then
    MsgBox s1 & ". Favor de verificar los datos", vbOKOnly + vbInformation, "Validación de datos"
    Exit Sub
End If
bAceptar = True
Unload Me
End Sub

Private Sub comboCausa_GotFocus()
If comboCausa.ListIndex >= 0 Then
    miLeyCausa = comboCausa.ItemData(comboCausa.ListIndex)
Else
    miLeyCausa = 0
End If
End Sub

Private Sub comboCausa_LostFocus()
If Not mbCambio Then
    If comboCausa.ListIndex >= 0 Then
        mbCambio = (comboCausa.ItemData(comboCausa.ListIndex) <> miLeyCausa)
    Else
        mbCambio = (0 <> miLeyCausa)
    End If
End If
End Sub

Private Sub ComboUnidad_Click()
If Val(txtCampo(2).Text) > 0 Then ValidaMonto
End Sub

Private Sub ComboUnidad_GotFocus()
If ComboUnidad.ListIndex >= 0 Then
    miLeyCausa = ComboUnidad.ItemData(ComboUnidad.ListIndex)
Else
    miLeyCausa = 0
End If
End Sub

Private Sub ComboUnidad_LostFocus()
If Not mbCambio Then
    If ComboUnidad.ListIndex >= 0 Then
        mbCambio = (ComboUnidad.ItemData(ComboUnidad.ListIndex) <> miMoneda)
    Else
        mbCambio = (0 <> miMoneda)
    End If
End If
End Sub

Private Sub Form_Activate()
Dim Y As Byte, s As String, s1 As String, i As Integer, n As Integer, nFinReg As Integer
Dim adors As New ADODB.Recordset
On Error GoTo salir:
LlenaCombo ComboUnidad, "select id, descripción from unidadmonetaria where fechabaja is null", "", True
mlAnálisis = Val(gs1)
mlSeguimiento = Val(gs2)
s = gs
Call LlenaCombo(comboCausa, "select id,paq_conceptos.ley(idley)||' ('||paq_conceptos.causa(idcau)||')' as descripción from análisiscausas where idana=" & mlAnálisis, "", True)
If InStr(s, "|") Then
    s1 = Mid(s, 1, InStr(s, "|") - 1)
    txtCampo(0).Text = s1
    s = Mid(s, InStr(s, "|") + 1)
    's1 = Mid(s, 1, InStr(s, "|") - 1)
    'txtCampo(1).Text = s1
    's = Mid(s, InStr(s, "|") + 1)
    Y = 0
    ListView1.ListItems.Clear
    n = 1
    nFinReg = InStr(InStr(InStr(InStr(InStr(InStr(InStr(s, "|") + 1, s, "|") + 1, s, "|") + 1, s, "|") + 1, s, "|") + 1, s, "|") + 1, s, "|")
    Do While nFinReg > 0
        s1 = Mid(s, 1, InStr(s, "|") - 1)
        '1er dato de la cadena ley_causa
        If adors.State Then adors.Close
        adors.Open "select paq_conceptos.leycausa(" & s1 & ") from dual", gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            ListView1.ListItems.Add n, , adors(0)
        Else
            ListView1.ListItems.Add n, , "LeyCau No Localizado"
        End If
        'ListView1.ListItems(n).Tag = s1 & "|"
        
        'Datos completos para el Tag
        ListView1.ListItems(n).Tag = Mid(s, 1, nFinReg)
        
        'Segundo dato Fecha
        s = Mid(s, InStr(s, "|") + 1)
        s1 = Mid(s, 1, InStr(s, "|") - 1)
        ListView1.ListItems(n).SubItems(1) = s1
        
        '3er dato subindice
        s = Mid(s, InStr(s, "|") + 1)
        s1 = Mid(s, 1, InStr(s, "|") - 1)
        ListView1.ListItems(n).SubItems(2) = s1
        
        '4o dato Procede
        s = Mid(s, InStr(s, "|") + 1)
        If Val(s) > 0 Then
            'ListView1.ListItems.Add n, , "Procede"
            ListView1.ListItems(n).SubItems(3) = "Procede"
            'ListView1.ListItems(n).Tag = ListView1.ListItems(n).Tag & "1|"
        Else
            'ListView1.ListItems.Add n, , "No Procede"
            ListView1.ListItems(n).SubItems(3) = "No Procede"
            'ListView1.ListItems(n).Tag = ListView1.ListItems(n).Tag & "0|"
        End If
        
        '5o dato Moneda
        s = Mid(s, InStr(s, "|") + 1)
        s1 = Mid(s, 1, InStr(s, "|") - 1)
        i = BuscaCombo(ComboUnidad, s1, True)
        If i >= 0 Then
            ListView1.ListItems(n).SubItems(4) = ComboUnidad.List(i)
        Else
            ListView1.ListItems(n).SubItems(4) = ""
        End If
        
        '6o dato Monto
        s = Mid(s, InStr(s, "|") + 1)
        s1 = Mid(s, 1, InStr(s, "|") - 1)
        ListView1.ListItems(n).SubItems(5) = Format(Val(s1), "###,###,###.00")
        s = Mid(s, InStr(s, "|") + 1)
        s1 = Mid(s, 1, InStr(s, "|") - 1)
        
        '7o dato Monto pesos
        ListView1.ListItems(n).SubItems(6) = Format(Val(s1), "###,###,###.00")
        s = Mid(s, InStr(s, "|") + 1)
        'ListView1.ListItems(i).Tag = adors(5) 'Guarda el id
        n = n + 1
        nFinReg = InStr(InStr(InStr(InStr(InStr(InStr(InStr(s, "|") + 1, s, "|") + 1, s, "|") + 1, s, "|") + 1, s, "|") + 1, s, "|") + 1, s, "|")
    Loop
    nCausas = n - 1
    
    
cmdAccion(1).Enabled = True
    
Else
    'If adors.State Then adors.Close
    'adors.Open "select id,paq_conceptos.ley(idley)||' ('||paq_conceptos.causa(idcau)||')' from análisiscausas where idana=" & mlAnálisis, gConSql, adOpenStatic, adLockReadOnly
    'ListView1.ListItems.Clear
    'n = 1
    'Do While Not adors.EOF
    '    ListView1.ListItems.Add n, , adors(1)
    '    s1 = adors(0) & "|" & Mid(sSub, n, 1) & "|-1|0|0|0|"
    '    ListView1.ListItems(n).Tag = s1
    '    'ListView1.ListItems.Add n, , adors(0)
    '    ListView1.ListItems(n).SubItems(1) = Mid(sSub, n, 1)
    '    ListView1.ListItems(n).SubItems(2) = "no definido"
    '    ListView1.ListItems(n).SubItems(3) = "no definido"
    '    ListView1.ListItems(n).SubItems(4) = Format(0, "###,###,###.00")
    '    ListView1.ListItems(n).SubItems(5) = Format(0, "###,###,###.00")
    '    n = n + 1
    '    adors.MoveNext
    'Loop
    'nCausas = n - 1
    cmdAccion_Click (0)
End If
If Len(Trim(txtCampo(0).Text)) = 0 Then
    'Dim adors As New ADODB.Recordset
    If adors.State Then adors.Close
    adors.Open "select f_nuevofolio(4,0," & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        If Not IsNull(adors(0)) Then
            If InStr(adors(0), "???") Then
                txtCampo(0).Text = Replace(adors(0), "???", "1")
            Else
                txtCampo(0).Text = adors(0)
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

Private Sub Form_Load()
msSub = "abcdefghijklmnñopqrstuvwxyz"
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim s As String, i As Integer, l As Single
s = Item.Tag
ActualizaCampos (s)
End Sub


Private Sub txtcampo_GotFocus(Index As Integer)
If Index = 1 Then 'Fecha
    msFecha = txtCampo(1).Text
ElseIf Index = 2 Then 'Monto
    msMonto = txtCampo(2).Text
ElseIf Index = 3 Then 'Monto
    msPesos = txtCampo(3).Text
End If
End Sub

Private Sub txtCampo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Static i As Integer
If KeyCode = 27 And Index = 3 Then
    txtCampo(Index) = ""
ElseIf KeyCode = 119 And Index = 6 Then
'    If lAsuIns > 0 Then
'        Set adors = New ADODB.Recordset
'        adors.Open "select count(*) from sanciones s inner join avances av on s.idava=av.id where s.imp_multa>0 and av.idasuins=" & lAsuIns, gConSql, adOpenStatic, adLockReadOnly
'        i = IIf(IsNull(adors(0)), 0, adors(0)) + 1
'        txtCampo(Index).Text = ColocaConsecutivo(txtCampo(Index).Text, i)
'    End If
End If

End Sub

Private Sub txtCampo_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = TeclaOprimida(txtCampo(Index), KeyAscii, txtCampo(Index).Tag, yTipoAcción = 3)
End Sub

Private Sub txtCampo_LostFocus(Index As Integer)
Static bValida As Boolean, d As Date, rs As DAO.Recordset, adors As New ADODB.Recordset
If txtCampo(Index).Tag = "f" Then
    If IsDate(txtCampo(Index)) Then
        txtCampo(Index) = Format(CDate(txtCampo(Index)), gsFormatoFecha)
        If Index = 1 And IsDate(txtCampo(1)) Then
            d = AhoraServidor
            If CDate(txtCampo(Index).Text) - d > 0 Then
                MsgBox etiTexto(1).Caption & " no puede ser mayor a la fecha de hoy ", vbCritical + vbOKOnly, "Validación de captura"
                If Not bValida Then
                    txtCampo(Index).SetFocus
                    bValida = True
                    Exit Sub
                End If
            End If
            If mdFechaOficio < CDate(txtCampo(Index).Text) Then
                MsgBox etiTexto(Index).Caption & " no puede ser mayor que la fecha de la sanción", vbCritical + vbOKOnly, "Validación de captura"
                If Not bValida Then
                    txtCampo(Index).SetFocus
                    bValida = True
                    Exit Sub
                End If
            End If
            If ComboUnidad.ListIndex >= 0 Then
                If ComboUnidad.ItemData(ComboUnidad.ListIndex) = 13 Then 'Sal Min
                    Call txtCampo_LostFocus(2)
                End If
            End If
        End If
    ElseIf Len(Trim(txtCampo(Index))) > 0 Then
        MsgBox "Fecha inválida. Favor de Corregirla", 0, "Validación de captura"
        If Not bValida Then
            txtCampo(Index).SetFocus
            bValida = True
            Exit Sub
        End If
    End If
End If
If Index = 2 And Val(txtCampo(2).Text) > 0 Then
    If ComboUnidad.ListIndex > 0 Then
        ValidaMonto
    End If
    If ComboUnidad.ListIndex >= 0 Then
        If ComboUnidad.ItemData(ComboUnidad.ListIndex) = 1 Then
            txtCampo(3).Text = txtCampo(2).Text
        End If
        If ComboUnidad.ItemData(ComboUnidad.ListIndex) = 13 Or ComboUnidad.ItemData(ComboUnidad.ListIndex) = 14 Then 'salarios mínimos o UMAS
            If IsDate(txtCampo(1).Text) Then
                If adors.State Then adors.Close
                adors.Open "select f_salmin_UMA(to_date('" & Format(CDate(txtCampo(1).Text), "dd/mm/yyyy") & "','dd/mm/yyyy')," & IIf(ComboUnidad.ItemData(ComboUnidad.ListIndex) = 13, 0, 1) & ") from dual", gConSql, adOpenStatic, adLockReadOnly
                If adors(0) > 0 Then
                    txtCampo(3).Text = Val(Replace(Replace(txtCampo(2).Text, ",", ""), "$", "")) * adors(0)
                End If
            Else
                MsgBox "Para calcular el monto en pesos de los salários mínimos o UMAS capture primeramente la Fecha de la Infracción", vbOKOnly + vbInformation, "Informativo"
                Exit Sub
            End If
        End If
    End If
End If
If Not mbCambio Then
    If Index = 1 Then 'Fecha
        mbCambio = (msFecha <> txtCampo(1).Text)
    ElseIf Index = 2 Then 'Monto
        mbCambio = (msMonto <> txtCampo(2).Text)
    ElseIf Index = 3 Then 'Monto
        mbCambio = (msPesos <> txtCampo(3).Text)
    End If
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
Dim i As Integer, b As Boolean, n As Integer, s As String
s = "|"
If Len(Trim(txtCampo(0))) > 0 Then
    For i = 1 To ListView1.ListItems.Count
        If InStr(s, "|" & Val(ListView1.ListItems(i).Tag) & "|") = 0 Then
            s = s & Val(ListView1.ListItems(i).Tag) & "|"
            n = n + 1
        End If
    Next
    b = (n >= comboCausa.ListCount)
End If
If b Then
    If Not cmdBotón(0).Enabled Then cmdBotón(0).Enabled = True
Else
    If cmdBotón(0).Enabled Then cmdBotón(0).Enabled = False
End If
End Sub

Private Sub ValidaMonto()
Dim i As Integer, adors As New ADODB.Recordset
If ComboUnidad.ListIndex >= 0 And Val(txtCampo(2).Text) Then
    i = ComboUnidad.ItemData(ComboUnidad.ListIndex)
    adors.Open "select max_sanción from unidadmonetaria where id=" & i, gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        If Not IsNull(adors(0)) Then
            If adors(0) > 0 Then
                If Val(txtCampo(2).Text) > adors(0) Then
                    MsgBox "El monto de " & ComboUnidad.Text & " no puede ser mayor a " & adors(0), vbInformation + vbOKOnly, "Validación"
                    txtCampo(2).Text = adors(0)
                    Exit Sub
                End If
            End If
        End If
    End If
End If
End Sub
