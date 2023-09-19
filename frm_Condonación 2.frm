VERSION 5.00
Begin VB.Form Condonación 
   Appearance      =   0  'Flat
   Caption         =   "Información de la Condonación (Multa)"
   ClientHeight    =   2625
   ClientLeft      =   2025
   ClientTop       =   1995
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1620
      Left            =   90
      TabIndex        =   1
      Top             =   135
      Width           =   7500
      Begin VB.TextBox txtCampo 
         BackColor       =   &H80000000&
         DataSource      =   "datAsunto"
         Height          =   285
         Index           =   4
         Left            =   5670
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   6
         Tag             =   "n"
         Top             =   1125
         Width           =   1605
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H80000000&
         DataSource      =   "datAsunto"
         Height          =   285
         Index           =   3
         Left            =   4005
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "n"
         Top             =   1125
         Width           =   1605
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H80000000&
         DataSource      =   "datAsunto"
         Height          =   285
         Index           =   2
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   4
         Tag             =   "n"
         Top             =   1125
         Width           =   1530
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   150
         MaxLength       =   3
         TabIndex        =   3
         Tag             =   "n"
         Top             =   1125
         Width           =   1020
      End
      Begin VB.TextBox txtCampo 
         DataSource      =   "datAsunto"
         Height          =   285
         Index           =   0
         Left            =   90
         MaxLength       =   70
         TabIndex        =   0
         Tag             =   "c"
         Top             =   450
         Width           =   4395
      End
      Begin VB.TextBox txtCampo 
         DataSource      =   "datAsunto"
         Height          =   285
         Index           =   1
         Left            =   4545
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "f"
         Top             =   450
         Width           =   2400
      End
      Begin VB.Label EtiTexto 
         Caption         =   "Monto a pagar:"
         Height          =   255
         Index           =   4
         Left            =   5670
         TabIndex        =   16
         Top             =   855
         Width           =   1545
      End
      Begin VB.Label EtiTexto 
         Caption         =   "Monto Condonado:"
         Height          =   255
         Index           =   3
         Left            =   4005
         TabIndex        =   15
         Top             =   855
         Width           =   1545
      End
      Begin VB.Label EtiTexto 
         Caption         =   "Monto total de la Multa:"
         Height          =   255
         Index           =   2
         Left            =   2250
         TabIndex        =   14
         Top             =   855
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Porcentaje de Condonación"
         Height          =   255
         Left            =   150
         TabIndex        =   13
         Top             =   855
         Width           =   2055
      End
      Begin VB.Label Porcentaje 
         Caption         =   "%"
         Height          =   255
         Left            =   1305
         TabIndex        =   12
         Top             =   1125
         Width           =   135
      End
      Begin VB.Label EtiTexto 
         Caption         =   "No. Oficio de Condonación:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   225
         Width           =   2385
      End
      Begin VB.Label EtiTexto 
         Caption         =   "Fecha:"
         Height          =   255
         Index           =   1
         Left            =   4635
         TabIndex        =   7
         Top             =   225
         Width           =   1740
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   1530
      TabIndex        =   9
      Top             =   1800
      Width           =   4755
      Begin VB.CommandButton cmdBotón 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   2700
         TabIndex        =   11
         Top             =   240
         Width           =   1185
      End
      Begin VB.CommandButton cmdBotón 
         Caption         =   "&Aceptar"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   10
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "Condonación"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const cLímInf = 500
Const cLímSup = 2000
Const cnGrisclaro = &HE0E0E0
Dim msCondonación As String
Public mdFechaOficio 'Fecha del oficio de condonacioón (fecha de la actividad que invoca este formulario)
Public myAcción 'Indica que tipo de acción se realizará 0:consulta; 1:Alta; 2: Modificación
Dim bAceptar As Boolean 'Variable lógica que indica si fueron aceptadados los cambios
Dim adors As New ADODB.Recordset
Dim mlAnálisis As Long, mlSeguimiento As Long

Private Sub cmdBotón_Click(Index As Integer)
Dim Y As Byte, adors As New ADODB.Recordset
If Index = 1 Or Index = 0 And myAcción = 0 Then
    gs = "cancelar"
    Unload Me
    Exit Sub
End If
'Validad datos
If Len(Trim(txtCampo(0).Text)) = 0 Then
    MsgBox "El número de oficio de condonación es requerido. Favor de capturarlo", vbOKOnly + vbInformation
    txtCampo(0).SetFocus
    Exit Sub
End If
adors.Open "select count(*) from seguimientocondonación where oficio='" & Replace(txtCampo(0).Text, "'", "''") & "' and idseg<>" & mlSeguimiento, gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    MsgBox "El número de oficio de condonación ya existe. Favor de verificar y cambiar el oficio", vbOKOnly + vbInformation
    txtCampo(0).SetFocus
    Exit Sub
End If
If Not IsDate(txtCampo(1).Text) Then
    MsgBox "Todo los datos son requeridos. Favor de capturar la fecha de infracción", vbOKOnly + vbInformation
    txtCampo(1).SetFocus
    Exit Sub
End If
gs = txtCampo(0).Text & "|" & Format(txtCampo(1).Text, "dd/mm/yyyy") & "|" & Trim(Text1.Text) & "|"
bAceptar = True
Unload Me
End Sub

Private Sub ComboUnidad_Click()
HabilitaAceptar
If Val(txtCampo(2).Text) > 0 Then ValidaMonto
End Sub

Private Sub Form_Activate()
Dim Y As Byte, s As String, s1 As String, i As Integer
Dim adors As New ADODB.Recordset
On Error GoTo salir:
'LlenaCombo ComboUnidad, "select id, descripción from unidadmonetaria where fechabaja is null", "", True
mlAnálisis = Val(gs1)
mlSeguimiento = Val(gs2)
s = gs
Y = 0
Do While InStr(s, "|")
    s1 = Mid(s, 1, InStr(s, "|") - 1)
    If Y = 2 Then
        'ComboUnidad.ListIndex = BuscaCombo(ComboUnidad, Val(s1), True)
    ElseIf Y > 1 Then
        txtCampo(Y - 1).Text = s1
    Else
        txtCampo(Y).Text = s1
    End If
    s = Mid(s, InStr(s, "|") + 1)
    Y = Y + 1
Loop
If Len(Trim(txtCampo(0).Text)) = 0 Then
    adors.Open "select f_nuevofolio(5,0," & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
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
On Error GoTo 0
If adors.State > 0 Then adors.Close
adors.Open "select f_analisis_sancion_imp(" & mlAnálisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
If Not adors.EOF Then
    txtCampo(2).Text = Format(adors(0), "$###,###,###.00")
End If
If Val(Text1.Text) > 0 Then
    Text1_LostFocus
End If
salir:
End Sub

Private Sub Label3_Click()
End Sub

Private Sub Text1_Change()
HabilitaAceptar
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = TeclaOprimida(txtCampo(Index), KeyAscii, txtCampo(Index).Tag, True)
End Sub

Private Sub Text1_LostFocus()
Dim lMulta As Single
If Val(Text1.Text) >= 100 And Text1.Enabled Then
    Call MsgBox("El porcentaje no puede ser mayor o igual a cien", vbInformation + vbOKOnly)
    Text1.Text = 50
    Exit Sub
End If
If Val(Text1.Text) < 0 Then
    Call MsgBox("El porcentaje no puede ser menor a cero", vbInformation + vbOKOnly)
    Text1.Text = 0
    Exit Sub
End If
lMulta = Val(Replace(Replace(txtCampo(2).Text, "$", ""), ",", ""))
    txtCampo(3).Text = Format(Val(Text1.Text) * lMulta / 100, "$###,###,###.00")
    txtCampo(4).Text = Format((100 - Val(Text1.Text)) * lMulta / 100, "$###,###,###.00")
End Sub

Private Sub txtCampo_Change(Index As Integer)
Dim Y As Byte
HabilitaAceptar
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
Static bValida As Boolean, d As Date, rs As DAO.Recordset, adors As ADODB.Recordset
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
                MsgBox etiTexto(Index).Caption & " no puede ser mayor que la fecha de la condonación", vbCritical + vbOKOnly, "Validación de captura"
                If Not bValida Then
                    txtCampo(Index).SetFocus
                    bValida = True
                    Exit Sub
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
If IsDate(txtCampo(1).Text) And Len(txtCampo(2).Text) > 0 And Val(Replace(Text1.Text, ",", "")) > 0 Or myAcción = 0 Then
    cmdBotón(0).Enabled = True
ElseIf cmdBotón(0).Enabled Then
    cmdBotón(0).Enabled = False
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


