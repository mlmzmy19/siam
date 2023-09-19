VERSION 5.00
Begin VB.Form Condonaci�n 
   Appearance      =   0  'Flat
   Caption         =   "Informaci�n de la Sanci�n (Multa)"
   ClientHeight    =   2475
   ClientLeft      =   2025
   ClientTop       =   1995
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   1350
      TabIndex        =   6
      Top             =   1665
      Width           =   4755
      Begin VB.CommandButton cmdBot�n 
         Caption         =   "&Aceptar"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton cmdBot�n 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   2700
         TabIndex        =   5
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1620
      Left            =   45
      TabIndex        =   8
      Top             =   45
      Width           =   7335
      Begin VB.ComboBox ComboUnidad 
         Height          =   315
         Left            =   90
         TabIndex        =   1
         Top             =   1080
         Width           =   1995
      End
      Begin VB.TextBox txtCampo 
         DataSource      =   "datAsunto"
         Height          =   285
         Index           =   0
         Left            =   90
         MaxLength       =   70
         TabIndex        =   7
         Tag             =   "c"
         Top             =   450
         Width           =   4395
      End
      Begin VB.TextBox txtCampo 
         DataSource      =   "datAsunto"
         Height          =   285
         Index           =   2
         Left            =   2205
         MaxLength       =   20
         TabIndex        =   2
         Tag             =   "n"
         Top             =   1095
         Width           =   2280
      End
      Begin VB.TextBox txtCampo 
         DataSource      =   "datAsunto"
         Height          =   285
         Index           =   1
         Left            =   4545
         MaxLength       =   30
         TabIndex        =   0
         Tag             =   "f"
         Top             =   450
         Width           =   2400
      End
      Begin VB.TextBox txtCampo 
         DataSource      =   "datAsunto"
         Height          =   285
         Index           =   3
         Left            =   4590
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "n"
         Top             =   1095
         Width           =   2325
      End
      Begin VB.Label Label1 
         Caption         =   "Unidad de la Multa"
         Height          =   285
         Left            =   90
         TabIndex        =   13
         Top             =   810
         Width           =   1635
      End
      Begin VB.Label EtiTexto 
         Caption         =   "No. Oficio de Sanci�n:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   225
         Width           =   2385
      End
      Begin VB.Label EtiTexto 
         Caption         =   "Monto de la Multa:"
         Height          =   255
         Index           =   2
         Left            =   2205
         TabIndex        =   11
         Top             =   855
         Width           =   1815
      End
      Begin VB.Label EtiTexto 
         Caption         =   "Fecha de la infracci�n:"
         Height          =   255
         Index           =   1
         Left            =   4635
         TabIndex        =   10
         Top             =   225
         Width           =   1740
      End
      Begin VB.Label EtiTexto 
         Caption         =   "Importe en pesos:"
         Height          =   255
         Index           =   3
         Left            =   4605
         TabIndex        =   9
         Top             =   825
         Width           =   2325
      End
   End
End
Attribute VB_Name = "Condonaci�n"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const cL�mInf = 500
Const cL�mSup = 2000
Const cnGrisclaro = &HE0E0E0
Dim msSanci�n As String
Public mdFechaOficio 'Fecha del oficio de sanci�n (fecha de la actividad que invoca este formulario)
Public myAcci�n 'Indica que tipo de acci�n se realizar� 0:consulta; 1:Alta; 2: Modificaci�n
Dim bAceptar As Boolean 'Variable l�gica que indica si fueron aceptadados los cambios
Dim adors As New ADODB.Recordset
Dim mlAn�lisis As Long, mlSeguimiento As Long

Private Sub cmdBot�n_Click(Index As Integer)
Dim Y As Byte, adors As New ADODB.Recordset
If Index = 1 Or Index = 0 And myAcci�n = 0 Then
    gs = "cancelar"
    Unload Me
    Exit Sub
End If
'Validad datos
If Len(Trim(txtcampo(0).Text)) = 0 Then
    MsgBox "El n�mero de oficio de sanci�n es requerido. Favor de capturarlo", vbOKOnly + vbInformation
    txtcampo(0).SetFocus
    Exit Sub
End If
adors.Open "select count(*) from seguimientosanci�n where oficio='" & Replace(txtcampo(0).Text, "'", "''") & "' and idseg<>" & mlSeguimiento, gConSql, adOpenStatic, adLockReadOnly
If adors(0) > 0 Then
    MsgBox "El n�mero de oficio de sanci�n ya existe. Favor de verificar y cambiar el oficio", vbOKOnly + vbInformation
    txtcampo(0).SetFocus
    Exit Sub
End If
If Not IsDate(txtcampo(1).Text) Then
    MsgBox "Todo los datos son requeridos. Favor de capturar la fecha de infracci�n", vbOKOnly + vbInformation
    txtcampo(1).SetFocus
    Exit Sub
End If
If ComboUnidad.ListIndex < 0 Then
    MsgBox "Todo los datos son requeridos. Favor de capturar la Unidad de la Sanci�n", vbOKOnly + vbInformation
    ComboUnidad.SetFocus
    Exit Sub
End If
If Val(txtcampo(2).Text) = 0 Or Val(txtcampo(3).Text) = 0 Then
    MsgBox "Todo los datos son requeridos. Favor de capturar los montos", vbOKOnly + vbInformation
    If Val(txtcampo(2).Text) = 0 Then
        txtcampo(2).SetFocus
    Else
        txtcampo(3).SetFocus
    End If
    Exit Sub
End If
gs = txtcampo(0).Text & "|" & Format(txtcampo(1).Text, "dd/mm/yyyy") & "|" & ComboUnidad.ItemData(ComboUnidad.ListIndex) & "|" & Val(Replace(txtcampo(2).Text, ",", "")) & "|" & Val(Replace(txtcampo(3).Text, ",", "")) & "|"
bAceptar = True
Unload Me
End Sub

Private Sub ComboUnidad_Click()
HabilitaAceptar
If Val(txtcampo(2).Text) > 0 Then ValidaMonto
End Sub

Private Sub Form_Activate()
Dim Y As Byte, s As String, s1 As String, i As Integer
On Error GoTo salir:
LlenaCombo ComboUnidad, "select id, descripci�n from unidadmonetaria where fechabaja is null", "", True
mlAn�lisis = Val(gs1)
mlSeguimiento = Val(gs2)
s = gs
Y = 0
Do While InStr(s, "|")
    s1 = Mid(s, 1, InStr(s, "|") - 1)
    If Y = 2 Then
        ComboUnidad.ListIndex = BuscaCombo(ComboUnidad, Val(s1), True)
    ElseIf Y > 1 Then
        txtcampo(Y - 1).Text = s1
    Else
        txtcampo(Y).Text = s1
    End If
    s = Mid(s, InStr(s, "|") + 1)
    Y = Y + 1
Loop
If Len(Trim(txtcampo(0).Text)) = 0 Then
    Dim adors As New ADODB.Recordset
    adors.Open "select f_nuevofolio(4,0," & mlAn�lisis & ") from dual", gConSql, adOpenStatic, adLockReadOnly
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
salir:
End Sub

Private Sub txtCampo_Change(Index As Integer)
Dim Y As Byte
HabilitaAceptar
End Sub

Private Sub txtCampo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Static i As Integer
If KeyCode = 27 And Index = 3 Then
    txtcampo(Index) = ""
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
    KeyAscii = TeclaOprimida(txtcampo(Index), KeyAscii, txtcampo(Index).Tag, yTipoAcci�n = 3)
End Sub

Private Sub txtCampo_LostFocus(Index As Integer)
Static bValida As Boolean, d As Date, rs As DAO.Recordset, adors As ADODB.Recordset
If txtcampo(Index).Tag = "f" Then
    If IsDate(txtcampo(Index)) Then
        txtcampo(Index) = Format(CDate(txtcampo(Index)), gsFormatoFecha)
        If Index = 1 And IsDate(txtcampo(1)) Then
            d = AhoraServidor
            If CDate(txtcampo(Index).Text) - d > 0 Then
                MsgBox etiTexto(1).Caption & " no puede ser mayor a la fecha de hoy ", vbCritical + vbOKOnly, "Validaci�n de captura"
                If Not bValida Then
                    txtcampo(Index).SetFocus
                    bValida = True
                    Exit Sub
                End If
            End If
            If mdFechaOficio < CDate(txtcampo(Index).Text) Then
                MsgBox etiTexto(Index).Caption & " no puede ser mayor que la fecha de la sanci�n", vbCritical + vbOKOnly, "Validaci�n de captura"
                If Not bValida Then
                    txtcampo(Index).SetFocus
                    bValida = True
                    Exit Sub
                End If
            End If
        End If
    ElseIf Len(Trim(txtcampo(Index))) > 0 Then
        MsgBox "Fecha inv�lida. Favor de Corregirla", 0, "Validaci�n de captura"
        If Not bValida Then
            txtcampo(Index).SetFocus
            bValida = True
            Exit Sub
        End If
    End If
End If
If Index = 2 And Val(txtcampo(2).Text) > 0 Then
    If ComboUnidad.ListIndex > 0 Then
        ValidaMonto
    End If
    If ComboUnidad.ListIndex >= 0 Then
        If ComboUnidad.ItemData(ComboUnidad.ListIndex) = 1 Then
            txtcampo(3).Text = txtcampo(2).Text
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
If IsDate(txtcampo(1).Text) And Val(Replace(txtcampo(2).Text, ",", "")) > 0 And Val(Replace(txtcampo(3).Text, ",", "")) > 0 And ComboUnidad.ListIndex >= 0 Or myAcci�n = 0 Then
    cmdBot�n(0).Enabled = True
ElseIf cmdBot�n(0).Enabled Then
    cmdBot�n(0).Enabled = False
End If
End Sub

Private Sub ValidaMonto()
Dim i As Integer, adors As New ADODB.Recordset
If ComboUnidad.ListIndex >= 0 And Val(txtcampo(2).Text) Then
    i = ComboUnidad.ItemData(ComboUnidad.ListIndex)
    adors.Open "select max_sanci�n from unidadmonetaria where id=" & i, gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        If Not IsNull(adors(0)) Then
            If adors(0) > 0 Then
                If Val(txtcampo(2).Text) > adors(0) Then
                    MsgBox "El monto de " & ComboUnidad.Text & " no puede ser mayor a " & adors(0), vbInformation + vbOKOnly, "Validaci�n"
                    txtcampo(2).Text = adors(0)
                    Exit Sub
                End If
            End If
        End If
    End If
End If
End Sub
