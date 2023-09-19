VERSION 5.00
Begin VB.Form frmProgramaActividad 
   Caption         =   "Programación de la actividad"
   ClientHeight    =   2265
   ClientLeft      =   1620
   ClientTop       =   1485
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Height          =   2235
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4905
      Begin VB.TextBox txtEtiqueta 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Height          =   465
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "frmProgF.frx":0000
         Top             =   1320
         Visible         =   0   'False
         Width           =   2500
      End
      Begin VB.TextBox txtCampo 
         DataField       =   "fecha"
         DataSource      =   "datEncuestas"
         Height          =   315
         Index           =   2
         Left            =   3750
         MaxLength       =   5
         TabIndex        =   4
         Tag             =   "h"
         ToolTipText     =   "Hora programada hh:mm"
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   1260
         Width           =   4575
      End
      Begin VB.OptionButton opcTipoDía 
         Caption         =   "Días hábiles"
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   630
         Width           =   1305
      End
      Begin VB.OptionButton opcTipoDía 
         Caption         =   "Días naturales"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   300
         Width           =   1395
      End
      Begin VB.CommandButton cmdBotón 
         Caption         =   "&Aceptar"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   990
         TabIndex        =   6
         Top             =   1680
         Width           =   1305
      End
      Begin VB.CommandButton cmdBotón 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   2790
         TabIndex        =   7
         Top             =   1680
         Width           =   1305
      End
      Begin VB.TextBox txtCampo 
         BackColor       =   &H00C0C0C0&
         DataField       =   "fechaplaneada"
         DataSource      =   "datEncuestas"
         Height          =   315
         Index           =   0
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   2
         Tag             =   "n"
         Top             =   600
         Width           =   405
      End
      Begin VB.TextBox txtCampo 
         DataField       =   "fecha"
         DataSource      =   "datEncuestas"
         Height          =   315
         Index           =   1
         Left            =   2400
         TabIndex        =   3
         Tag             =   "f"
         ToolTipText     =   "Fecha programada dd/mmm/aaaa"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Responsable:"
         Height          =   285
         Left            =   180
         TabIndex        =   11
         Top             =   990
         Width           =   1905
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Número días:"
         Height          =   405
         Index           =   0
         Left            =   1650
         TabIndex        =   10
         Top             =   180
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha y hora programada:"
         Height          =   225
         Index           =   1
         Left            =   2520
         TabIndex        =   9
         Top             =   300
         Width           =   1965
      End
   End
End
Attribute VB_Name = "frmProgramaActividad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const lMinutosAudiencia = 60
Public dInicio As Date
Public iResponsable As Integer
Public bDíasNaturales As Boolean
Public sProgramada As String
Public iDías As Integer
Dim lSegundos As Long 'Segundos transcurridos para mostrar la etiqueta
Dim iTarea As Integer, iActividad As Integer
Public iPlazoMínimo As Integer
Public iPlazoMáximo As Integer
Public iPlazoEstandar As Integer

Private Sub cmdBotón_Click(Index As Integer)
Dim s As String, Y As Byte, d As Date
If Index = 0 Then
    If InStr(txtCampo(1).Text, " ") Then
        txtCampo(1).Text = Mid(txtCampo(1), 1, InStr(txtCampo(1).Text, " ") - 1)
    End If
    If CDate(txtCampo(1) + " " + txtCampo(2)) < dInicio Then
        MsgBox "La fecha programada no puede ser menor a la fecha de inico de la actividad que le precede", vbOKOnly + vbCritical, "Validación de captura"
        Exit Sub
    End If
    If iPlazoMínimo > 0 Then
        If bDíasNaturales Then
            d = dInicio + iPlazoMínimo
        Else
            d = DíasHábiles(dInicio, iPlazoMínimo)
        End If
        If DateDiff("d", CDate(txtCampo(1) + " " + txtCampo(2)), d) > 0 Then s = "La fecha programada es menor a la fecha correspondiente según el plazo mínimo definido (" + Format(d, gsFormatoFecha) + ")"
    End If
    If iPlazoMáximo > 0 Then
        If bDíasNaturales Then
            d = dInicio + iPlazoMáximo
        Else
            d = DíasHábiles(dInicio, iPlazoMáximo)
        End If
        If DateDiff("d", d, CDate(txtCampo(1) + " " + txtCampo(2))) > 0 Then s = "La fecha programada es mayor a la fecha correspondiente según el plazo máximo definido (" + Format(d, gsFormatoFecha) + ")"
    End If
    If Not CompruebaPrograma Then Exit Sub
    If Len(s) > 0 Then
        If MsgBox(s + " ¿Está seguro de continuar?", vbYesNo + vbQuestion, "Confirmación") = vbNo Then Exit Sub
    End If
    gs = Format(CDate(txtCampo(1).Text + " " + txtCampo(2)), "dd/mm/yyyy hh:mm") + "|" + Trim(Str(Combo1.ItemData(Combo1.ListIndex)))
Else
    gs = ""
End If
On Error GoTo OcultaFormulario:
Unload Me
Exit Sub
OcultaFormulario:
Me.Tag = "Inválido"
Me.Hide
End Sub

Private Sub Combo1_Click()
If txtEtiqueta.Visible Then txtEtiqueta.Visible = False
lSegundos = -1
ProgramaFecha
End Sub

Private Sub Form_Activate()
Dim i As Integer, adors As New ADODB.Recordset
opcTipoDía(IIf(bDíasNaturales, 0, 1)).Value = True
txtCampo(0) = IIf(iPlazoEstandar < 0, IIf(iPlazoMáximo < 0, IIf(iPlazoMínimo < 0, "", iPlazoMínimo), iPlazoMáximo), iPlazoEstandar)
Combo1.ListIndex = BuscaCombo(Combo1, Trim(Str(iResponsable)), True)
'txtCampo(1).SetFocus
If IsDate(sProgramada) Then
    txtCampo(1) = Format(CDate(sProgramada), gsFormatoFecha)
    txtCampo(2) = Format(CDate(sProgramada), "hh:mm")
Else
    If adors.State Then adors.Close
    adors.Open "select diasprog, diash from relacióntareaactividad where idtar=" & iTarea & " and idact=" & iActividad, gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        txtCampo(0).Text = IIf(IsNull(adors(0)), 0, adors(1))
        If adors(1) <> 0 Then
            opcTipoDía(1).Value = True
        Else
            opcTipoDía(0).Value = True
        End If
    End If
    txtCampo(2).Text = ""
End If
If IsDate(txtCampo(0).Text) Then
    If Year(txtCampo(0).Text) < 2000 Then
        txtCampo(0).Text = ""
    End If
End If
End Sub

Private Sub Form_Load()
Dim s As String
Call LlenaComboClave(Combo1, "USUARIOSSISTEMA", "responsable<>0 and baja=0")
'dInicio = Null
iTarea = gi1
iActividad = gi2
If Not gs = "no iniciar var" Then
    iPlazoEstandar = 0
    iPlazoMáximo = 0
    iPlazoMínimo = 0
    sProgramada = ""
    iResponsable = 0
    bDíasNaturales = False
    iActividad = 0
    iDías = 0
End If
End Sub

Private Sub Frame_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If txtEtiqueta.Visible Then txtEtiqueta.Visible = False
lSegundos = -1
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim aValor(6) As Integer
'Call MuestraEtiqueta(Label1(Index), txtEtiqueta, 0, lSegundos, aValor)
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim aValor(6) As Integer
'Call MuestraEtiqueta(Label2, txtEtiqueta, 0, lSegundos, aValor)
End Sub

Private Sub opcTipoDía_Click(Index As Integer)
ProgramaFecha
If txtCampo(0).Visible And txtCampo(0).Enabled Then txtCampo(0).SetFocus
End Sub

Private Sub opcTipoDía_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim aValor(6) As Integer
'Call MuestraEtiqueta(opcTipoDía(Index), txtEtiqueta, 0, lSegundos, aValor)
End Sub

Private Sub txtCampo_Change(Index As Integer)
If Index = 0 Then
    ProgramaFecha
Else
    cmdBotón(0).Enabled = IsDate(txtCampo(1)) And IsDate(txtCampo(2)) And Combo1.ListIndex >= 0
End If
End Sub

Private Sub ProgramaFecha()
If Len(Trim(txtCampo(0))) > 0 Then
    If opcTipoDía(0) Then
        txtCampo(1).Text = Format(dInicio + Val(txtCampo(0)), gsFormatoFecha)
        If InStr(Me.Caption, "07 Audiencia") = 0 Then txtCampo(2) = Format(DateAdd("n", 1, dInicio + Val(txtCampo(0))), "hh:mm")
    Else
        txtCampo(1).Text = Format(DíasHábiles(dInicio, Val(txtCampo(0))), gsFormatoFecha)
        If InStr(Me.Caption, "07 Audiencia") = 0 Then txtCampo(2) = Format(DateAdd("n", 1, dInicio + Val(txtCampo(0))), "hh:mm")
    End If
Else
    txtCampo(1) = ""
End If
cmdBotón(0).Enabled = IsDate(txtCampo(1)) And IsDate(txtCampo(2)) And Combo1.ListIndex >= 0
End Sub

Private Sub txtcampo_GotFocus(Index As Integer)
If txtEtiqueta.Visible Then txtEtiqueta.Visible = False
lSegundos = -1
End Sub

Private Sub txtCampo_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = TeclaOprimida(txtCampo(Index), KeyAscii, txtCampo(Index).Tag, False)
End Sub

Function CompruebaPrograma() As Boolean
Dim rs As Recordset, adors As New ADODB.Recordset, s As String
If Len(s) > 0 Then
    If MsgBox(s, vbYesNo + vbQuestion + vbDefaultButton2, "") = vbNo Then
        CompruebaPrograma = False
        Exit Function
    End If
End If
CompruebaPrograma = True
End Function

Private Sub txtCampo_LostFocus(Index As Integer)
If Index = 1 Then
    If IsDate(txtCampo(Index)) Then
        If Hour(CDate(txtCampo(Index))) = 0 Then
            txtCampo(Index) = Format(CDate(txtCampo(Index)), gsFormatoFecha)
         End If
    End If
End If
End Sub
