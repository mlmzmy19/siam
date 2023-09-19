VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form SelProceso 
   Caption         =   "Selección de actividad"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12480
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   12480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelTodos 
      Caption         =   "&Quitar Selección a Todos"
      Height          =   580
      Index           =   1
      Left            =   10800
      TabIndex        =   4
      Top             =   7155
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelTodos 
      Caption         =   "Seleccionar &Todos"
      Height          =   580
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   7110
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2475
      Top             =   8085
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6810
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   12012
      _Version        =   393217
      HideSelection   =   0   'False
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   1305
      TabIndex        =   5
      Top             =   6975
      Width           =   9210
      Begin VB.CommandButton cmdAcción 
         Caption         =   "C&ontraer Nodos"
         Height          =   300
         Index           =   4
         Left            =   7320
         TabIndex        =   14
         Top             =   450
         Width           =   1440
      End
      Begin VB.CommandButton cmdAcción 
         Caption         =   "&Expandir Nodos"
         Height          =   300
         Index           =   3
         Left            =   7320
         TabIndex        =   13
         Top             =   135
         Width           =   1440
      End
      Begin VB.CommandButton cmdAcción 
         Caption         =   "&Sol. Inf. Usuario"
         Height          =   435
         Index           =   2
         Left            =   3600
         TabIndex        =   12
         Top             =   225
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.CommandButton cmdAcción 
         Caption         =   "&Cancelar"
         Height          =   435
         Index           =   1
         Left            =   5280
         TabIndex        =   3
         Top             =   240
         Width           =   1170
      End
      Begin VB.CommandButton cmdAcción 
         Caption         =   "&Aceptar"
         Height          =   435
         Index           =   0
         Left            =   2040
         TabIndex        =   2
         Top             =   225
         Width           =   1215
      End
   End
   Begin VB.TextBox txtComentarios 
      Height          =   1365
      Left            =   45
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   4860
      Visible         =   0   'False
      Width           =   6765
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   510
      Left            =   45
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   4200
      Begin VB.OptionButton OpcCompleta 
         Caption         =   "Información Incompleta"
         Height          =   330
         Index           =   1
         Left            =   2025
         TabIndex        =   10
         Top             =   135
         Width           =   2040
      End
      Begin VB.OptionButton OpcCompleta 
         Caption         =   "Información Completa"
         Height          =   330
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   135
         Value           =   -1  'True
         Width           =   1905
      End
   End
   Begin VB.CheckBox chkInfAdi 
      Caption         =   "Solicitar Inf. Adi. al Usuario"
      Height          =   375
      Left            =   4410
      TabIndex        =   11
      Top             =   4050
      Width           =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "Comentarios o documentos adicionales a la UNE:"
      Height          =   330
      Left            =   90
      TabIndex        =   7
      Top             =   4590
      Visible         =   0   'False
      Width           =   6675
   End
End
Attribute VB_Name = "SelProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public yTipoSel As Byte 'indica si se puede seleccionar con doble click
Dim bVariosNiveles As Boolean 'Se utiliza para varias instituciones únicamente
Dim iIns As Long 'Instprincipal  de las varias inst
Dim iCla As Long 'Clase principal
Dim iIndIns As Long
Dim iIndCla As Long
Dim lValor As Long
Dim sQuitaNodos As String
Dim sActuales As String
Dim bNoNuevo As Boolean
Dim bVerifica As Boolean
Dim bUnes As Boolean
Dim miCambio As Integer 'indica si se a realizado algún cambio
Public psSeleccionados As String
Public piSel As Long 'Valor del elemento seleccionado
Public psSel As String 'Contiene las claves del elemento seleccionado
Public piValida As Integer 'Indicador si valida que el elemento seleccionado sea del último nivel desagregado

Private Sub chkInfAdi_Click()
If chkInfAdi.Value Then
    TreeView1.Visible = False
    Frame2.Visible = False
    txtComentarios.Visible = False
    Label1.Visible = False
Else
    TreeView1.Visible = True
    Frame2.Visible = True
    txtComentarios.Visible = True
    Label1.Visible = True
End If
End Sub

Sub cmdAcción_Click(Index As Integer)
Dim i As Long
If Index = 3 Or Index = 4 Then 'Expandir/Contraer Nodos
    For i = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(i).Children > 0 Then
            TreeView1.Nodes(i).Expanded = (Index = 3)
        End If
    Next
    Exit Sub
End If
glProceso = 0
gs = ""
If bUnes Then gs1 = ""
gs2 = ""
gs3 = ""
gs1 = "" 'Obtiene los key de los nodos seleccionados
If Index = 0 Then 'Aceptar
    If miCambio = 0 Then
        MsgBox "No se realizaron cambios", vbOKOnly + vbInformation, ""
        cmdAcción_Click 1
        Exit Sub
    End If
    If iIndIns > 0 And iIndIns <= TreeView1.Nodes.Count Then
        If TreeView1.Nodes(iIndIns).Checked Then
            gs = Right(TreeView1.Nodes(iIndIns).Key, IIf(bVariosNiveles, 10, 3)) + ","
            gs3 = Mid(Right(TreeView1.Nodes(iIndIns).Key, IIf(bVariosNiveles, 20, 6)), 1, IIf(bVariosNiveles, 10, 3)) + ","
            gs2 = Trim(TreeView1.Nodes(iIndIns).Text) + ", "
        End If
    End If
    If TreeView1.CheckBoxes Then
        For i = 1 To TreeView1.Nodes.Count
            If TreeView1.Nodes(i).Children = 0 Then
                If TreeView1.Nodes(i).Checked And i <> iIndIns Then
                    gs = gs & Val(Right(TreeView1.Nodes(i).Key, IIf(bVariosNiveles, 10, 3))) & ","
                    gs3 = gs3 & Val(Mid(Right(TreeView1.Nodes(i).Key, IIf(bVariosNiveles, 20, 6)), 1, IIf(bVariosNiveles, 10, 3))) & ","
                    gs2 = gs2 + Trim(TreeView1.Nodes(i).Text) + ", "
                    gs1 = gs1 & Mid(TreeView1.Nodes(i).Key, 2) & "|"
                ElseIf TreeView1.Nodes(i).Checked Then
                    gs1 = gs1 & Mid(TreeView1.Nodes(i).Key, 2) & "|"
                End If
            Else
                If TreeView1.Nodes(i).Checked Then
                    iCla = Val(Right(TreeView1.Nodes(i).Key, IIf(bVariosNiveles, 10, 3)))
                    iIndCla = i
                End If
            End If
        Next
        If Len(gs) = 0 Then gs = "nada"
        If bUnes Then gs1 = txtComentarios
    Else
        gs = lValor
        gs1 = psSel
    End If
    If bUnes And chkInfAdi.Value Then gs = "solinf"
ElseIf Index = 2 Then
    gs = "solinf"
End If
Unload Me
End Sub

Private Sub cmdSelTodos_Click(Index As Integer)
Dim b As Boolean, i As Integer
b = Index = 0
'If Index = 0 Then
    For i = 1 To TreeView1.Nodes.Count
        'If TreeView1.Nodes(i).Children = 0 And TreeView1.Nodes(i).Checked <> b Then
            TreeView1.Nodes(i).Checked = b
        'End If
    Next
'End If
End Sub

Private Sub Form_Activate()
Dim Y As Byte, nodo As MSComctlLib.Node
Dim n As Integer
On Error GoTo ErrorAbrirArbol:
If bVariosNiveles Then 'Varios niveles desagregación
    iIns = IIf(Val(gs2) > 0, Val(gs2), -1)
    iCla = IIf(Val(Mid(gs2, InStr(gs2, ",") + 1)), Val(Mid(gs2, InStr(gs2, ",") + 1)), -1)
    sAnteriores = gs1
    sActuales = sAnteriores
End If
If gs2 Like "-->*" Then 'obtiene el dato de la llave con dato
    n = 200
    If TreeView1.CheckBoxes Then
        For i = 1 To TreeView1.Nodes.Count
            If InStr("|" & Mid(gs2, 4), "|" & Mid(TreeView1.Nodes(i).Key, 2) & "|") > 0 Then
                TreeView1.Nodes(i).Checked = True
                TreeView1.Nodes(i).Selected = True
                
                'Exit For
            End If
        Next
    Else
        For i = 1 To TreeView1.Nodes.Count
            If TreeView1.Nodes(i).Key = "r" & Mid(gs2, 4) Then
                TreeView1.Nodes(i).Selected = True
                Exit For
            End If
        Next
    End If
ElseIf InStr(gs4, ",") > 0 Then
    For i = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(i).Children = 0 Then
            If InStr("," & gs4, "," & Val(Right(TreeView1.Nodes(i).Key, IIf(bVariosNiveles, 10, 3))) & ",") > 0 Then
                TreeView1.Nodes(i).Checked = True
                If TreeView1.Nodes(i).Children Then TreeView1.Nodes(i).Parent.Expanded = True
                iIndIns = i
            End If
        Else
            If Val(Right(TreeView1.Nodes(i).Key, IIf(bVariosNiveles, 10, 3))) = iCla Then
                TreeView1.Nodes(i).Checked = True
                TreeView1.Nodes(i).Expanded = True
                iIndCla = i
            End If
        End If
    Next
End If
If gs3 Like "comentarios:*" Then
    Frame2.Visible = True
    txtComentarios.Visible = True
    txtComentarios.Text = Mid(gs3, 13)
    Label1.Visible = True
    TreeView1.Height = Frame2.Top - TreeView1.Top - 50
    bUnes = True
    If TreeView1.Nodes.Count > 1 Then
        Set nodo = TreeView1.Nodes(2)
        Call TreeView1_NodeCheck(nodo)
    End If
ElseIf gs3 Like "expandir todos*" Then
    For i = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(i).Children > 0 Then
            TreeView1.Nodes(i).Expanded = True
        End If
    Next
End If
If TreeView1.Nodes.Count > 0 And Not TreeView1.CheckBoxes Then
    If n = 0 Then
        TreeView1_NodeClick TreeView1.Nodes(1)
    End If
End If
Exit Sub
ErrorAbrirArbol:
If Err.Number = 2013 Then Resume Next
Y = MsgBox("Error: " + Err.Description, vbAbortRetryIgnore + vbCritical, "Error no esperado (" + Str(Err.Number) + ")")
If Y = vbRetry Then
    Resume
ElseIf Y = vbIgnore Then
    Resume Next
End If
End Sub

Private Sub Form_Load()
Dim s As String
bVariosNiveles = False
bUnes = False
piValida = 0
piSel = 0
psSel = ""
miCambio = 0
If gs Like "{*}" Then
    s = gs
    gs = ""
    s = LCase(s)
    Call CargaDatosArbolVariosNiveles10(TreeView1, s, 0, False, , Val(gs4) <> 200, True)
    bVariosNiveles = True
ElseIf gs Like "ArbolVarios-->*" Then
    s = Mid(gs, 15)
    gs = ""
    s = LCase(s)
    Call CargaDatosArbolVariosNiveles10(TreeView1, s, IIf(InStr("FoliosAsuntos,UsuariosSistema,", Trim(gs2)) > 0 And Len(gs2) > 0, 1, 2), False, , False, gs3 <> "varios")
    bVariosNiveles = True
ElseIf gs Like "-->*" Then
    s = Mid(gs, 4)
    gs = ""
    If InStr(LCase(s), "relación") And InStr(LCase(s), "docsol") Then
        Call CargaDatosArbol(TreeView1, s, False, 0)
    Else
        Call CargaDatosArbol(TreeView1, s)
    End If
Else
    s = "SELECT a.id,b.id,c.id,d.id,a.descripción,b.descripción,c.descripción,d.descripción FROM ((actividades AS a LEFT JOIN actividades AS b ON a.id=b.idpad) LEFT JOIN actividades AS c ON b.id=c.idpad) LEFT JOIN actividades AS d ON c.id=d.idpad WHERE a.nivel=1 and (b.nivel=2 or b.nivel is null) and (c.nivel=3 or c.nivel is null) ORDER BY a.descripción,b.descripción,c.descripción,d.descripción"
    Call CargaDatosArbol(TreeView1, s)
End If
If TreeView1.CheckBoxes Then
    cmdSelTodos(0).Visible = True
    cmdSelTodos(1).Visible = True
End If
'For i = 1 To TreeView1.Nodes.Count
'    If TreeView1.Nodes(i).Children > 0 Then TreeView1.Nodes(i).Expanded = True
'Next
'TreeView1.Style = tvwTextOnly
End Sub

Private Sub Form_Resize()
If Me.Width > 7095 Then
    TreeView1.Width = Me.Width - 300
Else
    If TreeView1.Width <> 7095 Then
        TreeView1.Width = 7095
    End If
End If
End Sub

Private Sub OpcCompleta_Click(Index As Integer)
If OpcCompleta(0).Value Then
    Label1.Caption = "Comentarios o documentos adicionales a la UNE:"
    Label1.FontBold = False
    cmdAcción(0).Enabled = True
Else
    cmdAcción(0).Enabled = (Len(Trim(txtComentarios.Text)) > 0)
    Label1.Caption = "Razón por la cual no se manda completa la Información:"
    Label1.FontBold = True
End If
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
Cancel = 1
End Sub

Private Sub TreeView1_DblClick()
If TreeView1.CheckBoxes Then Exit Sub
If piValida > 0 Then 'Valida que sea de último nivel
    If TreeView1.SelectedItem.Children > 0 Then 'No se toma como bueno
        MsgBox "Debe seleccionar un elemento de último nivel del árbol", vbInformation + vbOKOnly, ""
        Exit Sub
    End If
End If
glProceso = Val(Right(TreeView1.SelectedItem.Key, 3))
If Not TreeView1.CheckBoxes Then
    miCambio = 1
End If
If glProceso <> 0 Then
    cmdAcción_Click (0)
End If
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim i As Long, nodo As MSComctlLib.Node
Dim i1 As Integer
Dim b As Boolean
On Error GoTo salir
If Node.Children > 0 Then
    b = Node.Checked
    Set nodo = Node.Child
    i1 = Node.Children
    For i = 1 To i1
        If nodo.Checked <> b Then
            nodo.Checked = b
            If nodo.Children > 0 Then
                TreeView1_NodeCheck nodo
            End If
        End If
        Set nodo = nodo.Next
    Next
Else
    If Val(Right(Node.Key, 10)) = iIns And Not Node.Checked Then
        iIns = -1
        iIndIns = -1
    ElseIf Node.Checked And iIns = -1 Then
        iIns = Val(Right(Node.Key, 10))
        iIndIns = Node.Index
    End If
End If
miCambio = 1
salir:
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
Dim i As Integer, iNodo As Integer
lValor = Val(Right(Node.Key, IIf(bVariosNiveles, 10, 3)))
psSel = Mid(Node.Key, 2)
If Not TreeView1.CheckBoxes Then
    miCambio = 1
End If
End Sub

Private Sub txtComentarios_Change()
Dim b As Boolean
b = Len(Trim(txtComentarios.Text))
If b Then
    If Not cmdAcción(0).Enabled Then cmdAcción(0).Enabled = True
Else
    If OpcCompleta(1).Value Then
        If cmdAcción(0).Enabled Then cmdAcción(0).Enabled = False
    Else
        If Not cmdAcción(0).Enabled Then cmdAcción(0).Enabled = True
    End If
End If
End Sub
