VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRefCruz 
   Caption         =   "Informaci�n para generar informe de referencias cruzadas"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   1230
      Left            =   0
      TabIndex        =   24
      Top             =   45
      Width           =   8970
      Begin VB.OptionButton opcRango 
         Caption         =   "Anual"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   29
         Top             =   270
         Width           =   780
      End
      Begin VB.OptionButton opcRango 
         Caption         =   "Bimestral"
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   28
         Top             =   495
         Width           =   1050
      End
      Begin VB.OptionButton opcRango 
         Caption         =   "Mensual"
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   27
         Top             =   855
         Width           =   1005
      End
      Begin VB.OptionButton opcRango 
         Caption         =   "Semanal"
         Height          =   285
         Index           =   3
         Left            =   1710
         TabIndex        =   26
         Top             =   405
         Width           =   1095
      End
      Begin VB.OptionButton opcRango 
         Caption         =   "Otro"
         Height          =   285
         Index           =   4
         Left            =   1710
         TabIndex        =   25
         Top             =   765
         Width           =   915
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   330
         Index           =   0
         Left            =   3780
         TabIndex        =   30
         Top             =   270
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   582
         _Version        =   393216
         Format          =   102760448
         CurrentDate     =   37739
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   825
         Left            =   8505
         TabIndex        =   31
         Top             =   225
         Width           =   270
         _ExtentX        =   450
         _ExtentY        =   1455
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   330
         Index           =   1
         Left            =   3780
         TabIndex        =   32
         Top             =   720
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   582
         _Version        =   393216
         Format          =   102760448
         CurrentDate     =   37743
      End
      Begin VB.Label Label3 
         Caption         =   "Del:"
         Height          =   240
         Index           =   0
         Left            =   3330
         TabIndex        =   7
         Top             =   315
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Al:"
         Height          =   240
         Index           =   1
         Left            =   3420
         TabIndex        =   33
         Top             =   810
         Width           =   285
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5505
      Left            =   0
      TabIndex        =   3
      Top             =   1305
      Width           =   8970
      Begin VB.TextBox txtSubt�tulo 
         Height          =   1005
         Left            =   210
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   4320
         Width           =   7305
      End
      Begin VB.TextBox txtT�tulo 
         Height          =   375
         Left            =   210
         MaxLength       =   250
         TabIndex        =   21
         Top             =   3810
         Width           =   7275
      End
      Begin VB.Frame Frame3 
         Height          =   3675
         Left            =   7560
         TabIndex        =   17
         Top             =   1200
         Width           =   1290
         Begin VB.CommandButton cmdBot�n 
            Caption         =   "&Procesa informe"
            Enabled         =   0   'False
            Height          =   555
            Index           =   0
            Left            =   150
            TabIndex        =   20
            Top             =   360
            Width           =   1000
         End
         Begin VB.CommandButton cmdBot�n 
            Caption         =   "&Salir"
            Height          =   405
            Index           =   1
            Left            =   135
            TabIndex        =   19
            Top             =   2565
            Width           =   1000
         End
         Begin VB.CommandButton cmdBot�n 
            Caption         =   "&Exportar a &Excel"
            Enabled         =   0   'False
            Height          =   1125
            Index           =   2
            Left            =   150
            Picture         =   "R.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1140
            Width           =   1035
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
         Begin VB.TextBox txtRengl�n 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   1035
            Width           =   2200
         End
         Begin VB.TextBox txtRengl�n 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   765
            Width           =   2200
         End
         Begin VB.TextBox txtRengl�n 
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
         ItemData        =   "R.frx":0442
         Left            =   3960
         List            =   "R.frx":0444
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "n"
         ToolTipText     =   "Valor que desea como t�tulo en las columnas"
         Top             =   495
         Width           =   3400
      End
      Begin VB.ListBox List1 
         Height          =   840
         Index           =   0
         ItemData        =   "R.frx":0446
         Left            =   300
         List            =   "R.frx":0450
         TabIndex        =   2
         Top             =   1110
         Width           =   3400
      End
      Begin VB.ComboBox ComboVarios 
         DataField       =   "idtip"
         DataSource      =   "datAsunto"
         Height          =   315
         Index           =   0
         ItemData        =   "R.frx":0467
         Left            =   270
         List            =   "R.frx":04B9
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "n"
         ToolTipText     =   "Valor que desea como t�tulo en los renglones"
         Top             =   480
         Width           =   3400
      End
      Begin Crystal.CrystalReport CReport 
         Left            =   7830
         Top             =   135
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label1 
         Caption         =   "T�tulo y subt�tulo del informe:"
         Height          =   225
         Left            =   240
         TabIndex        =   23
         Top             =   3600
         Width           =   2385
      End
      Begin VB.Label EtiCombo 
         Caption         =   "Valor que se desea como t�tulo de la columna:"
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
         Caption         =   "Valor que se desea como t�tulo de la fila:"
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
Const ciDelegaci�n = 2
Dim sQueryPrin As String
Dim db As DAO.Database

Private Sub cmdBot�n_Click(Index As Integer)
Dim s As String, sFrom As String, rs As Recordset, ss As String, s1 As String, s2 As String, s3 As String, yCampos As Integer, y1 As Byte, y2 As Byte, rs1 As Recordset
Dim s_s(3) As String, s_ss(3) As String, s_sfrom(3) As String, b_Otros(1) As Boolean, s_gs(3) As String
Dim adors As New ADODB.Recordset
Dim sW1 As String
Dim qdef As QueryDef
Dim Hoja As Excel.Worksheet
Dim LibroExcel As Excel.Workbook
Dim ApExcel As Excel.Application

Me.MousePointer = 11

If Index = 1 Then 'SALIR
    Unload Me
    Exit Sub
End If

'Verifica si se manda informaci�n del Rengl�n y Columna para el informe de Ref Cruzadas
If ComboVarios(0).ListIndex < 0 Then
    MsgBox "Se requiere especificar la variable del Rengl�n para el informe", vbInformation + vbOKOnly, "Validaci�n"
    Exit Sub
End If
If ComboVarios(1).ListIndex < 0 Then
    MsgBox "Se requiere especificar la variable de la Columna para el informe", vbInformation + vbOKOnly, "Validaci�n"
    Exit Sub
End If
If Frame4.Visible Then
    If DTP(0).Value > DTP(1).Value Then
        MsgBox "El rango de fechas es incorrecto", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
End If
'Asigna nombre del informe (archivo.rpt)
CReport.ReportFileName = gsDirReportes & "\Reporte referencia cruzadas.rpt"
'asigna el valor del par�metro en caso de haber
'If Frame3.Visible Then
'    CReport.ParameterFields(0) = msPar�metro & ";" & ComboVarios.ItemData(ComboVarios.ListIndex) & ";true"
'Else
'    CReport.ParameterFields(0) = ""
'End If
'asigna el rango de fechas en caso de haber
If Frame4.Visible Then
    CReport.ParameterFields(1) = "psInicio;" & Format(DTP(0).Value, "dd/mm/yyyy") & ";true"
    CReport.ParameterFields(2) = "psTermino;" & Format(DTP(1).Value, "dd/mm/yyyy") & ";true"
Else
    CReport.ParameterFields(1) = ""
    CReport.ParameterFields(2) = ""
End If
CReport.ParameterFields(3) = "piRow;" & ComboVarios(0).ItemData(ComboVarios(0).ListIndex) & ";true"
CReport.ParameterFields(4) = "piCol;" & ComboVarios(1).ItemData(ComboVarios(1).ListIndex) & ";true"

'Asigna la conexi�n
CReport.Connect = gConSql.ConnectionString '& ";dsn=siam"
CReport.Connect = "FILEDSN=c:\siam\siam.dsn;pwd=siam_desa"

CReport.Action = 1

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
    cmdBot�n(0).Enabled = False
    cmdBot�n(2).Enabled = False
    ComboVarios(1).Clear
    For Y = 0 To ComboVarios(0).ListCount - 1
        If Y <> ComboVarios(0).ListIndex Then
            ComboVarios(1).AddItem ComboVarios(0).List(Y)
            ComboVarios(1).ItemData(ComboVarios(1).ListCount - 1) = ComboVarios(0).ItemData(Y)
        End If
    Next
    txtColumna(0) = ComboVarios(0).Text
    For Y = 0 To 2
        txtRengl�n(Y) = ComboVarios(0).Text + Str(Y + 1)
    Next
    ComboVarios(1).ListIndex = -1
Else
    For Y = 1 To 3
        txtColumna(Y) = ComboVarios(1).Text + Str(Y)
    Next
    txtT�tulo = "Informe de referencias cruzadas (" + ComboVarios(0).Text + " vs " + ComboVarios(1).Text + ")"
End If
chkOtros(Index).Visible = InStr("Causas**Instituci�n", ComboVarios(Index).Text) Or ComboVarios(Index).Text Like "Producto Nivel*"
For Y = 0 To 1
    If ComboVarios(Y).ListIndex < 0 Then Exit For
Next
cmdBot�n(0).Enabled = Y > 1
cmdBot�n(2).Enabled = Y > 1
End Sub

Private Sub Form_Activate()
Dim i As Integer
List1(0).ListIndex = 0
txtT�tulo = "" '"Informe de referencias cruzadas (" + ComboVarios(0).Text + " vs " + ComboVarios(1).Text + ")"
txtSubt�tulo = gsT�tulo
'If InStr(gsQueryPrin, "select ") = 0 Then
    If gSQLACC = cyAccess Then
        sQueryPrin = "select a.id,a.fecha,a.idrec,a.idsec,a.idcla,a.idpr1,a.idpr2,a.idpr3,a.idcau,a.atenci�n,a.idusi,b.id,b.a�o,b.iddel,b.consecutivo,b.fechahechos,nm.monto,nob.observaciones,b.clase,ni.noidentificaci�n,ni.idide,nm.iduni,null,null,ao.observaciones,e.idprc,e.referencia,e.fechaprocedencia,e.foliooficial�a," + _
                     "g.[n�mero de cuenta],h.[n�mero de contrato],h.[Nombre del promotor],h.[Fecha de operaci�n impugnada],i.[N�mero de Seguridad Social],i.RFC,i.CURP,i.[Lugar de nacimiento],nsb.beneficiarios,j.[N�mero de P�liza],j.[Nombre del Fiado],j.[Nombre del Beneficiario],j.[Contrato origen de la Fianza],j.[Inicio de Vigencia],j.[T�rmino de Vigencia],k.[Inicio de Vigencia],k.[T�rmino de Vigencia],k.[N�mero de P�liza],k.[Nombre del Contratante],k.[Nombre del Asegurado],k.beneficiarios,k.[Nombre del Agente],k.[Nombre del Ajustador],k.[N�mero de Siniestro],k.[Lugar del siniestro],k.[Suma Asegurada]" + _
                     " from ((((((((((((asuntos a left join Nominales as b on a.id=b.idAsu) left join nominalesmontos nm on b.id=nm.idnom) left join nominalesobs as nob on b.id=nob.idnom) left join an�nimos as d on a.id=d.idAsu) left join an�nimosobs as ao on d.id=ao.idanon) left join procedencias as e on a.id=e.idAsu) left join nominalesbancos g on b.id=g.idnom) left join nominalesburs�til h on b.id=h.idnom) left join nominalessar i on b.id=i.idnom) left join nominalesfianzas j on b.id=j.idnom) left join nominalesseguros k on b.id=k.idnom) left join nominalesidentificaci�n ni on b.id=ni.idnom) left join nominalessarben nsb on b.id=nsb.idnom"
    Else
        If gSQLACC = cyOracle Then
            sQueryPrin = "select a.id as a_id,a.fecha,a.idrec,a.idsec,a.idcla,a.idpr1,a.idpr2,a.idpr3,a.idcau,a.atenci�n,a.idusi,b.id as b_id,b.a�o,b.iddel,b.consecutivo,b.fechahechos,nm.monto,nob.observaciones as b_observaciones,b.clase,ni.noidentificaci�n,ni.idide,nm.iduni,null,null,ao.observaciones as d_observaciones,e.idprc,e.referencia,e.fechaprocedencia,e.foliooficial�a,f.iddes,f.idsen,em.montorecuperado,f.idact,f.favorable,f.idins as InstFavo,eo.observaciones as f_observaciones,f.fechaconclusi�n," + _
                     "g.n�mero_de_cuenta,h.n�mero_de_contrato,h.Nombre_del_promotor,h.Fecha_de_operaci�n_impugnada,i.N�mero_de_Seguridad_Social,i.RFC,i.CURP,i.Lugar_de_nacimiento,nsb.beneficiarios as i_beneficiarios,j.N�mero_de_P�liza as J_N�mero_de_P�liza,j.Nombre_del_Fiado,j.Nombre_del_Beneficiario,j.Contrato_origen_de_la_Fianza,j.Inicio_de_Vigencia as J_Inicio_de_Vigencia,j.T�rmino_de_Vigencia as J_T�rmino_de_Vigencia,k.Inicio_de_Vigencia as K_Inicio_de_Vigencia,k.T�rmino_de_Vigencia as K_T�rmino_de_Vigencia,k.N�mero_de_P�liza as K_N�mero_de_P�liza,k.Nombre_del_Contratante,k.Nombre_del_Asegurado,k.beneficiarios as k_beneficiarios,k.Nombre_del_Agente,k.Nombre_del_Ajustador,k.N�mero_de_Siniestro,k.Lugar_del_siniestro,k.Suma_Asegurada" + _
                     " from asuntos a left join Nominales b on a.id=b.idAsu left join nominalesmontos nm on b.id=nm.idnom left join nominalesobs nob on b.id=nob.idnom left join an�nimos d on a.id=d.idAsu left join an�nimosobs ao on d.id=ao.idanon left join procedencias e on a.id=e.idAsu left join evaluaci�n f on a.id=f.idasu left join nominalesbancos g on b.id=g.idnom left join nominalesburs�til h on b.id=h.idnom left join nominalessar i on b.id=i.idnom left join nominalesfianzas j on b.id=j.idnom left join nominalesseguros k on b.id=k.idnom left join nominalesidentificaci�n ni on b.id=ni.idnom left join nominalessarben nsb on b.id=nsb.idnom left join evaluaci�nmontos em on f.id=em.ideva left join evaluaci�nobs eo on f.id=eo.ideva"
        Else
            sQueryPrin = "select a.id as a_id,a.fecha,a.idrec,a.idsec,a.idcla,a.idpr1,a.idpr2,a.idpr3,a.idcau,a.atenci�n,a.idusi,b.id as b_id,b.a�o,b.iddel,b.consecutivo,b.fechahechos,nm.monto,nob.observaciones as b_observaciones,b.clase,ni.noidentificaci�n,ni.idide,nm.iduni,null,null,ao.observaciones as d_observaciones,e.idprc,e.referencia,e.fechaprocedencia,e.foliooficial�a,f.iddes,f.idsen,em.montorecuperado,f.idact,f.favorable,f.idins as InstFavo,eo.observaciones as f_observaciones,f.fechaconclusi�n," + _
                     "g.[n�mero de cuenta],h.[n�mero de contrato],h.[Nombre del promotor],h.[Fecha de operaci�n impugnada],i.[N�mero de Seguridad Social],i.RFC,i.CURP,i.[Lugar de nacimiento],nsb.beneficiarios as i_beneficiarios,j.[N�mero de P�liza] as [J_N�mero de P�liza],j.[Nombre del Fiado],j.[Nombre del Beneficiario],j.[Contrato origen de la Fianza],j.[Inicio de Vigencia] as [J_Inicio de Vigencia],j.[T�rmino de Vigencia] as [J_T�rmino de Vigencia],k.[Inicio de Vigencia] as [K_Inicio de Vigencia],k.[T�rmino de Vigencia] as [K_T�rmino de Vigencia],k.[N�mero de P�liza] as [K_N�mero de P�liza],k.[Nombre del Contratante],k.[Nombre del Asegurado],k.beneficiarios as k_beneficiarios,k.[Nombre del Agente],k.[Nombre del Ajustador],k.[N�mero de Siniestro],k.[Lugar del siniestro],k.[Suma Asegurada]" + _
                     " from asuntos a left join Nominales b on a.id=b.idAsu left join nominalesmontos nm on b.id=nm.idnom left join nominalesobs nob on b.id=nob.idnom left join an�nimos d on a.id=d.idAsu left join an�nimosobs ao on d.id=ao.idanon left join procedencias e on a.id=e.idAsu) left join evaluaci�n f on a.id=f.idasu left join nominalesbancos g on b.id=g.idnom left join nominalesburs�til h on b.id=h.idnom left join nominalessar i on b.id=i.idnom left join nominalesfianzas j on b.id=j.idnom left join nominalesseguros k on b.id=k.idnom left join nominalesidentificaci�n ni on b.id=ni.idnom left join nominalessarben nsb on b.id=nsb.idnom left join evaluaci�nmontos em on f.id=em.ideva left join evaluaci�nobs eo on f.id=eo.ideva"
        End If
    End If
'Else
'    sQueryPrin = gsQueryPrin
'End If
If gyDelegaci�n <> 90 Then
    i = BuscaCombo(ComboVarios(0), Str(ciDelegaci�n), True)
    If i >= 0 Then
        ComboVarios(0).RemoveItem i
    End If
End If
If gSQLACC = cyAccess Then Set db = OpenDatabase("z:\rpt.mdb", False, False, ";uid=;pwd=837379")
ActualizaColorFormulario Me
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
If Index = 0 Then
    List1(1).Clear
    List1(1).AddItem "Cuenta"
    List1(1).ItemData(0) = 0
    If List1(0).ListIndex > 0 Then
        List1(1).AddItem "Suma"
        List1(1).ItemData(1) = 1
        List1(1).AddItem "Promedio"
        List1(1).ItemData(2) = 2
        List1(1).AddItem "M�ximo"
        List1(1).ItemData(3) = 3
        List1(1).AddItem "M�nimo"
        List1(1).ItemData(4) = 4
    End If
    List1(1).ListIndex = 0
    txtContenido = List1(1).Text + "(" + IIf(List1(0).ListIndex >= 0, List1(0).Text, "") + ")"
ElseIf List1(1).ListIndex >= 0 Then
    txtContenido = List1(1).Text + "(" + IIf(List1(0).ListIndex >= 0, List1(0).Text, "") + ")"
End If
End Sub

Private Function ArmaCadenaColumna(ByRef sCadena As String, ByRef sFrom As String, yRengl�nColumna As Byte, ByRef sDescripci�n As String)
Dim s As String, s1 As String
Select Case ComboVarios(yRengl�nColumna).ItemData(ComboVarios(yRengl�nColumna).ListIndex)
Case 0 'Causa
    If gSQLACC = cyAccess Then
        If (chkOtros(yRengl�nColumna).Value = 1) Then
            s = "iif(mcla.id<>a.idcla,'Otra causa (Indirecta)',iif(lcase(mm.descripci�n) like 'otr[ao]*' and len(mm.descripci�n)<=5,'OTRO: '+coc.descripci�n,mm.descripci�n))"
        Else
            s = "iif(mcla.id=a.idcla or a.idcla=7,mm.descripci�n,'Otra causa (Indirecta)')"
        End If
    Else
        If InStr(sFrom, "claseinstituci�n ci") = 0 Then
            sFrom = sFrom + " left join claseinstituci�n ci on ai.idcla=ci.id"
        End If
        If (chkOtros(yRengl�nColumna).Value = 1) Then
            If gSQLACC = cyOracle Then
                s = "case when ci.id<>a.idcla then 'Otra causa (Indirecta)' else case when lower(mm.descripci�n) in ('otro','otra','otros','otras') then 'OTRO: '+coc.descripci�n else mm.descripci�n end end as Rengl�n"
            Else
                s = "Rengl�n = case when mcla.id<>a.idcla then 'Otra causa (Indirecta)' else case when lower(mm.descripci�n) like 'otr[ao]%' and len(mm.descripci�n)<=5 then 'OTRO: '+coc.descripci�n else mm.descripci�n end end"
            End If
        Else
            If gSQLACC = cyOracle Then
                s = "case when (ci.id=a.idcla or a.idcla=7) then mm.descripci�n else 'Otra causa (Indirecta)' end as Rengl�n"
            Else
                s = "Rengl�n = case when (mcla.id=a.idcla or a.idcla=7) then mm.descripci�n else 'Otra causa (Indirecta)' end"
            End If
        End If
    End If
    If (chkOtros(yRengl�nColumna).Value = 1) Then
        If gSQLACC = cyAccess Then
            sFrom = "((" + sFrom + ") left join causas mm* on a.idcau=mm*.id) left join conotroscausas coc on a.id=coc.idasu"
        Else
            sFrom = sFrom + " left join causas mm* on a.idcau=mm*.id left join conotroscausas coc on a.id=coc.idasu"
        End If
    Else
        If gSQLACC = cyAccess Then
            sFrom = "(" + sFrom + ") left join causas mm* on a.idcau=mm*.id"
        Else
            sFrom = sFrom + " left join causas mm* on a.idcau=mm*.id"
        End If
    End If
    s1 = "Causa        "
Case 1 'Clase de instituci�n
    If gSQLACC = cyAccess Then
        s = "iif(isnull(mcla.id) and not isnull(a.idcla),ci.descripci�n,mcla.descripci�n)"
        sFrom = "(" + sFrom + ") left join claseinstituci�n ci on a.idcla=ci.id"
    Else
        s = "nvl(ci.descripci�n,'Clase no capturada') as Rengl�n"
        If InStr(sFrom, "claseinstituci�n ci") = 0 Then sFrom = sFrom + " left join claseinstituci�n ci on ai.idcla=ci.id"
    End If
    s1 = "Clase de instituci�n"
Case 2 'Delegaci�n de Condusef
    If gSQLACC = cyAccess Then
        s = "iif(isnull(deln.descripci�n),dela.descripci�n,deln.descripci�n)"
        sFrom = "((" + sFrom + ") left join delegaciones deln on b.iddel=deln.id) left join delegaciones dela on d.iddel=dela.id"
    Else
        If glProceso = 100 Then
            s = "d.descripci�n"
            sFrom = sFrom + " left join nominales n on a.id=n.idasu left join delegaciones d on n.iddel=d.id"
        Else
            s = "case when deln.descripci�n is null then dela.descripci�n else deln.descripci�n end as Rengl�n"
            sFrom = sFrom + " left join nominales n on a.id=n.idasu left join delegaciones deln on n.iddel=deln.id left join an�nimos a_o on a.id=a_o.idasu left join delegaciones dela on a_o.iddel=dela.id"
        End If
    End If
    s1 = "Delegaci�n de Condusef"
Case 3 'Desenlace del asunto
    s = "mm.descripci�n"
    If gSQLACC = cyAccess Then
        sFrom = "((" + sFrom + ") left join evaluaci�n p on a.id=p.idasu) left join Desenlaces mm* on p.iddes=mm*.id"
    Else
        sFrom = sFrom + " left join evaluaci�n p on a.id=p.idasu left join Desenlaces mm* on p.iddes=mm*.id"
    End If
    s1 = "Desenlace del asunto"
Case 4 'Estado
    s = "mm.descripci�n"
    If InStr(sFrom, " asus1 ") = 0 Then
        If gSQLACC = cyAccess Then
            sFrom = "(((" + sFrom + ") left join (select ai.* from asuntousuario ai inner join (select idasu,min(idusu) as idusuario from asuntousuario group by idasu) as asus1 on ai.idasu=asus1.idasu and ai.idusu=asus1.idusuario) as au on a.id=au.idasu) left join usuarios c on au.idusu=c.id) left join estados mm* on c.idedo=mm*.id"
        Else
            sFrom = sFrom + " left join (select ai.* from asuntousuario ai inner join (select idasu,min(idusu) as idusuario from asuntousuario group by idasu) asus1 on ai.idasu=asus1.idasu and ai.idusu=asus1.idusuario) au on a.id=au.idasu left join usuarios c on au.idusu=c.id left join estados mm* on c.idedo=mm*.id"
        End If
    Else
        If gSQLACC = cyAccess Then
            sFrom = "(" + sFrom + ") left join estados mm* on c.idedo=mm*.id"
        Else
            sFrom = sFrom + " left join estados mm* on c.idedo=mm*.id"
        End If
    End If
    s1 = "Estado"
Case 5 'Fecha de recepci�n
    If gSQLACC = cyAccess Then
        s = "format(a.fecha,'" & gsFormatoFecha & "')"
    Else
        If gSQLACC = cyAccess Then
            s = "convert(nvarchar,a.fecha,105)"
        Else
            s = "to_char(a.fecha,'DD-MM-YYYY')"
        End If
    End If
    s1 = "Fecha de recepci�n"
Case 6 'Forma de recepci�n
    If gSQLACC = cyAccess Then
        s = "iif(a.idrec=0,'Personal',iif(a.idrec=1,'Telef�nica',iif(a.idrec=2,'E-Mail',iif(a.idrec=3,'Escrito',iif(a.idrec=4,'Fax',iif(a.idrec=5,'CAT','Otro'))))))"
    Else
        s = "r.descripci�n as Rengl�n"
        'sFrom = Replace(sFrom, "asuntos a left join Nominales b on a.id=b.idAsu", "asuntos a left join Recepci�n r ON a.idrec = r.id left join Nominales b on a.id=b.idAsu")
        sFrom = sFrom & " left join recepci�n r ON a.idrec = r.id"
    End If
    s1 = "Forma de recepci�n"
Case 7 'Instituci�n
    If gSQLACC = cyAccess Then
        If (chkOtros(yRengl�nColumna).Value = 1) Then
            s = "iif(a.idsec=7 or (lcase(mins.descripci�n) like 'otr[ao]*' and len(mins.descripci�n)<=5),'OTRA: '+coi.descripci�n,mins.descripci�n)"
            sFrom = "(" + sFrom + ") left join conotrosinstituciones coi on a.id=coi.idasu"
        Else
            s = "mins.descripci�n"
        End If
    Else
        If (chkOtros(yRengl�nColumna).Value = 1) Then
            If gSQLACC = cyOracle Then
                s = "case when a.idsec=7 or (lower(i.descripci�n) like 'otr_%' and substr(lower(i.descripci�n),4,1) in ('a','o') and length(i.descripci�n)<=5) then 'OTRA: '+coi.descripci�n else i.descripci�n end as Rengl�n"
            Else
                s = "Rengl�n = case when a.idsec=7 or (lower(mins.descripci�n) like 'otr[ao]%' and len(mins.descripci�n)<=5) then'OTRA: '+coi.descripci�n else mins.descripci�n end"
            End If
            sFrom = sFrom & " left join instituciones i ON ai.idins = i.id left join conotrosinstituciones coi on a.id=coi.idasu"
        Else
            s = "i.descripci�n"
            sFrom = sFrom & " left join instituciones i ON ai.idins = i.id"
        End If
    End If
    s1 = "Instituci�n"
Case 8 'Mes de recepci�n
    If gSQLACC = cyOracle Then
        s = "'Mes: '||to_char(a.fecha,'MM')"
    Else
        s = "'Mes: '+str(Month(a.fecha))"
    End If
    s1 = "Mes de recepci�n"
Case 9 'Municipio
    s = "mm.descripci�n"
    If InStr(sFrom, " asus1 ") = 0 Then
        If gSQLACC = cyAccess Then
            sFrom = "(((" + sFrom + ") left join (select ai.* from asuntousuario ai inner join (select idasu,min(idusu) as idusuario from asuntousuario group by idasu) as asus1 on ai.idasu=asus1.idasu and ai.idusu=asus1.idusuario) as au on a.id=au.idasu) left join usuarios c on au.idusu=c.id) left join municipios mm* on c.idmun=mm*.id"
        Else
            sFrom = sFrom + " left join (select ai.* from asuntousuario ai inner join (select idasu,min(idusu) as idusuario from asuntousuario group by idasu) asus1 on ai.idasu=asus1.idasu and ai.idusu=asus1.idusuario) au on a.id=au.idasu left join usuarios c on au.idusu=c.id left join municipios mm* on c.idmun=mm*.id"
        End If
    Else
        If gSQLACC = cyAccess Then
            sFrom = "(" + sFrom + ") left join municipios mm* on c.idmun=mm*.id"
        Else
            sFrom = sFrom + " left join municipios mm* on c.idmun=mm*.id"
        End If
    End If
    s1 = "Municipio"
Case 10 'Personalidad voy aqu�
    If gSQLACC = cyAccess Then
        s = "iif(c.permoral=0,'F�sica','Moral')"
        If InStr(sFrom, " asus1 ") = 0 Then
            sFrom = "((" + sFrom + ") left join (select ai.* from asuntousuario ai inner join (select idasu,min(idusu) as idusuario from asuntousuario group by idasu) as asus1 on ai.idasu=asus1.idasu and ai.idusu=asus1.idusuario) as au on a.id=au.idasu) left join usuarios c on au.idusu=c.id"
        End If
    Else
        s = "case when c.permoral=0 then 'F�sica' else 'Moral' end as Rengl�n"
        If InStr(sFrom, " asus1 ") = 0 Then
            sFrom = sFrom + " left join (select ai.* from asuntousuario ai inner join (select idasu,min(idusu) as idusuario from asuntousuario group by idasu) asus1 on ai.idasu=asus1.idasu and ai.idusu=asus1.idusuario) au on a.id=au.idasu left join usuarios c on au.idusu=c.id"
        End If
    End If
    s1 = "Persona F�sica o Moral"
Case 11 'Procedencia
    s = "mm.descripci�n"
    If gSQLACC = cyAccess Then
        sFrom = "(" + sFrom + ") left join tipoprocedencia mm* on e.idprc=mm*.id"
    Else
        sFrom = sFrom + " left join procedencias p on a.id=p.idasu left join tipoprocedencia mm* on p.idprc=mm*.id"
    End If
    s1 = "Procedencia"
Case 12 'Producto nivel 1
    If gSQLACC = cyAccess Then
        If (chkOtros(yRengl�nColumna).Value = 1) Then
            s = "iif(mcla.id<>a.idcla,'Otro producto nivel 1 (Indirecto)',iif(lcase(mm.descripci�n) like 'otr[ao]*' and len(mm.descripci�n)<=5,'OTRO: '+cop1.descripci�n,mm.descripci�n))"
        Else
            s = "iif(mcla.id<>a.idcla,'Otro producto nivel 1 (Indirecto)',mm.descripci�n)"
        End If
        If (chkOtros(yRengl�nColumna).Value = 1) Then
            sFrom = "((" + sFrom + ") left join productosnivel1 mm* on a.idpr1=mm*.id) left join conotrosproductosn1 cop1 on a.id=cop1.idasu"
        Else
            sFrom = "(" + sFrom + ") left join productosnivel1 mm* on a.idpr1=mm*.id"
        End If
    Else
        If InStr(sFrom, "claseinstituci�n ci") = 0 Then
            sFrom = sFrom + " left join claseinstituci�n ci on ai.idcla=ci.id"
        End If
        If (chkOtros(yRengl�nColumna).Value = 1) Then
            If gSQLACC = cyOracle Then
                s = "case when ci.id<>a.idcla then 'Otro producto nivel 1 (Indirecto)' else case when lower(mm.descripci�n) like 'otr_%' and substr(mm.descripci�n,4,1) in ('a','o') and length(mm.descripci�n)<=5 then 'OTRO: '+cop1.descripci�n else mm.descripci�n end end as Rengl�n"
            Else
                s = "case when ci.id<>a.idcla then 'Otro producto nivel 1 (Indirecto)' else case when lower(mm.descripci�n) like 'otr[ao]%' and len(mm.descripci�n)<=5 then 'OTRO: '+cop1.descripci�n else mm.descripci�n end end as Rengl�n"
            End If
        Else
            s = "case when ci.id<>a.idcla then 'Otro producto nivel 1 (Indirecto)' else mm.descripci�n end as Rengl�n"
        End If
        If (chkOtros(yRengl�nColumna).Value = 1) Then
            sFrom = sFrom + " left join productosnivel1 mm* on a.idpr1=mm*.id left join conotrosproductosn1 cop1 on a.id=cop1.idasu"
        Else
            sFrom = sFrom + " left join productosnivel1 mm* on a.idpr1=mm*.id"
        End If
    End If
    s1 = "Producto Nivel 1"
Case 13 'Producto nivel 2
    If gSQLACC = cyAccess Then
        If (chkOtros(yRengl�nColumna).Value = 1) Then
            s = "iif(mcla.id<>a.idcla,'Otro producto nivel 2 (Indirecto)',iif(lcase(mm.descripci�n) like 'otr[ao]*' and len(mm.descripci�n)<=5,'OTRO: '+cop2.descripci�n,mm.descripci�n))"
        Else
            s = "iif(mcla.id<>a.idcla,'Otro producto nivel 2 (Indirecto)',mm.descripci�n)"
        End If
        If (chkOtros(yRengl�nColumna).Value = 1) Then
            sFrom = "((" + sFrom + ") left join productosnivel2 mm* on a.idpr2=mm*.id) left join conotrosproductosn2 cop2 on a.id=cop2.idasu"
        Else
            sFrom = "(" + sFrom + ") left join productosnivel2 mm* on a.idpr2=mm*.id"
        End If
    Else
        If InStr(sFrom, "claseinstituci�n ci") = 0 Then
            sFrom = sFrom + " left join claseinstituci�n ci on ai.idcla=ci.id"
        End If
        If (chkOtros(yRengl�nColumna).Value = 1) Then
            If gSQLACC = cyOracle Then
                s = "case when ci.id<>a.idcla then 'Otro producto nivel 2 (Indirecto)' else  case  when lower(mm.descripci�n) like 'otr_%' and substr(mm.descripci�n,4,1) in ('a','o') and length(mm.descripci�n)<=5 then 'OTRO: '+cop2.descripci�n else mm.descripci�n end end as Rengl�n"
            Else
                s = "Rengl�n = case when ci.id<>a.idcla then 'Otro producto nivel 2 (Indirecto)' else  case  when lower(mm.descripci�n) like 'otr[ao]%' and len(mm.descripci�n)<=5 then 'OTRO: '+cop2.descripci�n else mm.descripci�n end end"
            End If
        Else
            s = "case when ci.id<>a.idcla then 'Otro producto nivel 2 (Indirecto)' else mm.descripci�n end as Rengl�n"
        End If
        If (chkOtros(yRengl�nColumna).Value = 1) Then
            sFrom = sFrom + " left join productosnivel2 mm* on a.idpr2=mm*.id left join conotrosproductosn2 cop2 on a.id=cop2.idasu"
        Else
            sFrom = sFrom + " left join productosnivel2 mm* on a.idpr2=mm*.id"
        End If
    End If
    s1 = "Producto Nivel 2"
Case 14 'Producto nivel 3
    If gSQLACC = cyAccess Then
        If (chkOtros(yRengl�nColumna).Value = 1) Then
            s = "iif(mcla.id<>a.idcla,'Otro producto nivel 3 (Indirecto)',iif(lcase(mm.descripci�n) like 'otr[ao]*' and len(mm.descripci�n)<=5,'OTRO: '+cop3.descripci�n,mm.descripci�n))"
        Else
            s = "iif(mcla.id<>a.idcla,'Otro producto nivel 3 (Indirecto)',mm.descripci�n)"
        End If
        If (chkOtros(yRengl�nColumna).Value = 1) Then
            sFrom = "((" + sFrom + ") left join productosnivel3 mm* on a.idpr3=mm*.id) left join conotrosproductosn3 cop3 on a.id=cop3.idasu"
        Else
            sFrom = "(" + sFrom + ") left join productosnivel3 mm* on a.idpr3=mm*.id"
        End If
    Else
        If InStr(sFrom, "claseinstituci�n ci") = 0 Then
            sFrom = sFrom + " left join claseinstituci�n ci on ai.idcla=ci.id"
        End If
        If (chkOtros(yRengl�nColumna).Value = 1) Then
            If gSQLACC = cyOracle Then
                s = "case when ci.id<>a.idcla then 'Otro producto nivel 3 (Indirecto)' else  case  when lower(mm.descripci�n) like 'otr_%' and substr(mm.descripci�n,4,1) in ('a','o') and length(mm.descripci�n)<=5 then 'OTRO: '+cop3.descripci�n else mm.descripci�n end end as Rengl�n"
            Else
                s = "Rengl�n = case when ci.id<>a.idcla then 'Otro producto nivel 3 (Indirecto)' else case lower(mm.descripci�n) like 'otr[ao]%' and len(mm.descripci�n)<=5 then 'OTRO: '+cop3.descripci�n else mm.descripci�n end"
            End If
        Else
            s = "case when ci.id<>a.idcla then 'Otro producto nivel 3 (Indirecto)' else mm.descripci�n end as Rengl�n"
        End If
        If (chkOtros(yRengl�nColumna).Value = 1) Then
            sFrom = sFrom + " left join productosnivel3 mm* on a.idpr3=mm*.id left join conotrosproductosn3 cop3 on a.id=cop3.idasu"
        Else
            sFrom = sFrom + " left join productosnivel3 mm* on a.idpr3=mm*.id"
        End If
    End If
    s1 = "Producto nivel 3"
Case 15 'Responsable de la 1a Actividad
    s = "mm.descripci�n"
    If gSQLACC = cyAccess Then
        sFrom = "((" + sFrom + ") left join (select * from avances where id=idant) as av1 on m.id=av1.idasuins) left join responsables mm* on av1.idres=mm*.id"
    Else
        'If InStr(sFrom, "avances av") = 0 Then
        '    sFrom = sFrom + " left join avances av on ai.id=av.idasuins"
        'End If
        sFrom = sFrom + " left join (select * from avances where id=idant) av2 on ai.id=av2.idasuins left join responsables mm* on av2.idres=mm*.id"
    End If
    s1 = "Responsable de 1a Actividad"
Case 16 'Sector financiero
    If gSQLACC = cyAccess Then
        s = "iif(not isnull(m.idcla),mm.descripci�n,'No especificado')"
        sFrom = "(" + sFrom + ") left join sectorfinanciero mm* on mcla.idsec=mm*.id"
    Else
        s = "nvl(mm.descripci�n,'No especificado') as Rengl�n"
        If InStr(sFrom, "claseinstituci�n ci") = 0 Then
            sFrom = sFrom + " left join claseinstituci�n ci on ai.idcla=ci.id left join sectorfinanciero mm* on ci.idsec=mm*.id"
        Else
            sFrom = sFrom + " left join sectorfinanciero mm* on ci.idsec=mm*.id"
        End If
    End If
    s1 = "Sector financiero"
Case 17 'Tipo de asunto
    If gSQLACC = cyAccess Then
        s = "iif(a.atenci�n<>0,'Nominativa','An�nimo')"
    Else
        s = "case when a.atenci�n<>0 then 'Nominativa' else 'An�nimo' end as Rengl�n"
    End If
    s1 = "Tipo de asunto"
Case 18 'Tipo identificaci�n
    s = "mm.descripci�n"
    If gSQLACC = cyAccess Then
        sFrom = "((" + sFrom + ") left join nominalesidentificaci�n nni on b.id=nni.idnom) left join tipoidentificaci�n mm* on nni.idide=mm*.id"
    Else
        If InStr(sFrom, "nominales n") = 0 Then
            sFrom = sFrom + " left join nominales n on a.id=n.idasu left join nominalesidentificaci�n nni on n.id=nni.idnom left join tipoidentificaci�n mm* on nni.idide=mm*.id"
        Else
            sFrom = sFrom + " left join nominalesidentificaci�n nni on n.id=nni.idnom left join tipoidentificaci�n mm* on nni.idide=mm*.id"
        End If
    End If
    s1 = "Tipo identificaci�n"
Case 19 '�ltima Actividad/Desenlace AT
    s = "mm.�ltimaActividad_AT"
    If gSQLACC = cyAccess Then
        sFrom = "(" + sFrom + ") left join (select av.idasuins,t.descripci�n + ' / ' + iif(d.descripci�n is null,'',d.descripci�n) as �ltimaActividad_AT from ((avances av inner join actividades t on av.idtar=t.id) left join desenlaces d on av.iddes=d.id) inner join (select a.idasuins,max(a.id) as idava from avances a inner join actividades ac on a.idtar=ac.id where a.fecha is not null and ac.idpad in (2,96) group by a.idasuins) maxAT on av.id=maxAT.idava) mm* on m.id=mm*.idasuins"
    Else
        If gSQLACC = cyOracle Then
            'sFrom = sFrom + " left join (select av.idasuins,t.descripci�n ||' / '|| case when d.descripci�n is null then '' else d.descripci�n end as �ltimaActividad_AT from avances av inner join actividades t on av.idtar=t.id left join desenlaces d on av.iddes=d.id inner join (select a.idasuins,max(a.id) as idava from avances a inner join actividades ac on a.idtar=ac.id where a.fecha is not null and ac.idpad in (2,96) group by a.idasuins) maxAT on av.id=maxAT.idava) mm* on av.idasuins=mm*.idasuins"
            sFrom = sFrom + " left join (select av.idasuins,t.descripci�n || ' / ' || nvl(d.descripci�n,'') as �ltimaActividad_AT from avances av inner join actividades t on av.idtar=t.id left join desenlaces d on av.iddes=d.id inner join (select av.idasuins,max(av.id) as idava from avances av inner join actividades ac on av.idtar=ac.id inner join asuntoinstituci�n ai on av.idasuins=ai.id inner join asuntos a on ai.idasu=a.id where av.fecha is not null and ac.idpad in (2,96)" & gsSeparador & " group by av.idasuins) maxAT on av.id=maxAT.idava) mm* on ai.id=mm*.idasuins"
        Else
            sFrom = sFrom + " left join (select av.idasuins,t.descripci�n + ' / ' + case when d.descripci�n is null then '' else d.descripci�n end as �ltimaActividad_AT from avances av inner join actividades t on av.idtar=t.id left join desenlaces d on av.iddes=d.id inner join (select a.idasuins,max(a.id) as idava from avances a inner join actividades ac on a.idtar=ac.id where a.fecha is not null and ac.idpad in (2,96) group by a.idasuins) maxAT on av.id=maxAT.idava) mm* on av.idasuins=mm*.idasuins"
        End If
    End If
    s1 = "�ltima Actividad/Desenlace AT"
Case 20 '�ltima Actividad/Desenlace CO
    s = "mm.�ltimaActividad_CO"
    If gSQLACC = cyAccess Then
        sFrom = "(" + sFrom + ") left join (select av.idasuins,t.descripci�n + ' / ' + iif(d.descripci�n is null,'',d.descripci�n) as �ltimaActividad_CO from ((avances av inner join actividades t on av.idtar=t.id) left join desenlaces d on av.iddes=d.id) inner join (select a.idasuins,max(a.id) as idava from avances a inner join actividades ac on a.idtar=ac.id where a.fecha is not null and ac.idpad=4 group by a.idasuins) maxCO on av.id=maxCO.idava) mm* on m.id=mm*.idasuins"
    Else
        If gSQLACC = cyOracle Then
            'sFrom = sFrom + " left join (select av.idasuins,t.descripci�n ||' / '|| case when d.descripci�n is null then '' else d.descripci�n end as �ltimaActividad_CO from ((avances av inner join actividades t on av.idtar=t.id) left join desenlaces d on av.iddes=d.id) inner join (select a.idasuins,max(a.id) as idava from avances a inner join actividades ac on a.idtar=ac.id where a.fecha is not null and ac.idpad=4 group by a.idasuins) maxCO on av.id=maxCO.idava) mm* on av.idasuins=mm*.idasuins"
            sFrom = sFrom + " left join (select av.idasuins,t.descripci�n || ' / ' || nvl(d.descripci�n,'') as �ltimaActividad_CO from avances av inner join actividades t on av.idtar=t.id left join desenlaces d on av.iddes=d.id inner join (select av.idasuins,max(av.id) as idava from avances av inner join actividades ac on av.idtar=ac.id inner join asuntoinstituci�n ai on av.idasuins=ai.id inner join asuntos a on ai.idasu=a.id where av.fecha is not null and ac.idpad=4" & gsSeparador & " group by av.idasuins) maxCO on av.id=maxCO.idava) mm* on ai.id=mm*.idasuins"
        Else
            sFrom = sFrom + " left join (select av.idasuins,t.descripci�n + ' / ' + case when d.descripci�n is null then '' else d.descripci�n end as �ltimaActividad_CO from ((avances av inner join actividades t on av.idtar=t.id) left join desenlaces d on av.iddes=d.id) inner join (select a.idasuins,max(a.id) as idava from avances a inner join actividades ac on a.idtar=ac.id where a.fecha is not null and ac.idpad=4 group by a.idasuins) maxCO on av.id=maxCO.idava) mm* on av.idasuins=mm*.idasuins"
        End If
    End If
    s1 = "�ltima Actividad/Desenlace CO"
Case 21 'Usuario del SIO
    s = "mm.descripci�n"
    If gSQLACC = cyAccess Then
        sFrom = "(" + sFrom + ") left join usuariossistema mm* on a.idusi=mm*.id"
    Else
        If glProceso = 100 Then
            sFrom = sFrom + " left join usuariossistema mm* on av.idusi=mm*.id"
        Else
            sFrom = sFrom + " left join usuariossistema mm* on a.idusi=mm*.id"
        End If
    End If
    s1 = "Usuario del SIO"
End Select
If InStr(s, "Rengl�n") = 0 Then
    s = s + " as Rengl�n,"
Else
    s = s + ","
End If
If yRengl�nColumna > 0 Then
    s = Replace(Replace(Replace(s, "Rengl�n", "Columna"), "mm*", "nn"), "mm.", "nn.")
    sFrom = Replace(sFrom, "mm*", "nn")
Else
    s = Replace(s, "mm*", "mm")
    sFrom = Replace(sFrom, "mm*", "mm")
    sDescripci�n = s1
End If
sCadena = sCadena + s
End Function

Private Sub opcRango_Click(Index As Integer)
Dim d As Date
Select Case Index
Case 0
    DTP(0).Value = CDate("01/01/" & (Year(Date) - 1))
    DTP(1).Value = DateAdd("yyyy", 1, DTP(0).Value)
    DTP(1).Value = DTP(1).Value - 1
Case 1
    d = DateAdd("m", IIf(Month(Date) Mod 2 = 0, -3, -2), Date)
    DTP(0).Value = d - Day(d) + 1
    d = DateAdd("m", 2, d)
    DTP(1).Value = d - Day(d)
Case 2
    DTP(1).Value = Date - Day(Date)
    DTP(0).Value = DTP(1).Value - Day(DTP(1).Value) + 1
Case 3
    DTP(1).Value = Date - Weekday(Date, vbSaturday)
    DTP(0).Value = DTP(1).Value - 4
Case 4
    If Not DTP(0).Enabled Then
        DTP(0).Enabled = True
        DTP(1).Enabled = True
    End If
End Select
If Index < 4 And DTP(0).Enabled Then
    DTP(0).Enabled = False
    DTP(1).Enabled = False
End If
End Sub

Private Sub UpDown_DownClick()
Dim d As Date, i As Integer
For i = 0 To opcRango.UBound
    If opcRango(i).Value Then Exit For
Next
Select Case i
Case 0
    DTP(0).Value = CDate("01/01/" & (Year(DTP(0).Value) - 1))
    DTP(1).Value = DateAdd("yyyy", 1, DTP(0).Value) - 1
Case 1
    d = DateAdd("m", -2, DTP(0).Value)
    DTP(0).Value = d - Day(d) + 1
    d = DateAdd("m", 2, d)
    DTP(1).Value = d - Day(d)
Case 2
    DTP(1).Value = DTP(1).Value - Day(DTP(1).Value)
    DTP(0).Value = DTP(1).Value - Day(DTP(1).Value) + 1
Case 3
    DTP(1) = DTP(1) - 7
    DTP(0).Value = DTP(1).Value - 4
Case 4
    DTP(0).Value = DTP(0).Value - 1
    DTP(1).Value = DTP(1).Value - 1
End Select
End Sub

Private Sub UpDown_UpClick()
Dim d As Date, i As Integer
For i = 0 To opcRango.UBound
    If opcRango(i).Value Then Exit For
Next
Select Case i
Case 0
    DTP(0).Value = CDate("01/01/" & (Year(DTP(0).Value) + 1))
    DTP(1).Value = DateAdd("yyyy", 1, DTP(0).Value) - 1
Case 1
    d = DateAdd("m", 2, DTP(0).Value)
    DTP(0).Value = d - Day(d) + 1
    d = DateAdd("m", 2, d)
    DTP(1).Value = d - Day(d)
Case 2
    d = DateAdd("m", 2, DTP(1).Value)
    DTP(1).Value = d - Day(d)
    DTP(0).Value = DTP(1).Value - Day(DTP(1).Value) + 1
Case 3
    DTP(1) = DTP(1) + 7
    DTP(0).Value = DTP(1).Value - 4
Case 4
    DTP(0).Value = DTP(0).Value + 1
    DTP(1).Value = DTP(1).Value + 1
End Select
End Sub

