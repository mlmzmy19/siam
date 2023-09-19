Attribute VB_Name = "Reportes"
Sub PR_CargaComboReportes()
Dim S_Condi As String, RE_CargaReportes As ADODB.Recordset

On Error GoTo ErrRecataInfRep

S_Condi = "Select r.n_cvereporte,r.s_descrip,r.s_ruta From c_reportes r (Nolock) Where n_cvereporte_p = " & gn_OpcionReporte & " Order by r.n_cvereporte"

Set RE_CargaReportes = New ADODB.Recordset
RE_CargaReportes.Open S_Condi, gcn

With Frm_ReportesGen
    .Cmb_RepGrupo.Clear
    Do While Not RE_CargaReportes.EOF()
        .Cmb_RepGrupo.AddItem Trim(RE_CargaReportes(0)) & " " & Trim(RE_CargaReportes(1))
        .Cmb_RepGrupo.ItemData(.Cmb_RepGrupo.NewIndex) = Trim(RE_CargaReportes(0))
        RE_CargaReportes.MoveNext
    Loop
End With

RE_CargaReportes.Close
Set RE_CargaReportes = Nothing
Exit Sub

ErrRecataInfRep:
If Err.Number <> 0 Then
    MsgBox Str(Err.Number) & ": " & Err.Description, 0 + 64, "[PR_CargaComboReportes]"
    Exit Sub
End If
'************************************************************************************************************************
'*Rescata información de un registro especifico dada una condición, de alguna tabla X
'***********************************************************************************************************************
End Sub

Sub PR_SelecionaReporteVariable()
Dim S_MicondicionBus As String

With Frm_ReportesGen
    S_MicondicionBus = "n_cvereporte = " & .Lbl_RepCveGrupo.Caption
    .Lbl_RepNombre = FU_RescataInfX_CampoS("s_nombre", "c_reportes", S_MicondicionBus)
    .Lbl_RepRuta = FU_RescataInfX_CampoS("s_ruta", "c_reportes", S_MicondicionBus)
    .StatusBar_Rep.SimpleText = Trim(.Lbl_RepRuta) & "\" & Trim(.Lbl_RepNombre) & ".rpt"
End With
End Sub
