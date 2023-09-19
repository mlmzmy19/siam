Attribute VB_Name = "Module2"
Public gcn As New ADODB.Connection
Public MidB As Database
Public S_Path As String
Public gs_usuario As Long
Public gn_Miperfil As Integer
Public gs_Proyecto As Long
Public gs_Etapa As Long
Public gs_Cuestionario As String

Public gs_ProcCuestionario As Integer

Public RE_Etapas As ADODB.Recordset
Public RE_Perfiles As ADODB.Recordset
Public RE_Usuarios As ADODB.Recordset
Public RE_Proyectos As ADODB.Recordset
Public gn_OpcionReporte As Integer


Function FU_ExistenciaValor_ResumenContactos()
Dim N_Suma As Integer, N_Valor As Variant

FU_ExistenciaValor_ResumenContactos = 0
With Frm_Resumen
    For li_Col = 1 To Val(.MSFlexGrid1.Cols - 1)
        N_Valor = Trim(.MSFlexGrid1.TextMatrix(2, li_Col))
        N_Valor = IIf(IsNumeric(N_Valor), Val(N_Valor), 0)
        N_Suma = N_Suma + N_Valor
    Next

    If N_Suma = 0 Then
        MsgBox "Al menos debe de existir un valor númerico mayor a 0.", 0 + 16, "Aún no exiten datos"
        .MSFlexGrid1.SetFocus
        .MSFlexGrid1.Col = 1
        FU_ExistenciaValor_ResumenContactos = -1
     End If
End With
End Function

Function FU_ExtraePosiblesValosAEscojer(N_Mivalor, N_MiModulo) As String
Dim RE_MisDatosExt As ADODB.Recordset, S_CondiCargaExt As String, S_CadenaValores As String

FU_ExtraePosiblesValosAEscojer = ""
S_CadenaValores = "-"
If Val(N_MiModulo) = 1 Or Val(N_MiModulo) = 2 Or Val(N_MiModulo) = 3 Then
     S_CondiCargaExt = "Select n_opcion_base From T_Caratulas (Nolock)" & _
    "Where n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And n_modulo = " & N_MiModulo
ElseIf Val(N_MiModulo) = 4 Then
    S_CondiCargaExt = "Select n_opcion_det,n_opcion_base  From T_CaratulasDet (Nolock)" & _
    "Where n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And n_modulo = " & N_MiModulo & " And n_opcion_Base = " & N_Mivalor
ElseIf Val(N_MiModulo) = 8 Or Val(N_MiModulo) = 9 Then
    S_CondiCargaExt = "Select n_opcion_base From t_caratulas (Nolock)" & _
    "Where n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And n_modulo = " & N_MiModulo & " And c_tipo_opcion = 'L'"
End If

Select Case Val(N_MiModulo)
    Case 1 To 3
        Set RE_MisDatosExt = New ADODB.Recordset
        RE_MisDatosExt.Open S_CondiCargaExt, gcn
        
        If Val(RE_MisDatosExt.RecordCount) = 0 Then Exit Function
        Do While Not RE_MisDatosExt.EOF()
             S_CadenaValores = S_CadenaValores & Trim(Str(RE_MisDatosExt(0))) & "-"
             RE_MisDatosExt.MoveNext
        Loop
        FU_ExtraePosiblesValosAEscojer = S_CadenaValores
    
    Case 4
        Set RE_MisDatosExt = New ADODB.Recordset
        RE_MisDatosExt.Open S_CondiCargaExt, gcn
        
        If Val(RE_MisDatosExt.RecordCount) = 0 Then Exit Function
        Do While Not RE_MisDatosExt.EOF()
             S_CadenaValores = S_CadenaValores & Trim(Str(RE_MisDatosExt(0))) & "-"
             RE_MisDatosExt.MoveNext
        Loop
        FU_ExtraePosiblesValosAEscojer = S_CadenaValores
        
    Case 8
        Set RE_MisDatosExt = New ADODB.Recordset
        RE_MisDatosExt.Open S_CondiCargaExt, gcn
        
        If Val(RE_MisDatosExt.RecordCount) = 0 Then Exit Function
        Do While Not RE_MisDatosExt.EOF()
             S_CadenaValores = S_CadenaValores & Trim(Str(RE_MisDatosExt(0))) & "-"
             RE_MisDatosExt.MoveNext
        Loop
        FU_ExtraePosiblesValosAEscojer = S_CadenaValores
End Select

RE_MisDatosExt.Close
Set RE_MisDatosExt = Nothing
'----------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------
End Function

Function FU_ExtraePosiblesValosAEscojerCATALOGOS(S_Control) As String
Dim RE_MisDatosExt As ADODB.Recordset, S_CondiCargaExt As String, S_CadenaValores As String

FU_ExtraePosiblesValosAEscojerCATALOGOS = ""
S_CadenaValores = "-"

Select Case UCase(Trim(S_Control))
    Case "SUPERVISION-GDV"
        S_CondiCargaExt = "SELECT a.n_cvecatgral From c_catalogosanexos a (Nolock) Where a.n_cvecatgral_p = 1"
        
        Set RE_MisDatosExt = New ADODB.Recordset
        RE_MisDatosExt.Open S_CondiCargaExt, gcn
        
        If Val(RE_MisDatosExt.RecordCount) = 0 Then Exit Function
        Do While Not RE_MisDatosExt.EOF()
             S_CadenaValores = S_CadenaValores & Trim(Str(RE_MisDatosExt(0))) & "-"
             RE_MisDatosExt.MoveNext
        Loop
        FU_ExtraePosiblesValosAEscojerCATALOGOS = S_CadenaValores
        
    Case "SUPERVISION-OUT"
        S_CondiCargaExt = "SELECT a.n_cvecatgral From c_catalogosanexos a (Nolock) Where a.n_cvecatgral_p = 2"
        
        Set RE_MisDatosExt = New ADODB.Recordset
        RE_MisDatosExt.Open S_CondiCargaExt, gcn
        
        If Val(RE_MisDatosExt.RecordCount) = 0 Then Exit Function
        Do While Not RE_MisDatosExt.EOF()
             S_CadenaValores = S_CadenaValores & Trim(Str(RE_MisDatosExt(0))) & "-"
             RE_MisDatosExt.MoveNext
        Loop
        FU_ExtraePosiblesValosAEscojerCATALOGOS = S_CadenaValores
        
    Case "ENTREVISTA-GDV"
        S_CondiCargaExt = "SELECT a.n_cvecatgral From c_catalogosanexos a (Nolock) Where a.n_cvecatgral_p = 3"
        
        Set RE_MisDatosExt = New ADODB.Recordset
        RE_MisDatosExt.Open S_CondiCargaExt, gcn
        
        If Val(RE_MisDatosExt.RecordCount) = 0 Then Exit Function
        Do While Not RE_MisDatosExt.EOF()
             S_CadenaValores = S_CadenaValores & Trim(Str(RE_MisDatosExt(0))) & "-"
             RE_MisDatosExt.MoveNext
        Loop
        FU_ExtraePosiblesValosAEscojerCATALOGOS = S_CadenaValores
        
    Case "PERSONA-ENT"
        S_CondiCargaExt = "SELECT a.n_cvecatgral From c_catalogosanexos a (Nolock) Where a.n_cvecatgral_p = 4"
        
        Set RE_MisDatosExt = New ADODB.Recordset
        RE_MisDatosExt.Open S_CondiCargaExt, gcn
        
        If Val(RE_MisDatosExt.RecordCount) = 0 Then Exit Function
        Do While Not RE_MisDatosExt.EOF()
             S_CadenaValores = S_CadenaValores & Trim(Str(RE_MisDatosExt(0))) & "-"
             RE_MisDatosExt.MoveNext
        Loop
        FU_ExtraePosiblesValosAEscojerCATALOGOS = S_CadenaValores
        
    Case "SEXO"
        S_CondiCargaExt = "SELECT a.n_cvecatgral From c_catalogosanexos a (Nolock) Where a.n_cvecatgral_p = 5"
        
        Set RE_MisDatosExt = New ADODB.Recordset
        RE_MisDatosExt.Open S_CondiCargaExt, gcn
        
        If Val(RE_MisDatosExt.RecordCount) = 0 Then Exit Function
        Do While Not RE_MisDatosExt.EOF()
             S_CadenaValores = S_CadenaValores & Trim(Str(RE_MisDatosExt(0))) & "-"
             RE_MisDatosExt.MoveNext
        Loop
        FU_ExtraePosiblesValosAEscojerCATALOGOS = S_CadenaValores
    
    Case "AUDITOR-GDV"
        S_CondiCargaExt = "SELECT a.n_cvecatgral From c_catalogosanexos a (Nolock) Where a.n_cvecatgral_p = 6"
        
        Set RE_MisDatosExt = New ADODB.Recordset
        RE_MisDatosExt.Open S_CondiCargaExt, gcn
        
        If Val(RE_MisDatosExt.RecordCount) = 0 Then Exit Function
        Do While Not RE_MisDatosExt.EOF()
             S_CadenaValores = S_CadenaValores & Trim(Str(RE_MisDatosExt(0))) & "-"
             RE_MisDatosExt.MoveNext
        Loop
        FU_ExtraePosiblesValosAEscojerCATALOGOS = S_CadenaValores
        
    Case "AUDITOR-OUT"
        S_CondiCargaExt = "SELECT a.n_cvecatgral From c_catalogosanexos a (Nolock) Where a.n_cvecatgral_p = 7"
        
        Set RE_MisDatosExt = New ADODB.Recordset
        RE_MisDatosExt.Open S_CondiCargaExt, gcn
        
        If Val(RE_MisDatosExt.RecordCount) = 0 Then Exit Function
        Do While Not RE_MisDatosExt.EOF()
             S_CadenaValores = S_CadenaValores & Trim(Str(RE_MisDatosExt(0))) & "-"
             RE_MisDatosExt.MoveNext
        Loop
        FU_ExtraePosiblesValosAEscojerCATALOGOS = S_CadenaValores
End Select
    
RE_MisDatosExt.Close
Set RE_MisDatosExt = Nothing
    
End Function


Function FU_GuardaInf_GridCaratulaBD(S_Cualmod As String) As String
Dim N_Modulo As Integer, li_Row As Integer, S_ArmaQryEje As String, N_Reglones As Integer, S_ComBita As String
Dim S_DescripBita As String, S_CadenotaAviso As String, Bo_EntraCiclo As Boolean
'n_cveproyecto n_cveetapa n_modulo n_opcion_base s_descrip_base s_valorasig c_tipo_opcion n_interno

On Error GoTo ErrAccionGGrid

FU_GuardaInf_GridCaratulaBD = ""
Select Case UCase(Trim(S_Cualmod))
    Case "MODULO1"
        N_Modulo = 1
        With Frm_Caratula
            N_Reglones = Val(.Grd_Mod1.Rows - 1)
            If N_Reglones = 1 And Trim(.Grd_Mod1.TextMatrix(1, 1)) = "" And Trim(.Grd_Mod1.TextMatrix(1, 2)) = "" Then Exit Function
            For li_Row = 1 To N_Reglones
                S_ArmaQryEje = "Insert into t_caratulas values (" & _
                Trim(.Txt_CarProyecto) & "," & Trim(.Lbl_CveEtapa) & "," & N_Modulo & "," & _
                Val(Trim(.Grd_Mod1.TextMatrix(li_Row, 2))) & ",'" & Trim(.Grd_Mod1.TextMatrix(li_Row, 1)) & _
                "'," & Val(li_Row) & ",'L',0)"
                
                gcn.Execute S_ArmaQryEje
            Next
            S_ComBita = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo)
            S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (" & Trim(N_Reglones) & " Reg)"
            Call PR_GrabaBitacora(6, 1, S_DescripBita, S_ComBita)
        End With
        If N_Reglones > 0 Then FU_GuardaInf_GridCaratulaBD = S_DescripBita & vbCrLf
    
    Case "MODULO2"
        N_Modulo = 2
        With Frm_Caratula
            N_Reglones = Val(.Grd_Mod2.Rows - 1)
            If N_Reglones = 1 And Trim(.Grd_Mod2.TextMatrix(1, 1)) = "" And Trim(.Grd_Mod2.TextMatrix(1, 2)) = "" Then Exit Function
            For li_Row = 1 To N_Reglones
                S_ArmaQryEje = "Insert into t_caratulas values (" & _
                Trim(.Txt_CarProyecto) & "," & Trim(.Lbl_CveEtapa) & "," & N_Modulo & "," & _
                Val(Trim(.Grd_Mod2.TextMatrix(li_Row, 2))) & ",'" & Trim(.Grd_Mod2.TextMatrix(li_Row, 1)) & _
                "'," & Val(li_Row) & ",'L',0)"
                
                gcn.Execute S_ArmaQryEje
            Next
            S_ComBita = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo)
            S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (" & Trim(N_Reglones) & " Reg)"
            Call PR_GrabaBitacora(6, 1, S_DescripBita, S_ComBita)
        End With
        If N_Reglones > 0 Then FU_GuardaInf_GridCaratulaBD = S_DescripBita & vbCrLf
        
    Case "MODULO3"
        N_Modulo = 3
        With Frm_Caratula
            N_Reglones = Val(.Grd_Mod3.Rows - 1)
            If N_Reglones = 1 And Trim(.Grd_Mod3.TextMatrix(1, 1)) = "" And Trim(.Grd_Mod3.TextMatrix(1, 2)) = "" Then Exit Function
            For li_Row = 1 To N_Reglones
                S_ArmaQryEje = "Insert into t_caratulas values (" & _
                Trim(.Txt_CarProyecto) & "," & Trim(.Lbl_CveEtapa) & "," & N_Modulo & "," & _
                Val(Trim(.Grd_Mod3.TextMatrix(li_Row, 2))) & ",'" & Trim(.Grd_Mod3.TextMatrix(li_Row, 1)) & _
                "'," & Val(li_Row) & ",'L',0)"
                
                gcn.Execute S_ArmaQryEje
            Next
            S_ComBita = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo)
            S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (" & Trim(N_Reglones) & " Reg)"
            Call PR_GrabaBitacora(6, 1, S_DescripBita, S_ComBita)
        End With
        If N_Reglones > 0 Then FU_GuardaInf_GridCaratulaBD = S_DescripBita & vbCrLf
    
    Case "MODULO4"
        N_Modulo = 4
        With Frm_Caratula
            N_Reglones = Val(.Grd_Mod4.Rows - 1)
            If N_Reglones = 1 And Trim(.Grd_Mod4.TextMatrix(1, 1)) = "" And Trim(.Grd_Mod4.TextMatrix(1, 2)) = "" Then Exit Function
            For li_Row = 1 To N_Reglones
                S_ArmaQryEje = "Insert into t_caratulas values (" & _
                Trim(.Txt_CarProyecto) & "," & Trim(.Lbl_CveEtapa) & "," & N_Modulo & "," & _
                Val(Trim(.Grd_Mod4.TextMatrix(li_Row, 2))) & ",'" & Trim(.Grd_Mod4.TextMatrix(li_Row, 1)) & _
                "'," & Val(li_Row) & ",'L',0)"
                
                gcn.Execute S_ArmaQryEje
            Next
            S_ComBita = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo)
            S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (" & Trim(N_Reglones) & " Reg)"
            Call PR_GrabaBitacora(6, 1, S_DescripBita, S_ComBita)
        End With
        If N_Reglones > 0 Then FU_GuardaInf_GridCaratulaBD = S_DescripBita & vbCrLf
        
    Case "MODULO5-DEF"
      
      
      
      
      
      
      
      
      
                
      
      
      
      
      
      
      
        
    Case "MODULO5"
        N_Modulo = 4
        With Frm_Caratula
            N_Reglones = Val(.Grd_Mod5.Rows - 1)
            If N_Reglones = 1 And Trim(.Grd_Mod5.TextMatrix(1, 1)) = "" And Trim(.Grd_Mod5.TextMatrix(1, 2)) = "" Then Exit Function
            For li_Row = 1 To N_Reglones
                S_ArmaQryEje = "Insert into t_caratulasDet values (" & _
                Trim(.Txt_CarProyecto) & "," & Trim(.Lbl_CveEtapa) & "," & N_Modulo & "," & _
                Val(Trim(.Grd_Mod5.TextMatrix(li_Row, 4))) & "," & Trim(.Grd_Mod5.TextMatrix(li_Row, 3)) & ",'" & Trim(.Grd_Mod5.TextMatrix(li_Row, 2)) & _
                "'," & Val(li_Row) & ",0)"
                
                gcn.Execute S_ArmaQryEje
            Next
            N_Modulo = 5
            S_ComBita = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo)
            S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (" & Trim(N_Reglones) & " Reg)"
            Call PR_GrabaBitacora(6, 1, S_DescripBita, S_ComBita)
        End With
        If N_Reglones > 0 Then FU_GuardaInf_GridCaratulaBD = S_DescripBita & vbCrLf
        
    Case "MODULO6"
        N_Modulo = 6
        N_Reglones = FU_GrabaInformacionChecks("Modulo6")
        If N_Reglones > 0 Then
            S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (" & Trim(N_Reglones) & " Reg)"
            FU_GuardaInf_GridCaratulaBD = S_DescripBita & vbCrLf
        End If
        
    Case "MODULO7"
        N_Modulo = 7
        N_Reglones = FU_GrabaInformacionChecks("Modulo7")
        If N_Reglones > 0 Then
            S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (" & Trim(N_Reglones) & " Reg)"
            FU_GuardaInf_GridCaratulaBD = S_DescripBita & vbCrLf
        End If
    
    Case "MODULO8"
        N_Modulo = 8
        Bo_EntraCiclo = True
        N_Reglones = 0
        
        With Frm_Caratula
            If .Chk_CarMod8(3).Value = 1 Then
                N_Reglones = Val(.Grd_Mod8.Rows - 1)
                If N_Reglones = 1 And Trim(.Grd_Mod8.TextMatrix(1, 1)) = "" And Trim(.Grd_Mod8.TextMatrix(1, 2)) = "" Then Bo_EntraCiclo = False
                If Bo_EntraCiclo Then
                    For li_Row = 1 To N_Reglones
                        S_ArmaQryEje = "Insert into t_caratulas values (" & _
                        Trim(.Txt_CarProyecto) & "," & Trim(.Lbl_CveEtapa) & "," & N_Modulo & "," & _
                        Val(Trim(.Grd_Mod8.TextMatrix(li_Row, 2))) & ",'" & Trim(.Grd_Mod8.TextMatrix(li_Row, 1)) & _
                        "'," & Val(li_Row) & ",'L',0)"
                        
                        gcn.Execute S_ArmaQryEje
                    Next
                End If
            End If
            
            N_RegMod8 = FU_GrabaInformacionChecks("Modulo8")
            N_Reglones = N_Reglones + N_RegMod8
            
            S_ComBita = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo)
            S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (" & Trim(N_Reglones) & " Reg)"
            Call PR_GrabaBitacora(6, 1, S_DescripBita, S_ComBita)
            S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (" & Trim(N_Reglones) & " Reg)"
        End With
        If N_Reglones > 0 Then FU_GuardaInf_GridCaratulaBD = S_DescripBita & vbCrLf
        
     Case "MODULO9"
        N_Modulo = 9
         Bo_EntraCiclo = True
        N_Reglones = 0
        
        With Frm_Caratula
            If .Chk_CarMod9(9).Value = 1 Then
                N_Reglones = Val(.Grd_Mod9.Rows - 1)
                If N_Reglones = 1 And Trim(.Grd_Mod9.TextMatrix(1, 1)) = "" And Trim(.Grd_Mod9.TextMatrix(1, 2)) = "" Then Bo_EntraCiclo = False
                If Bo_EntraCiclo Then
                    For li_Row = 1 To N_Reglones
                        S_ArmaQryEje = "Insert into t_caratulas values (" & _
                        Trim(.Txt_CarProyecto) & "," & Trim(.Lbl_CveEtapa) & "," & N_Modulo & "," & _
                        Val(Trim(.Grd_Mod9.TextMatrix(li_Row, 2))) & ",'" & Trim(.Grd_Mod9.TextMatrix(li_Row, 1)) & _
                        "'," & Val(li_Row) & ",'L',0)"
                       
                        gcn.Execute S_ArmaQryEje
                    Next
                End If
            End If
            
            N_RegMod9 = FU_GrabaInformacionChecks("Modulo9")
            N_Reglones = N_Reglones + N_RegMod9
            
            S_ComBita = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo)
            S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (" & Trim(N_Reglones) & " Reg)"
            Call PR_GrabaBitacora(6, 1, S_DescripBita, S_ComBita)
            S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (" & Trim(N_Reglones) & " Reg)"
        End With
        
        If N_Reglones > 0 Then FU_GuardaInf_GridCaratulaBD = S_DescripBita & vbCrLf
        
End Select

Exit Function

ErrAccionGGrid:
If Err.Number <> 0 Then
    MsgBox "Ha ocurrido un error al guardar los datos", 0 + 32, "[FU_GuardaInf_GridCaratulaBD]"
    Frm_Caratula.MousePointer = 0
    Exit Function
End If
End Function

Function Fu_LeeDatosArchConfig(I_IdLinea, S_CualArchivo) As String
Dim S_RutaIni As String, S_MiArchivo As String, I_Linea As Integer

If UCase(Trim(S_CualArchivo)) = "CENTRAL" Then
    S_RutaIni = Trim(S_Path) & "\conecta.ini"
ElseIf UCase(Trim(S_CualArchivo)) = "LOCAL" Then
  
End If

S_MiArchivo = Dir(S_RutaIni, vbArchive)
If S_MiArchivo = "" Then
    Fu_LeeDatosArchConfig = ""
Else
    I_Linea = 0
    Open S_RutaIni For Input As #1
    Do While Not EOF(1)
        I_Linea = I_Linea + 1
        Line Input #1, S_Contenido
        If I_Linea = I_IdLinea Then Exit Do
    Loop
    Close #1
    
    If Len(Trim(S_Contenido)) > 0 Then
        S_Contenido = UCase(Left(S_Contenido, 1)) & LCase(Right(S_Contenido, Len(Trim(S_Contenido)) - 1))
    End If
    Fu_LeeDatosArchConfig = S_Contenido
End If
'********************************************************************************************
'*Rescata nombre de la base de datos, de acuerdo a la posición del número de linea
'*3 - Nombre de la base de Datos
'*4 - Dirección IP
'********************************************************************************************
End Function
Function FU_DatosServerExt() As Boolean
Dim S_CadenaconExt As String, S_ServerExt As String, S_BaseDatosExt As String
Dim S_LogExt As String, S_PassExt As String

FU_DatosServerExt = False

S_LogExt = Fu_LeeDatosArchConfig(1, "Central")
S_PassExt = Fu_LeeDatosArchConfig(2, "Central")
S_BaseDatosExt = Fu_LeeDatosArchConfig(3, "Central")
S_ServerExt = Fu_LeeDatosArchConfig(4, "Central")
'-----------------------------------------------
'S_LogExt = FUsDeCodifica(Trim(S_LogExt))
'S_PassExt = FUsDeCodifica(Trim(S_PassExt))
'--------------------------------------------------------------------------------------------
S_CadenaconExt = "User ID=" & S_LogExt & ";Password=" & S_PassExt & ";Data Source=" & S_ServerExt & ";Initial Catalog=" & S_BaseDatosExt
If Not FUConecta(S_CadenaconExt) Then Exit Function
FU_DatosServerExt = True
'*******************************************************************************************
'*Arma la cadena de conexión
'*******************************************************************************************
End Function

Function FU_RescataInfX_CampoS(S_Campo As String, S_Tabla As String, S_CondiParam As String) As String
Dim S_Condi As String, RE_InformNumX As ADODB.Recordset

On Error GoTo ErrRecataInfNX
FU_RescataInfX_CampoS = ""
S_Condi = "Select " & S_Campo & " From " & S_Tabla & " Where " & S_CondiParam

Set RE_InformNumX = New ADODB.Recordset
RE_InformNumX.Open S_Condi, gcn

If Not RE_InformNumX.EOF() Then
    FU_RescataInfX_CampoS = Trim(RE_InformNumX(0))
End If
RE_InformNumX.Close
Set RE_InformNumX = Nothing
Exit Function

ErrRecataInfNX:
If Err.Number <> 0 Then
    MsgBox Str(Err.Number) & ": " & Err.Description, 0 + 64, "FU_RescataInfX_CampoS"
    Exit Function
End If
'************************************************************************************************************************
'*Rescata información de un registro especifico dada una condición, de alguna tabla X
'***********************************************************************************************************************
End Function

Function FU_TipoCaraTele(S_NumProyectoVal, N_EtapaVal) As Integer
Dim S_Condi As String, RE_SumaCaraTele As ADODB.Recordset

On Error GoTo ErrRecataInfNX
FU_TipoCaraTele = 0
S_Condi = "Select Sum(c.n_interno) From t_caratulas c (Nolock) " & _
"Where c.n_cveproyecto = " & S_NumProyectoVal & " And c.n_cveetapa = " & N_EtapaVal & " And c.n_modulo = 9 And c.c_tipo_opcion = 'C'"

Set RE_SumaCaraTele = New ADODB.Recordset
RE_SumaCaraTele.Open S_Condi, gcn

If Not RE_SumaCaraTele.EOF() Then
    If IsNull(RE_SumaCaraTele(0)) Then
        FU_TipoCaraTele = 0
    Else
        FU_TipoCaraTele = Val(RE_SumaCaraTele(0))
    End If
End If
RE_SumaCaraTele.Close
Set RE_SumaCaraTele = Nothing
Exit Function

ErrRecataInfNX:
If Err.Number <> 0 Then
    MsgBox Str(Err.Number) & ": " & Err.Description, 0 + 64, "FU_TipoCaraTele"
    Exit Function
End If
End Function

Function FU_ValorCeldaEfectiva() As Integer
Dim N_Col_NE As Integer, N_Efectiva

Frm_Resumen.MSFlexGrid1.Col = 2
FU_ValorCeldaEfectiva = 0
With Frm_Resumen
    N_Col_NE = 0
    For i = 1 To Val(.MSFlexGrid1.Cols - 1)
        If UCase(Trim(.MSFlexGrid1.TextMatrix(0, i))) = "EFECTIVA" Then
            N_Col_NE = i
            Exit For
        End If
    Next i
    If Val(N_Col_NE) = 0 Then Exit Function
    N_Efectiva = IIf(IsNumeric(.MSFlexGrid1.TextMatrix(2, N_Col_NE)), Val(.MSFlexGrid1.TextMatrix(2, N_Col_NE)), 0)
    If Val(N_Efectiva) <> 1 Then
        MsgBox "El valor de la celda Efectiva, debe ser 1.", 0 + 16, "Verificar"
        .MSFlexGrid1.SetFocus
        FU_ValorCeldaEfectiva = -1
        Frm_Resumen.MSFlexGrid1.Col = N_Col_NE
    End If
End With
End Function

Public Function FUConecta(sconexion, Optional S_IP As String = "99.99.99.99") As Boolean
On Error GoTo ErrorConexion
Dim bynum As Byte

bynum = 0

With gcn
    .CursorLocation = adUseClient
    .CommandTimeout = 50
    .Provider = "SQLOLEDB"
    .ConnectionString = sconexion
    .Open
End With
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

FUConecta = True
Exit Function

ErrorConexion:
    N_CuentaSinConec = N_CuentaSinConec + 1
    If InStr(Err.Description, "08001") > 0 Or (InStr(Err.Description, "01000") > 0 And InStr(Err.Description, "gethostbyname()()") > 0) Then
        MsgBox "No se puede establecer la comunicación." & vbCrLf & "Verifique la configuración de la red y/o que el servidor este activo o en su defecto que exista.", vbOKOnly, "Error en la conexión, intente nuevamente."
    ElseIf InStr(Err.Description, "01000") > 0 And InStr(Err.Description, "connect()") > 0 Then
        MsgBox "No se puede establecer la comunicación. Verifique la conexión física de la red.", vbOKOnly, "Error en la conexión, intente nuevamente."
    ElseIf InStr(Err.Description, "08004") > 0 Or (InStr(Err.Description, "01000") > 0 And InStr(Err.Description, "Changed database") > 0) Then
        MsgBox "No se localizó la base datos. Verifique el nombre de la base de datos que capturó.", vbOKOnly, "Error en la conexión, intente nuevamente."
    ElseIf InStr(Err.Description, "28000") > 0 And InStr(Err.Description, "Login failed") > 0 Then
        MsgBox "Error en la contraseña. Verifique la información que capturó.", vbOKOnly, "Error en la conexión, intente nuevamente."
    ElseIf (InStr(Err.Description, "Time") > 0 Or InStr(Err.Description, "Tiempo") > 0) And InStr(Err.Description, "S1T00") > 0 And bynum <= 2 Then
        bynum = bynum + 1
        Resume
    Else
        If Val(Err.Number) = -2147217843 Then
            MsgBox "No se realizó la conexión. Descripción del error: " & Err.Description & " IP: " & S_IP, 0 + 16, "Verificar el usuario en la local"
        Else
            MsgBox "No se realizó la conexión. Descripción del error: " & Err.Description & " IP: " & S_IP, 0 + 16, "Error en la conexión, intente nuevamente."
        End If
    End If
'*******************************************************************************************
'*Sirve para conectarse al servidor central
'*******************************************************************************************
End Function
Sub PR_Desconecta()
        If gcn.State = 1 Then
            gcn.Close
            Set gcn = Nothing
        End If
'***********************************************************************************************************************
'*Sirve para desconectarse del servidor central
'***********************************************************************************************************************

End Sub


Function FU_VerificaArchivoConfig(S_CualArch As String) As Boolean
Dim S_Rutita As String

FU_VerificaArchivoConfig = False

S_Rutita = S_Path & "\" & Trim(S_CualArch)
Select Case UCase(Trim(S_CualArch))
    Case "CONECTA.INI"
        If Dir(S_Rutita, vbArchive) = "" Then
            MsgBox "No se encuentra el archivo de configuración. No se puede accesar al servidor central.", 0 + 16, "Falta archivo de configuración"
        End If
    
    Case Else
        MsgBox "Archivo no cotemplado para la configuración [" & Trim(S_CualArch) & "]", 0 + 16, "Error Interno de lógica"
        Exit Function
End Select
FU_VerificaArchivoConfig = True
'********************************************************************************************
'*Checa si existen los archivos para el uso adecuado de la aplicación
'********************************************************************************************
End Function
Function FU_ExtraeFechaServer() As Date
Dim S_Condi As String, RE_Fecha As ADODB.Recordset

FU_ExtraeFechaServer = "01/01/1900 23:59:59"
S_Condi = "SELECT getdate() as F_Ser"
Set RE_Fecha = New ADODB.Recordset
RE_Fecha.Open S_Condi, gcn

If Not RE_Fecha.EOF() Then
    FU_ExtraeFechaServer = RE_Fecha(0)
End If
RE_Fecha.Close
'*******************************************************************************************
'*Rescata la fecha del servidor
'*******************************************************************************************
End Function
Function FU_Cuenta_Registros(ps_Tabla As String, S_Condi As String) As Long
Dim RE_TablaX As ADODB.Recordset, sConsec As String

On Error GoTo ErrNvaClave
sConsec = "SELECT Count(*) FROM  " & ps_Tabla & " (Nolock) Where " & S_Condi

Set RE_TablaX = New ADODB.Recordset
RE_TablaX.Open sConsec, gcn

If Val(RE_TablaX.RecordCount) > 0 Then
      FU_Cuenta_Registros = IIf(IsNull(RE_TablaX(0)), 0, RE_TablaX(0))
Else
      FU_Cuenta_Registros = 0
End If
RE_TablaX.Close

Exit Function
ErrNvaClave:

If Err.Number <> 0 Then
    MsgBox Err.Description, 0 + 64, "Error número: " & Err.Number
    FU_Cuenta_Registros = 0
    Exit Function
End If
'*******************************************************************************************
'*Cuenta resgistros de alguna tabla con cierta condición
'*******************************************************************************************
End Function

Function FU_RescataInfX_CampoN(S_Campo As String, S_Tabla As String, S_CondiParam As String) As Integer
Dim S_Condi As String, RE_InformNumX As ADODB.Recordset

On Error GoTo ErrRecataInfNX
FU_RescataInfX_CampoN = 0
S_Condi = "Select " & S_Campo & " From " & S_Tabla & " Where " & S_CondiParam

Set RE_InformNumX = New ADODB.Recordset
RE_InformNumX.Open S_Condi, gcn

If Not RE_InformNumX.EOF() Then
    FU_RescataInfX_CampoN = RE_InformNumX(0)
End If
RE_InformNumX.Close
Set RE_InformNumX = Nothing
Exit Function

ErrRecataInfNX:
If Err.Number <> 0 Then
    MsgBox Str(Err.Number) & ": " & Err.Description, 0 + 64, "FU_RescataInfX_CampoN"
    Exit Function
End If
'************************************************************************************************************************
'*Rescata información de un registro especifico dada una condición, de alguna tabla X
'***********************************************************************************************************************
End Function
Function FU_ValidaInforSeg(S_Pantalla As String) As Boolean
Dim I_Ren As Integer, S_ValorX As Variant, li_Contrl As Integer, S_Micondicion As String

FU_ValidaInforSeg = True

Select Case UCase(Trim(S_Pantalla))
    Case "SEGURIDAD"
        With Frm_Seguridad
            If Len(Trim(.Txt_SegCta)) = 0 Then
                MsgBox "Es necesario el dato de la cuenta.", vbInformation, "Validación"
                .Txt_SegCta.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Len(Trim(.Txt_SegCon)) = 0 Then
                MsgBox "Es necesario el dato de la contraseña.", vbInformation, "Validación"
                .Txt_SegCon.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
        End With
        
    Case "CAT_ETAPAS"
        With Frm_Etapas
            If Len(Trim(.Txt_EtaCveEta)) = 0 Then
                MsgBox "Es necesario el dato de la clave.", vbInformation, "Validación"
                .Txt_EtaCveEta.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Len(Trim(.Txt_EtaCveEta)) = 0 Then
                MsgBox "La clave de la etapa es hasta 3 dígitos.", vbInformation, "Validación"
                .Txt_EtaCveEta.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Val(Trim(.Txt_EtaCveEta)) > 255 Then
                MsgBox "La clave de la etapa no debe ser mayor a 255.", vbInformation, "Validación"
                .Txt_EtaCveEta.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Len(Trim(.Txt_EtaDescrip)) = 0 Then
                MsgBox "Es necesario el dato de la descripción de la Etapa.", vbInformation, "Validación"
                .Txt_EtaDescrip.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
       End With
       
    Case "CAT_PERFILES"
        With Frm_Perfiles
            If Len(Trim(.Txt_PerCvePer)) = 0 Then
                MsgBox "Es necesario el dato de la clave.", vbInformation, "Validación"
                .Txt_PerCvePer.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Len(Trim(.Txt_PerCvePer)) > 3 Then
                MsgBox "La clave del perfil es hasta 3 dígitos.", vbInformation, "Validación"
                .Txt_PerCvePer.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Val(Trim(.Txt_PerCvePer)) > 255 Then
                MsgBox "La clave del perfil no debe ser mayor a 255", vbInformation, "Validación"
                .Txt_PerCvePer.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Len(Trim(.Txt_PerDescrip)) = 0 Then
                MsgBox "Es necesario el dato de la descripción del perfil.", vbInformation, "Validación"
                .Txt_PerDescrip.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
        End With
        
    Case "CAT_USUARIOS"
        With Frm_Usuarios
            If Len(Trim(.Txt_UsuClave)) = 0 Then
                MsgBox "Es necesario el dato de la clave.", vbInformation, "Validación"
                .Txt_UsuClave.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Len(Trim(.Txt_UsuNombre)) = 0 Then
                MsgBox "Es necesario el dato del nombre.", vbInformation, "Validación"
                .Txt_UsuNombre.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If

            If Len(Trim(.Txt_UsuPaterno)) = 0 Then
                MsgBox "Es necesario el dato del apellido paterno.", vbInformation, "Validación"
                .Txt_UsuPaterno.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If

            If Len(Trim(.Txt_UsuMaterno)) = 0 Then
                MsgBox "Es necesario el dato del apellido materno.", vbInformation, "Validación"
                .Txt_UsuMaterno.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If

            If Len(Trim(.Txt_UsuContra)) = 0 Then
                MsgBox "Es necesario el dato de la contraseña.", vbInformation, "Validación"
                .Txt_UsuContra.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Len(Trim(.Txt_UsuContra)) < 4 Then
                MsgBox "El dato de la contraseña, debe tener mínimo 4 carácteres.", vbInformation, "Validación"
                .Txt_UsuContra.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Trim(.Txt_UsuContra) <> Trim(.Txt_UsuContra.Tag) Then
                MsgBox "La contraseña no es igual a la contraseña confirmada.", vbInformation, "Validación"
                .Txt_UsuContra.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If

            If Len(Trim(.Cmb_UsuPerfil)) = 0 Then
                MsgBox "Es necesario selccionar un perfil.", vbInformation, "Validación"
                .Cmb_UsuPerfil.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
        End With
        
    Case "PROYECTOS"
        With Frm_Proyecto
            If Len(Trim(.Txt_ProyNumero)) = 0 Then
                MsgBox "Es necesario el dato de la clave del nombre del proyecto.", vbInformation, "Validación"
                .Txt_ProyNumero.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Len(Trim(.Txt_ProyNombre)) = 0 Then
                MsgBox "Es necesario el dato del nombre del proyecto.", vbInformation, "Validación"
                .Txt_ProyNombre.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Len(Trim(.Cmb_ProyEjecutivoSC)) = 0 Then
                MsgBox "Es necesario seleccionar un ejecutivo de servicio al cliente.", vbInformation, "Validación"
                .Cmb_ProyEjecutivoSC.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Len(Trim(.Cmb_ProyEjecutivoPR)) = 0 Then
                MsgBox "Es necesario seleccionar un ejecutivo de procesamiento.", vbInformation, "Validación"
                .Cmb_ProyEjecutivoPR.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Len(Trim(.Cmb_ProyProg)) = 0 Then
                MsgBox "Es necesario seleccionar un programador.", vbInformation, "Validación"
                .Cmb_ProyProg.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Len(Trim(.Txt_ProyProd)) = 0 Then
                MsgBox "Es necesario el dato del nivel de productividad.", vbInformation, "Validación"
                .Txt_ProyProd.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Len(Trim(.Cmb_ProyEstatus)) = 0 Then
                MsgBox "Es necesario seleccionar un estatus.", vbInformation, "Validación"
                .Cmb_ProyEstatus.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
        End With
        
    Case "CARATULA-ENCABEZADO"
        With Frm_Caratula
            If Val(.Lbl_CarCveProyecto.Caption) = 0 Then
                MsgBox "Es necesario seleccionar un proyecto válido para la definir la etapa.", vbInformation, "Validación"
                .Txt_CarProyecto.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Val(.Lbl_CveEtapa.Caption) = 0 Then
                MsgBox "Es necesario seleccionar una etapa válida.", vbInformation, "Validación"
                .Cmb_CarEtapa.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Val(.Lbl_CarCveEstatus.Caption) = 0 Then
                MsgBox "Es necesario seleccionar un estatus para la etapa.", vbInformation, "Validación"
                .Cmb_CarEstatus.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
        End With
    
    Case "CARATULA-EDAD"
        With Frm_Caratula
            If .Chk_CarMod8(3).Value = 1 Then
                I_Ren = Val(.Grd_Mod8.Rows) - 1
                If I_Ren = 1 Then
                    If .Grd_Mod8.TextMatrix(I_Ren, 1) = "" And .Grd_Mod8.TextMatrix(I_Ren, 2) = "" Then
                        MsgBox "Cuando se selecciona la Edad, es necesario que al menos tenga una opción.", vbInformation, "Validación módulo 8"
                        FU_ValidaInforSeg = False
                        Exit Function
                    End If
                End If
            End If
        End With
        
    Case "CARATULA-CUOTAS"
        If FU_ValidaExistenciaCuotasDefinidas = -1 Then
            MsgBox "Cuando se tienen cuotas definidas, al menos debe tener una opción por cuota.", vbInformation, "Validación módulo 5"
            FU_ValidaInforSeg = False
            Exit Function
        End If
        
    Case "CARATULA-NOELEGIBLE"
        With Frm_Caratula
            If .Chk_CarMod9(9).Value = 1 Then
                I_Ren = Val(.Grd_Mod9.Rows) - 1
                If I_Ren = 1 Then
                    If .Grd_Mod9.TextMatrix(I_Ren, 1) = "" And .Grd_Mod9.TextMatrix(I_Ren, 2) = "" Then
                        MsgBox "Cuando se selecciona la opción No Elegible, es necesario que al menos tenga una opción.", vbInformation, "Validación módulo 9"
                        FU_ValidaInforSeg = False
                        Exit Function
                    End If
                End If
            End If
        End With
        
    Case "T_ENCCUESTIONARIO"
         With Frm_Cuestionario
            If Len(Trim(.Cmb_CueProyecto)) = 0 Then
                MsgBox "Es necesario el dato del número del proyecto.", vbInformation, "Validación"
                .Cmb_CueProyecto.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Len(Trim(.Cmb_CueEtapa)) = 0 Then
                MsgBox "Es necesario el dato de la etapa del proyecto.", vbInformation, "Validación"
                .Cmb_CueEtapa.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Len(Trim(.Txt_CueCuestionario)) = 0 Then
                MsgBox "Es necesario el dato del número del cuestionario.", vbInformation, "Validación"
                .Txt_CueCuestionario.SetFocus
                FU_ValidaInforSeg = False
                Exit Function
            End If
            
            If Len(Trim(.Txt_CueCuestionario2)) = 0 Then
                If UCase(Trim(.StatusBarCue.Panels(2).Text)) = "CUESTIONARIO NUEVO" Then
                    MsgBox "Es necesario el dato de la confirmación número del cuestionario.", vbInformation, "Validación"
                  
                    .Txt_CueCuestionario2.Enabled = True
                    .Txt_CueCuestionario2.SetFocus
                    FU_ValidaInforSeg = False
                    Exit Function
                End If
            End If
            
            If Len(Trim(.Txt_CueCuestionario)) > 0 And Len(Trim(.Txt_CueCuestionario2)) > 0 Then
                If Trim(.Txt_CueCuestionario) <> Trim(.Txt_CueCuestionario2) Then
                    MsgBox "El dato de la confirmación número del cuestionario, es diferente.", vbInformation, "Validación"
                    .Txt_CueCuestionario2.SetFocus
                    FU_ValidaInforSeg = False
                    Exit Function
                End If
            End If
            
         End With
         
    Case "BITACORA-CALIDAD"
        With Frm_Auditoria
            For li_Contrl = 1 To Val(.Txt_AudMayor.UBound)
                S_ValorX = .Txt_AudMayor(li_Contrl)
                If Not IsNumeric(S_ValorX) Then
                    FU_ValidaInforSeg = False
                    Exit Function
                End If
            Next
        End With
        
    Case "AUDITORIA-BITACORA"
      
      
      
      
      
      
      
      

End Select
'***********************************************************************************************************************
'*
'***********************************************************************************************************************
End Function
Function FU_ValidaDatosSeguridad() As Boolean
Dim S_Cuenta As String, S_Contrasena As String

FU_ValidaDatosSeguridad = True
With Frm_Seguridad
    S_Cuenta = Trim(.Txt_SegCta)
    S_Contrasena = Trim(.Txt_SegCon)
    If Not FU_ValidaSeguridad_CtaCon("Cuenta", S_Cuenta, S_Contrasena) Then
        MsgBox "La cuenta de usuario que proporcionó no tiene acceso al sistema", vbCritical, "Verifique la cuenta de usuario"
        FU_ValidaDatosSeguridad = False
        .Txt_SegCta.SelStart = 0
        .Txt_SegCta.SelLength = Len(.Txt_SegCta)
        .Txt_SegCta.SetFocus
        Exit Function
    End If
    If Not FU_ValidaSeguridad_CtaCon("Contrasena", S_Cuenta, S_Contrasena) Then
        MsgBox "La contraseña que proporcionó no es correcta y no tiene permisos de acceso.", vbCritical, "Verifique la contraseña"
        FU_ValidaDatosSeguridad = False
        .Txt_SegCon.SelStart = 0
        .Txt_SegCon.SelLength = Len(.Txt_SegCon)
        .Txt_SegCon.SetFocus
        Exit Function
    End If
End With
'***********************************************************************************************************************
'*
'***********************************************************************************************************************
End Function

Function FU_ValidaSeguridad_CtaCon(S_Tipo As String, S_ParamCta As String, S_ParamCon As String) As Boolean
Dim RE_SegUsuarios As ADODB.Recordset, S_Filtro As String

FU_ValidaSeguridad_CtaCon = False
Select Case UCase(Trim(S_Tipo))
    Case "CUENTA"
      
        S_Filtro = "Select n_cveusuario,s_pwd From c_segusuarios Where n_cveusuario = '" & Trim(S_ParamCta) & "'"
        
        Set RE_SegUsuarios = New ADODB.Recordset
        RE_SegUsuarios.Open S_Filtro, gcn
        
        If Not RE_SegUsuarios.EOF() Then
            FU_ValidaSeguridad_CtaCon = True
        End If
        RE_SegUsuarios.Close
        Set RE_SegUsuarios = Nothing
    Case "CONTRASENA"
        S_Filtro = "Select n_cveusuario,s_pwd From c_segusuarios Where n_cveusuario = '" & Trim(S_ParamCta) & "' And s_pwd = '" & S_ParamCon & "'"

        Set RE_SegUsuarios = New ADODB.Recordset
        RE_SegUsuarios.Open S_Filtro, gcn
        
        If Not RE_SegUsuarios.EOF() Then
            FU_ValidaSeguridad_CtaCon = True
        End If
        RE_SegUsuarios.Close
        Set RE_SegUsuarios = Nothing
End Select
'*******************************************************************************************
'*Determina si existe una cuenta o contraseña
'*******************************************************************************************
End Function


Function FU_vte_ValTecla(pintTecla As Integer, _
                            pbytPatron As Byte, _
                            Optional pstrCadena As String) As Integer
'******************************************************************************************
'Función                    : gf_ValTecla
'Autor                      : J.M.F.
'Descripción                : Valida que la tecla pulsada en un control se encuentre en el patrón
'                             especificado. La suma de valores de pbytpatrón determinan buscar en
'                             más de uno
'Fecha de Creación          : 26/Ene/1999
'Fecha de Liberación        : 26/Ene/1999
'Fecha de Modificación      :
'Autor de la Modificación   :
'Usuario que solicita la modificación:
'
' Parámetros:   Tipo        Nombre         Descripción
'    Entrada:   Integer    pintTecla     Tecla pulsada en el control
'               String     pstrCadena    Cadena a validar
'               Byte       pbytPatron    Patrón de valores permitidos de acuerdo a
'                                       1=Letras (sólo mayusculas A..Z)
'                                       3=Numeros (0..9)
'                                       5=String especial (pstrCadena) opcional
'                                       4=Letras y Números
'                                       6=Letras y String especial
'                                       8=Numeros y String especial
'                                       9=Letras,Numeros y String especial
'     salida:   Integer    gfint_vte_ValTecla  Ascci de la tecla pulsada en el control si
'                                              existe en el patrón de búsqueda
'                                              0: En caso contrario
'******************************************************************************************
Dim lintAscii As Integer

'No procesa: comillas o apostrofe retorna 0
   If pintTecla = 34 Or pintTecla = 39 Then
      FU_vte_ValTecla = 0
      Exit Function
   End If
'Valida si faltan parametros si se eligió la opcion de patrón por cadena especial
   If (InStr("5689", CStr(pbytPatron)) > 0) And pstrCadena = "" Then
      MsgBox "Falta el parametro de la cadena", vbOKOnly, "Eligió validación por cadena"
      FU_vte_ValTecla = 0
      Exit Function
   End If
'Valida si se eligió al menos un patrón de busqueda
   If InStr("12345689", CStr(pbytPatron)) <= 0 Then
      MsgBox "No selecciono patrón de búsqueda", vbOKOnly, "Patrón de validación erroneo"
      FU_vte_ValTecla = 0
      Exit Function
   End If
   If pintTecla < vbKeySpace Then
      FU_vte_ValTecla = pintTecla
      Exit Function
   End If
   lintAscii = 0
   If pintTecla >= 65 And pintTecla <= 90 Then
      pintTecla = Asc(UCase(Chr(pintTecla)))
   ElseIf pintTecla >= 97 And pintTecla <= 122 Then
      pintTecla = Asc(LCase(Chr(pintTecla)))
   ElseIf pintTecla = 225 Or pintTecla = 243 Or pintTecla = 237 Or pintTecla = 243 Or pintTecla = 250 Then
      pintTecla = Asc(LCase(Chr(pintTecla)))
   ElseIf pintTecla = 193 Or pintTecla = 201 Or pintTecla = 205 Or pintTecla = 211 Or pintTecla = 218 Then
      pintTecla = Asc(UCase(Chr(pintTecla)))
   ElseIf pintTecla = 46 Or pintTecla = 32 Or pintTecla = 160 Then
      pintTecla = Asc(UCase(Chr(pintTecla)))
   
   End If
 
   If InStr("12469", CStr(pbytPatron)) > 0 Then lintAscii = FU_vte_EsLetra(pintTecla)
   If InStr("32489", CStr(pbytPatron)) > 0 And lintAscii = 0 Then lintAscii = FU_vte_EsNumero(pintTecla)
   If InStr("52689", CStr(pbytPatron)) > 0 And lintAscii = 0 Then lintAscii = FU_vte_EnCadena(pintTecla)
   If lintAscii >= 48 And lintAscii <= 57 Then FU_vte_ValTecla = lintAscii
   If lintAscii >= 65 And lintAscii <= 90 Then
      FU_vte_ValTecla = Asc(UCase(Chr(lintAscii)))
   ElseIf lintAscii >= 97 And lintAscii <= 122 Then
      FU_vte_ValTecla = Asc(LCase(Chr(lintAscii)))
   ElseIf lintAscii = 225 Or lintAscii = 233 Or lintAscii = 237 Or lintAscii = 243 Or lintAscii = 250 Then
      FU_vte_ValTecla = Asc(LCase(Chr(lintAscii)))
   ElseIf lintAscii = 193 Or lintAscii = 201 Or lintAscii = 205 Or lintAscii = 211 Or lintAscii = 218 Then
      FU_vte_ValTecla = Asc(UCase(Chr(lintAscii)))
   ElseIf lintAscii = 46 Or lintAscii = 32 Or lintAscii = 160 Then
      FU_vte_ValTecla = Asc(LCase(Chr(lintAscii)))

  
   End If
   
End Function
Sub PR_AcomodaDatosPerfiles()
'*PERFILES*
On Error GoTo ErrorAviso

    If RE_Perfiles.EOF = True And RE_Perfiles.BOF = True Then
        MsgBox "No existen más registros ", vbInformation, "Perfiles"
        Exit Sub
    Else
        With Frm_Perfiles
            .Txt_PerCvePer = RE_Perfiles(0)
            .Txt_PerDescrip = Trim(RE_Perfiles(1))
            .Txt_PerDescripLar = IIf(IsNull(RE_Perfiles(2)), "", Trim(RE_Perfiles(2)))
            
            .Cmb_PerEstatus = Trim(RE_Perfiles(4))
            .Lbl_PerCveEstatus.Caption = RE_Perfiles(3)
            
            .Cmb_PerEstatus.ListIndex = FU_RetornaIndiceCombo(.Cmb_PerEstatus)
            
            .txt_TotRegPerfiles = "    " & RE_Perfiles.AbsolutePosition & "/" & RE_Perfiles.RecordCount
            .Lbl_RegistroActualPerfiles.Caption = RE_Perfiles.AbsolutePosition
        End With
          
    End If
    
    Exit Sub
ErrorAviso:
     If Error = "No hay ningún registro activo." Then
        MsgBox "No hay más registros", vbInformation, "Perfiles"
     Else
        MsgBox "Descripción del error: " & Err.Description, vbCritical, "Error Número: " & Err.Number
     End If
     Exit Sub
End Sub

Sub PR_AcomodaDatosEtapas()
'*ETAPAS*
On Error GoTo ErrorAviso

    If RE_Etapas.EOF = True And RE_Etapas.BOF = True Then
        MsgBox "No existen más registros ", vbInformation, "Etapas"
        Exit Sub
    Else
        With Frm_Etapas
            .Txt_EtaCveEta = RE_Etapas(0)
            .Txt_EtaDescrip = Trim(RE_Etapas(1))
            .Txt_EtaDescripLar = IIf(IsNull(RE_Etapas(2)), "", Trim(RE_Etapas(2)))
            
            .Cmb_EtaEstatus = Trim(RE_Etapas(4))
            .Lbl_EtaCveEstatus.Caption = RE_Etapas(3)
            
            .Cmb_EtaEstatus.ListIndex = FU_RetornaIndiceCombo(.Cmb_EtaEstatus)
            
            .txt_TotRegEtapas = "    " & RE_Etapas.AbsolutePosition & "/" & RE_Etapas.RecordCount
            .Lbl_RegistroActualEtapas.Caption = RE_Etapas.AbsolutePosition
        End With
          
    End If
    
    Exit Sub
ErrorAviso:
     If Error = "No hay ningún registro activo." Then
        MsgBox "No hay más registros", vbInformation, "Etapas"
     Else
        MsgBox "Descripción del error: " & Err.Description, vbCritical, "Error Número: " & Err.Number
     End If
     Exit Sub
End Sub
Sub PR_AcomodaDatosProyectos()
'*PROYECTOS *
On Error GoTo ErrorAvisoPro

    If RE_Proyectos.EOF = True And RE_Proyectos.BOF = True Then
        MsgBox "No existen más registros ", vbInformation, "Proyectos"
        Exit Sub
    Else
        With Frm_Proyecto
            .Txt_ProyNumero = RE_Proyectos(0)
            .Txt_ProyNombre = Trim(RE_Proyectos(1))
            
            .Cmb_ProyEjecutivoSC = FU_ExtraeNombreUsuario(Trim(RE_Proyectos(3)))
            .Txt_ProyNumEjecutivoSC = Trim(RE_Proyectos(3))
            .Cmb_ProyEjecutivoPR = FU_ExtraeNombreUsuario(Trim(RE_Proyectos(4)))
            .Txt_ProyNumEjecutivoPR = Trim(RE_Proyectos(4))
            .Cmb_ProyProg = FU_ExtraeNombreUsuario(Trim(RE_Proyectos(5)))
            .Txt_ProyNumProg = Trim(RE_Proyectos(5))
            
            .Txt_ProyProd = Trim(RE_Proyectos(6))
            .Txt_ProyComen = Trim(RE_Proyectos(7))
            .Cmb_ProyEstatus = Trim(RE_Proyectos(8))
            .Lbl_ProyCveEstatus.Caption = Trim(RE_Proyectos(2))
            
            .Cmb_ProyEjecutivoSC.ListIndex = FU_RetornaIndiceCombo(.Cmb_ProyEjecutivoSC)
            .Cmb_ProyEjecutivoPR.ListIndex = FU_RetornaIndiceCombo(.Cmb_ProyEjecutivoPR)
            .Cmb_ProyProg.ListIndex = FU_RetornaIndiceCombo(.Cmb_ProyProg)

            .txt_TotRegProy = "    " & RE_Proyectos.AbsolutePosition & "/" & RE_Proyectos.RecordCount
            .Lbl_RegistroActualProy.Caption = RE_Proyectos.AbsolutePosition
        End With
    End If
    
    Exit Sub
ErrorAvisoPro:
     If Error = "No hay ningún registro activo." Then
        MsgBox "No hay más registros", vbInformation, "Proyectos"
     Else
        MsgBox "Descripción del error: " & Err.Description, vbCritical, "Error Número: " & Err.Number
     End If
     Exit Sub
End Sub
Sub PR_AcomodaDatosUsuarios()
Dim i As Long, adors As New ADODB.Recordset
'*USUARIOS *
On Error GoTo ErrorAvisoUsu

    If RE_Usuarios.EOF = True And RE_Usuarios.BOF = True Then
        MsgBox "No existen más registros ", vbInformation, "Perfiles"
        Exit Sub
    Else
        With Frm_Usuarios
            .Txt_UsuClave = RE_Usuarios(0)
            .Txt_UsuNombre = Trim(RE_Usuarios(1))
            .Txt_UsuPaterno = Trim(RE_Usuarios(2))
            .Txt_UsuMaterno = Trim(RE_Usuarios(3))
            .Txt_UsuContra = Trim(RE_Usuarios(4))
            .Txt_UsuContra.Tag = Trim(RE_Usuarios(4))
            .Cmb_UsuPerfil = Trim(RE_Usuarios(8))
            .Lbl_UsuCvePerfil.Caption = Trim(RE_Usuarios(5))
            
            .Cmb_UsuEstatus = Trim(RE_Usuarios(10))
            .Lbl_UsuCveEstatus.Caption = Trim(RE_Usuarios(9))

            .Cmb_UsuPerfil.ListIndex = FU_RetornaIndiceCombo(.Cmb_UsuPerfil)
            .Cmb_UsuEstatus.ListIndex = FU_RetornaIndiceCombo(.Cmb_UsuEstatus)
            
            .txt_TotRegUsuarios = "    " & RE_Usuarios.AbsolutePosition & "/" & RE_Usuarios.RecordCount
            .Lbl_RegistroActualUsu.Caption = RE_Usuarios.AbsolutePosition
            
          
            i = BuscaCombo(.ComboVarios(0), IIf(IsNull(RE_Usuarios!n_cvepuesto), 0, RE_Usuarios!n_cvepuesto), True)
            If i >= 0 Then
                .ComboVarios(0).ListIndex = i
            Else
                .ComboVarios(0).ListIndex = -1
            End If
'            i = IIf(IsNull(RE_Usuarios!n_cvepuesto), 0, RE_Usuarios!n_cvepuesto)
'            If i > 0 Then
'                If adors.State Then adors.Close
'                adors.Open "select case when s_clave is null then '---' else s_clave end+' ('+convert(nvarchar,n_cvepuesto)+')' from t_rhpuestos where n_cvepuesto=" & i, gConSql, adOpenStatic, adLockReadOnly
'                If Not adors.EOF Then
'                    .txtCampo(0).Text = adors(0)
'                Else
'                    .txtCampo(0).Text = ""
'                End If
'            Else
'                .txtCampo(0).Text = ""
'            End If
'            .txtCampo(0).Tag = i
'            i = IIf(IsNull(RE_Usuarios!n_cveempleado), 0, RE_Usuarios!n_cveempleado)
'            If i > 0 Then
'                If adors.State Then adors.Close
'                adors.Open "select right('0'+datename(d,f_alta),2)+'/'+right('0'+convert(nvarchar,month(f_alta)),2)+'/'+datename(yy,f_alta)+' ('+convert(nvarchar,n_cveempleado)+')' from t_rhempleados where n_cveempleado=" & i, gConSql, adOpenStatic, adLockReadOnly
'                If Not adors.EOF Then
'                    .txtCampo(1).Text = adors(0)
'                Else
'                    .txtCampo(1).Text = ""
'                End If
'            Else
'                .txtCampo(1).Text = ""
'            End If
'            .txtCampo(1).Tag = i
'            i = IIf(IsNull(RE_Usuarios!n_cvepersona), 0, RE_Usuarios!n_cvepersona)
'            If i > 0 Then
'                If adors.State Then adors.Close
'                adors.Open "select dbo.f_responsable(" & i & ")", gConSql, adOpenStatic, adLockReadOnly
'                If Not adors.EOF Then
'                    .txtCampo(2).Text = adors(0) & " (" & i & ")"
'                Else
'                    If i > 0 Then
'                        .txtCampo(2).Text = "(" & i & ")"
'                    Else
'                        .txtCampo(2).Text = ""
'                    End If
'                End If
'            Else
'                .txtCampo(2).Text = ""
'            End If
'            .txtCampo(2).Tag = i
        End With
    End If
    
    Exit Sub
ErrorAvisoUsu:
     If Error = "No hay ningún registro activo." Then
        MsgBox "No hay más registros", vbInformation, "Usuarios"
     Else
        MsgBox "Descripción del error: " & Err.Description, vbCritical, "Error Número: " & Err.Number
     End If
     Resume
     Exit Sub
End Sub

Sub PR_ActualizaInf_GridCaratula(S_CualModulo)
Dim N_CtlsChecks As Integer, i As Integer, S_CondiEje As String, N_Modulo As Integer, N_Reglones As Integer
Dim S_ComBitaLiga As String, S_DescripBita As String, S_Mov As String, S_Where As String, N_RegBorraM8 As Integer
Dim S_Micondicion As String, Bo_Edad As Boolean, Bo_NoElegible As Boolean, S_CadOption As String
Dim S_MicondiAct As String, S_OpcionCheckApre As String, S_EtiquetaCheckApreAmai As String, N_TipoCarTel As Integer

On Error GoTo GrabaChecksAct

Bo_Edad = False
Select Case UCase(Trim(S_CualModulo))
    Case "MODULO6"
        N_Reglones = 0
        N_Modulo = 6
        With Frm_Caratula
            For i = 0 To Val(.Chk_CarMod6.UBound)
                If Val(.Chk_CarMod6(i).Value) <> Val(.Chk_CarMod6(i).Tag) Then
                    If Val(.Chk_CarMod6(i).Value) > Val(.Chk_CarMod6(i).Tag) Then
                        S_Mov = "A"
                        .Chk_CarMod6(i).Tag = 1
                      
                        
                        S_CondiEje = "Insert into t_caratulas values (" & _
                        Trim(.Txt_CarProyecto) & "," & Trim(.Lbl_CveEtapa.Caption) & "," & N_Modulo & "," & _
                        IIf(Val(i) > 0, Val(i) * -1, Val(i)) & ",'" & Trim(.Chk_CarMod6(i).Caption) & _
                        "',Null,'C',0)"
                        
                        gcn.Execute S_CondiEje
                    ElseIf Val(.Chk_CarMod6(i).Value) < Val(.Chk_CarMod6(i).Tag) Then
                        S_Mov = "B"
                        .Chk_CarMod6(i).Tag = 0
                      
                        
                        S_Where = "Where n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And n_cveetapa = " & Trim(.Lbl_CveEtapa.Caption) & " And n_modulo = " & N_Modulo & " And n_opcion_base = " & IIf(Val(i) > 0, Val(i) * -1, Val(i))
                        S_CondiEje = "Delete From t_caratulas " & S_Where
                        
                        gcn.Execute S_CondiEje
                    End If
                        
                    If S_Mov = "A" Then
                        S_ComBitaLiga = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo) & "-" & Trim(IIf(Val(i) > 0, Val(i) * -1, Val(i)))
                        S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (1 Reg)"
                        Call PR_GrabaBitacora(6, 1, S_DescripBita, S_ComBitaLiga)
                    ElseIf S_Mov = "B" Then
                         S_ComBitaLiga = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo) & "-" & Trim(IIf(Val(i) > 0, Val(i) * -1, Val(i)))
                        S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (1 Reg)"
                        Call PR_GrabaBitacora(6, 2, S_DescripBita, S_ComBitaLiga)
                    End If
                End If
            Next i
        End With
        
    Case "MODULO7"
        N_Reglones = 0
        N_Modulo = 7
        With Frm_Caratula
            For i = 0 To Val(.Chk_CarMod7.UBound)
                If Val(.Chk_CarMod7(i).Value) <> Val(.Chk_CarMod7(i).Tag) Then
                    If Val(.Chk_CarMod7(i).Value) > Val(.Chk_CarMod7(i).Tag) Then
                        S_Mov = "A"
                        .Chk_CarMod7(i).Tag = 1
                      
                        
                        S_CondiEje = "Insert into t_caratulas values (" & _
                        Trim(.Txt_CarProyecto) & "," & Trim(.Lbl_CveEtapa.Caption) & "," & N_Modulo & "," & _
                        IIf(Val(i) > 0, Val(i) * -1, Val(i)) & ",'" & Trim(.Chk_CarMod7(i).Caption) & _
                        "',Null,'C',0)"
                        
                        gcn.Execute S_CondiEje
                    ElseIf Val(.Chk_CarMod7(i).Value) < Val(.Chk_CarMod7(i).Tag) Then
                        S_Mov = "B"
                        .Chk_CarMod7(i).Tag = 0
                      
                        
                        S_Where = "Where n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And n_cveetapa = " & Trim(.Lbl_CveEtapa.Caption) & " And n_modulo = " & N_Modulo & " And n_opcion_base = " & IIf(Val(i) > 0, Val(i) * -1, Val(i))
                        S_CondiEje = "Delete From t_caratulas " & S_Where
                        
                        gcn.Execute S_CondiEje
                    End If
                        
                    If S_Mov = "A" Then
                        S_ComBitaLiga = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo) & "-" & Trim(IIf(Val(i) > 0, Val(i) * -1, Val(i)))
                        S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (1 Reg)"
                        Call PR_GrabaBitacora(6, 1, S_DescripBita, S_ComBitaLiga)
                    ElseIf S_Mov = "B" Then
                         S_ComBitaLiga = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo) & "-" & Trim(IIf(Val(i) > 0, Val(i) * -1, Val(i)))
                        S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (1 Reg)"
                        Call PR_GrabaBitacora(6, 2, S_DescripBita, S_ComBitaLiga)
                    End If
                End If
            Next i
        End With
        
    Case "MODULO8"
        N_Reglones = 0
        N_Modulo = 8
        With Frm_Caratula
          
            If .Chk_CarMod8(0).Value = 1 Then
                If .Opt_Mod8Apre(0).Value = True Then
                    S_OpcionCheckApre = "00"
                    S_EtiquetaCheckApreAmai = Trim(.Opt_Mod8Apre(0).Caption)
                Else
                    S_OpcionCheckApre = "01"
                    S_EtiquetaCheckApreAmai = Trim(.Opt_Mod8Apre(1).Caption)
                End If
                If Trim(.Fra_Mod8_1.Tag) <> Trim(S_OpcionCheckApre) Then
                    .Fra_Mod8_1.Tag = Trim(S_OpcionCheckApre)
                    S_MicondiAct = "UPDATE t_caratulas SET s_descrip_base = '" & Trim(.Chk_CarMod8(0).Caption) & Trim(S_EtiquetaCheckApreAmai) & "'" & _
                    " Where n_cveproyecto = " & Trim(.Lbl_CarCveProyecto.Caption) & " And n_cveetapa = " & Trim(.Lbl_CveEtapa.Caption) & " And N_Modulo = 8 And n_opcion_base = 0"
                    gcn.Execute S_MicondiAct
                End If
            End If
            If .Chk_CarMod8(1).Value = 1 Then
                If .Opt_Mod8Amai(0).Value = True Then
                    S_OpcionCheckApre = "10"
                    S_EtiquetaCheckApreAmai = Trim(.Opt_Mod8Amai(0).Caption)
                Else
                    S_OpcionCheckApre = "11"
                    S_EtiquetaCheckApreAmai = Trim(.Opt_Mod8Amai(1).Caption)
                End If
                If Trim(.Fra_Mod8_2.Tag) <> Trim(S_OpcionCheckApre) Then
                    .Fra_Mod8_2.Tag = Trim(S_OpcionCheckApre)
                    S_MicondiAct = "UPDATE t_caratulas SET s_descrip_base = '" & Trim(.Chk_CarMod8(1).Caption) & Trim(S_EtiquetaCheckApreAmai) & "'" & _
                    " Where n_cveproyecto = " & Trim(.Lbl_CarCveProyecto.Caption) & " And n_cveetapa = " & Trim(.Lbl_CveEtapa.Caption) & " And N_Modulo = 8 And n_opcion_base = -1"
                    gcn.Execute S_MicondiAct
                End If
            End If
          
            
            For i = 0 To Val(.Chk_CarMod8.UBound)
                S_CadOption = ""
                If Val(.Chk_CarMod8(i).Value) <> Val(.Chk_CarMod8(i).Tag) Then
                    If Val(.Chk_CarMod8(i).Value) > Val(.Chk_CarMod8(i).Tag) Then
                        S_Mov = "A"
                        .Chk_CarMod8(i).Tag = 1
                      
                      
                        If i = 0 Then
                            S_CadOption = .Opt_Mod8Apre(0).Caption
                            If .Opt_Mod8Apre(1).Value = True Then S_CadOption = .Opt_Mod8Apre(1).Caption
                        End If
                        If i = 1 Then
                            S_CadOption = .Opt_Mod8Amai(0).Caption
                            If .Opt_Mod8Amai(1).Value = True Then S_CadOption = .Opt_Mod8Amai(1).Caption
                        End If
                      
                        S_CondiEje = "Insert into t_caratulas values (" & _
                        Trim(.Txt_CarProyecto) & "," & Trim(.Lbl_CveEtapa.Caption) & "," & N_Modulo & "," & _
                        IIf(Val(i) > 0, Val(i) * -1, Val(i)) & ",'" & Trim(.Chk_CarMod8(i).Caption) & S_CadOption & _
                        "',Null,'C',0)"
                        
                        gcn.Execute S_CondiEje
                    ElseIf Val(.Chk_CarMod8(i).Value) < Val(.Chk_CarMod8(i).Tag) Then
                        S_Mov = "B"
                        .Chk_CarMod8(i).Tag = 0
                        If Trim(.Chk_CarMod8(i).Caption) = "Edad" Then Bo_Edad = True
                      
                        
                        S_Where = "Where n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And n_cveetapa = " & Trim(.Lbl_CveEtapa.Caption) & " And n_modulo = " & N_Modulo & " And n_opcion_base = " & IIf(Val(i) > 0, Val(i) * -1, Val(i))
                        S_CondiEje = "Delete From t_caratulas " & S_Where
                        
                        gcn.Execute S_CondiEje
                    End If
                        
                    If S_Mov = "A" Then
                        S_ComBitaLiga = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo) & "-" & Trim(IIf(Val(i) > 0, Val(i) * -1, Val(i)))
                        S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (1 Reg)"
                        Call PR_GrabaBitacora(6, 1, S_DescripBita, S_ComBitaLiga)
                    ElseIf S_Mov = "B" Then
                        S_ComBitaLiga = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo) & "-" & Trim(IIf(Val(i) > 0, Val(i) * -1, Val(i)))
                        S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (1 Reg)"
                        Call PR_GrabaBitacora(6, 2, S_DescripBita, S_ComBitaLiga)
                    End If
                End If
            Next i
            
            If .Chk_CarMod8(3).Value = 0 And Bo_Edad Then
                S_Micondicion = "n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And n_cveetapa = " & Trim(.Lbl_CveEtapa.Caption) & " And n_modulo = " & N_Modulo & " And c_tipo_opcion = 'L'"
                N_RegBorraM8 = FU_Cuenta_Registros("t_caratulas", S_Micondicion)
                
                S_Where = "Where n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And n_cveetapa = " & Trim(.Lbl_CveEtapa.Caption) & " And n_modulo = " & N_Modulo & " And c_tipo_opcion = 'L'"
                S_CondiEje = "Delete From t_caratulas " & S_Where
                
                gcn.Execute S_CondiEje
                
                If N_RegBorraM8 > 0 Then
                    S_ComBitaLiga = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo)
                    S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (" & Trim(N_RegBorraM8) & " Reg)"
                    Call PR_GrabaBitacora(6, 2, S_DescripBita, S_ComBitaLiga)
                End If
            End If
        End With
        
    Case "MODULO9"
        N_Reglones = 0
        N_Modulo = 9
        With Frm_Caratula
            For i = 0 To Val(.Chk_CarMod9.UBound)
                If Val(.Chk_CarMod9(i).Value) <> Val(.Chk_CarMod9(i).Tag) Then
                    If Val(.Chk_CarMod9(i).Value) > Val(.Chk_CarMod9(i).Tag) Then
                        N_TipoCarTel = 0
                        If Frm_Caratula.Opt_CaraTele(1).Value = True Then N_TipoCarTel = 1
                        S_Mov = "A"
                        .Chk_CarMod9(i).Tag = 1
                      
                        
                        S_CondiEje = "Insert into t_caratulas values (" & _
                        Trim(.Txt_CarProyecto) & "," & Trim(.Lbl_CveEtapa.Caption) & "," & N_Modulo & "," & _
                        IIf(Val(i) > 0, Val(i) * -1, Val(i)) & ",'" & Trim(.Chk_CarMod9(i).Caption) & _
                        "',Null,'C'," & N_TipoCarTel & ")"
                        
                        gcn.Execute S_CondiEje
                    ElseIf Val(.Chk_CarMod9(i).Value) < Val(.Chk_CarMod9(i).Tag) Then
                        S_Mov = "B"
                        .Chk_CarMod9(i).Tag = 0
                        If Trim(.Chk_CarMod9(i).Caption) = "No Elegible" Then Bo_NoElegible = True
                      
                        
                        S_Where = "Where n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And n_cveetapa = " & Trim(.Lbl_CveEtapa.Caption) & " And n_modulo = " & N_Modulo & " And n_opcion_base = " & IIf(Val(i) > 0, Val(i) * -1, Val(i))
                        S_CondiEje = "Delete From t_caratulas " & S_Where
                        
                        gcn.Execute S_CondiEje
                    End If
                        
                    If S_Mov = "A" Then
                        S_ComBitaLiga = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo) & "-" & Trim(IIf(Val(i) > 0, Val(i) * -1, Val(i)))
                        S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (1 Reg)"
                        Call PR_GrabaBitacora(6, 1, S_DescripBita, S_ComBitaLiga)
                    ElseIf S_Mov = "B" Then
                         S_ComBitaLiga = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo) & "-" & Trim(IIf(Val(i) > 0, Val(i) * -1, Val(i)))
                        S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (1 Reg)"
                        Call PR_GrabaBitacora(6, 2, S_DescripBita, S_ComBitaLiga)
                    End If
                End If
            Next i
            
            If .Chk_CarMod9(9).Value = 0 Then
                S_Micondicion = "n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And n_cveetapa = " & Trim(.Lbl_CveEtapa.Caption) & " And n_modulo = " & N_Modulo & " And c_tipo_opcion = 'L'"
                N_RegBorraM8 = FU_Cuenta_Registros("t_caratulas", S_Micondicion)
                
                S_Where = "Where n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And n_cveetapa = " & Trim(.Lbl_CveEtapa.Caption) & " And n_modulo = " & N_Modulo & " And c_tipo_opcion = 'L'"
                S_CondiEje = "Delete From t_caratulas " & S_Where
                
                gcn.Execute S_CondiEje
                
                If N_RegBorraM8 > 0 Then
                    S_ComBitaLiga = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo)
                    S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (" & Trim(N_RegBorraM8) & " Reg)"
                    Call PR_GrabaBitacora(6, 2, S_DescripBita, S_ComBitaLiga)
                End If
            End If
            
        End With
        
End Select

Exit Sub

GrabaChecksAct:
If Err.Number <> 0 Then
    MsgBox "Ha ocurrido un error al actualizar los datos", 0 + 32, "[PR_ActualizaInf_GridCaratula]"
    Exit Sub
End If
End Sub

Sub PR_ActualizaRegEtapas(N_Param)
Dim sQuery As String, Lposi As Long

On Error GoTo ErrAccionA

If gf_Confirma("¿Desea actualizar el registro actual?", "Confirmación") Then
    With Frm_Etapas
        sQuery = "Update c_etapa set s_descrip = '" & Trim(.Txt_EtaDescrip) & "',s_comentario = '" & Trim(.Txt_EtaDescripLar) & "', n_estatuscat = " & Val(.Lbl_EtaCveEstatus.Caption) & _
        " Where n_cveetapa = " & N_Param
    End With
    
    gcn.Execute sQuery
    MsgBox "Registro actualizado satisfactoriamente.", 0 + 64, "Etapas"
      
    Lposi = RE_Etapas.AbsolutePosition
  
    RE_Etapas.Requery
    RE_Etapas.AbsolutePosition = Lposi
End If
Exit Sub

ErrAccionA:
If Err.Number <> 0 Then
    MsgBox "Ha ocurrido un error al actualizar los datos", 0 + 32, "[PR_ActualizaRegEtapas]"
    Exit Sub
End If
End Sub

Sub PR_ActualizaRegistroIndividual_Pest(N_Modulo, S_DescripRegIndNue, N_ValorRegIndNue, S_DescripRegIndOri, N_ValorRegIndOri, S_CualCambio, N_QuintoValor)
Dim S_EjeBorrado As String, S_Where As String, S_ComBita As String, S_DescripBita As String, S_Cambios As String

On Error GoTo ErrAccionAInd
If N_Modulo = 5 Then
    With Frm_Caratula
        N_ModuloTruco = 4
        S_Where = " Where n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And n_cveetapa = " & Trim(.Lbl_CveEtapa.Caption) & " And N_Modulo = " & N_ModuloTruco & " And n_opcion_base = " & N_QuintoValor & " And n_opcion_det = " & N_ValorRegIndOri
        
        Select Case UCase(S_CualCambio)
            Case "D"
                S_EjeBorrado = "Update T_caratulasDet Set s_descrip_det = '" & S_DescripRegIndNue & "'"
                S_Cambios = S_DescripRegIndOri
            Case "V"
                S_EjeBorrado = "Update T_caratulasDet Set n_opcion_det = " & N_ValorRegIndNue
                S_Cambios = N_ValorRegIndOri
            Case "A"
                S_EjeBorrado = "Update T_caratulasDet Set n_opcion_det = " & N_ValorRegIndNue & " ,s_descrip_det = '" & S_DescripRegIndNue & "'"
                S_Cambios = S_DescripRegIndOri & " (" & N_ValorRegIndOri & ")"
        End Select
        
        S_EjeBorrado = S_EjeBorrado & S_Where
        gcn.Execute S_EjeBorrado
       
        S_ComBita = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo) & "-" & Trim(N_QuintoValor) & "-" & Trim(N_ValorRegIndNue)
        S_DescripBita = S_CualCambio & ": " & S_Cambios
        Call PR_GrabaBitacora(6, 3, S_DescripBita, S_ComBita)
    End With
Else
    With Frm_Caratula
        S_Where = " Where n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And n_cveetapa = " & Trim(.Lbl_CveEtapa.Caption) & " And N_Modulo = " & N_Modulo & " And n_opcion_base = " & N_ValorRegIndOri
        
        Select Case UCase(S_CualCambio)
            Case "D"
                S_EjeBorrado = "Update T_caratulas Set s_descrip_base = '" & S_DescripRegIndNue & "'"
                S_Cambios = S_DescripRegIndOri
            Case "V"
                S_EjeBorrado = "Update T_caratulas Set n_opcion_base = " & N_ValorRegIndNue
                S_Cambios = N_ValorRegIndOri
            Case "A"
                S_EjeBorrado = "Update T_caratulas Set n_opcion_base = " & N_ValorRegIndNue & " ,s_descrip_base = '" & S_DescripRegIndNue & "'"
                S_Cambios = S_DescripRegIndOri & " (" & N_ValorRegIndOri & ")"
        End Select
        
        S_EjeBorrado = S_EjeBorrado & S_Where
        gcn.Execute S_EjeBorrado
       
        S_ComBita = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo) & "-" & Trim(N_ValorRegIndNue)
        S_DescripBita = S_CualCambio & ": " & S_Cambios
        Call PR_GrabaBitacora(6, 3, S_DescripBita, S_ComBita)
    End With
End If
Exit Sub

ErrAccionAInd:
If Err.Number <> 0 Then
    MsgBox "Ha ocurrido un error al actualizar los datos", 0 + 32, "[PR_ActualizaRegistroIndividual_Pest]"
    Exit Sub
End If
End Sub

Sub PR_ActualizaRegPerfiles(N_Param)
Dim sQuery As String, Lposi As Long

On Error GoTo ErrAccionP

If gf_Confirma("¿Desea actualizar el registro actual?", "Confirmación") Then
    With Frm_Perfiles
        sQuery = "Update c_perfil set s_descrip = '" & Trim(.Txt_PerDescrip) & "',s_comentario = '" & Trim(.Txt_PerDescripLar) & "',n_estatuscat = " & Val(.Lbl_PerCveEstatus.Caption) & _
        " Where n_cveperfil = " & N_Param
    End With
    
    gcn.Execute sQuery
    MsgBox "Registro actualizado satisfactoriamente.", 0 + 64, "Perfiles"
      
    Lposi = RE_Perfiles.AbsolutePosition
  
    RE_Perfiles.Requery
    RE_Perfiles.AbsolutePosition = Lposi
End If
Exit Sub

ErrAccionP:
If Err.Number <> 0 Then
    MsgBox "Ha ocurrido un error al actualizar los datos", 0 + 32, "[PR_ActualizaRegPerfiles]"
    Exit Sub
End If

End Sub

Sub PR_ActualizaRegProyectos(N_Param)
Dim sQuery As String, Lposi As Long, F_Sistema

On Error GoTo ErrAccionA
'n_cveproyecto , s_nombre, n_cveestatus, n_cveejec, n_cveprog, n_productividad, s_comentario, n_interno
If gf_Confirma("¿Desea actualizar el registro actual?", "Confirmación") Then
    F_Sistema = "'" & Format(FU_ExtraeFechaServer(), "dd/mm/yyyy hh:mm:ss") & "'"
    With Frm_Proyecto
        sQuery = "Update t_proyectos set " & _
        "s_nombre = '" & Trim(.Txt_ProyNombre) & "', n_cveestatus = " & Val(Trim(.Lbl_ProyCveEstatus.Caption)) & "," & _
        "n_cveejec_sc = " & Val(Trim(.Txt_ProyNumEjecutivoSC)) & "," & _
        "n_cveejec_pr = " & Val(Trim(.Txt_ProyNumEjecutivoPR)) & "," & _
        "n_cveprog = " & Val(Trim(.Txt_ProyNumProg)) & ", n_productividad = " & Val(Trim(.Txt_ProyProd)) & "," & _
        "s_comentario = '" & Trim(.Txt_ProyComen) & "' Where n_cveproyecto = " & N_Param
    End With
    
    gcn.Execute sQuery
    MsgBox "Registro actualizado satisfactoriamente.", 0 + 64, "Proyectos"
      
    Call PR_GrabaBitacora(4, 3, "t_proyectos", Trim(N_Param))
    Lposi = RE_Proyectos.AbsolutePosition
  
    RE_Proyectos.Requery
    RE_Proyectos.AbsolutePosition = Lposi
End If
Exit Sub

ErrAccionA:
If Err.Number <> 0 Then
    MsgBox "Ha ocurrido un error al actualizar los datos", 0 + 32, "[PR_ActualizaRegProyectos]"
    Exit Sub
End If
End Sub

Sub PR_ActualizaRegUsuarios(N_Param)
Dim sQuery As String, Lposi As Long, F_Sistema, S_Micondicion As String

On Error GoTo ErrAccionA

If gf_Confirma("¿Desea actualizar el registro actual?", "Confirmación") Then
    F_Sistema = "'" & Format(FU_ExtraeFechaServer(), "dd/mm/yyyy hh:mm:ss") & "'"
    With Frm_Usuarios
        sQuery = "Update c_segusuarios set " & _
        "s_nombre = '" & Trim(.Txt_UsuNombre) & "', s_paterno = '" & Trim(.Txt_UsuPaterno) & "'," & _
        "s_materno = '" & Trim(.Txt_UsuMaterno) & "'," & _
        "s_pwd = '" & Trim(.Txt_UsuContra) & "', n_cveperfil = " & Val(.Lbl_UsuCvePerfil.Caption) & "," & _
        "n_estatuscat = " & Val(.Lbl_UsuCveEstatus.Caption) & "," & _
        "f_movi = " & F_Sistema & "," & _
        "n_cvepuesto= " & IIf(Val(.txtCampo(0).Tag) = 0, "null", Val(.txtCampo(0).Tag)) & "," & _
        "n_cveempleado= " & IIf(Val(.txtCampo(1).Tag) = 0, "null", Val(.txtCampo(1).Tag)) & "," & _
        "n_cvepersona= " & IIf(Val(.txtCampo(2).Tag) = 0, "null", Val(.txtCampo(2).Tag)) & " Where n_cveusuario = " & N_Param
    End With
    
    gcn.Execute sQuery
    MsgBox "Registro actualizado satisfactoriamente.", 0 + 64, "Usuarios"
      
    Call PR_GrabaBitacora(1, 3, "c_segusuarios", Trim(N_Param))
    Lposi = RE_Usuarios.AbsolutePosition
  
    RE_Usuarios.Requery
    RE_Usuarios.AbsolutePosition = Lposi
Else
  
    S_Micondicion = "n_cveusuario = " & Trim(Frm_Usuarios.Txt_UsuClave)
    If FU_Cuenta_Registros("C_Segusuarios", S_Micondicion) > 0 Then
        Frm_Usuarios.Txt_UsuContra.Tag = FU_RescataInfX_CampoS("s_pwd", "c_segusuarios", S_Micondicion)
        Frm_Usuarios.Txt_UsuContra = Trim(Frm_Usuarios.Txt_UsuContra.Tag)
    End If
End If
Exit Sub

ErrAccionA:
If Err.Number <> 0 Then
    MsgBox "Ha ocurrido un error al actualizar los datos", 0 + 32, "[PR_ActualizaRegUsuarios]"
    Exit Sub
End If
End Sub

Function FU_cbo_InicializaUsuarios(Ctl_Control As Control, S_Condi As String)
Dim RE_MiTabla As ADODB.Recordset

Set RE_MiTabla = New ADODB.Recordset
RE_MiTabla.Open S_Condi, gcn

Ctl_Control.Clear
Do While Not RE_MiTabla.EOF()
    Ctl_Control.AddItem Trim(RE_MiTabla(1)) & " " & Trim(RE_MiTabla(2)) & " " & Trim(RE_MiTabla(3))
    Ctl_Control.ItemData(Ctl_Control.NewIndex) = Trim(RE_MiTabla(0))
    RE_MiTabla.MoveNext
Loop
RE_MiTabla.Close

End Function
Sub PR_BorraInfPerfiles()
Dim sQuery As String, Lposis As Long

On Error GoTo ErrBorrado

Lposis = RE_Perfiles.AbsolutePosition
If gf_Confirma("¿Desea borrar el registro actual?", "Confirmación") Then
    With Frm_Perfiles
        sQuery = "DELETE FROM c_perfil Where n_cveperfil = " & Trim(.Txt_PerCvePer)
        gcn.Execute sQuery
        
        RE_Perfiles.Requery
        .txt_TotRegPerfiles = " "
            
        If (RE_Perfiles.EOF = True And RE_Perfiles.BOF = True) Then
           MsgBox "No hay más registros", vbInformation, "Sistema"
           Call PR_LimpiaDatos("Perfiles")
           Exit Sub
        Else
            RE_Perfiles.AbsolutePosition = Lposis
            If RE_Perfiles.EOF = True Then
                RE_Perfiles.MovePrevious
            End If
            If RE_Perfiles.BOF = True Then
                MsgBox "No hay más registros", vbInformation, "Sistema"
                Call PR_LimpiaDatos("Perfiles")
                Exit Sub
            End If
            Call PR_AcomodaDatosPerfiles
            .txt_TotRegPerfiles = RE_Perfiles.AbsolutePosition & "/" & RE_Perfiles.RecordCount
            
        End If
    End With
End If
Exit Sub

ErrBorrado:
If Err.Number <> 0 Then
    MsgBox "Existe un error al borrar el registro actual. " & " " & Err.Description, 0 + 16, Err.Number
    Exit Sub
End If
End Sub
Sub PR_BorraInfEtapas()
Dim sQuery As String, Lposis As Long

On Error GoTo ErrBorrado

Lposis = RE_Etapas.AbsolutePosition
If gf_Confirma("¿Desea borrar el registro actual?", "Confirmación") Then
    With Frm_Etapas
        sQuery = "DELETE FROM c_etapa  Where n_cveetapa = " & Trim(.Txt_EtaCveEta)
        gcn.Execute sQuery
        
        RE_Etapas.Requery
        .txt_TotRegEtapas = " "
            
        If (RE_Etapas.EOF = True And RE_Etapas.BOF = True) Then
           MsgBox "No hay más registros", vbInformation, "Sistema"
           Call PR_LimpiaDatos("Etapas")
           Exit Sub
        Else
            RE_Etapas.AbsolutePosition = Lposis
            If RE_Etapas.EOF = True Then
                RE_Etapas.MovePrevious
            End If
            If RE_Etapas.BOF = True Then
                MsgBox "No hay más registros", vbInformation, "Sistema"
                Call PR_LimpiaDatos("Etapas")
                Exit Sub
            End If
            Call PR_AcomodaDatosEtapas
            .txt_TotRegEtapas = RE_Etapas.AbsolutePosition & "/" & RE_Etapas.RecordCount
            
        End If
    End With
End If
Exit Sub

ErrBorrado:
If Err.Number <> 0 Then
    MsgBox "Existe un error al borrar el registro actual. " & " " & Err.Description, 0 + 16, Err.Number
    Exit Sub
End If
End Sub

Sub PR_AnteriorReg_Eta()
If RE_Etapas.BOF = True Then
   MsgBox "No hay más registros", vbInformation, "Etapas"
Else
   RE_Etapas.MovePrevious
   If RE_Etapas.BOF = True Then
        MsgBox "No hay más registros", vbInformation, "Etapas"
        RE_Etapas.MoveNext
        Exit Sub
   End If
   
   Call PR_AcomodaDatosEtapas
   Exit Sub
End If
End Sub
Sub PR_BorraInfProyectos()
Dim sQuery As String, Lposis As Long

On Error GoTo ErrBorradoP

Lposis = RE_Proyectos.AbsolutePosition
If gf_Confirma("¿Desea borrar el registro actual?", "Confirmación") Then
    With Frm_Proyecto
        sQuery = "DELETE FROM t_proyectos Where n_cveproyecto = " & Trim(Frm_Proyecto.Txt_ProyNumero)
        gcn.Execute sQuery
                                
        RE_Proyectos.Requery
        .txt_TotRegProy = " "
        
        Call PR_GrabaBitacora(4, 2, "t_proyectos", Trim(Frm_Proyecto.Txt_ProyNumero))
            
        If (RE_Proyectos.EOF = True And RE_Proyectos.BOF = True) Then
           MsgBox "No hay más registros", vbInformation, "Sistema"
           Call PR_LimpiaDatos("Proyectos")
           Exit Sub
        Else
            RE_Proyectos.AbsolutePosition = Lposis
            If RE_Proyectos.EOF = True Then
                RE_Proyectos.MovePrevious
            End If
            If RE_Proyectos.BOF = True Then
                MsgBox "No hay más registros", vbInformation, "Sistema"
                Call PR_LimpiaDatos("Proyectos")
                Exit Sub
            End If
            Call PR_AcomodaDatosProyectos
            .txt_TotRegProy = RE_Proyectos.AbsolutePosition & "/" & RE_Proyectos.RecordCount
        End If
    End With
End If

Exit Sub

ErrBorradoP:
If Err.Number <> 0 Then
    MsgBox "Existe un error al borrar el registro actual. " & " " & Err.Description, 0 + 16, Err.Number
    Exit Sub
End If
End Sub

Sub PR_BorraInfUsuarios()
Dim sQuery As String, Lposis As Long

On Error GoTo ErrBorrado

Lposis = RE_Usuarios.AbsolutePosition
If gf_Confirma("¿Desea borrar el registro actual?", "Confirmación") Then
    With Frm_Usuarios
        sQuery = "DELETE FROM c_segusuarios Where n_cveusuario = " & Trim(Frm_Usuarios.Txt_UsuClave)
        gcn.Execute sQuery
                                
        RE_Usuarios.Requery
        .txt_TotRegUsuarios = " "
        
        Call PR_GrabaBitacora(1, 2, "c_segusuarios", Trim(Frm_Usuarios.Txt_UsuClave))
            
        If (RE_Usuarios.EOF = True And RE_Usuarios.BOF = True) Then
           MsgBox "No hay más registros", vbInformation, "Sistema"
           Call PR_LimpiaDatos("Usuarios")
           Exit Sub
        Else
            RE_Usuarios.AbsolutePosition = Lposis
            If RE_Usuarios.EOF = True Then
                RE_Usuarios.MovePrevious
            End If
            If RE_Usuarios.BOF = True Then
                MsgBox "No hay más registros", vbInformation, "Sistema"
                Call PR_LimpiaDatos("Usuarios")
                Exit Sub
            End If
            Call PR_AcomodaDatosUsuarios
            .txt_TotRegUsuarios = RE_Usuarios.AbsolutePosition & "/" & RE_Usuarios.RecordCount
            
        End If
    End With
End If
Exit Sub

ErrBorrado:
If Err.Number <> 0 Then
    MsgBox "Existe un error al borrar el registro actual. " & " " & Err.Description, 0 + 16, Err.Number
    Exit Sub
End If
Exit Sub
End Sub

Sub PR_BorraRegistroIndividual_Pest(N_Modulo, N_ValorRegInd, N_QuintoValor)
Dim S_EjeBorrado As String, S_EjeWhere As String, S_ComBita As String, S_DescripBita As String

On Error GoTo ErrAccionBInd

With Frm_Caratula
    If Val(N_Modulo) = 5 Then
        N_ModuloTruco = 4
        S_EjeBorrado = "Delete From T_caratulasDet "
        S_EjeWhere = "Where n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And n_cveetapa = " & Trim(.Lbl_CveEtapa.Caption) & " And N_Modulo = " & N_ModuloTruco & " And n_opcion_base = " & Val(N_QuintoValor) & " And n_opcion_det = " & Val(N_ValorRegInd)
    Else
        S_EjeBorrado = "Delete From T_caratulas "
        S_EjeWhere = "Where n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And n_cveetapa = " & Trim(.Lbl_CveEtapa.Caption) & " And N_Modulo = " & N_Modulo & " And n_opcion_base = " & N_ValorRegInd
    End If
    S_EjeBorrado = S_EjeBorrado & S_EjeWhere
    
    gcn.Execute S_EjeBorrado
    If Val(N_Modulo) = 5 Then
        S_ComBita = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo) & "-" & Trim(N_QuintoValor) & "-" & Trim(N_ValorRegInd)
    Else
        S_ComBita = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo) & "-" & Trim(N_ValorRegInd)
    End If
    S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (1 Reg)"
    Call PR_GrabaBitacora(6, 2, S_DescripBita, S_ComBita)
End With

Exit Sub

ErrAccionBInd:
If Err.Number <> 0 Then
    MsgBox "Ha ocurrido un error al borrar los datos", 0 + 32, "[PR_BorraRegistroIndividual_Pest]"
    Exit Sub
End If
End Sub

Sub PR_CargaInf_Auditoria()
Dim RE_Bitacora As ADODB.Recordset, S_CondiCarga As String, N_Campo As Integer
Dim S_MicondicionBus As String

With Frm_Auditoria
     S_CondiCarga = "Select b.n_aud1,b.n_aud2,b.n_aud3,b.n_aud4,b.n_aud5,b.n_aud6,b.n_aud7,b.n_aud8,b.n_aud9,b.n_aud10," & _
    "b.n_aud11,b.n_aud12,b.n_aud13,b.n_aud14,b.n_aud15,b.n_aud16,b.n_aud17,b.n_aud18,b.n_aud19, " & _
    "b.n_cveestatusaud,b.n_cvetipoaud,b.s_auditorclave, b.s_comentario, b.n_ind_menor, b.n_ind_mayor, b.n_ind_critica, b.n_ind_calidad, e.S_Descrip " & _
    "From t_bitacoraAud b (Nolock),c_estatusauditoria e (Nolock) " & _
    "Where b.n_cveestatusaud = e.n_cveestatusaud " & _
    "And n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And s_numcuestionario = " & Trim(gs_Cuestionario)
                    
     Set RE_Bitacora = New ADODB.Recordset
     RE_Bitacora.Open S_CondiCarga, gcn
    
    If Val(RE_Bitacora.RecordCount) = 0 Then Exit Sub
    RE_Bitacora.MoveFirst
    
    For N_Campo = 0 To Val(.Txt_AudMayor.UBound)
        .Txt_AudMayor(N_Campo) = RE_Bitacora(N_Campo)
    Next N_Campo
    
    .Cmb_AudEstatus = RE_Bitacora(27)
    .Lbl_AudCveEstatus.Caption = RE_Bitacora(19)
    .Lbl_AudCveTipo.Caption = RE_Bitacora(20)
    .Txt_AudAuditor = Trim(RE_Bitacora(21))
    .Txt_AudComentario = RE_Bitacora(22)
    .Txt_AudIndices(0) = RE_Bitacora(23)
    .Txt_AudIndices(1) = RE_Bitacora(24)
    .Txt_AudIndices(2) = RE_Bitacora(25)
    .Txt_AudIndices(3) = RE_Bitacora(26)
    
    If Val(RE_Bitacora(20)) > 0 Then
        S_MicondicionBus = "n_cvetipoaud = " & Trim(RE_Bitacora(20))
        .Cmb_AudTipo = FU_RescataInfX_CampoS("s_descrip", "c_tipoauditoria", S_MicondicionBus)
    Else
        .Cmb_AudTipo = ""
    End If
     .Cmb_AudEstatus.ListIndex = FU_RetornaIndiceCombo(.Cmb_AudEstatus)
     .Cmb_AudTipo.ListIndex = FU_RetornaIndiceCombo(.Cmb_AudTipo)
 End With
 
 RE_Bitacora.Close
 Set RE_Bitacora = Nothing

End Sub

Sub PR_CargaInf_BitacoraCalidad()
Dim RE_Bitacora As ADODB.Recordset, S_CondiCarga As String, N_Campo As Integer

With Frm_Auditoria
     S_CondiCarga = "Select bi.n_aud1,bi.n_aud2,bi.n_aud3,bi.n_aud4,bi.n_aud5,bi.n_aud6,bi.n_aud7,bi.n_aud8,bi.n_aud9,bi.n_aud10," & _
    "bi.n_aud11,bi.n_aud12, bi.n_aud13, bi.n_aud14, bi.n_aud15, bi.n_aud16, bi.n_aud17, bi.n_aud18, bi.n_aud19 " & _
    "From t_bitacoraCal bi (Nolock) " & _
     "Where n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And s_numcuestionario = " & Trim(gs_Cuestionario)
                    
     Set RE_Bitacora = New ADODB.Recordset
     RE_Bitacora.Open S_CondiCarga, gcn
    
    If Val(RE_Bitacora.RecordCount) = 0 Then Exit Sub
    RE_Bitacora.MoveFirst
    
    For N_Campo = 0 To Val(.Txt_AudMayor.UBound)
        .Txt_AudMayor(N_Campo) = RE_Bitacora(N_Campo)
    Next N_Campo
 End With
 RE_Bitacora.Close
 Set RE_Bitacora = Nothing
 
End Sub
Sub PR_CargaInformacionCuestionario(S_ModuloSel, N_Cuestionario)
Dim N_Ctl As Integer, N_Modulo As Integer
Dim RE_Cuestion As ADODB.Recordset, S_CondiCarga As String, S_TablaTG As String

If gs_ProcCuestionario = 1 Then
    S_TablaTG = "t_cuestionarioPrue"
Else
    S_TablaTG = "t_cuestionario"
End If

Select Case S_ModuloSel
    Case "123"
        S_CondiCarga = "Select t.n_idctrl,t.s_valorasig,n_modulo from " & S_TablaTG & " t (Nolock) " & _
        "Where t.n_cveproyecto = " & gs_Proyecto & " And t.n_cveetapa = " & gs_Etapa & " And t.s_numcuestionario = " & N_Cuestionario & " And t.N_Modulo < 4 " & _
        "Order by t.n_modulo,t.n_idctrl"
        
        Set RE_Cuestion = New ADODB.Recordset
        RE_Cuestion.Open S_CondiCarga, gcn
        
        If Val(RE_Cuestion.RecordCount) = 0 Then Exit Sub
        
        N_Ctl = 0
        Do While Not RE_Cuestion.EOF()
            Frm_CuestionarioDin.Txt_CueDinMod1(N_Ctl) = RE_Cuestion(1)
            N_Ctl = N_Ctl + 1
            RE_Cuestion.MoveNext
        Loop
        RE_Cuestion.Close
        Set RE_Cuestion = Nothing
        
   Case "4"
        S_CondiCarga = "Select t.n_idctrl,t.s_valorasig,n_modulo from " & S_TablaTG & " t (Nolock) " & _
        "Where t.n_cveproyecto = " & gs_Proyecto & " And t.n_cveetapa = " & gs_Etapa & " And t.s_numcuestionario = " & N_Cuestionario & " And t.N_Modulo = 4 " & _
        "Order by t.n_modulo,t.n_idctrl"
        
        Set RE_Cuestion = New ADODB.Recordset
        RE_Cuestion.Open S_CondiCarga, gcn
        
        If Val(RE_Cuestion.RecordCount) = 0 Then Exit Sub
        
        N_Ctl = 0
        Do While Not RE_Cuestion.EOF()
            Frm_CuestionarioDin.Txt_CueDinMod45(N_Ctl) = RE_Cuestion(1)
            N_Ctl = N_Ctl + 1
            RE_Cuestion.MoveNext
        Loop
        RE_Cuestion.Close
        Set RE_Cuestion = Nothing
        
    Case "6"
        S_CondiCarga = "Select t.n_idctrl,t.s_valorasig,n_modulo from " & S_TablaTG & " t (Nolock) " & _
        "Where t.n_cveproyecto = " & gs_Proyecto & " And t.n_cveetapa = " & gs_Etapa & " And t.s_numcuestionario = " & N_Cuestionario & " And t.N_Modulo = 6 " & _
        "Order by t.n_modulo,t.n_idctrl"
        
        Set RE_Cuestion = New ADODB.Recordset
        RE_Cuestion.Open S_CondiCarga, gcn
        
        If Val(RE_Cuestion.RecordCount) = 0 Then Exit Sub
        
        N_Ctl = 0
        Do While Not RE_Cuestion.EOF()
            Frm_CuestionarioDin.Txt_CueDinMod6(N_Ctl) = RE_Cuestion(1)
            N_Ctl = N_Ctl + 1
            RE_Cuestion.MoveNext
        Loop
        RE_Cuestion.Close
        Set RE_Cuestion = Nothing
        
    Case "7"
        S_CondiCarga = "Select t.n_idctrl,t.s_valorasig,n_modulo from " & S_TablaTG & " t (Nolock) " & _
        "Where t.n_cveproyecto = " & gs_Proyecto & " And t.n_cveetapa = " & gs_Etapa & " And t.s_numcuestionario = " & N_Cuestionario & " And t.N_Modulo = 7 " & _
        "Order by t.n_modulo,t.n_idctrl"
        
        Set RE_Cuestion = New ADODB.Recordset
        RE_Cuestion.Open S_CondiCarga, gcn
        
        If Val(RE_Cuestion.RecordCount) = 0 Then Exit Sub
        
        N_Ctl = 0
        Do While Not RE_Cuestion.EOF()
          
          
          
          
          
          
            
            Frm_CuestionarioDin.Txt_CueDinMod7(N_Ctl) = RE_Cuestion(1)
            N_Ctl = N_Ctl + 1
            RE_Cuestion.MoveNext
        Loop
        RE_Cuestion.Close
        Set RE_Cuestion = Nothing
        
    Case "8"
        S_CondiCarga = "Select t.n_idctrl,t.s_valorasig,n_modulo from " & S_TablaTG & " t (Nolock) " & _
        "Where t.n_cveproyecto = " & gs_Proyecto & " And t.n_cveetapa = " & gs_Etapa & " And t.s_numcuestionario = " & N_Cuestionario & " And t.N_Modulo = 8 " & _
        "Order by t.n_modulo,t.n_idctrl"
        
        Set RE_Cuestion = New ADODB.Recordset
        RE_Cuestion.Open S_CondiCarga, gcn
        
        If Val(RE_Cuestion.RecordCount) = 0 Then Exit Sub
        
        N_Ctl = 0
        Do While Not RE_Cuestion.EOF()
            Frm_CuestionarioDin.Txt_CueDinMod8(N_Ctl) = RE_Cuestion(1)
            N_Ctl = N_Ctl + 1
            RE_Cuestion.MoveNext
        Loop
        RE_Cuestion.Close
        Set RE_Cuestion = Nothing
    
End Select

Exit Sub
 
End Sub

Function FU_ColocaLetreros_ResumenContactos() As Integer
Dim S_MiCuenta As String, N_TotChecks9 As Integer, N_Suma As Integer, N_Valor As Variant
Dim N_ColFiltro As Integer, N_Col_NE As Integer

FU_ColocaLetreros_ResumenContactos = 0
With Frm_Resumen
    S_MiCuenta = "n_cveproyecto = " & gs_Proyecto & " and n_cveetapa = " & gs_Etapa & " And n_modulo = 9 And c_tipo_opcion = 'C'"
    N_TotChecks9 = FU_Cuenta_Registros("t_Caratulas", S_MiCuenta)
    If N_TotChecks9 = 0 Then Exit Function
    N_Reg = Val(.MSFlexGrid1.Rows - 1)
    
  
    N_Col_NE = 0
    For i = 1 To Val(.MSFlexGrid1.Cols - 1)
        If UCase(Trim(.MSFlexGrid1.TextMatrix(0, i))) = "NO ELEGIBLE" Then
            N_Col_NE = i
            Exit For
        End If
    Next i
   
    If N_Col_NE > 0 Then
        For li_Row = 2 To N_Reg
            N_Suma = 0
            For li_Col = (N_Col_NE + 1) To Val(.MSFlexGrid1.Cols - 1)
                N_Valor = Trim(.MSFlexGrid1.TextMatrix(li_Row, li_Col))
                N_Valor = IIf(IsNumeric(N_Valor), Val(N_Valor), 0)
                N_Suma = N_Suma + N_Valor
            Next
          
            If Val(.MSFlexGrid1.TextMatrix(2, N_Col_NE)) <> N_Suma Then
                MsgBox "La suma de las preguntas filtro no coincide con el valor de la celda No Elegible.", 0 + 16, "Verificar"
                .MSFlexGrid1.SetFocus
                .MSFlexGrid1.Col = N_Col_NE
                FU_ColocaLetreros_ResumenContactos = -1
                Exit For
            End If
        Next
    End If
    
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
End With
End Function

Sub PR_ConsultaInfPestanas()
Call PR_AcomodaInformacionCaratula("Modulo1")
Call PR_AcomodaInformacionCaratula("Modulo2")
Call PR_AcomodaInformacionCaratula("Modulo3")
Call PR_AcomodaInformacionCaratula("Modulo4")
Call PR_AcomodaInformacionCaratula("Modulo5")
Call PR_AcomodaInformacionCaratula("Modulo5-Def")
Call PR_AcomodaInformacionCaratula("Modulo6")
Call PR_AcomodaInformacionCaratula("Modulo7")
Call PR_AcomodaInformacionCaratula("Modulo8")
Call PR_AcomodaInformacionCaratula("Modulo9")

End Sub
Sub PR_GrabaBitacora(N_Pan As Integer, N_Mov As Integer, S_Descrip As String, S_Liga As String)
Dim sQueryBita As String
Dim F_Sistema

On Error GoTo ErrAccionBit
''gs_usuario = 10001

F_Sistema = "'" & Format(FU_ExtraeFechaServer(), "dd/mm/yyyy hh:mm:ss") & "'"
With Frm_Usuarios
    sQueryBita = "Insert Into t_bitacora Values (" & _
    Val(gs_usuario) & "," & Val(N_Pan) & "," & Val(N_Mov) & "," & _
    Trim(F_Sistema) & ",'" & Trim(S_Descrip) & "','" & Trim(S_Liga) & "')"
End With

gcn.Execute sQueryBita
Exit Sub

ErrAccionBit:
If Err.Number <> 0 Then
    MsgBox "Ha ocurrido un error al guardar los datos", 0 + 32, "[PR_GrabaBitacora]"
    Exit Sub
End If
'Graba en la bitacora
End Sub
Sub PR_GuardaBitacoraCalidad()
Dim S_CadenaBase As String, S_CadenaInsert  As String, S_CadenaEje As String, S_ValorX As Variant
Dim S_Borrar As String, S_ComBita As String, S_DescripBita As String, li_Row As Integer
Dim S_InsertaAudi As String, N_ValorTexto As Variant, N_ValAsignado As Long

On Error GoTo GrabaAuditoria
With Frm_Auditoria
      
    S_Borrar = "Delete From t_bitacoraCal " & _
    "Where n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And s_numcuestionario = " & gs_Cuestionario
     gcn.Execute S_Borrar
    
    S_CadenaBase = "Insert into t_bitacoraCal Values (" & gs_Proyecto & "," & gs_Etapa & "," & gs_Cuestionario & ","

    S_CadenaInsert = ""
    
    For li_Row = 0 To Val(.Txt_AudMayor.UBound)
        If Trim(.Txt_AudMayor(li_Row)) = "" Then
            S_ValorX = Null
        Else
            S_ValorX = Val(Trim(.Txt_AudMayor(li_Row)))
        End If
        S_CadenaInsert = S_CadenaInsert & S_ValorX & ","
    Next
    S_CadenaInsert = Left(S_CadenaInsert, Len(Trim(S_CadenaInsert)) - 1) & ")"
    S_CadenaEje = S_CadenaBase & S_CadenaInsert
  
    gcn.Execute S_CadenaEje
    
  
    For i = 8 To 18
        N_ValorTexto = IIf(IsNumeric(Trim(.Txt_AudMayor(i))), Val(.Txt_AudMayor(i)), 0)
        N_ValAsignado = N_ValAsignado + N_ValorTexto
    Next i
    If N_ValAsignado > 0 Then
      
        S_CondiAct = "Update t_cuestionario Set s_valorasig = " & Val(.Lbl_AudCveEstatus.Caption) & ", s_valorformato = '" & UCase(Trim(.Cmb_AudEstatus)) & _
        "' Where n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And s_numcuestionario = " & gs_Cuestionario & " And n_modulo = 7 And Rtrim(Ltrim(s_descrip_base)) = 'Entrevista'"
        gcn.Execute S_CondiAct
    End If
  
    
    S_ComBita = Trim(gs_Proyecto) & "-" & Trim(gs_Etapa) & "-" & Trim(gs_Cuestionario)
    If gs_ProcCuestionario = 13 Then
        MsgBox "Registros de la bitácora de calidad, actualizados satisfactoriamente.", 0 + 64, "Bitacora de Calidad"
        S_DescripBita = "t_bitacoraCal" & " - Actualización"
        Call PR_GrabaBitacora(9, 3, S_DescripBita, S_ComBita)
      
        S_ActualizaEnc = "Update t_enccuestionario set n_exportar = 2 " & _
        "Where n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And s_numcuestionario = '" & Trim(gs_Cuestionario) & "'"
        gcn.Execute S_ActualizaEnc
      
    Else
        MsgBox "Registros de la bitácora de calidad, guardados satisfactoriamente.", 0 + 64, "Bitacora de Calidad"
        S_DescripBita = "t_bitacoraCal"
        Call PR_GrabaBitacora(9, 1, S_DescripBita, S_ComBita)
    End If
End With

Exit Sub

GrabaAuditoria:
If Err.Number <> 0 Then
    MsgBox Str(Err.Number) & ": " & Err.Description, 0 + 64, "PR_GuardaBitacoraAuditoria"
    Exit Sub
End If
'-----------------------------------------------------------------------------------------------------------------------
'n_cveproyecto,n_cveetapa,s_numcuestionario,n_aud1,n_aud2,n_aud3,n_aud4,n_aud5,n_aud6,n_aud7,n_aud8,n_aud9,n_aud10,n_aud11,n_aud12,n_aud13,n_aud14,n_aud15,n_aud16,n_aud17,n_aud18,n_aud19
'-----------------------------------------------------------------------------------------------------------------------
End Sub

Sub PR_GuardaInf_GridCaratula()
Dim S_CabeceraAviso As String, S_CadenotaAviso As String, S_CadenaPaso As String

S_CabeceraAviso = "Se guardo la siguente información:" & vbCrLf
S_CadenotaAviso = ""

Frm_Caratula.MousePointer = 11
S_CadenaPaso = FU_GuardaInf_GridCaratulaBD("Modulo1")
S_CadenotaAviso = S_CadenotaAviso & S_CadenaPaso

S_CadenaPaso = FU_GuardaInf_GridCaratulaBD("Modulo2")
S_CadenotaAviso = S_CadenotaAviso & S_CadenaPaso

S_CadenaPaso = FU_GuardaInf_GridCaratulaBD("Modulo3")
S_CadenotaAviso = S_CadenotaAviso & S_CadenaPaso

S_CadenaPaso = FU_GuardaInf_GridCaratulaBD("Modulo4")
S_CadenotaAviso = S_CadenotaAviso & S_CadenaPaso

'S_CadenaPaso = FU_GuardaInf_GridCaratulaBD("Modulo5-Def")
'S_CadenotaAviso = S_CadenotaAviso & S_CadenaPaso Truco para que no mande mensaje

S_CadenaPaso = FU_GuardaInf_GridCaratulaBD("Modulo5")
S_CadenotaAviso = S_CadenotaAviso & S_CadenaPaso

S_CadenaPaso = FU_GuardaInf_GridCaratulaBD("Modulo6")
S_CadenotaAviso = S_CadenotaAviso & S_CadenaPaso

S_CadenaPaso = FU_GuardaInf_GridCaratulaBD("Modulo7")
S_CadenotaAviso = S_CadenotaAviso & S_CadenaPaso

S_CadenaPaso = FU_GuardaInf_GridCaratulaBD("Modulo8")
S_CadenotaAviso = S_CadenotaAviso & S_CadenaPaso

S_CadenaPaso = FU_GuardaInf_GridCaratulaBD("Modulo9")
S_CadenotaAviso = S_CadenotaAviso & S_CadenaPaso
If Len(Trim(S_CadenotaAviso)) > 0 Then
  
    MsgBox "Se guardo la información satisfactoriamente.", 0 + 64, "Información Guardada"
  
    PR_LimpiaControlesPestanas
    PR_LimpiaDatos ("Caratula-Enc")
    PR_LimpiaDatos ("Caratula-Pestanas")
    With Frm_Caratula
        .Bot_CarHabilitar.Enabled = False
        .Txt_CarProyecto.SetFocus
        .Bot_Caratula(2).Enabled = False
    End With
Else
    MsgBox "Este proyecto aún no tiene definida la caratula.", 0 + 48, "Sin Información"
End If
Frm_Caratula.MousePointer = 0


End Sub
Sub PR_GuardaRegUsarios()
Dim sQuery As String
Dim Cancela As Boolean, F_Sistema

On Error GoTo ErrAccionG

F_Sistema = "'" & Format(FU_ExtraeFechaServer(), "dd/mm/yyyy hh:mm:ss") & "'"
With Frm_Usuarios
    sQuery = "Insert Into c_segusuarios Values (" & _
    Val(Trim(.Txt_UsuClave)) & ",'" & Trim(.Txt_UsuNombre) & "','" & Trim(.Txt_UsuPaterno) & "','" & _
    Trim(.Txt_UsuMaterno) & "','" & Trim(.Txt_UsuContra) & "'," & Val(Trim(.Lbl_UsuCvePerfil.Caption)) & "," & F_Sistema & "," & _
    F_Sistema & "," & Val(Trim(.Lbl_UsuCveEstatus.Caption)) & "," & IIf(Val(.txtCampo(0).Tag) = 0, "null", Val(.txtCampo(0).Tag)) & "," & IIf(Val(.txtCampo(1).Tag) = 0, "null", Val(.txtCampo(1).Tag)) & "," & IIf(Val(.txtCampo(2).Tag) = 0, "null", Val(.txtCampo(2).Tag)) & ")"
End With

gcn.Execute sQuery
MsgBox "Registro grabado satisfactoriamente.", 0 + 64, "Usuarios"
Call PR_GrabaBitacora(1, 1, "c_segusuarios", Trim(Frm_Usuarios.Txt_UsuClave))
Exit Sub

ErrAccionG:
If Err.Number <> 0 Then
    MsgBox "Ha ocurrido un error al guardar los datos", 0 + 32, "[PR_GuardaRegUsarios]"
    Exit Sub
End If
'Graba en c_segusuarios
End Sub
Sub PR_GuardaResumen()
Dim S_CadenaBase As String, S_CadenaInsert  As String, S_CadenaEje As String, S_ValorX As Variant
Dim S_Borrar As String, S_ComBita As String, S_DescripBita As String

On Error GoTo GrabaResumen

With Frm_Resumen
    S_Borrar = "Delete From t_resumen " & _
    "Where n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And s_numcuestionario = '" & gs_Cuestionario & "'"
     gcn.Execute S_Borrar
    
    S_CadenaBase = "Insert into T_Resumen Values (" & gs_Proyecto & "," & gs_Etapa & ",'" & gs_Cuestionario & "',"

    N_Reg = Val(.MSFlexGrid1.Rows - 1)
    N_Col = Val(.MSFlexGrid1.Cols - 1)
    S_CadenaInsert = ""
    For li_Col = 1 To N_Col
        S_CadenaInsert = Trim(.MSFlexGrid1.TextMatrix(1, li_Col)) & ","
        
        For li_Row = 2 To N_Reg
            If Trim(.MSFlexGrid1.TextMatrix(li_Row, li_Col)) = "" Then
                S_ValorX = "Null"
            Else
                S_ValorX = Trim(.MSFlexGrid1.TextMatrix(li_Row, li_Col))
            End If
            S_CadenaInsert = S_CadenaInsert & S_ValorX & ","
        Next
        S_CadenaInsert = Val(li_Col) & "," & Left(S_CadenaInsert, Len(Trim(S_CadenaInsert)) - 1) & ")"
        S_CadenaEje = S_CadenaBase & S_CadenaInsert
      
        gcn.Execute S_CadenaEje
    Next
End With

S_ComBita = gs_Proyecto & "-" & gs_Etapa & "-" & Trim(gs_Cuestionario)
If gs_ProcCuestionario = 113 Then
  
    MsgBox "Registros del Resumen de Contactos, actualizados satisfactoriamente.", 0 + 64, "Resumen de Contactos"
    S_DescripBita = "t_resumen" & " - Actualización"
    Call PR_GrabaBitacora(10, 3, S_DescripBita, S_ComBita)
  
    S_ActualizaEnc = "Update t_enccuestionario set n_exportar = 2 " & _
    "Where n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And s_numcuestionario = '" & Trim(gs_Cuestionario) & "'"
    gcn.Execute S_ActualizaEnc
  
Else
  
    MsgBox "Resumen de Contactos guardado satisfactoriamente.", 0 + 64, "Resumen de Contactos"
    S_DescripBita = "t_resumen"
    Call PR_GrabaBitacora(10, 1, S_DescripBita, S_ComBita)
End If

Exit Sub

GrabaResumen:
If Err.Number <> 0 Then
    MsgBox Str(Err.Number) & ": " & Err.Description, 0 + 64, "PR_GuardaResumen"
    Exit Sub
End If
'***********************************************************************************************************************
'*Guarda información del resumen de contactos
'***********************************************************************************************************************
End Sub
Sub PR_LimpiaCuestionario(S_CualPantalla)
Dim N_Ctl As Integer
Select Case UCase(Trim(S_CualPantalla))
    Case "CUESTIONARIO-PARCIAL"
            With Frm_CuestionarioDin
              
                For N_Ctl = 0 To Val(.Txt_CueDinMod1.UBound)
                    .Txt_CueDinMod1(N_Ctl) = ""
                Next N_Ctl
                
              
                For N_Ctl = 0 To Val(.Txt_CueDinMod45.UBound)
                    .Txt_CueDinMod45(N_Ctl) = ""
                  
                  
                  
                  
                  
                  
                  
                Next N_Ctl
                
              
                For N_Ctl = 0 To Val(.Txt_CueDinMod6.UBound)
                    .Txt_CueDinMod6(N_Ctl) = ""
                Next N_Ctl
                
              
                For N_Ctl = 0 To Val(.Txt_CueDinMod7.UBound)
                    .Txt_CueDinMod7(N_Ctl) = ""
                Next N_Ctl
                
              
                For N_Ctl = 0 To Val(.Txt_CueDinMod8.UBound)
                    .Txt_CueDinMod8(N_Ctl) = ""
                Next N_Ctl
        End With
        
    Case "CUESTIONARIO-COMPLETO"
        With Frm_CuestionarioDin
          
            For N_Ctl = 0 To Val(.Txt_CueDinMod1.UBound)
                .Txt_CueDinMod1(N_Ctl) = ""
                .Txt_CueDinMod1(N_Ctl).Tag = ""
            Next N_Ctl
            
          
            For N_Ctl = 0 To Val(.Txt_CueDinMod45.UBound)
                .Txt_CueDinMod45(N_Ctl) = ""
                .Txt_CueDinMod45(N_Ctl).Tag = ""
            Next N_Ctl
            
          
            For N_Ctl = 0 To Val(.Txt_CueDinMod6.UBound)
                .Txt_CueDinMod6(N_Ctl) = ""
                .Txt_CueDinMod6(N_Ctl).Tag = ""
            Next N_Ctl
            
          
            For N_Ctl = 0 To Val(.Txt_CueDinMod7.UBound)
                .Txt_CueDinMod7(N_Ctl) = ""
                .Txt_CueDinMod7(N_Ctl).Tag = ""
            Next N_Ctl
            
          
            For N_Ctl = 0 To Val(.Txt_CueDinMod8.UBound)
                .Txt_CueDinMod8(N_Ctl) = ""
                .Txt_CueDinMod8(N_Ctl).Tag = ""
            Next N_Ctl
        End With
        
    Case "OTRO"
End Select

End Sub
Sub PR_PrimerReg_Eta()
If RE_Etapas.BOF = True Then
   MsgBox "No hay más registros", vbInformation, "Etapas"
Else
   RE_Etapas.MoveFirst
   Call PR_AcomodaDatosEtapas
   Exit Sub
End If
End Sub
Sub PR_SiguienteReg_Eta()
If RE_Etapas.EOF = True Then
   MsgBox "No hay más registros", vbInformation, "Etapas"
Else
   RE_Etapas.MoveNext
   If RE_Etapas.EOF = True Then
        MsgBox "No hay más registros", vbInformation, "Etapas"
        RE_Etapas.MovePrevious
        Exit Sub
   End If
   Call PR_AcomodaDatosEtapas
   Exit Sub
End If
End Sub
Sub PR_UltimoReg_Usu()
'*USUARIOS
If RE_Usuarios.EOF = True Then
   MsgBox "No hay más registros", vbInformation, "Usuarios"
Else
   RE_Usuarios.MoveLast
   Call PR_AcomodaDatosUsuarios
   Exit Sub
End If
End Sub

Sub PR_UltimoReg_Eta()
If RE_Etapas.EOF = True Then
   MsgBox "No hay más registros", vbInformation, "Etapas"
Else
   RE_Etapas.MoveLast
   Call PR_AcomodaDatosEtapas
   Exit Sub
End If
End Sub
Sub PR_ConsultaInfPerfiles()
'**PERFILES**
Dim sQuery As String
Dim S_Condi As String, L_CuantosReg As Long
Dim S_Select As String, S_From As String, S_Where As String, S_Order As String
Dim S_Sel, S_Fro, S_Whe As String, S_Como As String

S_Sel = ""
S_Fro = ""
S_Whe = ""

S_Select = "SELECT a.n_cveperfil,a.s_descrip,a.s_comentario,a.n_estatuscat,e.s_descrip "

S_From = "FROM c_perfil a (Nolock), c_estatuscat e (Nolock) "
S_Where = "WHERE a.n_estatuscat = e.n_estatuscat "
S_Como = FU_SeleccionQry("Perfiles")

If S_Como <> "" Then
    S_Where = S_Where & S_Como
End If

S_Order = " ORDER BY a.n_cveperfil"
sQuery = S_Select & S_From & S_Where & S_Order
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Call PR_DestruyeCursoGral("C_Perfil")
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Set RE_Perfiles = New ADODB.Recordset
RE_Perfiles.Open sQuery, gcn

 On Error GoTo Errorconsulta
         If RE_Perfiles.BOF And RE_Perfiles.EOF Then
             MsgBox "No existen registros con este criterio de búsqueda", vbInformation, "Perfiles"
             BOEstado = True
             Frm_Perfiles.ToolbarPerfiles.Buttons(1).ToolTipText = "Grabar"
             Frm_Perfiles.ToolbarPerfiles.Buttons(3).Enabled = False
             Frm_Perfiles.ToolbarPerfiles.Buttons(6).Enabled = False
             Frm_Perfiles.ToolbarPerfiles.Buttons(7).Enabled = False
             Exit Sub
         End If
         RE_Perfiles.MoveLast
         Frm_Perfiles.txt_TotRegPerfiles = RE_Perfiles.RecordCount
         RE_Perfiles.MoveFirst
         Call PR_AcomodaDatosPerfiles
Exit Sub

Errorconsulta:
     MsgBox "Descripción del error" & Err.Description & "Número de error:" & Err.Number
     Exit Sub
End Sub
Sub PR_ConsultaInfEtapas()
'**ETAPAS**
Dim sQuery As String
Dim S_Condi As String, L_CuantosReg As Long
Dim S_Select As String, S_From As String, S_Where As String, S_Order As String
Dim S_Sel, S_Fro, S_Whe As String, S_Como As String

S_Sel = ""
S_Fro = ""
S_Whe = ""

S_Select = "SELECT a.n_cveetapa,a.s_descrip,a.s_comentario,a.n_estatuscat,e.s_descrip "

S_From = "FROM c_etapa a (Nolock),c_estatuscat e (Nolock) "
S_Where = "WHERE a.n_estatuscat = e.n_estatuscat"
S_Como = FU_SeleccionQry("Etapas")

If S_Como <> "" Then
    S_Where = S_Where & S_Como
End If

S_Order = " ORDER BY a.n_cveetapa"
sQuery = S_Select & S_From & S_Where & S_Order
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Call PR_DestruyeCursoGral("C_Etapa")
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Set RE_Etapas = New ADODB.Recordset
RE_Etapas.Open sQuery, gcn

 On Error GoTo Errorconsulta
         If RE_Etapas.BOF And RE_Etapas.EOF Then
             MsgBox "No existen registros con este criterio de búsqueda", vbInformation, "Etapas"
             BOEstado = True
             Frm_Etapas.ToolbarEtapas.Buttons(1).ToolTipText = "Grabar"
             Frm_Etapas.ToolbarEtapas.Buttons(3).Enabled = False
             Frm_Etapas.ToolbarEtapas.Buttons(6).Enabled = False
             Frm_Etapas.ToolbarEtapas.Buttons(7).Enabled = False
             Exit Sub
         End If
         RE_Etapas.MoveLast
         Frm_Etapas.txt_TotRegEtapas = RE_Etapas.RecordCount
         RE_Etapas.MoveFirst
         Call PR_AcomodaDatosEtapas
Exit Sub

Errorconsulta:
     MsgBox "Descripción del error" & Err.Description & "Número de error:" & Err.Number
     Exit Sub
End Sub


Sub PR_AnteriorReg_Per()
'*PERFILES
If RE_Perfiles.BOF = True Then
   MsgBox "No hay más registros", vbInformation, "Perfiles"
Else
   RE_Perfiles.MovePrevious
   If RE_Perfiles.BOF = True Then
        MsgBox "No hay más registros", vbInformation, "Perfiles"
        RE_Perfiles.MoveNext
        Exit Sub
   End If
   
   Call PR_AcomodaDatosPerfiles
   Exit Sub
End If
End Sub

Sub PR_PrimerReg_Per()
'*PERFILES
If RE_Perfiles.BOF = True Then
   MsgBox "No hay más registros", vbInformation, "Perfiles"
Else
   RE_Perfiles.MoveFirst
   Call PR_AcomodaDatosPerfiles
   Exit Sub
End If
End Sub

Sub PR_SiguienteReg_Usu()
'*USUARIOS
If RE_Usuarios.EOF = True Then
   MsgBox "No hay más registros", vbInformation, "Usuarios"
Else
   RE_Usuarios.MoveNext
   If RE_Usuarios.EOF = True Then
        MsgBox "No hay más registros", vbInformation, "Usuarios"
        RE_Usuarios.MovePrevious
        Exit Sub
   End If
   Call PR_AcomodaDatosUsuarios
   Exit Sub
End If
End Sub
Sub PR_SiguienteReg_Per()
'*PERFILES
If RE_Perfiles.EOF = True Then
   MsgBox "No hay más registros", vbInformation, "Perfiles"
Else
   RE_Perfiles.MoveNext
   If RE_Perfiles.EOF = True Then
        MsgBox "No hay más registros", vbInformation, "Perfiles"
        RE_Perfiles.MovePrevious
        Exit Sub
   End If
   Call PR_AcomodaDatosPerfiles
   Exit Sub
End If
End Sub
Sub PR_UltimoReg_Per()
'*PERFILES
If RE_Perfiles.EOF = True Then
   MsgBox "No hay más registros", vbInformation, "Perfiles"
Else
   RE_Perfiles.MoveLast
   Call PR_AcomodaDatosPerfiles
   Exit Sub
End If
End Sub
Sub PR_DestruyeCursoGral(Cual As String)

On Error GoTo ErrorCerrar

Select Case UCase(Trim(Cual))
    Case "C_ETAPA"
        RE_Etapas.Close
        Set RE_Etapas = nothig
        Exit Sub
        
    Case "C_PERFIL"
        RE_Perfiles.Close
        Set RE_Perfiles = nothig
        Exit Sub
        
    Case "C_SEGUSUARIOS"
        RE_Usuarios.Close
        Set RE_Usuarios = nothig
        Exit Sub
        
    Case "T_PROYECTOS"
        RE_Proyectos.Close
        Set RE_Proyectos = nothig
        Exit Sub
End Select

ErrorCerrar:
If Err.Number <> 0 Then Exit Sub
'*******************************************************************************************
'*Sirve para cerrar el cursor general, siempre y cuando este abierto
'*******************************************************************************************
End Sub

Function FU_SeleccionQry(S_MiSeleccion As String) As String
Dim S_Patron As String, S_conector As String

S_Patron = ""
Select Case UCase(Trim(S_MiSeleccion))
    Case "ETAPAS"
        S_conector = " AND"
        With Frm_Etapas
            If Trim(.Txt_EtaCveEta.Text) <> "" Then
                S_Patron = S_Patron & " And a.n_cveetapa = " & Trim(.Txt_EtaCveEta)
            End If
            
            If Trim(.Txt_EtaDescrip.Text) <> "" Then
                S_Patron = S_Patron & " And ltrim(a.s_descrip) Like '" & Trim(.Txt_EtaDescrip) & "%'"
            End If
            
        End With
        FU_SeleccionQry = S_Patron
    
    Case "PERFILES"
        S_conector = " AND"
        With Frm_Perfiles
            If Trim(.Txt_PerCvePer.Text) <> "" Then
                S_Patron = S_Patron & " AND a.n_cveperfil = " & Trim(.Txt_PerCvePer)
            End If
                                    
            If Trim(.Txt_PerDescrip.Text) <> "" Then
                S_Patron = S_Patron & " AND ltrim(a.s_descrip) Like '" & Trim(.Txt_PerDescrip) & "%'"
               
            End If
            
        End With
        FU_SeleccionQry = S_Patron
        
    Case "USUARIOS"
        With Frm_Usuarios
            If Trim(.Txt_UsuClave.Text) <> "" Then
                S_Patron = S_Patron & " AND ltrim(s.n_cveusuario) Like '" & Trim(.Txt_UsuClave) & "%'"
            End If
            
            If Trim(.Txt_UsuNombre.Text) <> "" Then
                S_Patron = S_Patron & " AND ltrim(s.s_nombre) Like '" & Trim(.Txt_UsuNombre) & "%'"
            End If
            
        End With
        FU_SeleccionQry = S_Patron
    
    Case "PROYECTOS"
        With Frm_Proyecto
            If Trim(.Txt_ProyNumero.Text) <> "" Then
                S_Patron = S_Patron & " AND ltrim(p.n_cveproyecto) Like '" & Trim(.Txt_ProyNumero) & "%'"
            End If
            
            If Trim(.Txt_ProyNombre.Text) <> "" Then
                S_Patron = S_Patron & " AND ltrim(p.s_nombre) Like '" & Trim(.Txt_ProyNombre) & "%'"
            End If
            
        End With
        FU_SeleccionQry = S_Patron
        
    Case "OTRO"
    
End Select
'*******************************************************************************************
'*
'*******************************************************************************************
End Function

Function FUconsecutivo(ps_Tabla As String, ps_Campo As String) As Long
Dim rdo           As ADODB.Recordset
Dim sClave        As Long
Dim sConsec       As String

sConsec = "SELECT Max(" & ps_Campo & ") FROM  " & ps_Tabla & ""

Set rdo = New ADODB.Recordset
rdo.Open sConsec, gcn

On Error GoTo ErrNvaClave

If Val(rdo.RecordCount) > 0 Then
      FUconsecutivo = IIf(IsNull(rdo(0)), 0, rdo(0)) + 1
Else
      FUconsecutivo = 1
End If
rdo.Close
Exit Function

ErrNvaClave:
    FUconsecutivo = 0
    Exit Function

End Function

Function FU_vte_EsNumero(pintNum As Integer) As Integer
'******************************************************************************************
'Función                    : gf_EsNumero
'Autor                      : J.M.F.
'Descripción                : Valida si el parametro esta en el rango de asccii de números
'Fecha de Creación          : 26/Enero/1999
'Fecha de Liberación        : 26/Enero/1999
'Fecha de Modificación      :
'Autor de la Modificación   :
'Usuario que solicita la modificación:
'
' Parámetros:    Tipo        Nombre               Descripción
'    Entrada:   Integer      pintNum              Valor númerico
'     salida:   Integer      gfint_vte_EsNumero   pintNum si esta en el rango
'          0:   En caso contrario
'******************************************************************************************
    FU_vte_EsNumero = 0
    If pintNum >= 48 And pintNum <= 57 Then FU_vte_EsNumero = pintNum
End Function

Function FU_vte_EnCadena(pintNum As Integer) As Integer
'******************************************************************************************
'Función                    : gfint_vte_EnCadena
'Autor                      : J.M.F.
'Descripción                : Valida si el parametro esta en el rango de asccii en la cadena
'Fecha de Creación          : 26/Enero/1999
'Fecha de Liberación        : 26/Enero/1999
'Fecha de Modificación      :
'Autor de la Modificación   :
'Usuario que solicita la modificación:
'
' Parámetros:   Tipo       Nombre        Descripción
'    Entrada:   Integer    pintNum       Valor númerico
'               String     pstrCad       String donde se busca
'     salida:   Integer    gp_Escadena   pintNum si esta en el rango
'                                        0: En caso contrario
'******************************************************************************************
Dim lstrCar As String * 1 'Ascci convertido en el caracter
   FU_vte_EnCadena = 0
   
'
'
'
      
    If pintNum >= 33 And pintNum <= 191 Then
       FU_vte_EnCadena = Asc(UCase(Chr(pintNum)))

    End If
End Function

Function FU_RescataConsecutivoTablaX(S_TablaX As String) As Long
Dim RE_TablaFuente As ADODB.Recordset, S_Condi As String, S_MiTablaArg As String

FU_RescataConsecutivoTablaX = 0

Select Case UCase(Trim(S_TablaX))
    Case "C_ETAPA"
        S_Condi = "select Max(n_cveetapa) from " & S_TablaX & " where n_cveetapa < 99"

        Set RE_TablaFuente = New ADODB.Recordset
        RE_TablaFuente.Open S_Condi, gcn
        
        If Not RE_TablaFuente.EOF() Then
            FU_RescataConsecutivoTablaX = RE_TablaFuente(0) + 1
        End If
        RE_TablaFuente.Close
        Set RE_TablaFuente = Nothing
        Exit Function
        
    Case "C_PERFIL"
        S_Condi = "select Max(n_cveperfil) from " & S_TablaX

        Set RE_TablaFuente = New ADODB.Recordset
        RE_TablaFuente.Open S_Condi, gcn
        
        If Not RE_TablaFuente.EOF() Then
            FU_RescataConsecutivoTablaX = RE_TablaFuente(0) + 1
        End If
        RE_TablaFuente.Close
        Set RE_TablaFuente = Nothing
        Exit Function
        
    Case "OTRO"
End Select
'********************************************************************************************
'*
'********************************************************************************************
End Function
Function Ejecuta_Query(sQuery As String) As Boolean
Dim QueryCmd As ADODB.Command

On Error GoTo queryerror
Set QueryCmd = New ADODB.Command

Set QueryCmd.ActiveConnection = gcn
QueryCmd.CommandText = sQuery
QueryCmd.Prepared = True
QueryCmd.Execute

Ejecuta_Query = True
Exit Function

queryerror:
    MsgBox "Error No." & Err.Number & " " & Err.Description, vbCritical, "Error en el proceso de actualización en la base de datos"
  
    Ejecuta_Query = False
    Exit Function
    Resume
End Function
Sub PR_GuardaRegEtapas()
Dim sQuery As String
Dim N_claveexp As String

On Error GoTo ErrAccionG

With Frm_Etapas
    N_claveexp = FU_AgregaCeros_Izquierda(2, Trim(.Txt_EtaCveEta))
    sQuery = "Insert Into c_etapa Values (" & Val(Trim(.Txt_EtaCveEta)) & ",'" & Trim(.Txt_EtaDescrip) & "','" & Trim(.Txt_EtaDescripLar) & "','" & N_claveexp & "'," & Val(.Lbl_EtaCveEstatus.Caption) & ")"
End With

gcn.Execute sQuery
MsgBox "Registro grabado satisfactoriamente.", 0 + 64, ""

Exit Sub

ErrAccionG:
If Err.Number <> 0 Then
    MsgBox "Ha ocurrido un error al guardar los datos", 0 + 32, "[PR_GuardaRegEtapas]"
    Exit Sub
End If
'Graba en c_etapas
End Sub

Sub PR_GuardaRegPerfiles()
Dim sQuery As String
Dim N_claveexp As String

On Error GoTo ErrAccionP
With Frm_Perfiles
    N_claveexp = FU_AgregaCeros_Izquierda(2, Trim(.Txt_PerCvePer))
    sQuery = "Insert Into c_perfil Values (" & Val(Trim(.Txt_PerCvePer)) & ",'" & Trim(.Txt_PerDescrip) & "','" & Trim(.Txt_PerDescripLar) & "','" & N_claveexp & "'," & Val(.Lbl_PerCveEstatus.Caption) & ")"
End With

gcn.Execute sQuery
MsgBox "Registro grabado satisfactoriamente.", 0 + 64, "Perfiles"

Exit Sub

ErrAccionP:
If Err.Number <> 0 Then
    MsgBox "Ha ocurrido un error al guardar los datos", 0 + 32, "[PR_GuardaRegPerfiles]"
    Exit Sub
End If
'Graba en c_perfil
End Sub
Sub PR_AnteriorReg_Usu()
'*USUARIOS
If RE_Usuarios.BOF = True Then
   MsgBox "No hay más registros", vbInformation, "Usuarios"
Else
   RE_Usuarios.MovePrevious
   If RE_Usuarios.BOF = True Then
        MsgBox "No hay más registros", vbInformation, "Usuarios"
        RE_Usuarios.MoveNext
        Exit Sub
   End If
   
   Call PR_AcomodaDatosUsuarios
   Exit Sub
End If
End Sub
Sub PR_PrimerReg_Usu()
'*USUARIOS
If RE_Usuarios.BOF = True Then
   MsgBox "No hay más registros", vbInformation, "Usuarios"
Else
   RE_Usuarios.MoveFirst
   Call PR_AcomodaDatosUsuarios
   Exit Sub
End If
End Sub
Function FU_ExtraeNombreUsuario(N_cveusuario As Long) As String
Dim S_Condi As String, RE_InformNumX As ADODB.Recordset

On Error GoTo ErrRecataInfNX
FU_ExtraeNombreUsuario = ""
S_Condi = "Select s_nombre,s_paterno,s_materno FROM c_segusuarios Where n_cveusuario = " & N_cveusuario

Set RE_InformNumX = New ADODB.Recordset
RE_InformNumX.Open S_Condi, gcn

If Not RE_InformNumX.EOF() Then
    FU_ExtraeNombreUsuario = Trim(RE_InformNumX(0)) & " " & Trim(RE_InformNumX(1)) & " " & Trim(RE_InformNumX(2))
End If
RE_InformNumX.Close
Set RE_InformNumX = Nothing
Exit Function

ErrRecataInfNX:
If Err.Number <> 0 Then
    MsgBox Str(Err.Number) & ": " & Err.Description, 0 + 64, "FU_ExtraeNombreUsuario"
    Exit Function
End If
'************************************************************************************************************************
'*
'***********************************************************************************************************************
End Function
Sub PR_AnteriorReg_Pro()
'*PROYECTOS
If RE_Proyectos.BOF = True Then
   MsgBox "No hay más registros", vbInformation, "Proyectos"
Else
   RE_Proyectos.MovePrevious
   If RE_Proyectos.BOF = True Then
        MsgBox "No hay más registros", vbInformation, "Proyectos"
        RE_Proyectos.MoveNext
        Exit Sub
   End If
   
   Call PR_AcomodaDatosProyectos
   Exit Sub
End If
End Sub

Sub PR_PrimerReg_Pro()
'*PROYECTOS
If RE_Proyectos.BOF = True Then
   MsgBox "No hay más registros", vbInformation, "Proyectos"
Else
   RE_Proyectos.MoveFirst
   Call PR_AcomodaDatosProyectos
   Exit Sub
End If
End Sub

Sub PR_SiguienteReg_Pro()
'*PROYECTOS
If RE_Proyectos.EOF = True Then
   MsgBox "No hay más registros", vbInformation, "Proyectos"
Else
   RE_Proyectos.MoveNext
   If RE_Proyectos.EOF = True Then
        MsgBox "No hay más registros", vbInformation, "Proyectos"
        RE_Proyectos.MovePrevious
        Exit Sub
   End If
   Call PR_AcomodaDatosProyectos
   Exit Sub
End If
End Sub

Sub PR_UltimoReg_Pro()
'*PROYECTOS
If RE_Proyectos.EOF = True Then
   MsgBox "No hay más registros", vbInformation, "Proyectos"
Else
   RE_Proyectos.MoveLast
   Call PR_AcomodaDatosProyectos
   Exit Sub
End If
End Sub
Sub PR_GuardaRegProyectos()
Dim sQuery As String
Dim F_Sistema

'n_cveproyecto , s_nombre, n_cveestatus, n_cveejec, n_cveprog, n_productividad, s_comentario, n_interno
On Error GoTo ErrAccionP

F_Sistema = "'" & Format(FU_ExtraeFechaServer(), "dd/mm/yyyy hh:mm:ss") & "'"
With Frm_Proyecto
    sQuery = "Insert Into t_proyectos Values (" & _
    Val(Trim(.Txt_ProyNumero)) & ",'" & Trim(.Txt_ProyNombre) & "'," & Val(.Lbl_ProyCveEstatus.Caption) & "," & _
    Val(Trim(.Txt_ProyNumEjecutivoSC)) & "," & Val(Trim(.Txt_ProyNumEjecutivoPR)) & "," & Trim(.Txt_ProyNumProg) & "," & Val(Trim(.Txt_ProyProd)) & ",'" & _
    Trim(.Txt_ProyComen) & "',0)"
End With

gcn.Execute sQuery
MsgBox "Registro grabado satisfactoriamente.", 0 + 64, "Proyectos"
Call PR_GrabaBitacora(4, 1, "t_proyectos", Trim(Frm_Proyecto.Txt_ProyNumero))
Exit Sub

ErrAccionP:
If Err.Number <> 0 Then
    MsgBox "Ha ocurrido un error al guardar los datos", 0 + 32, "[PR_GuardaRegProyectos]"
    Exit Sub
End If
'Graba en t_proyectos
End Sub

Sub PR_ConsultaInfProyectos()
'**PROYECTOS**
Dim sQuery As String
Dim S_Condi As String, L_CuantosReg As Long
Dim S_Select As String, S_From As String, S_Where As String, S_Order As String
Dim S_Sel, S_Fro, S_Whe As String, S_Como As String

S_Sel = ""
S_Fro = ""
S_Whe = ""

S_Select = "SELECT p.n_cveproyecto,p.s_nombre,p.n_cveestatus,p.n_cveejec_sc,p.n_cveejec_pr,p.n_cveprog,p.n_productividad,p.s_comentario,e.s_descrip "
S_From = "FROM t_proyectos p (Nolock),c_estatus e (Nolock) "
S_Where = "WHERE p.n_cveestatus = e.n_cveestatus "
S_Como = FU_SeleccionQry("Proyectos")

If S_Como <> "" Then
    S_Where = S_Where & S_Como
End If

S_Order = " ORDER BY p.n_cveproyecto"
sQuery = S_Select & S_From & S_Where & S_Order
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Call PR_DestruyeCursoGral("t_proyectos")
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Set RE_Proyectos = New ADODB.Recordset
RE_Proyectos.Open sQuery, gcn

 On Error GoTo Errorconsulta
         If RE_Proyectos.BOF And RE_Proyectos.EOF Then
             MsgBox "No existen registros con este criterio de búsqueda", vbInformation, "Proyectos"
             BOEstado = True
             Frm_Proyecto.ToolbarProyectos.Buttons(1).ToolTipText = "Grabar"
             Frm_Proyecto.ToolbarProyectos.Buttons(3).Enabled = False
             Frm_Proyecto.ToolbarProyectos.Buttons(6).Enabled = False
             Frm_Proyecto.ToolbarProyectos.Buttons(7).Enabled = False
             Exit Sub
         End If
         RE_Proyectos.MoveLast
         Frm_Proyecto.txt_TotRegProy = RE_Proyectos.RecordCount
         RE_Proyectos.MoveFirst
         Call PR_AcomodaDatosProyectos
         Call PR_GrabaBitacora(4, 4, Trim(CStr(RE_Proyectos.RecordCount)) & " Registros", "")
Exit Sub

Errorconsulta:
     MsgBox "Descripción del error" & Err.Description & "Número de error:" & Err.Number
     Exit Sub

End Sub
Sub PR_ConsultaInfUsuarios()
'**USUARIOS**
Dim sQuery As String
Dim S_Condi As String, L_CuantosReg As Long
Dim S_Select As String, S_From As String, S_Where As String, S_Order As String
Dim S_Sel, S_Fro, S_Whe As String, S_Como As String

S_Sel = ""
S_Fro = ""
S_Whe = ""

S_Select = "SELECT s.n_cveusuario,s.s_nombre,s.s_paterno,s.s_materno,s.s_pwd,s.n_cveperfil,s.f_alta,s.f_movi,p.s_descrip,s.n_estatuscat,e.s_descrip, n_cvepuesto "

S_From = "FROM c_segusuarios s (Nolock), c_perfil p (Nolock), c_estatuscat e (Nolock) "
S_Where = "WHERE s.n_cveperfil = p.n_cveperfil And s.n_estatuscat = e.n_estatuscat  "
S_Como = FU_SeleccionQry("Usuarios")

If S_Como <> "" Then
    S_Where = S_Where & S_Como
End If

S_Order = " ORDER BY s.n_cveusuario"
sQuery = S_Select & S_From & S_Where & S_Order
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Call PR_DestruyeCursoGral("C_SegUsuarios")
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Set RE_Usuarios = New ADODB.Recordset
RE_Usuarios.Open sQuery, gcn

 On Error GoTo Errorconsulta
         If RE_Usuarios.BOF And RE_Usuarios.EOF Then
             MsgBox "No existen registros con este criterio de búsqueda", vbInformation, "Usuarios"
             BOEstado = True
             Frm_Usuarios.ToolbarUsuarios.Buttons(1).ToolTipText = "Grabar"
             Frm_Usuarios.ToolbarUsuarios.Buttons(3).Enabled = False
             Frm_Usuarios.ToolbarUsuarios.Buttons(6).Enabled = False
             Frm_Usuarios.ToolbarUsuarios.Buttons(7).Enabled = False
             Exit Sub
         End If
         RE_Usuarios.MoveLast
         Frm_Usuarios.txt_TotRegUsuarios = RE_Usuarios.RecordCount
         RE_Usuarios.MoveFirst
         Call PR_AcomodaDatosUsuarios
         Call PR_GrabaBitacora(1, 4, Trim(CStr(RE_Usuarios.RecordCount)) & " Registros", "")
Exit Sub

Errorconsulta:
     MsgBox "Descripción del error" & Err.Description & "Número de error:" & Err.Number
     Exit Sub
End Sub
Sub PR_GuardaRegEnc_Cuestionario()
Dim sQuery As String, S_ComBita As String, F_Sistema, S_Marca_Prue_Prod As String

On Error GoTo ErrAccionC
If gs_ProcCuestionario = 1 Then
    S_Marca_Prue_Prod = "P"
ElseIf gs_ProcCuestionario = 2 Then
    S_Marca_Prue_Prod = "X"
End If
F_Sistema = "'" & Format(FU_ExtraeFechaServer(), "dd/mm/yyyy hh:mm:ss") & "'"
With Frm_Cuestionario
    S_ComBita = Trim(.Cmb_CueProyecto) & "-" & Trim(.Lbl_CveCueEtapa.Caption)
    sQuery = "Insert Into t_enccuestionario Values (" & _
    Val(Trim(.Cmb_CueProyecto)) & "," & Val(Trim(.Lbl_CveCueEtapa.Caption)) & "," & _
    Trim(.Txt_CueCuestionario) & "," & F_Sistema & "," & F_Sistema & ",'" & S_Marca_Prue_Prod & "',0)"
End With

gcn.Execute sQuery
MsgBox "Registro grabado satisfactoriamente.", 0 + 64, "Encabezado del Cuestionario"
Call PR_GrabaBitacora(7, 1, "t_enccuestionario", S_ComBita)
Exit Sub

ErrAccionC:
If Err.Number <> 0 Then
    MsgBox "Ha ocurrido un error al guardar los datos", 0 + 32, "[PR_GuardaRegEnc_Cuestionario]"
    Exit Sub
End If
'Graba en t_enccuestionario
End Sub

Sub PR_GuardaRegEnc_Caratula()
Dim sQuery As String, S_ComBita As String, F_Sistema

On Error GoTo ErrAccionC

F_Sistema = "'" & Format(FU_ExtraeFechaServer(), "dd/mm/yyyy hh:mm:ss") & "'"
With Frm_Caratula
    S_ComBita = Trim(Frm_Caratula.Txt_CarProyecto) & "-" & Trim(Frm_Caratula.Lbl_CveEtapa.Caption)
    sQuery = "Insert Into t_enccaratulas Values (" & _
    Val(Trim(.Txt_CarProyecto)) & "," & Val(Trim(.Lbl_CveEtapa.Caption)) & "," & _
    Val(Trim(.Lbl_CarCveEstatus.Caption)) & "," & F_Sistema & "," & F_Sistema & ",0)"
End With

gcn.Execute sQuery
MsgBox "Registro grabado satisfactoriamente.", 0 + 64, "Encabezado de la Carátula"
Call PR_GrabaBitacora(5, 1, "t_enccaratulas", S_ComBita)
Exit Sub

ErrAccionC:
If Err.Number <> 0 Then
    MsgBox "Ha ocurrido un error al guardar los datos", 0 + 32, "[PR_GuardaRegEnc_Caratula]"
    Exit Sub
End If
'Graba en t_enccaratulas
End Sub

Function FU_BuscaValorEnGrid(S_CualGrid As String, N_ValParam As Integer) As Integer
Dim N_Reglones As Integer, li_Row As Integer, N_Modulo As Integer, S_ValorPadre As Variant

FU_BuscaValorEnGrid = 0

Select Case UCase(Trim(S_CualGrid))
    Case "MODULO1"
        With Frm_Caratula
            N_Reglones = Val(.Grd_Mod1.Rows - 1)
            For li_Row = 1 To N_Reglones
                If Val(N_ValParam) = Val(Trim(.Grd_Mod1.TextMatrix(li_Row, 2))) Then
                    FU_BuscaValorEnGrid = 1
                    Exit For
                End If
            Next
        End With
            
    Case "MODULO2"
        With Frm_Caratula
            N_Reglones = Val(.Grd_Mod2.Rows - 1)
            For li_Row = 1 To N_Reglones
                If Val(N_ValParam) = Val(Trim(.Grd_Mod2.TextMatrix(li_Row, 2))) Then
                    FU_BuscaValorEnGrid = 1
                    Exit For
                End If
            Next
        End With
        
    Case "MODULO3"
        With Frm_Caratula
            N_Reglones = Val(.Grd_Mod3.Rows - 1)
            For li_Row = 1 To N_Reglones
                If Val(N_ValParam) = Val(Trim(.Grd_Mod3.TextMatrix(li_Row, 2))) Then
                    FU_BuscaValorEnGrid = 1
                    Exit For
                End If
            Next
        End With
        
    Case "MODULO4"
        With Frm_Caratula
            N_Reglones = Val(.Grd_Mod4.Rows - 1)
            For li_Row = 1 To N_Reglones
                If Val(N_ValParam) = Val(Trim(.Grd_Mod4.TextMatrix(li_Row, 2))) Then
                    FU_BuscaValorEnGrid = 1
                    Exit For
                End If
            Next
        End With
        
    Case "MODULO5"
        With Frm_Caratula
            S_ValorPadre = Trim(.Lbl_Mod5Opcion.Tag)
            N_Reglones = Val(.Grd_Mod5.Rows - 1)
            For li_Row = 1 To N_Reglones
                If S_ValorPadre & Trim(N_ValParam) = Trim(.Grd_Mod5.TextMatrix(li_Row, 4)) & Trim(.Grd_Mod5.TextMatrix(li_Row, 3)) Then
                    FU_BuscaValorEnGrid = 1
                    Exit For
                End If
            Next
        End With
        
    Case "MODULO8"
        With Frm_Caratula
            N_Reglones = Val(.Grd_Mod8.Rows - 1)
            For li_Row = 1 To N_Reglones
                If Val(N_ValParam) = Val(Trim(.Grd_Mod8.TextMatrix(li_Row, 2))) Then
                    FU_BuscaValorEnGrid = 1
                    Exit For
                End If
            Next
        End With
        
    Case "MODULO9"
        With Frm_Caratula
            N_Reglones = Val(.Grd_Mod9.Rows - 1)
            For li_Row = 1 To N_Reglones
                If Val(N_ValParam) = Val(Trim(.Grd_Mod9.TextMatrix(li_Row, 2))) Then
                    FU_BuscaValorEnGrid = 1
                    Exit For
                End If
            Next
        End With
End Select
End Function
Sub PR_AcomodaInformacionCaratula(S_NumPesta As String)
Dim RE_MiCara As ADODB.Recordset, S_CondiCar As String, s_Cadena As String, N_NumModulo As Integer
Dim S_CamposBasicos As String, S_CamposBasicosDet As String, S_CamposBasicosCheck As String
Dim N_TipoCaraTele As Integer

S_CamposBasicos = "Select c.n_opcion_base,c.s_descrip_base From T_Caratulas c (Nolock) "
S_CamposBasicosDef = "Select c.s_descrip_base,d.n_opcion_base,d.n_opcion_det,d.s_descrip_det From T_CaratulasDet d (Nolock),T_Caratulas c (Nolock) "
S_CamposBasicosCheck = "Select abs(c.n_opcion_base),c.s_descrip_base From T_Caratulas c (Nolock) "

Select Case UCase(Trim(S_NumPesta))
    Case "MODULO1"
        With Frm_Caratula
            N_NumModulo = 1
            S_CondiCar = S_CamposBasicos & _
            "Where c.n_cveproyecto = '" & Trim(.Txt_CarProyecto) & "' And c.n_cveetapa = " & Val(.Lbl_CveEtapa.Caption) & " And c.n_modulo = " & N_NumModulo & " And c.c_tipo_opcion = 'L'"
     
            Set RE_MiCara = New ADODB.Recordset
            RE_MiCara.Open S_CondiCar, gcn
    
            If Val(RE_MiCara.RecordCount) = 0 Then Exit Sub
    
            Do While Not RE_MiCara.EOF()
                s_Cadena = S_Marca & Chr(9) & Trim(RE_MiCara(1)) & Chr(9) & Val(RE_MiCara(0))
                .Grd_Mod1.AddItem s_Cadena
                RE_MiCara.MoveNext
            Loop
            RE_MiCara.Close
            .Grd_Mod1.RemoveItem 1
        End With
    
    Case "MODULO2"
        With Frm_Caratula
            N_NumModulo = 2
            S_CondiCar = S_CamposBasicos & _
            "Where c.n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And c.n_cveetapa = " & Val(.Lbl_CveEtapa.Caption) & " And c.n_modulo = " & N_NumModulo & " And c.c_tipo_opcion = 'L'"
     
            Set RE_MiCara = New ADODB.Recordset
            RE_MiCara.Open S_CondiCar, gcn
    
            If Val(RE_MiCara.RecordCount) = 0 Then Exit Sub
    
            Do While Not RE_MiCara.EOF()
                s_Cadena = S_Marca & Chr(9) & Trim(RE_MiCara(1)) & Chr(9) & Val(RE_MiCara(0))
                .Grd_Mod2.AddItem s_Cadena
                RE_MiCara.MoveNext
            Loop
            RE_MiCara.Close
            .Grd_Mod2.RemoveItem 1
        End With
        
    Case "MODULO3"
        With Frm_Caratula
            N_NumModulo = 3
            S_CondiCar = S_CamposBasicos & _
            "Where c.n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And c.n_cveetapa = " & Val(.Lbl_CveEtapa.Caption) & " And c.n_modulo = " & N_NumModulo & " And c.c_tipo_opcion = 'L'"
     
            Set RE_MiCara = New ADODB.Recordset
            RE_MiCara.Open S_CondiCar, gcn
    
            If Val(RE_MiCara.RecordCount) = 0 Then Exit Sub
    
            Do While Not RE_MiCara.EOF()
                s_Cadena = S_Marca & Chr(9) & Trim(RE_MiCara(1)) & Chr(9) & Val(RE_MiCara(0))
                .Grd_Mod3.AddItem s_Cadena
                RE_MiCara.MoveNext
            Loop
            RE_MiCara.Close
            .Grd_Mod3.RemoveItem 1
        End With
    
    Case "MODULO4"
        With Frm_Caratula
            N_NumModulo = 4
            S_CondiCar = S_CamposBasicos & _
            "Where c.n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And c.n_cveetapa = " & Val(.Lbl_CveEtapa.Caption) & " And c.n_modulo = " & N_NumModulo & " And c.c_tipo_opcion = 'L'"
     
            Set RE_MiCara = New ADODB.Recordset
            RE_MiCara.Open S_CondiCar, gcn
    
            If Val(RE_MiCara.RecordCount) = 0 Then Exit Sub
    
            Do While Not RE_MiCara.EOF()
                s_Cadena = S_Marca & Chr(9) & Trim(RE_MiCara(1)) & Chr(9) & Val(RE_MiCara(0))
                .Grd_Mod4.AddItem s_Cadena
                RE_MiCara.MoveNext
            Loop
            RE_MiCara.Close
            .Grd_Mod4.RemoveItem 1
        End With
        
    Case "MODULO5"
        With Frm_Caratula
            N_NumModulo = 4
            S_CondiCar = S_CamposBasicos & _
            "Where c.n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And c.n_cveetapa = " & Val(.Lbl_CveEtapa.Caption) & " And c.n_modulo = " & N_NumModulo & " And c.c_tipo_opcion = 'L'"
     
            Set RE_MiCara = New ADODB.Recordset
            RE_MiCara.Open S_CondiCar, gcn
    
            If Val(RE_MiCara.RecordCount) = 0 Then Exit Sub
    
            Do While Not RE_MiCara.EOF()
                s_Cadena = S_Marca & Chr(9) & Trim(RE_MiCara(1)) & Chr(9) & Val(RE_MiCara(0))
                .Grd_Mod5CuotasDef.AddItem s_Cadena
                RE_MiCara.MoveNext
            Loop
            RE_MiCara.Close
            .Grd_Mod5CuotasDef.RemoveItem 1
        End With
        
     Case "MODULO5-DEF"
        With Frm_Caratula
            N_NumModulo = 4
            S_CondiCar = S_CamposBasicosDef & _
            "Where d.n_cveproyecto = c.n_cveproyecto And d.n_cveetapa = c.n_cveetapa And d.n_modulo = c.n_modulo And d.n_opcion_base = c.n_opcion_base " & _
            "And d.n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And d.n_cveetapa = " & Val(.Lbl_CveEtapa.Caption) & " And d.n_modulo = " & N_NumModulo
     
            Set RE_MiCara = New ADODB.Recordset
            RE_MiCara.Open S_CondiCar, gcn
    
            If Val(RE_MiCara.RecordCount) = 0 Then Exit Sub
    
            Do While Not RE_MiCara.EOF()
                s_Cadena = S_Marca & Chr(9) & Trim(RE_MiCara(0)) & Chr(9) & Trim(RE_MiCara(3)) & Chr(9) & Val(RE_MiCara(2)) & Chr(9) & Val(RE_MiCara(1))
                .Grd_Mod5.AddItem s_Cadena
                RE_MiCara.MoveNext
            Loop
            RE_MiCara.Close
            .Grd_Mod5.RemoveItem 1
        End With
        
     Case "MODULO6"
        With Frm_Caratula
            N_NumModulo = 6
            S_CondiCar = S_CamposBasicosCheck & _
            "Where c.n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And c.n_cveetapa = " & Val(.Lbl_CveEtapa.Caption) & " And c.n_modulo = " & N_NumModulo & " And c.c_tipo_opcion = 'C' " & _
            "Order by abs(c.n_opcion_base) "
     
            Set RE_MiCara = New ADODB.Recordset
            RE_MiCara.Open S_CondiCar, gcn
    
            If Val(RE_MiCara.RecordCount) = 0 Then Exit Sub
    
            Do While Not RE_MiCara.EOF()
                .Chk_CarMod6(Val(RE_MiCara(0))).Value = 1
                .Chk_CarMod6(Val(RE_MiCara(0))).Tag = 1
                RE_MiCara.MoveNext
            Loop
            RE_MiCara.Close
        End With
        
    Case "MODULO7"
        With Frm_Caratula
            N_NumModulo = 7
            S_CondiCar = S_CamposBasicosCheck & _
            "Where c.n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And c.n_cveetapa = " & Val(.Lbl_CveEtapa.Caption) & " And c.n_modulo = " & N_NumModulo & " And c.c_tipo_opcion = 'C' " & _
            "Order by abs(c.n_opcion_base) "
     
            Set RE_MiCara = New ADODB.Recordset
            RE_MiCara.Open S_CondiCar, gcn
    
            If Val(RE_MiCara.RecordCount) = 0 Then Exit Sub
    
            Do While Not RE_MiCara.EOF()
                .Chk_CarMod7(Val(RE_MiCara(0))).Value = 1
                .Chk_CarMod7(Val(RE_MiCara(0))).Tag = 1
                RE_MiCara.MoveNext
            Loop
            RE_MiCara.Close
        End With
        
    Case "MODULO8"
      
        With Frm_Caratula
            N_NumModulo = 8
            S_CondiCar = S_CamposBasicosCheck & _
            "Where c.n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And c.n_cveetapa = " & Val(.Lbl_CveEtapa.Caption) & " And c.n_modulo = " & N_NumModulo & " And c.c_tipo_opcion = 'C' " & _
            "Order by abs(c.n_opcion_base) "
     
            Set RE_MiCara = New ADODB.Recordset
            RE_MiCara.Open S_CondiCar, gcn
    
            If Val(RE_MiCara.RecordCount) = 0 Then Exit Sub
    
            Do While Not RE_MiCara.EOF()
                .Chk_CarMod8(Val(RE_MiCara(0))).Value = 1
                .Chk_CarMod8(Val(RE_MiCara(0))).Tag = 1
              
                If Val(RE_MiCara(0)) = 0 Then
                    If InStr(1, Trim(RE_MiCara(1)), "4") > 0 Then
                        .Opt_Mod8Apre(0).Value = True
                        .Fra_Mod8_1.Tag = "00"
                    Else
                        .Opt_Mod8Apre(1).Value = True
                        .Fra_Mod8_1.Tag = "01"
                    End If
                End If
                If Val(RE_MiCara(0)) = 1 Then
                    If InStr(1, Trim(RE_MiCara(1)), "4") > 0 Then
                        .Opt_Mod8Amai(0).Value = True
                        .Fra_Mod8_2.Tag = "10"
                    Else
                        .Opt_Mod8Amai(1).Value = True
                        .Fra_Mod8_2.Tag = "11"
                    End If
                End If
                
                RE_MiCara.MoveNext
            Loop
            RE_MiCara.Close
        End With
      
        With Frm_Caratula
            N_NumModulo = 8
            S_CondiCar = S_CamposBasicos & _
            "Where c.n_cveproyecto = '" & Trim(.Txt_CarProyecto) & "' And c.n_cveetapa = " & Val(.Lbl_CveEtapa.Caption) & " And c.n_modulo = " & N_NumModulo & " And c.c_tipo_opcion = 'L'"
     
            Set RE_MiCara = New ADODB.Recordset
            RE_MiCara.Open S_CondiCar, gcn
    
            If Val(RE_MiCara.RecordCount) = 0 Then Exit Sub
    
            Do While Not RE_MiCara.EOF()
                s_Cadena = S_Marca & Chr(9) & Trim(RE_MiCara(1)) & Chr(9) & Val(RE_MiCara(0))
                .Grd_Mod8.AddItem s_Cadena
                RE_MiCara.MoveNext
            Loop
            RE_MiCara.Close
            .Grd_Mod8.RemoveItem 1
        End With
        
    Case "MODULO9"
        SN_TipoCaraTele = FU_TipoCaraTele(Trim(Frm_Caratula.Txt_CarProyecto), Val(Frm_Caratula.Lbl_CveEtapa.Caption))
        If SN_TipoCaraTele = 0 Then
            Frm_Caratula.Opt_CaraTele(0).Value = True
             Call PR_colocaLetrerosModulo9("CARA_CARA")
             Frm_Caratula.Fra_CaraTele.Tag = 0
        Else
            Frm_Caratula.Opt_CaraTele(1).Value = True
            Call PR_colocaLetrerosModulo9("TELEFONICO")
            Frm_Caratula.Fra_CaraTele.Tag = 1
        End If
      
        With Frm_Caratula
            N_NumModulo = 9
            S_CondiCar = S_CamposBasicosCheck & _
            "Where c.n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And c.n_cveetapa = " & Val(.Lbl_CveEtapa.Caption) & " And c.n_modulo = " & N_NumModulo & " And c.c_tipo_opcion = 'C' " & _
            "Order by abs(c.n_opcion_base) "
     
            Set RE_MiCara = New ADODB.Recordset
            RE_MiCara.Open S_CondiCar, gcn
    
            If Val(RE_MiCara.RecordCount) = 0 Then Exit Sub
    
            Do While Not RE_MiCara.EOF()
                .Chk_CarMod9(Val(RE_MiCara(0))).Value = 1
                .Chk_CarMod9(Val(RE_MiCara(0))).Tag = 1
                RE_MiCara.MoveNext
            Loop
            RE_MiCara.Close
        End With
      
        With Frm_Caratula
            N_NumModulo = 9
            S_CondiCar = S_CamposBasicos & _
            "Where c.n_cveproyecto = '" & Trim(.Txt_CarProyecto) & "' And c.n_cveetapa = " & Val(.Lbl_CveEtapa.Caption) & " And c.n_modulo = " & N_NumModulo & " And c.c_tipo_opcion = 'L'"
     
            Set RE_MiCara = New ADODB.Recordset
            RE_MiCara.Open S_CondiCar, gcn
    
            If Val(RE_MiCara.RecordCount) = 0 Then Exit Sub
    
            Do While Not RE_MiCara.EOF()
                s_Cadena = S_Marca & Chr(9) & Trim(RE_MiCara(1)) & Chr(9) & Val(RE_MiCara(0))
                .Grd_Mod9.AddItem s_Cadena
                RE_MiCara.MoveNext
            Loop
            RE_MiCara.Close
            Set RE_MiCara = Nothing
            .Grd_Mod9.RemoveItem 1
        End With
    
End Select
'Coloca información en las pestañas
End Sub
Sub PR_GuardaRegistroIndividual_Pest(N_Modulo, S_DescripRegInd, N_ValorRegInd, N_QuintoValor)
Dim S_ArmaQryEje As String, S_ComBita As String, S_DescripBita As String
Dim N_RegConsecBD As Integer

On Error GoTo ErrAccionGInd

If Val(N_Modulo) = 5 Then
    With Frm_Caratula
        N_ModuloTruco = 4
        
        S_CondiMax = "n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And n_cveetapa = " & Trim(.Lbl_CveEtapa.Caption) & " And n_modulo = " & N_ModuloTruco
        N_RegConsecBD = FU_ConsecutivoX("t_caratulasdet", "s_valorasig", S_CondiMax)

      
        S_ArmaQryEje = "Insert into T_caratulasDet values (" & _
        Trim(.Txt_CarProyecto) & "," & Trim(.Lbl_CveEtapa.Caption) & "," & N_ModuloTruco & "," & _
        Val(N_QuintoValor) & "," & Val(N_ValorRegInd) & ",'" & Trim(S_DescripRegInd) & "'," & N_RegConsecBD & ",0)"
                
        gcn.Execute S_ArmaQryEje
        
        S_ComBita = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo) & "-" & N_QuintoValor & "-" & N_ValorRegInd
        S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (1 Reg)"
        Call PR_GrabaBitacora(6, 1, S_DescripBita, S_ComBita)
    End With
Else
    With Frm_Caratula
        S_CondiMax = "n_cveproyecto = " & Trim(.Txt_CarProyecto) & " And n_cveetapa = " & Trim(.Lbl_CveEtapa.Caption) & " And n_modulo = " & N_Modulo
        N_RegConsecBD = FU_ConsecutivoX("t_caratulas", "s_valorasig", S_CondiMax)
        
        S_ArmaQryEje = "Insert into T_caratulas values (" & _
        Trim(.Txt_CarProyecto) & "," & Trim(.Lbl_CveEtapa.Caption) & "," & N_Modulo & "," & _
        Val(N_ValorRegInd) & ",'" & Trim(S_DescripRegInd) & "'," & N_RegConsecBD & ",'L',0)"
                
        gcn.Execute S_ArmaQryEje
        
        S_ComBita = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo) & "-" & N_ValorRegInd
        S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (1 Reg)"
        Call PR_GrabaBitacora(6, 1, S_DescripBita, S_ComBita)
    End With
End If

Exit Sub

ErrAccionGInd:
If Err.Number <> 0 Then
    MsgBox "Ha ocurrido un error al guardar los datos", 0 + 32, "[PR_GuardaRegistroIndividual_Pest]"
    Exit Sub
End If
End Sub

Sub PR_GrabaCuestionarioBD(S_Mov)
Dim N_Ctl As Integer, N_Mod As Integer, S_Cominserta As String
Dim S_DescripBita As String, S_ComBita As String, S_ActualizaEnc As String, S_TablaTG As String
Dim S_MicondiCta As String, N_CtlsXModulo As Integer, S_ValAsigFormato As String, S_MicondicionBus As String
Dim N_Poside As Integer, S_CvePapa As String

On Error GoTo GuardaCuest
If gs_ProcCuestionario = 1 Then
    S_TablaTG = "t_cuestionarioPrue"
Else
    S_TablaTG = "t_cuestionario"
End If

With Frm_CuestionarioDin
  
    For N_Ctl = 0 To Val(.Txt_CueDinMod1.UBound)
      
        If N_Ctl < 4 Then
            N_Mod = 0
            If N_Ctl < 3 Then
                S_ValAsigFormato = FU_AgregaCeros_Izquierda(2, Trim(.Txt_CueDinMod1(N_Ctl)))
            ElseIf N_Ctl = 3 Or N_Ctl = 4 Then
                S_ValAsigFormato = FU_AgregaCeros_Izquierda(3, Trim(.Txt_CueDinMod1(N_Ctl)))
            End If
        End If
        
        If InStr(1, UCase(Trim(Frm_CuestionarioDin.Lbl_CueDinMod1(N_Ctl).Caption)), "TIPO") > 0 Then
            N_Mod = 1
            S_MicondicionBus = "n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And N_Modulo = " & N_Mod & " And n_opcion_base = " & Trim(.Txt_CueDinMod1(N_Ctl))
            S_ValAsigFormato = FU_RescataInfX_CampoS("s_descrip_base", "T_CARATULAS", S_MicondicionBus)
        End If
        
        If InStr(1, UCase(Trim(Frm_CuestionarioDin.Lbl_CueDinMod1(N_Ctl).Caption)), "CELD") > 0 Then
            N_Mod = 2
            S_MicondicionBus = "n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And N_Modulo = " & N_Mod & " And n_opcion_base = " & Trim(.Txt_CueDinMod1(N_Ctl))
            S_ValAsigFormato = FU_RescataInfX_CampoS("s_descrip_base", "T_CARATULAS", S_MicondicionBus)
        End If
        If InStr(1, UCase(Trim(Frm_CuestionarioDin.Lbl_CueDinMod1(N_Ctl).Caption)), "ROTA") > 0 Then
            N_Mod = 3
            S_MicondicionBus = "n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And N_Modulo = " & N_Mod & " And n_opcion_base = " & Trim(.Txt_CueDinMod1(N_Ctl))
            S_ValAsigFormato = FU_RescataInfX_CampoS("s_descrip_base", "T_CARATULAS", S_MicondicionBus)
        End If
        
        S_Cominserta = "Insert Into " & S_TablaTG & " Values (" & _
        gs_Proyecto & "," & gs_Etapa & "," & Val(.Lbl_CueDinCuesti.Caption) & "," & N_Mod & "," & Val(N_Ctl) & ",'" & Trim(.Lbl_CueDinMod1(N_Ctl).Caption) & "','" & Trim(.Txt_CueDinMod1(N_Ctl)) & "','" & S_ValAsigFormato & "','" & S_Mov & "')"
        
        gcn.Execute S_Cominserta
    Next N_Ctl
    
  
    N_Mod = 4
    S_MicondiCta = "n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And n_modulo = " & N_Mod
    N_CtlsXModulo = FU_Cuenta_Registros("t_caratulas", S_MicondiCta)
    
    If N_CtlsXModulo > 0 Then
        For N_Ctl = 0 To Val(Frm_CuestionarioDin.Txt_CueDinMod45.UBound)
          
            If InStr(1, Trim(.Lbl_CueDinMod45(N_Ctl).Tag), "-") > 0 Then
                N_Poside = InStr(1, Trim(.Lbl_CueDinMod45(N_Ctl).Caption), "-")
                S_CvePapa = Left(Trim(.Lbl_CueDinMod45(N_Ctl).Caption), N_Poside - 1)
                If Len(Trim(.Txt_CueDinMod45(N_Ctl))) > 0 Then
                    S_MicondicionBus = "n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And N_Modulo = " & N_Mod & " And n_opcion_base = " & S_CvePapa & " And n_opcion_det = " & Trim(.Txt_CueDinMod45(N_Ctl))
                    S_ValAsigFormato = FU_RescataInfX_CampoS("s_descrip_det", "T_CARATULASDET", S_MicondicionBus)
                Else
                    S_ValAsigFormato = ""
                End If
            Else
                S_ValAsigFormato = Trim(.Txt_CueDinMod45(N_Ctl))
            End If
          
            S_Cominserta = "Insert Into " & S_TablaTG & " Values (" & _
            gs_Proyecto & "," & gs_Etapa & "," & Val(.Lbl_CueDinCuesti.Caption) & "," & N_Mod & "," & Val(N_Ctl) & ",'" & Trim(.Lbl_CueDinMod45(N_Ctl).Caption) & "','" & Trim(.Txt_CueDinMod45(N_Ctl)) & "','" & S_ValAsigFormato & "','" & S_Mov & "')"
            
            gcn.Execute S_Cominserta
        Next N_Ctl
    End If
    
  
    N_Mod = 6
    S_MicondiCta = "n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And n_modulo = " & N_Mod
    N_CtlsXModulo = FU_Cuenta_Registros("t_caratulas", S_MicondiCta)
    
    If N_CtlsXModulo > 0 Then
        For N_Ctl = 0 To Val(Frm_CuestionarioDin.Txt_CueDinMod6.UBound)
          
            Select Case Left(UCase(Trim(.Lbl_CueDinMod6(N_Ctl).Caption)), 4)
                Case "MUNI"
                    S_ValAsigFormato = FU_AgregaCeros_Izquierda(5, Trim(.Txt_CueDinMod6(N_Ctl)))
                Case "LOCA", "MANZ", "NO.D"
                    S_ValAsigFormato = FU_AgregaCeros_Izquierda(3, Trim(.Txt_CueDinMod6(N_Ctl)))
                Case "AGEB"
                    S_ValAsigFormato = FU_AgregaCeros_IzquierdaCad(6, Trim(.Txt_CueDinMod6(N_Ctl)))
                Case "ORIG"
                    S_MicondicionBus = "n_cvecatgral = " & Trim(.Txt_CueDinMod6(N_Ctl)) & " And n_cvecatgral_p = 4"
                    S_ValAsigFormato = FU_RescataInfX_CampoS("s_descrip", "c_catalogosanexos", S_MicondicionBus)
                Case Else
                    S_ValAsigFormato = Trim(.Txt_CueDinMod6(N_Ctl))
            End Select
          
            S_Cominserta = "Insert Into " & S_TablaTG & " Values (" & _
            gs_Proyecto & "," & gs_Etapa & "," & Val(.Lbl_CueDinCuesti.Caption) & "," & N_Mod & "," & Val(N_Ctl) & ",'" & Trim(.Lbl_CueDinMod6(N_Ctl).Caption) & "','" & Trim(.Txt_CueDinMod6(N_Ctl)) & "','" & S_ValAsigFormato & "','" & S_Mov & "')"
            
            gcn.Execute S_Cominserta
        Next N_Ctl
    End If
    
  
    N_Mod = 7
    S_MicondiCta = "n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And n_modulo = " & N_Mod
    N_CtlsXModulo = FU_Cuenta_Registros("t_caratulas", S_MicondiCta)
    
    If N_CtlsXModulo > 0 Then
        For N_Ctl = 0 To Val(Frm_CuestionarioDin.Txt_CueDinMod7.UBound)
          
            If InStr(1, Trim(.Lbl_CueDinMod7(N_Ctl).Tag), "-") > 0 Then
                Select Case UCase(Trim(.Lbl_CueDinMod7(N_Ctl).Caption))
                    Case "TIPO SUPERVISION GDV"
                        If Len(Trim(.Txt_CueDinMod7(N_Ctl))) > 0 Then
                            S_MicondicionBus = "n_cvecatgral = " & Trim(.Txt_CueDinMod7(N_Ctl)) & " And n_cvecatgral_p = 1"
                            S_ValAsigFormato = FU_RescataInfX_CampoS("s_descrip", "c_catalogosanexos", S_MicondicionBus)
                        Else
                            S_ValAsigFormato = ""
                        End If
                        
                    Case "TIPO SUPERVISION OUTSOURCING"
                        If Len(Trim(.Txt_CueDinMod7(N_Ctl))) > 0 Then
                            S_MicondicionBus = "n_cvecatgral = " & Trim(.Txt_CueDinMod7(N_Ctl)) & " And n_cvecatgral_p = 2"
                            S_ValAsigFormato = FU_RescataInfX_CampoS("s_descrip", "c_catalogosanexos", S_MicondicionBus)
                        Else
                            S_ValAsigFormato = ""
                        End If
                        
                    Case "ENTREVISTA"
                        S_MicondicionBus = "n_cvecatgral = " & Trim(.Txt_CueDinMod7(N_Ctl)) & " And n_cvecatgral_p = 3"
                        S_ValAsigFormato = FU_RescataInfX_CampoS("s_descrip", "c_catalogosanexos", S_MicondicionBus)
                        
                    Case "AUDITOR DE CALIDAD G."
                        If Len(Trim(.Txt_CueDinMod7(N_Ctl))) > 0 Then
                            S_MicondicionBus = "n_cvecatgral = " & Trim(.Txt_CueDinMod7(N_Ctl)) & " And n_cvecatgral_p = 6"
                            S_ValAsigFormato = FU_RescataInfX_CampoS("s_descrip", "c_catalogosanexos", S_MicondicionBus)
                        Else
                            S_ValAsigFormato = ""
                        End If
                        
                    Case "AUDITOR DE CALIDAD O."
                        If Len(Trim(.Txt_CueDinMod7(N_Ctl))) > 0 Then
                            S_MicondicionBus = "n_cvecatgral = " & Trim(.Txt_CueDinMod7(N_Ctl)) & " And n_cvecatgral_p = 7"
                            S_ValAsigFormato = FU_RescataInfX_CampoS("s_descrip", "c_catalogosanexos", S_MicondicionBus)
                        Else
                            S_ValAsigFormato = ""
                        End If
                End Select
            Else
                S_ValAsigFormato = Trim(.Txt_CueDinMod7(N_Ctl))
                If Left(UCase(Trim(.Lbl_CueDinMod7(i).Caption)), 2) = "ID" Then
                    If Len(Trim(.Txt_CueDinMod7(N_Ctl))) > 0 Then
                        S_ValAsigFormato = FU_AgregaCeros_Izquierda(5, Trim(.Txt_CueDinMod7(N_Ctl)))
                    Else
                        S_ValAsigFormato = ""
                    End If
                End If
            End If
          
            S_Cominserta = "Insert Into " & S_TablaTG & " Values (" & _
            gs_Proyecto & "," & gs_Etapa & "," & Val(.Lbl_CueDinCuesti.Caption) & "," & N_Mod & "," & Val(N_Ctl) & ",'" & Trim(.Lbl_CueDinMod7(N_Ctl).Caption) & "','" & Trim(.Txt_CueDinMod7(N_Ctl)) & "','" & S_ValAsigFormato & "','" & S_Mov & "')"
            
            gcn.Execute S_Cominserta
        Next N_Ctl
    End If
    
  
    N_Mod = 8
    S_MicondiCta = "n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And n_modulo = " & N_Mod
    N_CtlsXModulo = FU_Cuenta_Registros("t_caratulas", S_MicondiCta)
    
    If N_CtlsXModulo > 0 Then
        For N_Ctl = 0 To Val(Frm_CuestionarioDin.Txt_CueDinMod8.UBound)
          
            Select Case UCase(Trim(.Lbl_CueDinMod8(N_Ctl).Caption))
                Case "EDAD"
                    S_MicondicionBus = "n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And N_Modulo = " & N_Mod & " And n_opcion_base = " & Trim(.Txt_CueDinMod8(N_Ctl))
                    S_ValAsigFormato = FU_RescataInfX_CampoS("s_descrip_base", "T_CARATULAS", S_MicondicionBus)
                Case "SEXO"
                    S_MicondicionBus = "n_cvecatgral = " & Trim(.Txt_CueDinMod8(N_Ctl)) & " And n_cvecatgral_p = 5"
                    S_ValAsigFormato = FU_RescataInfX_CampoS("s_descrip", "c_catalogosanexos", S_MicondicionBus)
                Case "APRECIATIVO(6 X 4)", "AMAI(6 X 4)"
                    Select Case Val(Trim(.Txt_CueDinMod8(N_Ctl)))
                        Case 1
                            S_ValAsigFormato = "ABC+"
                        Case 2
                            S_ValAsigFormato = "C"
                        Case 3
                            S_ValAsigFormato = "D+"
                        Case 4
                            S_ValAsigFormato = "D/E"
                    End Select
                    
                Case "APRECIATIVO(13 X 6)", "AMAI(13 X 6)"
                    Select Case Val(Trim(.Txt_CueDinMod8(N_Ctl)))
                        Case 1
                            S_ValAsigFormato = "A/B"
                        Case 2
                            S_ValAsigFormato = "C+"
                        Case 3
                            S_ValAsigFormato = "C"
                        Case 4
                            S_ValAsigFormato = "D+"
                        Case 5
                            S_ValAsigFormato = "D"
                        Case 6
                            S_ValAsigFormato = "E"
                    End Select
                Case Else
                    S_ValAsigFormato = Trim(.Txt_CueDinMod8(N_Ctl))
            End Select
                    
            S_Cominserta = "Insert Into " & S_TablaTG & " Values (" & _
            gs_Proyecto & "," & gs_Etapa & "," & Val(.Lbl_CueDinCuesti.Caption) & "," & N_Mod & "," & Val(N_Ctl) & ",'" & Trim(.Lbl_CueDinMod8(N_Ctl).Caption) & "','" & Trim(.Txt_CueDinMod8(N_Ctl)) & "','" & S_ValAsigFormato & "','" & S_Mov & "')"
            
            gcn.Execute S_Cominserta
        Next N_Ctl
    End If
End With

S_ComBita = Trim(gs_Proyecto) & "-" & Trim(gs_Etapa) & "-" & Val(Frm_CuestionarioDin.Lbl_CueDinCuesti.Caption)
If gs_ProcCuestionario = 1 Then
    S_DescripBita = Trim(S_TablaTG) & " - " & S_Mov
    Call PR_GrabaBitacora(8, 1, S_DescripBita, S_ComBita)
ElseIf gs_ProcCuestionario = 2 Then 'Producción
    S_DescripBita = Trim(S_TablaTG) & " - " & S_Mov
    Call PR_GrabaBitacora(8, 1, S_DescripBita, S_ComBita)
ElseIf gs_ProcCuestionario = 3 Then
    S_DescripBita = Trim(S_TablaTG) & " - Actualización"
    Call PR_GrabaBitacora(8, 3, S_DescripBita, S_ComBita)
End If

If gs_ProcCuestionario = 2 Then
  
    S_ActualizaEnc = "Update t_enccuestionario set c_estatusinf = '" & S_Mov & "'" & _
    "Where n_cveproyecto = " & Trim(gs_Proyecto) & " And n_cveetapa = " & Trim(gs_Etapa) & " And s_numcuestionario = '" & Val(Frm_CuestionarioDin.Lbl_CueDinCuesti.Caption) & "'"
    gcn.Execute S_ActualizaEnc
End If
Exit Sub

GuardaCuest:
If Err.Number <> 0 Then
    MsgBox Str(Err.Number) & ": " & Err.Description, 0 + 64, "PR_GrabaCuestionarioBD"
  
    Exit Sub
End If
End Sub
Function FU_ExtraeCadenaGridValoresID(N_Proyecto, N_Etapa) As String
Dim RE_ResumenID As ADODB.Recordset, S_CondiCargaExt As String, S_CadenaValores0 As String, S_CadenaValores1 As String
Dim N_Id As Integer

FU_ExtraeCadenaGridValoresID = ""
S_CadenaValores0 = ""

On Error GoTo CadenaGridID

S_CondiCargaExt = "Select Abs(n_opcion_base),s_descrip_base From t_Caratulas (Nolock) " & _
"Where n_cveproyecto = " & N_Proyecto & " and n_cveetapa = " & N_Etapa & " And n_modulo = 9 And c_tipo_opcion = 'C' " & _
"Order by Abs(n_opcion_base)"

Set RE_ResumenID = New ADODB.Recordset
RE_ResumenID.Open S_CondiCargaExt, gcn

N_Id = 0
If Val(RE_ResumenID.RecordCount) > 0 Then
    Do While Not RE_ResumenID.EOF()
        N_Id = N_Id + 1
         Frm_Resumen.MSFlexGrid1.TextMatrix(1, N_Id) = Val(RE_ResumenID(0))
         RE_ResumenID.MoveNext
    Loop
     RE_ResumenID.Close
     Set RE_ResumenID = Nothing
End If
'---------------------------------------------------------------------------------------------------------
S_CondiCargaExt = "Select Abs(n_opcion_base),s_descrip_base From t_Caratulas (Nolock) " & _
"Where n_cveproyecto = " & N_Proyecto & " and n_cveetapa = " & N_Etapa & " And n_modulo = 9 And c_tipo_opcion = 'L' " & _
"Order by s_valorasig"

Set RE_ResumenID = New ADODB.Recordset
RE_ResumenID.Open S_CondiCargaExt, gcn

If Val(RE_ResumenID.RecordCount) = 0 Then Exit Function
Do While Not RE_ResumenID.EOF()
    N_Id = N_Id + 1
    Frm_Resumen.MSFlexGrid1.TextMatrix(1, N_Id) = Trim(RE_ResumenID(0) + 100)
    RE_ResumenID.MoveNext
Loop
 RE_ResumenID.Close
 Set RE_ResumenID = Nothing
 
Exit Function

CadenaGridID:
If Err.Number <> 0 Then
    MsgBox Str(Err.Number) & ": " & Err.Description, 0 + 64, "FU_ExtraeCadenaGridID"
    FU_ExtraeCadenaGridValoresID = "ERROR"
    Exit Function
End If
'------------------------------------------------------------------------------------------------------------------

End Function
Function FU_ExtraeCadenaGrid(N_Proyecto, N_Etapa) As String
Dim RE_Resumen As ADODB.Recordset, S_CondiCargaExt As String, S_CadenaValores0 As String, S_CadenaValores1 As String
Dim N_PreguntaFiltro As Integer

FU_ExtraeCadenaGrid = ""
S_CadenaValores0 = ""

On Error GoTo CadenaGrid

S_CondiCargaExt = "Select Abs(n_opcion_base),s_descrip_base From t_Caratulas " & _
"Where n_cveproyecto = " & N_Proyecto & " and n_cveetapa = " & N_Etapa & " And n_modulo = 9 And c_tipo_opcion = 'C' " & _
"Order by Abs(n_opcion_base)"

Set RE_Resumen = New ADODB.Recordset
RE_Resumen.Open S_CondiCargaExt, gcn

If Val(RE_Resumen.RecordCount) > 0 Then
    Do While Not RE_Resumen.EOF()
         S_CadenaValores0 = S_CadenaValores0 & Trim(RE_Resumen(1)) & "|"
         RE_Resumen.MoveNext
    Loop
     RE_Resumen.Close
     Set RE_Resumen = Nothing
     
    S_CadenaValores0 = Left(S_CadenaValores0, Len(Trim(S_CadenaValores0)) - 1)
End If
'---------------------------------------------------------------------------------------------------------
S_CadenaValores1 = ""
N_PreguntaFiltro = 0

S_CondiCargaExt = "Select Abs(n_opcion_base),s_descrip_base From t_Caratulas " & _
"Where n_cveproyecto = " & N_Proyecto & " and n_cveetapa = " & N_Etapa & " And n_modulo = 9 And c_tipo_opcion = 'L' " & _
"Order by Abs(n_opcion_base)"

Set RE_Resumen = New ADODB.Recordset
RE_Resumen.Open S_CondiCargaExt, gcn

If Val(RE_Resumen.RecordCount) = 0 Then
    FU_ExtraeCadenaGrid = Trim(S_CadenaValores0)
    Exit Function
End If
Do While Not RE_Resumen.EOF()
    N_PreguntaFiltro = N_PreguntaFiltro + 1
    S_CadenaValores1 = S_CadenaValores1 & Trim(RE_Resumen(1)) & "|"
    RE_Resumen.MoveNext
Loop
 RE_Resumen.Close
 Set RE_Resumen = Nothing
 
S_CadenaValores1 = Left(S_CadenaValores1, Len(Trim(S_CadenaValores1)) - 1)

If Len(Trim(S_CadenaValores0)) > 0 And Len(Trim(S_CadenaValores1)) > 0 Then
    FU_ExtraeCadenaGrid = S_CadenaValores0 & "|" & S_CadenaValores1
ElseIf Len(Trim(S_CadenaValores0)) > 0 Then
     FU_ExtraeCadenaGrid = S_CadenaValores0
ElseIf Len(Trim(S_CadenaValores0)) > 0 Then
    FU_ExtraeCadenaGrid = S_CadenaValores1
End If

Exit Function

CadenaGrid:
If Err.Number <> 0 Then
    MsgBox Str(Err.Number) & ": " & Err.Description, 0 + 64, "FU_ExtraeCadenaGrid"
    Exit Function
End If
'-----------------------------------------------------------------------------------------------------------------------
End Function

Sub PR_CargaInf_Resumen()
Dim RE_MisDatos As ADODB.Recordset, S_CondiCarga As String, N_Valor As Variant
Dim N_Reg As Integer, N_Col As Integer

With Frm_Resumen
    N_Reg = Val(.MSFlexGrid1.Rows - 1)
    N_Col = Val(.MSFlexGrid1.Cols - 1)
  
     S_CondiCarga = "Select r.n_AB From dbo.t_resumen r (Nolock) " & _
     "Where r.n_cveproyecto = " & gs_Proyecto & " And r.n_cveetapa = " & gs_Etapa & " And r.s_numcuestionario = '" & Trim(gs_Cuestionario) & "' " & _
     "Order by r.n_idcolumna"
     
     Set RE_MisDatos = New ADODB.Recordset
     RE_MisDatos.Open S_CondiCarga, gcn
    
     If Val(RE_MisDatos.RecordCount) = 0 Then Exit Sub
     
     li_Col = 0
 
    Do While Not RE_MisDatos.EOF()
        li_Col = li_Col + 1
        .MSFlexGrid1.TextMatrix(2, li_Col) = IIf(IsNull(RE_MisDatos(0)), "", Trim(RE_MisDatos(0)))
      
      
      
      
      
      
        RE_MisDatos.MoveNext
    Loop
 End With
 RE_MisDatos.Close
 Set RE_MisDatos = Nothing
 
End Sub

Function FU_ConsecutivoX(ps_Tabla, ps_Campo, S_Micondicion) As Long
Dim rdo           As ADODB.Recordset
Dim sClave        As Long
Dim sConsec       As String

sConsec = "SELECT Max(" & Trim(ps_Campo) & ") FROM  " & Trim(ps_Tabla) & " Where " & Trim(S_Micondicion)

Set rdo = New ADODB.Recordset
rdo.Open sConsec, gcn

On Error GoTo ErrNvaClave

If Val(rdo.RecordCount) > 0 Then
      FU_ConsecutivoX = IIf(IsNull(rdo(0)), 0, rdo(0)) + 1
Else
      FU_ConsecutivoX = 1
End If
rdo.Close
Exit Function

ErrNvaClave:
    FU_ConsecutivoX = 0
    Exit Function
End Function

Sub PR_GuardaAuditoriaBitacora()
Dim S_CadenaBase As String, S_CadenaInsert  As String, S_CadenaEje As String, S_ValorX As Variant
Dim S_Borrar As String, S_ComBita As String, S_DescripBita As String, li_Row As Integer
Dim S_InsertaAudi As String, S_CadValSolos As String
Dim N_Val1 As Double, N_Val2 As Double, N_Val3 As Double, N_Val4 As Double, S_AuditorClave As String
Dim S_CondiAct As String

On Error GoTo GrabaAuditoriaBit
S_AuditorClave = ""
With Frm_Auditoria
      
    S_Borrar = "Delete From t_bitacoraAud " & _
    "Where n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And s_numcuestionario = " & Trim(gs_Cuestionario)
     gcn.Execute S_Borrar
    
    S_CadenaBase = "Insert into t_bitacoraAud Values (" & gs_Proyecto & "," & gs_Etapa & "," & Trim(gs_Cuestionario) & ","

    S_CadenaInsert = ""
    
    For li_Row = 0 To Val(.Txt_AudMayor.UBound)
        If Trim(.Txt_AudMayor(li_Row)) = "" Then
            S_ValorX = "Null"
        Else
            S_ValorX = Val(Trim(.Txt_AudMayor(li_Row)))
        End If
        S_CadenaInsert = S_CadenaInsert & S_ValorX & ","
    Next
    S_CadenaInsert = Left(S_CadenaInsert, Len(Trim(S_CadenaInsert)) - 1)
  
  
    N_Val1 = IIf(IsNumeric(Trim(.Txt_AudIndices(0))), Val(Trim(.Txt_AudIndices(0))), 0)
    N_Val2 = IIf(IsNumeric(Trim(.Txt_AudIndices(1))), Val(Trim(.Txt_AudIndices(1))), 0)
    N_Val3 = IIf(IsNumeric(Trim(.Txt_AudIndices(2))), Val(Trim(.Txt_AudIndices(2))), 0)
    N_Val4 = IIf(IsNumeric(Trim(.Txt_AudIndices(3))), Val(Trim(.Txt_AudIndices(3))), 0)
    If Len(Trim(.Txt_AudAuditor)) > 0 Then S_AuditorClave = Trim(.Txt_AudAuditor)
    
    S_CadValSolos = Val(.Lbl_AudCveEstatus.Caption) & "," & Val(.Lbl_AudCveTipo.Caption) & ",'" & S_AuditorClave & "','" & Trim(.Txt_AudComentario) & "'," & N_Val1 & "," & N_Val2 & "," & N_Val3 & "," & N_Val4 & ")"
  
    S_CadenaEje = S_CadenaBase & S_CadenaInsert & "," & S_CadValSolos
  
    gcn.Execute S_CadenaEje
    
  
  
    S_CondiAct = "Update t_cuestionario Set s_valorasig = " & Val(.Lbl_AudCveEstatus.Caption) & ", s_valorformato = '" & UCase(Trim(.Cmb_AudEstatus)) & _
    "' Where n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And s_numcuestionario = " & Trim(gs_Cuestionario) & " And n_modulo = 7 And Rtrim(Ltrim(s_descrip_base)) = 'Entrevista'"
    gcn.Execute S_CondiAct
  
    
    S_ComBita = Trim(gs_Proyecto) & "-" & Trim(gs_Etapa) & "-" & Trim(gs_Cuestionario)
    If InStr(1, UCase(Trim(.Lbl_AudTitulo.Caption)), "EDICIÓN") > 0 Then
        MsgBox "Registros actualizados satisfactoriamente.", 0 + 64, "Auditoria de la Bitácora"
        S_DescripBita = "t_bitacoraAud"
        Call PR_GrabaBitacora(11, 3, S_DescripBita, S_ComBita)
      
        S_ActualizaEnc = "Update t_enccuestionario set n_exportar = 2 " & _
        "Where n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And Ltrim(Rtrim(s_numcuestionario)) = '" & Trim(gs_Cuestionario) & "'"
        gcn.Execute S_ActualizaEnc
      
    Else
        MsgBox "Registros guardados satisfactoriamente.", 0 + 64, "Auditoria de la Bitácora"
        S_DescripBita = "t_bitacoraAud"
        Call PR_GrabaBitacora(11, 1, S_DescripBita, S_ComBita)
      
        S_ActualizaEnc = "Update t_enccuestionario set n_exportar = 2 " & _
        "Where n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And Ltrim(Rtrim(s_numcuestionario)) = '" & Trim(gs_Cuestionario) & "'"
        gcn.Execute S_ActualizaEnc
      
    End If
    
End With

Exit Sub

GrabaAuditoriaBit:
If Err.Number <> 0 Then
    MsgBox Str(Err.Number) & ": " & Err.Description, 0 + 64, "PR_GuardaAuditoriaBitacora"
    Exit Sub
End If
End Sub

Function FU_DetectaError_Mod4VsMod5(N_CveProy, N_CveEtapa) As Integer
Dim RE_Caratulas As ADODB.Recordset, S_CondiProy_Etapa As String, Bo_ExisteRen As Boolean

On Error GoTo DetectaVal1

FU_DetectaError_Mod4VsMod5 = 0
S_CondiProy_Etapa = "Select c.n_cveproyecto,c.n_cveetapa,c.n_modulo,c.n_opcion_base From t_caratulas c (Nolock)" & _
"Where n_cveproyecto = " & N_CveProy & " And n_cveetapa = " & N_CveEtapa & " And n_modulo = 4"

Set RE_Caratulas = New ADODB.Recordset
RE_Caratulas.Open S_CondiProy_Etapa, gcn

If RE_Caratulas.RecordCount = 0 Then Exit Function

Do While Not RE_Caratulas.EOF()
    Bo_ExisteRen = FU_DetectaError_Mod4VsMod5Complem(RE_Caratulas(0), RE_Caratulas(1), RE_Caratulas(3))
    If Not Bo_ExisteRen Then
        FU_DetectaError_Mod4VsMod5 = -1
        Exit Do
    End If
    RE_Caratulas.MoveNext
Loop
RE_Caratulas.Close
Set RE_Caratulas = Nothing

Exit Function

DetectaVal1:
If Err.Number <> 0 Then
    FU_DetectaError_Mod4VsMod5 = -1
    MsgBox Str(Err.Number) & ": " & Err.Description, 0 + 64, "FU_DetectaError_Mod4VsMod5"
    Exit Function
End If

End Function
Function FU_DetectaError_Mod4VsMod5Complem(N_ValorProy, N_ValorEta, S_ValorPadre) As Boolean
Dim RE_CaratulasDet As ADODB.Recordset, S_CondiDet As String

On Error GoTo DetectaVal2

FU_DetectaError_Mod4VsMod5Complem = True
S_CondiDet = "Select c.n_cveproyecto,c.n_cveetapa,c.n_modulo,c.n_opcion_base From t_caratulasDet c (Nolock)" & _
" Where n_cveproyecto = " & N_ValorProy & " And n_cveetapa = " & N_ValorEta & " And n_modulo = 4 And n_opcion_base = " & S_ValorPadre

Set RE_CaratulasDet = New ADODB.Recordset
RE_CaratulasDet.Open S_CondiDet, gcn

If RE_CaratulasDet.RecordCount = 0 Then FU_DetectaError_Mod4VsMod5Complem = False
RE_CaratulasDet.Close
Set RE_CaratulasDet = Nothing
Exit Function

DetectaVal2:
If Err.Number <> 0 Then
    FU_DetectaError_Mod4VsMod5Complem = False
    MsgBox Str(Err.Number) & ": " & Err.Description, 0 + 64, "FU_DetectaError_Mod4VsMod5Complem"
    Exit Function
End If
End Function
