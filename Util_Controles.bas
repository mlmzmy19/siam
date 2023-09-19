Attribute VB_Name = "Util_Controles"

Function FU_CreaControles123(N_PosicionIzqCtlsPadres As Integer, N_PosicionTopeCtlsPadres As Integer, N_Totctrls As Integer, M_MiMatriz As Variant) As Integer
Dim i As Integer, N_TopLbl As Integer, N_TopTxt As Integer
Dim N_OlgVer As Integer, N_OlgHor As Integer, N_OlgHorCarril As Integer, N_IzqPadres As Integer
Dim N_CabenHor As Integer, N_Carril As Integer

FU_CreaControles123 = 0
N_OlgVer = 10
N_OlgHor = 160
N_OlgHorCarril = 800   '1000

N_TopLbl = N_PosicionTopeCtlsPadres
N_TopTxt = N_TopLbl + 300 + N_OlgVer    '550
N_IzqPadres = N_PosicionIzqCtlsPadres

N_CabenHor = 9      'Fijo, numero de controles que caben por carril horizontal, por el momento no se usa
If N_Totctrls = 0 Then
    Frm_CuestionarioDin.Lbl_CueDinMod1(0).Caption = "XXXX"
    Exit Function
End If

For i = 1 To N_Totctrls
    Select Case i
        Case 1 To 8
            N_Carril = 1
        Case 9 To 17
            N_Carril = 2
        Case 18 To 26
            N_Carril = 3
        Case 27 To 35
            N_Carril = 4
         Case 36 To 44
            N_Carril = 5
         Case 45 To 53
            N_Carril = 6
         Case 54 To 62
            N_Carril = 7
    End Select
    N_IzqPadres = N_IzqPadres + 1500 + N_OlgHor 'Donde 1500 es el ancho de los cotroles padres
    If N_Carril = 2 And i = 9 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 3 And i = 18 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 4 And i = 27 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 5 And i = 36 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 6 And i = 45 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 7 And i = 54 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    '-------------------------------------------------------------------
    With Frm_CuestionarioDin
        Load .Lbl_CueDinMod1(i)
        Load .Txt_CueDinMod1(i)
        .Lbl_CueDinMod1(i).Caption = M_MiMatriz(i, 0)
        .Lbl_CueDinMod1(i).Tag = M_MiMatriz(i, 1)
        .Lbl_CueDinMod1(i).Top = N_TopLbl
        .Txt_CueDinMod1(i).Top = N_TopTxt
        .Lbl_CueDinMod1(i).Left = N_IzqPadres
        .Txt_CueDinMod1(i).Left = N_IzqPadres
        .Lbl_CueDinMod1(i).Visible = True
        .Txt_CueDinMod1(i).Visible = True
    End With
Next i

FU_CreaControles123 = Frm_CuestionarioDin.Txt_CueDinMod1(Val(Frm_CuestionarioDin.Txt_CueDinMod1.UBound)).Top + 375
End Function

Function FU_CreaControles7(N_PosicionIzqCtlsPadres As Integer, N_PosicionTopeCtlsPadres As Integer, N_Totctrls As Integer) As Integer
Dim i As Integer, N_TopLbl As Integer, N_TopTxt As Integer
Dim N_OlgVer As Integer, N_OlgHor As Integer, N_OlgHorCarril As Integer, N_IzqPadres As Integer
Dim N_CabenHor As Integer, N_Carril As Integer
Dim RE_MisDatos As ADODB.Recordset, S_CondiCarga As String, M_Lista() As Variant, N_Localidad As Integer

FU_CreaControles7 = 0
N_OlgVer = 10
N_OlgHor = 160
N_OlgHorCarril = 800   '1000

N_TopLbl = N_PosicionTopeCtlsPadres
N_TopTxt = N_TopLbl + 300 + N_OlgVer    '550
N_IzqPadres = N_PosicionIzqCtlsPadres

N_CabenHor = 6      'Fijo, numero de controles que caben por carril horizontal, por el momento no se usa

If N_Totctrls = -1 Then
    Frm_CuestionarioDin.Lbl_CueDinMod7(0).Visible = False
    Frm_CuestionarioDin.Txt_CueDinMod7(0).Visible = False
    FU_CreaControles7 = N_PosicionTopeCtlsPadres - 80
    Exit Function
End If
'----------------------CARGA INFORMACIÓN----------------------
 S_CondiCarga = "Select n_opcion_base,s_descrip_base  From t_caratulas " & _
 "Where n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And n_modulo = 7 Order by Abs(n_opcion_base)"
 Set RE_MisDatos = New ADODB.Recordset
 RE_MisDatos.Open S_CondiCarga, gcn

 If Val(RE_MisDatos.RecordCount) = 0 Then Exit Function
 ReDim M_Lista(Val(RE_MisDatos.RecordCount))
 N_Localidad = 0
 
 Do While Not RE_MisDatos.EOF()
     M_Lista(N_Localidad) = Trim(RE_MisDatos(1))
     N_Localidad = N_Localidad + 1
     RE_MisDatos.MoveNext
 Loop
 RE_MisDatos.Close
 Set RE_MisDatos = Nothing
'----------------------------------------------------
Frm_CuestionarioDin.Lbl_CueDinMod7(0).Top = N_TopLbl
Frm_CuestionarioDin.Txt_CueDinMod7(0).Top = N_TopTxt
Frm_CuestionarioDin.Lbl_CueDinMod7(0).Caption = M_Lista(0)
'----------------------------------------------------

For i = 1 To N_Totctrls
    Select Case i
        Case 1 To 5
            N_Carril = 1
        Case 6 To 11
            N_Carril = 2
    End Select
    N_IzqPadres = N_IzqPadres + 1500 + N_OlgHor  'Donde 1500 es el ancho de los cotroles padres
    If N_Carril = 2 And i = 6 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 3 And i = 18 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    '-------------------------------------------------------------------
    With Frm_CuestionarioDin
        Load .Lbl_CueDinMod7(i)
        Load .Txt_CueDinMod7(i)
        .Lbl_CueDinMod7(i).Caption = Trim(M_Lista(i))
        .Lbl_CueDinMod7(i).Tag = ""
        Select Case UCase(Trim(.Lbl_CueDinMod7(i).Caption))
            Case "TIPO SUPERVISION GDV"
                .Lbl_CueDinMod7(i).Tag = FU_ExtraePosiblesValosAEscojerCATALOGOS("SUPERVISION-OUT")
            Case "TIPO SUPERVISION OUTSOURCING"
                .Lbl_CueDinMod7(i).Tag = FU_ExtraePosiblesValosAEscojerCATALOGOS("SUPERVISION-GDV")
            Case "ENTREVISTA"
                .Lbl_CueDinMod7(i).Tag = FU_ExtraePosiblesValosAEscojerCATALOGOS("ENTREVISTA-GDV")
            Case "AUDITOR DE CALIDAD G."
                .Lbl_CueDinMod7(i).Tag = FU_ExtraePosiblesValosAEscojerCATALOGOS("AUDITOR-GDV")
            Case "AUDITOR DE CALIDAD O."
                .Lbl_CueDinMod7(i).Tag = FU_ExtraePosiblesValosAEscojerCATALOGOS("AUDITOR-OUT")
        End Select
        .Lbl_CueDinMod7(i).Top = N_TopLbl
        .Txt_CueDinMod7(i).Top = N_TopTxt
        .Lbl_CueDinMod7(i).Left = N_IzqPadres
        .Txt_CueDinMod7(i).Left = N_IzqPadres
        .Lbl_CueDinMod7(i).Visible = True
        .Txt_CueDinMod7(i).Visible = True
    End With
Next i

FU_CreaControles7 = Frm_CuestionarioDin.Txt_CueDinMod7(Val(Frm_CuestionarioDin.Txt_CueDinMod7.UBound)).Top + 375

End Function

Function FU_CreaControles8(N_PosicionIzqCtlsPadres As Integer, N_PosicionTopeCtlsPadres As Integer, N_Totctrls As Integer) As Integer
Dim i As Integer, N_TopLbl As Integer, N_TopTxt As Integer
Dim N_OlgVer As Integer, N_OlgHor As Integer, N_OlgHorCarril As Integer, N_IzqPadres As Integer
Dim N_CabenHor As Integer, N_Carril As Integer
Dim RE_MisDatos As ADODB.Recordset, S_CondiCarga As String, M_Lista() As Variant, N_Localidad As Integer

FU_CreaControles8 = 0
N_OlgVer = 10
N_OlgHor = 160
N_OlgHorCarril = 800   '1000

N_TopLbl = N_PosicionTopeCtlsPadres
N_TopTxt = N_TopLbl + 300 + N_OlgVer    '550
N_IzqPadres = N_PosicionIzqCtlsPadres

N_CabenHor = 9      'Fijo, numero de controles que caben por carril horizontal, por el momento no se usa

If N_Totctrls = -1 Then
    Frm_CuestionarioDin.Lbl_CueDinMod8(0).Visible = False
    Frm_CuestionarioDin.Txt_CueDinMod8(0).Visible = False
    FU_CreaControles8 = N_PosicionTopeCtlsPadres - 80
    Exit Function
End If
'----------------------CARGA INFORMACIÓN----------------------
 S_CondiCarga = "Select n_opcion_base,s_descrip_base  From t_caratulas (Nolock)" & _
 "Where n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And n_modulo = 8 And c_tipo_opcion = 'C' Order by Abs(n_opcion_base)"
 Set RE_MisDatos = New ADODB.Recordset
 RE_MisDatos.Open S_CondiCarga, gcn

 If Val(RE_MisDatos.RecordCount) = 0 Then Exit Function
 ReDim M_Lista(Val(RE_MisDatos.RecordCount), 1)
 N_Localidad = 0
 
 Do While Not RE_MisDatos.EOF()
    If Abs(Val(RE_MisDatos(0))) = 3 Then        'Edad   antes tenia = 2
        M_Lista(N_Localidad, 0) = Trim(RE_MisDatos(1))
        M_Lista(N_Localidad, 1) = FU_ExtraePosiblesValosAEscojer(Val(RE_MisDatos(0)), 8)
    ElseIf Abs(Val(RE_MisDatos(0))) = 2 Then    'Sexo   antes tenia = 3
        M_Lista(N_Localidad, 0) = Trim(RE_MisDatos(1))
        M_Lista(N_Localidad, 1) = FU_ExtraePosiblesValosAEscojerCATALOGOS("SEXO")
    ElseIf UCase(RE_MisDatos(1)) = "APRECIATIVO(6 X 4)" Then    'Apreciativo
        M_Lista(N_Localidad, 0) = Trim(RE_MisDatos(1))
        M_Lista(N_Localidad, 1) = "-1-2-3-4-"
    ElseIf UCase(RE_MisDatos(1)) = "APRECIATIVO(13 X 6)" Then    'Apreciativo
        M_Lista(N_Localidad, 0) = Trim(RE_MisDatos(1))
        M_Lista(N_Localidad, 1) = "-1-2-3-4-5-6-"
    ElseIf UCase(RE_MisDatos(1)) = "AMAI(6 X 4)" Then    'AMAI
        M_Lista(N_Localidad, 0) = Trim(RE_MisDatos(1))
        M_Lista(N_Localidad, 1) = "-1-2-3-4-"
    ElseIf UCase(RE_MisDatos(1)) = "AMAI(13 X 6)" Then    'AMAI
        M_Lista(N_Localidad, 0) = Trim(RE_MisDatos(1))
        M_Lista(N_Localidad, 1) = "-1-2-3-4-5-6-"
    Else
        M_Lista(N_Localidad, 0) = Trim(RE_MisDatos(1))
        M_Lista(N_Localidad, 1) = ""
    End If
    
     N_Localidad = N_Localidad + 1
     RE_MisDatos.MoveNext
 Loop
 RE_MisDatos.Close
 Set RE_MisDatos = Nothing
'----------------------------------------------------
Frm_CuestionarioDin.Lbl_CueDinMod8(0).Top = N_TopLbl
Frm_CuestionarioDin.Txt_CueDinMod8(0).Top = N_TopTxt
Frm_CuestionarioDin.Lbl_CueDinMod8(0).Caption = M_Lista(0, 0)
Frm_CuestionarioDin.Lbl_CueDinMod8(0).Tag = M_Lista(0, 1)
'----------------------------------------------------

For i = 1 To N_Totctrls
    Select Case i
        Case 1 To 8
            N_Carril = 1
        Case 9 To 17
            N_Carril = 2
        Case 18 To 26
            N_Carril = 3
        Case 27 To 35
            N_Carril = 4
         Case 36 To 44
            N_Carril = 5
         Case 45 To 53
            N_Carril = 6
         Case 54 To 62
            N_Carril = 7
    End Select
    N_IzqPadres = N_IzqPadres + 1500 + N_OlgHor  'Donde 1500 es el ancho de los cotroles padres
    If N_Carril = 2 And i = 9 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 3 And i = 18 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 4 And i = 27 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 5 And i = 36 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 6 And i = 45 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 7 And i = 54 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    '-------------------------------------------------------------------
    With Frm_CuestionarioDin
        Load .Lbl_CueDinMod8(i)
        Load .Txt_CueDinMod8(i)
        .Lbl_CueDinMod8(i).Caption = Trim(M_Lista(i, 0))
        .Lbl_CueDinMod8(i).Tag = Trim(M_Lista(i, 1))
        .Lbl_CueDinMod8(i).Top = N_TopLbl
        .Txt_CueDinMod8(i).Top = N_TopTxt
        .Lbl_CueDinMod8(i).Left = N_IzqPadres
        .Txt_CueDinMod8(i).Left = N_IzqPadres
        .Lbl_CueDinMod8(i).Visible = True
        .Txt_CueDinMod8(i).Visible = True
    End With
Next i

FU_CreaControles8 = Frm_CuestionarioDin.Txt_CueDinMod8(Val(Frm_CuestionarioDin.Txt_CueDinMod8.UBound)).Top + 375

End Function
Function FU_cuentaChecksSeleccionado_Modulo9(S_CualBoton) As Integer
Dim N_SumaCheck As Integer

FU_cuentaChecksSeleccionado_Modulo9 = 0
N_SumaCheck = 0
With Frm_Caratula
    Select Case UCase(Trim(S_CualBoton))
        Case "CARA_CARA"
            If .Chk_CarMod9(0).Value = 1 Then N_SumaCheck = N_SumaCheck + 1
            If .Chk_CarMod9(1).Value = 1 Then N_SumaCheck = N_SumaCheck + 1
            If .Chk_CarMod9(2).Value = 1 Then N_SumaCheck = N_SumaCheck + 1
            If .Chk_CarMod9(3).Value = 1 Then N_SumaCheck = N_SumaCheck + 1
            If .Chk_CarMod9(4).Value = 1 Then N_SumaCheck = N_SumaCheck + 1
            If .Chk_CarMod9(5).Value = 1 Then N_SumaCheck = N_SumaCheck + 1
            If .Chk_CarMod9(8).Value = 1 Then N_SumaCheck = N_SumaCheck + 1
            If .Chk_CarMod9(9).Value = 1 Then N_SumaCheck = N_SumaCheck + 1
        Case "TELEFONICO"
            If .Chk_CarMod9(0).Value = 1 Then N_SumaCheck = N_SumaCheck + 1
            If .Chk_CarMod9(1).Value = 1 Then N_SumaCheck = N_SumaCheck + 1
            If .Chk_CarMod9(2).Value = 1 Then N_SumaCheck = N_SumaCheck + 1
            If .Chk_CarMod9(3).Value = 1 Then N_SumaCheck = N_SumaCheck + 1
            If .Chk_CarMod9(4).Value = 1 Then N_SumaCheck = N_SumaCheck + 1
            If .Chk_CarMod9(5).Value = 1 Then N_SumaCheck = N_SumaCheck + 1
            If .Chk_CarMod9(6).Value = 1 Then N_SumaCheck = N_SumaCheck + 1
            If .Chk_CarMod9(7).Value = 1 Then N_SumaCheck = N_SumaCheck + 1
            If .Chk_CarMod9(8).Value = 1 Then N_SumaCheck = N_SumaCheck + 1
            If .Chk_CarMod9(9).Value = 1 Then N_SumaCheck = N_SumaCheck + 1
    End Select
End With
FU_cuentaChecksSeleccionado_Modulo9 = N_SumaCheck
End Function

Function FU_ValidaDiferenciaHoras_Arra_Ini() As Boolean
Dim S_HoraArr As String, S_HoraIni As String, S_HoraArrValor As String, S_HoraIniValor As String

FU_ValidaDiferenciaHoras_Arra_Ini = False
S_HoraArr = ""
S_HoraIni = ""
With Frm_CuestionarioDin
    For i = 0 To Val(.Txt_CueDinMod6.UBound)
        If Len(Trim(.Txt_CueDinMod6(i))) > 0 Then
            If UCase(Trim(.Lbl_CueDinMod6(i))) = "HORA DE ARRANQUE" Then S_HoraArr = Trim(.Txt_CueDinMod6(i))
            If UCase(Trim(.Lbl_CueDinMod6(i))) = "HORA DE INICIO" Then S_HoraIni = Trim(.Txt_CueDinMod6(i))
        End If
    Next i
End With

If Len(S_HoraArr) > 0 And Len(S_HoraIni) > 0 Then
    N_pos = InStr(1, Trim(S_HoraArr), ":")
    If N_pos <> 3 Then Exit Function
    S_Hora = Left(Trim(S_HoraArr), N_pos - 1)
    S_Min = Right(Trim(S_HoraArr), Len(Trim(S_HoraArr)) - N_pos)
    S_HoraArrValor = Trim(S_Hora) & Trim(S_Min)
    '---
    N_pos = InStr(1, Trim(S_HoraIni), ":")
    If N_pos <> 3 Then Exit Function
    S_Hora = Left(Trim(S_HoraIni), N_pos - 1)
    S_Min = Right(Trim(S_HoraIni), Len(Trim(S_HoraIni)) - N_pos)
    S_HoraIniValor = Trim(S_Hora) & Trim(S_Min)
    
    If Val(S_HoraArrValor) <= Val(S_HoraIniValor) Then
        FU_ValidaDiferenciaHoras_Arra_Ini = True
    Else
        MsgBox "Verificar que la 'Hora Inicio' sea mayor o igual a la 'Hora de Arranque'", 0 + 48, "Verificar Hora Arranque/Inicio"
    End If
Else    'Si alguno de los datos esta en Blanco
    FU_ValidaDiferenciaHoras_Arra_Ini = True
End If

End Function

Function FU_ValidaDiferenciaHoras_Ini_Fin() As Boolean
Dim S_HoraIni As String, S_HoraFin As String, S_HoraIniValor As String, S_HoraFinValor As String

FU_ValidaDiferenciaHoras_Ini_Fin = False
S_HoraIni = ""
S_HoraFin = ""
With Frm_CuestionarioDin
    For i = 0 To Val(.Txt_CueDinMod6.UBound)
        If Len(Trim(.Txt_CueDinMod6(i))) > 0 Then
            If UCase(Trim(.Lbl_CueDinMod6(i))) = "HORA DE INICIO" Then S_HoraIni = Trim(.Txt_CueDinMod6(i))
            If UCase(Trim(.Lbl_CueDinMod6(i))) = "HORA FIN" Then S_HoraFin = Trim(.Txt_CueDinMod6(i))
        End If
    Next i
End With

If Len(S_HoraIni) > 0 And Len(S_HoraFin) > 0 Then
    N_pos = InStr(1, Trim(S_HoraIni), ":")
    If N_pos <> 3 Then Exit Function
    S_Hora = Left(Trim(S_HoraIni), N_pos - 1)
    S_Min = Right(Trim(S_HoraIni), Len(Trim(S_HoraIni)) - N_pos)
    S_HoraIniValor = Trim(S_Hora) & Trim(S_Min)
    '---
    N_pos = InStr(1, Trim(S_HoraFin), ":")
    If N_pos <> 3 Then Exit Function
    S_Hora = Left(Trim(S_HoraFin), N_pos - 1)
    S_Min = Right(Trim(S_HoraFin), Len(Trim(S_HoraFin)) - N_pos)
    S_HoraFinValor = Trim(S_Hora) & Trim(S_Min)
    
    If Val(S_HoraIniValor) < Val(S_HoraFinValor) Then
        FU_ValidaDiferenciaHoras_Ini_Fin = True
    Else
        MsgBox "Verificar que la 'Hora Fin' sea mayor que la 'Hora de Inicio'", 0 + 48, "Verificar Hora Inicio/Fin"
    End If
Else    'Si alguno de los datos esta en Blanco
    FU_ValidaDiferenciaHoras_Ini_Fin = True
End If
End Function


Function FU_ValidaExistenciaCuotasDefinidas() As Integer
Dim N_Reglones As Integer, N_ValorVerificar As Integer

FU_ValidaExistenciaCuotasDefinidas = 0
With Frm_Caratula
    N_Reglones = Val(.Grd_Mod5CuotasDef.Rows - 1)
    If N_Reglones = 1 And Trim(.Grd_Mod5CuotasDef.TextMatrix(1, 1)) = "" And Trim(.Grd_Mod5CuotasDef.TextMatrix(1, 2)) = "" Then Exit Function
    '------------------------------------------------------------------------------------------------------------------
    N_Reg = Val(.Grd_Mod5.Rows - 1)
    If N_Reg = 1 And Trim(.Grd_Mod5.TextMatrix(1, 1)) = "" And Trim(.Grd_Mod5.TextMatrix(1, 2)) = "" Then
        FU_ValidaExistenciaCuotasDefinidas = -1
        Exit Function
    End If
    '------------------------------------------------------------------------------------------------------------------
    For li_Row = 1 To N_Reglones
        N_ValorVerificar = Val(Trim(.Grd_Mod5CuotasDef.TextMatrix(li_Row, 2)))
        If FU_ValidaExistenciaCuotasDefinidasDetalle(N_ValorVerificar) = -1 Then
            FU_ValidaExistenciaCuotasDefinidas = -1
            Exit For
        End If
    Next
End With
'----------------------------------------------------------------------------------------------------------------------
'Sirve para validar que cuando exista alguna cuota, deba tener al menos una opción
'----------------------------------------------------------------------------------------------------------------------
End Function
Function FU_ValidaExistenciaCuotasDefinidasDetalle(N_ClaveX) As Integer
Dim N_Reg As Integer

FU_ValidaExistenciaCuotasDefinidasDetalle = 0
With Frm_Caratula
    N_Reg = Val(.Grd_Mod5.Rows - 1)
    For li_Row = 1 To N_Reg
        If N_ClaveX = Val(Trim(.Grd_Mod5.TextMatrix(li_Row, 4))) Then Exit Function
    Next
End With
FU_ValidaExistenciaCuotasDefinidasDetalle = -1  'Error no esta en la lista
End Function

Sub PR_ColocaMarcaRenglonGrid5()
Dim I_Ren As Integer

With Frm_Caratula
    I_Ren = .Grd_Mod5CuotasDef.RowSel
    .Grd_Mod5CuotasDef.Col = 0
    .Grd_Mod5CuotasDef.TextMatrix(I_Ren, 0) = ">"
    .Lbl_Mod5Opcion.Caption = Trim(.Grd_Mod5CuotasDef.TextMatrix(I_Ren, 1))
    .Lbl_Mod5Opcion.Tag = Val(.Grd_Mod5CuotasDef.TextMatrix(I_Ren, 2))
    
    .Txt_Mod5Opcion = ""
    .Txt_Mod5Valor = ""
End With
End Sub
Function FU_CreaControles45(N_PosicionIzqCtlsPadres As Integer, N_PosicionTopeCtlsPadres As Integer, N_Totctrls As Integer) As Integer
Dim i As Integer, N_TopLbl As Integer, N_TopTxt As Integer
Dim N_OlgVer As Integer, N_OlgHor As Integer, N_OlgHorCarril As Integer, N_IzqPadres As Integer
Dim N_CabenHor As Integer, N_Carril As Integer
Dim RE_MisDatos As ADODB.Recordset, S_CondiCarga As String, M_Lista() As Variant, N_Localidad As Integer
Dim N_UnicoValor As Integer, S_CadValoresPermitidos As String

FU_CreaControles45 = 0
N_OlgVer = 10
N_OlgHor = 160
N_OlgHorCarril = 800   '1000

N_TopLbl = N_PosicionTopeCtlsPadres
N_TopTxt = N_TopLbl + 300 + N_OlgVer    '550
N_IzqPadres = N_PosicionIzqCtlsPadres

N_CabenHor = 9      'Fijo, numero de controles que caben por carril horizontal, por el momento no se usa

If N_Totctrls = -1 Then
    Frm_CuestionarioDin.Lbl_CueDinMod45(0).Visible = False
    Frm_CuestionarioDin.Txt_CueDinMod45(0).Visible = False
    FU_CreaControles45 = N_PosicionTopeCtlsPadres - 80
    Exit Function
End If
'----------------------CARGA INFORMACIÓN----------------------
 S_CondiCarga = "Select n_opcion_base,s_descrip_base  From t_caratulas " & _
 "Where n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And n_modulo = 4"
 Set RE_MisDatos = New ADODB.Recordset
 RE_MisDatos.Open S_CondiCarga, gcn

 N_UnicoValor = Val(RE_MisDatos.RecordCount)     'Cuando es un único valor
 If Val(RE_MisDatos.RecordCount) = 0 Then Exit Function
 
 ReDim M_Lista(Val(RE_MisDatos.RecordCount), 1)
 N_Localidad = 0
 
 Do While Not RE_MisDatos.EOF()
     M_Lista(N_Localidad, 0) = Trim(Str(RE_MisDatos(0))) & "-" & Trim(RE_MisDatos(1))
     M_Lista(N_Localidad, 1) = FU_ExtraePosiblesValosAEscojer(Val(RE_MisDatos(0)), 4)
     N_Localidad = N_Localidad + 1
     RE_MisDatos.MoveNext
 Loop
 RE_MisDatos.Close
 Set RE_MisDatos = Nothing
'----------------------------------------------------
Frm_CuestionarioDin.Lbl_CueDinMod45(0).Top = N_TopLbl
Frm_CuestionarioDin.Txt_CueDinMod45(0).Top = N_TopTxt
Frm_CuestionarioDin.Lbl_CueDinMod45(0).Caption = M_Lista(0, 0)
Frm_CuestionarioDin.Lbl_CueDinMod45(0).Tag = M_Lista(0, 1)
'----------------------------------------------------

For i = 1 To N_Totctrls
    Select Case i
        Case 1 To 8
            N_Carril = 1
        Case 9 To 17
            N_Carril = 2
        Case 18 To 26
            N_Carril = 3
        Case 27 To 35
            N_Carril = 4
         Case 36 To 44
            N_Carril = 5
         Case 45 To 53
            N_Carril = 6
         Case 54 To 62
            N_Carril = 7
    End Select
    N_IzqPadres = N_IzqPadres + 1500 + N_OlgHor  'Donde 1500 es el ancho de los cotroles padres
    If N_Carril = 2 And i = 9 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 3 And i = 18 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 4 And i = 27 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 5 And i = 36 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 6 And i = 45 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 7 And i = 54 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    '-------------------------------------------------------------------
    With Frm_CuestionarioDin
        Load .Lbl_CueDinMod45(i)
        Load .Txt_CueDinMod45(i)
        .Lbl_CueDinMod45(i).Caption = Trim(M_Lista(i, 0))
        .Lbl_CueDinMod45(i).Tag = Trim(M_Lista(i, 1))
        .Lbl_CueDinMod45(i).Top = N_TopLbl
        .Txt_CueDinMod45(i).Top = N_TopTxt
        .Lbl_CueDinMod45(i).Left = N_IzqPadres
        .Txt_CueDinMod45(i).Left = N_IzqPadres
        .Lbl_CueDinMod45(i).Visible = True
        .Txt_CueDinMod45(i).Visible = True
    End With
Next i
'-------Caso especial cuando es un único valor [Documentado 16-Oct-2008]--------------------------------
'For i = 0 To N_Totctrls
'    If Len(Trim(Frm_CuestionarioDin.Lbl_CueDinMod45(i).Tag)) > 0 Then    'Cuando es único valor para la cuota
'        S_CadValoresPermitidos = Trim(Frm_CuestionarioDin.Lbl_CueDinMod45(i).Tag)
'        S_CadValoresPermitidos = Left(Trim(S_CadValoresPermitidos), Len(Trim(S_CadValoresPermitidos)) - 1)
'        S_CadValoresPermitidos = Right(S_CadValoresPermitidos, Len(S_CadValoresPermitidos) - 1)
'        If InStr(1, S_CadValoresPermitidos, "-") = 0 Then Frm_CuestionarioDin.Txt_CueDinMod45(i) = Trim(S_CadValoresPermitidos)
'    End If
'Next i
'-------------------------------------------------------------------------------
FU_CreaControles45 = Frm_CuestionarioDin.Txt_CueDinMod45(Val(Frm_CuestionarioDin.Txt_CueDinMod45.UBound)).Top + 375
End Function
Sub PR_ComboLost(cmb_ctrl As ComboBox, lbl_ctrl As Label)
If Len(Trim(cmb_ctrl)) > 0 Then
    If cmb_ctrl.ListIndex < 0 Then
        lbl_ctrl.Caption = 0
        cmb_ctrl.ForeColor = &HFF&           'Rojo
     Else
        lbl_ctrl.Caption = Val(cmb_ctrl.ItemData(cmb_ctrl.ListIndex))
        cmb_ctrl.ForeColor = &H80000008      'Negro
     End If
 Else
    lbl_ctrl.Caption = 0      'SE IMPLEMENTA CUANDO EL COMBO ESTE VACIO Y POR LO TANTO LA CLAVE ES CERO
 End If
End Sub

Sub PR_CuestationarioAsignaTAG()
With Frm_CuestionarioDin
    'Para el módulo 1,2,3
    For N_Ctl = 0 To Val(.Txt_CueDinMod1.UBound)
        .Txt_CueDinMod1(N_Ctl).Tag = Trim(.Txt_CueDinMod1(N_Ctl))
    Next N_Ctl
    
    'Para el módulo 4-5
    For N_Ctl = 0 To Val(.Txt_CueDinMod45.UBound)
        .Txt_CueDinMod45(N_Ctl).Tag = Trim(.Txt_CueDinMod45(N_Ctl))
    Next N_Ctl
    
    'Para el módulo 6
    For N_Ctl = 0 To Val(.Txt_CueDinMod6.UBound)
        .Txt_CueDinMod6(N_Ctl).Tag = Trim(.Txt_CueDinMod6(N_Ctl))
    Next N_Ctl
    
    'Para el módulo 7
    For N_Ctl = 0 To Val(.Txt_CueDinMod7.UBound)
        .Txt_CueDinMod7(N_Ctl).Tag = Trim(.Txt_CueDinMod7(N_Ctl))
    Next N_Ctl
    
    'Para el módulo 8
    For N_Ctl = 0 To Val(.Txt_CueDinMod8.UBound)
        .Txt_CueDinMod8(N_Ctl).Tag = Trim(.Txt_CueDinMod8(N_Ctl))
    Next N_Ctl
End With
End Sub
Sub PR_IniciazaChecksModulo9()
'NCJ
With Frm_Caratula
    For N_NumCheck = 0 To Val(.Chk_CarMod9.UBound)
        .Chk_CarMod9(N_NumCheck).Value = 1
        .Chk_CarMod9(N_NumCheck).Tag = 1
    Next N_NumCheck
    
    If .Opt_CaraTele(0).Value = True Then
        .Chk_CarMod9(6).Value = 0
        .Chk_CarMod9(7).Value = 0
    End If
End With
PR_LimpiaGridVarios ("Modulo9")
End Sub


Sub PR_InsertaRenglon(I_Renglon As Integer, S_Pestaña As String)

Select Case UCase(Trim(S_Pestaña))
    Case "DEF_TIPO_CONCEPTO"
        'With Frm_ReglaT1
        '   I_Renglon = .Lbl_ReglaT1RenActivo.Caption
        '   If Val(I_Renglon) = 0 Then I_Renglon = 1
        '   .Grd_ReglaT1.TextMatrix(I_Renglon, 1) = .Txt_ReglaT1Descrip
        '   .Grd_ReglaT1.TextMatrix(I_Renglon, 2) = .Lbl_ReglaT1RenActivo.Tag
        'End With
    
    Case "MODULO1"
        With Frm_Caratula
           I_Renglon = .Lbl_Mod1RenActivo.Caption
           If Val(I_Renglon) = 0 Then I_Renglon = 1
           .Grd_Mod1.TextMatrix(I_Renglon, 1) = Trim(.Txt_Mod1TipoCuota)
           .Grd_Mod1.TextMatrix(I_Renglon, 2) = Trim(.Txt_Mod1Valor)
        End With
        
    Case "MODULO2"
        With Frm_Caratula
           I_Renglon = .Lbl_Mod2RenActivo.Caption
           If Val(I_Renglon) = 0 Then I_Renglon = 1
           .Grd_Mod2.TextMatrix(I_Renglon, 1) = Trim(.Txt_Mod2Celdas)
           .Grd_Mod2.TextMatrix(I_Renglon, 2) = Trim(.Txt_Mod2Valor)
        End With
        
    Case "MODULO3"
        With Frm_Caratula
           I_Renglon = .Lbl_Mod3RenActivo.Caption
           If Val(I_Renglon) = 0 Then I_Renglon = 1
           .Grd_Mod3.TextMatrix(I_Renglon, 1) = Trim(.Txt_Mod3Rotacion)
           .Grd_Mod3.TextMatrix(I_Renglon, 2) = Trim(.Txt_Mod3Valor)
        End With
        
    Case "MODULO4"
        With Frm_Caratula
           I_Renglon = .Lbl_Mod4RenActivo.Caption
           If Val(I_Renglon) = 0 Then I_Renglon = 1
           .Grd_Mod4.TextMatrix(I_Renglon, 1) = Trim(.Txt_Mod4Cuotas)
           .Grd_Mod4.TextMatrix(I_Renglon, 2) = Trim(.Txt_Mod4Valor)
        End With
        
    Case "MODULO5"
        With Frm_Caratula
           I_Renglon = .Lbl_Mod5RenActivo.Caption
           If Val(I_Renglon) = 0 Then I_Renglon = 1
           .Grd_Mod5.TextMatrix(I_Renglon, 1) = UCase(Trim(.Lbl_Mod5Opcion.Caption))
           .Grd_Mod5.TextMatrix(I_Renglon, 2) = Trim(.Txt_Mod5Opcion)
           .Grd_Mod5.TextMatrix(I_Renglon, 3) = Trim(.Txt_Mod5Valor)
           .Grd_Mod5.TextMatrix(I_Renglon, 4) = Val(.Lbl_Mod5Opcion.Tag)
        End With
        
    Case "MODULO8"
        With Frm_Caratula
           I_Renglon = .Lbl_Mod8RenActivo.Caption
           If Val(I_Renglon) = 0 Then I_Renglon = 1
           .Grd_Mod8.TextMatrix(I_Renglon, 1) = Trim(.Txt_Mod8OpcionX)
           .Grd_Mod8.TextMatrix(I_Renglon, 2) = Trim(.Txt_Mod8Valor)
        End With
        
    Case "MODULO9"
        With Frm_Caratula
           I_Renglon = .Lbl_Mod9RenActivo.Caption
           If Val(I_Renglon) = 0 Then I_Renglon = 1
           .Grd_Mod9.TextMatrix(I_Renglon, 1) = Trim(.Txt_Mod9PreguntaF)
           .Grd_Mod9.TextMatrix(I_Renglon, 2) = Trim(.Txt_Mod9Valor)
        End With
End Select
'***********************************************************************************************************************
'***********************************************************************************************************************
End Sub

Sub PR_LimpiaCampoValoresDiferentes()
With Frm_CuestionarioDin
    'Para el módulo 1,2,3
    For N_Ctl = 0 To Val(.Txt_CueDinMod1.UBound)
        If .Txt_CueDinMod1(N_Ctl).ForeColor = &HFF& Then    'Rojo
            .Txt_CueDinMod1(N_Ctl) = ""
            .Txt_CueDinMod1(N_Ctl).BackColor = &H80FF&      'Color Anaranjado
            .Txt_CueDinMod1(N_Ctl).ForeColor = &H80000012   'Negro
        End If
    Next N_Ctl
    
    'Para el módulo 4-5
    For N_Ctl = 0 To Val(.Txt_CueDinMod45.UBound)
         If .Txt_CueDinMod45(N_Ctl).ForeColor = &HFF& Then
            .Txt_CueDinMod45(N_Ctl) = ""
            .Txt_CueDinMod45(N_Ctl).BackColor = &H80FF&      'Color Anaranjado
            .Txt_CueDinMod45(N_Ctl).ForeColor = &H80000012   'Negro
        End If
    Next N_Ctl
    
    'Para el módulo 6
    For N_Ctl = 0 To Val(.Txt_CueDinMod6.UBound)
          If .Txt_CueDinMod6(N_Ctl).ForeColor = &HFF& Then
            .Txt_CueDinMod6(N_Ctl) = ""
            .Txt_CueDinMod6(N_Ctl).BackColor = &H80FF&      'Color Anaranjado
            .Txt_CueDinMod6(N_Ctl).ForeColor = &H80000012   'Negro
        End If
    Next N_Ctl
    
    'Para el módulo 7
    For N_Ctl = 0 To Val(.Txt_CueDinMod7.UBound)
          If .Txt_CueDinMod7(N_Ctl).ForeColor = &HFF& Then
            .Txt_CueDinMod7(N_Ctl) = ""
            .Txt_CueDinMod7(N_Ctl).BackColor = &H80FF&      'Color Anaranjado
            .Txt_CueDinMod7(N_Ctl).ForeColor = &H80000012   'Negro
        End If
    Next N_Ctl
    
    'Para el módulo 8
    For N_Ctl = 0 To Val(.Txt_CueDinMod8.UBound)
        If .Txt_CueDinMod8(N_Ctl).ForeColor = &HFF& Then
            .Txt_CueDinMod8(N_Ctl) = ""
            .Txt_CueDinMod8(N_Ctl).BackColor = &H80FF&      'Color Anaranjado
            .Txt_CueDinMod8(N_Ctl).ForeColor = &H80000012   'Negro
        End If
    Next N_Ctl
End With
End Sub

Sub PR_LimpiaControlesPestanas()

PR_LimpiaGridVarios ("Modulo1")
PR_LimpiaGridVarios ("Modulo2")
PR_LimpiaGridVarios ("Modulo3")
PR_LimpiaGridVarios ("Modulo4")
PR_LimpiaGridVarios ("Modulo5")
PR_LimpiaGridVarios ("Modulo5-Def")
PR_LimpiaGridVarios ("Modulo8")
PR_LimpiaGridVarios ("Modulo9")
PR_LimpiaDatosChecksCaratula        'Para las pestañas 6 y 7

End Sub
Sub PR_LimpiaMarcaRenglonGrid5()
Dim N_Reg As Integer

With Frm_Caratula
    N_Reg = Val(.Grd_Mod5CuotasDef.Rows - 1)
    If N_Reg = 1 And .Grd_Mod5CuotasDef.TextMatrix(1, 1) = "" Then Exit Sub
    For li_Row = 1 To N_Reg
        .Grd_Mod5CuotasDef.TextMatrix(li_Row, 0) = ""
    Next
End With
End Sub

Sub PR_LimpiaPantalla(S_Pestaña As String)
Dim N_Ctl As Integer

Select Case UCase(Trim(S_Pestaña))
    Case "DEF_TIPO_CONCEPTO"
        'With Frm_ReglaT1
        '   .Lbl_ReglaT1RenActivo.Caption = 0
        '   .Txt_ReglaT1Descrip = ""
        'End With
    
    Case "MODULO1"
        With Frm_Caratula
           .Lbl_Mod1RenActivo.Caption = 0
           .Txt_Mod1TipoCuota = ""
           .Txt_Mod1Valor = ""
           .Txt_Mod1Valor.Tag = ""
        End With
        
    Case "MODULO2"
        With Frm_Caratula
           .Lbl_Mod2RenActivo.Caption = 0
           .Txt_Mod2Celdas = ""
           .Txt_Mod2Valor = ""
           .Txt_Mod2Valor.Tag = ""
        End With
        
    Case "MODULO3"
        With Frm_Caratula
           .Lbl_Mod3RenActivo.Caption = 0
           .Txt_Mod3Rotacion = ""
           .Txt_Mod3Valor = ""
           .Txt_Mod3Valor.Tag = ""
        End With
        
    Case "MODULO4"
        With Frm_Caratula
           .Lbl_Mod4RenActivo.Caption = 0
           .Txt_Mod4Cuotas = ""
           .Txt_Mod4Valor = ""
           .Txt_Mod4Valor.Tag = ""
        End With
        
    Case "MODULO5"
        With Frm_Caratula
           .Lbl_Mod5RenActivo.Caption = 0
           .Txt_Mod5Opcion = ""
           .Txt_Mod5Valor = ""
           .Txt_Mod5Valor.Tag = ""
        End With
        
    Case "MODULO8"
        With Frm_Caratula
           .Lbl_Mod8RenActivo.Caption = 0
           .Txt_Mod8OpcionX = ""
           .Txt_Mod8Valor = ""
           .Txt_Mod8Valor.Tag = ""
        End With
        
    Case "MODULO9"
        With Frm_Caratula
           .Lbl_Mod9RenActivo.Caption = 0
           .Txt_Mod9PreguntaF = ""
           .Txt_Mod9Valor = ""
        End With
        
    Case "PEC"
        With Frm_PEC
            .Cmb_PECProyecto = ""
            .Cmb_PECEtapa = ""
            .Cmb_PECCuestionario = ""
           .Lbl_CvePECProyecto.Caption = 0
           .Lbl_CvePECEtapa.Caption = 0
           .Lbl_CvePECCuestionario.Caption = 0
        End With
        
    Case "BITACORA/AUDITORIA"
        With Frm_Auditoria
            For N_Ctl = 0 To Val(.Txt_AudMayor.UBound)
                .Txt_AudMayor(N_Ctl) = ""
            Next N_Ctl
            
            For N_Ctl = 0 To Val(.Txt_AudIndices.UBound)
                .Txt_AudIndices(N_Ctl) = ""
            Next N_Ctl
            
            .Cmb_AudEstatus = ""
            .Lbl_AudCveEstatus.Caption = 0
            .Cmb_AudTipo = ""
            .Lbl_AudCveTipo.Caption = 0
            .Txt_AudAuditor = ""
            .Txt_AudComentario = ""
            
            .Txt_AudInvestigador = ""
            .Txt_AudSupervisor = ""
            .Txt_AudSupGDV = ""
            .Txt_AudSupOUT = ""
        End With
    
End Select
    
End Sub
Sub PR_ConfiguraPestañas(S_Pestaña As String)

Select Case UCase(Trim(S_Pestaña))

    Case "MODULO1"
        With Frm_Caratula
            .Grd_Mod1.Cols = 3
            .Grd_Mod1.ColWidth(0) = 300
            .Grd_Mod1.RowHeight(0) = 300
            
            .Grd_Mod1.Col = 1
            .Grd_Mod1.Row = 0
            .Grd_Mod1.ColWidth(1) = 2800
            .Grd_Mod1.Text = "Tipo de Cuota"
            .Grd_Mod1.Font.Bold = True
            .Grd_Mod1.CellAlignment = flexAlignCenterCenter
            
            .Grd_Mod1.Col = 2
            .Grd_Mod1.Row = 0
            .Grd_Mod1.ColWidth(2) = 600
            .Grd_Mod1.Text = "Valor"
            .Grd_Mod1.Font.Bold = True
            .Grd_Mod1.CellAlignment = flexAlignCenterCenter
        End With
        
    Case "MODULO2"
        With Frm_Caratula
            .Grd_Mod2.Cols = 3
            .Grd_Mod2.ColWidth(0) = 300
            .Grd_Mod2.RowHeight(0) = 300
            
            .Grd_Mod2.Col = 1
            .Grd_Mod2.Row = 0
            .Grd_Mod2.ColWidth(1) = 2800
            .Grd_Mod2.Text = "Celdas"
            .Grd_Mod2.Font.Bold = True
            .Grd_Mod2.CellAlignment = flexAlignCenterCenter
            
            .Grd_Mod2.Col = 2
            .Grd_Mod2.Row = 0
            .Grd_Mod2.ColWidth(2) = 600
            .Grd_Mod2.Text = "Valor"
            .Grd_Mod2.Font.Bold = True
            .Grd_Mod2.CellAlignment = flexAlignCenterCenter
        End With
        
    Case "MODULO3"
        With Frm_Caratula
            .Grd_Mod3.Cols = 3
            .Grd_Mod3.ColWidth(0) = 300
            .Grd_Mod3.RowHeight(0) = 300
            
            .Grd_Mod3.Col = 1
            .Grd_Mod3.Row = 0
            .Grd_Mod3.ColWidth(1) = 2800
            .Grd_Mod3.Text = "Rotación"
            .Grd_Mod3.Font.Bold = True
            .Grd_Mod3.CellAlignment = flexAlignCenterCenter
            
            .Grd_Mod3.Col = 2
            .Grd_Mod3.Row = 0
            .Grd_Mod3.ColWidth(2) = 600
            .Grd_Mod3.Text = "Valor"
            .Grd_Mod3.Font.Bold = True
            .Grd_Mod3.CellAlignment = flexAlignCenterCenter
        End With
    
    Case "MODULO4"
        With Frm_Caratula
            .Grd_Mod4.Cols = 3
            .Grd_Mod4.ColWidth(0) = 300
            .Grd_Mod4.RowHeight(0) = 300
            
            .Grd_Mod4.Col = 1
            .Grd_Mod4.Row = 0
            .Grd_Mod4.ColWidth(1) = 3400    '2800
            .Grd_Mod4.Text = "Cuotas"
            .Grd_Mod4.Font.Bold = True
            .Grd_Mod4.CellAlignment = flexAlignCenterCenter
            
            .Grd_Mod4.Col = 2
            .Grd_Mod4.Row = 0
            .Grd_Mod4.ColWidth(2) = 600
            .Grd_Mod4.Text = "Valor"
            .Grd_Mod4.Font.Bold = True
            .Grd_Mod4.CellAlignment = flexAlignCenterCenter
        End With
        
    Case "MODULO5"
        With Frm_Caratula
            .Grd_Mod5.Cols = 5
            .Grd_Mod5.ColWidth(0) = 300
            .Grd_Mod5.RowHeight(0) = 300
            
            .Grd_Mod5.Col = 1
            .Grd_Mod5.Row = 0
            .Grd_Mod5.ColWidth(1) = 2800
            .Grd_Mod5.Text = "Cuotas Definidas"
            .Grd_Mod5.Font.Bold = True
            .Grd_Mod5.CellAlignment = flexAlignCenterCenter
            
             .Grd_Mod5.Col = 2
            .Grd_Mod5.Row = 0
            .Grd_Mod5.ColWidth(2) = 2800
            .Grd_Mod5.Text = "Opción"
            .Grd_Mod5.Font.Bold = True
            .Grd_Mod5.CellAlignment = flexAlignCenterCenter
            
            .Grd_Mod5.Col = 3
            .Grd_Mod5.Row = 0
            .Grd_Mod5.ColWidth(3) = 600
            .Grd_Mod5.Text = "Valor"
            .Grd_Mod5.Font.Bold = True
            .Grd_Mod5.CellAlignment = flexAlignCenterCenter
            
            .Grd_Mod5.Col = 4         'Para el ID de la cuota padre
            .Grd_Mod5.Row = 0
            .Grd_Mod5.ColWidth(4) = 0
        End With
        
    Case "MODULO5_CUOTASDEF"
        With Frm_Caratula
            .Grd_Mod5CuotasDef.Cols = 3
            .Grd_Mod5CuotasDef.ColWidth(0) = 300
            .Grd_Mod5CuotasDef.RowHeight(0) = 300
            
            .Grd_Mod5CuotasDef.Col = 1
            .Grd_Mod5CuotasDef.Row = 0
            .Grd_Mod5CuotasDef.ColWidth(1) = 3120
            .Grd_Mod5CuotasDef.Text = "Cuotas Definidas"
            .Grd_Mod5CuotasDef.Font.Bold = True
            .Grd_Mod5CuotasDef.CellAlignment = flexAlignCenterCenter
            
            .Grd_Mod5CuotasDef.Col = 2         'Para el ID del registro
            .Grd_Mod5CuotasDef.Row = 0
            .Grd_Mod5CuotasDef.ColWidth(2) = 0
        End With
        
    
           
   Case "MODULO8"
        With Frm_Caratula
            .Grd_Mod8.Cols = 3
            .Grd_Mod8.ColWidth(0) = 300
            .Grd_Mod8.RowHeight(0) = 300
            
            .Grd_Mod8.Col = 1
            .Grd_Mod8.Row = 0
            .Grd_Mod8.ColWidth(1) = 2600
            .Grd_Mod8.Text = "Opciones de Edad"
            .Grd_Mod8.Font.Bold = True
            .Grd_Mod8.CellAlignment = flexAlignCenterCenter
            
            .Grd_Mod8.Col = 2
            .Grd_Mod8.Row = 0
            .Grd_Mod8.ColWidth(2) = 600
            .Grd_Mod8.Text = "Valor"
            .Grd_Mod8.Font.Bold = True
            .Grd_Mod8.CellAlignment = flexAlignCenterCenter
        End With
        
    Case "MODULO9"
        With Frm_Caratula
            .Grd_Mod9.Cols = 3
            .Grd_Mod9.ColWidth(0) = 300
            .Grd_Mod9.RowHeight(0) = 300
            
            .Grd_Mod9.Col = 1
            .Grd_Mod9.Row = 0
            .Grd_Mod9.ColWidth(1) = 2800
            .Grd_Mod9.Text = "Pregunta Filtro"
            .Grd_Mod9.Font.Bold = True
            .Grd_Mod9.CellAlignment = flexAlignCenterCenter
            
            .Grd_Mod9.Col = 2
            .Grd_Mod9.Row = 0
            .Grd_Mod9.ColWidth(2) = 600
            .Grd_Mod9.Text = "Valor"
            .Grd_Mod9.Font.Bold = True
            .Grd_Mod9.CellAlignment = flexAlignCenterCenter
        End With
       
    Case "DEF_TIPO_CONCEPTO"
        'With Frm_ReglaT1
        '    .Grd_ReglaT1.Cols = 3
        '    .Grd_ReglaT1.ColWidth(0) = 300
        '    .Grd_ReglaT1.RowHeight(0) = 300
        
        '    .Grd_ReglaT1.Col = 2
        '    .Grd_ReglaT1.Row = 0
        '    .Grd_ReglaT1.ColWidth(1) = 5000
        '    .Grd_ReglaT1.Text = "Descripción de la opción"
        '    .Grd_ReglaT1.Font.Bold = True
        '    .Grd_ReglaT1.CellAlignment = flexAlignCenterCenter
        
        '    .Grd_ReglaT1.Col = 2        'Para el ID del registro
        '    .Grd_ReglaT1.Row = 0
        '    .Grd_ReglaT1.ColWidth(2) = 0
        'End With
        
    Case "PRUEBAS"
End Select

'********************************************************************************************
'*Configura los tamaño y número de la columnas para algún Grid
'********************************************************************************************
End Sub

Public Function FU_Confirma(ls_Mensaje As String, ls_Titulo As String)
    FU_Confirma = (MsgBox(ls_Mensaje, vbQuestion + vbYesNo, ls_Titulo) = vbYes)
End Function
Sub PR_BajaRenglonActualizar(S_Pestaña As String)
Dim I_Renglon As Integer

Select Case UCase(Trim(S_Pestaña))
    Case "DEF_TIPO_CONCEPTO"
        'With Frm_ReglaT1
        '    I_Renglon = .Lbl_ReglaT1RenActivo.Caption
        '    .Grd_ReglaT1.TextMatrix(I_Renglon, 1) = Trim(.Txt_ReglaT1Descrip)
        '    .Grd_ReglaT1.TextMatrix(I_Renglon, 2) = ""
        'End With
        
    Case "MODULO1"
        With Frm_Caratula
            I_Renglon = .Lbl_Mod1RenActivo.Caption
            .Grd_Mod1.TextMatrix(I_Renglon, 1) = Trim(.Txt_Mod1TipoCuota)
            .Grd_Mod1.TextMatrix(I_Renglon, 2) = Trim(.Txt_Mod1Valor)
        End With
        
    Case "OTRO"
    
End Select
'********************************************************************************************
'********************************************************************************************
End Sub

Function FU_ValidacionSencillaTipo() As Boolean
FU_ValidacionSencillaTipo = True

Select Case UCase(Trim(S_Pestaña))
    Case "DEF_TIPO_CONCEPTO"
       If Not IsNumeric(TipoDato) Then
            FU_ValidacionSencillaTipo = False
       ElseIf Not IsDate(TipoDato) Then
            FU_ValidacionSencillaTipo = False
       End If
        
    Case "OTRO"
End Select
'********************************************************************************************
'
'********************************************************************************************
End Function

Sub Buscar(Tv As TreeView, texto As String, sMayMin As Boolean)

Dim nodo As Node
Dim i As Integer
Dim bEncontrado As Boolean

' recorre todos los nodos
    With Tv
        For i = 1 To .Nodes.Count
            ' Diferencia de mayúsculas /Minúsculas
            If sMayMin Then
               If .Nodes(i).Text = texto Then
                  bEncontrado = True
               End If
            End If
            ' No Diferencia de mayúsculas /Minúsculas
            If sMayMin = False Then
               If Trim(LCase(.Nodes(i).Text)) = Trim(LCase(texto)) Then
                  bEncontrado = True
               End If
            End If
            ' si se encontró
            If bEncontrado Then
               ' referencia al nodo actual
               Set nodo = .Nodes(i)
               ' Expanded / para expandir el nodo encontrado
               Do Until nodo.Parent Is Nothing
                  nodo.Parent.Expanded = True
                  Set nodo = nodo.Parent
               Loop
               ' selecciona el item encontrado
               .Nodes.Item(i).Selected = True
               ' Le pone el Foco al treeview
               .SetFocus
               Exit Sub
            End If
        Next
    End With
    
    MsgBox "Dato no encontrado", vbExclamation
'***********************************************************************************************************************
'*Rutina para encontrar un texto en un control TreeView
'***********************************************************************************************************************
End Sub

Sub PR_ColocaMarcaRenglonGrid()
Dim I_Ren As Integer, S_RutaBmp As String

S_RutaBmp = Trim(S_Path) & "\Depura3.bmp"
S_RutaBmpVer = Trim(S_Path) & "\Ball_Green.bmp"
With Frm_Utileria
    I_Ren = .Grd_Utileria.RowSel
    .Grd_Utileria.Col = 0
    If Trim(.Grd_Utileria.TextMatrix(I_Ren, 0)) = "X" Then Exit Sub
    If Val(.Grd_Utileria.CellPicture) = 0 Then           'NO contiene la marca
        Set .Grd_Utileria.CellPicture = LoadPicture(S_RutaBmpVer)
    Else                                                 'SI contiene la marca
        Set .Grd_Utileria.CellPicture = Nothing
    End If
End With
End Sub

Sub PR_LeeOpcionesSeleccionadas()
'Dim N_renglones As Integer

    'N_renglones = Val(Frm_Caratula.Lst_CarMod7.ListCount)
    'If Val(Frm_Caratula.Lst_CarMod7.ListIndex) = -1 Then Exit Sub
     
    'For i = 0 To (N_renglones - 1)
    '     If Frm_Caratula.Lst_CarMod7.Selected(i) = True Then
    '         MsgBox Frm_Caratula.Lst_CarMod7.ItemData(i)
    '     End If
    'Next
'***********************************************************************************************************************
'NO SE USA POR EL MOMENTO
'***********************************************************************************************************************
End Sub

Sub PR_ManejoBotonesListas()
'Dim N_pos As Integer, S_Cadbak As String, N_renglones As Integer, s_cadenota As String

'Select Case Index
'    Case 0  'Agregar
'        If Len(Trim(Frm_Caratula.Txt_CarMod1)) = 0 Then Exit Sub
'        s_cadenota = Trim(Frm_Caratula.Txt_CarMod1)
'        Frm_Caratula.Lst_CarMod1.AddItem s_cadenota
'        Frm_Caratula.Txt_CarMod1 = ""
'        Frm_Caratula.Txt_CarMod1.SetFocus
        
'    Case 1    'Quitar Elemento
'        If Val(Lst_CarMod1.ListIndex) = -1 Then Exit Sub
'        N_pos = Val(Lst_CarMod1.ListIndex)
'        Lst_CarMod1.RemoveItem N_pos
        
'    Case 2  'Subir
'        N_renglones = Val(Lst_CarMod1.ListCount)
'        If Val(Lst_CarMod1.ListIndex) = -1 Then Exit Sub
'        If N_renglones = 1 Then Exit Sub
'        If Lst_CarMod1.Selected(0) = True Then Exit Sub
        
 '       N_pos = Val(Lst_CarMod1.ListIndex)
 '       S_Cadbak = Lst_CarMod1.List(N_pos - 1)
 '       Lst_CarMod1.List(N_pos - 1) = Lst_CarMod1.List(N_pos)
 '       Lst_CarMod1.List(N_pos) = S_Cadbak
 '       Lst_CarMod1.Selected(N_pos - 1) = True
       
 '   Case 3  'Bajar
 '       N_renglones = Val(Lst_CarMod1.ListCount)
 '       If Val(Lst_CarMod1.ListIndex) = -1 Then Exit Sub
 '       If N_renglones = 1 Then Exit Sub
 '       If Lst_CarMod1.Selected(N_renglones - 1) = True Then Exit Sub
        
 '       N_pos = Val(Lst_CarMod1.ListIndex)
 '       S_Cadbak = Lst_CarMod1.List(N_pos + 1)
 '       Lst_CarMod1.List(N_pos + 1) = Lst_CarMod1.List(N_pos)
 '       Lst_CarMod1.List(N_pos) = S_Cadbak
 '       Lst_CarMod1.Selected(N_pos + 1) = True
'End Select

End Sub

Sub PR_MueveFocoControlesIndexados()
Dim N_TotControles As Integer

With Frm_XXX
    N_TotControles = Val(Txt_Test.UBound)
    If N_TotControles = Index Then Exit Sub
    If KeyAscii = 13 Then Txt_Test(Index + 1).SetFocus
End With
'***********************************************************************************************************************
'*Mueve el Foco al siguiente caja de texto indexada
'***********************************************************************************************************************
End Sub

Function FU_vte_EsLetra(pintNum As Integer) As Integer
'******************************************************************************************
'Función                    : gf_EsLetra
'Autor                      : J.M.F.
'Descripción                : Valida si el parametro esta en el rango de asccii de letras mayusculas
'Fecha de Creación          : 26/Enero/1999
'Fecha de Liberación        : 26/Enero/1999
'Fecha de Modificación      :
'Autor de la Modificación   :
'Usuario que solicita la modificación:
'
' Parámetros:   Tipo        Nombre              Descripción
'    Entrada:   Integer      pintNum             Valor númerico
'     salida:   Integer      gfint_vte_EsLetra    pintNum si esta en el rango
'          0:   En caso contrario
'******************************************************************************************
    FU_vte_EsLetra = 0
    If pintNum >= 65 And pintNum <= 90 Or UCase(Chr(pintNum)) = "Ü" Or UCase(Chr(pintNum)) = " " Or UCase(Chr(pintNum)) = "." Then
       FU_vte_EsLetra = pintNum
    ElseIf pintNum >= 97 And pintNum <= 122 Then
       FU_vte_EsLetra = Asc(LCase(Chr(pintNum)))
    ElseIf pintNum = 225 Or pintNum = 233 Or pintNum = 237 Or pintNum = 243 Or pintNum = 250 Then
       FU_vte_EsLetra = Asc(LCase(Chr(pintNum)))
    ElseIf pintNum = 193 Or pintNum = 201 Or pintNum = 205 Or pintNum = 211 Or pintNum = 218 Then
       FU_vte_EsLetra = Asc(UCase(Chr(pintNum)))
    ElseIf pintNum = 46 Or pintNum = 32 Or pintNum = 160 Then
       FU_vte_EsLetra = Asc(LCase(Chr(pintNum)))
    End If
    
End Function
Sub PR_LimpiaDatosChecksCaratula()
Dim N_NumCheck As Integer

    With Frm_Caratula
        For N_NumCheck = 0 To Val(.Chk_CarMod6.UBound)
            .Chk_CarMod6(N_NumCheck).Value = 0
            .Chk_CarMod6(N_NumCheck).Tag = 0
        Next N_NumCheck
        
        For N_NumCheck = 0 To Val(.Chk_CarMod7.UBound)
            .Chk_CarMod7(N_NumCheck).Value = 0
            .Chk_CarMod7(N_NumCheck).Tag = 0
        Next N_NumCheck
        
        For N_NumCheck = 0 To Val(.Chk_CarMod8.UBound)
            .Chk_CarMod8(N_NumCheck).Value = 0
            .Chk_CarMod8(N_NumCheck).Tag = 0
        Next N_NumCheck
        
        For N_NumCheck = 0 To Val(.Chk_CarMod9.UBound)
            .Chk_CarMod9(N_NumCheck).Value = 0
            .Chk_CarMod9(N_NumCheck).Tag = 0
        Next N_NumCheck
    End With
'------------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------------
End Sub
Sub PR_LimpiaDatos(S_Pantalla As String)

Select Case UCase(Trim(S_Pantalla))
    Case "ETAPAS"
        With Frm_Etapas
            .Txt_EtaCveEta = ""
            .Txt_EtaDescrip = ""
            .Txt_EtaDescripLar = ""
            .Lbl_RegistroActualEtapas.Caption = 0
            
            .Cmb_EtaEstatus = ""
            .Lbl_EtaCveEstatus.Caption = 0
            
        End With
        
    Case "PERFILES"
        With Frm_Perfiles
            .Txt_PerCvePer = ""
            .Txt_PerDescrip = ""
            .Txt_PerDescripLar = ""
            .Lbl_RegistroActualPerfiles.Caption = 0
            
            .Cmb_PerEstatus = ""
            .Lbl_PerCveEstatus.Caption = 0
            
        End With
        
    Case "USUARIOS"
        With Frm_Usuarios
            .Txt_UsuClave = ""
            .Txt_UsuNombre = ""
            .Txt_UsuPaterno = ""
            .Txt_UsuMaterno = ""
            .Txt_UsuContra = ""
            .Cmb_UsuPerfil = ""
            .Lbl_RegistroActualUsu.Caption = 0
            .Lbl_UsuCvePerfil.Caption = 0
            
            .Cmb_UsuEstatus = ""
            .Lbl_UsuCveEstatus.Caption = 0
            .Txt_UsuContra.Tag = ""             'Dato de confrimación de la contraseña
            'MMM
            .ComboVarios(0).ListIndex = 0 'Puesto (empleado)
            .txtcampo(0).Text = "" 'Descripción del puesto
            .txtcampo(1).Text = "" 'Alta del Empleado
            .txtcampo(2).Text = "" 'Descripción de la Persona
        End With
        
    Case "PROYECTOS"
        With Frm_Proyecto
            .Txt_ProyNumero = ""
            .Txt_ProyNombre = ""
            .Cmb_ProyEjecutivoSC = ""
            .Txt_ProyNumEjecutivoSC = ""
            .Cmb_ProyEjecutivoPR = ""
            .Txt_ProyNumEjecutivoPR = ""
            .Cmb_ProyProg = ""
            .Txt_ProyNumProg = ""
            .Txt_ProyProd = ""
            .Cmb_ProyEstatus = ""
            .Txt_ProyComen = ""
            .Lbl_ProyCveEstatus.Caption = 0
            .Lbl_RegistroActualProy.Caption = 0
        End With
        
    Case "CARATULA-ENC"
        With Frm_Caratula
            .Txt_CarProyecto = ""
            .Cmb_CarEtapa = ""
            .Lbl_CarCveProyecto.Caption = 0
            .Lbl_CveEtapa.Caption = 0
            .Txt_CarNombreProyecto = ""
            .Cmb_CarEstatus = ""
            .Lbl_CarCveEstatus.Caption = 0
        End With
        
    Case "CARATULA-PESTANAS"
        With Frm_Caratula
            .Lbl_CarProceso.Caption = "NUEVO"
            .Lbl_Mod1RenActivo.Caption = 0
            .Lbl_Mod2RenActivo.Caption = 0
            .Lbl_Mod3RenActivo.Caption = 0
            .Lbl_Mod4RenActivo.Caption = 0
            .Lbl_Mod5RenActivo.Caption = 0
            .Lbl_Mod8RenActivo.Caption = 0
            .Lbl_Mod9RenActivo.Caption = 0
            
            .Lbl_CarEstadoGral.Caption = "X"
        End With
        
    Case "CUESTIONARIO-ENC"
        With Frm_Cuestionario
            .Cmb_CueProyecto = ""
            .Cmb_CueEtapa = ""
            .Txt_CueCuestionario = ""
            .Txt_CueCuestionario2 = ""
            .Lbl_CveCueEtapa.Caption = 0
        End With
        
    Case "PESTANAS-CHECKS"
        With Frm_Caratula
            For i = 0 To Val(.Chk_CarMod6.UBound)
                .Chk_CarMod6(i).Enabled = True
            Next i
            For i = 0 To Val(.Chk_CarMod7.UBound)
                .Chk_CarMod7(i).Enabled = True
            Next i
            For i = 0 To Val(.Chk_CarMod8.UBound)
                .Chk_CarMod8(i).Enabled = True
            Next i
            For i = 0 To Val(.Chk_CarMod9.UBound)
                .Chk_CarMod9(i).Enabled = True
            Next i
            .Opt_CaraTele(0).Value = True
        End With
        
 End Select

End Sub

Public Function gf_Confirma(ls_Mensaje As String, ls_Titulo As String)
    gf_Confirma = (MsgBox(ls_Mensaje, vbQuestion + vbYesNo, ls_Titulo) = vbYes)
End Function

Sub PR_HabilitaNavegadores(S_Navegador As String, Bo_EstadoBoton As Boolean)

Select Case UCase(Trim(S_Navegador))
    Case "ETAPAS"
        With Frm_Etapas
            .ToolbarEtapas.Buttons(10).Enabled = Bo_EstadoBoton
            .ToolbarEtapas.Buttons(11).Enabled = Bo_EstadoBoton
            .ToolbarEtapas.Buttons(21).Enabled = Bo_EstadoBoton
            .ToolbarEtapas.Buttons(22).Enabled = Bo_EstadoBoton
        End With
        
   Case "PERFILES"
        With Frm_Perfiles
            .ToolbarPerfiles.Buttons(10).Enabled = Bo_EstadoBoton
            .ToolbarPerfiles.Buttons(11).Enabled = Bo_EstadoBoton
            .ToolbarPerfiles.Buttons(21).Enabled = Bo_EstadoBoton
            .ToolbarPerfiles.Buttons(22).Enabled = Bo_EstadoBoton
        End With
        
   Case "USUARIOS"
        With Frm_Usuarios
            .ToolbarUsuarios.Buttons(10).Enabled = Bo_EstadoBoton
            .ToolbarUsuarios.Buttons(11).Enabled = Bo_EstadoBoton
            .ToolbarUsuarios.Buttons(21).Enabled = Bo_EstadoBoton
            .ToolbarUsuarios.Buttons(22).Enabled = Bo_EstadoBoton
        End With
        
    Case "PROYECTOS"
          With Frm_Proyecto
            .ToolbarProyectos.Buttons(10).Enabled = Bo_EstadoBoton
            .ToolbarProyectos.Buttons(11).Enabled = Bo_EstadoBoton
            .ToolbarProyectos.Buttons(21).Enabled = Bo_EstadoBoton
            .ToolbarProyectos.Buttons(22).Enabled = Bo_EstadoBoton
        End With
        
    Case "PESTANAS-BOTONES"
        With Frm_Caratula
            .Bot_Mod1(0).Enabled = Bo_EstadoBoton
            .Bot_Mod1(1).Enabled = Bo_EstadoBoton
            .Bot_Mod1(2).Enabled = Bo_EstadoBoton
            
            .Bot_Mod2(0).Enabled = Bo_EstadoBoton
            .Bot_Mod2(1).Enabled = Bo_EstadoBoton
            .Bot_Mod2(2).Enabled = Bo_EstadoBoton
            
            .Bot_Mod3(0).Enabled = Bo_EstadoBoton
            .Bot_Mod3(1).Enabled = Bo_EstadoBoton
            .Bot_Mod3(2).Enabled = Bo_EstadoBoton
            
            .Bot_Mod4(0).Enabled = Bo_EstadoBoton
            .Bot_Mod4(1).Enabled = Bo_EstadoBoton
            .Bot_Mod4(2).Enabled = Bo_EstadoBoton
            
            .Bot_Mod5(0).Enabled = Bo_EstadoBoton
            .Bot_Mod5(1).Enabled = Bo_EstadoBoton
            .Bot_Mod5(2).Enabled = Bo_EstadoBoton
            
            .Bot_Mod8(0).Enabled = Bo_EstadoBoton
            .Bot_Mod8(1).Enabled = Bo_EstadoBoton
            .Bot_Mod8(2).Enabled = Bo_EstadoBoton
            
            .Bot_Mod9(0).Enabled = Bo_EstadoBoton
            .Bot_Mod9(1).Enabled = Bo_EstadoBoton
            .Bot_Mod9(2).Enabled = Bo_EstadoBoton
        End With
        
    Case "PESTANAS-FRAMES"
        With Frm_Caratula
            .Fra_Mod1.Enabled = Bo_EstadoBoton
            .Fra_Mod2.Enabled = Bo_EstadoBoton
            .Fra_Mod3.Enabled = Bo_EstadoBoton
            .Fra_Mod4.Enabled = Bo_EstadoBoton
            .Fra_Mod5.Enabled = Bo_EstadoBoton
            .Fra_Mod6.Enabled = Bo_EstadoBoton
            .Fra_Mod7.Enabled = Bo_EstadoBoton
            .Fra_Mod8.Enabled = Bo_EstadoBoton
            .Fra_Mod9.Enabled = Bo_EstadoBoton
        End With
        
    Case "PESTANAS-CHECKS"
        With Frm_Caratula
            For i = 0 To Val(.Chk_CarMod6.UBound)
                If .Chk_CarMod6(i).Value = 1 Then .Chk_CarMod6(i).Enabled = Bo_EstadoBoton
            Next i
            
            For i = 0 To Val(.Chk_CarMod7.UBound)
                If .Chk_CarMod7(i).Value = 1 Then .Chk_CarMod7(i).Enabled = Bo_EstadoBoton
            Next i
            
            For i = 0 To Val(.Chk_CarMod8.UBound)
                If .Chk_CarMod8(i).Value = 1 Then .Chk_CarMod8(i).Enabled = Bo_EstadoBoton
            Next i
            
            For i = 0 To Val(.Chk_CarMod9.UBound)
                If .Chk_CarMod9(i).Value = 1 Then .Chk_CarMod9(i).Enabled = Bo_EstadoBoton
            Next i
            .Fra_CaraTele.Enabled = Bo_EstadoBoton
        End With
        
    Case "CHECKS"
        With Frm_Caratula
            For i = 0 To Val(.Chk_CarMod6.UBound)
                .Chk_CarMod6(i).Enabled = Bo_EstadoBoton
            Next i
            
            For i = 0 To Val(.Chk_CarMod7.UBound)
                .Chk_CarMod7(i).Enabled = Bo_EstadoBoton
            Next i
            
            For i = 0 To Val(.Chk_CarMod8.UBound)
                .Chk_CarMod8(i).Enabled = Bo_EstadoBoton
            Next i
            
            For i = 0 To Val(.Chk_CarMod9.UBound)
                .Chk_CarMod9(i).Enabled = Bo_EstadoBoton
            Next i
            .Fra_CaraTele.Enabled = Bo_EstadoBoton
        End With
    
End Select

End Sub
Sub PR_LimpiaGridVarios(Cual As String)
Dim I_Ren As Integer, I_Itera As Integer
Dim N_Reg As Integer, N_Col As Integer, li_Col As Integer, li_Row As Integer

Select Case UCase(Trim(Cual))
    Case "MODULO1"
        With Frm_Caratula
            I_Ren = Val(.Grd_Mod1.Rows) - 1
            For I_Itera = I_Ren To 1 Step -1
                If I_Itera >= 2 Then
                    .Grd_Mod1.RemoveItem I_Itera
                Else
                    .Grd_Mod1.TextMatrix(1, 1) = ""
                    .Grd_Mod1.TextMatrix(1, 2) = ""
                End If
            Next I_Itera
        End With
        
    Case "MODULO2"
        With Frm_Caratula
            I_Ren = Val(.Grd_Mod2.Rows) - 1
            For I_Itera = I_Ren To 1 Step -1
                If I_Itera >= 2 Then
                    .Grd_Mod2.RemoveItem I_Itera
                Else
                    .Grd_Mod2.TextMatrix(1, 1) = ""
                    .Grd_Mod2.TextMatrix(1, 2) = ""
                End If
            Next I_Itera
        End With
        
    Case "MODULO3"
        With Frm_Caratula
            I_Ren = Val(.Grd_Mod3.Rows) - 1
            For I_Itera = I_Ren To 1 Step -1
                If I_Itera >= 2 Then
                    .Grd_Mod3.RemoveItem I_Itera
                Else
                    .Grd_Mod3.TextMatrix(1, 1) = ""
                    .Grd_Mod3.TextMatrix(1, 2) = ""
                End If
            Next I_Itera
        End With
        
    Case "MODULO4"
        With Frm_Caratula
            I_Ren = Val(.Grd_Mod4.Rows) - 1
            For I_Itera = I_Ren To 1 Step -1
                If I_Itera >= 2 Then
                    .Grd_Mod4.RemoveItem I_Itera
                Else
                    .Grd_Mod4.TextMatrix(1, 1) = ""
                    .Grd_Mod4.TextMatrix(1, 2) = ""
                End If
            Next I_Itera
        End With
        
    Case "MODULO5-DEF"
        With Frm_Caratula
            I_Ren = Val(.Grd_Mod5CuotasDef.Rows) - 1
            For I_Itera = I_Ren To 1 Step -1
                If I_Itera >= 2 Then
                    .Grd_Mod5CuotasDef.RemoveItem I_Itera
                Else
                    .Grd_Mod5CuotasDef.TextMatrix(1, 1) = ""
                    .Grd_Mod5CuotasDef.TextMatrix(1, 2) = ""
                End If
            Next I_Itera
        End With
        
    Case "MODULO5"
        With Frm_Caratula
            I_Ren = Val(.Grd_Mod5.Rows) - 1
            For I_Itera = I_Ren To 1 Step -1
                If I_Itera >= 2 Then
                    .Grd_Mod5.RemoveItem I_Itera
                Else
                    .Grd_Mod5.TextMatrix(1, 1) = ""
                    .Grd_Mod5.TextMatrix(1, 2) = ""
                    .Grd_Mod5.TextMatrix(1, 3) = ""
                    .Grd_Mod5.TextMatrix(1, 4) = ""
                End If
            Next I_Itera
        End With
        
    Case "MODULO8"
        With Frm_Caratula
            I_Ren = Val(.Grd_Mod8.Rows) - 1
            For I_Itera = I_Ren To 1 Step -1
                If I_Itera >= 2 Then
                    .Grd_Mod8.RemoveItem I_Itera
                Else
                    .Grd_Mod8.TextMatrix(1, 1) = ""
                    .Grd_Mod8.TextMatrix(1, 2) = ""
                End If
            Next I_Itera
        End With
        
    Case "MODULO9"
        With Frm_Caratula
            I_Ren = Val(.Grd_Mod9.Rows) - 1
            For I_Itera = I_Ren To 1 Step -1
                If I_Itera >= 2 Then
                    .Grd_Mod9.RemoveItem I_Itera
                Else
                    .Grd_Mod9.TextMatrix(1, 1) = ""
                    .Grd_Mod9.TextMatrix(1, 2) = ""
                End If
            Next I_Itera
        End With
        
    Case "RESUMEN"
        With Frm_Resumen
            N_Reg = Val(.MSFlexGrid1.Rows - 1)
            N_Col = Val(.MSFlexGrid1.Cols - 1)
            
            For li_Col = 1 To N_Col
                For li_Row = 2 To N_Reg
                    .MSFlexGrid1.TextMatrix(li_Row, li_Col) = ""
                Next
            Next
        End With
        
End Select

End Sub

Function FU_GrabaInformacionChecks(S_CualModulo As String) As Integer
Dim N_CtlsChecks As Integer, i As Integer, S_CondiEje As String, N_Modulo As Integer, N_Reglones As Integer
Dim S_ComBitaLiga As String, S_DescripBita As String
Dim N_TipoCarTel As Integer

FU_GrabaInformacionChecks = 0
On Error GoTo GrabaChecks

Select Case UCase(Trim(S_CualModulo))
    Case "MODULO6"
        N_Reglones = 0
        N_Modulo = 6
        With Frm_Caratula
            For i = 0 To Val(.Chk_CarMod6.UBound)
                If .Chk_CarMod6(i).Value = 1 Then
                    N_Reglones = N_Reglones + 1
                    S_CondiEje = "Insert into t_caratulas values (" & _
                    Trim(.Txt_CarProyecto) & "," & Trim(.Lbl_CveEtapa) & "," & N_Modulo & "," & _
                    IIf(Val(i) > 0, Val(i) * -1, Val(i)) & ",'" & Trim(.Chk_CarMod6(i).Caption) & _
                    "',Null,'C',0)"
                    gcn.Execute S_CondiEje
                End If
            Next i
            
            If N_Reglones > 0 Then
                S_ComBitaLiga = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo)
                S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (" & Trim(N_Reglones) & " Reg)"
                Call PR_GrabaBitacora(6, 1, S_DescripBita, S_ComBitaLiga)
                FU_GrabaInformacionChecks = N_Reglones
            End If
        End With
        
    Case "MODULO7"
        N_Reglones = 0
        N_Modulo = 7
        With Frm_Caratula
            For i = 0 To Val(.Chk_CarMod7.UBound)
                If .Chk_CarMod7(i).Value = 1 Then
                    N_Reglones = N_Reglones + 1
                    S_CondiEje = "Insert into t_caratulas values (" & _
                    Trim(.Txt_CarProyecto) & "," & Trim(.Lbl_CveEtapa) & "," & N_Modulo & "," & _
                    IIf(Val(i) > 0, Val(i) * -1, Val(i)) & ",'" & Trim(.Chk_CarMod7(i).Caption) & _
                    "',Null,'C',0)"
                    gcn.Execute S_CondiEje
                End If
            Next i
            
            If N_Reglones > 0 Then
                S_ComBitaLiga = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo)
                S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (" & Trim(N_Reglones) & " Reg)"
                Call PR_GrabaBitacora(6, 1, S_DescripBita, S_ComBitaLiga)
                FU_GrabaInformacionChecks = N_Reglones
            End If
        End With
    
    Case "MODULO8"
        Dim S_CadOption As String
        
        N_Reglones = 0
        N_Modulo = 8
        With Frm_Caratula
            For i = 0 To Val(.Chk_CarMod8.UBound)
                S_CadOption = ""
                If .Chk_CarMod8(i).Value = 1 Then
                    N_Reglones = N_Reglones + 1
                    If i = 0 Then    'Apreciativo
                        S_CadOption = .Opt_Mod8Apre(0).Caption
                        If .Opt_Mod8Apre(1).Value = True Then S_CadOption = .Opt_Mod8Apre(1).Caption
                    End If
                    If i = 1 Then    'Amai
                        S_CadOption = .Opt_Mod8Amai(0).Caption
                        If .Opt_Mod8Amai(1).Value = True Then S_CadOption = .Opt_Mod8Amai(1).Caption
                    End If
                    S_CondiEje = "Insert into t_caratulas values (" & _
                    Trim(.Txt_CarProyecto) & "," & Trim(.Lbl_CveEtapa) & "," & N_Modulo & "," & _
                    IIf(Val(i) > 0, Val(i) * -1, Val(i)) & ",'" & Trim(.Chk_CarMod8(i).Caption) & S_CadOption & _
                    "',Null,'C',0)"
                    gcn.Execute S_CondiEje
                End If
            Next i
            
            If N_Reglones > 0 Then
                S_ComBitaLiga = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo)
                S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (" & Trim(N_Reglones) & " Reg)"
                Call PR_GrabaBitacora(6, 1, S_DescripBita, S_ComBitaLiga)
                FU_GrabaInformacionChecks = N_Reglones
            End If
        End With
    
    Case "MODULO9"
        N_Reglones = 0
        N_Modulo = 9
        N_TipoCarTel = 0
        If Frm_Caratula.Opt_CaraTele(1).Value = True Then N_TipoCarTel = 1
        With Frm_Caratula
            For i = 0 To Val(.Chk_CarMod9.UBound)
                If .Chk_CarMod9(i).Value = 1 Then
                    N_Reglones = N_Reglones + 1
                    S_CondiEje = "Insert into t_caratulas values (" & _
                    Trim(.Txt_CarProyecto) & "," & Trim(.Lbl_CveEtapa) & "," & N_Modulo & "," & _
                    IIf(Val(i) > 0, Val(i) * -1, Val(i)) & ",'" & Trim(.Chk_CarMod9(i).Caption) & _
                    "',Null,'C'," & N_TipoCarTel & ")"
                    gcn.Execute S_CondiEje
                End If
            Next i
            
            If N_Reglones > 0 Then
                S_ComBitaLiga = Trim(.Txt_CarProyecto) & "-" & Trim(.Lbl_CveEtapa.Caption) & "-" & Trim(N_Modulo)
                S_DescripBita = "Mod " & (Trim(N_Modulo)) & " (" & Trim(N_Reglones) & " Reg)"
                Call PR_GrabaBitacora(6, 1, S_DescripBita, S_ComBitaLiga)
                FU_GrabaInformacionChecks = N_Reglones
            End If
        End With
End Select

Exit Function

GrabaChecks:
If Err.Number <> 0 Then
    MsgBox Str(Err.Number) & ": " & Err.Description, 0 + 64, "FU_GrabaInformacionChecks"
    Exit Function
End If
End Function

Function FU_AsignaConsecutivoGrid(S_Mod As String) As Integer
Dim N_Reglones  As Integer

FU_AsignaConsecutivoGrid = 1
Select Case UCase(Trim(S_Mod))

    Case "MODULO3"
        With Frm_Caratula
            N_Reglones = Val(.Grd_Mod3.Rows)
            If Trim(.Grd_Mod3.TextMatrix(1, 1)) = "" And Trim(.Grd_Mod3.TextMatrix(1, 2)) = "" Then
                Exit Function
            Else
                If N_Reglones = 2 And Trim(.Grd_Mod3.TextMatrix(1, 1)) <> "" Then
                    FU_AsignaConsecutivoGrid = N_Reglones
                Else
                    FU_AsignaConsecutivoGrid = N_Reglones
                End If
            End If
        End With

    Case "MODULO8"
        With Frm_Caratula
            N_Reglones = Val(.Grd_Mod8.Rows)
            If Trim(.Grd_Mod8.TextMatrix(1, 1)) = "" And Trim(.Grd_Mod8.TextMatrix(1, 2)) = "" Then
                Exit Function
            Else
                If N_Reglones = 2 And Trim(.Grd_Mod8.TextMatrix(1, 1)) <> "" Then
                    FU_AsignaConsecutivoGrid = N_Reglones
                Else
                    FU_AsignaConsecutivoGrid = N_Reglones
                End If
            End If
        End With
        
    Case "MODULO9"
        With Frm_Caratula
            N_Reglones = Val(.Grd_Mod9.Rows)
            If Trim(.Grd_Mod9.TextMatrix(1, 1)) = "" And Trim(.Grd_Mod9.TextMatrix(1, 2)) = "" Then
                Exit Function
            Else
                If N_Reglones = 2 And Trim(.Grd_Mod9.TextMatrix(1, 1)) <> "" Then
                    FU_AsignaConsecutivoGrid = N_Reglones
                Else
                    FU_AsignaConsecutivoGrid = N_Reglones
                End If
            End If
        End With
End Select
End Function
Function FU_CreaControles6(N_PosicionIzqCtlsPadres As Integer, N_PosicionTopeCtlsPadres As Integer, N_Totctrls As Integer) As Integer
Dim i As Integer, N_TopLbl As Integer, N_TopTxt As Integer
Dim N_OlgVer As Integer, N_OlgHor As Integer, N_OlgHorCarril As Integer, N_IzqPadres As Integer
Dim N_CabenHor As Integer, N_Carril As Integer
Dim RE_MisDatos As ADODB.Recordset, S_CondiCarga As String, M_Lista() As Variant, N_Localidad As Integer

FU_CreaControles6 = 0
N_OlgVer = 10
N_OlgHor = 160
N_OlgHorCarril = 800   '1000

N_TopLbl = N_PosicionTopeCtlsPadres
N_TopTxt = N_TopLbl + 300 + N_OlgVer    '550
N_IzqPadres = N_PosicionIzqCtlsPadres

N_CabenHor = 9      'Fijo, numero de controles que caben por carril horizontal, por el momento no se usa

If N_Totctrls = -1 Then
    Frm_CuestionarioDin.Lbl_CueDinMod6(0).Visible = False
    Frm_CuestionarioDin.Txt_CueDinMod6(0).Visible = False
    FU_CreaControles6 = N_PosicionTopeCtlsPadres - 80
    Exit Function
End If
'----------------------CARGA INFORMACIÓN----------------------
 S_CondiCarga = "Select n_opcion_base,s_descrip_base  From t_caratulas " & _
 "Where n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And n_modulo = 6 Order by Abs(n_opcion_base)"
 Set RE_MisDatos = New ADODB.Recordset
 RE_MisDatos.Open S_CondiCarga, gcn

 If Val(RE_MisDatos.RecordCount) = 0 Then Exit Function
 ReDim M_Lista(Val(RE_MisDatos.RecordCount))
 N_Localidad = 0
 
 Do While Not RE_MisDatos.EOF()
     M_Lista(N_Localidad) = Trim(RE_MisDatos(1))
     N_Localidad = N_Localidad + 1
     RE_MisDatos.MoveNext
 Loop
 RE_MisDatos.Close
 Set RE_MisDatos = Nothing
'----------------------------------------------------
Frm_CuestionarioDin.Lbl_CueDinMod6(0).Top = N_TopLbl
Frm_CuestionarioDin.Txt_CueDinMod6(0).Top = N_TopTxt
Frm_CuestionarioDin.Lbl_CueDinMod6(0).Caption = M_Lista(0)
'----------------------------------------------------

For i = 1 To N_Totctrls
    Select Case i
        Case 1 To 8
            N_Carril = 1
        Case 9 To 17
            N_Carril = 2
        Case 18 To 26
            N_Carril = 3
        Case 27 To 35
            N_Carril = 4
         Case 36 To 44
            N_Carril = 5
         Case 45 To 53
            N_Carril = 6
         Case 54 To 62
            N_Carril = 7
    End Select
    N_IzqPadres = N_IzqPadres + 1500 + N_OlgHor  'Donde 1500 es el ancho de los cotroles padres
    If N_Carril = 2 And i = 9 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 3 And i = 18 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 4 And i = 27 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 5 And i = 36 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 6 And i = 45 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    If N_Carril = 7 And i = 54 Then
        N_IzqPadres = N_PosicionIzqCtlsPadres
        N_TopLbl = N_TopLbl + N_OlgHorCarril
        N_TopTxt = N_TopLbl + 300 + N_OlgVer
    End If
    '-------------------------------------------------------------------
    With Frm_CuestionarioDin
        Load .Lbl_CueDinMod6(i)
        Load .Txt_CueDinMod6(i)
        .Lbl_CueDinMod6(i).Caption = Trim(M_Lista(i))
        
        .Lbl_CueDinMod6(i).Tag = ""
        Select Case UCase(Trim(.Lbl_CueDinMod6(i).Caption))
            Case "ORIGINAL/SUSTITUTO"
                .Lbl_CueDinMod6(i).Tag = FU_ExtraePosiblesValosAEscojerCATALOGOS("PERSONA-ENT")
            Case "AGEB"
                .Txt_CueDinMod6(i).Width = 700
        End Select
        
        .Lbl_CueDinMod6(i).Top = N_TopLbl
        .Txt_CueDinMod6(i).Top = N_TopTxt
        .Lbl_CueDinMod6(i).Left = N_IzqPadres
        .Txt_CueDinMod6(i).Left = N_IzqPadres
        .Lbl_CueDinMod6(i).Visible = True
        .Txt_CueDinMod6(i).Visible = True
    End With
Next i

FU_CreaControles6 = Frm_CuestionarioDin.Txt_CueDinMod6(Val(Frm_CuestionarioDin.Txt_CueDinMod6.UBound)).Top + 375

End Function

Function FU_cbo_InicializaVarios(Ctl_Control As Control, S_Condi As String)
Dim RE_MiTabla As ADODB.Recordset

Set RE_MiTabla = New ADODB.Recordset
RE_MiTabla.Open S_Condi, gcn

Ctl_Control.Clear
Do While Not RE_MiTabla.EOF()
    Ctl_Control.AddItem Trim(RE_MiTabla(1))
    Ctl_Control.ItemData(Ctl_Control.NewIndex) = Trim(RE_MiTabla(0))
    RE_MiTabla.MoveNext
Loop
RE_MiTabla.Close

End Function
Function FU_ValidaInformacionCuestionarioTAG() As Boolean
FU_ValidaInformacionCuestionarioTAG = True

With Frm_CuestionarioDin
    'Para el módulo 1,2,3
    For N_Ctl = 0 To Val(.Txt_CueDinMod1.UBound)
        If Trim(.Txt_CueDinMod1(N_Ctl)) <> Trim(.Txt_CueDinMod1(N_Ctl).Tag) Then
            .Txt_CueDinMod1(N_Ctl).ForeColor = &HFF&       'Rojo
            FU_ValidaInformacionCuestionarioTAG = False
        Else
            .Txt_CueDinMod1(N_Ctl).ForeColor = &H80000012  'Negro
        End If
    Next N_Ctl
    
    'Para el módulo 4-5
    For N_Ctl = 0 To Val(.Txt_CueDinMod45.UBound)
         If Trim(.Txt_CueDinMod45(N_Ctl)) <> Trim(.Txt_CueDinMod45(N_Ctl).Tag) Then
            .Txt_CueDinMod45(N_Ctl).ForeColor = &HFF&       'Rojo
            FU_ValidaInformacionCuestionarioTAG = False
        Else
            .Txt_CueDinMod45(N_Ctl).ForeColor = &H80000012  'Negro
        End If
    Next N_Ctl
    
    'Para el módulo 6
    For N_Ctl = 0 To Val(.Txt_CueDinMod6.UBound)
          If Trim(.Txt_CueDinMod6(N_Ctl)) <> Trim(.Txt_CueDinMod6(N_Ctl).Tag) Then
            .Txt_CueDinMod6(N_Ctl).ForeColor = &HFF&       'Rojo
            FU_ValidaInformacionCuestionarioTAG = False
        Else
            .Txt_CueDinMod6(N_Ctl).ForeColor = &H80000012  'Negro
        End If
    Next N_Ctl
    
    'Para el módulo 7
    For N_Ctl = 0 To Val(.Txt_CueDinMod7.UBound)
          If Trim(.Txt_CueDinMod7(N_Ctl)) <> Trim(.Txt_CueDinMod7(N_Ctl).Tag) Then
            .Txt_CueDinMod7(N_Ctl).ForeColor = &HFF&       'Rojo
            FU_ValidaInformacionCuestionarioTAG = False
        Else
            .Txt_CueDinMod7(N_Ctl).ForeColor = &H80000012  'Negro
        End If
    Next N_Ctl
    
    'Para el módulo 8
    For N_Ctl = 0 To Val(.Txt_CueDinMod8.UBound)
        If Trim(.Txt_CueDinMod8(N_Ctl)) <> Trim(.Txt_CueDinMod8(N_Ctl).Tag) Then
            .Txt_CueDinMod8(N_Ctl).ForeColor = &HFF&       'Rojo
            FU_ValidaInformacionCuestionarioTAG = False
        Else
            .Txt_CueDinMod8(N_Ctl).ForeColor = &H80000012  'Negro
        End If
    Next N_Ctl
End With
End Function
Function FU_ValidaInformacionCuestionario() As Boolean
Dim N_Ctl As Integer, N_Mod As Integer, N_TamanoCampo As Integer
Dim S_CadValoresPermitidos As String, N_ValorUnico As Integer   '30-Sep-08

FU_ValidaInformacionCuestionario = True
N_Mod = 0
With Frm_CuestionarioDin
    'Para el módulo 1,2,3
    For N_Ctl = 0 To Val(.Txt_CueDinMod1.UBound)
        If Len(Trim(.Txt_CueDinMod1(N_Ctl))) = 0 Or .Txt_CueDinMod1(N_Ctl).BackColor = &H80FF& Then  'Color Anaranjado
            'MsgBox "Verificar datos capturados del Tipo de Cuota, Celda o Rotación.", 0 + 16, "Validación de Captura"
            MsgBox "Verificar datos capturados en el campo: " & Trim(.Lbl_CueDinMod1(N_Ctl).Caption), 0 + 16, "Validación de Captura"
            FU_ValidaInformacionCuestionario = False
            Exit Function
        End If
    Next N_Ctl
    
    If FU_ValidaDiaFebrero < 0 Then         'Validación incorporada 02-Mar-09
        FU_ValidaInformacionCuestionario = False
        Exit Function
    End If
    
    I_MesCap = Val(Trim(Frm_CuestionarioDin.Txt_CueDinMod1(1)))
    If FU_ValidaDiaMeses_con30dias(I_MesCap) < 0 Then  'Validación incorporada 20-Mar-09
        FU_ValidaInformacionCuestionario = False
        Exit Function
    End If
    
    'Para el módulo 4-5
    N_Mod = 4
    For N_Ctl = 0 To Val(.Txt_CueDinMod45.UBound)
        If Val(.Txt_CueDinMod45.UBound) = 0 And .Txt_CueDinMod45(0).Visible = False Then Exit For
        '-->30-Sep-08
        N_ValorUnico = 0
        S_CadValoresPermitidos = Trim(Frm_CuestionarioDin.Lbl_CueDinMod45(N_Ctl).Tag)
        S_CadValoresPermitidos = Left(Trim(S_CadValoresPermitidos), Len(Trim(S_CadValoresPermitidos)) - 1)
        S_CadValoresPermitidos = Right(S_CadValoresPermitidos, Len(S_CadValoresPermitidos) - 1)
        If InStr(1, S_CadValoresPermitidos, "-") = 0 Then N_ValorUnico = -1
        '<--
        If N_ValorUnico = 0 Then
            If Len(Trim(.Txt_CueDinMod45(N_Ctl))) = 0 Or .Txt_CueDinMod45(N_Ctl).BackColor = &H80FF& Then  'Color Anaranjado
                MsgBox "Verificar datos capturados en las Cuotas Definidas", 0 + 16, "Validación de Captura"
                FU_ValidaInformacionCuestionario = False
                Exit Function
            End If
        End If
    Next N_Ctl
    
    'Para el módulo 6
    N_Mod = 6
    For N_Ctl = 0 To Val(.Txt_CueDinMod6.UBound)
        If Val(.Txt_CueDinMod6.UBound) = 0 And .Txt_CueDinMod6(0).Visible = False Then Exit For
        If Len(Trim(.Txt_CueDinMod6(N_Ctl))) = 0 Or .Txt_CueDinMod6(N_Ctl).BackColor = &H80FF& Then  'Color Anaranjado
            'MsgBox "Verificar datos capturados al módulo 6", 0 + 16, "Validación de Captura"
            MsgBox "Verificar datos capturados en el campo: " & Trim(.Lbl_CueDinMod6(N_Ctl).Caption), 0 + 16, "Validación de Captura"
            FU_ValidaInformacionCuestionario = False
            Exit Function
        End If
    Next N_Ctl
    
    'Para el módulo 7
    N_Mod = 7
    For N_Ctl = 0 To Val(.Txt_CueDinMod7.UBound)
        If Val(.Txt_CueDinMod7.UBound) = 0 And .Txt_CueDinMod7(0).Visible = False Then Exit For
        'Truco cuando no tenga valor Tipos de Supervisión y Tipos de Auditor
        '-->
        N_TamanoCampo = Len(Trim(.Txt_CueDinMod7(N_Ctl)))
        Select Case UCase(Trim(.Lbl_CueDinMod7(N_Ctl).Caption))
            Case "TIPO SUPERVISION GDV"
                If N_TamanoCampo = 0 Then N_TamanoCampo = 1000
            Case "TIPO SUPERVISION OUTSOURCING"
                If N_TamanoCampo = 0 Then N_TamanoCampo = 1000
            Case "AUDITOR DE CALIDAD G."
                If N_TamanoCampo = 0 Then N_TamanoCampo = 1000
            Case "AUDITOR DE CALIDAD O."
                If N_TamanoCampo = 0 Then N_TamanoCampo = 1000
                
            Case "ID SUPERVISOR OUT", "ID COORDINADOR", "ID CODIFICADOR"
                If N_TamanoCampo = 0 Then N_TamanoCampo = 1000
        End Select
        '<--
        If N_TamanoCampo = 0 Or .Txt_CueDinMod7(N_Ctl).BackColor = &H80FF& Then  'Color Anaranjado
            'MsgBox "Verificar datos capturados al módulo 7", 0 + 16, "Validación de Captura"
            MsgBox "Verificar datos capturados en el campo: " & Trim(.Lbl_CueDinMod7(N_Ctl).Caption), 0 + 16, "Validación de Captura"
            FU_ValidaInformacionCuestionario = False
            Exit Function
        End If
    Next N_Ctl
    
    'Para el módulo 8
    N_Mod = 8
    For N_Ctl = 0 To Val(.Txt_CueDinMod8.UBound)
        If Val(.Txt_CueDinMod8.UBound) = 0 And .Txt_CueDinMod8(0).Visible = False Then Exit For
        If Len(Trim(.Txt_CueDinMod8(N_Ctl))) = 0 Or .Txt_CueDinMod8(N_Ctl).BackColor = &H80FF& Then  'Color Anaranjado
            'MsgBox "Verificar datos capturados al módulo 8", 0 + 16, "Validación de Captura"
            MsgBox "Verificar datos capturados en el campo: " & Trim(.Lbl_CueDinMod8(N_Ctl).Caption), 0 + 16, "Validación de Captura"
            FU_ValidaInformacionCuestionario = False
            Exit Function
        End If
    Next N_Ctl
    
    If Not FU_ValidaDiferenciaHoras_Ini_Fin Then
        FU_ValidaInformacionCuestionario = False
        Exit Function
    End If
    
    If Not FU_ValidaDiferenciaHoras_Arra_Ini Then
        FU_ValidaInformacionCuestionario = False
        Exit Function
    End If
    
    
End With
End Function

Sub PR_CreaGrid_Resumen()
Dim S_MiCadenota As String, S_ini As String, S_s As String, N_RegModulo9 As Integer, N_ColI As Integer
Dim S_Micondicion As String

'-----------------------------------------------
S_Micondicion = "n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And n_modulo = 9"
N_RegModulo9 = FU_Cuenta_Registros("t_caratulas", S_Micondicion)
If N_RegModulo9 = 0 Then Exit Sub
'------------------------------------------------
S_Cadconfig = "*|"
S_MiCadenota = FU_ExtraeCadenaGrid(gs_Proyecto, gs_Etapa)

S_Cadconfig = S_Cadconfig & S_MiCadenota
With Frm_Resumen.MSFlexGrid1
        ' añadir las columnas con FormatString
        .RowHeight(0) = 400
        .RowHeight(1) = 0
        .FormatString = S_Cadconfig
        .FixedCols = 1
  
        ' Propiedades para el MsFlexgrid
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        
        '-------------Identificador de la Columna
        .TextMatrix(2, 0) = "Total"
        '.AddItem " "
        '--------------[DOCUMENTADO POR LA MINUTA DEL DÍA 15/OCT/08]-------------------------------
        '.TextMatrix(3, 0) = "C+"
        '.AddItem " "
        '.TextMatrix(4, 0) = "C"
        '.AddItem " "
        '.TextMatrix(5, 0) = "D+"
        '.AddItem " "
        '.TextMatrix(6, 0) = "D"
        '.AddItem " "
        '.TextMatrix(7, 0) = "E"
        
        'Asignar el ancho de los encabezados
        .ColWidth(0) = 500
        If N_RegModulo9 > 5 Then N_RegModulo9 = 5
        For N_ColI = 1 To N_RegModulo9
            .ColWidth(N_ColI) = 1000
        Next N_ColI
    End With
S_s = FU_ExtraeCadenaGridValoresID(gs_Proyecto, gs_Etapa)
End Sub

Sub PR_HaceVisblesCtrls(S_MiPantalla, Bo_VisibleSN As Boolean)
Select Case UCase(Trim(S_MiPantalla))
    Case "AUDITORIA"
        With Frm_Auditoria
            .Fra_Aud0.Visible = Bo_VisibleSN
            .Fra_Aud1.Visible = Bo_VisibleSN
            .Lbl_AudEstatus.Visible = Bo_VisibleSN
            .Cmb_AudEstatus.Visible = Bo_VisibleSN
            .Lbl_AudTipo.Visible = Bo_VisibleSN
            .Cmb_AudTipo.Visible = Bo_VisibleSN
            .Lbl_AudAuditor.Visible = Bo_VisibleSN
            .Txt_AudAuditor.Visible = Bo_VisibleSN
            .Lbl_AudComentario.Visible = Bo_VisibleSN
            .Txt_AudComentario.Visible = Bo_VisibleSN
        End With
End Select
End Sub
Function FU_ValidaHora(S_HoraParam) As Boolean
Dim N_pos  As Integer, S_Hora As String, S_Min As String

FU_ValidaHora = False
If Len(Trim(S_HoraParam)) <> 5 Then Exit Function
N_pos = InStr(1, Trim(S_HoraParam), ":")
If N_pos = 0 Then Exit Function
If N_pos <> 3 Then Exit Function

S_Hora = Left(Trim(S_HoraParam), N_pos - 1)
S_Min = Right(Trim(S_HoraParam), Len(Trim(S_HoraParam)) - N_pos)
If Val(S_Hora) >= 24 Then Exit Function
If Val(S_Min) > 59 Then Exit Function
FU_ValidaHora = True

End Function

Sub Carga_Arbol(I_CveRaiz As Integer, S_LetreroR As String, CtlArbol As TreeView)
Dim Re_Tabla As ADODB.Recordset
Dim I_Index As Integer                      'Variable para el índice del nodo actual.
Dim S_Patron As String, S_Condi As String, N_MiEtapa As Integer

N_MiEtapa = 0
S_Condi = "SELECT p.n_cveproyecto,Rtrim(Ltrim(str(e.n_cveetapa))) + ' ' + Rtrim(Ltrim(t.s_descrip)) Etapa,e.n_cveetapa " & _
"FROM t_proyectos p (Nolock), t_enccuestionario e (Nolock),c_etapa t (Nolock) " & _
"Where p.n_cveproyecto = e.n_cveproyecto And e.n_cveetapa = t.n_cveetapa And e.n_cveproyecto = " & I_CveRaiz & " " & _
"Order by e.n_cveetapa,e.s_numcuestionario"

Set Re_Tabla = New ADODB.Recordset
Re_Tabla.Open S_Condi, gcn

Set mNode = CtlArbol.Nodes.Add(, , "GDV", S_LetreroR)
Do Until Re_Tabla.EOF
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    If N_MiEtapa <> Re_Tabla(2) Then
        Set mNode = CtlArbol.Nodes.Add("GDV", tvwChild, "G-" & Trim(CStr(Re_Tabla(2))), Trim(Re_Tabla(1)))
        I_Index = mNode.Index
        S_Condi = "e.n_cveproyecto = " & I_CveRaiz & " And e.n_cveetapa = " & Re_Tabla(2)
        Call Ciclo_Interno(S_Condi, "G-" & Trim(CStr(Re_Tabla(2))), CtlArbol)
        N_MiEtapa = Re_Tabla(2)
    End If
    Re_Tabla.MoveNext
Loop
Re_Tabla.Close
CtlArbol.Nodes(1).Expanded = True
'*******************************************************************************************
'*******************************************************************************************
End Sub

Sub Ciclo_Interno(S_Condi As String, S_Relacion As String, CtlArbol As TreeView)
Dim Re_TablaInt1 As ADODB.Recordset, S_CondiInt As String
Dim mNode As Node

S_CondiInt = "SELECT p.n_cveproyecto,e.n_cveetapa,e.s_numcuestionario + '  (' + " & _
"Case e.c_estatusinf " & _
"When 'X' then 'Sin cuestionario' " & _
"When 'G1' then 'Primera Captura' " & _
"When 'G2' then 'Captura de Producción'" & _
"Else 'Verificar Estado' " & _
"End + ')' EstadoCues,e.s_numcuestionario " & _
"FROM t_proyectos p (Nolock), t_enccuestionario e (Nolock) " & _
"Where p.n_cveproyecto = e.n_cveproyecto And " & S_Condi & " " & _
"Order by e.n_cveetapa,e.s_numcuestionario "

Set Re_TablaInt1 = New ADODB.Recordset
Re_TablaInt1.Open S_CondiInt, gcn

Do Until Re_TablaInt1.EOF
    Set mNode = CtlArbol.Nodes.Add(S_Relacion, tvwChild, "B-" & Re_TablaInt1(3), Trim(Re_TablaInt1(2)))
    Re_TablaInt1.MoveNext
Loop
'*******************************************************************************************
'*
'*******************************************************************************************
End Sub

Sub PR_ComboKeyUp(cmb_ctrl As ComboBox, Codigo As Integer)
Dim LenText As Long, ret As Long
   
'Si los caracteres presionados están entre el 0 y la Z
If Codigo >= vbKey0 And Codigo <= vbKeyZ Or (Codigo >= 96 And Codigo <= 105) Then
    ret = SendMessage(cmb_ctrl.hwnd, &H14C&, -1, ByVal cmb_ctrl.Text)
    
    If ret >= 0 Then
       LenText = Len(cmb_ctrl.Text)
       cmb_ctrl.ListIndex = ret
       cmb_ctrl.Text = cmb_ctrl.List(ret)
       cmb_ctrl.SelStart = LenText
       cmb_ctrl.SelLength = Len(cmb_ctrl.Text) - LenText
    End If
End If
End Sub
Sub PR_PoneCeroCapturaAB()
Dim li_Contrl As Integer

With Frm_Auditoria
    For li_Contrl = 0 To Val(.Txt_AudMayor.UBound)
        .Txt_AudMayor(li_Contrl) = 0
    Next
End With
End Sub
Function FU_RetornaIndiceCombo(cmb_ctrl As ComboBox) As Integer
Dim LenText As Long, ret As Long

FU_RetornaIndiceCombo = -1
 ret = SendMessage(cmb_ctrl.hwnd, &H14C&, -1, ByVal cmb_ctrl.Text)
    
    If ret >= 0 Then
       cmb_ctrl.ListIndex = ret
    End If
    FU_RetornaIndiceCombo = ret
End Function

Sub PR_PoneIndices_AuditoriaBitacora(S_CualIndice)
Dim N_ValAsignado As Double, N_ValorTexto As Variant

N_ValAsignado = 0
With Frm_Auditoria
    Select Case UCase(Trim(S_CualIndice))
        Case "MENOR"
            N_ValorTexto = 0
            If Len(Trim(.Txt_AudMayor(0))) Then
                N_ValorTexto = IIf(IsNumeric(Trim(.Txt_AudMayor(0))), Val(.Txt_AudMayor(0)), 0)
                'If Val(N_ValorTexto) > 0 Then N_ValAsignado = N_ValAsignado + 3    Minuta 15-Oct-2008
                N_ValAsignado = N_ValAsignado + N_ValorTexto
            End If
            If Len(Trim(.Txt_AudMayor(1))) Then
                N_ValorTexto = IIf(IsNumeric(Trim(.Txt_AudMayor(1))), Val(.Txt_AudMayor(1)), 0)
                'If Val(N_ValorTexto) > 0 Then N_ValAsignado = N_ValAsignado + 4    Minuta 15-Oct-2008
                N_ValAsignado = N_ValAsignado + N_ValorTexto
            End If
            If Len(Trim(.Txt_AudMayor(2))) Then
                N_ValorTexto = IIf(IsNumeric(Trim(.Txt_AudMayor(2))), Val(.Txt_AudMayor(2)), 0)
                'If Val(N_ValorTexto) > 0 Then N_ValAsignado = N_ValAsignado + 6    Minuta 15-Oct-2008
                N_ValAsignado = N_ValAsignado + N_ValorTexto
            End If
            If Len(Trim(.Txt_AudMayor(3))) Then
                N_ValorTexto = IIf(IsNumeric(Trim(.Txt_AudMayor(3))), Val(.Txt_AudMayor(3)), 0)
                'If Val(N_ValorTexto) > 0 Then N_ValAsignado = N_ValAsignado + 6    Minuta 15-Oct-2008
                N_ValAsignado = N_ValAsignado + N_ValorTexto
            End If
            .Txt_AudIndices(0) = N_ValAsignado  'Menor
            
        Case "MAYOR"
            N_ValorTexto = 0
            If Len(Trim(.Txt_AudMayor(4))) Then
                N_ValorTexto = IIf(IsNumeric(Trim(.Txt_AudMayor(4))), Val(.Txt_AudMayor(4)), 0)
                'If Val(N_ValorTexto) > 0 Then N_ValAsignado = N_ValAsignado + 7    Minuta 15-Oct-2008
                N_ValAsignado = N_ValAsignado + N_ValorTexto
            End If
            If Len(Trim(.Txt_AudMayor(5))) Then
                N_ValorTexto = IIf(IsNumeric(Trim(.Txt_AudMayor(5))), Val(.Txt_AudMayor(5)), 0)
                'If Val(N_ValorTexto) > 0 Then N_ValAsignado = N_ValAsignado + 7    Minuta 15-Oct-2008
                N_ValAsignado = N_ValAsignado + N_ValorTexto
            End If
            If Len(Trim(.Txt_AudMayor(6))) Then
                N_ValorTexto = IIf(IsNumeric(Trim(.Txt_AudMayor(6))), Val(.Txt_AudMayor(6)), 0)
                'If Val(N_ValorTexto) > 0 Then N_ValAsignado = N_ValAsignado + 10   Minuta 15-Oct-2008
                N_ValAsignado = N_ValAsignado + N_ValorTexto
            End If
            If Len(Trim(.Txt_AudMayor(7))) Then
                N_ValorTexto = IIf(IsNumeric(Trim(.Txt_AudMayor(7))), Val(.Txt_AudMayor(7)), 0)
                'If Val(N_ValorTexto) > 0 Then N_ValAsignado = N_ValAsignado + 10   Minuta 15-Oct-2008
                N_ValAsignado = N_ValAsignado + N_ValorTexto
            End If
            .Txt_AudIndices(1) = N_ValAsignado  'Mayor
            
        Case "CRITICA"
            For i = 8 To 18
                N_ValorTexto = IIf(IsNumeric(Trim(.Txt_AudMayor(i))), Val(.Txt_AudMayor(i)), 0)
                N_ValAsignado = N_ValAsignado + N_ValorTexto
            Next i
            
            If N_ValAsignado > 0 Then
                N_ValAsignado = 100
            Else
                N_ValAsignado = 0
            End If
            .Txt_AudIndices(2) = N_ValAsignado  'Critica
            
            '-------------PARA EL INDICE DE CALIDAD------------------
            If Val(N_ValAsignado) = 100 Then
                .Txt_AudIndices(3) = "0.00"    'Indice de Calidad
            Else
                .Txt_AudIndices(3) = 100 - (Val(.Txt_AudIndices(0)) * 0.19 + Val(.Txt_AudIndices(1)) * 0.34)  'Indice de Calidad
            End If
            
    End Select
End With
End Sub

Sub PR_HabilitaPestanas()
    If Frm_Caratula.Lbl_CarCveEstatus.Caption = 1 Then
        Call PR_HabilitaNavegadores("PESTANAS-FRAMES", True)
        Call PR_HabilitaNavegadores("CHECKS", True)
        Exit Sub
    End If
    
    If Frm_Caratula.Lbl_CarCveEstatus.Caption = 2 Then
        Call PR_HabilitaNavegadores("PESTANAS-CHECKS", False)
    Else
        Call PR_HabilitaNavegadores("PESTANAS-CHECKS", True)
    End If

    If Frm_Caratula.Lbl_CarCveEstatus.Caption = 3 Then
        Call PR_HabilitaNavegadores("PESTANAS-FRAMES", False)
    Else
        Call PR_HabilitaNavegadores("PESTANAS-FRAMES", True)
    End If
End Sub
Sub PR_colocaLetrerosModulo9(S_CualBoton)
With Frm_Caratula
    Select Case UCase(Trim(S_CualBoton))
        Case "CARA_CARA"
            .Chk_CarMod9(6).Visible = False
            .Chk_CarMod9(7).Visible = False
            
            .Chk_CarMod9(0).Caption = "Efectiva"
            .Chk_CarMod9(1).Caption = "Deshabitada"
            .Chk_CarMod9(2).Caption = "Ausente"
            .Chk_CarMod9(3).Caption = "No Abrio"
            .Chk_CarMod9(4).Caption = "No Coopero"
            .Chk_CarMod9(5).Caption = "Cortada"
            '.Chk_CarMod9(6).Caption = "No Contesta"
            '.Chk_CarMod9(7).Caption = "Equivocado"
            .Chk_CarMod9(8).Caption = "No vive en Cd."
            .Chk_CarMod9(9).Caption = "No Elegible"
        Case "TELEFONICO"
            .Chk_CarMod9(6).Visible = True
            .Chk_CarMod9(7).Visible = True
            
            .Chk_CarMod9(0).Caption = "Efectiva"
            .Chk_CarMod9(1).Caption = "No Coopero"
            .Chk_CarMod9(2).Caption = "Ausente"
            .Chk_CarMod9(3).Caption = "Cortada"
            .Chk_CarMod9(4).Caption = "Equivocado"
            .Chk_CarMod9(5).Caption = "No Existe"
            .Chk_CarMod9(6).Caption = "No Contesta"
            .Chk_CarMod9(7).Caption = "Suspendido/Fuera de Serv."
            .Chk_CarMod9(8).Caption = "No vive ahí"
            .Chk_CarMod9(9).Caption = "No Elegible"
    End Select
End With
End Sub

Function FU_DetectaValoresTipoSupervision() As Boolean
Dim N_Ctl As Integer

FU_DetectaValoresTipoSupervision = False

With Frm_CuestionarioDin
    For N_Ctl = 0 To Val(.Txt_CueDinMod7.UBound)
    Select Case UCase(Trim(.Lbl_CueDinMod7(N_Ctl).Caption))
        Case "TIPO SUPERVISION GDV"
            If Len(Trim(.Txt_CueDinMod7(N_Ctl))) > 0 Then
                FU_DetectaValoresTipoSupervision = True
                Exit For
            End If
        Case "TIPO SUPERVISION OUTSOURCING"
             If Len(Trim(.Txt_CueDinMod7(N_Ctl))) > 0 Then
                FU_DetectaValoresTipoSupervision = True
                Exit For
            End If
       End Select
    Next N_Ctl
End With
End Function

Sub PR_CargaPantallaResumen()
Dim S_Micondicion As String

    Load Frm_Resumen
    Frm_Resumen.Lbl_ResTitulo.Caption = "Captura de Producción de Resumen de Contactos"
    Frm_Resumen.Caption = "Captura de Producción de Resumen de Contactos"
    Frm_Resumen.Bot_Res(0).Enabled = False
    '-->Cuando no existan registro del módulo 9 para la captura del Resumen de Contactos
    S_Micondicion = "n_cveproyecto = " & gs_Proyecto & " And n_cveetapa = " & gs_Etapa & " And n_modulo = 9"
    If FU_Cuenta_Registros("t_caratulas", S_Micondicion) = 0 Then
        Frm_Resumen.Bot_Res(1).Enabled = False
    End If
    '<--
    Frm_Resumen.Show vbModal
End Sub
