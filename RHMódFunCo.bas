Attribute VB_Name = "RHMód_FunFormCond"
Const cyAccess = 1

Function BuscaComboClave(ByRef Combo As ComboBox, ByVal sBus As String, bCve As Boolean, Optional bLike As Boolean) As Long
Dim y As Integer, b As Boolean, s As String
If IsNull(bLike) Then
    b = False
Else
    b = bLike
End If
BuscaComboClave = -1
If InStrRev(Combo.List(y), "(") = 0 Then Exit Function
If Not bCve Then
    For y = 0 To Combo.ListCount - 1
        s = Mid(Combo.List(y), InStrRev(Combo.List(y), "(") + 1)
        s = Mid(s, 1, Len(s) - 1)
        If bLike Then
            If s Like sBus + "*" Then
                BuscaComboClave = y
                Exit Function
            End If
        Else
            If s = sBus Then
                BuscaComboClave = y
                Exit Function
            End If
        End If
    Next
Else
    For y = 0 To Combo.ListCount - 1
        If Combo.ItemData(y) = Val(sBus) Then
            BuscaComboClave = y
            Exit Function
        End If
    Next
End If
End Function

Function BuscaCombo(ByRef Combo As ComboBox, ByVal sBus As String, bCve As Boolean, Optional bLike As Boolean, Optional bClave As Boolean) As Long
Dim y As Integer, b As Boolean
If IsNull(bLike) Then
    b = False
Else
    b = bLike
End If
BuscaCombo = -1
If Not bCve Then
    If bLike Then
        For y = 0 To Combo.ListCount - 1
            If UCase(Combo.List(y)) Like "*" + UCase(sBus) + "*" Then
                BuscaCombo = y
                Exit Function
            End If
        Next
    Else
        If bClave And InStrRev(Combo.List(y), "(") > 0 And InStrRev(sBus, "(") > 0 Then
            For y = 0 To Combo.ListCount - 1
                If Mid(Combo.List(y), 1, InStrRev(Combo.List(y), "(")) = Mid(sBus, 1, InStrRev(sBus, "(")) Then
                    BuscaCombo = y
                    Exit Function
                End If
            Next
        Else
            For y = 0 To Combo.ListCount - 1
                If Combo.List(y) = sBus Then
                    BuscaCombo = y
                    Exit Function
                End If
            Next
        End If
    End If
Else
    For y = 0 To Combo.ListCount - 1
        If Combo.ItemData(y) = Val(sBus) Then
            BuscaCombo = y
            Exit Function
        End If
    Next
End If
End Function

Function BuscaList(ByRef List As ListBox, ByVal sBus As String, bCve As Boolean, Optional bLike As Boolean) As Long
Dim y As Integer, b As Boolean
If IsNull(bLike) Then
    b = False
Else
    b = bLike
End If
BuscaList = -1
If Not bCve Then
    For y = 0 To List.ListCount - 1
        If bLike Then
            If List.List(y) Like sBus + "*" Then
                BuscaList = y
                Exit Function
            End If
        Else
            If List.List(y) = sBus Then
                BuscaList = y
                Exit Function
            End If
        End If
    Next
Else
    For y = 0 To List.ListCount - 1
        If List.ItemData(y) = Val(sBus) Then
            BuscaList = y
            Exit Function
        End If
    Next
End If
End Function

'Pone en el arreglo los valores 0 ó 1
'1: Debe el campo conservar su valor al Limpiar
Sub AsignaValor(ByRef yDatos() As Byte, yTabla As Byte, bEntrada As Boolean)
Dim yMax As Byte, y As Byte, l As Double
yMax = UBound(yDatos) + 1
If bEntrada Then  'Load de la forma
    'Set rs = gdbMidb.OpenRecordset("select * from nolimpiar where tabla=" + Str(yTabla), dbOpenDynaset)
    Set rs = gdbmiconfig.OpenRecordset("select * from nolimpiar where tabla=" + Str(yTabla) + " and idusi=" + Str(gs_usuario), dbOpenDynaset)  '****************************jps
    If Not rs.EOF Then
        l = rs(1)
        For y = 0 To yMax - 1
            If l >= (2 ^ (yMax - y)) Then
                l = l - (2 ^ (yMax - y))
                yDatos(y) = 1
            End If
        Next
    End If
Else  'UnLoad de la forma
    l = 0
    For y = 0 To yMax - 1
        If yDatos(y) = 1 Then
            l = l + (2 ^ (yMax - y))
        End If
    Next
    'gdbMidb.Execute "update nolimpiar set valor=" + Trim(Str(l)) + " where tabla=" + Str(yTabla)
    Set rs = gdbmiconfig.OpenRecordset("select count(*) from nolimpiar where idusi=" + Str(gs_usuario) + " and tabla=" + Str(yTabla), dbOpenSnapshot)
    If rs(0) > 0 Then
        gdbmiconfig.Execute "update nolimpiar set valor=" + Trim(Str(l)) + " where idusi=" + Str(gs_usuario) + " and tabla=" + Str(yTabla)
    Else
        gdbmiconfig.Execute "insert into nolimpiar (idusi,tabla,valor ) values (" + Str(gs_usuario) + "," + Str(yTabla) + "," + Trim(Str(l)) + ")"
    End If
End If
End Sub

Sub LlenaComboClave(ByRef Combo As ComboBox, sTabla As String, sCondición As String, Optional bQuery As Boolean, Optional bDescripciónÚnica As Boolean, Optional sNombreCorto As String, Optional bNoBorraCombo As Boolean)
Dim rs As Recordset, s As String, i As Integer, bSio As Boolean
sTabla = LCase(sTabla)
s = gsNoCatálogosFijos
If gbSIO_mdb Then
    Do While InStr(s, ",")
        If InStr(sTabla, Mid(s, 1, InStr(s, ",") - 1)) Then
            Exit Do
        End If
        s = Mid(s, InStr(s, ",") + 1)
    Loop
    bSio = (InStr(s, ",") = 0)
End If
'Do While InStr(s, ",") And Not bQuery
'    If InStr(LCase(sTabla), Mid(s, 1, InStr(s, ",") - 1)) Then
'        i = 200
'        Exit Do
'    End If
'    s = Mid(s, InStr(s, ",") + 1)
'Loop
's = ""
If Not bNoBorraCombo Then Combo.Clear
If gSQLACC = cyAccess Or bSio Then
    If bQuery Then
        Set rs = gdbmidb.OpenRecordset(sTabla, dbOpenSnapshot)
    Else
        Set rs = gdbmidb.OpenRecordset("select * from " + sTabla + IIf(Len(sCondición) > 0, " where " + sCondición, "") + " order by " + IIf(Len(sNombreCorto) > 0, "iif(len(" + sNombreCorto + ")>2," + sNombreCorto + ",descripción)", "descripción"), dbOpenSnapshot)
    End If
    Do While Not rs.EOF
        If Not bDescripciónÚnica Or s <> rs!descripción Then
            If Len(sNombreCorto) > 0 Then
                If Len(rs(sNombreCorto)) > 2 Then
                    Combo.AddItem rs(sNombreCorto) + " (" + Trim(IIf(IsNull(rs!clave), Str(rs!ID), rs!clave)) + ": " + rs!descripción + ")"
                Else
                    Combo.AddItem rs!descripción + " (" + Trim(IIf(IsNull(rs!clave), Str(rs!ID), rs!clave)) + ")"
                End If
            Else
                Combo.AddItem rs!descripción + " (" + Trim(IIf(IsNull(rs!clave), Str(rs!ID), rs!clave)) + ")"
            End If
            'Combo.AddItem rs!descripción + " (" + Trim(IIf(IsNull(rs!clave), Str(rs!ID), rs!clave)) + ")"
            Combo.ItemData(Combo.NewIndex) = rs!ID
            s = rs!descripción
        End If
        rs.MoveNext
    Loop
Else
    Dim rsSQLtabla As New ADODB.Recordset
    If bQuery Then
        rsSQLtabla.Open Replace(sTabla, "lcase", "lower"), gConSql, adOpenStatic, adLockReadOnly
    Else
        sCondición = Replace(sCondición, "lcase", "lower")
        rsSQLtabla.Open "select * from " + sTabla + IIf(Len(sCondición) > 0, " where " + sCondición, "") + " order by " + IIf(Len(sNombreCorto) > 0, "case when " + sNombreCorto + " is not null then (case when len(nombre_corto_if)>3 then nombre_corto_if else descripción end) else descripción end", "descripción"), gConSql, adOpenStatic, adLockReadOnly
    End If
    Do While Not rsSQLtabla.EOF
        If Not bDescripciónÚnica Or s <> rsSQLtabla!descripción Then
            If Len(sNombreCorto) > 0 Then
                If Len(rsSQLtabla(sNombreCorto)) > 2 Then
                    Combo.AddItem rsSQLtabla(sNombreCorto) + " (" + Trim(IIf(IsNull(rsSQLtabla!clave), Str(rsSQLtabla!ID), rsSQLtabla!clave)) + ")"
                Else
                    Combo.AddItem rsSQLtabla!descripción + " (" + Trim(IIf(IsNull(rsSQLtabla!clave), Str(rsSQLtabla!ID), rsSQLtabla!clave)) + ")"
                End If
            Else
                Combo.AddItem rsSQLtabla!descripción + " (" + Trim(IIf(IsNull(rsSQLtabla!clave), Str(rsSQLtabla!ID), rsSQLtabla!clave)) + ")"
            End If
            'Combo.AddItem rs!descripción + " (" + Trim(IIf(IsNull(rs!clave), Str(rs!ID), rs!clave)) + ")"
            Combo.ItemData(Combo.NewIndex) = rsSQLtabla!ID
            s = rsSQLtabla!descripción
        End If
        rsSQLtabla.MoveNext
    Loop
End If
End Sub

Sub LlenaCombo(ByRef Combo As ComboBox, sTabla As String, sCondición As String, Optional bQuery As Boolean, Optional bNoBorraCombo As Boolean)
Dim rs As Recordset, rsSQLCOMBO As New ADODB.Recordset, bSio As Boolean, s As String
sTabla = LCase(sTabla)
s = gsNoCatálogosFijos
If gbSIO_mdb Then
    Do While InStr(s, ",")
        If InStr(sTabla, Mid(s, 1, InStr(s, ",") - 1)) Then
            Exit Do
        End If
        s = Mid(s, InStr(s, ",") + 1)
    Loop
    bSio = (InStr(s, ",") = 0)
End If
'rsSQLCOMBO.CursorLocation = adUseClient
If bQuery Then
    rsSQLCOMBO.Open sTabla, gConSql, adOpenStatic, adLockReadOnly
Else
    rsSQLCOMBO.Open "select * from " + sTabla + IIf(Len(sCondición) > 0, " where " + sCondición, "") + " order by descripción", gConSql, adOpenStatic, adLockReadOnly
End If
If Not bNoBorraCombo Then Combo.Clear
Do While Not rsSQLCOMBO.EOF
    Combo.AddItem IIf(IsNull(rsSQLCOMBO(1)), "---", rsSQLCOMBO(1))
    Combo.ItemData(Combo.NewIndex) = rsSQLCOMBO(0)
    rsSQLCOMBO.MoveNext
Loop

End Sub

Sub LlenaLista(ByRef List As ListBox, sTabla As String, sCondición As String)
Dim rs As Recordset
If gSQLACC = cyAccess Then
    Set rs = gdbmidb.OpenRecordset("select * from " + sTabla + IIf(Len(sCondición) > 0, " where " + sCondición, "") + " order by descripción", dbOpenSnapshot)
    List.Clear
    Do While Not rs.EOF
        List.AddItem rs!descripción
        List.ItemData(List.ListCount - 1) = rs!ID
        rs.MoveNext
    Loop
Else
    Dim rsSQLtemp As New ADODB.Recordset
    'rsSQLtemp.CursorLocation = adUseClient
    rsSQLtemp.Open "select * from " + sTabla + IIf(Len(sCondición) > 0, " where " + sCondición, "") + " order by descripción", gCadSQL, adOpenStatic, adLockReadOnly, adCmdText
    List.Clear
    Do While Not rsSQLtemp.EOF
        List.AddItem rsSQLtemp!descripción
        List.ItemData(List.NewIndex) = rsSQLtemp!ID
        rsSQLtemp.MoveNext
    Loop
End If
End Sub

Function TeclaOprimida(ByRef Campo As TextBox, iAscii As Integer, sTipo As String, bAsterisco As Boolean) As Integer
Dim bIns As Boolean, s As String
TeclaOprimida = iAscii
If InStr(sTipo, "|") Then
    sTipo = Mid(sTipo, 1, InStr(sTipo, "|") - 1)
End If
If iAscii <> 8 Then
    If (InStr("*?", Chr(iAscii)) > 0 And Not bAsterisco) Or InStr("'|¬" + Chr(34), Chr(iAscii)) > 0 Then
        TeclaOprimida = 0
        Exit Function
    End If
    If sTipo = "f" Or sTipo = "fh" And InStr(Campo, " ") = 0 Then
        If InStr("*?", Chr(iAscii)) > 0 Then
            If (InStr(Campo.Text, ">") > 0 Or InStr(Campo.Text, "=") > 0 Or InStr(Campo.Text, "<") > 0 Or (InStr(Campo.Text, "-") > 0 And Len(Campo) > 1)) Then TeclaOprimida = 0
        ElseIf InStr("><=-", Chr(iAscii)) > 0 Then
            If Not bAsterisco Then TeclaOprimida = 0
            If (InStr(Campo.Text, "*") > 0 Or InStr(Campo.Text, "?") > 0) Then TeclaOprimida = 0
            If Chr(iAscii) = "=" And (Len(Campo) <> 1 Or InStr("<>", Campo) = 0) Then TeclaOprimida = 0
            If InStr("><", Chr(iAscii)) > 0 And Len(Campo) > 1 Then TeclaOprimida = 0
            If InStr("-", Chr(iAscii)) > 0 And Not IsDate(Campo) And Len(Trim(Campo)) > 0 Then TeclaOprimida = 0
        End If
        s = QuitaCadena(Campo.Text, "<>=")
        If TeclaOprimida = 0 Then Exit Function
        If InStr(Campo, "-") > 0 Then
            s = Mid(s, InStr(Campo, "-") + 1)
        End If
        If Len(Trim(s)) = 1 Then
            If InStr("0123456789", s) > 0 And InStr("0123456789", Chr(iAscii)) Then bIns = True
        ElseIf Len(Trim(s)) = 4 Then
            If InStr("0123456789", Mid(s, 4, 1)) > 0 And InStr("0123456789", Chr(iAscii)) Then bIns = True
        ElseIf Len(Trim(s)) = 5 Then
            If IsDate(Trim(s) + Chr(iAscii) + "/" + "00") Then bIns = True
        End If
        If bIns Then
            Campo = Trim(Campo) + Chr(iAscii) + "/"
            TeclaOprimida = 0
            Campo.SelStart = Len(Trim(Campo))
        End If
    ElseIf sTipo = "h" Or sTipo = "fh" And InStr(Campo, " ") > 0 Then
        If InStr("*?", Chr(iAscii)) > 0 Then
            If (InStr(Campo.Text, ">") > 0 Or InStr(Campo.Text, "=") > 0 Or InStr(Campo.Text, "<") > 0 Or InStr(Campo.Text, "-") > 0) Then TeclaOprimida = 0
        ElseIf InStr("><=-", Chr(iAscii)) > 0 Then
            If Not bAsterisco Then TeclaOprimida = 0
            If (InStr(Campo.Text, "*") > 0 Or InStr(Campo.Text, "?") > 0) Then TeclaOprimida = 0
            If Chr(iAscii) = "=" And (Len(Campo) <> 1 Or InStr("<>", Campo) = 0) Then TeclaOprimida = 0
            If InStr("><", Chr(iAscii)) > 0 And Len(Campo) > 1 Then TeclaOprimida = 0
            If InStr("-", Chr(iAscii)) > 0 And Not IsDate(Campo) Then TeclaOprimida = 0
        End If
        s = QuitaCadena(Campo.Text, "<>=")
        If TeclaOprimida = 0 Then Exit Function
        If InStr(Campo, "-") > 0 Then
            s = Mid(s, InStr(Campo, "-") + 1)
        End If
        If sTipo = "fh" And InStr(s, " ") > 0 Then
            s = Mid(s, InStr(s, " ") + 1)
        End If
        If Len(Trim(s)) = 1 Then
            If InStr("0123456789", s) > 0 And InStr("0123456789", Chr(iAscii)) Then bIns = True
        ElseIf Len(Trim(s)) = 4 Then
            If InStr("0123456789", Mid(s, 4, 1)) > 0 And InStr("0123456789", Chr(iAscii)) Then bIns = True
        End If
        If bIns Then
            Campo = Trim(Campo) + Chr(iAscii) + ":"
            TeclaOprimida = 0
            Campo.SelStart = Len(Trim(Campo))
        End If
    ElseIf (sTipo = "n" Or sTipo = "m") And InStr("0123456789.", Chr(iAscii)) = 0 Then
        If InStr("*?", Chr(iAscii)) > 0 Then
            If (InStr(Campo.Text, ">") > 0 Or InStr(Campo.Text, "=") > 0 Or InStr(Campo.Text, "<") > 0 Or InStr(Campo.Text, "-") > 0) Then TeclaOprimida = 0
        ElseIf InStr("><=-", Chr(iAscii)) > 0 Then
            If Not bAsterisco Then TeclaOprimida = 0
            If (InStr(Campo.Text, "*") > 0 Or InStr(Campo.Text, "?") > 0) Then TeclaOprimida = 0
            If Chr(iAscii) = "=" And (Len(Campo) <> 1 Or InStr("<>", Campo) = 0) Then TeclaOprimida = 0
            If InStr("><", Chr(iAscii)) > 0 And Len(Campo) > 1 Then TeclaOprimida = 0
            If InStr("-", Chr(iAscii)) > 0 And (Len(Campo) = 0 Or InStr(Campo, "-") > 0) Then TeclaOprimida = 0
        Else: TeclaOprimida = 0
        End If
    Else
        'TeclaOprimida = Asc(UCase(Chr(iAscii)))
    End If
End If
End Function

'yFormaAct_0_Bus_1_Ins_2 0:actualizar,1:buscar y 2:insertar
Function ArmaCadenaCampo(sCampo As String, sCon As String, sTipo As String, yFormaAct_0_Bus_1_Ins_2 As Byte, Optional bCrystal As Boolean) As String
Dim s As String, ss As String, d As Date, cyAccess As Byte
'Quita información de más cuando trae pipe
If InStr(sTipo, "|") Then
    sTipo = Mid(sTipo, 1, InStr(sTipo, "|") - 1)
End If
cyAccess = 1
If sTipo = "m" Then
    sCon = QuitaCadena(sCon, "$, ")
End If
If yFormaAct_0_Bus_1_Ins_2 = 1 And InStr("><", Mid(sCon, 1, 1)) > 0 And Len(sCon) > 0 Then
    ss = Mid(sCon, 1, IIf(Mid(sCon, 2, 1) = "=", 2, 1))
Else
    ss = "="
End If
If yFormaAct_0_Bus_1_Ins_2 = 1 And Len(sCon) > 2 And InStr(sCon, "-") > 1 And InStr("mnfh", sTipo) > 0 Then
    ArmaCadenaCampo = ArmaCadenaCampo(sCampo, "<=" + Mid(sCon, InStr(sCon, "-") + 1), sTipo, yFormaAct_0_Bus_1_Ins_2)
    ss = ">="
    sCon = Mid(sCon, 1, InStr(sCon, "-") - 1)
Else
    ArmaCadenaCampo = ""
End If
sCon = QuitaCadena(sCon, "><=")
s = IIf(yFormaAct_0_Bus_1_Ins_2 <= 1, sCampo + ss, "")
If sTipo = "n" Or sTipo = "m" Then
    If yFormaAct_0_Bus_1_Ins_2 = 1 And (InStr(sCon, "*") Or InStr(sCon, "?")) Then
        If sCon = "-*" Then
            If gSQLACC = cyAccess Then
                s = " (isnull(" + sCampo + ") " + IIf(bCrystal, ")", " or " + sCampo + "=0)")
            Else  'SQL
                s = " (" + sCampo + " is null " + IIf(bCrystal, ")", " or " + sCampo + "=0)")
            End If
        ElseIf sCon = "*" Then
            If gSQLACC = cyAccess Then
                s = " not (isnull(" + sCampo + ") " + IIf(bCrystal, ")", "or " + sCampo + "=0)")
            Else  'SQL
                s = " not (" + sCampo + " is null " + IIf(bCrystal, ")", " or " + sCampo + "=0)")
            End If
        Else
            If gSQLACC = cyAccess Then
                s = IIf(bCrystal, " (totext(", " ltrim(str(") + sCampo + ")) like '" + Trim(sCon) + "'"
            Else  'SQL
                sCon = Replace(Replace(sCon, "_", "[_]"), "%", "[%]")
                sCon = Replace(Replace(sCon, "*", "%"), "?", "[_]")
                s = IIf(bCrystal, " (totext(", " ltrim(str(") + sCampo + ")) like '" + Trim(sCon) + "'"
            End If
        End If
    Else
        s = s + IIf(sCon = "", "null", Str(Val(Trim(sCon))))
    End If
ElseIf sTipo = "c" Or Len(sTipo) = 0 Then
    If yFormaAct_0_Bus_1_Ins_2 = 1 And (InStr(sCon, "*") Or InStr(sCon, "?")) Then
        If sCon = "-*" Then
            If gSQLACC = cyAccess Then
                s = " (isnull(" + sCampo + ") " + IIf(bCrystal, ")", "or len(rtrim(" + sCampo + "))=0)")
            Else  'SQL
                s = " (" + sCampo + " is null " + IIf(bCrystal, ")", "or len(rtrim(" + sCampo + "))=0)")
            End If
        ElseIf sCon = "*" Then
            If gSQLACC = cyAccess Then
                s = " not (isnull(" + sCampo + ") " + IIf(bCrystal, ")", "or len(rtrim(" + sCampo + "))=0)")
            Else  'SQL
                s = " not (" + sCampo + " is null " + IIf(bCrystal, ")", "or len(rtrim(" + sCampo + "))=0)")
            End If
        Else
            sCon = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(LCase(sCon), "á", "a"), "é", "e"), "í", "i"), "ó", "o"), "ú", "u"), "à", "a"), "è", "e"), "ì", "i"), "ò", "o"), "ù", "u")
            sCon = Replace(Replace(Replace(Replace(Replace(sCon, "a", "[aáAÁ]"), "e", "[eéEÉ]"), "i", "[iíIÍ]"), "o", "[oóOÓ]"), "u", "[uúUÚüÜ]")
            If gSQLACC = cyAccess Then
                s = sCampo + " like '" + Trim(sCon) + "'"
            Else
                sCon = Replace(Replace(sCon, "_", "[_]"), "%", "[%]")
                sCon = Replace(Replace(sCon, "*", "%"), "?", "_")
                s = sCampo + " like '" + Trim(sCon) + "'"
            End If
            s = sCampo + " like '" + Trim(sCon) + "'"
        End If
    Else
        If (yFormaAct_0_Bus_1_Ins_2 = 2 Or yFormaAct_0_Bus_1_Ins_2 = 0) And Len(Trim(sCon)) = 0 Then
            s = s + "null"
        Else
            If yFormaAct_0_Bus_1_Ins_2 = 1 And Len(Trim(sCon)) > 0 Then
                sCon = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(LCase(sCon), "á", "a"), "é", "e"), "í", "i"), "ó", "o"), "ú", "u"), "à", "a"), "è", "e"), "ì", "i"), "ò", "o"), "ù", "u")
                sCon = Replace(Replace(Replace(Replace(Replace(sCon, "a", "[aáAÁ]"), "e", "[eéEÉ]"), "i", "[iíIÍ]"), "o", "[oóOÓ]"), "u", "[uúUÚüÜ]")
                If InStr(sCon, "[") > 0 And InStr(sCon, "]") > 0 Then
                    s = sCampo + " like '" + Trim(sCon) + "'"
                Else
                    s = s + "'" + Trim(sCon) + "'"
                End If
            Else
                s = s + "'" + Trim(sCon) + "'"
            End If
        End If
    End If
ElseIf sTipo = "fh" Then
    If bCrystal Then
        s = String(" ", 250) + " 1=1 and "
    Else
        If IsDate(sCon) And yFormaAct_0_Bus_1_Ins_2 = 1 Then
            d = CDate(sCon)
            If Hour(d) + Minute(d) + Second(d) = 0 And InStr(sCon, ":") = 0 Then
                If gSQLACC = cyAccess Then
                    s = "format(" + sCampo + ",'yyyy/mm/dd')" + ss + "'" + Format(CDate(sCon), "yyyy/mm/dd") + "'"
                Else
                    s = "convert(varchar," + sCampo + ",102)" + ss + "'" + Format(CDate(sCon), "yyyy.mm.dd") + "'"
                End If
            Else
                If gSQLACC = cyAccess Then
                    s = "format(" + sCampo + ",'yyyy/mm/dd hh:mm:ss')" + ss + "'" + Format(CDate(sCon), "yyyy/mm/dd hh:mm:ss") + "'"
                Else
                    s = "convert(varchar," + sCampo + ",120)" + ss + "'" + Format(CDate(sCon), "yyyy-mm-dd hh:mm:ss") + "'"
                End If
            End If
        ElseIf IsDate(sCon) Then
            If gSQLACC = cyAccess Then
                s = s + "cdate('" + sCon + "')"
            Else
                s = s & "convert(datetime,'" & Format(CDate(sCon), "dd-mm-yyyy hh:mm:ss") & "',105)"
                's = s + "'" + sCon + "'"
            End If
        ElseIf yFormaAct_0_Bus_1_Ins_2 = 1 And (InStr(sCon, "*") Or InStr(sCon, "?")) Then
            If sCon = "-*" Then
                If gSQLACC = cyAccess Then
                    s = " isnull(" + sCampo + ")"
                Else
                    s = " " + sCampo + " is null"
                End If
            ElseIf sCon = "*" Then
                If gSQLACC = cyAccess Then
                    s = " not isnull(" + sCampo + ")"
                Else
                    s = " " + sCampo + " is not null"
                End If
            Else
                If gSQLACC = cyAccess Then
                    s = "format(" + sCampo + ",'dd/mmm/yyyy hh:mm:ss')" + " like '" + sCon + "'"
                Else
                    sCon = Replace(Replace(sCon, "_", "[_]"), "%", "[%]")
                    sCon = Replace(Replace(sCon, "*", "%"), "?", "[_]")
                    If Not gSQLACC = cyAccess Then sCon = Replace(sCon, "/", "-")
                    s = "convert(varchar," + sCampo + ",120) like '" + sCon + "'"
                End If
            End If
        Else
            If yFormaAct_0_Bus_1_Ins_2 = 2 Or yFormaAct_0_Bus_1_Ins_2 = 0 Then
                s = s + "null"
            Else
                Exit Function
            End If
        End If
    End If
ElseIf sTipo = "f" Then
    If bCrystal Then
        s = sCampo + "=date(" + Str(Year(CDate(sCon))) + "," + Str(Month(CDate(sCon))) + "," + Str(Day(CDate(sCon))) + ")"
    Else
        If IsDate(sCon) And yFormaAct_0_Bus_1_Ins_2 = 1 Then
            If gSQLACC = cyAccess Then
                s = "format(" + sCampo + ",'yyyy/mm/dd')" + ss + "'" + Format(CDate(sCon), "yyyy/mm/dd") + "'"
            Else
                If ss = "=" Then
                    s = sCampo + "  between '" + Format(CDate(sCon), "dd/mm/yyyy") + " 00:00:00.000' and '" + Format(CDate(sCon), "dd/mm/yyyy") + " 23:59:59.999'"
                Else
                    s = sCampo + ss + "'" + Format(CDate(sCon), "dd/mm/yyyy") + IIf(InStr(s, ">"), " 00:00:00.000'", " 23:59:59.999'")
                End If
            End If
        ElseIf IsDate(sCon) Then
            If gSQLACC = cyAccess Then
                s = s + "cdate('" + sCon + "')"
            Else
                s = s & "convert(datetime,'" & Format(CDate(sCon), "dd-mm-yyyy") & "',105)"
            End If
        ElseIf yFormaAct_0_Bus_1_Ins_2 = 1 And (InStr(sCon, "*") Or InStr(sCon, "?")) Then
            If Not gSQLACC = cyAccess Then sCon = Replace(sCon, "/", ".")
            If sCon = "-*" Then
                If gSQLACC = cyAccess Then
                    s = " isnull(" + sCampo + ")"
                Else
                    s = " " + sCampo + " is null"
                End If
            ElseIf sCon = "*" Then
                If gSQLACC = cyAccess Then
                    s = " not isnull(" + sCampo + ")"
                Else
                    s = " " + sCampo + " is not null"
                End If
            Else
                If gSQLACC = cyAccess Then
                    s = "format(" + sCampo + ",'dd/mmm/yyyy')" + " like '" + sCon + "'"
                Else
                    sCon = Replace(Replace(sCon, "_", "[_]"), "%", "[%]")
                    sCon = Replace(Replace(sCon, "*", "%"), "?", "[_]")
                    s = "convert(varchar," + sCampo + ",102) like '" + sCon + "'"
                End If
            End If
        Else
            If yFormaAct_0_Bus_1_Ins_2 = 2 Or yFormaAct_0_Bus_1_Ins_2 = 0 Then
                s = s + "null"
            Else
                Exit Function
            End If
        End If
    End If
ElseIf sTipo = "h" Then
    If bCrystal Then
        s = String(" ", 250) + " 1=1 and "
    Else
        If IsDate(sCon) Then
            If yFormaAct_0_Bus_1_Ins_2 <= 1 Then
                If gSQLACC = cyAccess Then
                    s = "format(" + sCampo + ",'hh:mm:ss')" + ss + "'" + Trim(sCon) + "'"
                Else
                    s = "convert(varchar," + sCampo + ",8)" + ss + "'" + Format(CDate(sCon), "hh:mm:ss") + "'"
                End If
            Else
                If gSQLACC = cyAccess Then
                    s = s + "cdate('" + Trim(sCon) + "')"
                Else
                    s = s + "'" + Trim(sCon) + "'"
                End If
            End If
        ElseIf yFormaAct_0_Bus_1_Ins_2 = 1 And (InStr(sCon, "*") Or InStr(sCon, "?")) Then
            If sCon = "-*" Then
                s = " " + sCampo + " is null"
            ElseIf sCon = "*" Then
                s = " " + sCampo + " not is null"
            Else
                If gSQLACC = cyAccess Then
                    s = "format(" + sCampo + ",'hh:mm:ss') like '" + sCon + "'"
                Else
                    s = "convert(varchar," + sCampo + ",8) like '" + sCon + "'"
                End If
            End If
        Else
            If yFormaAct_0_Bus_1_Ins_2 = 2 Or yFormaAct_0_Bus_1_Ins_2 = 0 Then
                s = s + "null"
            Else
                Exit Function
            End If
        End If
    End If
End If
If Len(s) = 0 Then Exit Function
ArmaCadenaCampo = ArmaCadenaCampo + s + IIf(yFormaAct_0_Bus_1_Ins_2 = 1, " and ", ",")
End Function

Sub CargaDatosArbolVariosNiveles(ByRef tvArbol As TreeView, sQuery As String, yNoNiveles As Byte, Optional bNoBarraProgreso As Boolean, Optional bAbierto As Boolean)
Dim nPro(2) As Integer, y As Byte, n As Integer, s As String, sCve(2) As String, bMov As Boolean, yErr As Byte
Dim ynivel As Byte, iNodos As Integer, iValor() As Integer, l As Long
Dim sCvePadre As String, bSio As Boolean, i As Integer
Dim rs As Recordset
On Error GoTo ErrorCargaDatos:
sQuery = LCase(sQuery)
s = gsNoCatálogosFijos
If gbSIO_mdb Then
    Do While InStr(s, ",")
        If InStr(sQuery, Mid(s, 1, InStr(s, ",") - 1)) Then
            Exit Do
        End If
        s = Mid(s, InStr(s, ",") + 1)
    Loop
    bSio = (InStr(s, ",") = 0)
End If
l = Timer
If gSQLACC = cyAccess Or bSio Then
    Set rs = gdbmidb.OpenRecordset(sQuery, dbOpenSnapshot)
    tvArbol.Visible = False
    tvArbol.Nodes.Clear
    'If IsNull(yNoNiveles) Then yNoNiveles = 3
    
    ReDim iValor(yNoNiveles - 1)
    'Do While Not rs.EOF
    '    rs.MoveNext
    '    Debug.Print Str(rs.AbsolutePosition)
    'Loop
    'Debug.Print Timer - l
    If Not rs.EOF Then
        rs.MoveLast
        l = rs.RecordCount
        n = Int(l / 20)
        rs.MoveFirst
        If n > 0 And Not bNoBarraProgreso Then
            RHBarraProgreso.Show
            RHBarraProgreso.ProgressBar1.Value = 0
            RHBarraProgreso.Refresh
        End If
    End If
    Do While Not rs.EOF
        sCvePadre = "r"
        For ynivel = 0 To yNoNiveles - 1
            If IsNull(rs(ynivel)) Then Exit For
            sCvePadre = sCvePadre + Right("000" + Trim(Str(rs(ynivel))), 4)
            If iValor(ynivel) <> rs(ynivel) Then
                iValor(ynivel) = rs(ynivel)
                For y = ynivel + 1 To yNoNiveles - 1
                    iValor(y) = 0
                Next
                If ynivel = 0 Then
                    Call tvArbol.Nodes.Add(, , sCvePadre, rs(yNoNiveles + ynivel))
                Else
    'AgregaExistente:
                    Call tvArbol.Nodes.Add(Mid(sCvePadre, 1, Len(sCvePadre) - 4), tvwChild, sCvePadre, rs(yNoNiveles + ynivel))
                End If
            End If
        Next
        rs.MoveNext
        If n > 0 And Not bBarraProgreso Then
            If rs.AbsolutePosition Mod n = 0 Then
                RHBarraProgreso.ProgressBar1.Value = 100 * (rs.AbsolutePosition + 1) / l
                RHBarraProgreso.Refresh
            End If
        End If
        'Debug.Print Str(rs.AbsolutePosition)
    Loop
Else
    Dim rsSQL1 As New ADODB.Recordset
    'rsSQL1.CursorLocation = adUseClient
    'rsSQL1.CursorLocation = adUseServer
    rsSQL1.Open sQuery, gCadSQL, adOpenStatic, adLockReadOnly, adCmdText
    tvArbol.Visible = False
    tvArbol.Nodes.Clear
    
    ReDim iValor(yNoNiveles - 1)

    If Not rsSQL1.EOF Then
        rsSQL1.MoveLast
        l = rsSQL1.RecordCount
        n = Int(l / 20)
        rsSQL1.MoveFirst
        If n > 0 And Not bNoBarraProgreso Then
            RHBarraProgreso.Show
            RHBarraProgreso.ProgressBar1.Value = 0
            RHBarraProgreso.Refresh
        End If
    End If
    Do While Not rsSQL1.EOF
        sCvePadre = "r"
        For ynivel = 0 To yNoNiveles - 1
            If IsNull(rsSQL1(Val(ynivel))) Then Exit For
            sCvePadre = sCvePadre + Right("000" + Trim(Str(rsSQL1(Val(ynivel)))), 4)
            If iValor(ynivel) <> rsSQL1(Val(ynivel)) Then
                iValor(ynivel) = rsSQL1(Val(ynivel))
                For y = ynivel + 1 To yNoNiveles - 1
                    iValor(y) = 0
                Next
                If ynivel = 0 Then
                    Call tvArbol.Nodes.Add(, , sCvePadre, rsSQL1(Val(yNoNiveles) + Val(ynivel)))
                Else
                    Call tvArbol.Nodes.Add(Mid(sCvePadre, 1, Len(sCvePadre) - 4), tvwChild, sCvePadre, rsSQL1(Val(yNoNiveles) + Val(ynivel)))
                End If
            End If
        Next
        rsSQL1.MoveNext
        If n > 0 And Not bBarraProgreso Then
            If (Not (rsSQL1.EOF)) And (Not (rsSQL1.EOF)) Then
                If (rsSQL1.Bookmark Mod n = 0) Then
                    RHBarraProgreso.ProgressBar1.Value = 100 * (rsSQL1.Bookmark) / l
                    RHBarraProgreso.Refresh
                End If
            End If
        End If
    Loop
    rsSQL1.Close
End If
RHBarraProgreso.ProgressBar1.Value = 100
RHBarraProgreso.Refresh
Unload RHBarraProgreso
If bAbierto Then
    For y = 1 To tvArbol.Nodes.Count
        tvArbol.Nodes(y).Expanded = True
    Next
End If
tvArbol.LineStyle = tvwRootLines  ' Linestyle = 1
tvArbol.Visible = True
Exit Sub
ErrorCargaDatos:
If Err.Number = 35602 Or Err.Number = -1 Then
    Resume Next
End If
sError = "Error: " + Err.Description
y = MsgBox(sError, vbAbortRetryIgnore + vbCritical, "Error no esperado (" + Str(Err.Number) + ")")
If y = vbCancel Then
    Exit Sub
ElseIf y = vbRetry Then
    Resume
ElseIf y = vbIgnore Then
    Resume Next
End If
End Sub

Sub CargaDatosArbol(ByRef tvArbol As TreeView, sQuery As String, Optional bAbierto As Boolean, Optional ByVal lAsuins As Long, Optional db As DAO.Database, Optional bOtroDB As Boolean)
Dim nPro(3) As Integer, y As Integer, n As Integer, s As String, sCve(3) As String, bMov As Boolean
Dim rs As Recordset, rs2 As Recordset, bSio As Boolean
Dim ss As String
Dim rsSQL1 As New ADODB.Recordset, rsSQL2 As New ADODB.Recordset
sQuery = LCase(sQuery)
s = gsNoCatálogosFijos
If gbSIO_mdb Then
    Do While InStr(s, ",")
        If InStr(sQuery, Mid(s, 1, InStr(s, ",") - 1)) Then
            Exit Do
        End If
        s = Mid(s, InStr(s, ",") + 1)
    Loop
    bSio = (InStr(s, ",") = 0)
    If bSio Then
        ss = "personal  fax       internet  escrito   cat       telefónica"
        For y = 0 To 5
            sQuery = Replace(sQuery, Trim(Mid(ss, y * 10 + 1, 10)) & "=1", Trim(Mid(ss, y * 10 + 1, 10)) & "=-1")
        Next
    End If
End If
If gSQLACC = cyAccess Or bSio Then
    If Not bOtroDB Then
       Set db = gdbmidb
    End If
    Set rs = db.OpenRecordset(sQuery, dbOpenSnapshot)
    tvArbol.Nodes.Clear
    'tvArbol.Nodes.Item.
    'Call MueveSigActividad(rs)
    Do While Not rs.EOF
        bMov = False
        nPro(0) = rs(0)
        sCve(0) = "r" + Right("00" + Trim(Str(nPro(0))), 3) + IIf(IsNull(rs(1)), "", "000") + IIf(IsNull(rs(2)), "", "000") + IIf(IsNull(rs(3)), "", "000")
        Call tvArbol.Nodes.Add(, , sCve(0), rs(4))
        'tvarbol.Nodes(tvarbol.Nodes.Count).
        Do While nPro(0) = rs(0)
            bMov = False
            If Not IsNull(rs(1)) Then
                nPro(1) = rs(1)
                sCve(1) = Mid(sCve(0), 1, 4) + Right("00" + Trim(Str(nPro(1))), 3) + IIf(IsNull(rs(2)), "", "000") + IIf(IsNull(rs(3)), "", "000")
                Call tvArbol.Nodes.Add(sCve(0), tvwChild, sCve(1), rs(5))
                If lAsuins > 0 Then
                    Set rs2 = db.OpenRecordset("select count(*) from avances where idasuins=" + Str(lAsuins), dbOpenSnapshot)
                    If rs2(0) > 0 Then
                        'cambio último
                        'Set rs2 = db.OpenRecordset("select * from Arcos where iddestino=" + Str(nPro(1)) + " and idorigen in (select idact from avances where ultimo and idasuins=" + Str(lAsuIns) + ")", dbOpenSnapshot)
                        Set rs2 = db.OpenRecordset("select * from Arcos where iddestino=" + Str(nPro(1)) + " and idorigen in (select a.idact from avances a left join avances b on a.id=b.idant where b.id is null and a.idasuins=" + Str(lAsuins) + ")", dbOpenSnapshot)
                    Else
                        Set rs2 = db.OpenRecordset("select * from Arcos where iddestino=" + Str(nPro(1)) + " and idorigen=1", dbOpenSnapshot)
                    End If
                    If Not rs2.EOF Then
                        If rs2!valor = 1 Then
                            tvArbol.Nodes(sCve(1)).Tag = 1
                            tvArbol.Nodes(sCve(1)).Checked = True
                        End If
                    End If
                End If
                Do While nPro(0) = rs(0) And nPro(1) = rs(1)
                    bMov = False
                    If Not IsNull(rs(2)) Then
                        nPro(2) = rs(2)
                        sCve(2) = Mid(sCve(1), 1, 7) + Right("00" + Trim(Str(nPro(2))), 3) + IIf(IsNull(rs(2)), "", "000")
                        Call tvArbol.Nodes.Add(sCve(1), tvwChild, sCve(2), rs(6))
                        If lAsuins > 0 Then
                            Set rs2 = db.OpenRecordset("select count(*) from avances where idasuins=" + Str(lAsuins), dbOpenSnapshot)
                            If rs2(0) > 0 Then
                                'cambio último
                                'Set rs2 = db.OpenRecordset("select * from Arcos where iddestino=" + Str(nPro(1)) + " and idorigen in (select idact from avances where ultimo and idasuins=" + Str(lAsuIns) + ")", dbOpenSnapshot)
                                Set rs2 = db.OpenRecordset("select * from Arcos where iddestino=" + Str(nPro(1)) + " and idorigen in (select a.idact from avances a left join avances b on a.id=b.idant where b.id is null and a.idasuins=" + Str(lAsuins) + ")", dbOpenSnapshot)
                            Else
                                Set rs2 = db.OpenRecordset("select * from Arcos where iddestino=" + Str(nPro(1)) + " and idorigen=1", dbOpenSnapshot)
                            End If
                            If Not rs2.EOF Then
                                If rs2!valor = 1 Then
                                    tvArbol.Nodes(sCve(2)).Tag = 1
                                    tvArbol.Nodes(sCve(2)).Checked = True
                                End If
                            End If
                        End If
                        Do While nPro(0) = rs(0) And nPro(1) = rs(1) And nPro(2) = rs(2)
                            If Not IsNull(rs(3)) Then
                                nPro(3) = rs(3)
                                sCve(3) = Mid(sCve(2), 1, 10) + Right("00" + Trim(Str(nPro(3))), 3)
                                Call tvArbol.Nodes.Add(sCve(2), tvwChild, sCve(3), rs(7))
                            Else
                                Exit Do
                            End If
                            rs.MoveNext
                            bMov = True
                            If rs.EOF Then Exit Do
                        Loop
                    Else
                        Exit Do
                    End If
                    If Not bMov Then
                        rs.MoveNext
                        bMov = True
                    End If
                    If rs.EOF Then Exit Do
                Loop
            Else
                Exit Do
            End If
            If Not bMov Then
                rs.MoveNext
                bMov = True
            End If
            If rs.EOF Then Exit Do
        Loop
        If Not bMov Then
            rs.MoveNext
            bMov = True
        End If
    Loop
Else
    'rsSQL1.CursorLocation = adUseClient
    rsSQL1.Open sQuery, gConSql, adOpenStatic, adLockReadOnly, adCmdText
    tvArbol.Nodes.Clear
    'tvArbol.Nodes.Item.
    'Call MueveSigActividad(rs)
    Do While Not rsSQL1.EOF
        bMov = False
        nPro(0) = rsSQL1(0)
        sCve(0) = "r" + Right("00" + Trim(Str(nPro(0))), 3) + IIf(IsNull(rsSQL1(1)), "", "000") + IIf(IsNull(rsSQL1(2)), "", "000") + IIf(IsNull(rsSQL1(3)), "", "000")
        Call tvArbol.Nodes.Add(, , sCve(0), rsSQL1(4))
        'tvarbol.Nodes(tvarbol.Nodes.Count).
        Do While nPro(0) = rsSQL1(0)
            bMov = False
            If Not IsNull(rsSQL1(1)) Then
                nPro(1) = rsSQL1(1)
                sCve(1) = Mid(sCve(0), 1, 4) + Right("00" + Trim(Str(nPro(1))), 3) + IIf(IsNull(rsSQL1(2)), "", "000") + IIf(IsNull(rsSQL1(3)), "", "000")
                Call tvArbol.Nodes.Add(sCve(0), tvwChild, sCve(1), rsSQL1(5))
                If lAsuins > 0 Then
                    rsSQL2.Open "select count(*) from avances where idasuins=" + Str(lAsuins), gCadSQL, adOpenStatic, adLockReadOnly, adCmdText
                    If rsSQL2(0) > 0 Then
                        'cambio último
                        'rsSQL2.Open "select * from Arcos where iddestino=" + Str(nPro(1)) + " and idorigen in (select idact from avances where ultimo and idasuins=" + Str(lAsuIns) + ")", gCadSQL, adOpenStatic, adLockReadOnly, adCmdText
                        rsSQL2.Open "select * from Arcos where iddestino=" + Str(nPro(1)) + " and idorigen in (select a.idact from avances a left join avances b on a.id=b.idant where b.id is null and a.idasuins=" + Str(lAsuins) + ")", gCadSQL, adOpenStatic, adLockReadOnly, adCmdText
                    Else
                        rsSQL2.Open "select * from Arcos where iddestino=" + Str(nPro(1)) + " and idorigen=1", gCadSQL, adOpenStatic, adLockReadOnly, adCmdText
                    End If
                    If Not rsSQL2.EOF Then
                        If rsSQL2!valor = 1 Then
                            tvArbol.Nodes(sCve(1)).Tag = 1
                            tvArbol.Nodes(sCve(1)).Checked = True
                        End If
                    End If
                End If
                Do While nPro(0) = rsSQL1(0) And nPro(1) = rsSQL1(1)
                    bMov = False
                    If Not IsNull(rsSQL1(2)) Then
                        nPro(2) = rsSQL1(2)
                        sCve(2) = Mid(sCve(1), 1, 7) + Right("00" + Trim(Str(nPro(2))), 3) + IIf(IsNull(rsSQL1(2)), "", "000")
                        Call tvArbol.Nodes.Add(sCve(1), tvwChild, sCve(2), rsSQL1(6))
                        If lAsuins > 0 Then
                            rsSQL2.Open "select count(*) from avances where idasuins=" + Str(lAsuins), gCadSQL, adOpenStatic, adLockReadOnly, adCmdText
                            If rsSQL2(0) > 0 Then
                                'cambio último
                                'rsSQL2.Open "select * from Arcos where iddestino=" + Str(nPro(1)) + " and idorigen in (select idact from avances where ultimo and idasuins=" + Str(lAsuIns) + ")", gCadSQL, adOpenStatic, adLockReadOnly, adCmdText
                                rsSQL2.Open "select * from Arcos where iddestino=" + Str(nPro(1)) + " and idorigen in (select a.idact from avances a left join avances b on a.id=b.idant where b.id is null and a.idasuins=" + Str(lAsuins) + ")", gCadSQL, adOpenStatic, adLockReadOnly, adCmdText
                            Else
                                rsSQL2.Open "select * from Arcos where iddestino=" + Str(nPro(1)) + " and idorigen=1", gCadSQL, adOpenStatic, adLockReadOnly, adCmdText
                            End If
                            If Not rsSQL2.EOF Then
                                If rsSQL2!valor = 1 Then
                                    tvArbol.Nodes(sCve(2)).Tag = 1
                                    tvArbol.Nodes(sCve(2)).Checked = True
                                End If
                            End If
                        End If
                        Do While nPro(0) = rsSQL1(0) And nPro(1) = rsSQL1(1) And nPro(2) = rsSQL1(2)
                            If Not IsNull(rsSQL1(3)) Then
                                nPro(3) = rsSQL1(3)
                                sCve(3) = Mid(sCve(2), 1, 10) + Right("00" + Trim(Str(nPro(3))), 3)
                                Call tvArbol.Nodes.Add(sCve(2), tvwChild, sCve(3), rsSQL1(7))
                            Else
                                Exit Do
                            End If
                            rsSQL1.MoveNext
                            bMov = True
                            If rsSQL1.EOF Then Exit Do
                        Loop
                    Else
                        Exit Do
                    End If
                    If Not bMov Then
                        rsSQL1.MoveNext
                        bMov = True
                    End If
                    If rsSQL1.EOF Then Exit Do
                Loop
            Else
                Exit Do
            End If
            If Not bMov Then
                rsSQL1.MoveNext
                bMov = True
            End If
            If rsSQL1.EOF Then Exit Do
        Loop
        If Not bMov Then
            rsSQL1.MoveNext
            bMov = True
        End If
    Loop
End If
If bAbierto Then
    For y = 1 To tvArbol.Nodes.Count
        tvArbol.Nodes(y).Expanded = True
    Next
End If
tvArbol.LineStyle = tvwRootLines  ' Linestyle = 1
End Sub


Function TipoVariable(sTipo As String) As String
Select Case sTipo
Case "10"
    TipoVariable = "c"
Case "4"
    TipoVariable = "n"
Case "5"
    TipoVariable = "f"
Case Else
    TipoVariable = "c"
End Select
End Function

'Quita todos los caracteres contenidos en sQuita de sCadena
Function QuitaCadena(ByVal sCadena As String, ByVal sQuita As String) As String
Dim y As Byte
QuitaCadena = sCadena
For y = 1 To Len(sQuita)
    QuitaCadena = Replace(QuitaCadena, Mid(sQuita, y, 1), "")
Next
End Function

'Quita caracteres no dígitos
Function QuitaNoDígitos(ByVal sCadena As String) As String
Dim y  As Integer
For y = 1 To Len(sCadena)
    If InStr("0123456789", Mid(sCadena, y, 1)) > 0 Then
        QuitaNoDígitos = QuitaNoDígitos & Mid(sCadena, y, 1)
    End If
Next
End Function

'Última modif. por Miguel el Mié.30 de Enero del 2002
Sub RedAsunto(ByRef tvArbol As TreeView, lAsuins As Long)
Dim rs As DAO.Recordset, adors As New ADODB.Recordset, i As Integer, lAva As Long
tvArbol.Nodes.Clear
If gSQLACC = cyAccess Then
    'Set rs = gdbmidb.OpenRecordset("select c.id,a.descripción+' ('+b.descripción+') Fecha: '+format(c.fecha,'dd-mmm-yyyy')+', Tipo: '+iif(a.clase=1,'a','b') from ((select * from actividades where id in (select idact from avances where idasuins=" & lAsuIns & ")) as a left join actividades b on a.idpad=b.id) inner join avances c on a.id=c.idact where c.idasuins=" & lAsuIns & " and c.idant is null", dbOpenSnapshot)
    Set rs = gdbmidb.OpenRecordset("select distinct idact,idant,id from avances where idasuins=" & lAsuins, dbOpenSnapshot)
    If Not rs.EOF Then
        rs.FindFirst "idant=id"
        If Not rs.NoMatch Then
            i = rs(0)
            lAva = rs(2)
        Else
            i = -1
        End If
    Else
        i = -1
    End If
    Set rs = gdbmidb.OpenRecordset("select a.id,b.descripción+' ('+c.descripción+') Fecha: '+format(a.fecha,'dd-mmm-yyyy')+', Tipo: '+iif(b.clase=1,'a','b') from ((select * from avances where id=" & lAva & ") as a inner join (select * from actividades where id=" & i & ") as b on a.idact=b.id) left join actividades c on a.idtar=c.id  where a.fecha is not null", dbOpenSnapshot)
    Do While Not rs.EOF
        i = tvArbol.Nodes.Add(, , "_" & rs(0), rs(1)).Index
        Call AgregaNodoRedAsunto(tvArbol, rs(0), i)
        rs.MoveNext
    Loop
Else
    adors.Open "select distinct idact,idant,id from avances where idasuins=" & lAsuins & " and idant=id", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        i = adors(0)
        lAva = adors(2)
    Else
        i = -1
    End If
    If adors.State Then adors.Close
    adors.Open "select a.id,b.descripción+' ('+c.descripción+') Fecha: '+convert(varchar,a.fecha,100)+', Tipo: '+case when b.clase=1 then 'a' else 'b' end  from ((select * from avances where id=" & lAva & ") as a inner join (select * from actividades where id=" & i & ") as b on a.idact=b.id) left join actividades c on a.idtar=c.id  where a.fecha is not null", gConSql, adOpenStatic, adLockReadOnly
    Do While Not adors.EOF
        i = tvArbol.Nodes.Add(, , "_" & adors(0), adors(1)).Index
        Call AgregaNodoRedAsunto(tvArbol, adors(0), i)
        adors.MoveNext
    Loop
'    adors.Open "select c.id,a.descripción+' ('+b.descripción+') Fecha: '+convert(varchar,c.fecha,100)+', Tipo: '+case when a.clase=1 then 'a' else 'b' end from (actividades a left join actividades b on a.idpad=b.id) inner join avances c on a.id=c.idact where c.idant=c.id and c.idasuins=" + Str(lAsuIns), gConSql, adOpenStatic, adLockReadOnly
'    Do While Not adors.EOF
'        i = tvArbol.Nodes.Add(, , "_" & adors(0), adors(1)).Index
'        Call AgregaNodoRedAsunto(tvArbol, adors(0), i)
'        adors.MoveNext
'    Loop
End If
tvArbol.LineStyle = tvwRootLines  ' Linestyle = 1
For i = 1 To tvArbol.Nodes.Count
    tvArbol.Nodes(i).Expanded = True
Next
End Sub

'Agrega nodos hijos en la red de actividades del asunto lAsunIns
'Creación: 30 de enero del 20002
Sub AgregaNodoRedAsunto(ByRef tvArbol As TreeView, lAvancePadre As Long, iIndiceNodo As Integer)
Dim rs As Recordset, adors As New ADODB.Recordset, i As Integer
If gSQLACC = cyAccess Then
    Set rs = gdbmidb.OpenRecordset("select c.id,a.descripción+' ('+b.descripción+') Fecha: '+format(c.fecha,'dd-mmm-yyyy')+', Tipo: '+iif(a.clase=1,'a','b') from (avances c inner join actividades a on c.idact=a.id) left join actividades b on c.idtar=b.id where c.id<>c.idant and c.idant=" & lAvancePadre & " and c.fecha is not null", dbOpenSnapshot)
    Do While Not rs.EOF
        i = tvArbol.Nodes.Add(iIndiceNodo, tvwChild, tvArbol.Nodes(iIndiceNodo).Key + "_" & rs(0), rs(1)).Index
        Call AgregaNodoRedAsunto(tvArbol, rs(0), i)
        rs.MoveNext
    Loop
Else
    'adors.Open "select c.id,a.descripción+' ('+ (case when b.descripción is null then '' else b.descripción end) +') Fecha: '+ case when c.fecha is null then '' else convert(varchar,c.fecha,100) end +', Tipo: '+ case when a.clase=1 then 'a' else 'b' end,c.fecha,a.clase,b.descripción,a.descripción from (avances c inner join actividades a on c.idact=a.id) left join actividades b on c.idtar=b.id where c.id<>c.idant and c.idant=" & lAvancePadre, gConSql, adOpenStatic, adLockReadOnly
    adors.Open "select c.id,a.descripción+' ('+ (case when b.descripción is null then '' else b.descripción end) +') Fecha: '+ case when c.fecha is null then '' else convert(varchar,c.fecha,100) end +', Tipo: '+ case when a.clase=1 then 'a' else 'b' end from (avances c inner join actividades a on c.idact=a.id) left join actividades b on c.idtar=b.id where c.id<>c.idant and c.idant=" & lAvancePadre, gConSql, adOpenStatic, adLockReadOnly
    Do While Not adors.EOF
        i = tvArbol.Nodes.Add(iIndiceNodo, tvwChild, tvArbol.Nodes(iIndiceNodo).Key + "_" & adors(0), adors(1)).Index
        Call AgregaNodoRedAsunto(tvArbol, adors(0), i)
        adors.MoveNext
    Loop
End If
End Sub


Sub ArbolRed(ByRef tvArbol As TreeView, lProceso As Long)
Dim rs As Recordset, s As String, y As Byte, adors As New ADODB.Recordset
tvArbol.Nodes.Clear
gs = ""
For y = 0 To gyNiveles - 3
    s = s + " or idpad in (select id from actividades where idpad=" + Trim(Str(lProceso))
Next
s = s + String(gyNiveles - 2, ")")
If lProceso = 0 Then
    If gSQLACC = cyAccess Then
        Set rs = gdbmidb.OpenRecordset("SELECT distinct a.iddestino,b.descripción FROM arcos a left join actividades b on a.idorigen=b.id WHERE idorigen=1", dbOpenSnapshot)
        Call tvArbol.Nodes.Add(, , "r001", rs(1))
    Else
        If adors.State > 0 Then adors.Close
        adors.Open "SELECT distinct a.iddestino,b.descripción FROM arcos a left join actividades b on a.idorigen=b.id WHERE idorigen=1", gConSql, adOpenStatic, adLockReadOnly
        Call tvArbol.Nodes.Add(, , "r001", adors(1))
    End If
    Call AgregaRama(tvArbol, "1", lProceso, False)
Else
    If gSQLACC = cyAccess Then
        Set rs = gdbmidb.OpenRecordset("SELECT distinct a.idorigen,b.descripción FROM arcos a left join actividades b on a.idorigen=b.id WHERE idorigen not in (select id from actividades where idpad=" + Str(lProceso) + s + ") and iddestino in (select id from actividades where idpad=" + Str(lProceso) + s + ")", dbOpenSnapshot)
        Do While Not rs.EOF
            If InStr(gs, Right("   " + Trim(Str(rs(0))), 4)) > 0 Then
                Call tvArbol.Nodes.Add(, , "r" + Right("00" + Trim(Str(rs(0))), 3) + "000", rs(1) + "(...)")
            Else
                gs = gs + Right("   " + Trim(Str(rs(0))), 4)
                Call tvArbol.Nodes.Add(, , "r" + Right("00" + Trim(Str(rs(0))), 3), rs(1))
                Call AgregaRama(tvArbol, rs(0), lProceso, False)
            End If
            rs.MoveNext
        Loop
    Else
        If adors.State > 0 Then adors.Close
        adors.Open "SELECT distinct a.idorigen,b.descripción FROM arcos a left join actividades b on a.idorigen=b.id WHERE idorigen not in (select id from actividades where idpad=" + Str(lProceso) + s + ") and iddestino in (select id from actividades where idpad=" + Str(lProceso) + s + ")", gConSql, adOpenStatic, adLockReadOnly
        Do While Not adors.EOF
            If InStr(gs, Right("   " + Trim(Str(adors(0))), 4)) > 0 Then
                Call tvArbol.Nodes.Add(, , "r" + Right("00" + Trim(Str(adors(0))), 3) + "000", adors(1) + "(...)")
            Else
                gs = gs + Right("   " + Trim(Str(adors(0))), 4)
                Call tvArbol.Nodes.Add(, , "r" + Right("00" + Trim(Str(adors(0))), 3), adors(1))
                Call AgregaRama(tvArbol, adors(0), lProceso, False)
            End If
            adors.MoveNext
        Loop
    End If
End If
tvArbol.LineStyle = tvwRootLines  ' Linestyle = 1
End Sub

'Muestra el arbol de actividades con los desenlaces asociados
'Miguel 19 de julio del 2002
Sub ArbolRedConDesenlaces(ByRef tvArbol As TreeView, lProceso As Long)
Dim rs As Recordset, s As String, y As Byte, adors As New ADODB.Recordset
tvArbol.Nodes.Clear
gs = ""
For y = 0 To gyNiveles - 3
    s = s + " or idpad in (select id from actividades where idpad=" + Trim(Str(lProceso))
Next
s = s + String(gyNiveles - 2, ")")
If lProceso = 0 Then
    If gSQLACC = cyAccess Then
        Set rs = gdbmidb.OpenRecordset("SELECT distinct a.iddestino,b.descripción,'a' as tipo FROM arcos a left join actividades b on a.idorigen=b.id WHERE idorigen=1", dbOpenSnapshot)
        Call tvArbol.Nodes.Add(, , "ra0001", rs(1))
    Else
        If adors.State > 0 Then adors.Close
        adors.Open "SELECT distinct a.iddestino,b.descripción FROM arcos a left join actividades b on a.idorigen=b.id WHERE idorigen=1", gConSql, adOpenStatic, adLockReadOnly
        Call tvArbol.Nodes.Add(, , "r0001", adors(1))
    End If
    Call AgregaRamaConDesenlaces(tvArbol, "1", lProceso, False, 1)
Else
    If gSQLACC = cyAccess Then
        Set rs = gdbmidb.OpenRecordset("SELECT distinct a.idorigen,b.descripción,'a' as tipo FROM arcos a left join actividades b on a.idorigen=b.id WHERE idorigen not in (select id from actividades where idpad=" + Str(lProceso) + s + ") and iddestino in (select id from actividades where idpad=" + Str(lProceso) + s + ") union all select id,descripción,'d' as tipo from desenlaces where id in (select iddes from relacióndesenlaceactividad where idact in (select id from actividades where idpad=" + Str(lProceso) + s + "))", dbOpenSnapshot)
        Do While Not rs.EOF
            If InStr(gs, Right(String(4, rs(2)) + Trim(Str(rs(0))), 4)) > 0 Then
                Call tvArbol.Nodes.Add(, , "r" + rs(2) + Right("000" + Trim(Str(rs(0))), 4) + "0000", rs(1) + "(...)")
            Else
                gs = gs + Right(String(4, rs(2)) + Trim(Str(rs(0))), 4)
                Call tvArbol.Nodes.Add(, , "r" + rs(2) + Right("000" + Trim(Str(rs(0))), 4), rs(1))
                Call AgregaRamaConDesenlaces(tvArbol, rs(0), lProceso, False, IIf(rs(2) = "a", 1, 2))
            End If
            rs.MoveNext
        Loop
    Else
        If adors.State > 0 Then adors.Close
        adors.Open "SELECT distinct a.idorigen,b.descripción,'a' as tipo FROM arcos a left join actividades b on a.idorigen=b.id WHERE idorigen not in (select id from actividades where idpad=" + Str(lProceso) + s + ") and iddestino in (select id from actividades where idpad=" + Str(lProceso) + s + ") union all select id,descripción,'d' as tipo from desenlaces where id in (select iddes from relacióndesenlaceactividad where idact in (select id from actividades where idpad=" + Str(lProceso) + s + "))", gConSql, adOpenStatic, adLockReadOnly
        Do While Not adors.EOF
            If InStr(gs, Right("   " + Trim(Str(adors(0))), 4)) > 0 Then
                Call tvArbol.Nodes.Add(, , "r" + rs(2) + Right("000" + Trim(Str(adors(0))), 4) + "0000", adors(1) + "(...)")
            Else
                gs = gs + Right("   " + Trim(Str(adors(0))), 4)
                Call tvArbol.Nodes.Add(, , "r" + Right("000" + Trim(Str(adors(0))), 4), adors(1))
                Call AgregaRamaConDesenlaces(tvArbol, adors(0), lProceso, False, IIf(adors(2) = "a", 1, 2))
            End If
            adors.MoveNext
        Loop
    End If
End If
tvArbol.LineStyle = tvwRootLines  ' Linestyle = 1
End Sub


Sub AgregaRama(ByRef tvArbol As TreeView, nNodo As Integer, lProceso As Long, bRaiz As Boolean)
Dim rs As Recordset, s As String, yNiveles As Byte, y As Byte
Dim adors As New ADODB.Recordset
'gyNiveles
For y = 0 To gyNiveles - 3
    s = s + " or idpad in (select id from actividades where idpad=" + Trim(Str(lProceso))
Next
s = s + String(gyNiveles - 2, ")")
If gSQLACC = cyAccess Then
    Set rs = gdbmidb.OpenRecordset("select distinct a.id,a.descripción+' ('+b.descripción+')' from ((actividades a left join actividades b on a.idpad=b.id) inner join Arcos c on a.id=c.iddestino) inner join actividades d on c.idorigen=d.id where c.idorigen=" + Str(nNodo) + " and ((a.idpad=" + Trim(Str(lProceso)) + IIf(gyNiveles > 2, " or a." + Mid(s, 5), "") + ") or (d.idpad=" + Trim(Str(lProceso)) + IIf(gyNiveles > 2, " or d." + Mid(s, 5), "") + "))", dbOpenSnapshot)
    gs = gs + Right("   " + Trim(Str(nNodo)), 4)
    Do While Not rs.EOF
        If InStr(gs, Right("   " + Trim(Str(rs(0))), 4)) > 0 Then
            If Not bRaiz Then
                Call tvArbol.Nodes.Add("r" + Right("00" + Trim(Str(nNodo)), 3), tvwChild, "r" + Right("00" + Trim(Str(nNodo)), 3) + Right("00" + Trim(Str(rs(0))), 3), rs(1) + " (...)")
            Else
                Call tvArbol.Nodes.Add(, , "r" + Right("00" + Trim(Str(nNodo)), 3) + Right("00" + Trim(Str(rs(0))), 3), rs(1) + " (...)")
            End If
        Else
            If Not bRaiz Then
                Call tvArbol.Nodes.Add("r" + Right("00" + Trim(Str(nNodo)), 3), tvwChild, "r" + Right("00" + Trim(Str(rs(0))), 3), rs(1))
            Else
                Call tvArbol.Nodes.Add(, , "r" + Right("00" + Trim(Str(rs(0))), 3), rs(1))
            End If
            Call AgregaRama(tvArbol, rs(0), lProceso, False)
        End If
        rs.MoveNext
    Loop
Else
    If adors.State > 0 Then adors.Close
    adors.Open "select distinct a.id,a.descripción+' ('+b.descripción+')' from ((actividades a left join actividades b on a.idpad=b.id) inner join Arcos c on a.id=c.iddestino) inner join actividades d on c.idorigen=d.id where c.idorigen=" + Str(nNodo) + " and ((a.idpad=" + Trim(Str(lProceso)) + IIf(gyNiveles > 2, " or a." + Mid(s, 5), "") + ") or (d.idpad=" + Trim(Str(lProceso)) + IIf(gyNiveles > 2, " or d." + Mid(s, 5), "") + "))", gConSql, adOpenStatic, adLockReadOnly
    gs = gs + Right("   " + Trim(Str(nNodo)), 4)
    Do While Not adors.EOF
        If InStr(gs, Right("   " + Trim(Str(adors(0))), 4)) > 0 Then
            If Not bRaiz Then
                Call tvArbol.Nodes.Add("r" + Right("00" + Trim(Str(nNodo)), 3), tvwChild, "r" + Right("00" + Trim(Str(nNodo)), 3) + Right("00" + Trim(Str(adors(0))), 3), adors(1) + " (...)")
            Else
                Call tvArbol.Nodes.Add(, , "r" + Right("00" + Trim(Str(nNodo)), 3) + Right("00" + Trim(Str(adors(0))), 3), adors(1) + " (...)")
            End If
        Else
            If Not bRaiz Then
                Call tvArbol.Nodes.Add("r" + Right("00" + Trim(Str(nNodo)), 3), tvwChild, "r" + Right("00" + Trim(Str(adors(0))), 3), adors(1))
            Else
                Call tvArbol.Nodes.Add(, , "r" + Right("00" + Trim(Str(adors(0))), 3), adors(1))
            End If
            Call AgregaRama(tvArbol, adors(0), lProceso, False)
        End If
        adors.MoveNext
    Loop
End If
End Sub


Sub AgregaRamaConDesenlaces(ByRef tvArbol As TreeView, nNodo As Integer, lProceso As Long, bRaiz As Boolean, yTipo As Byte)
Dim rs As Recordset, s As String, yNiveles As Byte, y As Byte
Dim adors As New ADODB.Recordset, bDesenlace As Boolean
'gyNiveles: variable global que indica el número de niveles en que se clasifican las actividades
For y = 0 To gyNiveles - 3
    s = s + " or idpad in (select id from actividades where idpad=" + Trim(Str(lProceso))
Next
s = s + String(gyNiveles - 2, ")")
If gSQLACC = cyAccess Then
    If yTipo = 1 Then
        'Set rs = gdbmidb.OpenRecordset("select distinct a.id,a.descripción+' ('+b.descripción+')' as nombre,'a' as tipo from ((actividades a left join actividades b on a.idpad=b.id) inner join Arcos c on a.id=c.iddestino) inner join actividades d on c.idorigen=d.id where c.idorigen=" + Str(nNodo) + " and ((a.idpad=" + Trim(Str(lProceso)) + IIf(gyNiveles > 2, " or a." + Mid(s, 5), "") + ") or (d.idpad=" + Trim(Str(lProceso)) + IIf(gyNiveles > 2, " or d." + Mid(s, 5), "") + ")) union all select distinct id,descripción as nombre,'d' as tipo from desenlaces where id in (select iddes from relaciónactividaddesenlace where idact=" & nNodo & ")", dbOpenSnapshot)
        Set rs = gdbmidb.OpenRecordset("select distinct a.id,a.descripción as nombre,'a' as tipo from ((actividades a left join actividades b on a.idpad=b.id) inner join Arcos c on a.id=c.iddestino) inner join actividades d on c.idorigen=d.id where c.idorigen=" + Str(nNodo) + " and ((a.idpad=" + Trim(Str(lProceso)) + IIf(gyNiveles > 2, " or a." + Mid(s, 5), "") + ") or (d.idpad=" + Trim(Str(lProceso)) + IIf(gyNiveles > 2, " or d." + Mid(s, 5), "") + ")) union all select distinct id,descripción as nombre,'d' as tipo from desenlaces where id in (select iddes from relaciónactividaddesenlace where idact=" & nNodo & ")", dbOpenSnapshot)
        bDesenlace = False
    Else
        'Set rs = gdbmidb.OpenRecordset("select distinct a.id,a.descripción+' ('+b.descripción+')' as nombre,'a' as tipo from (actividades a left join actividades b on a.idpad=b.id) inner join relacióndesenlaceactividad c on a.id=c.idact where c.iddes=" + Str(nNodo) + " and (a.idpad=" + Trim(Str(lProceso)) + IIf(gyNiveles > 2, " or a." + Mid(s, 5), "") + ")", dbOpenSnapshot)
        Set rs = gdbmidb.OpenRecordset("select distinct a.id,a.descripción as nombre,'a' as tipo from (actividades a left join actividades b on a.idpad=b.id) inner join relacióndesenlaceactividad c on a.id=c.idact where c.iddes=" + Str(nNodo) + " and (a.idpad=" + Trim(Str(lProceso)) + IIf(gyNiveles > 2, " or a." + Mid(s, 5), "") + ")", dbOpenSnapshot)
        bDesenlace = True
    End If
    gs = gs + Right(String(5, IIf(yTipo = 1, "a", "d")) + Trim(Str(nNodo)), 5)
    Do While Not rs.EOF
        If InStr(gs, Right(String(5, rs(2)) + Trim(Str(rs(0))), 5)) > 0 Then
            If Not bRaiz Then
                Call tvArbol.Nodes.Add("r" + IIf(yTipo = 1, "a", "d") + Right("000" + Trim(Str(nNodo)), 4), tvwChild, "r" + rs(2) + Right("000" + Trim(Str(nNodo)), 4) + Right("000" + Trim(Str(rs(0))), 4), rs(1) + IIf(rs(2) = "a", " (...)", ""))
            Else
                Call tvArbol.Nodes.Add(, , "r" + rs(2) + Right("000" + Trim(Str(nNodo)), 4) + Right("000" + Trim(Str(rs(0))), 4), rs(1) + IIf(rs(2) = "a", " (...)", ""))
            End If
        Else
            If Not bRaiz Then
                Call tvArbol.Nodes.Add("r" + IIf(yTipo = 1, "a", "d") + Right("000" + Trim(Str(nNodo)), 4), tvwChild, "r" + rs(2) + Right("000" + Trim(Str(rs(0))), 4), rs(1))
            Else
                Call tvArbol.Nodes.Add(, , "r" + rs(2) + Right("000" + Trim(Str(rs(0))), 4), rs(1))
            End If
            Call AgregaRamaConDesenlaces(tvArbol, rs(0), lProceso, False, IIf(rs(2) = "a", 1, 2))
        End If
        rs.MoveNext
    Loop
Else
    If adors.State > 0 Then adors.Close
    If yTipo = 1 Then
        adors.Open "select distinct a.id,a.descripción+' ('+b.descripción+')' as nombre,'a' as tipo from ((actividades a left join actividades b on a.idpad=b.id) inner join Arcos c on a.id=c.iddestino) inner join actividades d on c.idorigen=d.id where c.idorigen=" + Str(nNodo) + " and ((a.idpad=" + Trim(Str(lProceso)) + IIf(gyNiveles > 2, " or a." + Mid(s, 5), "") + ") or (d.idpad=" + Trim(Str(lProceso)) + IIf(gyNiveles > 2, " or d." + Mid(s, 5), "") + ")) union all select distinct id,descripción as nombre,'d' as tipo from desenlaces where id in (select iddes from relaciónactividaddesenlace where idact=" & nNodo & ")", gConSql, adOpenStatic, adLockReadOnly
    Else
        adors.Open "select distinct a.id,a.descripción+' ('+b.descripción+')' as nombre,'a' as tipo from (actividades a left join actividades b on a.idpad=b.id) inner join relacióndesenlaceactividad c on a.id=c.idact where c.iddes=" + Str(nNodo) + " and (a.idpad=" + Trim(Str(lProceso)) + IIf(gyNiveles > 2, " or a." + Mid(s, 5), "") + ")", gConSql, adOpenStatic, adLockReadOnly
    End If
    gs = gs + Right("   " + Trim(Str(nNodo)), 5)
    Do While Not adors.EOF
        If InStr(gs, Right("   " + Trim(Str(adors(0))), 5)) > 0 Then
            If Not bRaiz Then
                Call tvArbol.Nodes.Add("r" + Right("000" + Trim(Str(nNodo)), 4), tvwChild, "r" + Right("000" + Trim(Str(nNodo)), 4) + Right("000" + Trim(Str(adors(0))), 4), adors(1) + " (...)")
            Else
                Call tvArbol.Nodes.Add(, , "r" + Right("000" + Trim(Str(nNodo)), 4) + Right("000" + Trim(Str(adors(0))), 4), adors(1) + " (...)")
            End If
        Else
            If Not bRaiz Then
                Call tvArbol.Nodes.Add("r" + Right("000" + Trim(Str(nNodo)), 4), tvwChild, "r" + Right("000" + Trim(Str(adors(0))), 4), adors(1))
            Else
                Call tvArbol.Nodes.Add(, , "r" + Right("000" + Trim(Str(adors(0))), 4), adors(1))
            End If
            Call AgregaRamaConDesenlaces(tvArbol, adors(0), lProceso, False, IIf(adors(2) = "a", 1, 2))
        End If
        adors.MoveNext
    Loop
End If
End Sub


Sub TransfiereElementosListas(ByRef ListaOrigen As ListBox, ByRef ListaDestino As ListBox, Optional bTodos As Boolean)
Dim y As Byte
y = 0
Do While y < ListaOrigen.ListCount
    If ListaOrigen.Selected(y) Or bTodos Then
        ListaDestino.AddItem ListaOrigen.List(y)
        ListaDestino.ItemData(ListaDestino.NewIndex) = ListaOrigen.ItemData(y)
        ListaOrigen.RemoveItem (y)
    Else
        y = y + 1
    End If
Loop
End Sub

Sub QuitaMemoriaForma(ByVal sForma As String, Optional yNoFormas As Byte)
Dim i As Long, y As Byte
For i = Forms.Count - 1 To 0 Step -1
    If Forms(i).Name = sForma Then
        Unload Forms(i)
        y = y + 1
        If y >= yNoFormas Then Exit Sub
    End If
Next
End Sub


Sub CalculaTop(ByRef Ctl As Control, ByRef lTop As Long, ByRef lLeft As Long, ByRef lmax As Long)
Dim y As Byte
On Error GoTo continua:
lTop = Ctl.Top + Ctl.Height
lLeft = Ctl.Left
If Mid(Ctl.Container.Name, 1, 3) <> "frm" Then
    lTop = lTop + Ctl.Container.Top
    lLeft = lLeft + Ctl.Container.Left
End If
lmax = Ctl.Container.Width
If Mid(Ctl.Container.Container.Name, 1, 3) <> "frm" Then
    lTop = lTop + Ctl.Container.Container.Top
    lLeft = lLeft + Ctl.Container.Container.Left
Else
    lmax = Ctl.Container.Container.Width
End If
If Mid(Ctl.Container.Container.Container.Name, 1, 3) <> "frm" Then
    lTop = lTop + Ctl.Container.Container.Container.Top
    lLeft = lLeft + Ctl.Container.Container.Container.Left
Else
    lmax = Ctl.Container.Container.Width
End If
If Mid(Ctl.Container.Container.Container.Container.Name, 1, 3) = "frm" Then
    lmax = Ctl.Container.Container.Width
End If
continua:
End Sub

'Carga elementos de query en el árbol igual que la función anterior
'La única diferencia esta en la clave que se compone por 10 caracteres en cada nivel.
Sub CargaDatosArbolVariosNiveles10(ByRef tvArbol As TreeView, sQuery As String, yNoNiveles As Integer, Optional bNoBarraProgreso As Boolean, Optional db As DAO.Database, Optional bAbierto As Boolean, Optional bNoBoxes As Boolean)
Dim nPro(2) As Integer, y As Byte, n As Integer, s As String, sCve(2) As String, bMov As Boolean, yErr As Byte
Dim ynivel As Integer, iNodos As Integer, iValor() As Long, l As Long, adors As New ADODB.Recordset
Dim sCvePadre As String, bSio As Boolean
Dim rs As Recordset
On Error GoTo ErrorCargaDatos:
If bNoBoxes Then
    tvArbol.Checkboxes = False
End If
sQuery = LCase(sQuery)
s = gsNoCatálogosFijos
If yNoNiveles = 1 Then bNoBarraProgreso = True
If gbSIO_mdb And IsNull(db) Then
    Do While InStr(s, ",")
        If InStr(sQuery, Mid(s, 1, InStr(s, ",") - 1)) Then
            Exit Do
        End If
        s = Mid(s, InStr(s, ",") + 1)
    Loop
    bSio = (InStr(s, ",") = 0)
End If
l = Timer
If gSQLACC = cyAccess Or bSio Then
    If Len(db.Name) = 0 Then
       Set db = gdbmidb
    End If
    Set rs = db.OpenRecordset(sQuery, dbOpenSnapshot)
    yNoNiveles = Int(rs.Fields.Count / 2)
    tvArbol.Visible = False
    tvArbol.Nodes.Clear
    gs = ""
    ReDim iValor(yNoNiveles - 1)
    If Not rs.EOF Then
        rs.MoveLast
        l = rs.RecordCount
        n = Int(l / 20)
        rs.MoveFirst
        If n > 0 And Not bNoBarraProgreso Then
            RHBarraProgreso.Show
            RHBarraProgreso.ProgressBar1.Value = 0
            RHBarraProgreso.Refresh
        End If
    End If
    Do While Not rs.EOF
        sCvePadre = "r"
        For ynivel = 0 To yNoNiveles - 1
            If IsNull(rs(ynivel)) Then Exit For
            sCvePadre = sCvePadre + Right(String(9, "0") & rs(ynivel), 10)
            If iValor(ynivel) <> rs(ynivel) Then
                iValor(ynivel) = rs(ynivel)
                For y = ynivel + 1 To yNoNiveles - 1
                    iValor(y) = 0
                Next
                If ynivel = 0 Then
                    Call tvArbol.Nodes.Add(, , sCvePadre, rs(yNoNiveles + ynivel))
                Else
    'AgregaExistente:
                    Call tvArbol.Nodes.Add(Mid(sCvePadre, 1, Len(sCvePadre) - 10), tvwChild, sCvePadre, rs(yNoNiveles + ynivel))
                End If
            End If
        Next
        rs.MoveNext
        If n > 0 And Not bBarraProgreso Then
            If rs.AbsolutePosition Mod n = 0 Then
                RHBarraProgreso.ProgressBar1.Value = 100 * (rs.AbsolutePosition + 1) / l
                RHBarraProgreso.Refresh
            End If
        End If
        'Debug.Print Str(rs.AbsolutePosition)
    Loop
Else
    adors.Open sQuery, gConSql, adOpenStatic, adLockReadOnly
    yNoNiveles = Int(adors.Fields.Count / 2)
    tvArbol.Visible = False
    tvArbol.Nodes.Clear
    gs = ""
    ReDim iValor(yNoNiveles - 1)
    If Not adors.EOF Then
        'adors.MoveLast
        l = adors.RecordCount
        n = Int(l / 20)
        'adors.MoveFirst
'        If n > 0 And Not bNoBarraProgreso Then
'            RHBarraProgreso.Show
'            RHBarraProgreso.ProgressBar1.Value = 0
'            RHBarraProgreso.Refresh
'        End If
    End If
    Do While Not adors.EOF
        sCvePadre = "r"
        For ynivel = 0 To yNoNiveles - 1
            If IsNull(adors(ynivel)) Then Exit For
            sCvePadre = sCvePadre + Right(String(9, "0") & adors(ynivel), 10)
            If iValor(ynivel) <> adors(ynivel) Then
                iValor(ynivel) = adors(ynivel)
                For y = ynivel + 1 To yNoNiveles - 1
                    iValor(y) = 0
                Next
                If ynivel = 0 Then
                    Call tvArbol.Nodes.Add(, , sCvePadre, adors(yNoNiveles + ynivel))
                Else
    'AgregaExistente:
                    Call tvArbol.Nodes.Add(Mid(sCvePadre, 1, Len(sCvePadre) - 10), tvwChild, sCvePadre, adors(yNoNiveles + ynivel))
                End If
            End If
        Next
        adors.MoveNext
        If n > 0 And Not bBarraProgreso Then
            If adors.AbsolutePosition Mod n = 0 And Not adors.EOF Then
                RHBarraProgreso.ProgressBar1.Value = 100 * (adors.AbsolutePosition + 1) / (l + 1)
                RHBarraProgreso.Refresh
            End If
        End If
        'Debug.Print Str(adors.AbsolutePosition)
    Loop
End If
If bAbierto Then
    For y = 1 To tvArbol.Nodes.Count
        tvArbol.Nodes(y).Expanded = True
    Next
End If
RHBarraProgreso.ProgressBar1.Value = 100
RHBarraProgreso.Refresh
Unload RHBarraProgreso
tvArbol.LineStyle = tvwRootLines  ' Linestyle = 1
tvArbol.Visible = True
Exit Sub
ErrorCargaDatos:
If Err.Number = 91 And yErr = 0 Then
    yErr = 1
    Resume Next
ElseIf Err.Number = 35602 Or Err.Number = -1 Then
    Resume Next
End If
sError = "Error: " + Err.Description
y = MsgBox(sError, vbAbortRetryIgnore + vbCritical, "Error no esperado (" + Str(Err.Number) + ")")
If y = vbCancel Then
    Exit Sub
ElseIf y = vbRetry Then
    Resume
ElseIf y = vbIgnore Then
    Resume Next
End If
End Sub



'Obtiene los datios del cuestionario capturados y de vuelve la cadena con la sig. estructura
'AsuntoInstitución->Cuestionario->Preguntas->Respuestas
'Miguel 17 de Junio
Function ObtieneDatosCapturados(ByVal lAvance As Long, iCuestionario As Integer, Optional bTexto As Boolean) As String
Dim rs As Recordset, sPrg As String, sDatos As String, i As Integer
Dim adors As New ADODB.Recordset
If gSQLACC = cyAccess Then
    If bTexto Then
        Set rs = gdbmidb.OpenRecordset("select * from respuestascuestionariotexto where idAvance=" & lAvance & " and idcuestionario=" & iCuestionario & " order by idpregunta", dbOpenSnapshot)
        If Not rs.EOF Then
            i = rs!idcuestionario
            ObtieneDatosCapturados = "Prg:°" & rs!idpregunta & gsSeparador & rs!Respuesta & gsSeparador
        End If
        Exit Function
    End If
    Set rs = gdbmidb.OpenRecordset("select * from respuestascuestionario where idAvance=" & lAvance & " order by idcuestionario,idpregunta", dbOpenSnapshot)
    Do While Not rs.EOF
        i = rs!idcuestionario
        sDatos = sDatos & " Cuestionario:°" & i & ", "
        Do While rs!idcuestionario = i
            sPrg = rs!idpregunta
            sDatos = sDatos & "Prg:°" & rs!idpregunta & ","
            Do While rs!idpregunta = sPrg
                sDatos = sDatos & rs!idrespuesta & ","
                rs.MoveNext
                If rs.EOF Then Exit Do
            Loop
            If rs.EOF Then Exit Do
        Loop
        sDatos = sDatos & ObtieneDatosCapturados(lAvance, i, True)
    Loop
Else
    If bTexto Then
        adors.Open "select * from respuestascuestionariotexto where idAvance=" & lAvance & " and idcuestionario=" & iCuestionario & " order by idpregunta", gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            i = adors!idcuestionario
            ObtieneDatosCapturados = "Prg:°" & adors!idpregunta & gsSeparador & adors!Respuesta & gsSeparador
        End If
        Exit Function
    End If
    adors.Open "select * from respuestascuestionario where idAvance=" & lAvance & " order by idcuestionario,idpregunta", gConSql, adOpenStatic, adLockReadOnly
    Do While Not adors.EOF
        i = adors!idcuestionario
        sDatos = sDatos & ", Cuestionario:°" & i & ", "
        Do While adors!idcuestionario = i
            sPrg = adors!idpregunta
            sDatos = sDatos & "Prg:°" & adors!idpregunta & ","
            Do While adors!idpregunta = sPrg
                sDatos = sDatos & adors!idrespuesta & ","
                adors.MoveNext
                If adors.EOF Then Exit Do
            Loop
            If adors.EOF Then Exit Do
        Loop
        sDatos = sDatos & ObtieneDatosCapturados(lAvance, i, True)
    Loop
End If
ObtieneDatosCapturados = sDatos
End Function


Function ObtenSM(dFecha As Date) As Currency
Dim rs As DAO.Recordset, adors As New ADODB.Recordset
If gSQLACC = cyAccess Then
    Set rs = gdbmidb.OpenRecordset("select monto from salmin where fecha<=cdate('" & Format(dFecha, gsFormatoFecha) & "') order by fecha desc", dbOpenSnapshot)
    If Not rs.EOF Then
        ObtenSM = rs(0)
    Else
        Call MsgBox("No se encuentra ningún Salario mínimo vigente a la fecha especificada (" & Format(dFecha, gsFormatoFecha) & ")", vbCritical + vbOKOnly, "")
    End If
Else
    adors.Open "select monto from salmin where fecha<=convert(datetime,'" & Format(dFecha, "dd-mm-yyyy") & "',105) order by fecha desc", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        ObtenSM = adors(0)
    Else
        Call MsgBox("No se encuentra ningún Salario mínimo vigente a la fecha especificada (" & Format(dFecha, gsFormatoFecha) & ")", vbCritical + vbOKOnly, "")
    End If
End If
End Function


Function ObtieneRefArchivo(lAsunto As Long) As String
Dim rs As DAO.Recordset, adors As ADODB.Recordset
Dim sArchivo As String
On Error GoTo Salir:
If gSQLACC = cyAccess Then
    Set rs = gdbmidb.OpenRecordset("select n.año & '_' & right('0'& n.iddel,3) & '_' & n.consecutivo as folio from nominales n where n.idasu=" & lAsunto, dbOpenSnapshot)
    If rs.EOF Then
        ObtieneRefArchivo = "z:\DocsSIO\"
    Else
        ObtieneRefArchivo = "z:\DocsSIO\" & rs(0) & "\"
    End If
Else
    Set adors = New ADODB.Recordset
    adors.Open "select convert(nvarchar,n.año) + '_' + right('0'+ convert(nvarchar,n.iddel),3) + '_' + convert(nvarchar,n.consecutivo) as folio from nominales n where n.idasu=" & lAsunto, gConSql, adOpenStatic, adLockReadOnly
    If adors.EOF Then
        ObtieneRefArchivo = Mid(gsDirDocumentos, 1, InStrRev(Mid(gsDirDocumentos, 1, Len(gsDirDocumentos) - 1), "\")) & "DocsSIO\"
    Else
        ObtieneRefArchivo = Mid(gsDirDocumentos, 1, InStrRev(Mid(gsDirDocumentos, 1, Len(gsDirDocumentos) - 1), "\")) & "DocsSIO\" & adors(0) & "\"
    End If
End If
Exit Function
Salir:
End Function




Function sEdad(sFecha As String) As String
Dim dFecha As Date
sEdad = ""
If IsDate(sFecha) Then
    dFecha = CDate(sFecha)
    sEdad = Str(DateDiff("yyyy", dFecha, Date) - IIf(Date < DateAdd("yyyy", DateDiff("yyyy", dFecha, Date), dFecha), 1, 0)) + " Años "
    dFecha = DateAdd("yyyy", DateDiff("yyyy", dFecha, Date) - IIf(Date < DateAdd("yyyy", DateDiff("yyyy", dFecha, Date), dFecha), 1, 0), dFecha)
    sEdad = sEdad + Str(DateDiff("m", dFecha, Date) - IIf(Date < DateAdd("m", DateDiff("m", dFecha, Date), dFecha), 1, 0)) + " Meses"
    'dFecha = DateAdd("m", DateDiff("m", dFecha, Date) - IIf(Date < DateAdd("m", DateDiff("m", dFecha, Date), dFecha), 1, 0), dFecha)
    'sEdad = sEdad + " Dias:" + Str(DateDiff("d", dFecha, Date))
End If
End Function

