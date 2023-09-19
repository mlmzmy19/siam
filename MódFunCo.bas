Attribute VB_Name = "Mód_FunFormCond"
Const cyAccess = 1

Function BuscaComboClave(ByRef Combo As ComboBox, ByVal sBus As String, bCve As Boolean, Optional bLike As Boolean) As Long
Dim Y As Integer, b As Boolean, s As String
If IsNull(bLike) Then
    b = False
Else
    b = bLike
End If
BuscaComboClave = -1
If InStrRev(Combo.List(Y), "(") = 0 Then Exit Function
If Not bCve Then
    For Y = 0 To Combo.ListCount - 1
        s = Mid(Combo.List(Y), InStrRev(Combo.List(Y), "(") + 1)
        s = Mid(s, 1, Len(s) - 1)
        If bLike Then
            If s Like sBus + "*" Then
                BuscaComboClave = Y
                Exit Function
            End If
        Else
            If s = sBus Then
                BuscaComboClave = Y
                Exit Function
            End If
        End If
    Next
Else
    For Y = 0 To Combo.ListCount - 1
        If Combo.ItemData(Y) = Val(sBus) Then
            BuscaComboClave = Y
            Exit Function
        End If
    Next
End If
End Function

Function BuscaCombo(ByRef Combo As ComboBox, ByVal sBus As String, bCve As Boolean, Optional bLike As Boolean, Optional bClave As Boolean, Optional iIndex As Integer) As Long
Dim Y As Integer, b As Boolean
If IsNull(bLike) Then
    b = False
Else
    b = bLike
End If
If IsNull(iIndex) Then
    iIndex = 0
End If
BuscaCombo = -1
If Not bCve Then
    If bLike Then
        For Y = iIndex To Combo.ListCount - 1
            If UCase(Combo.List(Y)) Like "*" + UCase(sBus) + "*" Then
                BuscaCombo = Y
                Exit Function
            End If
        Next
    Else
        If bClave And InStrRev(Combo.List(Y), "(") > 0 And InStrRev(sBus, "(") > 0 Then
            For Y = iIndex To Combo.ListCount - 1
                If Mid(Combo.List(Y), 1, InStrRev(Combo.List(Y), "(")) = Mid(sBus, 1, InStrRev(sBus, "(")) Then
                    BuscaCombo = Y
                    Exit Function
                End If
            Next
        Else
            For Y = iIndex To Combo.ListCount - 1
                If Combo.List(Y) = sBus Then
                    BuscaCombo = Y
                    Exit Function
                End If
            Next
        End If
    End If
Else
    For Y = iIndex To Combo.ListCount - 1
        If Combo.ItemData(Y) = Val(sBus) Then
            BuscaCombo = Y
            Exit Function
        End If
    Next
End If
End Function

Function BuscaList(ByRef List As ListBox, ByVal sBus As String, bCve As Boolean, Optional bLike As Boolean) As Long
Dim Y As Integer, b As Boolean
If IsNull(bLike) Then
    b = False
Else
    b = bLike
End If
BuscaList = -1
If Not bCve Then
    For Y = 0 To List.ListCount - 1
        If bLike Then
            If List.List(Y) Like sBus + "*" Then
                BuscaList = Y
                Exit Function
            End If
        Else
            If List.List(Y) = sBus Then
                BuscaList = Y
                Exit Function
            End If
        End If
    Next
Else
    For Y = 0 To List.ListCount - 1
        If List.ItemData(Y) = Val(sBus) Then
            BuscaList = Y
            Exit Function
        End If
    Next
End If
End Function

'Pone en el arreglo los valores 0 ó 1
'1: Debe el campo conservar su valor al Limpiar
Sub AsignaValor(ByRef yDatos() As Byte, yTabla As Byte, bEntrada As Boolean)
Dim yMax As Byte, Y As Byte, l As Double
yMax = UBound(yDatos) + 1
If bEntrada Then  'Load de la forma
    'Set rs = gdbMidb.OpenRecordset("select * from nolimpiar where tabla=" + Str(yTabla), dbOpenDynaset)
    Set rs = gdbmiconfig.OpenRecordset("select * from nolimpiar where tabla=" + Str(yTabla) + " and idusi=" + Str(gs_usuario), dbOpenDynaset)  '****************************jps
    If Not rs.EOF Then
        l = rs(1)
        For Y = 0 To yMax - 1
            If l >= (2 ^ (yMax - Y)) Then
                l = l - (2 ^ (yMax - Y))
                yDatos(Y) = 1
            End If
        Next
    End If
Else  'UnLoad de la forma
    l = 0
    For Y = 0 To yMax - 1
        If yDatos(Y) = 1 Then
            l = l + (2 ^ (yMax - Y))
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
                Combo.AddItem rsSQLtabla!descripción + " (" + Trim(Str(rsSQLtabla!ID)) + ")"
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
    If sTabla Like "{*}" Then 'sólo adelante
        rsSQLCOMBO.Open sTabla, gConSql, adOpenForwardOnly, adLockReadOnly
    Else
        rsSQLCOMBO.Open sTabla, gConSql, adOpenStatic, adLockReadOnly
    End If
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


Sub LlenaComboCursor(ByRef Combo As ComboBox, ByRef adors As ADODB.Recordset, Optional bNoBorraCombo As Boolean)
If Not bNoBorraCombo Then Combo.Clear
Do While Not adors.EOF
    Combo.AddItem IIf(IsNull(adors(1)), "---", adors(1))
    Combo.ItemData(Combo.NewIndex) = adors(0)
    adors.MoveNext
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
                s = " (" + sCampo + " is null " + IIf(bCrystal, ")", "or length(rtrim(" + sCampo + "))=0)")
            End If
        ElseIf sCon = "*" Then
            If gSQLACC = cyAccess Then
                s = " not (isnull(" + sCampo + ") " + IIf(bCrystal, ")", "or len(rtrim(" + sCampo + "))=0)")
            Else  'SQL
                s = " not (" + sCampo + " is null " + IIf(bCrystal, ")", "or length(rtrim(" + sCampo + "))=0)")
            End If
        Else
            'sCon = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(LCase(sCon), "á", "a"), "é", "e"), "í", "i"), "ó", "o"), "ú", "u"), "à", "a"), "è", "e"), "ì", "i"), "ò", "o"), "ù", "u")
            'sCon = Replace(Replace(Replace(Replace(Replace(sCon, "a", "[aá]"), "e", "[eé]"), "i", "[ií]"), "o", "[oó]"), "u", "[uúü]")
            If gSQLACC = cyAccess Then
                s = sCampo + " like '" + Trim(sCon) + "'"
            Else
                sCon = Replace(Replace(sCon, "_", "[_]"), "%", "[%]")
                sCon = Replace(Replace(sCon, "*", "%"), "?", "_")
                s = "lower(" & sCampo & ") like '" + Replace(LCase(sCon), "'", "''") + "'"
            End If
            's = sCampo + " like '" + Trim(sCon) + "'"
        End If
    Else
        If (yFormaAct_0_Bus_1_Ins_2 = 2 Or yFormaAct_0_Bus_1_Ins_2 = 0) And Len(Trim(sCon)) = 0 Then
            s = s + "null"
        Else
            If yFormaAct_0_Bus_1_Ins_2 = 1 And Len(Trim(sCon)) > 0 Then
                If InStr(sCon, "*") > 0 Then
                    'sCon = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(LCase(sCon), "á", "a"), "é", "e"), "í", "i"), "ó", "o"), "ú", "u"), "à", "a"), "è", "e"), "ì", "i"), "ò", "o"), "ù", "u")
                    'sCon = Replace(Replace(Replace(Replace(Replace(sCon, "a", "[aá]"), "e", "[eé]"), "i", "[ií]"), "o", "[oó]"), "u", "[uúü]")
                    s = "lower(" & sCampo & ") like '" & Replace(LCase(sCon), "'", "''") & "'"
                Else
                    s = sCampo + "='" + Replace(sCon, "'", "''") + "'"
                End If
            Else
                s = s + "'" + Replace(sCon, "'", "''") + "'"
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
                    's = "convert(varchar," + sCampo + ",102)" + ss + "'" + Format(CDate(sCon), "yyyy.mm.dd") + "'"
                    s = "to_char(" + sCampo + ",'yyyymmdd')" + ss + "'" + Format(CDate(sCon), "yyyymmdd") + "'"
                End If
            Else
                If gSQLACC = cyAccess Then
                    s = "format(" + sCampo + ",'yyyy/mm/dd hh:mm:ss')" + ss + "'" + Format(CDate(sCon), "yyyy/mm/dd hh:mm:ss") + "'"
                Else
                    's = "convert(varchar," + sCampo + ",120)" + ss + "'" + Format(CDate(sCon), "yyyy-mm-dd hh:mm:ss") + "'"
                    s = "to_char(" + sCampo + ",'yyyymmdd hh24:mi:ss')" + ss + "'" + Format(CDate(sCon), "yyyymmdd hh:mm:ss") + "'"
                End If
            End If
        ElseIf IsDate(sCon) Then
            If gSQLACC = cyAccess Then
                s = s + "cdate('" + sCon + "')"
            Else
                's = s & "convert(datetime,'" & Format(CDate(sCon), "dd-mm-yyyy hh:mm:ss") & "',105)"
                s = s & "to_date('" & Format(CDate(sCon), "ddmmyyyy hh:mm:ss") & "','ddmmyyyy hh24:mi:ss')"
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
                    's = "convert(varchar," + sCampo + ",120) like '" + sCon + "'"
                    s = "to_char(" + sCampo + ",'dd/mon/yyyy hh24:mi:ss') like '" + sCon + "'"
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
                    If gSQLACC = cyAccess Then
                        s = sCampo + "  between cdate('" + Format(CDate(sCon), "dd/mm/yyyy") + " 00:00:00.000') and cdate('" + Format(CDate(sCon), "dd/mm/yyyy") + " 23:59:59.999')"
                    Else 'Oracle
                        s = sCampo + "  between to_date('" + Format(CDate(sCon), "dd/mm/yyyy") & "','dd/mm/yyyy') and to_date('" + Format(CDate(sCon), "dd/mm/yyyy") + " 23:59:59','dd/mm/yyyy hh24:mi:ss')"
                    End If
                Else
                    s = sCampo + ss + "'" + Format(CDate(sCon), "dd/mm/yyyy") + IIf(InStr(s, ">"), " 00:00:00.000'", " 23:59:59.999'")
                End If
            End If
        ElseIf IsDate(sCon) Then
            If gSQLACC = cyAccess Then
                s = s + "cdate('" + sCon + "')"
            Else
                's = s & "convert(datetime,'" & Format(CDate(sCon), "dd-mm-yyyy") & "',105)"
                s = s & "to_date('" & Format(CDate(sCon), "dd-mm-yyyy") & "','dd-mm-yyyy')"
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
                    's = "convert(varchar," + sCampo + ",102) like '" + sCon + "'"
                    s = "to_char(" + sCampo + ",'dd/mmm/yyyy') like '" + sCon + "'"
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
                    's = "convert(varchar," + sCampo + ",8)" + ss + "'" + Format(CDate(sCon), "hh:mm:ss") + "'"
                    s = "to_char(" + sCampo + ",'hh24:mi:ss')" + ss + "'" + Format(CDate(sCon), "hh:mm:ss") + "'"
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
                    's = "convert(varchar," + sCampo + ",8) like '" + sCon + "'"
                    s = "to_char(" + sCampo + ",'hh24:mi:ss') like '" + sCon + "'"
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

Sub CargaDatosArbolVariosNiveles(ByRef tvArbol As TreeView, sQuery As String, yNoNiveles As Byte, Optional bNoBarraProgreso As Boolean, Optional bAbierto As Boolean, Optional bNoMinus As Boolean)
Dim nPro(2) As Integer, Y As Byte, n As Integer, s As String, sCve(2) As String, bMov As Boolean, yErr As Byte
Dim ynivel As Byte, iNodos As Integer, iValor() As Integer, l As Long
Dim sCvePadre As String, bSio As Boolean, i As Integer
Dim rs As Recordset
On Error GoTo ErrorCargaDatos:
If Not bNoMinus Then
    sQuery = LCase(sQuery)
End If
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
            BarraProgreso.Show
            BarraProgreso.ProgressBar1.Value = 0
            BarraProgreso.Refresh
        End If
    End If
    Do While Not rs.EOF
        sCvePadre = "r"
        For ynivel = 0 To yNoNiveles - 1
            If IsNull(rs(ynivel)) Then Exit For
            sCvePadre = sCvePadre + Right("000" + Trim(Str(rs(ynivel))), 4)
            If iValor(ynivel) <> rs(ynivel) Then
                iValor(ynivel) = rs(ynivel)
                For Y = ynivel + 1 To yNoNiveles - 1
                    iValor(Y) = 0
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
                BarraProgreso.ProgressBar1.Value = 100 * (rs.AbsolutePosition + 1) / l
                BarraProgreso.Refresh
            End If
        End If
        'Debug.Print Str(rs.AbsolutePosition)
    Loop
Else
    Dim rsSQL1 As New ADODB.Recordset
    'rsSQL1.CursorLocation = adUseClient
    'rsSQL1.CursorLocation = adUseServer
    rsSQL1.Open sQuery, gConSql, adOpenStatic, adLockReadOnly, adCmdText
    tvArbol.Visible = False
    tvArbol.Nodes.Clear
    
    ReDim iValor(yNoNiveles - 1)

    If Not rsSQL1.EOF Then
        rsSQL1.MoveLast
        l = rsSQL1.RecordCount
        n = Int(l / 20)
        rsSQL1.MoveFirst
        If n > 0 And Not bNoBarraProgreso Then
            BarraProgreso.Show
            BarraProgreso.ProgressBar1.Value = 0
            BarraProgreso.Refresh
        End If
    End If
    Do While Not rsSQL1.EOF
        sCvePadre = "r"
        For ynivel = 0 To yNoNiveles - 1
            If IsNull(rsSQL1(Val(ynivel))) Then Exit For
            sCvePadre = sCvePadre + Right("000" + Trim(Str(rsSQL1(Val(ynivel)))), 4)
            If iValor(ynivel) <> rsSQL1(Val(ynivel)) Then
                iValor(ynivel) = rsSQL1(Val(ynivel))
                For Y = ynivel + 1 To yNoNiveles - 1
                    iValor(Y) = 0
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
                    BarraProgreso.ProgressBar1.Value = 100 * (rsSQL1.Bookmark) / l
                    BarraProgreso.Refresh
                End If
            End If
        End If
    Loop
    rsSQL1.Close
End If
BarraProgreso.ProgressBar1.Value = 100
BarraProgreso.Refresh
Unload BarraProgreso
If bAbierto Then
    For Y = 1 To tvArbol.Nodes.Count
        tvArbol.Nodes(Y).Expanded = True
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
Y = MsgBox(sError, vbAbortRetryIgnore + vbCritical, "Error no esperado (" + Str(Err.Number) + ")")
If Y = vbCancel Then
    Exit Sub
ElseIf Y = vbRetry Then
    Resume
ElseIf Y = vbIgnore Then
    Resume Next
End If
End Sub

Sub CargaDatosArbol(ByRef tvArbol As TreeView, sQuery As String, Optional bAbierto As Boolean, Optional ByVal lAsuIns As Long, Optional db As DAO.Database, Optional bOtroDB As Boolean)
Dim nPro(3) As Integer, Y As Integer, n As Integer, s As String, sCve(3) As String, bMov As Boolean
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
        For Y = 0 To 5
            sQuery = Replace(sQuery, Trim(Mid(ss, Y * 10 + 1, 10)) & "=1", Trim(Mid(ss, Y * 10 + 1, 10)) & "=-1")
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
                If lAsuIns > 0 Then
                    Set rs2 = db.OpenRecordset("select count(*) from avances where idasuins=" + Str(lAsuIns), dbOpenSnapshot)
                    If rs2(0) > 0 Then
                        'cambio último
                        'Set rs2 = db.OpenRecordset("select * from Arcos where iddestino=" + Str(nPro(1)) + " and idorigen in (select idact from avances where ultimo and idasuins=" + Str(lAsuIns) + ")", dbOpenSnapshot)
                        Set rs2 = db.OpenRecordset("select * from Arcos where iddestino=" + Str(nPro(1)) + " and idorigen in (select a.idact from avances a left join avances b on a.id=b.idant where b.id is null and a.idasuins=" + Str(lAsuIns) + ")", dbOpenSnapshot)
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
                        If lAsuIns > 0 Then
                            Set rs2 = db.OpenRecordset("select count(*) from avances where idasuins=" + Str(lAsuIns), dbOpenSnapshot)
                            If rs2(0) > 0 Then
                                'cambio último
                                'Set rs2 = db.OpenRecordset("select * from Arcos where iddestino=" + Str(nPro(1)) + " and idorigen in (select idact from avances where ultimo and idasuins=" + Str(lAsuIns) + ")", dbOpenSnapshot)
                                Set rs2 = db.OpenRecordset("select * from Arcos where iddestino=" + Str(nPro(1)) + " and idorigen in (select a.idact from avances a left join avances b on a.id=b.idant where b.id is null and a.idasuins=" + Str(lAsuIns) + ")", dbOpenSnapshot)
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
                If lAsuIns > 0 Then
                    rsSQL2.Open "select count(*) from avances where idasuins=" + Str(lAsuIns), gCadSQL, adOpenStatic, adLockReadOnly, adCmdText
                    If rsSQL2(0) > 0 Then
                        'cambio último
                        'rsSQL2.Open "select * from Arcos where iddestino=" + Str(nPro(1)) + " and idorigen in (select idact from avances where ultimo and idasuins=" + Str(lAsuIns) + ")", gCadSQL, adOpenStatic, adLockReadOnly, adCmdText
                        rsSQL2.Open "select * from Arcos where iddestino=" + Str(nPro(1)) + " and idorigen in (select a.idact from avances a left join avances b on a.id=b.idant where b.id is null and a.idasuins=" + Str(lAsuIns) + ")", gCadSQL, adOpenStatic, adLockReadOnly, adCmdText
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
                        If lAsuIns > 0 Then
                            rsSQL2.Open "select count(*) from avances where idasuins=" + Str(lAsuIns), gCadSQL, adOpenStatic, adLockReadOnly, adCmdText
                            If rsSQL2(0) > 0 Then
                                'cambio último
                                'rsSQL2.Open "select * from Arcos where iddestino=" + Str(nPro(1)) + " and idorigen in (select idact from avances where ultimo and idasuins=" + Str(lAsuIns) + ")", gCadSQL, adOpenStatic, adLockReadOnly, adCmdText
                                rsSQL2.Open "select * from Arcos where iddestino=" + Str(nPro(1)) + " and idorigen in (select a.idact from avances a left join avances b on a.id=b.idant where b.id is null and a.idasuins=" + Str(lAsuIns) + ")", gCadSQL, adOpenStatic, adLockReadOnly, adCmdText
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
    For Y = 1 To tvArbol.Nodes.Count
        tvArbol.Nodes(Y).Expanded = True
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
Dim Y As Byte
QuitaCadena = sCadena
For Y = 1 To Len(sQuita)
    QuitaCadena = Replace(QuitaCadena, Mid(sQuita, Y, 1), "")
Next
End Function

'Quita caracteres no dígitos
Function QuitaNoDígitos(ByVal sCadena As String) As String
Dim Y  As Integer
For Y = 1 To Len(sCadena)
    If InStr("0123456789", Mid(sCadena, Y, 1)) > 0 Then
        QuitaNoDígitos = QuitaNoDígitos & Mid(sCadena, Y, 1)
    End If
Next
End Function

'Última modif. por Miguel el Mié.30 de Enero del 2002
Sub RedAsunto(ByRef tvArbol As TreeView, lAsuIns As Long)
Dim rs As DAO.Recordset, adors As New ADODB.Recordset, i As Integer, lAva As Long
tvArbol.Nodes.Clear
If gSQLACC = cyAccess Then
    'Set rs = gdbmidb.OpenRecordset("select c.id,a.descripción+' ('+b.descripción+') Fecha: '+format(c.fecha,'dd-mmm-yyyy')+', Tipo: '+iif(a.clase=1,'a','b') from ((select * from actividades where id in (select idact from avances where idasuins=" & lAsuIns & ")) as a left join actividades b on a.idpad=b.id) inner join avances c on a.id=c.idact where c.idasuins=" & lAsuIns & " and c.idant is null", dbOpenSnapshot)
    Set rs = gdbmidb.OpenRecordset("select distinct idact,idant,id from avances where idasuins=" & lAsuIns, dbOpenSnapshot)
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
    adors.Open "select distinct idact,idant,id from avances where idasuins=" & lAsuIns & " and idant=id", gConSql, adOpenStatic, adLockReadOnly
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
Dim rs As Recordset, s As String, Y As Byte, adors As New ADODB.Recordset
tvArbol.Nodes.Clear
gs = ""
For Y = 0 To gyNiveles - 3
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
Dim rs As Recordset, s As String, Y As Byte, adors As New ADODB.Recordset
tvArbol.Nodes.Clear
gs = ""
For Y = 0 To gyNiveles - 3
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
Dim rs As Recordset, s As String, yNiveles As Byte, Y As Byte
Dim adors As New ADODB.Recordset
'gyNiveles
For Y = 0 To gyNiveles - 3
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
Dim rs As Recordset, s As String, yNiveles As Byte, Y As Byte
Dim adors As New ADODB.Recordset, bDesenlace As Boolean
'gyNiveles: variable global que indica el número de niveles en que se clasifican las actividades
For Y = 0 To gyNiveles - 3
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
Dim Y As Byte
Y = 0
Do While Y < ListaOrigen.ListCount
    If ListaOrigen.Selected(Y) Or bTodos Then
        ListaDestino.AddItem ListaOrigen.List(Y)
        ListaDestino.ItemData(ListaDestino.NewIndex) = ListaOrigen.ItemData(Y)
        ListaOrigen.RemoveItem (Y)
    Else
        Y = Y + 1
    End If
Loop
End Sub

Sub QuitaMemoriaForma(ByVal sForma As String, Optional yNoFormas As Byte)
Dim i As Long, Y As Byte
For i = Forms.Count - 1 To 0 Step -1
    If Forms(i).Name = sForma Then
        Unload Forms(i)
        Y = Y + 1
        If Y >= yNoFormas Then Exit Sub
    End If
Next
End Sub


Sub CalculaTop(ByRef Ctl As Control, ByRef lTop As Long, ByRef lLeft As Long, ByRef lmax As Long)
Dim Y As Byte
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
Dim nPro(2) As Integer, Y As Integer, n As Integer, s As String, sCve(2) As String, bMov As Boolean, yErr As Byte
Dim ynivel As Integer, iNodos As Integer, iValor() As Currency, l As Long, adors As New ADODB.Recordset
Dim sCvePadre As String, bSio As Boolean
Dim rs As Recordset
On Error GoTo ErrorCargaDatos:
If bNoBoxes Then
    tvArbol.CheckBoxes = False
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
            BarraProgreso.Show
            BarraProgreso.ProgressBar1.Value = 0
            BarraProgreso.Refresh
        End If
    End If
    Do While Not rs.EOF
        sCvePadre = "r"
        For ynivel = 0 To yNoNiveles - 1
            If IsNull(rs(ynivel)) Then Exit For
            sCvePadre = sCvePadre + Right(String(9, "0") & rs(ynivel), 10)
            If iValor(ynivel) <> rs(ynivel) Then
                iValor(ynivel) = rs(ynivel)
                For Y = ynivel + 1 To yNoNiveles - 1
                    iValor(Y) = 0
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
                BarraProgreso.ProgressBar1.Value = 100 * (rs.AbsolutePosition + 1) / l
                BarraProgreso.Refresh
            End If
        End If
        'Debug.Print Str(rs.AbsolutePosition)
    Loop
Else
    If InStr(sQuery, "{") > 0 And InStr(sQuery, "}") > 0 Then
        adors.Open sQuery, gConSql, adOpenForwardOnly, adLockReadOnly
    Else
        adors.Open sQuery, gConSql, adOpenStatic, adLockReadOnly
    End If
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
'            BarraProgreso.Show
'            BarraProgreso.ProgressBar1.Value = 0
'            BarraProgreso.Refresh
'        End If
    End If
    Do While Not adors.EOF
        sCvePadre = "r"
        For ynivel = 0 To yNoNiveles - 1
            If IsNull(adors(ynivel)) Then Exit For
            sCvePadre = sCvePadre + Right(String(9, "0") & adors(ynivel), 10)
            If iValor(ynivel) <> adors(ynivel) Then
                iValor(ynivel) = adors(ynivel)
                For Y = ynivel + 1 To yNoNiveles - 1
                    iValor(Y) = 0
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
                BarraProgreso.ProgressBar1.Value = 100 * (adors.AbsolutePosition + 1) / (l + 1)
                BarraProgreso.Refresh
            End If
        End If
        'Debug.Print Str(adors.AbsolutePosition)
    Loop
End If
If bAbierto Then
    For Y = 1 To tvArbol.Nodes.Count
        tvArbol.Nodes(Y).Expanded = True
    Next
End If
BarraProgreso.ProgressBar1.Value = 100
BarraProgreso.Refresh
Unload BarraProgreso
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
Y = MsgBox(sError, vbAbortRetryIgnore + vbCritical, "Error no esperado (" + Str(Err.Number) + ")")
If Y = vbCancel Then
    Exit Sub
ElseIf Y = vbRetry Then
    Resume
ElseIf Y = vbIgnore Then
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
End Function


Function ObtieneRefArchivo(lAnálisis As Long) As String
Dim adors As ADODB.Recordset, s As String
Dim sArchivo As String
On Error GoTo salir:
    Set adors = New ADODB.Recordset
    adors.Open "select r.Expediente as folio from análisis ana, registroxif ri, registro r where ana.id=" & lAnálisis & " and ana.idregxif=ri.id and ri.idreg=r.id", gConSql, adOpenStatic, adLockReadOnly
    If adors.EOF Then
        If Len(gsDirDocumentos) = 0 Then
            ObtieneRefArchivo = CurDir & "\DocsSIAM\"
        Else
            ObtieneRefArchivo = Mid(gsDirDocumentos, 1, InStrRev(Mid(gsDirDocumentos, 1, Len(gsDirDocumentos) - 1), "\")) & "DocsSIAM\"
        End If
    Else
        s = adors(0)
        s = Replace(Replace(Replace(Replace(Replace(Replace(Replace(s, "/", "_"), ",", ""), "|", ""), "?", ""), "¿", ""), "¡", ""), "!", "")
        If Len(gsDirDocumentos) = 0 Then
            ObtieneRefArchivo = CurDir & "\DocsSIAM\" & s & "\"
        Else
            ObtieneRefArchivo = Mid(gsDirDocumentos, 1, InStrRev(Mid(gsDirDocumentos, 1, Len(gsDirDocumentos) - 1), "\")) & "DocsSIAM\" & s & "\"
        End If
    End If
Exit Function
salir:
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

Function NodoContieneFecha(ByRef nodo As Node) As Integer
Dim s As String
If InStrRev(nodo.Text, "(") < InStrRev(nodo.Text, " Resp.:") Then
    s = Mid(nodo.Text, InStrRev(nodo.Text, "(") + 1)
    s = Mid(s, 1, InStr(s, " Resp") - 1)
    If IsDate(s) Then NodoContieneFecha = InStrRev(nodo.Text, "(")
End If
End Function

Function VerificaExistenciaDocumentos(sActividades As String) As String
Dim rs As Recordset, adors As New ADODB.Recordset, Y As Byte, yy As Byte, s As String
'adors.CursorLocation = adUseClient
For yy = 1 To Len(sActividades) / 4 'verifica documentos
    If adors.State > 0 Then adors.Close
    'adors.CursorLocation = adUseClient
    adors.Open "select * from documentos where id=" + Mid(sActividades, (yy - 1) * 4 + 1, 4), gConSql, adOpenStatic, adLockReadOnly
    If adors.RecordCount > 0 Then
        If Len(adors!archivo) > 0 Then
            If Len(Dir(gsDirDocumentos + adors!archivo + ".doc")) = 0 Then
                yError = MsgBox("El documento " & IIf(IsNull(adors!descripción), "Sin descripción", adors!descripción) & " (" + adors!archivo + ".doc)  se encuentra en proceso de elaboración por el área jurídica correspondiente.", vbOKOnly + vbInformation, "")
                s = s & "general.doc" & gsSeparador
            Else
                s = s & adors!archivo & ".doc" & gsSeparador
            End If
        End If
    Else
        s = s & gsSeparador
    End If
 Next
VerificaExistenciaDocumentos = s
End Function

Function GeneraDocumento(ByRef adorsDoc As ADODB.Recordset, ByVal lAnálisis As Long, Optional ByVal lSeguimiento As Long, Optional ByVal sDocumento As String) As Boolean
Dim s As String, rsCamposVarios As Recordset, sFormato As String, Y As Byte, yErr As Byte, rsCampos As Recordset, rsDatosIni As Recordset
Dim DoctoWord As Word.Document, yOffice As Byte, Docto As Word.Document ', mdbAccess As Access.Application
Dim ApWord As Word.Application, i As Long, yIntentosActDoc As Byte, sArchivo As String
Dim rs As Recordset, adors As New ADODB.Recordset, adors1 As New ADODB.Recordset ', OBJ As Object
Dim bNuevo As Boolean
Dim bActualizaNombreArchivo As Boolean
'adors.CursorLocation = adUseClient
'adors1.CursorLocation = adUseClient
If Len(sDocumento) > 0 Then
    bNuevo = False
Else
    bNuevo = True
End If
Nuevo:
sArchivo = ObtieneRefArchivo(lAnálisis)
If Not bNuevo Then
    sArchivo = sArchivo & sDocumento
    sFormato = sArchivo
Else
    sFormato = gsDirDocumentos & adorsDoc!archivo + ".doc"
    If adors.State > 0 Then adors.Close
    'adors.CursorLocation = adUseClient
'    adors.Open "select a.id from seguimiento s, relacióntareadocumento rtd where s.idtar=rtd.idtar and (s.id=" & lSeguimiento & " or a.idant=" & lSeguimiento & ") and rtd.iddoc=" & adorsDoc!ID, gConSql, adOpenStatic, adLockReadOnly
'    If adors.RecordCount > 0 Then
'        If lSeguimiento <> adors(0) Then lSeguimiento = adors(0)
'    End If
    sArchivo = sArchivo & adorsDoc!archivo & lSeguimiento & ".doc"
    If bActualizaNombreArchivo Then
        If adors.State > 0 Then adors.Close
        adors.Open "select archivo from seguimientodoctos where idseg=" & lSeguimiento & " and  iddoc=" & adorsDoc!ID, gConSql, adOpenStatic, adLockReadOnly
        If Not adors.EOF Then
            If InStr(sArchivo, "\") > 0 Then
                If IsNull(adors(0)) Then gConSql.Execute "update seguimientodoctos set archivo='" & Mid(sArchivo, InStrRev(sArchivo, "\") + 1) & "' where idseg=" & lSeguimiento & " and  iddoc=" & adorsDoc!ID
            End If
        End If
    End If
End If
'If Len(sArchivo) > 0 And Len(Dir(sArchivo)) = 0 Then
'    y = MsgBox("No existe el achivo (" + sArchivo + ") ¿Deseas generarlo a partir del formato?", vbYesNo + vbQuestion, "Error: No se localizo el archivo")
'    If y = vbNo Then
'        Exit Function
'    End If
'End If
'If Len(sArchivo) = 0 Or y = vbYes Then
     If Len(Dir(sFormato)) = 0 Then
        If Not bNuevo Then
            If MsgBox("No existe el achivo (" + sFormato + ") ¿Desea generarlo nuevamente a partir de la plantilla?", vbYesNo + vbQuestion, "Error: No se localizo el archivo formato") = vbNo Then Exit Function
            bActualizaNombreArchivo = True
            bNuevo = True
            GoTo Nuevo:
        Else
            MsgBox "No existe el achivo (" + sFormato + ")", vbOKOnly + vbCritical, "Error: No se localizo el archivo formato"
            Exit Function
        End If
    End If
    On Error GoTo ErrWord:
    'ActiveForm.MousePointer = 11
    Y = ReplicaArchivo(sFormato, sArchivo, bNuevo)
    'If Len(Dir(sArchivo)) = 0 Then
    '    s = sArchivo
    '    FileCopy sFormato, s
    '    If Len(Trim(Dir(s))) > 0 Then
    '    End If
    'End If
    'If y > 200 Then
    '    MsgBox "Error: existe problema al generar archivo temporal", vbOKOnly, ""
    '    mdi.ActiveForm.MousePointer = 0
    '    Exit Sub
    'End If
    sFormato = sArchivo

CreaDocumento:
    Set DoctoWord = GetObject(sFormato, "Word.Document")  'con .Add añadimso Libros de trabajo de la aplicacion '2000
'End If

ActivaWordApp:
Set ApWord = GetObject(, "Word.Application") 'Método CreateObject y Application '97
Set DoctoWord = ApWord.Documents.Open(sFormato, ReadOnly:=False, PasswordDocument:="Condusef2000")  'con .Add añadimso Libros de trabajo de la aplicacion 97
'Set OBJ = GetObject("Z:\DocsSIO\2005_090_162746\FCO01451688701.doc")
'OBJ.View

If Y = 1 Then 'ya existía el documento por lo tanto ...salir
    Exit Function
End If
If Len(sDocumento) > 0 And lSeguimiento > 0 Then
    sFormato = gsDirDocumentos + sDocumento
Else
        sFormato = gsDirDocumentos + adorsDoc!archivo + ".doc"
End If
gs1 = ""

ContinuaSiguiente:

    Call ProcesaDatosDocumentoAdo(DoctoWord, lAnálisis, lSeguimiento)
'MDI.ActiveForm.MousePointer = 0
Exit Function

ErrWord:
    If Err.Number = 70 And Y < 200 Then
        Y = Y + 1
        s = gsDirDocumentos + "temp" + Trim(Str(Y)) + ".doc"
        Resume
    ElseIf Err.Number = 5825 Or Err.Number = 5941 Then
        GoTo ContinuaSiguiente:
    ElseIf Err.Number < -1000 And Y < 200 Then
        If Y < 100 Then Y = 190
        Y = Y + 1
        GoTo ActivaWordApp:
    ElseIf Err.Number = 429 Then
        GoTo ActivaWordApp:
    End If
    yErr = MsgBox(Err.Description + IIf(Err.Number = 287, ". Probablemente se encuente el documento Abierto.", ""), vbAbortRetryIgnore, "Error: " + Str(Err.Number))
    If yErr = vbRetry Then
        Resume
    ElseIf yErr = vbIgnore Then
        Resume Next
    End If
MDI.ActiveForm.MousePointer = 0
End Function

'Genera copia de archivo origen a Destino generando la ruta del destino en caso que no exista
'0: No existe; 1: Ya existe.
Function ReplicaArchivo(sOrigen, sDestino, bNuevo As Boolean) As Byte
Dim s As String, ss As String
On Error GoTo salir:
s = sDestino
Do While InStr(s, "\")
    ss = ss & Mid(s, 1, InStr(s, "\"))
    If Len(Dir(ss, vbDirectory)) = 0 Then
        MkDir ss
    End If
    s = Mid(s, InStr(s, "\") + 1)
Loop
If Len(Dir(sDestino)) > 0 And bNuevo Then
    Kill sDestino
End If
If Len(Dir(sDestino)) = 0 And bNuevo Then
    s = sDestino
    If Right(sOrigen, 1) <> "\" Then
        FileCopy sOrigen, s
    End If
    Exit Function
End If
ReplicaArchivo = 1
Exit Function
salir:
'Resume
Call MsgBox(Err.Description, vbOKOnly, Err.Number)
End Function

Sub BorraDocumentos(ByVal lSeguimiento As Long, lAsunto As Long, ByVal bTodos As Boolean)
Dim s As String, rs As DAO.Recordset, adors As ADODB.Recordset
On Error GoTo salir:
sArchivo = ObtieneRefArchivo(lAsunto)
If bTodos Then
    'En este caso lSeguimiento contiene el node asunto lSeguimiento=lasunto
    sDocumentos = ""
    If gSQLACC = cyAccess Then
        Set rs = gdbmidb.OpenRecordset("select tr.archivo from (tareasrealizadas tr inner join avances av on tr.idava=av.id) inner join asuntoinstitución ai on av.idasuins=ai.id where ai.idasu=" & lAsunto, dbOpenSnapshot)
        Do While Not rs.EOF
            If Not IsNull(rs(0)) Then
                sDocumentos = sDocumentos & rs(0) & gsSeparador
            End If
            rs.MoveNext
        Loop
        Set rs = gdbmidb.OpenRecordset("select dr.archivo from documentosrecibidos dr where dr.idasu=" & lAsunto, dbOpenSnapshot)
        Do While Not rs.EOF
            If Not IsNull(rs(0)) Then
                sDocumentos = sDocumentos & rs(0) & gsSeparador
            End If
            rs.MoveNext
        Loop
    Else
        Set adors = New ADODB.Recordset
        adors.Open "select tr.archivo from (tareasrealizadas tr inner join avances av on tr.idava=av.id) inner join asuntoinstitución ai on av.idasuins=ai.id where ai.idasu=" & lAsunto, gConSql, adOpenStatic, adLockReadOnly
        Do While Not adors.EOF
            If Not IsNull(adors(0)) Then
                sDocumentos = sDocumentos & adors(0) & gsSeparador
            End If
            adors.MoveNext
        Loop
        adors.Close
        adors.Open "select dr.archivo from documentosrecibidos dr where dr.idasu=" & lAsunto, gConSql, adOpenStatic, adLockReadOnly
        Do While Not adors.EOF
            If Not IsNull(adors(0)) Then
                sDocumentos = sDocumentos & adors(0) & gsSeparador
            End If
            adors.MoveNext
        Loop
    End If
    s = Dir(sArchivo & "*.*")
Else
    s = Dir(sArchivo & "*" & lSeguimiento & ".doc")
    sDocumentos = ""
    If gSQLACC = cyAccess Then
        Set rs = gdbmidb.OpenRecordset("select archivo from tareasrealizadas where idava=" & lSeguimiento, dbOpenSnapshot)
        Do While Not rs.EOF
            If Not IsNull(rs(0)) Then
                sDocumentos = sDocumentos & IIf(IsNull(rs(0)), "", rs(0)) & gsSeparador
            End If
            rs.MoveNext
        Loop
    Else
        Set adors = New ADODB.Recordset
        adors.Open "select archivo from tareasrealizadas where idava=" & lSeguimiento, gConSql, adOpenStatic, adLockReadOnly
        Do While Not adors.EOF
            If Not IsNull(adors(0)) Then
                sDocumentos = sDocumentos & IIf(IsNull(adors(0)), "", adors(0)) & gsSeparador
            End If
            adors.MoveNext
        Loop
    End If
    'sDocumentos = Replace(LCase(sDocumentos), ".doc", lSeguimiento & ".doc")
End If
Do While Len(s) > 0
    If InStr(gsSeparador & LCase(sDocumentos), gsSeparador & LCase(s) & gsSeparador) = 0 Then
        Kill sArchivo & s
    End If
    s = Dir
Loop
Exit Sub
salir:
MsgBox Err.Description, vbOKOnly + vbCritical, Err.Number
'Resume
End Sub

Sub ProcesaDatosDocumentoAdo(ByRef DoctoWord As Word.Document, ByRef lAnálisis As Long, lSeguimiento As Long)
Const cyActConciliación = 14
Const cyActRespIF = 10
Dim adors As New ADODB.Recordset, Y As Byte, sCampos As String, sFrom As String, sCampo() As String, s As String, ss As String, d As Date
Dim i As Long, yy As Byte, iAsu As Integer, iCve As Integer, iCla As Integer, y2 As Byte, sDesCam As String, dAhora As Date, lAsunto As Long, yNoIncremento As Byte
Dim sValida_Documento As String
Dim ii As Integer
sValida_Documento = "1234"
On Error GoTo ErrWordCampos:
DoctoWord.Application.Visible = True
DoctoWord.Application.Activate
DoctoWord.Activate
For ii = 1 To DoctoWord.FormFields.Count
    sDesCam = ""
    s = Trim(DoctoWord.FormFields(ii).Result)
    If adors.State Then adors.Close
    adors.Open "select f_Doctos_Campo(" & lAnálisis & "," & lSeguimiento & ",'" & s & "') from dual", gConSql, adOpenStatic, adLockReadOnly
    If Not adors.EOF Then
        sDesCam = IIf(IsNull(adors(0)), "???", adors(0))
    End If
    'DoctoWord.FormFields(y).Result = IIf(ComAscii(DoctoWord.FormFields(y).Result, UCase(DoctoWord.FormFields(y).Result)), UCase(sDesCam), sDesCam)
    DoctoWord.FormFields(ii).Result = sDesCam
    If yNoIncremento <> 0 Then
        ii = ii - 1
        yNoIncremento = 0
    End If
Next


Exit Sub
ErrWordCampos:
    If Err.Number = 70 And Y < 200 Then
        Y = Y + 1
        s = "temp" + Trim(Str(Y)) + ".doc"
        Resume
    ElseIf Err.Number = 4120 Or Err.Number = 5825 Or Err.Number = 5941 Then
        Resume Next
    ElseIf Err.Number < 1000 And Y < 200 Then
        If Y < 100 Then Y = 199
        Y = Y + 1
        Resume
    End If
    yErr = MsgBox(Err.Description, vbAbortRetryIgnore, "Error: " + Str(Err.Number))
    If yErr = vbRetry Then
        Resume
    ElseIf yErr = vbIgnore Then
        Resume Next
    End If
End Sub


Sub MuestraEtiqueta(ByRef Ctl As Control, ByRef txtEtiqueta As TextBox, yAcción As Byte, ByRef lSegundos As Long, iArregloCombo() As Integer)
Dim lTop As Long, lLeft As Long, lmax As Long, ss As String, l As Long, sin As Single
Dim rs As DAO.Recordset, adors As ADODB.Recordset, i As Integer, bCambiar As Boolean, s As String
Static iValorCombo(6) As Integer
Static sEstadística As String
Static sEtiqueta As String
If lSegundos <= 0 Then
    lSegundos = Timer
    Exit Sub
ElseIf Timer - lSegundos < 1 Then
    Exit Sub
End If
If yAcción <> 0 And LCase(Ctl.Name) <> "eticombo" And Not (yAcción = 1 And InStr("folio:folio int.:", LCase(Ctl.Caption)) > 0) Then Exit Sub
If Ctl.Name = "EtiCombo" Then
    If Ctl.Index = 2 Then
        s = "prod. o serv. Nivel 1"
    ElseIf Ctl.Index = 11 Then
        s = "prod. o serv. Nivel 2"
    ElseIf Ctl.Index = 3 Then
        s = "prod. o serv. Nivel 3"
    Else
        s = Ctl.Caption
    End If
Else
    s = Ctl.Caption
End If
If sEtiqueta <> s Or Not txtEtiqueta.Visible Then
    Call CalculaTop(Ctl, lTop, lLeft, lmax)
    If BuscaEtiqueta(s, lTop, lLeft, txtEtiqueta, lmax) Then
        sEtiqueta = s
    End If
Else
    If Not txtEtiqueta.Visible Then txtEtiqueta.Visible = True
    Exit Sub
End If
s = ""
'agrega a la etiqueta estadística de asuntos según valores de los combos
Exit Sub
If txtEtiqueta.Visible And iArregloCombo(0) > 0 Then 'verifica que esté visible la etiqueta y se esté en el combo correspondiente
    If iArregloCombo(0) = iValorCombo(0) Then 'Verifica si ya fue calculada la estadística para no volver a ejecutarla
        For i = 1 To iArregloCombo(0)
            If iArregloCombo(i) <> iValorCombo(i) Then
                bCambiar = True
                Exit For
            End If
        Next
    Else
        bCambiar = True
    End If
    If bCambiar Then
        For i = 1 To 11
            If iArregloCombo(i) <> iValorCombo(i) Then
                bCambiar = True
                Exit For
            End If
        Next
        For i = 1 To iArregloCombo(0)
            If iArregloCombo(i) > 0 Then
                ss = ss & Trim(Mid("Clase Inst.InstituciónProducto N1Producto N2Producto N3Causa      ", (i - 1) * 11 + 1, 11)) & ", "
                s = s & IIf(i <= 2, "ai.id", "a.id") & Mid("clainspr1pr2pr3cau", (i - 1) * 3 + 1, 3) & "=" & iArregloCombo(i) & " and "
            End If
        Next
        If gSQLACC = cyAccess Then
            Set rs = gdbmidb.OpenRecordset("select iif(isnull(e.id),'En trámite',iif(e.favorable=2,'Fav.Usuario',iif(e.favorable=1,'Fav.Institución','Fav.Ninguno'))) as a_Favor, av.proceso, count(*) as asuntos from ((asuntos a inner join asuntoinstitución ai on a.id=ai.idasu) left join evaluación e on a.id=e.idasu) left join (select ac1.descripción as proceso,av.idasuins from (((select max(id) as idava from avances where fecha is not null group by idasuins) a left join avances av on a.idava=av.id) left join actividades ac on av.idact=ac.id) left join actividades ac1 on ac.idpad=ac1.id) av on ai.id=av.idasuins" & IIf(Len(s) > 0, " where " & Mid(s, 1, Len(s) - 5), "") & " group by iif(isnull(e.id),'En trámite',iif(e.favorable=2,'Fav.Usuario',iif(e.favorable=1,'Fav.Institución','Fav.Ninguno'))),av.proceso ", dbOpenSnapshot)
            If rs.EOF Then
                sEstadística = Chr(13) & Chr(10) & "No existen asuntos de " & IIf(Len(ss) > 0, Mid(ss, 1, Len(ss) - 2), "") & " con la información especificada"
            Else
                sEstadística = Chr(13) & Chr(10) & "Resumen de los asuntos (" & IIf(Len(ss) > 0, Mid(ss, 1, Len(ss) - 2), "") & ") con la información especificada"
                i = 0
                s = ""
                Do While Not rs.EOF
                    i = i + 1
                    sEstadística = sEstadística & Chr(13) & Chr(10) & rs(2) & gsSeparador & i & " Asuntos (" & rs(0) & ") (" & IIf(IsNull(rs(1)), "Sin Proceso", rs(1)) & ")"
                    l = l + rs(2)
                    s = s & rs(2) & ","
                    rs.MoveNext
                Loop
                i = 0
                Do While InStr(s, ",")
                    sin = (100 * Val(s) / l)
                    s = Mid(s, InStr(s, ",") + 1)
                    i = i + 1
                    sEstadística = Replace(sEstadística, gsSeparador & i & " ", " (" & Mid(sin & "", 1, IIf(sin > 10, 5, 4)) & " %) ")
                Loop
            End If
        Else
            Set adors = New ADODB.Recordset
            adors.Open "select nvl2(e.id, case when e.favorable=2 then 'Fav.Usuario' else case when e.favorable=1 then 'Fav.Institución' else 'Fav.Ninguno' end end,'En trámite') as a_Favor, av.proceso, count(*) as asuntos from asuntos a inner join asuntoinstitución ai on a.id=ai.idasu left join evaluación e on a.id=e.idasu left join (select ac1.descripción as proceso,av.idasuins from (select max(id) as idava from avances where fecha is not null group by idasuins) a left join avances av on a.idava=av.id left join actividades ac on av.idact=ac.id left join actividades ac1 on ac.idpad=ac1.id) av on ai.id=av.idasuins" & IIf(Len(s) > 0, " where " & Mid(s, 1, Len(s) - 5), "") & " group by nvl2(e.id, case when e.favorable=2 then 'Fav.Usuario' else case when e.favorable=1 then 'Fav.Institución' else 'Fav.Ninguno' end end,'En trámite'), av.proceso", gConSql, adOpenStatic, adLockReadOnly
            If adors.EOF Then
                sEstadística = Chr(13) & Chr(10) & "No existen asuntos de " & IIf(Len(ss) > 0, Mid(ss, 1, Len(ss) - 2), "") & " con la información especificada"
            Else
                sEstadística = Chr(13) & Chr(10) & "Resumen de los asuntos (" & IIf(Len(ss) > 0, Mid(ss, 1, Len(ss) - 2), "") & ") con la información especificada"
                i = 0
                s = ""
                Do While Not adors.EOF
                    i = i + 1
                    sEstadística = sEstadística & Chr(13) & Chr(10) & adors(2) & gsSeparador & i & " Asuntos (" & adors(0) & ") (" & IIf(IsNull(adors(1)), "Sin Proceso", adors(1)) & ")"
                    l = l + adors(2)
                    s = s & adors(2) & ","
                    adors.MoveNext
                Loop
                i = 0
                Do While InStr(s, ",")
                    sin = (100 * Val(s) / l)
                    s = Mid(s, InStr(s, ",") + 1)
                    i = i + 1
                    sEstadística = Replace(sEstadística, gsSeparador & i & " ", " (" & Mid(sin & "", 1, IIf(sin > 10, 5, 4)) & " %) ")
                Loop
            End If
        End If
        For i = 0 To iArregloCombo(0)
            iValorCombo(i) = iArregloCombo(i)
        Next
    End If
    txtEtiqueta = txtEtiqueta & sEstadística
    txtEtiqueta.Height = txtEtiqueta.Height * 2
    'txtEtiqueta.ScrollBars = 2
    txtEtiqueta.Refresh
Else
    'If txtEtiqueta.ScrollBars <> 0 Then
    '    txtEtiqueta.ScrollBars = 0
    '    txtEtiqueta.Refresh
    'End If
End If
End Sub

Function BuscaEtiqueta(ByVal sCampo As String, lTop As Long, lLeft As Long, ByRef txtEtiqueta As TextBox, ByRef lmax As Long) As Boolean
Dim i As Long, l As Long
sCampo = LCase(Replace(Replace(sCampo, "&", ""), ":", ""))
grsEti.Index = "elemento"
grsEti.Seek "=", sCampo
If grsEti.NoMatch Then
    grsEti.Index = "campo"
    grsEti.Seek "=", sCampo
End If
If Not grsEti.NoMatch Then
    txtEtiqueta.Text = grsEti!descripción
    txtEtiqueta.Top = lTop
    i = Int((Len(grsEti!descripción) - 1) / 50)
    'txtEtiqueta.Width = (2 + i) * 1250
    txtEtiqueta.Height = IIf(Len(grsEti!descripción) <= 25, 230, 465)
    If lLeft + txtEtiqueta.Width > lmax - 50 Then
        txtEtiqueta.Left = lmax - 50 - txtEtiqueta.Width
    Else
        txtEtiqueta.Left = lLeft
    End If
    BuscaEtiqueta = True
    txtEtiqueta.Visible = True
Else
    txtEtiqueta.Visible = False
End If
End Function

Function ComAscii(sCadena1 As String, scadena2 As String) As Boolean
Dim i As Long
If Len(sCadena1) <> Len(scadena2) Then Exit Function
For i = 1 To Len(sCadena1)
    If Asc(Mid(sCadena1, i, 1)) <> Asc(Mid(scadena2, i, 1)) Then Exit Function
Next
ComAscii = True
End Function


Function BuscaAvance(ByVal lAva As Long, ByVal sActTar As String) As Long
Dim rs As DAO.Recordset, adors As ADODB.Recordset
Set adors = New ADODB.Recordset
adors.Open "select id,idant,idact,idtar from avances where id=" & lAva, gConSql, adOpenStatic, adLockReadOnly
Do While Not adors.EOF
    If InStr(sActTar, "," & IIf(IsNull(adors(2)), "-1", adors(2)) & ",") > 0 Or InStr(sActTar, "," & IIf(IsNull(adors(3)), "-1", adors(3)) & ",") > 0 Then
        BuscaAvance = lAva
        Exit Function
    End If
    If lAva = adors(1) Then
        Exit Function
    Else
        lAva = adors(1)
    End If
    adors.Close
    adors.Open "select id,idant,idact,idtar from avances where id=" & lAva, gConSql, adOpenStatic, adLockReadOnly
Loop
End Function

Function sMes(yMes As Byte, Optional bCorto As Boolean) As String
sMes = "enero     febrero   marzo     abril     mayo      junio     julio     agosto    septiembreoctubre   noviembre diciembre "
If bCorto Then
    sMes = Mid(sMes, (yMes - 1) * 10 + 1, 3)
Else
    sMes = Trim(Mid(sMes, (yMes - 1) * 10 + 1, 10))
End If
End Function

Function Importe(s As String, yDigitos As Byte) As String
Dim sCadenaUnidades As String, sCadenaDecenas As String, sCadenaCentenas
sCadenaCentenas = "        un      dos     tres    cuatro  quini   seis    sete    ocho    nove    diez    once    doce    trece   catorce quince  "
sCadenaUnidades = "        un      dos     tres    cuatro  cinco   seis    siete   ocho    nueve   diez    once    doce    trece   catorce quince  "
sCadenaDecenas = "         dieci    veinte   treinta  cuarenta cincuentasesenta  setenta  ochenta  noventa  "
Select Case yDigitos
Case 1
    Importe = Trim(Mid(sCadenaUnidades, Val(s) * 8 + 1, 8)) + " "
Case 2
    Importe = Trim(Mid(sCadenaDecenas, Val(Mid(s, 1, 1)) * 9 + 1, 9)) + " "
    If Val(Right(s, 1)) > 0 Then
        If Val(s) < 20 Then
            Importe = "dieci"
        ElseIf Val(s) < 30 Then
            Importe = "veinti"
        Else
            Importe = Importe + "y "
        End If
    End If
Case 3
    If Mid(s, 1, 1) = 1 Then
        If Val(Right(s, 2)) = 0 Then
            Importe = "cien "
        Else
            Importe = "ciento "
        End If
    Else
        If Val(Mid(s, 1, 1)) = 5 Then
            Importe = "quinientos "
        Else
            Importe = Trim(Mid(sCadenaCentenas, Val(Mid(s, 1, 1)) * 8 + 1, 8)) + IIf(Val(Mid(s, 1, 1)) > 0, "cientos ", "")
        End If
    End If
End Select
Importe = IIf(Len(Trim(Importe)) = 0, "", Importe)
Importe = Replace(Replace(Importe, "dieciseis", "dieciséis"), "veintiseis", "veintiséis")
End Function

'Convierte a mayúsculas las letras inicio de palabra y minúsculas las demás
'Es utilizada para estandarizar los campos de nombres y domicilios
Function AltasBajas_NuevaLetra(ByRef txtCampo As TextBox, ByVal KeyAscii As Byte) As Byte
AltasBajas_NuevaLetra = 0
If txtCampo.SelStart = 0 Then
    AltasBajas_NuevaLetra = 200
Else
    If InStr(" .", Mid(txtCampo.Text, txtCampo.SelStart, 1)) > 0 Then AltasBajas_NuevaLetra = 200
End If
If AltasBajas_NuevaLetra = 200 Then
    AltasBajas_NuevaLetra = Asc(UCase(Chr(KeyAscii)))
Else
    AltasBajas_NuevaLetra = Asc(LCase(Chr(KeyAscii)))
End If
End Function

'Convierte a mayúsculas las letras inicio de palabra y minúsculas las demás
'Es utilizada para estandarizar los campos de nombres y domicilios cuando pierde el foco
Function AltasBajas(ByRef txtCampo As TextBox) As String
Dim s As String, Y As Byte, y2 As Byte
s = ""
y2 = 1
For Y = 1 To Len(txtCampo.Text)
    If Len(s) > 0 Then
        If Right(s, 1) = " " And Mid(txtCampo.Text, Y, 1) = " " Then Y = Y + 1
    End If
    If y2 < 2 Then
        s = s + IIf(y2 = 1, UCase(Mid(txtCampo.Text, Y, 1)), LCase(Mid(txtCampo.Text, Y, 1)))
        y2 = IIf(InStr(" .", Mid(txtCampo.Text, Y, 1)) > 0, 1, 0)
    End If
Next
AltasBajas = s
s = "La  El  Los Las De  Del ParaPor Y   "
For Y = 1 To Len(s) / 4
    If InStr(AltasBajas + " ", " " + Trim(Mid(s, (Y - 1) * 4 + 1, 4)) + " ") > 0 Then
        AltasBajas = Trim(Replace(AltasBajas + " ", " " + Trim(Mid(s, (Y - 1) * 4 + 1, 4)) + " ", " " + LCase(Trim(Mid(s, (Y - 1) * 4 + 1, 4)) + " ")))
    End If
Next
End Function

'En caso de Access devuelve la hora de la tabla fechaservidor de la máquina que contiene SIO.MDB
'En caso de SQL devuelve Getdate() fecha del servidor
Function AhoraServidor(Optional yValida As Byte) As Date
Static rs As Recordset, d As Date, yError As Byte, adors As New ADODB.Recordset
On Error GoTo ErrorFechaServidor:
If Not gSQLACC = cyAccess Then 'sql
    If adors.State > 0 Then adors.Close
        If gSQLACC = cyOracle Then 'sql
            adors.Open "select sysdate from dual", gConSql, adOpenStatic, adLockReadOnly
        Else
            adors.Open "select getdate() from actividades where id=1", gConSql, adOpenStatic, adLockReadOnly
        End If
    If adors.RecordCount > 0 Then
        AhoraServidor = adors(0)
    Else
        AhoraServidor = Now
    End If
    Exit Function
End If
If yValida > 0 And InStr(gdbmidb.Name, "z:\") > 0 Then
    Set rs = gdbmidb.OpenRecordset("fechaservidor", dbOpenSnapshot)
    If Not rs.EOF Then
        AhoraServidor = rs(0)
        d = Now
        If Format(rs(0), gsFormatoFecha) <> Format(d, gsFormatoFecha) Then
            ''Aguas cambios específicos para central
            MsgBox "La fecha obtenida del servidor (" + Format(rs(0), gsFormatoFecha) + ") es diferente a la de su máquina (" + Format(d, gsFormatoFecha) + "). Favor de corregir la fecha de su máquina o verificar que el programa que actualiza la fecha en la base de datos del servidor se esté ejecutando", vbOKOnly, ""
            If yValida = 200 Then End
        End If
        If Abs(DateDiff("s", AhoraServidor, d)) > 120 Then
            ''Aguas cambios específicos para central
            If MsgBox("La hora del servidor (" + Format(AhoraServidor, "hh:mm:ss") + ") y la de su máquina (" + Format(Now, "hh:mm:ss") + ") son diferentes. Se actualizará la hora de su máquina a la hora del servidor", vbOKCancel + vbInformation, "") = vbCancel Then
                If yValida = 200 Then
                    End
                Else
                    MsgBox "Verifique la hora del servidor con la de su equipo.", vbOKOnly + vbInformation, "Importante"
                    Exit Function
                End If
            Else
                Time = AhoraServidor
            End If
        ElseIf Abs(DateDiff("s", AhoraServidor, d)) > 20 Then
            Time = AhoraServidor
        End If
    Else
        MsgBox "No se ha actualizado la fecha y hora del servidor", vbOKOnly + vbCritical, "Error"
    End If
    rs.Close
    Set rs = Nothing
Else
    AhoraServidor = Now
End If
Exit Function
ErrorFechaServidor:
If Err.Number = 91 Then Exit Function
yError = MsgBox("Error: " + Err.Description, vbAbortRetryIgnore, "Error: Fecha del servidor (" + Str(Err.Number) + " )")
If yError = vbRetry Then
    Resume
ElseIf yError = vbIgnore Then
    Resume Next
End If
End Function



Function ImporteLetra(dbl As Double, sUnidad As String) As String
Dim sImporte As String, s As String, s2 As String, Y As Byte, ss As String, yy As Byte, i As Long
sImporte = Trim(Str(dbl))
If InStr(sImporte, ".") > 0 Then
    s2 = Mid(sImporte, InStr(sImporte, "."))
    sImporte = Mid(sImporte, 1, InStr(sImporte, ".") - 1)
End If
Do While Len(sImporte) > 0
    If Len(sImporte) >= 3 Then
        s = Right(sImporte, 3)
        sImporte = Mid(sImporte, 1, Len(sImporte) - 3)
    Else
        s = Right("00" + sImporte, 3)
        sImporte = ""
    End If
    If Val(Right(s, 2)) < 16 Then
        ss = Importe(Right(s, 3), 3) + Importe(Right(s, 2), 1)
    Else
        ss = Importe(Right(s, 3), 3) + Importe(Right(s, 2), 2) + Importe(Right(s, 1), 1)
    End If
    Select Case yy
    Case 1, 3
        ss = ss + IIf(Len(ss) > 0, "mil ", "")
    Case 2
        ss = ss + IIf(Int(dbl / 1000000) = 1, "millón ", "millones ") + IIf(dbl Mod 1000000 = 0, "de ", "")
    Case Is = 4
        ss = ss + "????"
    End Select
    ImporteLetra = ss + ImporteLetra
    If InStr(ImporteLetra, "??") Then Exit Function
    yy = yy + 1
Loop
s = LCase(Trim(sUnidad))
sUnidad = ""
If dbl >= 1 And dbl < 2 Then
    Do While InStr(s, " ") > 0
        ss = Mid(s, 1, InStr(s, " ") - 1)
        sUnidad = sUnidad + IIf(Right(ss, 1) = "s", Mid(ss, 1, Len(ss) - IIf(Right(ss, 3) = "res", 2, 1)), ss) + " "
        s = Trim(Mid(s, InStr(s, " ") + 1))
    Loop
    sUnidad = sUnidad + IIf(Right(s, 1) = "s", Mid(s, 1, Len(s) - IIf(Right(s, 3) = "res", 2, 1)), s)
Else
    Do While InStr(s, " ") > 0
        ss = Mid(s, 1, InStr(s, " ") - 1)
        sUnidad = sUnidad + ss + IIf(Right(ss, 1) = "s", "", IIf(InStr("aeiou", LCase(Right(ss, 1))) > 0, "s", "es")) + " "
        s = Trim(Mid(s, InStr(s, " ") + 1))
    Loop
    sUnidad = sUnidad + s + IIf(Right(s, 1) = "s", "", IIf(InStr("aeiou", LCase(Right(s, 1))) > 0, "s", "es"))
End If
ImporteLetra = Replace(Replace(ImporteLetra, "dieciseis", "dieciséis"), "veintiseis", "veintiséis")
If Len(ImporteLetra) > 0 Then
    ImporteLetra = UCase("( " + IIf(LCase(Right(sUnidad, 1)) = "a", Mid(ImporteLetra, 1, Len(ImporteLetra) - 1) + "a ", ImporteLetra) + IIf(dbl < 1, "cero ", "") + sUnidad + IIf(InStr("pesos", LCase(sUnidad)) > 0 Or Val(s2) > 0, " " + Mid(Trim(Str(Round(Val(s2) * 100, 0))) + "0", 1, 2) + "/100", "") + IIf(InStr("pesos", LCase(sUnidad)) > 0, " MN", "") + " )")
Else
    ImporteLetra = UCase("( " + IIf(LCase(Right(sUnidad, 1)) = "a", "a ", ImporteLetra) + IIf(dbl < 1, "cero ", "") + sUnidad + IIf(InStr("pesos", LCase(sUnidad)) > 0 Or Val(s2) > 0, " " + Mid(Trim(Str(Round(Val(s2) * 100, 0))) + "0", 1, 2) + "/100", "") + IIf(InStr("pesos", LCase(sUnidad)) > 0, " MN", "") + " )")
End If
End Function

Function NúmeroLetra(dbl As Double) As String ' Máximo tres dígitos
Dim s As String
NúmeroLetra = ImporteLetra(dbl, "peso")
NúmeroLetra = Mid(NúmeroLetra, InStr(NúmeroLetra, "( ") + 2)
If InStr(NúmeroLetra, " PESOS ") > 0 Then
    NúmeroLetra = LCase(Mid(NúmeroLetra, 1, InStr(NúmeroLetra, " PESOS ") - 1))
ElseIf InStr(NúmeroLetra, " PESO ") > 0 Then
    NúmeroLetra = "uno" 'LCase(Mid(NúmeroLetra, 1, InStr(NúmeroLetra, " PESO ") - 1))
End If
s = Str(dbl)
If InStr(s, ".") Then
    s = Mid(s, InStr(s, ".") - 1, 1)
Else
    s = Right(s, 1)
End If
If s = "1" And Len(NúmeroLetra) > 0 Then
    If Right(NúmeroLetra, 2) = "un" Then NúmeroLetra = NúmeroLetra + "o"
End If

End Function

Sub ValidaFecha(ByRef txtFecha As TextBox, yHora As Byte, sForma As String, Optional sDescripciónDato As String)
Static Y As Byte
If Len(Trim(txtFecha)) > 0 Then
    If Not IsDate(txtFecha) And Y = 0 Then
        If MsgBox("La fecha es incorrecta", vbOKCancel, "Error de captura") = vbCancel Then
            txtFecha = ""
        Else
            Y = 1
            'If Forms.Name = sForma Then
                If txtFecha.Visible And txtFecha.Enabled Then txtFecha.SetFocus
            'End If
        End If
        Exit Sub
    End If
    If IsDate(txtFecha) Then
        If yHora <> 0 Then
            txtFecha = Format(txtFecha, gsFormatoFechaHora)
        Else
            txtFecha = Format(txtFecha, gsFormatoFecha)
        End If
    End If
End If
Y = 0
End Sub

'Devuelve la fecha correspondiente a iDías hábiles después de la fecha dF
Function DíasHábiles(ByVal df As Date, ByVal iDías As Integer) As Date
Dim Y As Integer, df1 As Date, yy As Byte
df = CDate(Format(df, gsFormatoFecha))
If iDías = 0 Then
    DíasHábiles = df
    Exit Function
End If
df1 = df
Do While iDías > 0
    If iDías > 4 Then
        Y = Int(iDías / 5) * 7
        df = DateAdd("d", Y, df)
        Y = iDías Mod 5
    Else
        Y = iDías
    End If
    For yy = 1 To Y
        df = DateAdd("d", 1, df)
        If Weekday(df, vbMonday) > 5 Then
            yy = yy - 1
        End If
    Next
    iDías = Festivos(df1 + 1, df)
    df1 = df
Loop
If Weekday(df, vbMonday) > 5 Then
    df = df + 5 - Weekday(df, vbMonday)
    Do While Festivos(df, df) > 0 Or Weekday(df, vbMonday) > 5
        df = df - 1
    Loop
End If
DíasHábiles = df
End Function

'Devuelve los días festivos que se encuentran entre las fechas dIni y dFin  diferentes a sábado y domingo
Function Festivos(dIni As Date, dFin As Date) As Integer
Dim rs As Recordset, d As Date, i As Long, adors As New ADODB.Recordset
If gSQLACC = cyAccess Then
    Set rs = gdbmidb.OpenRecordset("select * from DíasFestivos", dbOpenSnapshot)
    Do While Not rs.EOF
        If rs!periódico Then
            d = DateAdd("yyyy", Year(dIni) - Year(rs!díafestivo), rs!díafestivo)
            Do While d <= dFin
                If dIni <= d And d <= dFin And Weekday(d, vbMonday) <= 5 Then
                    Festivos = Festivos + 1
                End If
                d = DateAdd("yyyy", 1, d)
            Loop
        Else
            If dIni <= rs!díafestivo And rs!díafestivo <= dFin And Weekday(rs!díafestivo, vbMonday) <= 5 Then
                Festivos = Festivos + 1
            End If
        End If
        rs.MoveNext
    Loop
Else 'SQL
    If adors.State > 0 Then adors.Close
    adors.Open "select * from DíasFestivos", gConSql, adOpenStatic, adLockReadOnly
    Do While Not adors.EOF
        If adors!periódico Then
            d = DateAdd("yyyy", Year(dIni) - Year(adors!díafestivo), adors!díafestivo)
            Do While d <= dFin
                If dIni <= d And d <= dFin And Weekday(d, vbMonday) <= 5 Then
                    Festivos = Festivos + 1
                End If
                d = DateAdd("yyyy", 1, d)
            Loop
        Else
            If dIni <= adors!díafestivo And adors!díafestivo <= dFin And Weekday(adors!díafestivo, vbMonday) <= 5 Then
                Festivos = Festivos + 1
            End If
        End If
        adors.MoveNext
        'frmProgreso.Show
    
    Loop
End If
End Function

'Devuelve en una cadena los días festivos que se encuentran entre las fechas dIni y dFin  diferentes a sábado y domingo
Function DíasFestivos(dIni As Date, dFin As Date) As String
Dim d As Date, i As Long, adors As New ADODB.Recordset
If adors.State > 0 Then adors.Close
adors.Open "select * from DíasFestivos", gConSql, adOpenStatic, adLockReadOnly
DíasFestivos = ","
Do While Not adors.EOF
    If adors!periódico Then
        d = DateAdd("yyyy", Year(dIni) - Year(adors!díafestivo), adors!díafestivo)
        Do While d <= dFin
            If dIni <= d And d <= dFin And Weekday(d, vbMonday) <= 5 Then
                DíasFestivos = DíasFestivos & ((Year(d) - 2000) * 31 * 12 + Month(d) * 31 + Day(d)) & ","
            End If
            d = DateAdd("yyyy", 1, d)
        Loop
    Else
        If dIni <= adors!díafestivo And adors!díafestivo <= dFin And Weekday(adors!díafestivo, vbMonday) <= 5 Then
            d = adors!díafestivo
            DíasFestivos = DíasFestivos & ((Year(d) - 2000) * 31 * 12 + Month(d) * 31 + Day(d)) & ","
        End If
    End If
    adors.MoveNext
Loop
'If Len(DíasFestivos) Then DíasFestivos = Mid(DíasFestivos, 1, Len(DíasFestivos) - 1)
End Function

'devuelve los días Hábiles que hay entre dFin y dIni
Function DíasHábilesEntre(dIni As Date, dFin As Date) As Long
Dim Y As Byte, yy As Byte, lDias As Long
If dIni > dFin Then
    DíasHábilesEntre = 0
    Exit Function
End If
DíasHábilesEntre = Int((dFin - dIni) / 7) * 5 - Festivos(dIni + 1, dFin - ((dFin - dIni) Mod 7))
dIni = DateAdd("d", Int((dFin - dIni) / 7) * 7, dIni)
Do While dIni <= dFin And DateDiff("d", dIni, dFin) > 0
    DíasHábilesEntre = DíasHábilesEntre + IIf(Weekday(dIni, vbMonday) <= 5 And Festivos(dIni, dIni) = 0, 1, 0)
    dIni = DateAdd("d", 1, dIni)
Loop
End Function


Function F_PreguntaConsecutivo(ByVal yTipo As Byte, ByVal sFolio As String) As Long
Dim s As String
If yTipo = 1 Then 'Registro nuevo Expediente
    s = "Expediente"
ElseIf yTipo = 2 Then 'Análisis Oficio de Emplazamiento
    s = "Oficio de Emplazamiento"
ElseIf yTipo = 3 Then 'Análisis Acuerdo de improcedencia
    s = "Acuerdo de Improcedencia"
ElseIf yTipo = 4 Then 'Análisis Oficio de sanción
    s = "Oficio de Sanción"
End If
l = 1
Do While True
    l = InputBox("Consecutivo inicial del " & s & " tipo: " & sFolio, "Especifique el consecutivo inicial del " & s, l)
    If l > 0 Then
        If MsgBox("Está seguro de iniciar el Número de " & s & " en " & l, vbYesNo + vbQuestion, "") = vbYes Then
            F_PreguntaConsecutivo = l
            Exit Function
        End If
    Else
        l = -1
        Exit Function
        'Call MsgBox("El valor del consecutivo debe ser mayor a cero", vbOKOnly + vbInformation, "")
    End If
Loop
End Function


Function F_PreguntaFolio(ByVal yTipo As Byte, ByVal sFolio As String) As String
Dim s As String, ss As String
If yTipo = 1 Then 'Registro nuevo Expediente
    s = "Expediente"
ElseIf yTipo = 2 Then 'Análisis Oficio de Emplazamiento
    s = "Oficio de Emplazamiento"
ElseIf yTipo = 3 Then 'Análisis Acuerdo de improcedencia
    s = "Acuerdo de Improcedencia"
ElseIf yTipo = 4 Then 'Análisis Oficio de sanción
    s = "Oficio de Sanción"
End If
ss = sFolio
Do While True
    ss = InputBox("Consecutivo inicial del " & s & " tipo: " & sFolio, "Especifique o verifique el folio " & s & " que deba generarse", ss)
    If Len(ss) > 0 Then
        If MsgBox("Está seguro de generar el folio de " & s & ": " & ss, vbYesNo + vbQuestion, "") = vbYes Then
            F_PreguntaFolio = ss
            Exit Function
        End If
    Else
        F_PreguntaFolio = "salir"
        Exit Function
        'Call MsgBox("El valor del consecutivo debe ser mayor a cero", vbOKOnly + vbInformation, "")
    End If
Loop
End Function


Function ValidaFolio(sFolio As String, sNuevoFolio As String, Optional bNoMensaje As Boolean) As Integer
Dim i As Integer
    If ReemplazaNúmerosXAst(sFolio) <> ReemplazaNúmerosXAst(sNuevoFolio) Then
        ValidaFolio = 0
        If Not bNoMensaje Then
            i = MsgBox("El folio ingresado: " & sFolio & " no tiene la misma estructura del folio esperado: " & sNuevoFolio & ", Está seguro de que este es correcto", vbYesNoCancel)
            If i = vbYes Then
                ValidaFolio = 2
            ElseIf i = vbNo Then
                ValidaFolio = 0
            ElseIf i = vbCancel Then
                ValidaFolio = -1
            End If
        End If
    Else
        ValidaFolio = 1
    End If
End Function

Function ReemplazaNúmerosXAst(sCad As String) As String
Dim i As Integer, s As String, i1 As Integer
s = sCad
i = 1
Do While i <= Len(s)
    If InStr("0123456789", Mid(s, i, 1)) > 0 Then
        If i1 = 0 Then
            i1 = 1
        End If
        s = Mid(s, 1, i - 1) & Mid(s, i + 1)
    Else
        If i1 > 0 Then
            i1 = 0
            s = Mid(s, 1, i - 1) & "*" & Mid(s, i)
        End If
        i = i + 1
    End If
Loop
If i1 > 0 Then
    s = Mid(s, 1, i - 1) & "*"
End If
ReemplazaNúmerosXAst = s
End Function
