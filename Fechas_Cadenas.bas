Attribute VB_Name = "Module1"
Public Const LETRAS = 1                     '1=Letras (sólo mayusculas A..Z)
Public Const LETRAS_MAYUS_MINUS_CADENA = 2  '2=Letras (Mayusculas y Minusculas)
Public Const NUMEROS = 3                    '3=Numeros (0..9)
Public Const CADENA = 5                     '5=String especial (ps_cadena) opcional
Public Const LETRAS_NUMEROS = 4             '4=Letras y Números
Public Const LETRAS_CADENA = 6              '6=Letras y String especial
Public Const NUMEROS_CADENA = 8             '8=Numeros y String especial
Public Const LETRAS_NUMEROS_CADENA = 9      '9=Letras,Numeros y String especial
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Function FU_strTrans(sCad As String, sBus As String, sRem As String) As String
Dim s As String

s = sCad
FU_strTrans = ""
Do While InStr(s, sBus) > 0
    FU_strTrans = FU_strTrans + Mid(s, 1, InStr(s, sBus) - 1) + sRem
    s = IIf(InStr(s, sBus) + Len(s_Bus) >= Len(s), "", Mid(s, InStr(s, sBus) + Len(sBus)))
Loop
FU_strTrans = FU_strTrans + s
'***********************************************************************************************************************
'*Sustituye un caracter(es) por otro dentro de la cadena base
'***********************************************************************************************************************
End Function


Sub PR_ControlaCaraterFechaKeyUp(ctl_Fecha As Control, Codigo As Integer)
Dim S_SepFecha As String
S_SepFecha = "-"    'Configurable el separadador

     If Codigo = 8 Then
        If Len(ctl_Fecha) = 6 Then
           ctl_Fecha.SelStart = Len(Trim(ctl_Fecha))
        ElseIf Len(ctl_Fecha) = 2 Then
           ctl_Fecha.SelStart = Len(Trim(ctl_Fecha))
        ElseIf Len(ctl_Fecha) = 3 Then
           ctl_Fecha.SelStart = Len(Trim(ctl_Fecha))
        End If
        Exit Sub
    End If
    
    'Coloca el caracter separador
    If Len(ctl_Fecha) = 2 Then
        ctl_Fecha = ctl_Fecha + S_SepFecha
        ctl_Fecha.SelStart = Len(Trim(ctl_Fecha))
    ElseIf Len(ctl_Fecha) = 5 Then
        If (Chr(Asc(Mid(ctl_Fecha, 5, 1))) >= "0") And (Chr(Asc(Mid(ctl_Fecha, 5, 1))) <= "9") Then ctl_Fecha = ctl_Fecha + S_SepFecha: ctl_Fecha.SelStart = Len(Trim(ctl_Fecha))
    ElseIf Len(ctl_Fecha) = 6 Then
        If Not (((Chr(Asc(Mid(ctl_Fecha, 6, 1))) >= "0") And (Chr(Asc(Mid(ctl_Fecha, 6, 1))) <= "9")) Or (Chr(Asc(Mid(ctl_Fecha, 6, 1))) = S_SepFecha)) Then ctl_Fecha = ctl_Fecha + S_SepFecha
        ctl_Fecha.SelStart = Len(Trim(ctl_Fecha))
    End If
'***********************************************************************************************************************
'*Rutina para colocar separador de caracteres en la fecha
'***********************************************************************************************************************
End Sub

Function FU_ValidaFecha(ctl_Fecha As Control) As Boolean
Dim B_FechaValida As Boolean, S_Cadmes As String, I_DiaFeb As Integer
Dim S_CarDia As Variant, S_CarMes As Variant, I_PosMes As Integer, S_MesCarac As String

S_Cadmes = "EneFebMarAbrMayJunJulAgoSepOctNovDic"
B_FechaValida = True
I_DiaFeb = 28
I_PosMes = 0
S_CarDia = Mid(ctl_Fecha, 1, 2)             'Caracteres para el día
S_CarMes = Mid(ctl_Fecha, 4, 2)             'Caracteres para el mes

If Len(Trim(ctl_Fecha)) < 8 Or Not IsNumeric(S_CarDia) Then
    ctl_Fecha.SelStart = 0
    ctl_Fecha.SelLength = Len(Trim(ctl_Fecha))
    MsgBox "Fecha inválida.", 0 + 32, "Validación de captura."
    B_FechaValida = False
Else
    If IsNumeric(S_CarMes) Then     'Cuando el mes es Número
        If Val(S_CarMes) > 12 Or Val(S_CarMes) = 0 Then
            MsgBox "Fecha inválida.", 0 + 32, "Ver el No. de mes." & S_CarMes
            FU_ValidaFecha = False: Exit Function
        End If
    Else                             'Cuando tiene el mes con formato de letra
        S_MesCarac = Mid(ctl_Fecha, 4, 3)
        S_MesCarac = UCase(Left(S_MesCarac, 1)) & LCase(Right(S_MesCarac, Len(S_MesCarac) - 1))
        I_PosMes = InStr(1, S_Cadmes, S_MesCarac)
        If Val(I_PosMes) <> 0 Then
            S_CarMes = Val(Int(Val(I_PosMes) / 3) + 1)
        Else
            MsgBox "Fecha inválida.", 0 + 32, "Ver el No. de mes." & S_CarMes
            FU_ValidaFecha = False: Exit Function
        End If
    End If
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    If Val(S_CarDia) > 31 Or Val(S_CarDia) = 0 Then
        MsgBox "Fecha inválida.", 0 + 32, "Ver el No. de día." & S_CarDia
        FU_ValidaFecha = False: Exit Function
    End If
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    If Len(Trim(ctl_Fecha)) = 8 Then
        I_Ano = "20" & Trim(Mid(ctl_Fecha, 7, 2))
    ElseIf Len(Trim(ctl_Fecha)) = 9 Then
        I_Ano = "20" & Trim(Mid(ctl_Fecha, 8, 2))
    ElseIf Len(Trim(ctl_Fecha)) = 10 Then
        If IsNumeric(Mid(ctl_Fecha, 4, 2)) Then                             '04-Abr-2001
            I_Ano = Trim(Mid(ctl_Fecha, 7, Len(ctl_Fecha) - 6))
        Else                                                    '04-Abr-2001
            I_Ano = "20" & Trim(Mid(ctl_Fecha, 9, 2))           '04-Abr-2001
        End If                                                  '04-Abr-2001
    ElseIf Len(Trim(ctl_Fecha)) = 11 Then
        I_Ano = Trim(Mid(ctl_Fecha, 8, Len(ctl_Fecha) - 7))
    Else
        MsgBox "Fecha inválida.", 0 + 32, "Ver el No. de año."
        FU_ValidaFecha = False: Exit Function
    End If
    'If Val(I_Ano) < 2000 Or Val(I_Ano) <= 2100 Then MsgBox "Verifique el No. del año " & I_Ano & ".", 0 + 48, "Ver el No. de año."
    If Val(Val(I_Ano) Mod 4) = 0 Then I_DiaFeb = 29             'Para el año bisiesto
    
    'Mes con 28 ó 29 días
    If Val(S_CarMes) = 2 And Val(S_CarDia) > I_DiaFeb Then
        MsgBox "Fecha inválida.", 0 + 32, "Ver el No. de día." & S_CarDia
        FU_ValidaFecha = False: Exit Function
    End If
    'Meses con 30 días
    If (Val(S_CarMes) = 4 Or Val(S_CarMes) = 6 Or Val(S_CarMes) = 9 Or Val(S_CarMes) = 11) And Val(S_CarDia) > 30 Then
        MsgBox "Fecha inválida.", 0 + 32, "Ver el No. de día." & S_CarDia
        FU_ValidaFecha = False: Exit Function
    End If
    'Meses con 31 días
    If (Val(S_CarMes) = 1 Or Val(S_CarMes) = 3 Or Val(S_CarMes) = 5 Or Val(S_CarMes) = 7 Or Val(S_CarMes) = 8 Or Val(S_CarMes) = 10 Or Val(S_CarMes) = 12) And Val(S_CarDia) > 31 Then
        MsgBox "Fecha inválida.", 0 + 32, "Ver el No. de día." & S_CarDia
        FU_ValidaFecha = False: Exit Function
    End If
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    If IsDate(ctl_Fecha) Then
        ctl_Fecha = FU_CambioFormatoMes(Trim(ctl_Fecha))
    Else
        ctl_Fecha.SelStart = 0
        ctl_Fecha.SelLength = Len(Trim(ctl_Fecha))
        MsgBox "Fecha inválida.", 0 + 32, "Validación de captura."
        B_FechaValida = False
    End If
End If
FU_ValidaFecha = B_FechaValida
'*******************************************************************************************
'*Sirve para validar la fecha de un control de captura de fecha
'*******************************************************************************************
End Function
Function Proceso_KeyPress(KeyAscii, inicio, Ctl_ControlX As Control) As Integer
Dim Bo_MesLetra As Boolean
Bo_MesLetra = False

Proceso_KeyPress = KeyAscii

If KeyAscii = 13 Then       'RC
   Proceso_KeyPress = KeyAscii
   Exit Function
ElseIf KeyAscii = 45 Then   'Bloque de If agregado 15-Jul-2008
    If inicio = 0 Or inicio = 1 Or inicio >= 8 Then
        Proceso_KeyPress = 0
        Exit Function
    End If
End If

If KeyAscii = 8 Then
    KeyAscii = 8
    Exit Function
ElseIf (KeyAscii < 47 Or KeyAscii > 57) And inicio < 3 Then
    Proceso_KeyPress = 0
    Exit Function
ElseIf (KeyAscii = 32) Then     'Agregada 15-Jul-2008
    Proceso_KeyPress = 0
    Exit Function
End If

Select Case Val(inicio + 1)
    Case 1
        If Not Chr(KeyAscii) Like "[0-3]" Then Proceso_KeyPress = 0
    Case 2
        If Val(Right(Trim(Ctl_ControlX), 1)) = 0 Then
            If Not Chr(KeyAscii) Like "[1-9]" Then Proceso_KeyPress = 0
        ElseIf Val(Right(Trim(Ctl_ControlX), 1)) > 0 And Val(Right(Trim(Ctl_ControlX), 1)) < 3 Then
            If Not Chr(KeyAscii) Like "[0-9]" Then Proceso_KeyPress = 0
        Else
            If Not Chr(KeyAscii) Like "[0-1]" Then Proceso_KeyPress = 0
        End If
    Case 3
        If KeyAscii <> 45 Then Proceso_KeyPress = 0
    Case 4
        If IsNumeric(Chr(KeyAscii)) Then
            If Not Chr(KeyAscii) Like "[0-1]" Then Proceso_KeyPress = 0
        Else
            If Not UCase(Chr(KeyAscii)) Like "[EFMAJSOND]" Then Proceso_KeyPress = 0
        End If
    Case 5
        If IsNumeric(Right(Trim(Ctl_ControlX), 1)) Then
            If Val(Right(Trim(Ctl_ControlX), 1)) = 0 Then   'Cero para el primer caracter del mes
                If Not Chr(KeyAscii) Like "[1-9]" Then Proceso_KeyPress = 0
            Else                                            'Uno para el primer caracter del mes
                If Not Chr(KeyAscii) Like "[0-2]" Then Proceso_KeyPress = 0     '30/05/01 [1-2]
            End If
        Else                                                'Cuando ya contiene una letra
            If Not UCase(Chr(KeyAscii)) Like "[NEABUGCOI]" Then Proceso_KeyPress = 0
        End If
    Case 6
        If Not IsNumeric(Right(Trim(Ctl_ControlX), 1)) Then
            If Not UCase(Chr(KeyAscii)) Like "[EBRYNLOPTVC]" Then Proceso_KeyPress = 0
        End If
   Case 7                                                   'Para el Año
        If Not Chr(KeyAscii) Like "[0-9]" Then Proceso_KeyPress = 0
   Case 8
        If Val(Right(Trim(Ctl_ControlX), 1)) = 0 Then   'Cero para el primer caracter del año
            If Not Chr(KeyAscii) Like "[0-9]" Then Proceso_KeyPress = 0
        Else
            If Not Chr(KeyAscii) Like "[0-9]" Then Proceso_KeyPress = 0
        End If
        
End Select
'*******************************************************************************************
'*Sirve para validar los caracteres del día,mes y año
'*******************************************************************************************
End Function
Function FU_CambioFormatoMes(V_Fecha As Variant, Optional S_CarSeparador As String = "-") As String
Dim S_Mes As String, S_Dia As String, S_Cadmes As String, S_CarMes As Variant, I_PosMes, I_Ano

S_Cadmes = "EneFebMarAbrMayJunJulAgoSepOctNovDic"
S_CarMes = Mid(V_Fecha, 4, 2)

If IsNumeric(S_CarMes) Then     'Cuando el mes es Número
    S_CarMes = Val(S_CarMes)
Else                            'Cuando tiene el mes con formato de letra
    S_MesCarac = Mid(V_Fecha, 4, 3)
    S_MesCarac = UCase(Left(S_MesCarac, 1)) & LCase(Right(S_MesCarac, Len(S_MesCarac) - 1))
    I_PosMes = InStr(1, S_Cadmes, S_MesCarac)
    If Val(I_PosMes) <> 0 Then
        S_CarMes = Val(Int(Val(I_PosMes) / 3) + 1)
    End If
End If
    
Select Case S_CarMes
    Case 1
        S_Mes = "Ene"
    Case 2
        S_Mes = "Feb"
    Case 3
        S_Mes = "Mar"
    Case 4
        S_Mes = "Abr"
    Case 5
        S_Mes = "May"
    Case 6
        S_Mes = "Jun"
    Case 7
        S_Mes = "Jul"
    Case 8
        S_Mes = "Ago"
    Case 9
        S_Mes = "Sep"
    Case 10
        S_Mes = "Oct"
    Case 11
        S_Mes = "Nov"
    Case 12
        S_Mes = "Dic"
End Select

If Val(Mid(V_Fecha, 1, 2)) < 10 Then
    S_Dia = "0" & Trim(Str(Mid(V_Fecha, 1, 2)))
Else
    S_Dia = Trim(Str(Mid(V_Fecha, 1, 2)))
End If

If Len(Trim(V_Fecha)) = 8 Then
    If Val(Trim(Mid(V_Fecha, 7, 2))) >= 90 Then
        I_Ano = "19" & Trim(Mid(V_Fecha, 7, 2))
    Else
        I_Ano = "20" & Trim(Mid(V_Fecha, 7, 2))     'Original 22/09/2001
    End If
ElseIf Len(Trim(V_Fecha)) = 9 Then
    I_Ano = "20" & Trim(Mid(V_Fecha, 8, 2))
ElseIf Len(Trim(V_Fecha)) = 10 Then
    If IsNumeric(Mid(V_Fecha, 4, 2)) Then
        I_Ano = Trim(Mid(V_Fecha, 7, Len(V_Fecha) - 6))
    Else
        I_Ano = "20" & Trim(Mid(V_Fecha, 9, 2))
    End If
ElseIf Len(Trim(V_Fecha)) = 11 Then
    I_Ano = Trim(Mid(V_Fecha, 8, Len(V_Fecha) - 7))
End If

If Len(S_CarSeparador) > 0 Then S_CarSeparador = Left(S_CarSeparador, 1)
FU_CambioFormatoMes = Trim(S_Dia & S_CarSeparador & S_Mes & S_CarSeparador & I_Ano)
End Function

Function FU_cbo_Inicializa(pctlControl As Control, _
                       pstrNomTabla As String, _
                       pstrNomClave As String, _
                       pstrNomDescrip As String, _
                       Optional pstrWhere As String = "", _
                       Optional pintOrderBy As Byte = 2, _
                       Optional pbytDistinct As Boolean = True)
'******************************************************************************************
'Función                    : gp_cbo_Inicializa
'Autor                      :
'Descripción                : Llena ComboBox o ListBox con datos de una tabla.
'                             Copia Clave en el itemdata y Descripcion en el
'                             texto del control
'Fecha de Creación          :
'Fecha de Liberación        :
'Fecha de Modificación      :
'Autor de la Modificación   :
'Usuario que solicita la modificación:
'
' Parám.:   pctlControl    - Control a llenar.
'           pstrNomTabla   - Nombre de la tabla a seleccionar.
'           pstrNomClave   - Nombre del atributo clave en la tabla.
'           pstrWhere      - Expresión a considerar como criterios
'                            de selección de los registros.
'           pstrOrderBy    - Numero del campo por el cual se ordena el
'                            combo o lista (1-clave, 2-descripción).
'           pbytDistinct   - Valor booleano que indica si se incluye la
'                            clausula DISTINCT en la selección de
'                            registros.
'*****************************************************************************************
   Dim lrdoCatalogo As ADODB.Recordset
   Dim lstrQuery    As String

   If Not TypeOf pctlControl Is ComboBox And _
      Not TypeOf pctlControl Is ListBox Then Exit Function

   lstrQuery = "SELECT " & IIf(pbytDistinct, "DISTINCT ", "") & _
              pstrNomClave & "," & pstrNomDescrip & _
              " FROM " & pstrNomTabla & _
              IIf(pstrWhere = "", "", " WHERE " & pstrWhere) & _
              " ORDER BY " & IIf(pintOrderBy = 2, "2", "1")

   On Error GoTo ErrRecupera
   
   Set lrdoCatalogo = New ADODB.Recordset
   lrdoCatalogo.Open lstrQuery, gcn

   pctlControl.Clear
   If lrdoCatalogo.RecordCount > 0 Then
      Do While Not lrdoCatalogo.EOF
         pctlControl.AddItem (Trim(lrdoCatalogo(1)))
         pctlControl.ItemData(pctlControl.NewIndex) = lrdoCatalogo(0)
         lrdoCatalogo.MoveNext
      Loop
   End If
   Exit Function
   
ErrRecupera:
If Err.Number <> 0 Then
   Exit Function
End If
'***********************************************************************************************************************
'*
'***********************************************************************************************************************
End Function
Sub PR_ToolTipText(TipoControl, S_NombreCtl As Control, Optional S_Mensa As String = "")
Select Case UCase(Trim(TipoControl))
Case "ETIQUETA"
    If S_Mensa = "" Then
        S_NombreCtl.ToolTipText = S_NombreCtl.Caption
    Else
        S_NombreCtl.ToolTipText = Trim(S_Mensa)
    End If
Case "TEXTO"
    If S_Mensa = "" Then
        S_NombreCtl.ToolTipText = S_NombreCtl.Text
    Else
        S_NombreCtl.ToolTipText = Trim(S_Mensa)
    End If
End Select
'********************************************************************************************
'*Muestra el contenido de un control ó un mensaje en forma de ToolTipText
'*Recibe como parametro el control al cuál se debe de asignar el texto y es opcional el
'*mensaje.
'********************************************************************************************
End Sub
Function Fu_ValidaTeclaNew(KeyAscii) As Integer
Fu_ValidaTeclaNew = KeyAscii
Select Case KeyAscii
    Case 34             '""
        Fu_ValidaTeclaNew = 0
    Case 39             ''
        Fu_ValidaTeclaNew = 0
    Case 91             '[
        Fu_ValidaTeclaNew = 0
    Case 93             ']
        Fu_ValidaTeclaNew = 0
    Case 123 To 159     '{ en adelante
        Fu_ValidaTeclaNew = 0
    Case 168             '¨
        Fu_ValidaTeclaNew = 0
    Case 172             '¬
        Fu_ValidaTeclaNew = 0
    Case 176             '°
        Fu_ValidaTeclaNew = 0
    Case 180             '´
        Fu_ValidaTeclaNew = 0
End Select

'*******************************************************************************************
'*Esta función sirve para validar que el cararter tecleado es válido
'*Recibe como parametro el número equivalente en la tabla ANSI,es llamada desde el evento
'*KeyPress
'*******************************************************************************************
End Function
Function FU_AgregaCeros_Izquierda(N_Totdig As Integer, N_Cantidad As Long) As String
Dim s_ceros As String, N_Tama As Integer

FU_AgregaCeros_Izquierda = Trim(Str(N_Cantidad))
N_Tama = Len(Trim((N_Cantidad)))
If N_Tama >= N_Totdig Then Exit Function
s_ceros = ""
'N_Tama = Len(Trim((N_Cantidad)))
For i = 1 To (N_Totdig - N_Tama)
    s_ceros = s_ceros & "0"
Next i
FU_AgregaCeros_Izquierda = s_ceros & Trim(Str(N_Cantidad))
'***********************************************************************************************************************
'Da formato a un digito y lo regresa con ceros a la izquierda
'***********************************************************************************************************************
End Function
Function FU_AgregaCeros_IzquierdaCad(N_Totdig As Integer, N_Cantidad) As String
Dim s_ceros As String, N_Tama As Integer

FU_AgregaCeros_IzquierdaCad = Trim(N_Cantidad)
N_Tama = Len(Trim((N_Cantidad)))
If N_Tama >= N_Totdig Then Exit Function
s_ceros = ""

For i = 1 To (N_Totdig - N_Tama)
    s_ceros = s_ceros & "0"
Next i
FU_AgregaCeros_IzquierdaCad = s_ceros & Trim(N_Cantidad)
'***********************************************************************************************************************
'Da formato a una cadena y lo regresa con ceros a la izquierda
'***********************************************************************************************************************
End Function
Function FU_ValidaDiaFebrero() As Integer
Dim F_Sistema
Dim I_DiaFeb As Integer, I_Ano As Integer

FU_ValidaDiaFebrero = 0
'-->    Lineas agregadas,
If Val(Trim(Frm_CuestionarioDin.Txt_CueDinMod1(1))) = 2 Then
    F_Sistema = FU_ExtraeFechaServer()                          'Traer la fecha del server
    If IsNumeric(Trim(Frm_CuestionarioDin.Txt_CueDinMod1(2))) Then
        I_Ano = 2000 + Val(Trim(Frm_CuestionarioDin.Txt_CueDinMod1(2)))
    Else
        I_Ano = Val(Year(F_Sistema))
    End If
    I_DiaFeb = 28
    If Val(I_Ano Mod 4) = 0 Then I_DiaFeb = 29   'Para el año bisiesto
    If Len(Trim(Frm_CuestionarioDin.Txt_CueDinMod1(0))) > 0 Then
        If IsNumeric(Trim(Frm_CuestionarioDin.Txt_CueDinMod1(0))) Then
            If Val(Trim(Frm_CuestionarioDin.Txt_CueDinMod1(0))) > I_DiaFeb Then
                MsgBox "El día no es válido para el mes de Febrero", 0 + 16, "Verificar"
                Frm_CuestionarioDin.Txt_CueDinMod1(0).BackColor = &H80FF&
                FU_ValidaDiaFebrero = -1
            End If
        End If
    End If
End If
'***********************************************************************************************************************
'*Para validar día del Mes de Febrero: 02-Mar-09
'*NOTA: Cuando se modifique esta función es necesario cambiarla tambien en el evento Txt_CueDinMod1_LostFocus(2)
'***********************************************************************************************************************
End Function
Function FU_ValidaDiaMeses_con30dias(I_NumMes) As Integer
Dim F_Sistema, S_NomMes As String
Dim I_Ano As Integer

FU_ValidaDiaMeses_con30dias = 0

If I_NumMes = 4 Or I_NumMes = 6 Or I_NumMes = 9 Or I_NumMes = 11 Then
    If Len(Trim(Frm_CuestionarioDin.Txt_CueDinMod1(0))) > 0 Then
        If IsNumeric(Trim(Frm_CuestionarioDin.Txt_CueDinMod1(0))) Then
            If Val(Trim(Frm_CuestionarioDin.Txt_CueDinMod1(0))) > 30 Then
                Select Case I_NumMes
                    Case 4
                        S_NomMes = "Abril"
                    Case 6
                        S_NomMes = "Junio"
                    Case 9
                        S_NomMes = "Septiembre"
                    Case 11
                        S_NomMes = "Noviembre"
                End Select
                MsgBox "El día no es válido para el mes de " & S_NomMes & ".", 0 + 16, "Verificar"
                Frm_CuestionarioDin.Txt_CueDinMod1(0).BackColor = &H80FF&
                FU_ValidaDiaMeses_con30dias = -1
            End If
        End If
    End If
End If
'***********************************************************************************************************************
'*Para validar día de los meses que tienen sólo 30 días: 20-Mar-09
'*NOTA: Cuando se modifique esta función es necesario cambiarla tambien en el evento Txt_CueDinMod1_LostFocus(2)
'***********************************************************************************************************************
End Function
