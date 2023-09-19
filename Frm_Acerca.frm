VERSION 5.00
Begin VB.Form Frm_Acerca 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca del sistema"
   ClientHeight    =   3750
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7680
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2588.317
   ScaleMode       =   0  'User
   ScaleWidth      =   7211.917
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   1005
      Left            =   2475
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "Frm_Acerca.frx":0000
      Top             =   1530
      Width           =   4920
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   2745
      TabIndex        =   0
      Top             =   3015
      Width           =   2085
   End
   Begin VB.Image Image1 
      Height          =   2475
      Left            =   45
      Picture         =   "Frm_Acerca.frx":0075
      Stretch         =   -1  'True
      Top             =   45
      Width           =   2490
   End
   Begin VB.Label EtiActualización 
      Caption         =   "Última actualización: marzo 2014"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3060
      TabIndex        =   4
      Top             =   1215
      Width           =   3900
   End
   Begin VB.Label Lbl_ACS 
      Alignment       =   2  'Center
      Caption         =   "CONDUSEF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   2505
      TabIndex        =   3
      Top             =   90
      Width           =   4755
   End
   Begin VB.Label lblTitle 
      Caption         =   "Sistema de Administrativo de Multas (SIAM)"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3105
      TabIndex        =   1
      Top             =   585
      Width           =   3765
   End
   Begin VB.Label lblVersion 
      Caption         =   "Versión 1.03"
      Height          =   225
      Left            =   3105
      TabIndex        =   2
      Top             =   945
      Width           =   3795
   End
End
Attribute VB_Name = "Frm_Acerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Opciones de seguridad de claves del Registro...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Tipos principales de claves del Registro...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Cadena Unicode terminada en Null
Const REG_DWORD = 4                      ' Número de 32 bits

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Public Sub StartSysInfo()
    MsgBox "Unidad de Desarrollo y Evaluación del Proceso Operativo, Ext. 6032, 6032", 0 + 64, ""
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Contador de bucle
    Dim rc As Long                                          ' Código de retorno
    Dim hKey As Long                                        ' Controlador a una clave de Registro abierta
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Tipo de dato de una clave de Registro
    Dim tmpVal As String                                    ' Almacén temporal de una valor de clave de Registro
    Dim KeyValSize As Long                                  ' Tamaño de la variable de la clave de Registro
    '------------------------------------------------------------
    ' Abre la clave de Registro en la raíz {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Abre la clave de Registro
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Trata el error...
    
    tmpVal = String$(1024, 0)                               ' Asigna espacio para la variable
    KeyValSize = 1024                                       ' Marca el tamaño de la variable
    
    '------------------------------------------------------------
    ' Recupera valores de claves de Registro...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Obtiene o crea un valor de clave
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Trata el error
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 agrega una cadena terminada en Null...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Se encontró Null, se extrae de la cadena
    Else                                                    ' WinNT no tiene una cadena terminada en Null...
        tmpVal = Left(tmpVal, KeyValSize)                   ' No se encontró Null, sólo se extrae la cadena
    End If
    '------------------------------------------------------------
    ' Determina el tipo de valor de la clave para conversión...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Busca tipos de datos...
    Case REG_SZ                                             ' Tipo de dato de la cadena de la clave de Registro
        KeyVal = tmpVal                                     ' Copia el valor de la cadena
    Case REG_DWORD                                          ' El tipo de dato de la cadena de la clave es Double Word
        For i = Len(tmpVal) To 1 Step -1                    ' Convierte cada byte
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Genera el valor carácter a carácter
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convierte Double Word a String
    End Select
    
    GetKeyValue = True                                      ' Vuelve con éxito
    rc = RegCloseKey(hKey)                                  ' Cierra la clave de Registro
    Exit Function                                           ' Salir
    
GetKeyError:      ' Restaurar después de que ocurra un error...
    KeyVal = ""                                             ' Establece el valor de retorno para una cadena vacía
    GetKeyValue = False                                     ' Devuelve un error
    rc = RegCloseKey(hKey)                                  ' Cierra la clave de Registro
End Function

Private Sub Form_Load()
'lblVersion.Caption = App.FileDescription & " Versión: " & App.Revision
EtiActualización.Caption = "Última Actualización: " & Format(gdVersión, gsFormatoFecha)
lblVersion.Caption = App.EXEName & " - " & gsVersión
End Sub
