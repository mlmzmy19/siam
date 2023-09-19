VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   Caption         =   "Asignar nuevo responsable asunto para Análisis"
   ClientHeight    =   5445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   465
      Left            =   4770
      TabIndex        =   7
      Top             =   4410
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   315
      TabIndex        =   4
      Top             =   3465
      Width           =   5595
   End
   Begin VB.TextBox Text1 
      Height          =   1185
      Left            =   270
      TabIndex        =   2
      Top             =   1440
      Width           =   7215
   End
   Begin VB.TextBox txtExp 
      Height          =   375
      Left            =   3105
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   510
      Left            =   1125
      TabIndex        =   6
      Top             =   4365
      Width           =   2445
      Caption         =   "Actualizar"
      Size            =   "4313;900"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label3 
      Caption         =   "Asignar nuevo responsable:"
      Height          =   240
      Left            =   315
      TabIndex        =   5
      Top             =   3150
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Datos del Asunto:"
      Height          =   330
      Left            =   315
      TabIndex        =   3
      Top             =   1170
      Width           =   2715
   End
   Begin VB.Label Label1 
      Caption         =   "Favor de especificar No. el Expediente:"
      Height          =   285
      Left            =   135
      TabIndex        =   0
      Top             =   405
      Width           =   2850
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim miAct As Integer 'Por Actualizar
Dim miRes As Integer 'Nueva clave del responsable quien podrá realizar el análisis

Private Sub cmdSalir_Click()
If miAct Then
    
End If
End Sub
