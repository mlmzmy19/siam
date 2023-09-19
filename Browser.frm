VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form Browser 
   Caption         =   "Consulta de documento de Estrados Electrónicos"
   ClientHeight    =   10320
   ClientLeft      =   3060
   ClientTop       =   3345
   ClientWidth     =   15105
   LinkTopic       =   "Form1"
   ScaleHeight     =   10320
   ScaleWidth      =   15105
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Invocar nuevamente"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   10272
      Left            =   108
      TabIndex        =   0
      Top             =   0
      Width           =   14952
      ExtentX         =   26379
      ExtentY         =   18124
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6180
      Top             =   1500
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   2670
      Top             =   2325
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser.frx":005E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser.frx":00BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser.frx":011A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser.frx":0178
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser.frx":01D6
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public bCerrar As Boolean 'Indica si debe cerrar o ocultar la ventana
Public StartingAddress As String
Public yÚnicavez As Byte
Dim mbDontNavigateNow As Boolean


Private Sub brwWebBrowser_WindowClosing(ByVal IsChildWindow As Boolean, Cancel As Boolean)
Unload Me
End Sub

Private Sub cmd1_Click()
brwWebBrowser.Navigate gsWWW
End Sub

Private Sub cmd2_Click()
Unload Me
End Sub

Sub Form_Activate()
If yÚnicavez = 0 Or yÚnicavez = 200 Then
    brwWebBrowser.Navigate gsWWW
    
End If
yÚnicavez = 1
End Sub

Private Sub Form_Resize()
'    cboAddress.Width = Me.ScaleWidth - 100
    brwWebBrowser.Width = IIf(Me.ScaleWidth > 100, Me.ScaleWidth - 100, 0)
    brwWebBrowser.Height = Me.ScaleHeight '- (picAddress.Top + picAddress.Height) - 100
End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else
        Me.Caption = "Trabajando..."
    End If
End Sub
