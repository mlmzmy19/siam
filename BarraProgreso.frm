VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BarraProgreso 
   ClientHeight    =   900
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   7230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   900
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   7080
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   6870
         _ExtentX        =   12118
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "BarraProgreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
