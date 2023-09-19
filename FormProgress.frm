VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Progress"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Frame frProgress 
      Caption         =   "Progress"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin MSComctlLib.ProgressBar pbProgress 
         Height          =   285
         Left            =   225
         TabIndex        =   8
         Top             =   1035
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lProgress 
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label lDestFilename 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   5
         Top             =   720
         Width           =   45
      End
      Begin VB.Label lSourceFilename 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   45
      End
      Begin VB.Label lProcessed 
         Caption         =   "Processed:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lDest 
         Caption         =   "Destination file:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1215
         WordWrap        =   -1  'True
      End
      Begin VB.Label lSource 
         Caption         =   "Source file:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Canceled As Boolean

Private Sub btnCancel_Click()
  Canceled = True
End Sub

Private Sub Form_Load()
  Canceled = False
End Sub

