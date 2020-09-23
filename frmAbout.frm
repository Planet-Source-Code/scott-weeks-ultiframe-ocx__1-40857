VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5790
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   249
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   386
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "System Info..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4065
      TabIndex        =   3
      Top             =   3120
      Width           =   1470
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4065
      TabIndex        =   2
      Top             =   2670
      Width           =   1470
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Author"
      Height          =   195
      Left            =   1440
      TabIndex        =   7
      Top             =   1710
      Width           =   1140
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Develeoped By :"
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   1695
      Width           =   1215
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Ver#"
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   1350
      Width           =   1140
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Version :"
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   1335
      Width           =   675
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   7
      X2              =   376
      Y1              =   169
      Y2              =   169
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   7
      X2              =   376
      Y1              =   168
      Y2              =   168
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   210
      TabIndex        =   1
      Top             =   2655
      Width           =   3315
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "UltiFrame OCX"
      BeginProperty Font 
         Name            =   "Cataneo BT"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3585
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblCopyright.Caption = App.LegalCopyright
    lblVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
    lblAuthor.Caption = App.CompanyName
End Sub
