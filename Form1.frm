VERSION 5.00
Object = "*\AUltiFrameControl.vbp"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UltiFrame Demo"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8025
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin UltiFrameControl.UltiFrame UltiFrame3 
      Height          =   2175
      Left            =   120
      TabIndex        =   16
      Top             =   4440
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3836
      Caption         =   "Text"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdFont 
         Caption         =   "..."
         Height          =   255
         Left            =   3480
         TabIndex        =   31
         Top             =   1440
         Width           =   255
      End
      Begin VB.TextBox txtFontName 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1440
         Width           =   1815
      End
      Begin VB.ComboBox cboFont3D 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   1560
         List            =   "Form1.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdForeColor 
         Height          =   255
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Left            =   1560
         TabIndex        =   22
         Top             =   720
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   375
         Left            =   1320
         TabIndex        =   17
         Top             =   960
         Width           =   2535
         Begin VB.OptionButton optCaptionStyle 
            Caption         =   "Wrapped"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   19
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton optCaptionStyle 
            Caption         =   "Standard"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   18
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Label lblDescription 
         Caption         =   "Font Name"
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblDescription 
         Caption         =   "3D Font"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblDescription 
         Caption         =   "ForeColor"
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblDescription 
         Caption         =   "Caption"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblDescription 
         Caption         =   "Caption Style"
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   1335
      End
   End
   Begin UltiFrameControl.UltiFrame UltiFrame2 
      Height          =   4215
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   7435
      Caption         =   "Appearance"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdBorderColor 
         Height          =   255
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3840
         Width           =   495
      End
      Begin VB.CommandButton cmdShadowColor 
         Height          =   255
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton cmdHighlightColor 
         Height          =   255
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdBackColor 
         Height          =   255
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   840
         Width           =   495
      End
      Begin VB.ComboBox cboAppearance 
         Height          =   315
         ItemData        =   "Form1.frx":0004
         Left            =   1560
         List            =   "Form1.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox cboBorderStyle 
         Height          =   315
         ItemData        =   "Form1.frx":0008
         Left            =   1560
         List            =   "Form1.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   1200
         Width           =   2535
         Begin VB.OptionButton optBackStyle 
            Caption         =   " Transparent"
            Height          =   315
            Index           =   0
            Left            =   1080
            TabIndex        =   8
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton optBackStyle 
            Caption         =   "Opaque"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.ComboBox cboAlignment 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2520
         Width           =   2175
      End
      Begin VB.HScrollBar hscCornerRadius 
         Height          =   255
         Left            =   1560
         Max             =   50
         TabIndex        =   4
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label lblDescription 
         Caption         =   "Flat Border Color"
         Height          =   315
         Index           =   13
         Left            =   120
         TabIndex        =   37
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label lblDescription 
         Caption         =   "Border Shadow Color"
         Height          =   435
         Index           =   12
         Left            =   120
         TabIndex        =   35
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblDescription 
         Caption         =   "Border Highlight Color"
         Height          =   435
         Index           =   11
         Left            =   120
         TabIndex        =   33
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblDescription 
         Caption         =   "BackColor"
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblDescription 
         Caption         =   "Appearance"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblDescription 
         Caption         =   "Flat Border Style"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label lblDescription 
         Caption         =   "BackStyle"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblDescription 
         Caption         =   "Caption Alignment"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lblDescription 
         Caption         =   "Corner Radius"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   3000
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7440
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   1
      Min             =   8
   End
   Begin UltiFrameControl.UltiFrame UltiFrame1 
      Height          =   2895
      Left            =   4560
      TabIndex        =   0
      Top             =   480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   5106
      Caption         =   "UltiFrame1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About"
         Height          =   735
         Left            =   840
         TabIndex        =   2
         Top             =   1665
         Width           =   1575
      End
      Begin VB.CommandButton cmdPlySnd 
         Caption         =   "PlaySound"
         Height          =   735
         Left            =   840
         TabIndex        =   1
         Top             =   570
         Width           =   1575
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3360
      Left            =   4320
      Picture         =   "Form1.frx":000C
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboAlignment_Click()
    UltiFrame1.Alignment = cboAlignment.ItemData(cboAlignment.ListIndex)
End Sub

Private Sub cboAppearance_Click()
    UltiFrame1.Appearance = cboAppearance.ItemData(cboAppearance.ListIndex)
End Sub

Private Sub cboBorderStyle_Click()
    UltiFrame1.BorderStyle = cboBorderStyle.ItemData(cboBorderStyle.ListIndex)
End Sub

Private Sub cboFont3D_Click()
    UltiFrame1.Font3D = cboFont3D.ItemData(cboFont3D.ListIndex)
End Sub

Private Sub cmdAbout_Click()
    UltiFrame1.AboutBox
End Sub

Private Sub cmdBackColor_Click()
    CommonDialog1.ShowColor
    UltiFrame1.BackColor = CommonDialog1.Color
    cmdBackColor.BackColor = CommonDialog1.Color
End Sub

Private Sub cmdBorderColor_Click()
    CommonDialog1.ShowColor
    UltiFrame1.BorderColor = CommonDialog1.Color
    cmdBorderColor.BackColor = CommonDialog1.Color
End Sub

Private Sub cmdFont_Click()
    CommonDialog1.ShowFont
    UltiFrame1.Font = CommonDialog1.FontName
    txtFontName.Text = CommonDialog1.FontName
    UltiFrame1.Refresh
End Sub

Private Sub cmdForeColor_Click()
    CommonDialog1.ShowColor
    UltiFrame1.ForeColor = CommonDialog1.Color
    cmdForeColor.BackColor = CommonDialog1.Color
End Sub

Private Sub cmdHighlightColor_Click()
    CommonDialog1.ShowColor
    UltiFrame1.BorderHighLightColor = CommonDialog1.Color
    cmdHighlightColor.BackColor = CommonDialog1.Color
End Sub

Private Sub cmdPlySnd_Click()
    UltiFrame1.PlaySoundFile "SystemStart", ufPlaySystemSound
End Sub

Private Sub FillCaptionAlignment(Box As ComboBox)
    Box.AddItem "Top Left"
    Box.ItemData(Box.NewIndex) = ufTopLeft
    Box.ListIndex = Box.NewIndex
    Box.AddItem "Top Center"
    Box.ItemData(Box.NewIndex) = ufTopCenter
    Box.AddItem "Top Right"
    Box.ItemData(Box.NewIndex) = ufTopRight
    Box.AddItem "Bottom Left"
    Box.ItemData(Box.NewIndex) = ufBottomLeft
    Box.AddItem "Bottom Center"
    Box.ItemData(Box.NewIndex) = ufBottomCenter
    Box.AddItem "Bottom Right"
    Box.ItemData(Box.NewIndex) = ufBottomRight
End Sub

Private Sub FillAppearance(Box As ComboBox)
    Box.AddItem "Flat"
    Box.ItemData(Box.NewIndex) = ufFlat
    Box.AddItem "Etched"
    Box.ItemData(Box.NewIndex) = ufEtched
    Box.ListIndex = Box.NewIndex
    Box.AddItem "Bumped"
    Box.ItemData(Box.NewIndex) = ufBump
End Sub

Private Sub FillBorderStyle(Box As ComboBox)
    Box.AddItem "Solid"
    Box.ItemData(Box.NewIndex) = ufSolid
    Box.ListIndex = Box.NewIndex
    Box.AddItem "Dash"
    Box.ItemData(Box.NewIndex) = ufDash
    Box.AddItem "Dot"
    Box.ItemData(Box.NewIndex) = ufDot
    Box.AddItem "DashDot"
    Box.ItemData(Box.NewIndex) = ufDashDot
    Box.AddItem "DashDotDot"
    Box.ItemData(Box.NewIndex) = ufDashDotDot
End Sub

Private Sub FillFont3D(Box As ComboBox)
    Box.AddItem "No 3DFont"
    Box.ItemData(Box.NewIndex) = ufNoneFont3D
    Box.ListIndex = Box.NewIndex
    Box.AddItem "Raised Light"
    Box.ItemData(Box.NewIndex) = ufRaisedLight
    Box.AddItem "Raised Heavy"
    Box.ItemData(Box.NewIndex) = ufRaisedHeavy
    Box.AddItem "InsetLight"
    Box.ItemData(Box.NewIndex) = ufInsetLight
    Box.AddItem "InsetHeavy"
    Box.ItemData(Box.NewIndex) = ufInsetHeavy
    Box.AddItem "DropShadow"
    Box.ItemData(Box.NewIndex) = ufDropShadow
End Sub

Private Sub cmdShadowColor_Click()
    CommonDialog1.ShowColor
    UltiFrame1.BorderShadowColor = CommonDialog1.Color
    cmdShadowColor.BackColor = CommonDialog1.Color
End Sub

Private Sub Form_Load()

Call FillAppearance(cboAppearance)
Call FillBorderStyle(cboBorderStyle)
Call FillCaptionAlignment(cboAlignment)
Call FillFont3D(cboFont3D)
cmdBackColor.BackColor = UltiFrame1.BackColor
cmdBorderColor.BackColor = UltiFrame1.BorderColor
cmdForeColor.BackColor = UltiFrame1.ForeColor
cmdHighlightColor.BackColor = UltiFrame1.BorderHighLightColor
cmdShadowColor.BackColor = UltiFrame1.BorderShadowColor
hscCornerRadius.Min = 0
hscCornerRadius.Max = 50
hscCornerRadius.Value = 0
optBackStyle(1).Value = True
optCaptionStyle(0).Value = True
txtCaption.Text = UltiFrame1.Caption
txtFontName.Text = UltiFrame1.Font

End Sub

Private Sub hscCornerRadius_Change()
    UltiFrame1.CornerRadius = CLng(hscCornerRadius.Value)
End Sub

Private Sub optBackStyle_Click(Index As Integer)
    UltiFrame1.BackStyle = Abs(CLng(optBackStyle(1).Value))
End Sub

Private Sub optCaptionStyle_Click(Index As Integer)
    UltiFrame1.CaptionStyle = Abs(CLng(optCaptionStyle(1).Value))
End Sub

Private Sub txtCaption_Change()
    UltiFrame1.Caption = txtCaption.Text
End Sub
