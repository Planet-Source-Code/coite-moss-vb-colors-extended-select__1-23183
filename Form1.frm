VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label53 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Click to Change BackGround Color"
      Height          =   255
      Left            =   2520
      TabIndex        =   52
      Top             =   360
      Width           =   2595
   End
   Begin VB.Label Label52 
      BackColor       =   &H80000009&
      Caption         =   "Main Module by ""CodeJack"""
      Height          =   195
      Left            =   2820
      TabIndex        =   51
      Top             =   6780
      Width           =   2115
   End
   Begin VB.Label Label51 
      BackColor       =   &H80000009&
      Caption         =   "Coyote Code"
      Height          =   195
      Left            =   6360
      TabIndex        =   50
      Top             =   6840
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   6540
      Picture         =   "Form1.frx":0000
      Top             =   5940
      Width           =   645
   End
   Begin VB.Label Label50 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Mouse Over to View VB Color Value"
      Height          =   255
      Left            =   2460
      TabIndex        =   49
      Top             =   60
      Width           =   2715
   End
   Begin VB.Label Label49 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label49"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   6240
      TabIndex        =   48
      ToolTipText     =   "&H400040"
      Top             =   5520
      Width           =   1275
   End
   Begin VB.Label Label48 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label48"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   6240
      TabIndex        =   47
      ToolTipText     =   "&H800080"
      Top             =   4920
      Width           =   1275
   End
   Begin VB.Label Label47 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label47"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   6240
      TabIndex        =   46
      ToolTipText     =   "&HC000C0"
      Top             =   4320
      Width           =   1275
   End
   Begin VB.Label Label46 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label46"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   6240
      TabIndex        =   45
      ToolTipText     =   "&HFF00FF"
      Top             =   3720
      Width           =   1275
   End
   Begin VB.Label Label45 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label45"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   6240
      TabIndex        =   44
      ToolTipText     =   "&HFF00FF"
      Top             =   3120
      Width           =   1275
   End
   Begin VB.Label Label44 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label44"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   6240
      TabIndex        =   43
      ToolTipText     =   "&HFF80FF"
      Top             =   2520
      Width           =   1275
   End
   Begin VB.Label Label43 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label43"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   6240
      TabIndex        =   42
      ToolTipText     =   "&HFFC0FF"
      Top             =   1920
      Width           =   1275
   End
   Begin VB.Label Label42 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label42"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   6240
      TabIndex        =   41
      ToolTipText     =   "&H400000"
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label Label41 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label41"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   6240
      TabIndex        =   40
      ToolTipText     =   "&H800000"
      Top             =   720
      Width           =   1275
   End
   Begin VB.Label Label40 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label40"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   4740
      TabIndex        =   39
      ToolTipText     =   "&HC00000"
      Top             =   6120
      Width           =   1275
   End
   Begin VB.Label Label39 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label39"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   4740
      TabIndex        =   38
      ToolTipText     =   "&HFF0000"
      Top             =   5520
      Width           =   1275
   End
   Begin VB.Label Label38 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label38"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   4740
      TabIndex        =   37
      ToolTipText     =   "&HFF8080"
      Top             =   4920
      Width           =   1275
   End
   Begin VB.Label Label37 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label37"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   4740
      TabIndex        =   36
      ToolTipText     =   "&HFFC0C0"
      Top             =   4320
      Width           =   1275
   End
   Begin VB.Label Label36 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label36"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   4740
      TabIndex        =   35
      ToolTipText     =   "&H404000"
      Top             =   3720
      Width           =   1275
   End
   Begin VB.Label Label35 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label35"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   4740
      TabIndex        =   34
      ToolTipText     =   "&H808000"
      Top             =   3120
      Width           =   1275
   End
   Begin VB.Label Label34 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label34"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   4740
      TabIndex        =   33
      ToolTipText     =   "&HC0C000"
      Top             =   2520
      Width           =   1275
   End
   Begin VB.Label Label33 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label33"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   4740
      TabIndex        =   32
      ToolTipText     =   "&HFFFF00"
      Top             =   1920
      Width           =   1275
   End
   Begin VB.Label Label32 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label32"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   4740
      TabIndex        =   31
      ToolTipText     =   "&HFFFF80"
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label Label31 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label31"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   4740
      TabIndex        =   30
      ToolTipText     =   "&HFFFFC0"
      Top             =   720
      Width           =   1275
   End
   Begin VB.Label Label30 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label30"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   3240
      TabIndex        =   29
      ToolTipText     =   "&H4000&"
      Top             =   6120
      Width           =   1275
   End
   Begin VB.Label Label29 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label29"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   3240
      TabIndex        =   28
      ToolTipText     =   "&H8000&"
      Top             =   5520
      Width           =   1275
   End
   Begin VB.Label Label28 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label28"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   3240
      TabIndex        =   27
      ToolTipText     =   "&HC000&"
      Top             =   4920
      Width           =   1275
   End
   Begin VB.Label Label27 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label27"
      ForeColor       =   &H00000040&
      Height          =   435
      Left            =   3240
      TabIndex        =   26
      ToolTipText     =   "&HFF00&"
      Top             =   4320
      Width           =   1275
   End
   Begin VB.Label Label26 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label26"
      ForeColor       =   &H00000040&
      Height          =   435
      Left            =   3240
      TabIndex        =   25
      ToolTipText     =   "&H80FF80"
      Top             =   3720
      Width           =   1275
   End
   Begin VB.Label Label25 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label25"
      ForeColor       =   &H00000040&
      Height          =   435
      Left            =   3240
      TabIndex        =   24
      ToolTipText     =   "&HC0FFC0"
      Top             =   3120
      Width           =   1275
   End
   Begin VB.Label Label24 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label24"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   3240
      TabIndex        =   23
      ToolTipText     =   "&H4040&"
      Top             =   2520
      Width           =   1275
   End
   Begin VB.Label Label23 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label23"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   3240
      TabIndex        =   22
      ToolTipText     =   "&H8080&"
      Top             =   1920
      Width           =   1275
   End
   Begin VB.Label Label22 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label22"
      ForeColor       =   &H00000040&
      Height          =   435
      Left            =   3240
      TabIndex        =   21
      ToolTipText     =   "&HC0C0&"
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label Label21 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label21"
      ForeColor       =   &H00000040&
      Height          =   435
      Left            =   3240
      TabIndex        =   20
      ToolTipText     =   "&HFFFF&"
      Top             =   720
      Width           =   1275
   End
   Begin VB.Label Label20 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label20"
      ForeColor       =   &H00000040&
      Height          =   435
      Left            =   1740
      TabIndex        =   19
      ToolTipText     =   "&H80FFFF"
      Top             =   6120
      Width           =   1275
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label19"
      ForeColor       =   &H00000040&
      Height          =   435
      Left            =   1740
      TabIndex        =   18
      ToolTipText     =   "&HC0FFFF"
      Top             =   5520
      Width           =   1275
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label18"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   1740
      TabIndex        =   17
      ToolTipText     =   "&H404080"
      Top             =   4920
      Width           =   1275
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label17"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   1740
      TabIndex        =   16
      ToolTipText     =   "&H4080&"
      Top             =   4320
      Width           =   1275
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label16"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   1740
      TabIndex        =   15
      ToolTipText     =   "&H40C0&"
      Top             =   3720
      Width           =   1275
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label15"
      ForeColor       =   &H00000040&
      Height          =   435
      Left            =   1740
      TabIndex        =   14
      ToolTipText     =   "&H80FF&"
      Top             =   3120
      Width           =   1275
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label14"
      ForeColor       =   &H00000040&
      Height          =   435
      Left            =   1740
      TabIndex        =   13
      ToolTipText     =   "&H80C0FF"
      Top             =   2520
      Width           =   1275
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label13"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1740
      TabIndex        =   12
      ToolTipText     =   "&HC0E0FF"
      Top             =   1920
      Width           =   1275
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label12"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   1740
      TabIndex        =   11
      ToolTipText     =   "&H40&"
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label11"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   1740
      TabIndex        =   10
      ToolTipText     =   "&H80&"
      Top             =   720
      Width           =   1275
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label10"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   240
      TabIndex        =   9
      ToolTipText     =   "&HC0&"
      Top             =   6120
      Width           =   1275
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "&HFF"
      Top             =   5520
      Width           =   1275
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label8"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   240
      TabIndex        =   7
      ToolTipText     =   "&H8080FF"
      Top             =   4920
      Width           =   1275
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label7"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "&HC0C0FF"
      Top             =   4320
      Width           =   1275
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "&H0"
      Top             =   3720
      Width           =   1275
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "&H404040"
      Top             =   3120
      Width           =   1275
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "&H808080"
      Top             =   2520
      Width           =   1275
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "&HC0C0C0"
      Top             =   1920
      Width           =   1275
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "&HE0E0E0"
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "&HFFFFFF"
      Top             =   720
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Form and program build written by >>Coyote<<  Date 5/8/01
' Module1 submitted to PSC by "CodeJack"

Private Sub Form_Load()

Label1.BackColor = vbWhite
Label1.Caption = "vbWhite"

Label2.BackColor = vbLightGray
Label2.Caption = "vbLightGray"

Label3.BackColor = vbGray
Label3.Caption = "vbGray"

Label4.BackColor = vbMediumGray
Label4.Caption = "vbMediumGray"

Label5.BackColor = vbDarkGray
Label5.Caption = "vbDarkGray"

Label6.BackColor = vbBlack
Label6.Caption = "vbBlack"

Label7.BackColor = vbPaleRed
Label7.Caption = "VBPaleRed"

Label8.BackColor = vbLightRed
Label8.Caption = "vbLightRed"

Label9.BackColor = vbRed
Label9.Caption = "vbRed"

Label10.BackColor = vbMediumRed
Label10.Caption = "vbMediumRed"

Label11.BackColor = vbDarkRed
Label11.Caption = "vbDarkRed"

Label12.BackColor = vbBlackRed
Label12.Caption = "vbBlackRed"

Label13.BackColor = vbPaleOrange
Label13.Caption = "vbPaleOrange"

Label14.BackColor = vbLightOrange
Label14.Caption = "vbLightOrange"

Label15.BackColor = vbOrange
Label15.Caption = "vbOrange"

Label16.BackColor = vbMediumOrange
Label16.Caption = "vbMediumOrange"

Label17.BackColor = vbDarkOrange
Label17.Caption = "vbDarkOrange"

Label18.BackColor = vbBlackOrange
Label18.Caption = "vbBlackOrange"

Label19.BackColor = vbPaleYellow
Label19.Caption = "vbPaleYellow"

Label20.BackColor = vbLightYellow
Label20.Caption = "vbLightYellow"

Label21.BackColor = vbYellow
Label21.Caption = "vbYellow"

Label22.BackColor = vbMediumYellow
Label22.Caption = "vbMediumYellow"

Label23.BackColor = vbDarkYellow
Label23.Caption = "vbDarkYellow"

Label24.BackColor = vbBlackYellow
Label24.Caption = "vbBlackYellow"

Label25.BackColor = vbPaleGreen
Label25.Caption = "vbPaleGreen"

Label26.BackColor = vbLightGreen
Label26.Caption = "vbLightGreen"

Label27.BackColor = vbGreen
Label27.Caption = "vbGreen"

Label28.BackColor = vbMediumGreen
Label28.Caption = "vbMediumGreen"

Label29.BackColor = vbDarkGreen
Label29.Caption = "vbDarkGreen"

Label30.BackColor = vbBlackGreen
Label30.Caption = "vbBlackGreen"

Label31.BackColor = vbPaleCyan
Label31.Caption = "vbPaleCyan"

Label32.BackColor = vbLightCyan
Label32.Caption = "vbLightCyan"

Label33.BackColor = vbCyan
Label33.Caption = "vbCyan"

Label34.BackColor = vbMediumCyan
Label34.Caption = "vbMediumCyan"

Label35.BackColor = vbDarkCyan
Label35.Caption = "vbDarkCyan"

Label36.BackColor = vbBlackCyan
Label36.Caption = "vbBlackCyan"

Label37.BackColor = vbPaleBlue
Label37.Caption = "vbPaleBlue"

Label38.BackColor = vbLightBlue
Label38.Caption = "vbLightBlue"

Label39.BackColor = vbBlue
Label39.Caption = "vbBlue"

Label40.BackColor = vbMediumBlue
Label40.Caption = "vbMediumBlue"

Label41.BackColor = vbDarkBlue
Label41.Caption = "vbDarkBlue"

Label42.BackColor = vbBlackBlue
Label42.Caption = "vbBlackBlue"

Label43.BackColor = vbPalePurple
Label43.Caption = "vbPalePurple"

Label44.BackColor = vbLightPurple
Label44.Caption = "vbLightPurple"

Label45.BackColor = vbPurple
Label45.Caption = "vbPurple"

Label46.BackColor = vbMagenta
Label46.Caption = "vbMagenta"

Label47.BackColor = vbMediumPurple
Label47.Caption = "vbMediumPurple"

Label48.BackColor = vbdarkpurple
Label48.Caption = "vbDarkpurple"

Label49.BackColor = vbBlackPurple
Label49.Caption = "vbBlackPurple"



End Sub


Private Sub Label1_Click()
    Me.BackColor = &HFFFFFF
End Sub

Private Sub Label10_Click()
    Me.BackColor = &HC0&
End Sub

Private Sub Label11_Click()
    Me.BackColor = &H80&
End Sub

Private Sub Label12_Click()
    Me.BackColor = &H40&
End Sub

Private Sub Label13_Click()
    Me.BackColor = &HC0E0FF
End Sub

Private Sub Label14_Click()
    Me.BackColor = &H80C0FF
End Sub

Private Sub Label15_Click()
    Me.BackColor = &H80FF&
End Sub

Private Sub Label16_Click()
    Me.BackColor = &H40C0&
End Sub

Private Sub Label17_Click()
    Me.BackColor = &H4080&
End Sub

Private Sub Label18_Click()
    Me.BackColor = &H404080
End Sub

Private Sub Label19_Click()
    Me.BackColor = &HC0FFFF
End Sub

Private Sub Label2_Click()
    Me.BackColor = &HE0E0E0
End Sub

Private Sub Label20_Click()
    Me.BackColor = &H80FFFF
End Sub

Private Sub Label21_Click()
    Me.BackColor = &HFFFF&
End Sub

Private Sub Label22_Click()
    Me.BackColor = &HC0C0&
End Sub

Private Sub Label23_Click()
    Me.BackColor = &H8080&
End Sub

Private Sub Label24_Click()
    Me.BackColor = &H4040&
End Sub

Private Sub Label25_Click()
    Me.BackColor = &HC0FFC0
End Sub

Private Sub Label26_Click()
    Me.BackColor = &H80FF80
End Sub

Private Sub Label27_Click()
    Me.BackColor = &HFF00&
End Sub

Private Sub Label28_Click()
    Me.BackColor = &HC000&
End Sub

Private Sub Label29_Click()
    Me.BackColor = &H8000&
End Sub

Private Sub Label3_Click()
    Me.BackColor = &HC0C0C0
End Sub

Private Sub Label30_Click()
    Me.BackColor = &H4000&
End Sub

Private Sub Label31_Click()
    Me.BackColor = &HFFFFC0
End Sub

Private Sub Label32_Click()
    Me.BackColor = &HFFFF80
End Sub

Private Sub Label33_Click()
    Me.BackColor = &HFFFF00
End Sub

Private Sub Label34_Click()
    Me.BackColor = &HC0C000
End Sub

Private Sub Label35_Click()
    Me.BackColor = &H808000
End Sub

Private Sub Label36_Click()
    Me.BackColor = &H404000
End Sub

Private Sub Label37_Click()
    Me.BackColor = &HFFC0C0
End Sub

Private Sub Label38_Click()
    Me.BackColor = &HFF8080
End Sub

Private Sub Label39_Click()
    Me.BackColor = &HFF0000
End Sub

Private Sub Label4_Click()
    Me.BackColor = &H808080
End Sub

Private Sub Label40_Click()
    Me.BackColor = &HC00000
End Sub

Private Sub Label41_Click()
    Me.BackColor = &H800000
End Sub

Private Sub Label42_Click()
    Me.BackColor = &H400000
End Sub

Private Sub Label43_Click()
    Me.BackColor = &HFFC0FF
End Sub

Private Sub Label44_Click()
    Me.BackColor = &HFF80FF
End Sub

Private Sub Label45_Click()
    Me.BackColor = &HFF00FF
End Sub

Private Sub Label46_Click()
    Me.BackColor = &HFF00FF
End Sub

Private Sub Label47_Click()
    Me.BackColor = &HC000C0
End Sub

Private Sub Label48_Click()
    Me.BackColor = &H800080
End Sub

Private Sub Label49_Click()
    Me.BackColor = &H400040
End Sub

Private Sub Label5_Click()
    Me.BackColor = &H404040
End Sub

Private Sub Label6_Click()
    Me.BackColor = &H0
End Sub

Private Sub Label7_Click()
    Me.BackColor = &HC0C0FF
End Sub

Private Sub Label8_Click()
    Me.BackColor = &H8080FF
End Sub

Private Sub Label9_Click()
    Me.BackColor = &HFF
End Sub
