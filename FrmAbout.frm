VERSION 5.00
Begin VB.Form FrmAbout 
   Caption         =   "Thong Tin Phan Mem OCS10"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "VNI-Centur"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSelectTest 
      BackColor       =   &H8000000D&
      Caption         =   "Quay laïi"
      Height          =   420
      Left            =   4650
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
   Begin VB.PictureBox PictureAboutLogo 
      AutoRedraw      =   -1  'True
      Height          =   1380
      Left            =   120
      Picture         =   "FrmAbout.frx":0000
      ScaleHeight     =   88
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   5
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Hoaøn thaønh ngaøy 10-10-2023"
      Height          =   270
      Left            =   1680
      TabIndex        =   4
      Top             =   2790
      Width           =   2610
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Xaây döïng bôûi coâng ty Giaûi Phaùp Coâng Nghieäp"
      Height          =   270
      Left            =   1680
      TabIndex        =   3
      Top             =   2352
      Width           =   4620
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   $"FrmAbout.frx":2F71
      Height          =   1080
      Left            =   1680
      TabIndex        =   2
      Top             =   1005
      Width           =   4575
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Phieân baûn: 2.0"
      Height          =   270
      Left            =   1680
      TabIndex        =   1
      Top             =   564
      Width           =   1320
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      Caption         =   "Teân phaàn meàm: CS10"
      Height          =   270
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1995
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub btnSelectTest_Click()
  Unload Me
End Sub

Private Sub Form_Load()
Move (Screen.Width - Width) * 0.5!, (Screen.Height - Height) * 0.5!
 PictureAboutLogo.ScaleMode = 3
 PictureAboutLogo.AutoRedraw = True
 PictureAboutLogo.PaintPicture PictureAboutLogo.Picture, _
 0, 0, PictureAboutLogo.ScaleWidth, PictureAboutLogo.ScaleHeight, _
 0, 0, _
 PictureAboutLogo.Picture.Width / 26.46, _
 PictureAboutLogo.Picture.Height / 26.46
 PictureAboutLogo.Picture = PictureAboutLogo.Image
End Sub
