VERSION 5.00
Begin VB.Form FrmContactUs 
   Caption         =   "Industry Solution .JSC"
   ClientHeight    =   3585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6570
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
   ScaleHeight     =   3585
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PictureContactUs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      Picture         =   "FrmContactUs.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton btnSelectTest 
      BackColor       =   &H8000000D&
      Caption         =   "Thoaùt"
      Height          =   420
      Left            =   4920
      TabIndex        =   0
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Giaûi Phaùp Laø Tieân Phong"
      Height          =   270
      Left            =   3120
      TabIndex        =   8
      Top             =   2640
      Width           =   2340
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Tel: +84 43665 8985"
      Height          =   270
      Left            =   3120
      TabIndex        =   7
      Top             =   2130
      Width           =   1815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Email: info@thietbicongnghiep.vn"
      Height          =   270
      Left            =   3120
      TabIndex        =   6
      Top             =   1620
      Width           =   3090
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "http://thietbicongnghiep.vn"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   3960
      TabIndex        =   5
      Top             =   1110
      Width           =   2505
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Website: "
      Height          =   270
      Left            =   3120
      TabIndex        =   4
      Top             =   1110
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "INDUSTRY SOLUTION .JSC"
      Height          =   270
      Left            =   3120
      TabIndex        =   3
      Top             =   600
      Width           =   2715
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Coâng ty Giaûi Phaùp Coâng Nghieäp"
      BeginProperty Font 
         Name            =   "VNI-Centur"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3120
      TabIndex        =   2
      Top             =   90
      Width           =   3285
   End
End
Attribute VB_Name = "FrmContactUs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub btnSelectTest_Click()
  Unload Me
End Sub

Private Sub Form_Load()
Move (Screen.Width - Width) * 0.5!, (Screen.Height - Height) * 0.5
End Sub
