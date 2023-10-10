VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmPassword 
   Caption         =   "Doi Mat Khau"
   ClientHeight    =   2130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNewPassword 
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox txtPassword 
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   270
      Width           =   3255
   End
   Begin MSForms.Label Label2 
      Height          =   270
      Left            =   240
      TabIndex        =   5
      Top             =   760
      Width           =   1290
      VariousPropertyBits=   276824091
      Caption         =   "Maät khaåu môùi"
      Size            =   "2275;476"
      FontName        =   "VNI-Centur"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   270
      Left            =   240
      TabIndex        =   4
      Top             =   310
      Width           =   1650
      VariousPropertyBits=   276824091
      Caption         =   "Maät khaåu hieän taïi"
      Size            =   "2910;476"
      FontName        =   "VNI-Centur"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton btnCancel 
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   1410
      Width           =   975
      Caption         =   "Cancel"
      Size            =   "1720;873"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton btnOK 
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   1410
      Width           =   975
      Caption         =   "OK"
      Size            =   "1720;873"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "FrmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
  Unload Me
End Sub

Private Sub btnOK_Click()
  
End Sub

Private Sub Form_Load()
Move (Screen.Width - Width) * 0.5!, (Screen.Height - Height) * 0.5!
End Sub
 
