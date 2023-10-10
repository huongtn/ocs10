VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmLogin 
   Caption         =   "Login"
   ClientHeight    =   1755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPassword 
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   507
      Width           =   3255
   End
   Begin MSForms.CommandButton btnOK 
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   1080
      Width           =   975
      Caption         =   "OK"
      Size            =   "1720;873"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton btnCancel 
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   1080
      Width           =   975
      Caption         =   "Cancel"
      Size            =   "1720;873"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   555
      Width           =   1005
      VariousPropertyBits=   276824091
      Caption         =   "Maät khaåu"
      Size            =   "1773;476"
      FontName        =   "VNI-Centur"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Move (Screen.Width - Width) * 0.5!, (Screen.Height - Height) * 0.5!
End Sub
