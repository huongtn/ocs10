VERSION 5.00
Begin VB.Form FrmUpdateCheckingParameter 
   Caption         =   "Form1"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdEnd 
      Caption         =   "End"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton CmdNewName 
      Caption         =   "New Name"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton CmdNewChecking 
      Caption         =   "New Checking"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   4800
      Width           =   1455
   End
End
Attribute VB_Name = "FrmUpdateCheckingParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdEnd_Click()
Unload Me
End Sub
