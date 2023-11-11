VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmBackupDB 
   Caption         =   "BackUp Database"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   8010
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFileName 
      Height          =   405
      Left            =   4440
      TabIndex        =   5
      Top             =   1200
      Width           =   3255
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox txtLocation 
      Enabled         =   0   'False
      Height          =   405
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
   End
   Begin MSForms.CommandButton btnCancel 
      Height          =   495
      Left            =   5760
      TabIndex        =   7
      Top             =   1920
      Width           =   975
      Caption         =   "Cancel"
      Size            =   "1720;873"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton btnBackup 
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   1920
      Width           =   975
      Caption         =   "BackUp"
      Size            =   "1720;873"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   600
      Width           =   2775
      Caption         =   "Chon ten co so du lieu:"
      Size            =   "4895;661"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2535
      Caption         =   "Chon vi tri luu co so du lieu:"
      Size            =   "4471;661"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FrmBackupDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBackup_Click()
If txtLocation.Text = "" Then
    MsgBox ("Chon duong dan toi thu muc luu tru")
Else
    If txtFileName.Text = "" Then
    MsgBox ("Chon ten file luu tru")
    Else
    On Error GoTo ErrHandling
    FileSystem.FileCopy "C:\Program Files\OCS10" & FrmMain.DataBaseName, txtLocation.Text & "\" & txtFileName.Text & ".bak"
    MsgBox "Backup Database Successful!"
    Unload Me
    FrmMain.Show
ErrHandling:
        MsgBox "Error. Try Again please !"
    End If
End If
End Sub

Private Sub btnCancel_Click()
Unload Me
FrmMain.Show
End Sub

Private Sub Dir1_Change()
txtLocation.Text = Dir1
End Sub

Private Sub Drive1_Change()
On Error Resume Next
    Dir1.Path = Drive1
End Sub
