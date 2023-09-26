VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmLogin 
   Caption         =   "Login"
   ClientHeight    =   2325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPassword 
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   507
      Width           =   3255
   End
   Begin MSForms.CommandButton btnOK 
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   1320
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
      Left            =   2640
      TabIndex        =   2
      Top             =   1320
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
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   562
      Width           =   795
      VariousPropertyBits=   276824091
      Caption         =   "Mat khau"
      Size            =   "1402;423"
      FontName        =   "Tahoma"
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

Private Sub btnCancel_Click()
  Unload Me
End Sub

Private Sub btnOK_Click()
    'Dim connect As New ADODB.Connection
    'Dim rs As New ADODB.Recordset
    'If connect.State = 1 Then connect.Close
    'If rs.State = 1 Then rs.Close
    'connect.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source=" & App.Path & "\OCS10_DataBase_97.mdb;Persist Security Info=False"
    'rs.Open "Select * from Password", connect, adOpenDynamic, adLockOptimistic
    'Do Until rs.EOF
    'MsgBox rs!Password
    'rs.MoveNext
    'Loop
    If txtPassword.Text = "123456" Then
       FrmMain.BeginUpdate
       Unload Me
       Else
       
       MsgBox "Sai mat khau"
       
    End If
    
    
    
End Sub

