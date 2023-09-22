VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmTester 
   Caption         =   "Form Manage Testers"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Thong tin Tester"
      Height          =   4695
      Left            =   6000
      TabIndex        =   2
      Top             =   480
      Width           =   4215
      Begin VB.TextBox txtTesterID 
         DataField       =   "ID"
         DataSource      =   "DataTester"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtTesterName 
         DataField       =   "Name"
         DataSource      =   "DataTester"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Tester ID :"
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Tester Name :"
         Height          =   495
         Left            =   480
         TabIndex        =   5
         Top             =   1560
         Width           =   2535
      End
   End
   Begin VB.Data DataTester 
      Caption         =   "Data Tester"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   7800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5640
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7080
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTester1.frx":0000
            Key             =   "KeyNew"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTester1.frx":0112
            Key             =   "KeyEdit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTester1.frx":0224
            Key             =   "KeyAbort"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTester1.frx":0336
            Key             =   "KeySave"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTester1.frx":0448
            Key             =   "KeyDelete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTester1.frx":055A
            Key             =   "KeyUddate"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTester1.frx":066C
            Key             =   "KeyReport"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTester1.frx":077E
            Key             =   "KeyPara"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTester1.frx":0890
            Key             =   "KeyExit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TbrMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10710
      _ExtentX        =   18891
      _ExtentY        =   635
      ButtonWidth     =   1984
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "KeyNew"
            Object.ToolTipText     =   "Add new Tester"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Key             =   "KeyEdit"
            Object.ToolTipText     =   "Edit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Abort"
            Key             =   "KeyAbort"
            Object.ToolTipText     =   "Abort any change"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "KeySave"
            Object.ToolTipText     =   "Save changed Tester"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "KeyDelete"
            Object.ToolTipText     =   "Delete one Tester"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "   Refresh"
            Key             =   "KeyRefresh"
            Object.ToolTipText     =   "Update all new Tester"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "KeyExit"
            Object.ToolTipText     =   "Return Main Screen"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
      EndProperty
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FrmTester1.frx":09A2
      Height          =   4575
      Left            =   360
      OleObjectBlob   =   "FrmTester1.frx":09BB
      TabIndex        =   1
      Top             =   600
      Width           =   5295
   End
End
Attribute VB_Name = "FrmTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FileName   : FrmTester
'Author     : Thinhgacon
'Date       : 10/03/2011

Private Sub EnableAll()
txtTesterID.Enabled = True
txtTesterName.Enabled = True
End Sub

Private Sub DisableAll()
txtTesterID.Enabled = False
txtTesterName.Enabled = False
End Sub


Private Sub Form_Load()

DataTester.DatabaseName = App.Path & "\OCS10_DataBase_97.mdb"
DataTester.RecordSource = "select * from TblTesters  order by ID DESC"

DisableAll
TbrMain.Buttons(7).Enabled = False
TbrMain.Buttons(9).Enabled = False
End Sub

Private Sub TbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key

Case "KeyNew"
MnuAddNew_Click

Case "KeyEdit"
MnuEditResult_Click

Case "KeyAbort"
MnuAbort_Click

Case "KeySave"
MnuSave_Click

Case "KeyDelete"
MnuDelete_Click

Case "KeyExit"
MnuExit_Click

End Select

End Sub


Private Sub MnuAddNew_Click()
TbrMain.Buttons(3).Enabled = False
TbrMain.Buttons(5).Enabled = False
TbrMain.Buttons(11).Enabled = False
DBGrid1.Enabled = False
txtTesterName.Enabled = True
'txtTesterID.Text = ""
txtTesterID.BackColor = &H80000000
'txtTesterName.Text = ""
'txtTesterName.SetFocus

TbrMain.Buttons(7).Enabled = True
TbrMain.Buttons(9).Enabled = True

With DataTester.Recordset
        .AddNew
        !Name = "Tester"
        .Update
    End With
    DataTester.Refresh
DataTester.Recordset.Edit

End Sub

Private Sub MnuEditResult_Click()
DataTester.Recordset.Edit
TbrMain.Buttons(3).Enabled = False
TbrMain.Buttons(5).Enabled = False
TbrMain.Buttons(11).Enabled = False
txtTesterName.Enabled = True
txtTesterName.SetFocus
txtTesterName.SelStart = Len(txtTesterName.Text)

TbrMain.Buttons(7).Enabled = True
TbrMain.Buttons(9).Enabled = True
DBGrid1.Enabled = False

End Sub

Private Sub MnuAbort_Click()
'DataTester.Refresh
DataTester.Recordset.CancelUpdate
TbrMain.Buttons(3).Enabled = True
TbrMain.Buttons(5).Enabled = True
TbrMain.Buttons(7).Enabled = False
TbrMain.Buttons(9).Enabled = False
TbrMain.Buttons(11).Enabled = True
txtTesterName.Enabled = False
DBGrid1.Enabled = True

End Sub
Private Sub MnuSave_Click()
DataTester.Recordset.Edit
DataTester.Recordset.Update

TbrMain.Buttons(3).Enabled = True
TbrMain.Buttons(5).Enabled = True
TbrMain.Buttons(7).Enabled = False
TbrMain.Buttons(9).Enabled = False
TbrMain.Buttons(11).Enabled = True
txtTesterName.Enabled = False
DBGrid1.Enabled = True
End Sub
Private Sub MnuDelete_Click()

If MsgBox("Are you sure you want to delete this record?", _
                vbQuestion + vbYesNo + vbDefaultButton2, _
            "Confirm") = vbNo Then
            Exit Sub
        End If

With DataTester.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
    End With
End Sub
Private Sub MnuExit_Click()
Unload Me
End Sub


