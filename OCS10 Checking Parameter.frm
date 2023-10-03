VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmCheckingParameter 
   Caption         =   "Checking Parameters"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   ScaleHeight     =   12000
   ScaleMode       =   0  'User
   ScaleWidth      =   19200
   Begin MSDBGrid.DBGrid DBGCheckingUpdate 
      Bindings        =   "OCS10 Checking Parameter.frx":0000
      Height          =   4455
      Left            =   240
      OleObjectBlob   =   "OCS10 Checking Parameter.frx":0023
      TabIndex        =   0
      Top             =   720
      Width           =   18495
   End
   Begin VB.Data DatCheckingParameter 
      Caption         =   "Database Checking Parameter"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   16440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11520
      Top             =   120
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
            Picture         =   "OCS10 Checking Parameter.frx":0A00
            Key             =   "KeyNew"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OCS10 Checking Parameter.frx":0B12
            Key             =   "KeyEdit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OCS10 Checking Parameter.frx":0C24
            Key             =   "KeyAbort"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OCS10 Checking Parameter.frx":0D36
            Key             =   "KeySave"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OCS10 Checking Parameter.frx":0E48
            Key             =   "KeyDelete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OCS10 Checking Parameter.frx":0F5A
            Key             =   "KeyUddate"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OCS10 Checking Parameter.frx":106C
            Key             =   "KeyReport"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OCS10 Checking Parameter.frx":117E
            Key             =   "KeyPara"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OCS10 Checking Parameter.frx":1290
            Key             =   "KeyExit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TbrMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   635
      ButtonWidth     =   1984
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "KeyNew"
            Object.ToolTipText     =   "Add new Car's testing result"
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
            Object.ToolTipText     =   "Save changed Parameters"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "KeyDelete"
            Object.ToolTipText     =   "Delete one Car's Testing Result"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "   Refresh"
            Key             =   "KeyRefresh"
            Object.ToolTipText     =   "Update all new parameters of Database"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "KeyExit"
            Object.ToolTipText     =   "Return Main Screen"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmCheckingParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdEnd_Click()
Unload Me
End Sub

Private Sub Form_Load()

DatCheckingParameter.DatabaseName = FrmMain.DataBaseFolder & "\OCS10_DataBase_97.mdb"
DatCheckingParameter.RecordSource = "select * from TblCheckingParameter"

TbrMain.Buttons(7).Enabled = False
TbrMain.Buttons(9).Enabled = False
TbrMain.Buttons(13).Enabled = False
End Sub

Private Sub TbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key

Case "KeyNew"
MnuAddNew_Click

Case "KeyEdit"
MnuEditParameter_Click

Case "KeyAbort"
MnuAbort_Click

Case "KeySave"
MnuSave_Click

Case "KeyDelete"
MnuDeleteParameter_Click

Case "KeyRefresh"
MnuRefreshParameter_Click


Case "KeyExit"
MnuExit_Click

End Select
End Sub
Private Sub MnuAddNew_Click()
DatCheckingParameter.Recordset.AddNew
'DatCheckingParameter.Recordset.Update
DBGCheckingUpdate.AllowAddNew = True
DBGCheckingUpdate.AllowUpdate = True
TbrMain.Buttons(5).Enabled = False
TbrMain.Buttons(7).Enabled = True
TbrMain.Buttons(9).Enabled = True
TbrMain.Buttons(11).Enabled = False
End Sub
Private Sub MnuEditParameter_Click()
DBGCheckingUpdate.AllowUpdate = True
DatCheckingParameter.Recordset.Edit
TbrMain.Buttons(3).Enabled = False
TbrMain.Buttons(5).Enabled = False
TbrMain.Buttons(7).Enabled = True
TbrMain.Buttons(9).Enabled = True
TbrMain.Buttons(11).Enabled = False

End Sub
Private Sub MnuAbort_Click()
TbrMain.Buttons(3).Enabled = True
TbrMain.Buttons(5).Enabled = True
On Error GoTo ErrHandling
DatCheckingParameter.Recordset.CancelUpdate
DBGCheckingUpdate.AllowUpdate = False

TbrMain.Buttons(7).Enabled = False
TbrMain.Buttons(9).Enabled = False
TbrMain.Buttons(11).Enabled = True
EndIt:
Exit Sub ' or Exit Function

ErrHandling:
Select Case Err.Number
Case 3020
MsgBox "Need Edit or Add new Task Before !"
Case Else
MsgBox Err.Description
End Select
Resume EndIt
End Sub
Private Sub MnuSave_Click()

TbrMain.Buttons(3).Enabled = True
TbrMain.Buttons(5).Enabled = True
'-----------------------------------------------------
On Error GoTo ErrHandling

'Bat dau doan code chinh o day---------------------------------------
DatCheckingParameter.Recordset.Update
DBGCheckingUpdate.AllowUpdate = False
DBGCheckingUpdate.AllowAddNew = False

TbrMain.Buttons(7).Enabled = False
TbrMain.Buttons(9).Enabled = False
TbrMain.Buttons(11).Enabled = True
'Ket thuc Doan code chinh o day---------------------------------------
EndIt:
Exit Sub ' or Exit Function

ErrHandling:
Select Case Err.Number
Case 3022
MsgBox "The Name already exists !"
Case Else
MsgBox Err.Description
End Select
Resume EndIt
'-----------------------------------------------------

End Sub
Private Sub MnuDeleteParameter_Click()
 Dim RecordCurrent As Integer
 On Error GoTo Delete_Error

    If MsgBox("Are you sure you want to delete this record?", _
                vbQuestion + vbYesNo + vbDefaultButton2, _
            "Confirm") = vbNo Then
            Exit Sub
        End If

    'delete the current record
    RecordCurrent = DatCheckingParameter.Recordset.AbsolutePosition
    DatCheckingParameter.Recordset.Delete
    DatCheckingParameter.Refresh
    
    'move to a valid record
    'MoveNextRecord
    DatCheckingParameter.Recordset.Move (RecordCurrent)
    Exit Sub
Delete_Error:
    MsgBox "This record cannot be deleted. Error code = " _
           & CStr(Err.Number) & vbCrLf & Err.Description, _
           vbCritical, "Cannot Delete"
End Sub
Private Sub MnuRefreshParameter_Click()
DatCheckingParameter.Refresh
End Sub

Private Sub MnuExit_Click()
Unload Me
End Sub


