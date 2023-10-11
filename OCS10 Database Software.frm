VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmMain 
   Caption         =   "DBS10 - Database Software  -  Designed by INDUSTRY SOLUTION Co.  -   www.thietbicongnghiep.vn"
   ClientHeight    =   10680
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   19875
   FillColor       =   &H00808080&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "OCS10 Database Software.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10250.87
   ScaleMode       =   0  'User
   ScaleWidth      =   22755.52
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame11 
      Caption         =   "Thoâng tin xe môùi"
      BeginProperty Font 
         Name            =   "VNI-Centur"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   8295
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   8535
      Begin VB.ComboBox CboTester 
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   23
         Text            =   "Combo1"
         Top             =   2535
         Width           =   8055
      End
      Begin VB.TextBox TxtProducedNumber 
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   6720
         Width           =   8055
      End
      Begin VB.TextBox TxtEngineNumber 
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   5325
         Width           =   8055
      End
      Begin VB.TextBox TxtChassisNumber 
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   3930
         Width           =   8055
      End
      Begin VB.ComboBox CboName 
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   1140
         Width           =   8055
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "Soá maùy:"
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "Soá saûn xuaát:"
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   6120
         Width           =   2070
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Ngöôøi K.T:"
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   2040
         Width           =   1860
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Soá khung:"
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   3360
         Width           =   1680
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Loaïi xe:"
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   1335
      End
      Begin MSForms.CommandButton btnAdd 
         Height          =   735
         Left            =   6000
         TabIndex        =   17
         Top             =   7440
         Width           =   2295
         ForeColor       =   -2147483634
         BackColor       =   -2147483635
         Caption         =   "Theâm môùi"
         Size            =   "4048;1296"
         FontName        =   "VNI-Centur"
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   9480
      Top             =   12480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox TxtSelectedEngineNumber 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   13200
      TabIndex        =   10
      Top             =   1320
      Width           =   6375
   End
   Begin VB.TextBox TxtSelectedChassisNumber 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   6720
      TabIndex        =   9
      Top             =   1320
      Width           =   6375
   End
   Begin VB.TextBox TxtSelectedName 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   6375
   End
   Begin VB.PictureBox CommonDialog1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   13920
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   12360
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox txtCurrentID 
      DataField       =   "STT"
      DataSource      =   "DatTestingParameter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   6
      Text            =   "CurrentID"
      Top             =   12480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtSqlReport 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10680
      TabIndex        =   5
      Text            =   "SqlToReport"
      Top             =   12600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Tmr1 
      Left            =   8520
      Top             =   12360
   End
   Begin VB.Data DatCheckingParameter 
      Caption         =   "Database Checking Parameters"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   12480
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Data DatTestingParameter 
      Caption         =   "Database Testing Parameter"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   12360
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Frame Frame10 
      Caption         =   "Danh saùch xe"
      BeginProperty Font 
         Name            =   "VNI-Centur"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   8295
      Left            =   8880
      TabIndex        =   0
      Top             =   2280
      Width           =   10860
      Begin MSDBGrid.DBGrid DBGTestingUpdate 
         Bindings        =   "OCS10 Database Software.frx":0442
         Height          =   6480
         Left            =   240
         OleObjectBlob   =   "OCS10 Database Software.frx":0464
         TabIndex        =   4
         Top             =   1560
         Width           =   10260
      End
      Begin MSForms.CommandButton CommandButton1 
         Height          =   735
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   2295
         ForeColor       =   -2147483634
         BackColor       =   33023
         Caption         =   "Choïn xe test"
         Size            =   "4048;1296"
         FontName        =   "VNI-Centur"
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Xe ñang test"
      BeginProperty Font 
         Name            =   "VNI-Centur"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2295
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   19575
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         Caption         =   "Soá maùy"
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   13080
         TabIndex        =   14
         Top             =   600
         Width           =   1650
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "Soá khung"
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   6600
         TabIndex        =   13
         Top             =   600
         Width           =   2100
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         Caption         =   "Loaïi xe"
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1650
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DataBaseFolder As String
Dim connect As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim ColorGreen As String
Dim ColorOrange As String
Private SelectedTab As Integer
Public Sql As String
Public ArrayString As Integer
Dim RowCount As Integer

'-----------------------------------------------------------------------------------------
' initialize Checking Valve
Dim SpeedMin As Integer
Dim SpeedMax As Integer
Dim BrakeFrontSumMin As Integer
Dim BrakeRearSumMin As Integer
Dim BrakeStopSumMin As Integer
Dim BrakeFrontDifMax As Integer
Dim BrakeRearDifMax As Integer
Dim BrakeStopDifMax As Integer
Dim NoiseMax As Integer
Dim BuzzerMin As Integer
Dim BuzzerMax As Integer
Dim AlignMin As Integer
Dim AlignMax As Integer
Dim HCMax As Integer
Dim COMax As Integer
Dim CO2Max As Integer
Dim O2Max As Integer
Dim NOMax As Integer
Dim HSUMax As Integer
Dim HeSoDieselMax As Integer

Dim HLHighIntMin As Integer
Dim HLHighLRMin As Integer
Dim HLHighLRMax As Integer
Dim HLHighUDMin As Integer
Dim HLHighUDMax As Integer

Dim HLLowIntMin As Integer
Dim HLLowLRMin As Integer
Dim HLLowLRMax As Integer
Dim HLLowUDMin As Integer
Dim HLLowUDMax As Integer


'----------------------Khoi tao gia tri checking Parameter -----------

Private Sub InitializeCheckingParameter()
 End Sub

Private Sub SubErrHandling()
Select Case Err.Number
Case 3020
MsgBox " Need Edit or Add New Task Before "
Case 3021
MsgBox "No current Record"
Case Else
MsgBox Err.Description
End Select
End Sub
Private Sub SubEnableAll()
Dim EnableBit As Boolean
EnableBit = True


End Sub
Private Sub SubDisableAll()
Dim EnableBit As Boolean
EnableBit = False
End Sub

  

Private Sub btnAdd_Click()
 Dim Name As String
    Dim Tester As String
    Dim ChassisNumber As String
    Dim ProducedNumber As String
    Dim EngineNumber As String
    Name = CboName.Text
    Tester = CboTester.Text
    ChassisNumber = TxtChassisNumber.Text
    ProducedNumber = TxtProducedNumber.Text
    EngineNumber = TxtEngineNumber.Text
    With DatTestingParameter.Recordset
        .AddNew
        !Name = Name
        !Tester = Tester
        !ChassisNumber = ChassisNumber
        !ProducedNumber = ProducedNumber
        !EngineNumber = EngineNumber
        !Date = Date
        !SelectedDateTime = Now()
        .Update
    End With
    DatTestingParameter.Refresh
    TxtChassisNumber.Text = ""
    TxtProducedNumber.Text = ""
    TxtEngineNumber.Text = ""
    LoadSelectVehicle
    TxtChassisNumber.SetFocus
End Sub


Sub dbconnection()
If connect.State = 1 Then connect.Close
If rs.State = 1 Then rs.Close
connect.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source=" & FrmMain.DataBaseFolder & "\OCS10_DataBase_97.mdb;Persist Security Info=False"
End Sub
Private Sub addTester()
dbconnection
Dim FirstTester As String

rs.Open "Select * from TblTesters", connect, adOpenDynamic, adLockOptimistic
CboTester.Clear
FirstTester = rs(1)
Do Until rs.EOF
 CboTester.AddItem rs(1)
rs.MoveNext
CboTester.Text = FirstTester
Loop
End Sub

Private Sub addName()
dbconnection
Dim First As String

rs.Open "Select * from TblCheckingParameter", connect, adOpenDynamic, adLockOptimistic
CboName.Clear
First = rs(0)
Do Until rs.EOF
 CboName.AddItem rs(0)
rs.MoveNext
CboName.Text = First
Loop
End Sub


Private Sub CommandButton1_Click()
Dim connect As New ADODB.Connection
Dim RST As New ADODB.Recordset

If connect.State = 1 Then connect.Close
If RST.State = 1 Then RST.Close
connect.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source=" & DataBaseFolder & "\OCS10_DataBase_97.mdb;Persist Security Info=False"

Dim sSQL As String
sSQL = "Select * From TblTestingParameter Where STT = " & Val(txtCurrentID.Text) & ""
RST.Open sSQL, connect, adOpenDynamic, adLockOptimistic
If Not RST.EOF Then
RST("SelectedDateTime") = Now()
RST.Update
TxtSelectedName.Text = RST("Name")
If RST("ChassisNumber") <> "" Then
    TxtSelectedChassisNumber.Text = RST("ChassisNumber")
Else
    TxtSelectedChassisNumber.Text = ""
End If
  If RST("EngineNumber") <> "" Then
    TxtSelectedEngineNumber.Text = RST("EngineNumber")
Else
    TxtSelectedEngineNumber.Text = ""
End If
  
MsgBox "Ban da chon xe test(" & RST("ChassisNumber") & ")"
Else
MsgBox "Record Not Found..."
End If
RST.Close
End Sub

Private Sub DatTestingParameter_Reposition()
InitializeCheckingParameter
End Sub


 

Private Sub Form_Load()
DataBaseFolder = "\\Master\OCS10"
'DataBaseFolder = App.Path
txtSqlReport.Text = "SELECT * FROM TblTestingParameter"
DatTestingParameter.DatabaseName = DataBaseFolder & "\OCS10_DataBase_97.mdb"
DatTestingParameter.RecordSource = "select STT, Date, ChassisNumber,EngineNumber, Name,SelectedDateTime,Tester, ProducedNumber from TblTestingParameter order by STT desc"

DatCheckingParameter.DatabaseName = DataBaseFolder & "\OCS10_DataBase_97.mdb"
DatCheckingParameter.RecordSource = "select * from TblCheckingParameter"
LoadSelectVehicle
  
addTester
addName
End Sub

 

 

   
Private Sub MnuHelpAboutOCS10DBS_Click()
FrmContactUs.Show
End Sub

Private Sub MnuHelpGuide_Click()
FrmAbout.Show
End Sub

Private Sub MnuImportVehicles_Click()
CommonDialog2.Filter = "Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*"
CommonDialog2.DefaultExt = "txt"
CommonDialog2.DialogTitle = "Select File"
CommonDialog2.ShowOpen

Dim ExcelObj As Object
Dim ExcelBook As Object
Dim ExcelSheet As Object
Dim i As Integer
If CommonDialog2.FileName <> "" Then
    Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelSheet = CreateObject("Excel.Sheet")
    
    ExcelObj.WorkBooks.Open CommonDialog2.FileName
    
    Set ExcelBook = ExcelObj.WorkBooks(1)
    Set ExcelSheet = ExcelBook.WorkSheets(1)
     
    Dim curentTester As String
    Dim curentName As String
    Dim curentChassisNumber As String
    Dim curentProducedNumber As String
    Dim curentEngineNumber As String
    curentTester = CboTester.Text
    With ExcelSheet
    i = 3
    Do Until .cells(i, 2) & "" = ""
    curentName = .cells(i, 2)
    curentProducedNumber = .cells(i, 3)
    curentChassisNumber = .cells(i, 4)
    curentEngineNumber = .cells(i, 5)
       With DatTestingParameter.Recordset
        .AddNew
        !Name = curentName
        !ChassisNumber = curentChassisNumber
        !ProducedNumber = curentProducedNumber
        !EngineNumber = curentEngineNumber
        !Tester = curentTester
        !Date = Date
        .Update
        End With
        i = i + 1
    Loop
    
    End With
    
    ExcelObj.WorkBooks.Close
    
    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing
    DatTestingParameter.Refresh
    MsgBox "Thêm thành công " & CStr(i - 3) & " xe"
End If
End Sub
 
Private Sub MnuReportTotal_Click()
FrmReportSeperate.Show
End Sub

Private Sub MnuSaveAsDataBase_Click()
Unload Me
FrmBackupDB.Show
End Sub
 

Private Sub LoadSelectVehicle()
    Dim connect As New ADODB.Connection
    Dim RST As New ADODB.Recordset
    
    If connect.State = 1 Then connect.Close
    If RST.State = 1 Then RST.Close
    connect.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source=" & DataBaseFolder & "\OCS10_DataBase_97.mdb;Persist Security Info=False"
    
    Dim sSQL As String
    sSQL = "Select * From TblTestingParameter ORDER BY SelectedDateTime DESC"
    RST.Open sSQL, connect, adOpenDynamic, adLockOptimistic
    If Not RST.EOF Then
    TxtSelectedName.Text = RST("Name")
    If RST("ChassisNumber") <> "" Then
        TxtSelectedChassisNumber.Text = RST("ChassisNumber")
    Else
        TxtSelectedChassisNumber.Text = ""
    End If
    
   
    
      If RST("EngineNumber") <> "" Then
        TxtSelectedEngineNumber.Text = RST("EngineNumber")
    Else
        TxtSelectedEngineNumber.Text = ""
    End If
    Else
    End If
    RST.Close
End Sub
 
Private Sub TxtChassisNumber_Change()
If Len(TxtChassisNumber.Text) >= 17 Then
TxtEngineNumber.SetFocus
End If
End Sub

Private Sub TxtEngineNumber_Change()
If Len(TxtEngineNumber.Text) >= 12 Then
TxtProducedNumber.SetFocus
End If
End Sub

