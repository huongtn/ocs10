VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmAddCar 
   Caption         =   "Them Xe"
   ClientHeight    =   3930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4680
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
   ScaleHeight     =   3930
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CboTester 
      Height          =   390
      Left            =   1680
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Frame Frame11 
      Caption         =   "Thoâng tin chung"
      BeginProperty Font 
         Name            =   "VNI-Centur"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.ComboBox CboName 
         Height          =   390
         Left            =   1680
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   420
         Width           =   2535
      End
      Begin VB.TextBox TxtChassisNumber 
         DataField       =   "ChassisNumber"
         DataSource      =   "DatTestingParameter"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox TxtEngineNumber 
         DataField       =   "EngineNumber"
         DataSource      =   "DatTestingParameter"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox TxtProducedNumber 
         DataField       =   "ProducedNumber"
         DataSource      =   "DatTestingParameter"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   1920
         Width           =   2535
      End
      Begin MSForms.CommandButton btnCancel 
         Height          =   495
         Left            =   3240
         TabIndex        =   12
         Top             =   3000
         Width           =   975
         Caption         =   "Huûy"
         Size            =   "1720;873"
         FontName        =   "VNI-Centur"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton btnOK 
         Height          =   495
         Left            =   1680
         TabIndex        =   11
         Top             =   3000
         Width           =   1215
         Caption         =   "Theâm môùi"
         Size            =   "2143;873"
         FontName        =   "VNI-Centur"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Loaïi xe:"
         Height          =   270
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Soá khung:"
         Height          =   270
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Ngöôøi K.T:"
         Height          =   270
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "Soá saûn xuaát:"
         Height          =   270
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   1110
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "Soá maùy:"
         Height          =   270
         Left            =   240
         TabIndex        =   4
         Top             =   2400
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmAddCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connect As New ADODB.Connection
Dim rs As New ADODB.Recordset
Sub dbconnection()
If connect.State = 1 Then connect.Close
If rs.State = 1 Then rs.Close
connect.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source=" & App.Path & "\OCS10_DataBase_97.mdb;Persist Security Info=False"
End Sub
Private Sub addTester()
dbconnection
rs.Open "Select * from TblTesters", connect, adOpenDynamic, adLockOptimistic
CboTester.Clear
Do Until rs.EOF
 CboTester.AddItem rs(1)
rs.MoveNext
Loop
End Sub
