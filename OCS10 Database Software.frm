VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmMain 
   BorderStyle     =   0  'None
   Caption         =   "DBS10 - Database Software  -  Designed by INDUSTRY SOLUTION Co.  -   www.thietbicongnghiep.vn"
   ClientHeight    =   12360
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   17775
   FillColor       =   &H00808080&
   Icon            =   "OCS10 Database Software.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11863.37
   ScaleMode       =   0  'User
   ScaleWidth      =   20351.16
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame11 
      Caption         =   "Tho�ng tin xe m��i"
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
      Height          =   3615
      Left            =   12720
      TabIndex        =   46
      Top             =   1920
      Width           =   4935
      Begin VB.ComboBox CboTester 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   57
         Text            =   "Combo1"
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox TxtProducedNumber 
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
         TabIndex        =   50
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox TxtEngineNumber 
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
         TabIndex        =   49
         Top             =   2400
         Width           =   3015
      End
      Begin VB.TextBox TxtChassisNumber 
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
         TabIndex        =   48
         Top             =   960
         Width           =   3015
      End
      Begin VB.ComboBox CboName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   47
         Text            =   "Combo1"
         Top             =   420
         Width           =   3015
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "So� ma�y:"
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   56
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "So� sa�n xua�t:"
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   55
         Top             =   1920
         Width           =   1110
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Ng���i K.T:"
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   54
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "So� khung:"
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   53
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Loa�i xe:"
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   52
         Top             =   480
         Width           =   735
      End
      Begin MSForms.CommandButton btnAdd 
         Height          =   495
         Left            =   3240
         TabIndex        =   51
         Top             =   2880
         Width           =   1455
         ForeColor       =   -2147483634
         BackColor       =   -2147483635
         Caption         =   "The�m m��i"
         Size            =   "2566;873"
         FontName        =   "VNI-Centur"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.Frame freSearch 
      BorderStyle     =   0  'None
      Height          =   4665
      Index           =   4
      Left            =   7560
      TabIndex        =   19
      Top             =   3480
      Width           =   4359
      Begin VB.ListBox LstAll 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3900
         ItemData        =   "OCS10 Database Software.frx":0442
         Left            =   120
         List            =   "OCS10 Database Software.frx":0444
         TabIndex        =   21
         Top             =   600
         Width           =   4095
      End
      Begin VB.CommandButton CmdShowAll 
         Caption         =   "Click �e� xem"
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   4095
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   14280
      TabIndex        =   37
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox TxtSelectedProducedNumber 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   10590
      TabIndex        =   36
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox TxtSelectedChassisNumber 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   6900
      TabIndex        =   35
      Top             =   930
      Width           =   3015
   End
   Begin VB.TextBox TxtSelectedName 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   3210
      TabIndex        =   34
      Top             =   930
      Width           =   3015
   End
   Begin VB.CommandButton btnSelectTest 
      BackColor       =   &H8000000D&
      Caption         =   "CHO�N XE TEST"
      BeginProperty Font 
         Name            =   "VNI-Centur"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   480
      TabIndex        =   33
      Top             =   930
      Width           =   2055
   End
   Begin VB.PictureBox CommonDialog1 
      Height          =   480
      Left            =   13920
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   32
      Top             =   12360
      Visible         =   0   'False
      Width           =   1200
   End
   Begin MSACAL.Calendar cldToDate 
      Height          =   3015
      Left            =   15000
      TabIndex        =   26
      Top             =   7320
      Visible         =   0   'False
      Width           =   3975
      _Version        =   524288
      _ExtentX        =   7011
      _ExtentY        =   5318
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2012
      Month           =   3
      Day             =   20
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSACAL.Calendar cldFromDate 
      Height          =   3255
      Left            =   11760
      TabIndex        =   25
      Top             =   7200
      Visible         =   0   'False
      Width           =   3975
      _Version        =   524288
      _ExtentX        =   7011
      _ExtentY        =   5741
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2012
      Month           =   3
      Day             =   20
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame freSearch 
      BorderStyle     =   0  'None
      Height          =   4785
      Index           =   0
      Left            =   13080
      TabIndex        =   11
      Top             =   6720
      Width           =   4260
      Begin VB.TextBox TxtNameSearch 
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
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   3375
      End
      Begin VB.ListBox LstNameSearch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3900
         ItemData        =   "OCS10 Database Software.frx":0446
         Left            =   120
         List            =   "OCS10 Database Software.frx":0448
         TabIndex        =   12
         Top             =   600
         Width           =   4095
      End
      Begin MSForms.CommandButton CmdNameSearch 
         Height          =   375
         Left            =   3600
         TabIndex        =   27
         Top             =   120
         Width           =   615
         VariousPropertyBits=   25
         Caption         =   "Tim"
         Size            =   "1085;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.TextBox txtCurrentID 
      DataField       =   "STT"
      DataSource      =   "DatTestingParameter"
      Height          =   375
      Left            =   12120
      TabIndex        =   24
      Text            =   "CurrentID"
      Top             =   12480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtSqlReport 
      Height          =   405
      Left            =   10680
      TabIndex        =   23
      Text            =   "SqlToReport"
      Top             =   12600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame freSearch 
      BorderStyle     =   0  'None
      Height          =   5025
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Top             =   3840
      Width           =   4359
      Begin VB.ListBox LstDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3900
         ItemData        =   "OCS10 Database Software.frx":044A
         Left            =   120
         List            =   "OCS10 Database Software.frx":044C
         TabIndex        =   22
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox TxtDateFrom 
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
         Left            =   720
         TabIndex        =   18
         Text            =   "1/1/2012"
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton CmdDateSearchTo 
         Caption         =   "To"
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
         Left            =   2280
         TabIndex        =   9
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton CmdDateSearchFrom 
         Caption         =   "Fr"
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
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox TxtDateTo 
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
         Left            =   2880
         TabIndex        =   7
         Text            =   "12/30/2012"
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame freSearch 
      BorderStyle     =   0  'None
      Height          =   2500
      Index           =   1
      Left            =   7080
      TabIndex        =   5
      Top             =   7920
      Width           =   4359
      Begin VB.TextBox TxtChassisSearch 
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
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   3375
      End
      Begin VB.ListBox LstChassisSearch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3900
         ItemData        =   "OCS10 Database Software.frx":044E
         Left            =   120
         List            =   "OCS10 Database Software.frx":0450
         TabIndex        =   14
         Top             =   600
         Width           =   4095
      End
      Begin MSForms.CommandButton CmdChassisSearch 
         Height          =   375
         Left            =   3600
         TabIndex        =   29
         Top             =   120
         Width           =   615
         VariousPropertyBits=   25
         Caption         =   "Tim"
         Size            =   "1085;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.Frame freSearch 
      BorderStyle     =   0  'None
      Height          =   5745
      Index           =   2
      Left            =   4200
      TabIndex        =   10
      Top             =   3000
      Width           =   4380
      Begin VB.TextBox TxtEngineSearch 
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
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   3375
      End
      Begin VB.ListBox LstEngineSearch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3900
         ItemData        =   "OCS10 Database Software.frx":0452
         Left            =   120
         List            =   "OCS10 Database Software.frx":0454
         TabIndex        =   13
         Top             =   600
         Width           =   4095
      End
      Begin MSForms.CommandButton CmdEngineSearch 
         Height          =   375
         Left            =   3600
         TabIndex        =   28
         Top             =   120
         Width           =   615
         VariousPropertyBits=   25
         Caption         =   "Tim"
         Size            =   "1085;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   16440
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OCS10 Database Software.frx":0456
            Key             =   "KeyNew"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OCS10 Database Software.frx":0568
            Key             =   "KeyEdit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OCS10 Database Software.frx":067A
            Key             =   "KeyAbort"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OCS10 Database Software.frx":078C
            Key             =   "KeySave"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OCS10 Database Software.frx":089E
            Key             =   "KeyDelete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OCS10 Database Software.frx":09B0
            Key             =   "KeyUddate"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OCS10 Database Software.frx":0AC2
            Key             =   "KeyReport"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OCS10 Database Software.frx":0BD4
            Key             =   "KeyParameter"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OCS10 Database Software.frx":0CE6
            Key             =   "KeyExit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OCS10 Database Software.frx":0DF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OCS10 Database Software.frx":1472
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TbrMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   17775
      _ExtentX        =   31353
      _ExtentY        =   635
      ButtonWidth     =   2619
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Xoa"
            Key             =   "KeyDelete"
            Object.ToolTipText     =   "Delete one Car's Testing Result"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Import xe"
            Key             =   "KeyImport"
            Object.ToolTipText     =   "Import danh s�ch xe"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   11
            Style           =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Bao Cao"
            Key             =   "KeyReport"
            Object.ToolTipText     =   "Print report seperate"
            ImageIndex      =   7
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "KeyReportSelected"
                  Text            =   "Xe Hien Tai"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "KeyReportResultSearch"
                  Text            =   "Tat Ca Xe"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  Tieu Chuan  "
            Key             =   "KeyParameter"
            Object.ToolTipText     =   "Table Registered Parameters Of Cars"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "KeyRefresh"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "KeyExit"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
      EndProperty
   End
   Begin VB.Data DatTestingParameter 
      Caption         =   "Database Testing Parameter"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
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
      Caption         =   "Danh sa�ch xe"
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
      Height          =   9855
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   12255
      Begin MSDBGrid.DBGrid DBGTestingUpdate 
         Bindings        =   "OCS10 Database Software.frx":1AEC
         Height          =   9105
         Left            =   240
         OleObjectBlob   =   "OCS10 Database Software.frx":1B0E
         TabIndex        =   4
         Top             =   480
         Width           =   11820
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "T�m kie�m theo"
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
      Height          =   6015
      Left            =   12720
      TabIndex        =   0
      Top             =   5760
      Width           =   4935
      Begin MSComctlLib.TabStrip TabSearch 
         Height          =   5295
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   9340
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   5
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "So� SX"
               Key             =   "KeyName"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "So� khung"
               Key             =   "KeyChassisNo"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "So� ma�y"
               Key             =   "KeyEngineNo"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Nga�y KT"
               Key             =   "KeyDate"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Ta�t ca�"
               Key             =   "KeyAll"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "VNI-Centur"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Xe �ang test"
      BeginProperty Font 
         Name            =   "VNI-Centur"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   240
      TabIndex        =   38
      Top             =   360
      Width           =   17415
      Begin VB.Label Label64 
         Caption         =   "So� ma�y"
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   14040
         TabIndex        =   45
         Top             =   225
         Width           =   675
      End
      Begin VB.Label Label63 
         Caption         =   "So� sa�n xua�t"
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   10320
         TabIndex        =   44
         Top             =   225
         Width           =   1050
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "So� khung"
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6720
         TabIndex        =   43
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         Caption         =   "Loa�i xe"
         BeginProperty Font 
            Name            =   "VNI-Centur"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3000
         TabIndex        =   41
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Label Label62 
      Caption         =   "So� khung:"
      BeginProperty Font 
         Name            =   "VNI-Centur"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label59 
      Caption         =   "Loa�i xe:"
      BeginProperty Font 
         Name            =   "VNI-Centur"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label58 
      Caption         =   "Loa�i xe:"
      BeginProperty Font 
         Name            =   "VNI-Centur"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label56 
      AutoSize        =   -1  'True
      Caption         =   "C��ng ��(100xCd)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   1605
   End
   Begin VB.Label Label53 
      AutoSize        =   -1  'True
      Caption         =   "C��ng ��(100xCd)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   1605
   End
   Begin VB.Menu MnuFileOCS10 
      Caption         =   "He Thong"
      Begin VB.Menu MnuImportVehicles 
         Caption         =   "Import xe"
         Shortcut        =   ^I
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuReport 
      Caption         =   "Bao Cao"
      Begin VB.Menu MnuReportSelected 
         Caption         =   "Xe Hien Tai"
      End
      Begin VB.Menu i 
         Caption         =   "-"
      End
      Begin VB.Menu MnuReportTotal 
         Caption         =   "Tat Ca Xe"
      End
   End
   Begin VB.Menu MnuTable 
      Caption         =   "Du Lieu Khac"
      Begin VB.Menu MnuTester 
         Caption         =   "Nguoi Kiem Tra"
      End
      Begin VB.Menu MnuRegisteredParameter 
         Caption         =   "Bang Tieu Chuan Xe"
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "Tro Giup"
      Begin VB.Menu MnuHelpGuide 
         Caption         =   "Phan Mem OCS10"
      End
      Begin VB.Menu k 
         Caption         =   "-"
      End
      Begin VB.Menu MnuHelpAboutOCS10DBS 
         Caption         =   "ThietBiCongNghiep.vn"
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
        .Update
    End With
    DatTestingParameter.Refresh
    TxtChassisNumber.Text = ""
    TxtProducedNumber.Text = ""
    TxtEngineNumber.Text = ""
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

Private Sub btnSelectTest_Click()
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

 If RST("ProducedNumber") <> "" Then
    TxtSelectedProducedNumber.Text = RST("ProducedNumber")
Else
    TxtSelectedProducedNumber.Text = ""
End If

  If RST("EngineNumber") <> "" Then
    TxtSelectedEngineNumber.Text = RST("EngineNumber")
Else
    TxtSelectedEngineNumber.Text = ""
End If
  
MsgBox "Ban da chon xe test(" & RST("ProducedNumber") & ")"
Else
MsgBox "Record Not Found..."
End If
RST.Close
End Sub

Private Sub ChangePass_Click()
FrmPassword.Show
End Sub

 

Private Sub cldFromDate_Click()
TxtDateFrom.Text = cldFromDate.Value
cldFromDate.Visible = False
End Sub

Private Sub cldToDate_Click()
TxtDateTo.Text = cldToDate.Value
cldToDate.Visible = False

SearchFollowDate
End Sub

 

Private Sub CmdChassisSearch_Click()
Dim ChassisSearch As String
ChassisSearch = TxtChassisSearch.Text
'DatTestingParameter.RecordSource = "SELECT OrderMeasuringResult, Name, ChassisNumber, EngineNumber, Tester,  Date, ProducedNumber FROM TblTestingParameter WHERE Name = " & Chr$(34) & NameSearch & Chr$(34)
'Line code tren cho phep chi hien thi 04 colum, tuy nhien cac thong so khac khong show ben cac Textbox sau khi tim kiem, Cam kiem tra lai de chinh sua hoan chinh, tam thoi dung line code duoi day
DatTestingParameter.RecordSource = "SELECT * From TblTestingParameter WHERE ChassisNumber = " & Chr$(34) & ChassisSearch & Chr$(34)
txtSqlReport.Text = "SELECT * From TblTestingParameter WHERE ChassisNumber = " & Chr$(34) & ChassisSearch & Chr$(34)
DatTestingParameter.Refresh
End Sub

Private Sub CmdDateSearchFrom_Click()
cldToDate.Visible = False
If cldFromDate.Visible = False Then
cldFromDate.Visible = True
Else: cldFromDate.Visible = False
End If
End Sub

Private Sub CmdDateSearchTo_Click()
cldFromDate.Visible = False
If cldToDate.Visible = False Then
cldToDate.Visible = True
Else: cldToDate.Visible = False
End If
End Sub

Private Sub CmdEngineSearch_Click()
Dim EngineSearch As String
EngineSearch = TxtEngineSearch.Text
'DatTestingParameter.RecordSource = "SELECT OrderMeasuringResult, Name, ChassisNumber, EngineNumber, Tester,  Date, ProducedNumber FROM TblTestingParameter WHERE Name = " & Chr$(34) & NameSearch & Chr$(34)
'Line code tren cho phep chi hien thi 04 colum, tuy nhien cac thong so khac khong show ben cac Textbox sau khi tim kiem, Cam kiem tra lai de chinh sua hoan chinh, tam thoi dung line code duoi day
DatTestingParameter.RecordSource = "SELECT * From TblTestingParameter WHERE EngineNumber = " & Chr$(34) & EngineSearch & Chr$(34)
txtSqlReport.Text = "SELECT * From TblTestingParameter WHERE EngineNumber = " & Chr$(34) & EngineSearch & Chr$(34)
DatTestingParameter.Refresh
End Sub
 

Private Sub CmdNameSearch_Click()
Dim NameSearch As String
NameSearch = TxtNameSearch.Text
'DatTestingParameter.RecordSource = "SELECT OrderMeasuringResult, Name, ChassisNumber, EngineNumber, Tester,  Date, ProducedNumber FROM TblTestingParameter WHERE Name = " & Chr$(34) & NameSearch & Chr$(34)
'Line code tren cho phep chi hien thi 04 colum, tuy nhien cac thong so khac khong show ben cac Textbox sau khi tim kiem, Cam kiem tra lai de chinh sua hoan chinh, tam thoi dung line code duoi day
DatTestingParameter.RecordSource = "SELECT * From TblTestingParameter WHERE Name = " & Chr$(34) & NameSearch & Chr$(34)
txtSqlReport.Text = "SELECT * From TblTestingParameter WHERE Name = " & Chr$(34) & NameSearch & Chr$(34)
DatTestingParameter.Refresh
End Sub

Private Sub CmdShowAll_Click()
DatTestingParameter.Refresh
txtSqlReport.Text = "SELECT * From TblTestingParameter  order by STT desc"
End Sub


Private Sub DBGridOptionShow_Sub()
DatTestingParameter.RecordSource = "SELECT OrderMeasuringResult, Name, ChassisNumber, EngineNumber, Tester,  Date, ProducedNumber  FROM TblTestingParameter"
DatTestingParameter.Refresh
End Sub

Private Sub SearchFollowDate()
Dim FromDate As String
Dim ToDate As String
FromDate = TxtDateFrom.Text
ToDate = TxtDateTo.Text
DatTestingParameter.RecordSource = "SELECT * FROM TblTestingParameter WHERE Date >=#" & FromDate & "# AND Date <=#" & ToDate & "#"
txtSqlReport.Text = "SELECT * FROM TblTestingParameter WHERE Date >=#" & FromDate & "# AND Date <=#" & ToDate & "#"
DatTestingParameter.Refresh
End Sub


Private Sub DatTestingParameter_Reposition()
'Command1_Click
InitializeCheckingParameter
End Sub


 

Private Sub Form_Load()
DataBaseFolder = "\\Master\OCS10"
'DataBaseFolder = App.Path
txtSqlReport.Text = "SELECT * FROM TblTestingParameter"
DatTestingParameter.DatabaseName = DataBaseFolder & "\OCS10_DataBase_97.mdb"
DatTestingParameter.RecordSource = "select * from TblTestingParameter order by STT desc"

DatCheckingParameter.DatabaseName = DataBaseFolder & "\OCS10_DataBase_97.mdb"
DatCheckingParameter.RecordSource = "select * from TblCheckingParameter"

 
SearchingFramePosition
ListNameUpdate
ListNameSearch
ListChassisSearch
ListEngineSearch
ListTesterUpdate
LoadSelectVehicle

TxtDateFrom.Text = Date
TxtDateTo.Text = Date
cldFromDate.Value = Date
cldToDate.Value = Date
addTester
addName
End Sub

Private Sub TabSearchEnable()
'TxtNameSearch.Enabled = False
'TxtChassisNumber.Enabled = False
'TxtEngineNumber.Enabled = False
'TxtDate.Enabled = False
freSearch(0).Enabled = False
freSearch(1).Enabled = False
freSearch(2).Enabled = False
freSearch(3).Enabled = False
freSearch(4).Enabled = False
End Sub


Private Sub ListNameUpdate()
 '------------------------Cap nhan danh sach cho List Box  - Name of Car
Dim dbname_ln As String
Dim db_ln As Database
Dim rs_ln As Recordset

    ' Open the database.
    dbname_ln = DataBaseFolder
    If Right$(dbname_ln, 1) <> "\" Then dbname_ln = dbname_ln & "\"
    dbname_ln = dbname_ln & "OCS10_DataBase_97.mdb"

    Set db_ln = OpenDatabase(dbname_ln)
    Set rs_ln = db_ln.OpenRecordset( _
        "SELECT Name FROM TblCheckingParameter ORDER BY Name", _
        dbOpenSnapshot)

    ' Load the ComboBox.
    rs_ln.MoveFirst
    Do While Not rs_ln.EOF
        rs_ln.MoveNext
    Loop

    rs_ln.Close
    db_ln.Close

    ' Connect the Data control to the database.
   DatTestingParameter.DatabaseName = dbname_ln
 
End Sub
Private Sub ListNameSearch()
 '------------------------Cap nhan danh sach cho List Box  - Name Search Tested
Dim dbname_ln As String
Dim db_ln As Database
Dim rs_ln As Recordset

    ' Open the database.
    dbname_ln = DataBaseFolder
    If Right$(dbname_ln, 1) <> "\" Then dbname_ln = dbname_ln & "\"
    dbname_ln = dbname_ln & "OCS10_DataBase_97.mdb"

    Set db_ln = OpenDatabase(dbname_ln)
    Set rs_ln = db_ln.OpenRecordset( _
        "SELECT DISTINCT Name FROM TblTestingParameter ORDER BY Name", _
        dbOpenSnapshot)

    ' Load the ComboBox.
    If rs_ln.EOF = False Then
    rs_ln.MoveFirst
    Do While Not rs_ln.EOF
        LstNameSearch.AddItem rs_ln!Name
        rs_ln.MoveNext
    Loop
    End If

    rs_ln.Close
    db_ln.Close

    ' Connect the Data control to the database.
   DatTestingParameter.DatabaseName = dbname_ln

    ' Select the first choice.
    If LstNameSearch.ListCount > 0 Then
        LstNameSearch.ListIndex = 0
    End If
    
End Sub

Private Sub ListEngineSearch()
 '------------------------Cap nhan danh sach cho List Box  - Name Search Tested
Dim dbname_ln As String
Dim db_ln As Database
Dim rs_ln As Recordset

    ' Open the database.
    dbname_ln = DataBaseFolder
    If Right$(dbname_ln, 1) <> "\" Then dbname_ln = dbname_ln & "\"
    dbname_ln = dbname_ln & "OCS10_DataBase_97.mdb"

    Set db_ln = OpenDatabase(dbname_ln)
    Set rs_ln = db_ln.OpenRecordset( _
        "SELECT DISTINCT EngineNumber FROM TblTestingParameter ORDER BY EngineNumber", _
        dbOpenSnapshot)

    ' Load the ComboBox.
    If rs_ln.EOF = False Then
    rs_ln.MoveFirst
    Do While Not rs_ln.EOF
        If rs_ln!EngineNumber <> "" Then
            LstEngineSearch.AddItem rs_ln!EngineNumber
        End If
        
     
        rs_ln.MoveNext
    Loop
    End If
    rs_ln.Close
    db_ln.Close

    ' Connect the Data control to the database.
   DatTestingParameter.DatabaseName = dbname_ln

    ' Select the first choice.
     If LstEngineSearch.ListCount > 0 Then
        LstEngineSearch.ListIndex = 0
    End If
        
End Sub
Private Sub ListChassisSearch()
 '------------------------Cap nhan danh sach cho List Box  - Name Search Tested
Dim dbname_ln As String
Dim db_ln As Database
Dim rs_ln As Recordset

    ' Open the database.
    dbname_ln = DataBaseFolder
    If Right$(dbname_ln, 1) <> "\" Then dbname_ln = dbname_ln & "\"
    dbname_ln = dbname_ln & "OCS10_DataBase_97.mdb"

    Set db_ln = OpenDatabase(dbname_ln)
    Set rs_ln = db_ln.OpenRecordset( _
        "SELECT DISTINCT ChassisNumber FROM TblTestingParameter ORDER BY ChassisNumber", _
        dbOpenSnapshot)

    ' Load the ComboBox.
    If rs_ln.EOF = False Then
    rs_ln.MoveFirst
    Do While Not rs_ln.EOF
      If rs_ln!ChassisNumber <> "" Then
        LstChassisSearch.AddItem rs_ln!ChassisNumber
        End If
      rs_ln.MoveNext
    Loop
    End If
    rs_ln.Close
    db_ln.Close

    ' Connect the Data control to the database.
   DatTestingParameter.DatabaseName = dbname_ln
 
    ' Select the first choice.
    If LstChassisSearch.ListCount > 0 Then
        LstChassisSearch.ListIndex = 0
    End If
End Sub

Private Sub ListTesterUpdate()
'------------------------Cap nhan danh sach cho List Box  - Tester
Dim dbname_lt As String
Dim db_lt As Database
Dim rs_lt As Recordset

    ' Open the database.
    dbname_lt = DataBaseFolder
    If Right$(dbname_lt, 1) <> "\" Then dbname_lt = dbname_lt & "\"
    dbname_lt = dbname_lt & "OCS10_DataBase_97.mdb"

    Set db_lt = OpenDatabase(dbname_lt)
    Set rs_lt = db_lt.OpenRecordset( _
        "SELECT Name FROM TblTesters ORDER BY Name", _
        dbOpenSnapshot)

    ' Load the ComboBox.
    rs_lt.MoveFirst
    Do While Not rs_lt.EOF
        rs_lt.MoveNext
    Loop

    rs_lt.Close
    db_lt.Close

    ' Connect the Data control to the database.
   DatTestingParameter.DatabaseName = dbname_lt
 
End Sub
Private Sub SearchingFramePosition()
Dim i As Integer

    ' Move all the frames to the same position
    ' and make them all invisible.
    For i = 0 To freSearch.UBound
        freSearch(i).Move _
            freSearch(0).Left, _
            freSearch(0).Top, _
            freSearch(0).Width, _
            freSearch(0).Height
        freSearch(i).Visible = False
    Next i
    
    ' Select the first tab.
    SelectedTab = 5
    TabSearch.SelectedItem = TabSearch.Tabs(SelectedTab)
    freSearch(SelectedTab - 1).Visible = True
End Sub

 

 

 

 

Private Sub Login_Click()
FrmLogin.Show
End Sub
Sub BeginUpdate()

Dim Index As Integer
For Index = 1 To 13
    TbrMain.Buttons.Item(Index).Visible = True
Next Index
End Sub

Function GetPassword() As String
    Dim connect As New ADODB.Connection
    Dim RST As New ADODB.Recordset
    Dim password As String
    If connect.State = 1 Then connect.Close
    If RST.State = 1 Then RST.Close
    connect.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source=" & DataBaseFolder & "\OCS10_DataBase_97.mdb;Persist Security Info=False"
    
    Dim sSQL As String
    sSQL = "Select * From TblPassword"
    RST.Open sSQL, connect, adOpenDynamic, adLockOptimistic
    If Not RST.EOF Then
    password = RST("Password")
    RST.Close
    GetPassword = password
    End If
End Function

Sub SetPassword(newPassword As String)
    Dim connect As New ADODB.Connection
    Dim RST As New ADODB.Recordset
    Dim password As String
    If connect.State = 1 Then connect.Close
    If RST.State = 1 Then RST.Close
    connect.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source=" & DataBaseFolder & "\OCS10_DataBase_97.mdb;Persist Security Info=False"
    
    Dim sSQL As String
    sSQL = "Select * From TblPassword"
    RST.Open sSQL, connect, adOpenDynamic, adLockOptimistic
    If Not RST.EOF Then
    RST("Password") = newPassword
    RST.Update
    End If
End Sub
 

Private Sub LstChassisSearch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtChassisSearch = LstChassisSearch
End Sub

Private Sub LstEngineSearch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtEngineSearch = LstEngineSearch
End Sub


Private Sub LstNameSearch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtNameSearch = LstNameSearch
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
    MsgBox "Th�m th�nh c�ng " & CStr(i - 3) & " xe"
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
    
     If RST("ProducedNumber") <> "" Then
        TxtSelectedProducedNumber.Text = RST("ProducedNumber")
    Else
        TxtSelectedProducedNumber.Text = ""
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

Private Sub MnuTester_Click()
FrmTester.Show
End Sub

Private Sub TabSearch_Click()
    freSearch(SelectedTab - 1).Visible = False
    SelectedTab = TabSearch.SelectedItem.Index
    freSearch(SelectedTab - 1).Visible = True
    Select Case SelectedTab
    Case 1
        freSearch(0).Enabled = True
        TxtNameSearch.SetFocus
    Case 2
        freSearch(1).Enabled = True
        TxtChassisSearch.SetFocus
     
    Case 3
        freSearch(2).Enabled = True
        TxtEngineSearch.SetFocus
    
    Case 4
        freSearch(3).Enabled = True
    Case 5
        freSearch(4).Enabled = True
    DatTestingParameter.RecordSource = "SELECT * FROM TblTestingParameter order by STT desc"
    End Select
End Sub
  

Private Sub MnuAbort_Click()
DBGTestingUpdate.Enabled = True
Frame9.Enabled = True
freSearch(0).Enabled = True
freSearch(1).Enabled = True
freSearch(2).Enabled = True
freSearch(3).Enabled = True

TbrMain.Buttons(3).Enabled = True
TbrMain.Buttons(5).Enabled = True
SubDisableAll
On Error GoTo ErrHandling
DatTestingParameter.Recordset.CancelUpdate
TbrMain.Buttons(7).Enabled = False
TbrMain.Buttons(9).Enabled = False
TbrMain.Buttons(11).Enabled = True
TbrMain.Buttons(15).Enabled = True
TbrMain.Buttons(17).Enabled = True

 
MnuReportSelected.Enabled = True
MnuReportTotal.Enabled = True
MnuRegisteredParameter.Enabled = True
EndIt:
Exit Sub ' or Exit Function

ErrHandling:
SubErrHandling
Resume EndIt
End Sub

Private Sub MnuAddNew_Click()
TbrMain.Buttons(9).Enabled = True

On Error GoTo ErrHandling
'DatTestingParameter.Recordset.AddNew

With DatTestingParameter.Recordset
        .AddNew
        !Name = "Name"
        .Update
    End With
    DatTestingParameter.Refresh
DatTestingParameter.Recordset.Edit

TbrMain.Buttons(5).Enabled = False

TbrMain.Buttons(11).Enabled = False
TbrMain.Buttons(15).Enabled = False
TbrMain.Buttons(17).Enabled = False
 
MnuReportSelected.Enabled = False
MnuReportTotal.Enabled = False
MnuRegisteredParameter.Enabled = False

EndIt:
Exit Sub ' or Exit Function
ErrHandling:
SubErrHandling
Resume EndIt
End Sub
Private Sub MoveNextRecord()
DatTestingParameter.Recordset.MoveNext
    If DatTestingParameter.Recordset.EOF = True Then
        DatTestingParameter.Recordset.MoveLast
    End If
End Sub
Private Sub MnuDeleteParameter_Click()
 Dim RecordCurrent As Integer
 Dim RecordCount As Integer
 On Error GoTo Delete_Error

    If MsgBox("Are you sure you want to delete this record?", _
                vbQuestion + vbYesNo + vbDefaultButton2, _
            "Confirm") = vbNo Then
            Exit Sub
        End If

    'delete the current record
    RecordCurrent = DatTestingParameter.Recordset.AbsolutePosition
    RecordCount = DatTestingParameter.Recordset.RecordCount
    If RecordCurrent < RecordCount - 1 Then
    DatTestingParameter.Recordset.Delete
       DatTestingParameter.Refresh
       DatTestingParameter.Recordset.Move (RecordCurrent)
       Else: MsgBox (" Last Record should not be Deleted !")
    End If
    
EndIt:
    Exit Sub
Delete_Error:
SubErrHandling
Resume EndIt
End Sub

Private Sub MnuEditResult_Click()
DBGTestingUpdate.Enabled = False
Frame9.Enabled = False
freSearch(0).Enabled = False
freSearch(1).Enabled = False
freSearch(2).Enabled = False
freSearch(3).Enabled = False

On Error GoTo ErrHandling
SubEnableAll
DatTestingParameter.Recordset.Edit
TbrMain.Buttons(3).Enabled = False
TbrMain.Buttons(5).Enabled = False
TbrMain.Buttons(7).Enabled = True
TbrMain.Buttons(9).Enabled = True
TbrMain.Buttons(11).Enabled = False
TbrMain.Buttons(15).Enabled = False
TbrMain.Buttons(17).Enabled = False
MnuReportSelected.Enabled = False
MnuReportTotal.Enabled = False
MnuRegisteredParameter.Enabled = False

EndIt:
Exit Sub ' or Exit Function

ErrHandling:
SubErrHandling
Resume EndIt
End Sub

Private Sub MnuExit_Click()
If MsgBox("Quit now. Are you sure ?", _
                vbQuestion + vbYesNo + vbDefaultButton2, _
            "Confirm") = vbNo Then
            Exit Sub
        End If
End
End Sub

Private Sub MnuRegisteredParameter_Click()
FrmCheckingParameter.Show
End Sub

Private Sub MnuReportSeperate_Click()
FrmReportSeperate.Show
End Sub
Private Sub MnuReportSelected_Click()
If Len(txtCurrentID) > 0 Then
FrmReportSelected.Show
Else: MsgBox "No record"
End If
End Sub

Private Sub MnuSave_Click()
DBGTestingUpdate.Enabled = True
Frame9.Enabled = True
freSearch(0).Enabled = True
freSearch(1).Enabled = True
freSearch(2).Enabled = True
freSearch(3).Enabled = True

TbrMain.Buttons(3).Enabled = True
TbrMain.Buttons(5).Enabled = True
 

SubDisableAll
On Error GoTo ErrHandling
DatTestingParameter.Recordset.Update
TbrMain.Buttons(7).Enabled = False
TbrMain.Buttons(9).Enabled = False
TbrMain.Buttons(11).Enabled = True
TbrMain.Buttons(15).Enabled = True
TbrMain.Buttons(17).Enabled = True

MnuReportSelected.Enabled = False
MnuReportTotal.Enabled = False
MnuRegisteredParameter.Enabled = False


'DatTestingParameter.Recordset.MoveNext
'Dich chuyen xuong cuoi bang DBG cho phu hop de tien theo doi.
EndIt:
Exit Sub ' or Exit Function

ErrHandling:
SubErrHandling
Resume EndIt

End Sub

Private Sub MnuUpdateParameter_Click()
DatTestingParameter.Refresh
End Sub



Private Sub TbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key

Case "KeyNew"
Case "KeyEdit"
MnuEditResult_Click

Case "KeyAbort"
MnuAbort_Click

Case "KeySave"
MnuSave_Click

Case "KeyDelete"
MnuDeleteParameter_Click

Case "KeyRefresh"
MnuUpdateParameter_Click

Case "KeyImport"
MnuImportVehicles_Click

 

Case "KeyReport"
MnuReportSeperate_Click

Case "KeyParameter"
MnuRegisteredParameter_Click

Case "KeyExit"
MnuExit_Click

End Select

End Sub

Private Sub TbrMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
   Select Case ButtonMenu.Key
      Case "KeyReportSelected"
         MnuReportSelected_Click
      Case "KeyReportResultSearch"
        MnuReportSeperate_Click
   End Select
End Sub
  
'Author: Thinh Ga COn
'Date : 15/03/2012
'thuc hien tim kiem khi thay doi text search textbox
Private Sub TxtNameSearch_Change()
    'kiem tra truong nhap vao co trong k
    If TxtNameSearch.Text = "" Then
    CmdNameSearch.Enabled = False
    Else
    CmdNameSearch.Enabled = True
    End If
    'end
    
    Dim dbname_ln As String
    Dim db_ln As Database
    Dim rs_ln As Recordset
    
    LstNameSearch.Clear
    
    Dim strSearch As String
    Dim strSql As String
    
    strSearch = Trim(TxtNameSearch.Text)
    strSql = "SELECT DISTINCT Name FROM TblTestingParameter WHERE Name LIKE '*" & strSearch & "*'"
    
    dbname_ln = DataBaseFolder
    If Right$(dbname_ln, 1) <> "\" Then dbname_ln = dbname_ln & "\"
    dbname_ln = dbname_ln & "OCS10_DataBase_97.mdb"
    Set db_ln = OpenDatabase(dbname_ln)
    Set rs_ln = db_ln.OpenRecordset( _
        strSql, _
        dbOpenSnapshot)
        
        If rs_ln.RecordCount > 0 Then
    rs_ln.MoveFirst
    
    Do While Not rs_ln.EOF
        LstNameSearch.AddItem rs_ln!Name
        rs_ln.MoveNext
    Loop
    LstNameSearch.ListIndex = 0
    End If
    
    rs_ln.Close
    db_ln.Close
    DatTestingParameter.DatabaseName = dbname_ln
End Sub


Private Sub TxtChassisSearch_Change()
    'kiem tra truong nhap vao co trong k
    If TxtChassisSearch.Text = "" Then
    CmdChassisSearch.Enabled = False
    Else
    CmdChassisSearch.Enabled = True
    End If
    'end
    
    Dim dbname_ln As String
    Dim db_ln As Database
    Dim rs_ln As Recordset
    
    LstChassisSearch.Clear
    
    Dim strSearch As String
    Dim strSql As String
    
    strSearch = Trim(TxtChassisSearch.Text)
    strSql = "SELECT DISTINCT ChassisNumber FROM TblTestingParameter WHERE ChassisNumber LIKE '*" & strSearch & "*'"
    
    dbname_ln = DataBaseFolder
    If Right$(dbname_ln, 1) <> "\" Then dbname_ln = dbname_ln & "\"
    dbname_ln = dbname_ln & "OCS10_DataBase_97.mdb"
    Set db_ln = OpenDatabase(dbname_ln)
    Set rs_ln = db_ln.OpenRecordset( _
        strSql, _
        dbOpenSnapshot)
        
        If rs_ln.RecordCount > 0 Then
    rs_ln.MoveFirst
    
    Do While Not rs_ln.EOF
        LstChassisSearch.AddItem rs_ln!ChassisNumber
        rs_ln.MoveNext
    Loop
    LstChassisSearch.ListIndex = 0
    End If
    
    rs_ln.Close
    db_ln.Close
    DatTestingParameter.DatabaseName = dbname_ln
End Sub

Private Sub TxtEngineSearch_Change()
    'kiem tra truong nhap vao co trong k
    If TxtEngineSearch.Text = "" Then
    CmdEngineSearch.Enabled = False
    Else
    CmdEngineSearch.Enabled = True
    End If
    'end
    
    Dim dbname_ln As String
    Dim db_ln As Database
    Dim rs_ln As Recordset
    
    LstEngineSearch.Clear
    
    Dim strSearch As String
    Dim strSql As String
    
    strSearch = Trim(TxtEngineSearch.Text)
    strSql = "SELECT DISTINCT EngineNumber FROM TblTestingParameter WHERE EngineNumber LIKE '*" & strSearch & "*'"
    
    dbname_ln = DataBaseFolder
    If Right$(dbname_ln, 1) <> "\" Then dbname_ln = dbname_ln & "\"
    dbname_ln = dbname_ln & "OCS10_DataBase_97.mdb"
    Set db_ln = OpenDatabase(dbname_ln)
    Set rs_ln = db_ln.OpenRecordset( _
        strSql, _
        dbOpenSnapshot)
        
        If rs_ln.RecordCount > 0 Then
    rs_ln.MoveFirst
    
    Do While Not rs_ln.EOF
        LstEngineSearch.AddItem rs_ln!EngineNumber
        rs_ln.MoveNext
    Loop
    LstEngineSearch.ListIndex = 0
    End If
    
    rs_ln.Close
    db_ln.Close
    DatTestingParameter.DatabaseName = dbname_ln
End Sub

'enter key was pressed in txtNameSearch
Private Sub TxtNameSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ' enter key was pressed
        KeyAscii = 0 'suppress sound
        TxtNameSearch = LstNameSearch
        'dua con tro chuot ve cuoi textbox
        TxtNameSearch.SelStart = Len(TxtNameSearch.Text)
        'thuc hien tim kiem luon
        CmdNameSearch_Click
    End If
End Sub

'enter key was pressed in TxtChassisSearch
Private Sub TxtChassisSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ' enter key was pressed
        KeyAscii = 0 'suppress sound
        TxtChassisSearch = LstChassisSearch
        'dua con tro chuot ve cuoi textbox
        TxtChassisSearch.SelStart = Len(TxtChassisSearch.Text)
        
        CmdChassisSearch_Click
    End If
End Sub

'enter key was pressed in TxtEngineSearch
Private Sub TxtEngineSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ' enter key was pressed
        KeyAscii = 0 'suppress sound
        TxtEngineSearch = LstEngineSearch
        'dua con tro chuot ve cuoi textbox
        TxtEngineSearch.SelStart = Len(TxtEngineSearch.Text)
        
        CmdEngineSearch_Click
    End If
End Sub

 


 

