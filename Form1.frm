VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport cr 
      Left            =   360
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   8040
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim oapp As CRAXDDRT.Application
'Dim oreport As CRAXDDRT.Report
'rs.Open "select * from Table1", cn, adOpenKeyset, adLockOptimistic
'Set oapp = New CRAXDDRT.Application
'Set oreport = oapp.OpenReport(App.Path & "\DBS10_Report001.Rpt", 1)
'oreport.Database.SetDataSource rs, 3, 1
'CRViewer91.ReportSource = oreport
'CRViewer91.ViewReport
With cr
    .ReportFileName = App.Path & "\DBS10_Report002.Rpt"
    .WindowState = crptMaximized
    '.ReplaceSelectionFormula "{TblTestingParameter.OrderMeasuringResult}=17"
    .Destination = crptToWindow
    .Action = 1
            
End With



End Sub

