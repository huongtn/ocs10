VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form FrmReportSeperate 
   Caption         =   "Report Seperate"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   19080
   Icon            =   "ReportSeperate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   14723.81
   ScaleMode       =   0  'User
   ScaleWidth      =   19200
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   11415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18855
      lastProp        =   500
      _cx             =   33258
      _cy             =   20135
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "FrmReportSeperate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FileName   : FrmReportSeperate
'Author     : Thinhgacon
'Date       : 20/03/2011

'General Declarations
Dim myRS As DAO.Recordset
Dim myDB As DAO.Database
Dim Appl As New CRAXDRT.Application
Dim Report As New CRAXDRT.Report

'when load form
Private Sub Form_Load()
Dim sqlToReport As String
sqlToReport = FrmMain.txtSqlReport.Text
Set myDB = OpenDatabase(App.Path & "\OCS10_DataBase_97.mdb")

Set myRS = myDB.OpenRecordset(sqlToReport)
Set Report = Appl.OpenReport(".\OCS10Rpt.Rpt")

Report.Database.Tables(1).Location = App.Path & "\OCS10_DataBase_97.mdb"

Report.Database.SetDataSource myRS

CRViewer91.ReportSource = Report
CRViewer91.ViewReport
End Sub
'when risize form
'This above code resizes the Viewer Control to the size of the form
Private Sub Form_Resize()
 With CRViewer91
  .Top = 0
  .Left = 0
  .Width = Me.ScaleWidth
  .Height = Me.ScaleHeight
 End With

End Sub
