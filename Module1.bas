Attribute VB_Name = "Module1"
Public cn As ADODB.Connection
Public rs As ADODB.Recordset
Sub main()
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
cn.ConnectionString = "Provider=Microsoft.jet.oledb.4.0;data source =" & (App.Path & "\Test.mdb") & ";"
cn.Open
Form1.Show

End Sub


