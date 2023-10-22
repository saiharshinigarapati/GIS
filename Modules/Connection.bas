Attribute VB_Name = "Connection"
Public rs As New ADODB.Recordset
Public cn As New ADODB.Connection
Public lst As MSComctlLib.ListItem
Public ADD_REC                          As New ADODB.Command

Sub Main()
With cn
.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\GeograhicMap.mdb;Persist Security Info=False"
.Open
End With
'mmain.Show
frmLogin.Show 1

End Sub



