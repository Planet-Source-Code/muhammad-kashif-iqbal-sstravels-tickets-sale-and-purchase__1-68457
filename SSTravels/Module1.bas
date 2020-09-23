Attribute VB_Name = "Module1"
Public Cn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public kashif As String
Public Qty As Long
Public lst As ListItem

Sub Opendatabase()

con = "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=SSTravels;Data Source = home"

If Cn.State = 0 Then
    Cn.Open con
End If

End Sub
