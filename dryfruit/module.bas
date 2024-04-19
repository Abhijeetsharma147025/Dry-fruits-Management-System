Attribute VB_Name = "Module"
Public C As New ADODB.Connection
Public R As New ADODB.Recordset
Public SQL As String



Public Function CONN()
Set C = New ADODB.Connection
C.Open "Provider=MSDAORA.1;User ID=ANIKET/ARPIT;Persist Security Info=True"
End Function
