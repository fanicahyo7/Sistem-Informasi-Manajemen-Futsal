Attribute VB_Name = "Module2"
Public Koneksi As New ADODB.Connection
Public rsuser As ADODB.Recordset
Public Sub BukaDB()
On Error GoTo HELL

Set Koneksi = New ADODB.Connection
Set rsuser = New ADODB.Recordset
Koneksi.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=konekdb"
Exit Sub

HELL:
MsgBox "Database Tidak Terkoneksi", vbCritical + vbInformation, "Peringatan"
End
End Sub
Public Function TandaPetik(sText)
Dim sTemp
    If Not IsNull(sText) Then sTemp = Replace(sText, "'", "''")
    TandaPetik = sTemp
End Function
