VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} laprevisistok 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   11055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   35930
   _ExtentY        =   19500
   SectionData     =   "laprevisistok.dsx":0000
End
Attribute VB_Name = "laprevisistok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub mati()
Unload Me
End Sub

Private Sub ActiveReport_ReportStart()
Dim strsql As String
Call BukaDB
On Error Resume Next
If frmlaprevisi.Option2 = True Then
Label30 = "Periode " & Format(frmlaprevisi.DTPicker1, "DD/MM/YYYY") & " Sampai " & Format(frmlaprevisi.DTPicker2, "DD/MM/YYYY") & ""
strsql = "select * from TrRevisiStok where Tanggal>='" & Format(frmlaprevisi.DTPicker1, "YYYY/MM/DD") & "' and Tanggal<=' " & Format(frmlaprevisi.DTPicker2, "YYYY/MM/DD") & "'"
Else
If frmlaprevisi.Option1 = True Then
Label30 = ""
strsql = "select * from TrRevisiStok"
End If
End If
DataControl1.ConnectionString = Koneksi
DataControl1.Source = strsql
End Sub

Private Sub Detail_Format()
With DataControl1.Recordset
If Not .EOF Then
Field1.Text = !KodeTrans
Field2.Text = !Keterangan
Field3.Text = !Tanggal
Field4.Text = !NoUrut
Field5.Text = !KodeBarang
Field6.Text = !StokLama
Field7.Text = !StokBaru
End If
End With
End Sub
