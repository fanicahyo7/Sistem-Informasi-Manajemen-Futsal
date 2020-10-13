VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} LapPenjualan 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   9360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   Icon            =   "LapPenjualan.dsx":0000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   27093
   _ExtentY        =   16510
   SectionData     =   "LapPenjualan.dsx":9E4A
End
Attribute VB_Name = "LapPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
Dim strsql As String

On Error Resume Next
If frmlappenjualan.Option1 = True Then
Label18.Caption = ""
strsql = "select * from ItemPenjualan order by NoUrut"
ElseIf frmlappenjualan.Option2 = True Then
Label18.Caption = "Periode " & Format(frmlappenjualan.DTPicker1, "DD/MM/YYYY") & " Sampai " & Format(frmlappenjualan.DTPicker2, "DD/MM/YYYY") & ""
strsql = "select * from ItemPenjualan where Tanggal>='" & Format(frmlappenjualan.DTPicker1, "YYYY/MM/DD") & "' and Tanggal<=' " & Format(frmlappenjualan.DTPicker2, "YYYY/MM/DD") & "'"
End If

DataControl1.ConnectionString = Koneksi
DataControl1.Source = strsql

Dim jumlah As Long
With frmlappenjualan.Adodc1
.ConnectionString = Koneksi
.RecordSource = strsql
.Refresh
    .Recordset.MoveFirst
    Do Until .Recordset.EOF
        jumlah = jumlah + .Recordset!Total
        .Recordset.MoveNext
    Loop
    Field8.Text = "GrandTotal Rp. = " & jumlah
End With
End Sub

Private Sub Detail_Format()
With DataControl1.Recordset
If Not .EOF Then
Field7 = !KodePenjualan
Field1 = !NoUrut
Field2 = !Tanggal
Field3 = !KodeBarang
With frmlappembelian.Adodc2
.RecordSource = "select * from MstBarang where KodeBarang='" & Field3 & "'"
.Refresh
Field3 = .Recordset!NamaBarang
End With
Field4 = !jumlah
Field5 = !Harga
Field6 = !Total

End If
End With
End Sub
Sub mati()
Unload Me
End Sub

