VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} LapPembelian 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   11055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20370
   Icon            =   "LapPembelian.dsx":0000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   35930
   _ExtentY        =   19500
   SectionData     =   "LapPembelian.dsx":9E4A
End
Attribute VB_Name = "LapPembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
Dim strsql As String

On Error Resume Next
If frmlappembelian.Option2 = True Then
Label19 = "Periode " & Format(frmlappembelian.DTPicker1, "DD/MM/YYYY") & " Sampai " & Format(frmlappembelian.DTPicker2, "DD/MM/YYYY") & ""
strsql = "select * from ItemPembelian where Tanggal>='" & Format(frmlappembelian.DTPicker1, "YYYY/MM/DD") & "' and Tanggal<=' " & Format(frmlappembelian.DTPicker2, "YYYY/MM/DD") & "'"
Else
If frmlappembelian.Option1 = True Then
Label19 = ""
strsql = "select * from ItemPembelian"
End If
End If

DataControl1.ConnectionString = Koneksi
DataControl1.Source = strsql

Dim jumlah As Long
With frmlappembelian.Adodc1
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
Field1.Text = !KodePembelian
Field2.Text = !Tanggal
Field4.Text = !KodeBarang

With frmlappembelian.Adodc2
.RecordSource = "select * from MstBarang where KodeBarang='" & Field4 & "'"
.Refresh
Field4.Text = .Recordset!NamaBarang
End With

With frmlappembelian.Adodc1
.RecordSource = "select * from TrPembelian where KodePembelian='" & Field1 & "'"
.Refresh
Field3.Text = .Recordset!KodeSupplier
End With

With frmlappembelian.Adodc2
.RecordSource = "select * from MstSupplier where KodeSupplier='" & Field3.Text & "'"
.Refresh
Field3.Text = .Recordset!NamaSupplier
End With

Field5.Text = !Harga
Field6.Text = !jumlah
Field7.Text = !Total
End If
End With
End Sub

Sub mati()
Unload Me
End Sub
