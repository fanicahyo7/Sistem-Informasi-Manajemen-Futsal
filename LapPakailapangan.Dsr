VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} LapPakaiLapangan 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   11055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20370
   Icon            =   "LapPakailapangan.dsx":0000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   35930
   _ExtentY        =   19500
   SectionData     =   "LapPakailapangan.dsx":9E4A
End
Attribute VB_Name = "LapPakaiLapangan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
Dim strsql As String

On Error Resume Next
If frmlappakai.Option1 = True Then
strsql = "select * from TrPakaiLapangan"
Label18 = ""
ElseIf frmlappakai.Option2 = True Then
strsql = "select * from TrPakaiLapangan where Tanggal >='" & Format(frmlappakai.DTPicker1, "YYYY/MM/DD") & "' and Tanggal<='" & Format(frmlappakai.DTPicker2, "YYYY/MM/DD") & "'"
Label18 = "Periode " & Format(frmlappakai.DTPicker1, "DD/MM/YYYY") & " Sampai " & Format(frmlappakai.DTPicker2, "DD/MM/YYYY") & ""
End If
Call BukaDB
ado.ConnectionString = Koneksi
ado.Source = strsql

Dim jumlah As Long
With frmlappembelian.Adodc1
.ConnectionString = Koneksi
.RecordSource = strsql
.Refresh
    .Recordset.MoveFirst
    Do Until .Recordset.EOF
        jumlah = jumlah + .Recordset!GrandTotalharga
        .Recordset.MoveNext
    Loop
    Field8.Text = "GrandTotal Rp. = " & jumlah
End With
End Sub

Private Sub Detail_Format()
With ado.Recordset
    If Not .EOF Then
        Field1.Text = !NoPakaiLapangan
        Field2.Text = !NoBooking
        Field3.Text = !Kodelapangan
    With frmlappembelian.Adodc2
    .RecordSource = "select * from Mstlapangan where KodeLapangan='" & Field3 & "'"
    .Refresh
    Field3 = .Recordset!NamaLapangan
    End With
        Field4.Text = !Tanggal
        Field5.Text = !HargaSewalapangan
        Field6.Text = !TotalPembelian
        Field7.Text = !GrandTotalharga
    End If
End With
End Sub
Sub mati()
Unload Me
End Sub

