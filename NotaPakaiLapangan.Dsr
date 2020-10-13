VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} NotaPakaiLapangan 
   BorderStyle     =   0  'None
   ClientHeight    =   11055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20370
   Icon            =   "NotaPakaiLapangan.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   35930
   _ExtentY        =   19500
   SectionData     =   "NotaPakaiLapangan.dsx":9E4A
End
Attribute VB_Name = "NotaPakaiLapangan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportEnd()
frmpakailapangan.kosong
End Sub

Private Sub Detail_Format()
With ado.Recordset
If Not .EOF Then
Field11.Text = !KodeBarang
Field13.Text = !Jumlah
Field14.Text = !Harga
Field15.Text = !Total

With frmpakailapangan.Adodc3
.RecordSource = "select * from MstBarang where KodeBarang='" & Field11.Text & "'"
.Refresh
Field12.Text = .Recordset!NamaBarang
End With
End If
End With
End Sub

Private Sub GroupFooter1_Format()
Label20.Caption = frmMain.SSPanel3
End Sub

Private Sub PageHeader_Format()
'With frmtotalboju.Adodc1
'.RecordSource = "select * from TrPakaiLapangan where NoPakaiLapangan='" & frmpakailapangan.Text1 & "'"
'.Refresh
Field16.Visible = False
Call BukaDB
Dim anu As New ADODB.Recordset
anu.Open "select * from TrPakaiLapangan where NoPakaiLapangan='" & Field16 & "'", Koneksi
If Not anu.EOF Then
        Field1.Text = anu!NoPakaiLapangan
        Field2.Text = anu!NoBooking
        Field3.Text = anu!Kodelapangan
        Field4.Text = anu!DP
        Field5.Text = anu!HargaSewalapangan
        Field6.Text = anu!TotalPembelian
        If Field6 = 0 Then
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
End If
        Field7.Text = anu!GrandTotalHarga
    With frmpakailapangan.Adodc2
    .RecordSource = "select * from TrBooking where NoBooking='" & Field2.Text & "'"
    .Refresh
    Field9.Text = .Recordset!AtasNama
    End With
    With frmpakailapangan.Adodc4
    .RecordSource = "select * from Mstlapangan where Kodelapangan='" & Field3 & "'"
    .Refresh
    Field10.Text = .Recordset!NamaLapangan
    End With
    End If
'End With
End Sub
Sub mati()
Unload Me
End Sub
