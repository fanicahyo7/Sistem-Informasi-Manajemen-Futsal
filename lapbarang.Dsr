VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} LapBarang 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport2"
   ClientHeight    =   11055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15405
   Icon            =   "lapbarang.dsx":0000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   27173
   _ExtentY        =   19500
   SectionData     =   "lapbarang.dsx":9E4A
End
Attribute VB_Name = "LapBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub ActiveReport_ReportStart()
On Error Resume Next
Call BukaDB
strsql = "select * from MstBarang order by KodeBarang"
ado.ConnectionString = Koneksi
ado.Source = strsql
End Sub

Sub mati()
Unload Me
End Sub

Private Sub Detail_Format()
With ado.Recordset
    If Not .EOF Then
        Field1.Text = !KodeBarang
        Field2.Text = !NamaBarang
        Field3.Text = !HargaBeli
        Field4.Text = !HargaJual
        Field5.Text = !stok
    End If
End With
End Sub

