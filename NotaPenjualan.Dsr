VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} NotaPenjualan 
   BorderStyle     =   0  'None
   Caption         =   "NotaPenjualan"
   ClientHeight    =   11055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   35930
   _ExtentY        =   19500
   SectionData     =   "NotaPenjualan.dsx":0000
End
Attribute VB_Name = "NotaPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
With DataControl1.Recordset
If Not .EOF Then
Field9.Text = !NoUrut
Field4.Text = !KodeBarang
Field5.Text = !Jumlah
Field6.Text = !Harga
Field7.Text = !Total
End If
End With
End Sub

Private Sub GroupFooter1_Format()
Label24 = frmMain.SSPanel3
End Sub
Sub mati()
Unload Me
End Sub

