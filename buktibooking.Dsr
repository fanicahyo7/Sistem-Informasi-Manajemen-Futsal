VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} buktibooking 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11055
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   20370
   ControlBox      =   0   'False
   Icon            =   "buktibooking.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   35930
   _ExtentY        =   19500
   SectionData     =   "buktibooking.dsx":9E4A
End
Attribute VB_Name = "buktibooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Detail_Format()

With ado.Recordset
    If Not .EOF Then
        Field1.Text = !NoBooking
        Field2.Text = !Tanggal
        Field3.Text = !TanggalBooking
        Field4.Text = !JamMulai
        Field5.Text = !JamSelesai
        Field6.Text = !Kodelapangan
        Field7.Text = !NoRegister
        Field8.Text = !Harga
        Field9.Text = !DP
        Field10.Text = !AtasNama
        Field11.Text = Val(Field8.Text) - Val(Field9.Text)
    End If
End With

Label18.Caption = frmMain.SSPanel3.Caption
End Sub

Sub mati()
Unload Me
End Sub
