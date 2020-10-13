VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1575
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2778
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub SetLV()
With ListView1
    .View = lvwReport
    .Gridlines = True
    .MultiSelect = True
    .FullRowSelect = True
    .HotTracking = True
    .HoverSelection = True
    
    .ColumnHeaders.Add 1, , "No", 750
    .ColumnHeaders.Add 2, , "Kode Lapangan", 1250
    .ColumnHeaders.Add 3, , "Nama Lapangan", 2500
    Width = 10000
End With
End Sub
Public Sub TplGrid()
    Dim lst As ListItem, nmr As Integer
    With rsuser
    ListView1.ListItems.Clear
    Do While Not rsuser.EOF
    Set lst = ListView1.ListItems.Add
    nmr = nmr + 1
    lst.Text = nmr
    lst.SubItems(1) = rsuser!Kodelapangan
    lst.SubItems(2) = rsuser!NamaLapangan
    rsuser.MoveNext
    Loop
    End With
End Sub
Private Sub Form_Load()
Call SetLV
  Call BukaDB
rsuser.Open "Select * from MstLapangan Order By KodeLapangan", Koneksi
Call TplGrid
Set rsuser = Nothing
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Text3.Text = ListView1.SelectedItem.Text
Text1.Text = ListView1.SelectedItem.SubItems(1)
Text2.Text = ListView1.SelectedItem.SubItems(2)
End Sub
