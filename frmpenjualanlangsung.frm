VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.OCX"
Begin VB.Form frmpenjualanlangsung 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18585
   Icon            =   "frmpenjualanlangsung.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10440
   ScaleWidth      =   18585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10440
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18585
      _ExtentX        =   32782
      _ExtentY        =   18415
      _Version        =   262144
      AutoSize        =   1
      Locked          =   -1  'True
      PaneTree        =   "frmpenjualanlangsung.frx":9E4A
      Begin Threed.SSPanel SSPanel3 
         Height          =   2445
         Left            =   30
         TabIndex        =   3
         Top             =   7965
         Width           =   18525
         _ExtentX        =   32676
         _ExtentY        =   4313
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CommandButton Command4 
            Caption         =   "Simpan"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   3480
            Picture         =   "frmpenjualanlangsung.frx":9EBC
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Batal"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   1920
            Picture         =   "frmpenjualanlangsung.frx":BB86
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Baru"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   360
            Picture         =   "frmpenjualanlangsung.frx":D850
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   120
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   4020
         Left            =   30
         TabIndex        =   2
         Top             =   3855
         Width           =   18525
         _ExtentX        =   32676
         _ExtentY        =   7091
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComctlLib.ListView ListView1 
            Height          =   3975
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   20295
            _ExtentX        =   35798
            _ExtentY        =   7011
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   0
            Top             =   0
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   3735
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   18525
         _ExtentX        =   32676
         _ExtentY        =   6588
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox Text8 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3960
            TabIndex        =   26
            Top             =   1920
            Width           =   3255
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   495
            Left            =   3960
            TabIndex        =   24
            Top             =   480
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   112984065
            CurrentDate     =   41815
         End
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   11280
            TabIndex        =   19
            Top             =   3120
            Width           =   3255
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Tambah"
            BeginProperty Font 
               Name            =   "Puma Gaffer by Barreto"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   14880
            Picture         =   "frmpenjualanlangsung.frx":F51A
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   10800
            TabIndex        =   15
            Top             =   2280
            Width           =   3255
         End
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   10800
            TabIndex        =   13
            Top             =   1680
            Width           =   3255
         End
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   10800
            TabIndex        =   11
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   10800
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   480
            Width           =   3255
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3960
            TabIndex        =   7
            Top             =   2640
            Width           =   3255
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3960
            TabIndex        =   5
            Top             =   1200
            Width           =   3255
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "KodeJual                    :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   720
            TabIndex        =   25
            Top             =   1920
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal                       :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   720
            TabIndex        =   23
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Grand Total :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   8880
            TabIndex        =   18
            Top             =   3120
            Width           =   2295
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Total      :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8760
            TabIndex        =   14
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8760
            TabIndex        =   12
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga    :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8760
            TabIndex        =   10
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Barang   :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8760
            TabIndex        =   8
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Pakai Lapangan :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   6
            Top             =   2640
            Width           =   2775
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Urut                       :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   720
            TabIndex        =   4
            Top             =   1200
            Width           =   2655
         End
      End
   End
End
Attribute VB_Name = "frmpenjualanlangsung"
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

    .ColumnHeaders.Add 1, , "No Urut", 1000
    .ColumnHeaders.Add 2, , "KodeJual", 2000
    .ColumnHeaders.Add 3, , "Kode Pakai Lapangan", 2100
    .ColumnHeaders.Add 4, , "Kode Barang", 1500
    .ColumnHeaders.Add 5, , "Jumlah", 1000
    .ColumnHeaders.Add 6, , "Harga", 1500
    .ColumnHeaders.Add 7, , "Total", 1500
    .Width = 10700
End With
End Sub

Private Sub Command1_Click()
hidup
kosong
nurut
autokojul
Command1.Enabled = False
Command2.Enabled = True
Command4.Enabled = True
End Sub
Private Sub autokojul()
Call BukaDB
Dim Kode As String

Set rsuser = New ADODB.Recordset
    rsuser.Open "Select * From ItemPenjualan Where KodePenjualan Like '%" & Format(Date, "ddMMyy") & "%' ORDER BY KodePenjualan desc", Koneksi
    rsuser.Requery
    
    
If rsuser.EOF Then
Text8.Text = "JUL" & Format(Date, "ddMMyy") & "0001"
Else
Kode = rsuser!KodePenjualan
        Kode = Val(Right(Kode, 4))
        Kode = Kode + 1
    End If

'If rsuser.BOF Then
'        Text8.Text = "JUL" & Format(Date, "ddMMyy") & "0001"
'        Exit Sub
'    Else
'        rsuser.Requery
'
'        If (rsuser.EOF Or rsuser.BOF) Then
'            rsuser.MoveLast
'        End If
'        Kode = rsuser!KodePenjualan
'        Kode = Val(Right(Kode, 4))
'        Kode = Kode + 1
'    End If
    
    If Val(Kode) < 10 Then
 Kode = "JUL" & Format(Date, "ddMMyy") & "000" & Kode
        Text8.Text = Kode
    ElseIf Val(Kode) < 100 Then
        Kode = "JUL" & Format(Date, "ddMMyy") & "00" & Kode
        Text8.Text = Kode
    ElseIf Val(Kode) < 1000 Then
        Kode = "JUL" & Format(Date, "ddMMyy") & "0" & Kode
        Text8.Text = Kode
    ElseIf Val(Kode) < 10000 Then
        Kode = "JUL" & Format(Date, "ddMMyy") & "" & Kode
        Text8.Text = Kode
    Else
        MsgBox "Kapasitas Tidak Memadai!", _
        vbInformation + vbOKOnly, "Perhatian"
        Kode = ""
    End If
End Sub

Sub kosong()
Text1 = ""
Text2 = "0"
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
DTPicker1 = Now
End Sub

Private Sub Command2_Click()
kosong
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
ListView1.ListItems.Clear
mati
End Sub

Private Sub Command3_Click()
Set rsuser = New ADODB.Recordset
rsuser.Open "select * from MstBarang where KodeBarang = '" & Text3.Text & "'", Koneksi
If Val(Text4.Text) > rsuser.Fields("Stok") Then
MsgBox "Maaf Stock Barang Tidak Cukup", vbInformation, "Informasi"
Else
If Not rsuser.EOF Then
Dim a As Integer
Dim lst As ListItem
    Set lst = ListView1.ListItems.Add()
    lst.Text = Text1
    lst.SubItems(1) = Text8
    lst.SubItems(2) = Text2
    lst.SubItems(3) = Text3
    lst.SubItems(4) = Text4
    lst.SubItems(5) = Text5
    lst.SubItems(6) = Text6
    Text1 = ""
    Text2 = "-"
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    
a = MsgBox("Apakah Anda Akan Menambah Barang Lagi ?", vbQuestion + vbYesNo, "Konfirmasi")
If a = vbYes Then
Text1 = lst.Text + 1
Text3.SetFocus
ElseIf a = vbNo Then
Command3.Enabled = False
Command1.Enabled = False
Command4.Enabled = True
Command4.SetFocus
End If
Dim i, tot
For i = 1 To ListView1.ListItems.Count
tot = Val(tot) + Val(ListView1.ListItems(i).SubItems(6))
Next
Text7.Text = tot
End If
End If
End Sub

Private Sub Command4_Click()
If Text7 = "" Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
Else
frmtotallangsung.Show vbModal, Me
End If
End Sub

Private Sub Command6_Click()

End Sub

Private Sub Form_Load()
Call BukaDB
Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "select * from MstBarang"
Adodc1.Refresh
DTPicker1.Value = Format(Now, "mm/DD/yyyy")
mati
Call SetLV
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
End Sub

Private Sub nurut()
Call BukaDB
rsuser.Open ("SELECT * FROM ItemPenjualan WHERE NoUrut in(select max(NoUrut) from ItemPenjualan)order by NoUrut desc"), Koneksi
rsuser.Requery
    Dim Urut As String * 4
    Dim Hitung As Long
    With rsuser
        If .EOF Then
            Urut = "1"
            Text1 = Urut
        Else
            Hitung = Right(!NoUrut, 4) + 1
            Urut = Right("0" & Hitung, 4)
        End If
        Text1 = Urut
    End With
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
frmlihatbaranglagi.Show vbModal, Me

End Sub

Private Sub Text4_Change()
If Not IsNumeric(Text4) Then Text4 = "0"
Text6.Text = Val(Text5) * Val(Text4)


If Text4 = "0" Or Text4 = "" Then
Command3.Enabled = False
Else
Command3.Enabled = True
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
Dim stk
Call BukaDB
If KeyAscii = 13 Then
Adodc1.RecordSource = "select * from MstBarang where KodeBarang ='" & Text3.Text & "'"
Adodc1.Refresh
With Adodc1.Recordset
If .RecordCount > 0 Then
stk = Adodc1.Recordset!stok - (Val(Text4))
If stk < 0 Then
MsgBox "Stok Tidak memenuhi Permintaan", vbCritical + vbOKOnly, "Peringatan"
Text4.Text = ""
End If
End If
End With
End If
End Sub

Private Sub Text5_Change()
If Not IsNumeric(Text5) Then Text5 = "0"
End Sub

Private Sub Text7_Change()
'Me.Text7.Text = Format(CDbl(Text7.Text), "#,#")
End Sub
Sub mati()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
DTPicker1.Enabled = False
End Sub

Sub hidup()
DTPicker1.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
End Sub

