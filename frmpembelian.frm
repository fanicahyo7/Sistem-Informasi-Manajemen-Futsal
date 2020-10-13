VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.OCX"
Begin VB.Form frmpembelian 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19260
   Icon            =   "frmpembelian.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10095
   ScaleWidth      =   19260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Height          =   1095
      Left            =   1920
      Picture         =   "frmpembelian.frx":9E4A
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Baru"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      Picture         =   "frmpembelian.frx":BB14
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8160
      Width           =   1215
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19260
      _ExtentX        =   33973
      _ExtentY        =   17806
      _Version        =   262144
      AutoSize        =   1
      Locked          =   -1  'True
      PaneTree        =   "frmpembelian.frx":D7DE
      Begin Threed.SSPanel SSPanel3 
         Height          =   2115
         Left            =   30
         TabIndex        =   3
         Top             =   7950
         Width           =   19200
         _ExtentX        =   33867
         _ExtentY        =   3731
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CommandButton Command3 
            Caption         =   "Simpan"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   3360
            Picture         =   "frmpembelian.frx":D850
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   240
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   3585
         Left            =   30
         TabIndex        =   2
         Top             =   4275
         Width           =   19200
         _ExtentX        =   33867
         _ExtentY        =   6324
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComctlLib.ListView ListView1 
            Height          =   3495
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   13575
            _ExtentX        =   23945
            _ExtentY        =   6165
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
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   4155
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   19200
         _ExtentX        =   33867
         _ExtentY        =   7329
         _Version        =   262144
         BackColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Puma Gaffer by Barreto"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "Puma Gaffer by Barreto"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   15960
            TabIndex        =   34
            Top             =   4200
            Width           =   4215
         End
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   330
            Left            =   8400
            Top             =   5400
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
            Caption         =   "Adodc2"
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
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   5280
            Top             =   5400
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
         Begin VB.CommandButton Command4 
            Caption         =   "Tambah"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   14760
            Picture         =   "frmpembelian.frx":F51A
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox Text10 
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
            Left            =   14280
            TabIndex        =   28
            Top             =   1800
            Width           =   3135
         End
         Begin VB.TextBox Text9 
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
            Left            =   14280
            TabIndex        =   27
            Top             =   1080
            Width           =   3135
         End
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
            Left            =   9480
            TabIndex        =   26
            Top             =   3120
            Width           =   3135
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            ItemData        =   "frmpembelian.frx":111E4
            Left            =   9480
            List            =   "frmpembelian.frx":111E6
            TabIndex        =   25
            Text            =   "Pilih Kode Barang"
            Top             =   1800
            Width           =   3135
         End
         Begin VB.TextBox Text7 
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
            Left            =   9480
            TabIndex        =   24
            Top             =   2400
            Width           =   3135
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
            Height          =   495
            Left            =   9480
            TabIndex        =   23
            Top             =   1080
            Width           =   3135
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   495
            Left            =   9480
            TabIndex        =   22
            Top             =   480
            Width           =   3135
            _ExtentX        =   5530
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
            Format          =   124256257
            CurrentDate     =   41659
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
            Left            =   3480
            TabIndex        =   21
            Top             =   3120
            Width           =   3135
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
            Height          =   495
            Left            =   3480
            TabIndex        =   20
            Top             =   2400
            Width           =   3135
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
            Left            =   3480
            TabIndex        =   19
            Top             =   1680
            Width           =   3135
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            ItemData        =   "frmpembelian.frx":111E8
            Left            =   3480
            List            =   "frmpembelian.frx":111EA
            TabIndex        =   18
            Text            =   "Pilih Kode Supplier"
            Top             =   1080
            Width           =   3135
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
            Left            =   3480
            TabIndex        =   17
            Top             =   480
            Width           =   3135
         End
         Begin VB.Label Label14 
            Caption         =   "Total Pembelian :"
            BeginProperty Font 
               Name            =   "Puma Gaffer by Barreto"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   13080
            TabIndex        =   16
            Top             =   4200
            Width           =   2655
         End
         Begin VB.Label Label13 
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
            Left            =   13080
            TabIndex        =   15
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Barang :"
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
            Left            =   7680
            TabIndex        =   14
            Top             =   2400
            Width           =   2175
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga              :"
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
            Left            =   7680
            TabIndex        =   13
            Top             =   3120
            Width           =   1815
         End
         Begin VB.Label Label10 
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
            Left            =   13080
            TabIndex        =   12
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Barang :"
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
            Left            =   7680
            TabIndex        =   11
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Urut          :"
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
            Left            =   7680
            TabIndex        =   10
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "HP Penanggung Jawab :"
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
            Left            =   360
            TabIndex        =   9
            Top             =   3120
            Width           =   3135
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Telp                                    :"
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
            Left            =   360
            TabIndex        =   8
            Top             =   2400
            Width           =   3015
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Supplier                 :"
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
            Left            =   360
            TabIndex        =   7
            Top             =   1680
            Width           =   3015
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal          :"
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
            Left            =   7680
            TabIndex        =   6
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Supplier                  :"
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
            Left            =   360
            TabIndex        =   5
            Top             =   1080
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Pembelian              :"
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
            Left            =   360
            TabIndex        =   4
            Top             =   480
            Width           =   3015
         End
      End
   End
End
Attribute VB_Name = "frmpembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub autonumber()
Call BukaDB
rsuser.Open ("SELECT * FROM TrPembelian WHERE KodePembelian in(select max(KodePembelian) from TrPembelian)order by KodePembelian desc"), Koneksi
rsuser.Requery
    Dim Urut As String * 6
    Dim Hitung As Long
    With rsuser
        If .EOF Then
            Urut = "PBL" + "001"
            Text1 = Urut
        Else
            Hitung = (Right(!KodePembelian, 3)) + 1
            Urut = "PBL" + Right("00" & Hitung, 3)
        End If
        Text1 = Urut
    End With
End Sub
Private Sub nurut()
Call BukaDB
rsuser.Open ("SELECT * FROM ItemPembelian WHERE NoUrut in(select max(NoUrut) from ItemPembelian)order by NoUrut desc"), Koneksi
rsuser.Requery
     Dim Urut1 As String * 4
    Dim Hitung1 As Long
    With rsuser
        If .EOF Then
            Urut1 = "01"
            Text5 = Urut1
        Else
            Hitung1 = Right(!NoUrut, 4) + 1
            Urut1 = Right("0" & Hitung1, 4)
        End If
        Text5 = Urut1
    End With

End Sub

Private Sub Combo1_Click()
Adodc1.RecordSource = "select * from MstSupplier where KodeSupplier='" & Combo1 & "'"
Adodc1.Refresh
 Text2 = Adodc1.Recordset!NamaSupplier
 Text3 = Adodc1.Recordset!telp
 Text4 = Adodc1.Recordset!HPPenanggungJawab
End Sub

Private Sub Combo2_Click()
Adodc2.RecordSource = "select * from MstBarang where KodeBarang='" & Combo2 & "'"
Adodc2.Refresh
Text7 = Adodc2.Recordset!NamaBarang
Text8.SetFocus
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
Text9.Enabled = False
Text10.Enabled = False
DTPicker1.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
End Sub
Sub hidup()
Combo1.Enabled = True
Combo2.Enabled = True
DTPicker1.Enabled = True
Text9.Enabled = True
Text8.Enabled = True
End Sub

Private Sub Command1_Click()
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = True
hidup
kosong
autonumber
nurut
End Sub

Private Sub Command2_Click()
mati
kosong
ListView1.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
ListView1.ListItems.Clear
End Sub

Private Sub Command3_Click()
If Combo1.Text = "Pilih Kode Supplier" Or Combo2.Text = "Pilih Kode Barang" Then
MsgBox "Data Belum Lengkap", vbCritical + vbInformation, "Peringatan"
Else
frmhitungpembelian.Show vbModal, Me
End If
End Sub

Private Sub Command4_Click()
If Text8 = "" Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
Else
Dim a As Integer
Dim lst As ListItem
    Set lst = ListView1.ListItems.Add()
lst.Text = Text5
lst.SubItems(1) = Text1
lst.SubItems(2) = Combo2
lst.SubItems(3) = Text9
lst.SubItems(4) = Text8
lst.SubItems(5) = Text10
lst.SubItems(6) = Format(DTPicker1, "yyyy/MM/DD")
Combo2 = ""
Text9 = ""
Text8 = ""
Text10 = ""
Text7 = ""
DTPicker1 = Now

a = MsgBox("Apakah Anda Akan Menambah Barang Lagi ?", vbQuestion + vbYesNo)
If a = vbYes Then
Text5 = lst.Text + 1
ElseIf a = vbNo Then
Command3.Enabled = True
Command2.Enabled = True
Command1.Enabled = False
Command4.Enabled = False
End If
Dim i, tot
For i = 1 To ListView1.ListItems.Count
tot = Val(tot) + Val(ListView1.ListItems(i).SubItems(5))
Next
Text6.Text = tot
End If
End Sub

Private Sub Form_Load()
Call BukaDB
Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "select * from MstSupplier"
Adodc1.Refresh

Adodc2.ConnectionString = Koneksi
Adodc2.RecordSource = "select * from MstBarang"
Adodc2.Refresh

Call SetLV
mati
kosong
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
ListView1.ListItems.Clear

Adodc1.RecordSource = "select * from MstSupplier"
Adodc1.Refresh
        Do While Not Adodc1.Recordset.EOF
            Combo1.AddItem Adodc1.Recordset!KodeSupplier
            Adodc1.Recordset.MoveNext
        Loop
        
Adodc2.RecordSource = "select * from MstBarang order by KodeBarang"
Adodc2.Refresh
    Do While Not Adodc2.Recordset.EOF
    Combo2.AddItem Adodc2.Recordset!KodeBarang
    Adodc2.Recordset.MoveNext
    Loop
End Sub
Sub kosong()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Text10 = ""
DTPicker1.Value = Now
End Sub
Public Sub SetLV()
With ListView1
    .View = lvwReport
    .Gridlines = True
    .MultiSelect = True
    .FullRowSelect = True
    .HotTracking = True
    .HoverSelection = True

    .ColumnHeaders.Add 1, , "No Urut", 1000
    .ColumnHeaders.Add 2, , "Kode Pembelian", 2000
    .ColumnHeaders.Add 3, , "Kode Barang", 1500
    .ColumnHeaders.Add 4, , "Jumlah", 1000
    .ColumnHeaders.Add 5, , "Harga", 1500
    .ColumnHeaders.Add 6, , "Total", 1500
    .ColumnHeaders.Add 7, , "Tanggal", 1500
    .Width = 10100
End With
End Sub
 
Private Sub Text8_Change()
'If Not IsNumeric(Text8) Then Text8 = "0"
'Text10 = Val(Text8) * (Text9)
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text9.SetFocus
End If
End Sub

Private Sub Text9_Change()
If Not IsNumeric(Text9) Then Text9 = "0"
Text10 = Val(Text8) * Val(Text9)

If Text9 = "0" Or Text9 = "" Then
Command4.Enabled = False
Else
Command4.Enabled = True
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command4.SetFocus
End If
End Sub
