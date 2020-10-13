VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.OCX"
Begin VB.Form frmrevisi 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20490
   Icon            =   "frmrevisi.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11520
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20490
      _ExtentX        =   36142
      _ExtentY        =   20320
      _Version        =   262144
      AutoSize        =   1
      Locked          =   -1  'True
      PaneTree        =   "frmrevisi.frx":9E4A
      Begin Threed.SSPanel SSPanel3 
         Height          =   3645
         Left            =   30
         TabIndex        =   3
         Top             =   7845
         Width           =   20430
         _ExtentX        =   36036
         _ExtentY        =   6429
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CommandButton Command4 
            Caption         =   "Hapus"
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
            Left            =   3720
            Picture         =   "frmrevisi.frx":9EBC
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
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
            Left            =   2160
            Picture         =   "frmrevisi.frx":BB86
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Baru"
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
            Left            =   480
            Picture         =   "frmrevisi.frx":D850
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   120
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   3810
         Left            =   30
         TabIndex        =   2
         Top             =   3945
         Width           =   20430
         _ExtentX        =   36036
         _ExtentY        =   6720
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmrevisi.frx":F51A
            Height          =   3735
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   18135
            _ExtentX        =   31988
            _ExtentY        =   6588
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   24
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Revisi Stok"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
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
         Height          =   3825
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   20430
         _ExtentX        =   36036
         _ExtentY        =   6747
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
            Height          =   1455
            Left            =   12720
            TabIndex        =   25
            Top             =   720
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
            Left            =   7920
            TabIndex        =   23
            Top             =   1920
            Width           =   2415
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
            Left            =   7920
            TabIndex        =   21
            Top             =   1320
            Width           =   2415
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
            Height          =   405
            Left            =   7920
            TabIndex        =   15
            Top             =   2640
            Width           =   2415
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
            Height          =   405
            Left            =   7920
            TabIndex        =   13
            Top             =   720
            Width           =   2415
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
            Height          =   405
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   1920
            Width           =   2415
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
            Height          =   405
            Left            =   2760
            TabIndex        =   7
            Top             =   1320
            Width           =   2415
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
            Height          =   405
            Left            =   2760
            TabIndex        =   5
            Top             =   720
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   495
            Left            =   2760
            TabIndex        =   11
            Top             =   2520
            Width           =   2415
            _ExtentX        =   4260
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
            Format          =   123600897
            CurrentDate     =   41646
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Keterangan :"
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
            Left            =   11040
            TabIndex        =   24
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Bertambah   :"
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
            Left            =   6120
            TabIndex        =   22
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Berkurang    :"
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
            Left            =   6120
            TabIndex        =   20
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Stok Baru     :"
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
            Left            =   6120
            TabIndex        =   14
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Stok Lama    :"
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
            Left            =   6120
            TabIndex        =   12
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal                :"
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
            Left            =   480
            TabIndex        =   10
            Top             =   2520
            Width           =   2055
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Barang      :"
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
            Left            =   480
            TabIndex        =   8
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Transaksi :"
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
            Left            =   480
            TabIndex        =   6
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Urut               :"
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
            Left            =   480
            TabIndex        =   4
            Top             =   720
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "frmrevisi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub kosong()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
DTPicker1.Value = Now
End Sub

Private Sub Command3_Click()
Form_Load
End Sub

Private Sub Command4_Click()
If Text1 = "" Then
MsgBox "Data Belum Dipilih", vbCritical + vbOKOnly, "Peringatan"
Else
If MsgBox("Apakah Anda Yakin Menghapus Data Ini?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
Adodc1.Recordset.Delete
MsgBox "Data Berhasil Dihapus", vbInformation + vbOKOnly, "Informasi"
Form_Load
End If
End If
End Sub

Private Sub DataGrid1_Click()
If Not Adodc1.Recordset.EOF = False Then
MsgBox "Data Masih Kosong", vbInformation + vbOKOnly, "Informasi"
Else
Text1.Text = Adodc1.Recordset!NoUrut
Text8.Text = Adodc1.Recordset!Keterangan
Text2.Text = Adodc1.Recordset!KodeTrans
DTPicker1.Value = Adodc1.Recordset!Tanggal
Text3.Text = Adodc1.Recordset!KodeBarang
Text4.Text = Adodc1.Recordset!StokLama
Text5.Text = Adodc1.Recordset!StokBaru
End If
End Sub

Private Sub Form_Load()
Call BukaDB
Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "select * from trrevisistok"
Adodc1.Refresh

Command1.Enabled = True
Command3.Enabled = False
Command4.Enabled = True
Command1.Caption = "&Baru"
DataGrid1.Enabled = True
mati
kosong
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
frmlihatbarang.Show vbModal, Me
End Sub

Private Sub AutoKoTrans()
Call BukaDB
rsuser.Open ("SELECT * FROM TrRevisiStok WHERE KodeTrans in(select max(KodeTrans) from TrRevisiStok)order by KodeTrans desc"), Koneksi
rsuser.Requery
    Dim Urut As String * 5
    Dim Hitung As Long
    With rsuser
        If .EOF Then
            Urut = "TR" + "001"
            Text2 = Urut
        Else
            Hitung = Right(!KodeTrans, 3) + 1
            Urut = "TR" + Right("00" & Hitung, 3)
        End If
        Text2 = Urut
    End With
End Sub
Private Sub nurut()
Call BukaDB
rsuser.Open ("SELECT * FROM TrRevisiStok WHERE NoUrut in(select max(NoUrut) from TrRevisiStok)order by NoUrut desc"), Koneksi
rsuser.Requery
    Dim Urut As String * 4
    Dim Hitung As Long
    With rsuser
        If .EOF Then
            Urut = "01"
            Text1 = Urut
        Else
            Hitung = Right(!NoUrut, 4) + 1
            Urut = Right("0" & Hitung, 4)
        End If
        Text1 = Urut
    End With
End Sub

Private Sub Command1_Click()
If Command1.Caption = "&Baru" Then
Command1.Caption = "&Simpan"
Command3.Enabled = True
Command4.Enabled = False
hidup
kosong
DataGrid1.Enabled = False
Call AutoKoTrans
Call nurut
Else
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Then
MsgBox "Data BeLum Lengkap", vbInformation + vbOKOnly, "Pesan"
Else
'Dim rsitu As New ADODB.Recordset
'rsitu.Open "insert into TrRevisiStok (NoUrut,KodeTrans,Tanggal,KodeBarang,StokLama,StokBaru) values ('" & Text1 & "','" & Text2 & "','" & DTPicker1 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "')", Koneksi
Adodc1.Recordset.AddNew
Adodc1.Recordset!NoUrut = Text1.Text
Adodc1.Recordset!KodeTrans = Text2.Text
Adodc1.Recordset!Keterangan = Text8.Text
Adodc1.Recordset!Tanggal = Format(DTPicker1, "YYYY/MM/DD")
Adodc1.Recordset!KodeBarang = Text3.Text
Adodc1.Recordset!StokLama = Text4.Text
Adodc1.Recordset!StokBaru = Text5.Text
Adodc1.Recordset.Update
Adodc1.Recordset.Requery
stok
MsgBox "Data berhasil Disimpan", vbInformation + vbOKOnly, "Sukses"
Form_Load
End If
End If
End Sub
Sub stok()
Call BukaDB
Dim sqlstok As String
sqlstok = "update MstBarang set Stok ='" & Text5 & "' where KodeBarang ='" & Text3 & "'"
Koneksi.Execute sqlstok
End Sub

Private Sub Text4_Change()
If Not IsNumeric(Text4) Then Text4 = "0"
End Sub

Private Sub Text6_Change()
If Not IsNumeric(Text6) Then Text6 = "0"
Text5 = Val(Text4) - Val(Text6)
End Sub

Private Sub Text7_Change()
If Not IsNumeric(Text7) Then Text7 = "0"
Text5 = Val(Text4) + Val(Text7)
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
Text3.Enabled = True
Text4.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
DTPicker1.Enabled = True
End Sub
