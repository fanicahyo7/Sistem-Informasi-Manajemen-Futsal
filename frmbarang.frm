VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frmbarang 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10395
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20490
   Icon            =   "frmbarang.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10395
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10395
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20490
      _ExtentX        =   36142
      _ExtentY        =   18336
      _Version        =   262144
      AutoSize        =   1
      Locked          =   -1  'True
      PaneTree        =   "frmbarang.frx":9E4A
      Begin Threed.SSPanel SSPanel3 
         Height          =   2040
         Left            =   30
         TabIndex        =   3
         Top             =   8325
         Width           =   20430
         _ExtentX        =   36036
         _ExtentY        =   3598
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CommandButton Command6 
            Caption         =   "Eksport Ke Excel"
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
            Left            =   5880
            Picture         =   "frmbarang.frx":9EBC
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   120
            Width           =   1335
         End
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
            Height          =   1095
            Left            =   4440
            Picture         =   "frmbarang.frx":BB86
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Ubah"
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
            Left            =   3120
            Picture         =   "frmbarang.frx":D850
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   120
            Width           =   1215
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
            Height          =   1095
            Left            =   1800
            Picture         =   "frmbarang.frx":F51A
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   120
            Width           =   1215
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
            Height          =   1095
            Left            =   360
            Picture         =   "frmbarang.frx":111E4
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   120
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   3510
         Left            =   30
         TabIndex        =   2
         Top             =   4725
         Width           =   20430
         _ExtentX        =   36036
         _ExtentY        =   6191
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmbarang.frx":12EAE
            Height          =   3495
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   18135
            _ExtentX        =   31988
            _ExtentY        =   6165
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   24
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   14.25
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
            Caption         =   "Barang"
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
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   4605
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   20430
         _ExtentX        =   36036
         _ExtentY        =   8123
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   330
            Left            =   6960
            Top             =   4800
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
         Begin VB.CommandButton Command5 
            Caption         =   "Browse"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11520
            TabIndex        =   27
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox Text9 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8640
            TabIndex        =   25
            Top             =   2160
            Width           =   2775
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H8000000D&
            Height          =   3975
            Left            =   12360
            ScaleHeight     =   3915
            ScaleWidth      =   4875
            TabIndex        =   24
            Top             =   240
            Width           =   4935
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   1800
               Top             =   1320
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H8000000D&
            Caption         =   "Cari Barang"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   960
            TabIndex        =   15
            Top             =   2880
            Width           =   5175
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
               Height          =   405
               Left            =   2640
               TabIndex        =   17
               Top             =   960
               Width           =   2175
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
               Height          =   405
               Left            =   2640
               TabIndex        =   16
               Top             =   360
               Width           =   2175
            End
            Begin VB.Label Label10 
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
               Height          =   375
               Left            =   720
               TabIndex        =   19
               Top             =   960
               Width           =   1815
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Kode Barang  :"
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
               TabIndex        =   18
               Top             =   360
               Width           =   1695
            End
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
            Height          =   405
            Left            =   2880
            TabIndex        =   14
            Top             =   2040
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
            Left            =   8640
            TabIndex        =   12
            Top             =   1320
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
            Left            =   8640
            TabIndex        =   10
            Top             =   600
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
            Left            =   2880
            TabIndex        =   8
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
            Left            =   2880
            TabIndex        =   6
            Top             =   600
            Width           =   2415
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   375
            Left            =   8280
            Top             =   5640
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
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
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Lokasi Gambar :"
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
            Left            =   6600
            TabIndex        =   26
            Top             =   2160
            Width           =   1935
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Stok                   :"
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
            Left            =   840
            TabIndex        =   13
            Top             =   2040
            Width           =   1935
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Beli    :"
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
            Left            =   6600
            TabIndex        =   11
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Jual   :"
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
            Left            =   6600
            TabIndex        =   9
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Barang  :"
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
            Left            =   840
            TabIndex        =   7
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label1 
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
            Left            =   840
            TabIndex        =   5
            Top             =   600
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "frmbarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsMySQL As MYSQL_RS
Dim rsData As ADODB.Recordset
Dim sSQL As String
'Dim excel As New excel.Application
Public Sub dtgrid()
Set rsData = New ADODB.Recordset
    rsData.ActiveConnection = Koneksi
    rsData.CursorLocation = adUseClient
    rsData.CursorType = adOpenDynamic
    rsData.LockType = adLockOptimistic
    rsData.Source = "SELECT * FROM MstBarang"
    rsData.Open
End Sub

Sub mati()
Text1.Enabled = False
Text2.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text9.Enabled = False
End Sub
Sub hidup()
Text2.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
End Sub

Private Sub Command1_Click()
If Command1.Caption = "&Baru" Then
Command1.Caption = "&Simpan"
hidup
Text5.Enabled = False
Text6.Enabled = False
kosong
Picture1.Picture = LoadPicture("")
autonumber
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = True
DataGrid1.Enabled = False
Else
If Text1 = "" Or Text2 = "" Or Text4 = "" Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
Else
Dim cari As String
cari = "NamaBarang='" & Text2 & "'"
With Adodc1.Recordset
.Find cari
If Not .EOF Then
MsgBox "Data Sudah Ada", vbCritical + vbOKOnly, "Peringatan"
Else
'Dim rsanu As New ADODB.Recordset
'rsanu.Open "insert into MstBarang (KodeBarang,NamaBarang,HargaJual,Lokasi,Foto) values ('" & Text1 & "','" & Text2 & "','" & Text4 & "','" & Text9 & "','" & Picture1 & "')", Koneksi
Adodc2.Recordset.AddNew
Adodc2.Recordset!KodeBarang = Text1
Adodc2.Recordset!NamaBarang = Text2
Adodc2.Recordset!HargaJual = Text4
Adodc2.Recordset!Lokasi = Text9
Adodc2.Recordset!Foto = Picture1
Adodc2.Recordset.Update
Adodc2.Recordset.Requery
MsgBox "Data Berhasil Disimpan", vbInformation + vbOKOnly, "Informasi"
kosong
Form_Load
End If
End With
End If
End If
End Sub
Sub kosong()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text9 = ""
End Sub

Private Sub Command2_Click()
Form_Load
kosong
Picture1.Picture = LoadPicture("")
End Sub

Private Sub Command3_Click()
If Text1 = "" Then
MsgBox "Data Belum Dipilih", vbCritical + vbOKOnly, "Peringatan"
Else
If Command3.Caption = "&Ubah" Then
Command3.Caption = "&Simpan"
hidup
Command1.Enabled = False
Command2.Enabled = True
Command4.Enabled = False
Command5.Enabled = True
DataGrid1.Enabled = False
Else
If Text1 = "" Or Text2 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
Else
Adodc2.RecordSource = "select * from Mstbarang where KodeBarang='" & Text1 & "'"
Adodc2.Refresh
'Dim sqlupdate As String
'sqlupdate = "update MstBarang set NamaBarang='" & Text2 & "',HargaJual='" & Text4 & "',HargaBeli='" & Text5 & "',Stok='" & Text6 & "' where KodeBarang='" & Text1 & "'"
'Koneksi.Execute sqlupdate
Adodc2.Recordset!KodeBarang = Text1
Adodc2.Recordset!NamaBarang = Text2
Adodc2.Recordset!HargaJual = Text4
Adodc2.Recordset!HargaBeli = Text5
Adodc2.Recordset!stok = Text6
Adodc2.Recordset!Lokasi = Text9
Adodc2.Recordset!Foto = Picture1
Adodc2.Recordset.Update
Adodc2.Recordset.Requery
MsgBox "Data Berhasil Diubah", vbInformation + vbOKOnly, "Informasi"
kosong
Form_Load
End If
End If
End If
End Sub
Private Sub autonumber()
Call BukaDB
rsuser.Open ("SELECT * FROM MstBarang WHERE KodeBarang in(select max(KodeBarang) from MstBarang)order by KodeBarang desc"), Koneksi
rsuser.Requery
    Dim Urut As String * 6
    Dim Hitung As Long
    With rsuser
        If .EOF Then
            Urut = "BRG" + "001"
            Text1 = Urut
        Else
            Hitung = Right(!KodeBarang, 3) + 1
            Urut = "BRG" + Right("00" & Hitung, 3)
        End If
        Text1 = Urut
    End With

End Sub

Private Sub Command4_Click()
If Not Adodc1.Recordset.EOF = False Then
MsgBox "Data Kosong", vbCritical + vbOKOnly, "Peringatan"
Else
If MsgBox("Apakah Anda Yakin Akan Menghapus ?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
Dim sqldelete As String
sqldelete = "delete from MstBarang where KodeBarang='" & Text1 & "'"
Koneksi.Execute sqldelete
MsgBox "Data Berhasil Dihapus", vbInformation + vbOKOnly, "Informasi"
Form_Load
kosong
End If
End If
End Sub

Private Sub Command5_Click()
CommonDialog1.DialogTitle = "Pilih Gambar"
CommonDialog1.Filter = "File JPEG|*.jpg"
CommonDialog1.ShowOpen
Text9 = CommonDialog1.FileName
End Sub

Private Sub Command6_Click()
anuexcel
End Sub
Sub anuexcel()
    Set excel = excel.Application
    excel.Workbooks.Add
    excel.Worksheets(1).Activate
    

    For i = 0 To DataEnvironment1.rsRBarang.Fields.Count - 1
        excel.Worksheets(1).Cells(1) = "Data Barang"
        excel.Worksheets(1).Cells(2, i + 1) = DataEnvironment1.rsRBarang.Fields(i).Name
    Next
    

    If DataEnvironment1.rsRBarang.State = 0 Then DataEnvironment1.rsRBarang.Open
    If DataEnvironment1.rsRBarang.RecordCount > 0 Then DataEnvironment1.rsRBarang.MoveFirst
        For i = 1 To DataEnvironment1.rsRBarang.RecordCount
            For j = 0 To DataEnvironment1.rsRBarang.Fields.Count - 1
                excel.Worksheets(1).Cells(i + 2, j + 1) = DataEnvironment1.rsRBarang(j)
            Next
            DataEnvironment1.rsRBarang.MoveNext
        Next
        
    excel.Columns.AutoFit
    excel.Visible = True
    excel.Workbooks(1).Saved = False
End Sub
Private Sub DataGrid1_Click()
If Not Adodc1.Recordset.EOF = False Then
MsgBox "Data Masih Kosong", vbInformation + vbOKOnly, "Informasi"
Else
Text1 = DataGrid1.Columns(0)
Text2 = DataGrid1.Columns(1)
Text4 = DataGrid1.Columns(2)
Text5 = DataGrid1.Columns(3)
Text6 = DataGrid1.Columns(4)
Dim klo As New ADODB.Recordset
klo.Open "select * from MstBarang where KodeBarang='" & Text1 & "'", Koneksi
klo.Requery
Text9 = klo!Lokasi
End If
End Sub

Private Sub Form_Load()
Call BukaDB
Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "select KodeBarang,NamaBarang,HargaJual,HargaBeli,Stok from MstBarang order by KodeBarang"
Adodc1.Refresh
Adodc2.ConnectionString = Koneksi
Adodc2.RecordSource = "select * from MstBarang"
Adodc2.Refresh
'Call dtgrid
'Set DataGrid1.DataSource = rsData
'With DataGrid1
'End With
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = False
Command1.Caption = "&Baru"
Command3.Caption = "&Ubah"
DataGrid1.Enabled = True
Command6.Visible = False
mati
'Dim jumlah As Long
'DataGrid1.Refresh
'    Adodc1.Recordset.MoveFirst
'    Do Until Adodc1.Recordset.EOF
'        jumlah = jumlah + Adodc1.Recordset.Fields(4)
'        Adodc1.Recordset.MoveNext
'    Loop
'    Text10.Text = "Rp. = " & jumlah
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.SetFocus
End If
End Sub

Private Sub Text5_Change()
If Not IsNumeric(Text5) Then Text5 = "0"
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text6.SetFocus
End If
End Sub

Private Sub Text6_Change()
If Not IsNumeric(Text6) Then Text6 = "0"
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command5.SetFocus
End If
End Sub

Private Sub Text9_Change()
Picture1.Picture = LoadPicture(Text9)
End Sub
Private Sub Text7_Change()
Adodc1.RecordSource = "select KodeBarang,NamaBarang,HargaJual,HargaBeli,Stok from MstBarang where KodeBarang like '%" & TandaPetik(Text7) & "%'"
Adodc1.Refresh
If Not Adodc1.Recordset.EOF Then
    With DataGrid1
    Set .DataSource = Adodc1
        .Refresh
    End With
End If
End Sub

Private Sub Text8_Change()
Adodc1.RecordSource = "select KodeBarang,NamaBarang,HargaJual,HargaBeli,Stok from MstBarang where NamaBarang like '%" & TandaPetik(Text8) & "%'"
Adodc1.Refresh
If Not Adodc1.Recordset.EOF Then
    With DataGrid1
    Set .DataSource = Adodc1
        .Refresh
    End With
End If
End Sub
