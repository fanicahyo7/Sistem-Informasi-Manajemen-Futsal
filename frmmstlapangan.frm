VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.OCX"
Begin VB.Form frmmstlapangan 
   BorderStyle     =   0  'None
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20490
   Icon            =   "frmmstlapangan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11055
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20490
      _ExtentX        =   36142
      _ExtentY        =   19500
      _Version        =   262144
      AutoSize        =   1
      Locked          =   -1  'True
      PaneTree        =   "frmmstlapangan.frx":9E4A
      Begin Threed.SSPanel SSPanel3 
         Height          =   2775
         Left            =   30
         TabIndex        =   3
         Top             =   8250
         Width           =   20430
         _ExtentX        =   36036
         _ExtentY        =   4895
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
            Height          =   1095
            Left            =   4080
            Picture         =   "frmmstlapangan.frx":9EBC
            Style           =   1  'Graphical
            TabIndex        =   17
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
            Left            =   2880
            Picture         =   "frmmstlapangan.frx":BB86
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   120
            Width           =   1095
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
            Left            =   1680
            Picture         =   "frmmstlapangan.frx":D850
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   120
            Width           =   1095
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
            Picture         =   "frmmstlapangan.frx":F51A
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   120
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   4125
         Left            =   30
         TabIndex        =   2
         Top             =   4035
         Width           =   20430
         _ExtentX        =   36036
         _ExtentY        =   7276
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   375
            Left            =   5040
            Top             =   4680
            Width           =   1200
            _ExtentX        =   2117
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
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmmstlapangan.frx":111E4
            Height          =   4095
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   18135
            _ExtentX        =   31988
            _ExtentY        =   7223
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
            Caption         =   "Lapangan"
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
         Height          =   3915
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   20430
         _ExtentX        =   36036
         _ExtentY        =   6906
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   330
            Left            =   8040
            Top             =   4200
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
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   12600
            Top             =   1560
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Browse"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   7560
            TabIndex        =   20
            Top             =   1560
            Width           =   855
         End
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3240
            TabIndex        =   19
            Top             =   1560
            Width           =   4215
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H8000000D&
            Height          =   3495
            Left            =   10800
            ScaleHeight     =   3435
            ScaleWidth      =   5355
            TabIndex        =   18
            Top             =   120
            Width           =   5415
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H8000000D&
            Caption         =   "Cari Data Lapangan"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   840
            TabIndex        =   8
            Top             =   2160
            Width           =   6495
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
               Height          =   450
               Left            =   3000
               TabIndex        =   12
               Top             =   1080
               Width           =   2655
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
               Height          =   435
               Left            =   3000
               TabIndex        =   10
               Top             =   480
               Width           =   2655
            End
            Begin VB.Label Label4 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Nama Lapangan :"
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
               Left            =   600
               TabIndex        =   11
               Top             =   1080
               Width           =   2175
            End
            Begin VB.Label Label3 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Kode Lapangan  :"
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
               Left            =   600
               TabIndex        =   9
               Top             =   480
               Width           =   2175
            End
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
            Height          =   450
            Left            =   3240
            TabIndex        =   7
            Top             =   960
            Width           =   4215
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
            Height          =   450
            Left            =   3240
            TabIndex        =   5
            Top             =   360
            Width           =   4215
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Gambar                 :"
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
            Left            =   960
            TabIndex        =   21
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Lapangan :"
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
            Left            =   960
            TabIndex        =   6
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Lapangan  :"
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
            Left            =   960
            TabIndex        =   4
            Top             =   360
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "frmmstlapangan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim rsMySQL As MYSQL_RS
'Dim rsData As ADODB.Recordset
'Dim sSQL As String
'Public Sub dtgrid()
'Set rsData = New ADODB.Recordset
'    rsData.ActiveConnection = Koneksi
'    rsData.CursorLocation = adUseClient
'    rsData.CursorType = adOpenDynamic
'    rsData.LockType = adLockOptimistic
'    rsData.Source = "SELECT * FROM MstLapangan"
'    rsData.Open
'End Sub
'Sub grid()
'Dim l As Long
'Call BukaDB
'
'sSQL = "select * from Mstlapangan"
'Set rsMySQL = Koneksi.Execute(sSQL)
'
'Set rsData = New ADODB.Recordset
'rsData.Fields.Append rsMySQL.Fields(0).Name, adVarChar, 10, adFldIsNullable
'rsData.Fields.Append rsMySQL.Fields(1).Name, adVarChar, 50, adFldIsNullable
'rsData.Fields.Append rsMySQL.Fields(2).Name, adVarChar, 50, adFldIsNullable
'rsData.Open
'
'For l = 0 To rsMySQL.RecordCount - 1
'    rsData.AddNew
'    rsData.Fields(0).Value = rsMySQL.Fields(0).Value
'    rsData.Fields(1).Value = rsMySQL.Fields(1).Value
'    rsData.Fields(2).Value = rsMySQL.Fields(2).Value
'    rsData.Update
'    rsMySQL.MoveNext
'Next
'
'Set DataGrid1.DataSource = rsData
'
'With DataGrid1
'    .Columns(0).Caption = "Kode Lapangan"
'    .Columns(1).Caption = "Nama Kategori"
'    .Columns(2).Caption = "Lokasi"
'    .Columns(0).Width = 2000
'    .Columns(1).Width = 3000
'    .Columns(2).Width = 3000
'End With
'
'rsMySQL.CloseRecordset
'Set rsMySQL = Nothing
'End Sub

Private Sub autonumber()
Call BukaDB
rsuser.Open ("SELECT * FROM MstLapangan WHERE KodeLapangan in(select max(KodeLapangan) from MstLapangan)order by KodeLapangan desc"), Koneksi
rsuser.Requery
    Dim Urut As String * 5
    Dim Hitung As Long
    With rsuser
        If .EOF Then
            Urut = "FC" + "001"
            Text1 = Urut
        Else
            Hitung = Right(!Kodelapangan, 3) + 1
            Urut = "FC" + Right("00" & Hitung, 3)
        End If
        Text1 = Urut
    End With
End Sub

Private Sub Command1_Click()
If Command1.Caption = "&Baru" Then
Command1.Caption = "&Simpan"
DataGrid1.Enabled = False
Text1 = ""
Text2 = ""
Text5 = ""
autonumber
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = True
Text2.Enabled = True
Else
If Text1 = "" Or Text2 = "" Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
Else
'Dim rssimpan As New ADODB.Recordset
'rssimpan.Open "insert into Mstlapangan (KodeLapangan,NamaLapangan,Foto,Lokasi) values ('" & Text1 & "','" & Text2 & "','" & Picture1 & "','" & Text5 & "')", Koneksi
Adodc2.Recordset.AddNew
Adodc2.Recordset!Kodelapangan = Text1
Adodc2.Recordset!NamaLapangan = Text2
Adodc2.Recordset!Foto = Picture1
Adodc2.Recordset!Lokasi = Text5
Adodc2.Recordset.Update
Adodc2.Recordset.Requery
MsgBox "Data Berhasil Disimpan", vbInformation + vbOKOnly, "Informasi"
Form_Load
DataGrid1.Enabled = True
End If
End If
End Sub

Private Sub Command2_Click()
Form_Load
DataGrid1.Enabled = True
Picture1.Picture = LoadPicture("")
End Sub

Private Sub Command3_Click()
If Text1 = "" Or Text2 = "" Then
MsgBox "Pilih Data Dahulu", vbCritical + vbOKOnly, "Peringatan"
Else
If Command3.Caption = "&Ubah" Then
Command3.Caption = "&Simpan"
Command1.Enabled = False
Command2.Enabled = True
Command4.Enabled = False
DataGrid1.Enabled = False
Text2.Enabled = True
Command5.Enabled = True
Else
If Text1 = "" Or Text2 = "" Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
Else
Adodc2.ConnectionString = Koneksi
Adodc2.RecordSource = "select * from Mstlapangan where KodeLapangan='" & Text1 & "'"
Adodc2.Refresh

'Adodc2.Recordset!Kodelapangan = Text1
'Adodc2.Recordset!NamaLapangan = Text2
'Adodc2.Recordset!Lokasi = Text5
'Adodc2.Recordset!Foto = Picture1
'Adodc2.Recordset.Update
'Adodc2.Recordset.Requery
''Adodc3.ConnectionString = Koneksi
''Adodc3.RecordSource = "update MstLapangan set NamaLapangan='" & Text2 & "',Lokasi='" & Text5 & "' where KodeLapangan='" & Text1 & "'"
''Adodc3.Refresh
sqlupdate = "update MstLapangan set NamaLapangan='" & Text2 & "' where KodeLapangan='" & Text1 & "'"
Koneksi.Execute sqlupdate
Adodc2.Recordset!Lokasi = Text5
Adodc2.Recordset!Foto = Picture1
Adodc2.Recordset.Update
Adodc2.Recordset.Requery
MsgBox "Data Berhasil Diubah", vbInformation + vbOKOnly, "Informasi"
Form_Load
DataGrid1.Enabled = True
End If
End If
End If
End Sub

Private Sub Command4_Click()
'If Not Adodc1.Recordset.EOF = False Then
Dim ist As New ADODB.Recordset
ist.Open "select * from MstLapangan", Koneksi
ist.Requery
If Not ist.EOF = False Then
MsgBox "Data Kosong", vbCritical + vbOKOnly, "Peringatan"
Else
If MsgBox("Apakah Anda Yakin Akan Menghapus ?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
Dim sqldelete As String
sqldelete = "delete from MstLapangan where KodeLapangan='" & Text1 & "'"
Koneksi.Execute sqldelete
'Adodc1.Recordset.Delete
MsgBox "Data Berhasil Dihapus", vbInformation + vbOKOnly, "Informasi"
Form_Load
End If
End If
End Sub

Private Sub Command5_Click()
CommonDialog1.DialogTitle = "Pilih Gambar"
CommonDialog1.Filter = "File JPEG|*.jpg"
CommonDialog1.ShowOpen
Text5 = CommonDialog1.FileName
End Sub

Private Sub DataGrid1_Click()
Dim anu As New ADODB.Recordset
anu.Open "select * from MstLapangan", Koneksi
anu.Requery
If anu.EOF Then
'If Not Adodc1.Recordset.EOF = False Then
MsgBox "Data Masih Kosong", vbInformation + vbOKOnly, "Informasi"
Else
'Text1 = Adodc1.Recordset!KodeLapangan
'Text2 = Adodc1.Recordset!NamaLapangan
'Text5 = Adodc1.Recordset!lokasi
Text1 = DataGrid1.Columns(0)
Text2 = DataGrid1.Columns(1)
Dim klo As New ADODB.Recordset
klo.Open "select * from Mstlapangan where KodeLapangan='" & Text1 & "'", Koneksi
klo.Requery
Text5 = klo!Lokasi
'If DataGrid1.Columns(3) = "" Then
'Text5 = ""
'Else
'Text5 = DataGrid1.Columns(3)
'End If
End If
End Sub

Private Sub Form_Activate()
'Dim ist As New ADODB.Recordset
'ist.Open "select * from MstLapangan", Koneksi
'With ist
'    If Not (.BOF And .EOF) Then
'      mvBookMark = .Bookmark
'    End If
'End With
'Set DataGrid1.DataSource = ist.DataSource
'Call bukagrid
End Sub
'Function bukagrid()
'Dim rsbarang As New ADODB.Recordset
'    If rsbarang.State = 1 Then rsbarang.Close
'    rsbarang.Open "select * from MstLapangan", Koneksi
'    If Not rsbarang.BOF And rsbarang.EOF Then
'        bookkmark = rsbarang.Bookmark
'    End If
'    Set DataGrid1.DataSource = rsbarang.DataSource
'End Function
Private Sub Form_Load()
Call BukaDB
'Call grid
Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "select KodeLapangan,Namalapangan from MstLapangan order by KodeLapangan"
Adodc1.Refresh

Adodc2.ConnectionString = Koneksi
Adodc2.RecordSource = "select * from MstLapangan"
Adodc2.Refresh

'Call dtgrid
'Set DataGrid1.DataSource = rsData
'With DataGrid1
'End With

Text1 = ""
Text2 = ""
Text5 = ""
Text1.Enabled = False
Text2.Enabled = False
Text5.Enabled = False
Command5.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command1.Caption = "&Baru"
Command3.Caption = "&Ubah"
Picture1.Picture = LoadPicture("")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command5.SetFocus
End If
End Sub

Private Sub Text3_Change()
Call BukaDB
rsuser.CursorLocation = adUseClient
    rsuser.Open "Select * from MstLapangan where KodeLapangan like '%" & Text3 & "%'", Koneksi
    If Not rsuser.EOF Then
        With rsuser
            With DataGrid1
                Set .DataSource = rsuser
                    .Refresh
            End With
        End With
End If

''Adodc1.RecordSource = "Select * from MstLapangan where KodeLapangan like '%" & TandaPetik(Text3) & "%'"
''Adodc1.Refresh
''If Not Adodc1.Recordset.EOF Then
''    With DataGrid1
''    Set .DataSource = Adodc1
''        .Refresh
'    End With
'End If
End Sub

Private Sub Text4_Change()
'Adodc1.RecordSource = "Select * from MstLapangan where NamaLapangan like '%" & TandaPetik(Text4) & "%'"
'Adodc1.Refresh
'If Not Adodc1.Recordset.EOF Then
'    With DataGrid1
'    Set .DataSource = Adodc1
'        .Refresh
'    End With
'End If
Call BukaDB
rsuser.CursorLocation = adUseClient
    rsuser.Open "Select * from MstLapangan where NamaLapangan like '%" & Text4 & "%'", Koneksi
    If Not rsuser.EOF Then
        With rsuser
            With DataGrid1
                Set .DataSource = rsuser
                    .Refresh
            End With
        End With
End If
End Sub

Private Sub Text5_Change()
Picture1.Picture = LoadPicture(Text5)
End Sub
