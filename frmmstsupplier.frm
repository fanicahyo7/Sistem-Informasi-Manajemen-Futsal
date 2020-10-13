VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.OCX"
Begin VB.Form frmmstsupplier 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20490
   Icon            =   "frmmstsupplier.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20490
      _ExtentX        =   36142
      _ExtentY        =   16854
      _Version        =   262144
      AutoSize        =   1
      Locked          =   -1  'True
      PaneTree        =   "frmmstsupplier.frx":9E4A
      Begin Threed.SSPanel SSPanel1 
         Height          =   3900
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   20430
         _ExtentX        =   36036
         _ExtentY        =   6879
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Frame Frame1 
            BackColor       =   &H8000000D&
            Caption         =   "Cari Data Supplier"
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
            Left            =   7920
            TabIndex        =   17
            Top             =   1680
            Width           =   6735
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
               Height          =   450
               Left            =   3360
               TabIndex        =   21
               Top             =   1080
               Width           =   2895
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
               Height          =   450
               Left            =   3360
               TabIndex        =   19
               Top             =   480
               Width           =   2895
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Nama Supplier :"
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
               TabIndex        =   20
               Top             =   1080
               Width           =   1815
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Kode Supplier  :"
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
               Top             =   480
               Width           =   1935
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
            Height          =   450
            Left            =   11280
            TabIndex        =   10
            Top             =   1080
            Width           =   2895
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
            Height          =   525
            Left            =   11280
            TabIndex        =   9
            Top             =   360
            Width           =   2895
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
            Height          =   450
            Left            =   2880
            TabIndex        =   8
            Top             =   2760
            Width           =   2895
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
            Height          =   975
            Left            =   2880
            TabIndex        =   7
            Top             =   1560
            Width           =   2895
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
            Left            =   2880
            TabIndex        =   6
            Top             =   960
            Width           =   2895
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
            Left            =   2880
            TabIndex        =   5
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label8 
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
            Height          =   615
            Left            =   7920
            TabIndex        =   16
            Top             =   1080
            Width           =   3015
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Penanggung Jawab       :"
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
            Left            =   7920
            TabIndex        =   15
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Telp                     :"
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
            TabIndex        =   14
            Top             =   2760
            Width           =   2055
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat                :"
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
            TabIndex        =   13
            Top             =   1560
            Width           =   2055
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Supplier :"
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
            TabIndex        =   12
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Supplier  :"
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
            TabIndex        =   11
            Top             =   360
            Width           =   1815
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1335
         Left            =   30
         TabIndex        =   2
         Top             =   8190
         Width           =   20430
         _ExtentX        =   36036
         _ExtentY        =   2355
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
            Left            =   4800
            Picture         =   "frmmstsupplier.frx":9EBC
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   120
            Width           =   1095
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
            Left            =   3360
            Picture         =   "frmmstsupplier.frx":BB86
            Style           =   1  'Graphical
            TabIndex        =   24
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
            Left            =   1920
            Picture         =   "frmmstsupplier.frx":D850
            Style           =   1  'Graphical
            TabIndex        =   23
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
            Left            =   480
            Picture         =   "frmmstsupplier.frx":F51A
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   120
            Width           =   1095
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   4080
         Left            =   30
         TabIndex        =   1
         Top             =   4020
         Width           =   20430
         _ExtentX        =   36036
         _ExtentY        =   7197
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   5400
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
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmmstsupplier.frx":111E4
            Height          =   3975
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   18135
            _ExtentX        =   31988
            _ExtentY        =   7011
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
            Caption         =   "Supplier"
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
   End
End
Attribute VB_Name = "frmmstsupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub autonumber()
Call BukaDB
rsuser.Open ("SELECT * FROM MstSupplier WHERE KodeSupplier in(select max(KodeSupplier) from MstSupplier)order by KodeSupplier desc"), Koneksi
rsuser.Requery
    Dim Urut As String * 6
    Dim Hitung As Long
    With rsuser
        If .EOF Then
            Urut = "SPL" + "001"
            Text1 = Urut
        Else
            Hitung = Right(!KodeSupplier, 3) + 1
            Urut = "SPL" + Right("00" & Hitung, 3)
        End If
        Text1 = Urut
    End With
End Sub
Sub mati()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
End Sub
Sub hidup()
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
End Sub
Private Sub Command1_Click()
If Command1.Caption = "&Baru" Then
Command1.Caption = "&Simpan"
hidup
DataGrid1.Enabled = False
kosong
Call autonumber
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Else
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
Else
Dim rssimpan As New ADODB.Recordset
rssimpan.Open "insert into MstSupplier (KodeSupplier,NamaSupplier,Alamat,Telp,PenanggungJawab,HPPenanggungJawab) values ('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "','" & Text6 & "')", Koneksi
'Adodc1.Recordset.AddNew
'Adodc1.Recordset!KodeSupplier = Text1
'Adodc1.Recordset!NamaSupplier = Text2
'Adodc1.Recordset!alamat = Text3
'Adodc1.Recordset!telp = Text4
'Adodc1.Recordset!PenanggungJawab = Text5
'Adodc1.Recordset!HPPenanggungJawab = Text6
'Adodc1.Recordset.Update
'Adodc1.Recordset.Requery
MsgBox "Data Berhasil Disimpan", vbInformation + vbOKOnly, "Informasi"
Form_Load
DataGrid1.Enabled = True
End If
End If
End Sub

Private Sub Command2_Click()
Form_Load
kosong
DataGrid1.Enabled = True
End Sub

Private Sub Command3_Click()
If Text1 = "" Then
MsgBox "Data Belum DIpilih", vbCritical + vbOKOnly, "Peringatan"
Else
If Command3.Caption = "&Ubah" Then
Command3.Caption = "&Simpan"
Command1.Enabled = False
Command2.Enabled = True
Command4.Enabled = False
hidup
DataGrid1.Enabled = False
Else
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
Else
Dim ubah As String
ubah = "update MstSupplier set NamaSupplier='" & Text2 & "', Alamat='" & Text3 & "', Telp='" & Text4 & "', PenanggungJawab='" & Text5 & "', HPPenanggungJawab='" & Text6 & "' where KodeSupplier='" & Text1 & "'"
Koneksi.Execute ubah
'Adodc1.Recordset!KodeSupplier = Text1
'Adodc1.Recordset!NamaSupplier = Text2
'Adodc1.Recordset!alamat = Text3
'Adodc1.Recordset!telp = Text4
'Adodc1.Recordset!PenanggungJawab = Text5
'Adodc1.Recordset!HPPenanggungJawab = Text6
'Adodc1.Recordset.Update
'Adodc1.Recordset.Requery
MsgBox "Data Berhasil Diubah", vbInformation + vbOKOnly, "Informasi"
Form_Load
DataGrid1.Enabled = True
End If
End If
End If
End Sub

Private Sub Command4_Click()
If Not Adodc1.Recordset.EOF = False Then
MsgBox "Data Kosong", vbCritical + vbOKOnly, "Peringatan"
Else
If MsgBox("Apakah Anda Yakin Akan Menghapus ?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
Dim ddelete As String
ddelete = "delete from MstSupplier where KodeSupplier='" & Text1 & "'"
Koneksi.Execute ddelete
MsgBox "Data Berhasil Dihapus", vbInformation + vbOKOnly, "Informasi"
Form_Load
End If
End If
End Sub

Private Sub DataGrid1_Click()
If Not Adodc1.Recordset.EOF = False Then
MsgBox "Data Masih Kosong", vbInformation + vbOKOnly, "Informasi"
Else
Text1 = DataGrid1.Columns(0)
Text2 = DataGrid1.Columns(1)
Text3 = DataGrid1.Columns(2)
Text4 = DataGrid1.Columns(3)
Text5 = DataGrid1.Columns(4)
Text6 = DataGrid1.Columns(5)
End If
End Sub

Private Sub Form_Load()
 Call BukaDB
 Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "Select * from MstSupplier order by KodeSupplier"
Adodc1.Refresh
kosong
mati
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command1.Caption = "&Baru"
Command3.Caption = "&Ubah"
End Sub

Sub kosong()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub Text4_Change()
If Not IsNumeric(Text4) Then Text4 = "0"
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text6.SetFocus
End If
End Sub

Private Sub Text6_Change()
If Not IsNumeric(Text6) Then Text6 = "0"
End Sub

Private Sub Text7_Change()
Adodc1.RecordSource = "Select * from MstSupplier where KodeSupplier like '%" & TandaPetik(Text7) & "%'"
Adodc1.Refresh
If Not Adodc1.Recordset.EOF Then
    With DataGrid1
    Set .DataSource = Adodc1
        .Refresh
    End With
End If
End Sub

Private Sub Text8_Change()
Adodc1.RecordSource = "Select * from MstSupplier where NamaSupplier like '%" & TandaPetik(Text8) & "%'"
Adodc1.Refresh
If Not Adodc1.Recordset.EOF Then
    With DataGrid1
    Set .DataSource = Adodc1
        .Refresh
    End With
End If
End Sub
