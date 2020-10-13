VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.OCX"
Begin VB.Form frmmstmember 
   BorderStyle     =   0  'None
   Caption         =   "frmmstmember"
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20085
   Icon            =   "frmmstmember.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   9720
   ScaleWidth      =   20085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20085
      _ExtentX        =   35428
      _ExtentY        =   17145
      _Version        =   262144
      AutoSize        =   1
      Locked          =   -1  'True
      PaneTree        =   "frmmstmember.frx":9E4A
      Begin Threed.SSPanel SSPanel3 
         Height          =   1665
         Left            =   30
         TabIndex        =   2
         Top             =   8025
         Width           =   20025
         _ExtentX        =   35322
         _ExtentY        =   2937
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
            Left            =   5280
            Picture         =   "frmmstmember.frx":9EBC
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   360
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
            Left            =   3720
            Picture         =   "frmmstmember.frx":BB86
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   360
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
            Left            =   2280
            Picture         =   "frmmstmember.frx":D850
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   360
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
            Left            =   840
            Picture         =   "frmmstmember.frx":F51A
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   360
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   4395
         Left            =   30
         TabIndex        =   1
         Top             =   3540
         Width           =   20025
         _ExtentX        =   35322
         _ExtentY        =   7752
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   6360
            Top             =   -360
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
            Bindings        =   "frmmstmember.frx":111E4
            Height          =   4335
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   18135
            _ExtentX        =   31988
            _ExtentY        =   7646
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
            Caption         =   "Master Member"
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
         Height          =   3420
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   20025
         _ExtentX        =   35322
         _ExtentY        =   6033
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   330
            Left            =   7560
            Top             =   3480
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
            Left            =   4320
            TabIndex        =   12
            Top             =   2205
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
            Height          =   450
            Left            =   4320
            TabIndex        =   11
            Top             =   960
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
            Left            =   4320
            TabIndex        =   10
            Top             =   1560
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
            Left            =   4320
            TabIndex        =   9
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Biaya                                   :"
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
            TabIndex        =   16
            Top             =   2160
            Width           =   3015
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Jenis Member       :"
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
            TabIndex        =   15
            Top             =   960
            Width           =   3015
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Jangka Waktu (Bulan)    :"
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
            TabIndex        =   14
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Jenis Member        :"
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
            TabIndex        =   13
            Top             =   360
            Width           =   3135
         End
      End
   End
End
Attribute VB_Name = "frmmstmember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub autonumber()
Call BukaDB
rsuser.Open ("SELECT * FROM JenisMember WHERE KodeJenisMember in(select max(KodeJenisMember) from JenisMember)order by KodeJenisMember desc"), Koneksi
rsuser.Requery
    Dim Urut As String * 6
    Dim Hitung As Long
    With rsuser
        If .EOF Then
            Urut = "KMB" + "001"
            Text1 = Urut
        Else
            Hitung = Right(!KodeJenisMember, 3) + 1
            Urut = "KMB" + Right("00" & Hitung, 3)
        End If
        Text1 = Urut
    End With
End Sub
Private Sub Command1_Click()
If Command1.Caption = "&Baru" Then
Command1.Caption = "&Simpan"
hidup
DataGrid1.Enabled = False
kosong
autonumber

Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Else
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
Else
Adodc2.RecordSource = "select * from JenisMember"
Adodc2.Refresh
Dim cari As String
cari = "JumlahBulan='" & Text2 & "'"
With Adodc2.Recordset
.Find cari
If Not .EOF Then
MsgBox "Data Sudah Ada", vbCritical + vbOKOnly, "Peringatan"
Else
Dim anu As New ADODB.Recordset
anu.Open "insert into jenismember (KodeJenisMember,NamaJenisMember,JumlahBulan,Biaya) values ('" & Text1 & "','" & Text3 & "','" & Text2 & "','" & Text4 & "')", Koneksi
'Adodc1.Recordset.AddNew
'Adodc1.Recordset!KodeJenisMember = Text1
'Adodc1.Recordset!NamaJenisMember = Text3
'Adodc1.Recordset!JumlahBulan = Text2
'Adodc1.Recordset!Biaya = Text4
'Adodc1.Recordset.Update
'Adodc1.Recordset.Requery
MsgBox "Data Berhasil Disimpan", vbInformation + vbOKOnly, "Informasi"
Form_Load
DataGrid1.Enabled = True
End If
End With
End If
End If
End Sub

Private Sub Command2_Click()
Form_Load
DataGrid1.Enabled = True
End Sub

Private Sub Command3_Click()
If Text1 = "" Then
MsgBox "Data Belum Dipilih", vbCritical + vbOKOnly, "Peringatan"
Else
If Command3.Caption = "&Ubah" Then
Command3.Caption = "&Simpan"
DataGrid1.Enabled = False
hidup
Text2.Enabled = False
Command1.Enabled = False
Command2.Enabled = True
Command4.Enabled = False
Else

Adodc2.RecordSource = "select * from JenisMember where JumlahBulan like '" & Text2 & "'"
Adodc2.Refresh
'If Not Adodc2.Recordset.EOF Then
'MsgBox "Jumlah Bulan Sudah Ada", vbCritical + vbOKOnly, "Peringatan"
'Else
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
Else
Dim ubah As String
ubah = "update JenisMember set NamaJenisMember='" & Text3.Text & "', JumlahBulan='" & Text2.Text & "', Biaya='" & Text4.Text & "' where KodeJenisMember='" & Text1.Text & "'"
Koneksi.Execute ubah
'Adodc1.Recordset!KodeJenisMember = Text1
'Adodc1.Recordset!NamaJenisMember = Text3
'Adodc1.Recordset!JumlahBulan = Text2
'Adodc1.Recordset!Biaya = Text4
'Adodc1.Recordset.Update
'Adodc1.Recordset.Requery
MsgBox "Data Berhasil Disimpan", vbInformation + vbOKOnly, "Informasi"
Form_Load
DataGrid1.Enabled = True
'End If
End If
End If
End If
End Sub

Private Sub Command4_Click()
If Not Adodc1.Recordset.EOF = False Then
MsgBox "Data Kosong", vbCritical + vbOKOnly, "Peringatan"
Else
If MsgBox("Apakah Anda Yakin Akan Menghapus ?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
'Adodc1.Recordset.Delete
Dim anu As String
anu = "delete from JenisMember where KodeJenisMember='" & Text1 & "'"
Koneksi.Execute anu
MsgBox "Data Berhasil Dihapus", vbInformation + vbOKOnly, "Peringatan"
Form_Load
End If
End If
End Sub

Private Sub DataGrid1_Click()
If Not Adodc1.Recordset.EOF = False Then
MsgBox "Data Masih Kosong", vbInformation + vbOKOnly, "Informasi"
Else
Text1 = Adodc1.Recordset!KodeJenisMember
Text2 = Adodc1.Recordset!JumlahBulan
Text3 = Adodc1.Recordset!NamaJenisMember
Text4 = Adodc1.Recordset!Biaya
End If
End Sub

Private Sub Form_Load()
Call BukaDB
Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "select KodeJenisMember,NamaJenisMember,JumlahBulan,Biaya from JenisMember order by KodeJenisMember"
Adodc1.Refresh

Adodc2.ConnectionString = Koneksi

kosong
mati
Text1.Enabled = False
Command1.Caption = "&Baru"
Command3.Caption = "&Ubah"
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
End Sub
Sub kosong()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
End Sub
Sub mati()
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
End Sub
Sub hidup()
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
End Sub

Private Sub Text2_Change()
If Not IsNumeric(Text2) Then Text2 = "0"
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub

Private Sub Text4_Change()
If Not IsNumeric(Text4) Then Text4 = "0"
End Sub
