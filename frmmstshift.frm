VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.OCX"
Begin VB.Form frmmstshift 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18435
   Icon            =   "frmmstshift.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   18435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10290
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18435
      _ExtentX        =   32517
      _ExtentY        =   18150
      _Version        =   262144
      AutoSize        =   1
      Locked          =   -1  'True
      PaneTree        =   "frmmstshift.frx":9E4A
      Begin Threed.SSPanel SSPanel3 
         Height          =   2550
         Left            =   30
         TabIndex        =   3
         Top             =   7710
         Width           =   18375
         _ExtentX        =   32411
         _ExtentY        =   4498
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CommandButton Command4 
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
            Left            =   4920
            Picture         =   "frmmstshift.frx":9EBC
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
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
            Left            =   3480
            Picture         =   "frmmstshift.frx":BB86
            Style           =   1  'Graphical
            TabIndex        =   14
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
            Left            =   2040
            Picture         =   "frmmstshift.frx":D850
            Style           =   1  'Graphical
            TabIndex        =   13
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
            Left            =   600
            Picture         =   "frmmstshift.frx":F51A
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   360
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   3585
         Left            =   30
         TabIndex        =   2
         Top             =   4035
         Width           =   18375
         _ExtentX        =   32411
         _ExtentY        =   6324
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmmstshift.frx":111E4
            Height          =   3495
            Left            =   0
            TabIndex        =   15
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
            Caption         =   "Master Shift"
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
            Left            =   1440
            Top             =   1320
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
         Height          =   3915
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   18375
         _ExtentX        =   32411
         _ExtentY        =   6906
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
         Begin VB.Frame Frame1 
            BackColor       =   &H8000000D&
            Caption         =   "Cari Data"
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
            Left            =   600
            TabIndex        =   17
            Top             =   2040
            Width           =   5655
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
               Left            =   2280
               TabIndex        =   21
               Top             =   1080
               Width           =   2655
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
               Left            =   2280
               TabIndex        =   19
               Top             =   480
               Width           =   2655
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Jam Mulai :"
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
               TabIndex        =   20
               Top             =   1080
               Width           =   1455
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Kode Shift :"
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
               TabIndex        =   18
               Top             =   480
               Width           =   1335
            End
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   495
            Left            =   2160
            TabIndex        =   16
            Top             =   1200
            Width           =   2655
            _ExtentX        =   4683
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
            Format          =   123666434
            CurrentDate     =   41663
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
            Left            =   9720
            TabIndex        =   11
            Top             =   1200
            Width           =   2295
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
            Left            =   9720
            TabIndex        =   9
            Top             =   480
            Width           =   2295
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
            Height          =   435
            Left            =   2160
            TabIndex        =   5
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Member :"
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
            Left            =   7560
            TabIndex        =   10
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga                 :"
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
            Left            =   7560
            TabIndex        =   8
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Mulai :"
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
            TabIndex        =   7
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Shift :"
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
            Width           =   1575
         End
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   9600
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
   End
End
Attribute VB_Name = "frmmstshift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Command1.Caption = "&Baru" Then
Command1.Caption = "&Simpan"
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
kosong
autonumber
Text3.Enabled = True
Text4.Enabled = True
DTPicker1.Enabled = True
DataGrid1.Enabled = True
Else
If Text1 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
Else
Dim cari As String
cari = "JamMulai='" & Format(DTPicker1, "HH:MM:SS") & "'"
With Adodc1.Recordset
.Find cari
If Not .EOF Then
MsgBox "Data Sudah Ada", vbCritical + vbOKOnly, "Peringatan"

'Set rsuser = New ADODB.Recordset
'rsuser.LockType = adLockOptimistic
'rsuser.CursorType = adOpenDynamic
'rsuser.Open "select * from MstShift", Koneksi, , , adCmdText
'
'rsuser.Filter = " JamMulai= '" & Format(DTPicker1, "HH:MM:SS") & "'"
'If Not rsuser.EOF Then
'MsgBox "Jam Sudah Ada", vbCritical + vbOKOnly, "Informasi"
Else
Dim iku As New ADODB.Recordset
iku.Open "insert into MstShift (KodeShift,JamMulai,Harga,HargaMember) values ('" & Text1 & "','" & Format(DTPicker1, "HH:MM:SS") & "','" & Text3 & "','" & Text4 & "')", Koneksi
'Adodc1.Recordset.AddNew
'Adodc1.Recordset!KodeShift = Text1
'Adodc1.Recordset!JamMulai = Format(DTPicker1, "HH:MM:SS")
'Adodc1.Recordset!Harga = Text3
'Adodc1.Recordset!HargaMember = Text4
'Adodc1.Recordset.Update
'Adodc1.Recordset.Requery
MsgBox "Data berhasil Disimpan", vbInformation + vbOKOnly, "Informasi"
Form_Load
End If
End With
End If
End If
End Sub

Private Sub Command2_Click()
Form_Load
End Sub

Private Sub Command3_Click()
If Not Adodc1.Recordset.EOF = False Then
MsgBox "Data Kosong", vbCritical + vbOKOnly, "Peringatan"
Else
If MsgBox("Apakah Anda Yakin Akan Menghapus?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
'Adodc1.Recordset.Delete
'Adodc1.Recordset.Update
'Adodc1.Recordset.Requery
Dim hps As String
hps = "delete from MstShift where KodeShift='" & Text1 & "'"
Koneksi.Execute hps
MsgBox "Data Berhasil Dihapus", vbInformation + vbOKOnly, "Informasi"
Form_Load
End If
End If
End Sub

Private Sub Command4_Click()
If Text1 = "" Then
MsgBox "Data Belum Dipilih", vbCritical + vbOKOnly, "Peringatan"
Else
If Command4.Caption = "&Ubah" Then
Command4.Caption = "&Simpan"
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
Text3.Enabled = True
Text4.Enabled = True
DataGrid1.Enabled = False
Else
If Text1 = "" Or Text3 = "" Or Text4 = "" Or DTPicker1 = 0 Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly
Else
'Set rsuser = New ADODB.Recordset
'rsuser.Open "select * from MstShift", Koneksi
'rsuser.Filter = " JamMulai= '" & DTPicker1 & "'"
'If Not rsuser.EOF Then
'MsgBox "Jam Sudah Ada", vbCritical + vbOKOnly, "Informasi"
'Else
'Adodc1.Recordset!KodeShift = Text1
'Adodc1.Recordset!JamMulai = DTPicker1
'Adodc1.Recordset!Harga = Text3
'Adodc1.Recordset!HargaMember = Text4
'Adodc1.Recordset.Update
'Adodc1.Recordset.Requery
Dim ubah As String
ubah = "update MstShift set JamMulai ='" & Format(DTPicker1, "HH:MM:SS") & "', Harga='" & Text3 & "', HargaMember='" & Text4 & "' where KodeShift='" & Text1 & "'"
Koneksi.Execute ubah
MsgBox "Data Berhasil Diubah", vbInformation + vbOKOnly, "Informasi"
Form_Load
End If
End If
End If
'End If
End Sub

Private Sub DataGrid1_Click()
If Not Adodc1.Recordset.EOF = False Then
MsgBox "Data masih Kosong", vbInformation + vbOKOnly, "Informasi"
Else
Text1 = Adodc1.Recordset!KodeShift
Text3 = Adodc1.Recordset!Harga
Text4 = Adodc1.Recordset!HargaMember
DTPicker1 = Adodc1.Recordset!JamMulai
End If
End Sub

Private Sub Form_Load()
Call BukaDB
Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "select KodeShift,JamMulai,Harga,HargaMember from MstShift order by KodeShift"
Adodc1.Refresh

Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command1.Caption = "&Baru"
Command4.Caption = "&Ubah"
kosong
DTPicker1.Value = Format(DTPicker1, "HH:MM:SS")
Text1.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
DTPicker1.Enabled = False
DataGrid1.Enabled = True
End Sub
Sub autonumber()
Call BukaDB
rsuser.Open ("SELECT * FROM MstShift WHERE KodeShift in(select max(KodeShift) from MstShift)order by KodeShift desc"), Koneksi
rsuser.Requery
    Dim Urut As String * 6
    Dim Hitung As Long
    With rsuser
        If .EOF Then
            Urut = "SHF" + "001"
            Text1 = Urut
        Else
            Hitung = (Right(!KodeShift, 3)) + 1
            Urut = "SHF" + Right("00" & Hitung, 3)
        End If
        Text1 = Urut
    End With
End Sub
Sub kosong()
Text1 = ""
Text3 = ""
Text4 = ""
DTPicker1 = 0
End Sub

Private Sub Text3_Change()
If Not IsNumeric(Text3) Then Text3 = "0"
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub Text4_Change()
If Not IsNumeric(Text4) Then Text4 = "0"
End Sub

Private Sub Text5_Change()
Adodc1.RecordSource = "Select KodeShift,JamMulai,Harga,HargaMember from MstShift where KodeShift like '%" & TandaPetik(Text5) & "%'"
Adodc1.Refresh
If Not Adodc1.Recordset.EOF Then
    With DataGrid1
    Set .DataSource = Adodc1
        .Refresh
    End With
End If
End Sub

Private Sub Text6_Change()
Adodc1.RecordSource = "Select KodeShift,JamMulai,Harga,HargaMember from MstShift where JamMulai like '%" & TandaPetik(Text6) & "%'"
Adodc1.Refresh
If Not Adodc1.Recordset.EOF Then
    With DataGrid1
    Set .DataSource = Adodc1
        .Refresh
    End With
End If
End Sub
