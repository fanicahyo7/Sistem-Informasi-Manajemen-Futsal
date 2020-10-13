VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.OCX"
Begin VB.Form formMember 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   Icon            =   "formMember.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9885
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9885
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20490
      _ExtentX        =   36142
      _ExtentY        =   17436
      _Version        =   262144
      AutoSize        =   1
      Locked          =   -1  'True
      PaneTree        =   "formMember.frx":9E4A
      Begin Threed.SSPanel SSPanel3 
         Height          =   3885
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   20430
         _ExtentX        =   36036
         _ExtentY        =   6853
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   330
            Left            =   4560
            Top             =   4080
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
         Begin VB.CheckBox Check2 
            BackColor       =   &H8000000D&
            Caption         =   "Tidak Aktif/Mati"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   14760
            TabIndex        =   37
            Top             =   2520
            Width           =   2175
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H8000000D&
            Caption         =   "Aktif/Hidup"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   14760
            TabIndex        =   36
            Top             =   2040
            Width           =   1575
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
            Height          =   375
            Left            =   9240
            TabIndex        =   34
            Top             =   2040
            Width           =   2655
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
            Height          =   375
            Left            =   9240
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   1440
            Width           =   2655
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
            Height          =   375
            Left            =   9240
            TabIndex        =   30
            Top             =   2640
            Width           =   2655
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H8000000D&
            Caption         =   "Tidak"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   10200
            TabIndex        =   22
            Top             =   960
            Width           =   855
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H8000000D&
            Caption         =   "Ya"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9240
            TabIndex        =   21
            Top             =   960
            Width           =   615
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   14760
            TabIndex        =   18
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
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
            Format          =   189071361
            CurrentDate     =   41658
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   495
            Left            =   9240
            TabIndex        =   16
            Top             =   240
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
            Format          =   189071361
            CurrentDate     =   41658
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
            Height          =   375
            Left            =   14760
            TabIndex        =   14
            Top             =   1200
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
            Height          =   375
            Left            =   2520
            TabIndex        =   13
            Top             =   2880
            Width           =   2655
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
            Height          =   615
            Left            =   2520
            TabIndex        =   12
            Top             =   2040
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
            Height          =   375
            Left            =   2520
            TabIndex        =   11
            Top             =   1440
            Width           =   2655
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
            Height          =   375
            Left            =   2520
            TabIndex        =   10
            Top             =   840
            Width           =   2655
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
            Height          =   375
            Left            =   2520
            TabIndex        =   9
            Top             =   240
            Width           =   2655
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   6360
            Top             =   3960
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
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Status Member    :"
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
            Left            =   12600
            TabIndex        =   35
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Label Label13 
            BackColor       =   &H8000000D&
            Caption         =   "Nama Jenis Member :"
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
            Left            =   6480
            TabIndex        =   33
            Top             =   2040
            Width           =   2535
         End
         Begin VB.Label Label12 
            BackColor       =   &H8000000D&
            Caption         =   "Kode Jenis Member  :"
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
            Left            =   6480
            TabIndex        =   31
            Top             =   1440
            Width           =   2535
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Bulan              :"
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
            Left            =   6480
            TabIndex        =   29
            Top             =   2640
            Width           =   2535
         End
         Begin VB.Label Label10 
            BackColor       =   &H8000000D&
            Caption         =   "Tanpa Batas               :"
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
            Left            =   6480
            TabIndex        =   20
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Biaya                      :"
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
            Left            =   12600
            TabIndex        =   19
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Hangus :"
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
            Left            =   12600
            TabIndex        =   17
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label7 
            BackColor       =   &H8000000D&
            Caption         =   "Tanggal Daftar           :"
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
            Left            =   6480
            TabIndex        =   15
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label5 
            BackColor       =   &H8000000D&
            Caption         =   "Telp               :"
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
            TabIndex        =   8
            Top             =   2880
            Width           =   1695
         End
         Begin VB.Label Label4 
            BackColor       =   &H8000000D&
            Caption         =   "Alamat          :"
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
            TabIndex        =   7
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000D&
            Caption         =   "Atas Nama   :"
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
            TabIndex        =   6
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            Caption         =   "Nama Tim     :"
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
            TabIndex        =   5
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H8000000D&
            Caption         =   "No. Register :"
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
            TabIndex        =   4
            Top             =   240
            Width           =   1695
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1545
         Left            =   30
         TabIndex        =   2
         Top             =   8310
         Width           =   20430
         _ExtentX        =   36036
         _ExtentY        =   2725
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CommandButton Command5 
            Caption         =   "&Perpanjang"
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
            Left            =   5760
            Picture         =   "formMember.frx":9EBC
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   120
            Width           =   1215
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
            Left            =   4320
            Picture         =   "formMember.frx":FACE
            Style           =   1  'Graphical
            TabIndex        =   27
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
            Left            =   3000
            Picture         =   "formMember.frx":11798
            Style           =   1  'Graphical
            TabIndex        =   26
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
            Picture         =   "formMember.frx":13462
            Style           =   1  'Graphical
            TabIndex        =   25
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
            Picture         =   "formMember.frx":1512C
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   120
            Width           =   1095
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   4215
         Left            =   30
         TabIndex        =   1
         Top             =   4005
         Width           =   20430
         _ExtentX        =   36036
         _ExtentY        =   7435
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "formMember.frx":16DF6
            Height          =   4215
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   18135
            _ExtentX        =   31988
            _ExtentY        =   7435
            _Version        =   393216
            AllowUpdate     =   0   'False
            BorderStyle     =   0
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
            Caption         =   "Member"
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
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   495
      Left            =   9600
      TabIndex        =   28
      Top             =   5280
      Width           =   1215
   End
End
Attribute VB_Name = "formMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim lt As Boolean

Private Sub Check1_Click()
If Check1.Value = 1 Then
Check2.Value = 0
Else
If Check2.Value = 1 Then
Check1.Value = 0
End If
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Check1.Value = 0
Else
If Check1.Value = 1 Then
Check2.Value = 0
End If
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
Check4.Value = 0
lt = "True"
DTPicker2.Value = Format(Date, "31 / 12 / 9999")
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
Check3.Value = 0
lt = "False"
DTPicker2.Value = Format(Now, "MM/DD/yyyy")
End If
End Sub

Private Sub Command1_Click()
If Command1.Caption = "&Baru" Then
Command1.Caption = "&Simpan"
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
hidup
DataGrid1.Enabled = False
kosong
autonumber
DTPicker2.Enabled = False
Else
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7 = "" Or DTPicker1.Value = "" Or DTPicker2.Value = "" Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
ElseIf Check3.Value = 0 And Check4.Value = 0 Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
ElseIf DTPicker1.Value = 0 And DTPicker2.Value = 0 Then
MsgBox "Data Belum Lepngkap", vbCritical + vbOKOnly, "Peringatan"
ElseIf DTPicker2 > Now Then
Dim rspl1 As String
rspl1 = "insert into Member (NoRegister,NamaTim,AtasNama,Alamat,Telp,Tanggaldaftar,TanggalHangus,Biaya,KodeJenisMember,Berbatas,isMember) values ('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "','" & Format(DTPicker1, "yyyy/mm/dd") & "','" & Format(DTPicker2, "yyyy/mm/dd") & "','" & Text6 & "','" & Text7 & "','" & lt & "','-1')"
Koneksi.Execute rspl1
Form_Load
MsgBox "Data Disimpan", vbInformation + vbOKOnly, "Informasi"
ElseIf DTPicker2 < Now Then
Dim rspl2 As String
rspl2 = "insert into Member (NoRegister,NamaTim,AtasNama,Alamat,Telp,Tanggaldaftar,TanggalHangus,Biaya,KodeJenisMember,Berbatas,isMember) values ('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "','" & Format(DTPicker1, "yyyy/mm/dd") & "','" & Format(DTPicker2, "yyyy/mm/dd") & "','" & Text6 & "','" & Text7 & "','" & lt & "','0')"
Koneksi.Execute rspl2
Form_Load
MsgBox "Data Disimpan", vbInformation + vbOKOnly, "Informasi"
'Adodc1.Recordset.AddNew
'Adodc1.Recordset!NoRegister = Text1.Text
'Adodc1.Recordset!NamaTim = Text2.Text
'Adodc1.Recordset!AtasNama = Text3.Text
'Adodc1.Recordset!alamat = Text4.Text
'Adodc1.Recordset!telp = Text5.Text
'Adodc1.Recordset!TanggalDaftar = Format(DTPicker1.Value, "MM/DD/yyyy")
'Adodc1.Recordset!TanggalHangus = Format(DTPicker2.Value, "MM/DD/yyyy")
'Adodc1.Recordset!Biaya = Text6.Text
'Adodc1.Recordset!KodeJenisMember = Text7.Text
'Adodc1.Recordset!Berbatas = lt
'If DTPicker2 > Now Then
'Adodc1.Recordset!isMember = "-1"
'ElseIf DTPicker2 < Now Then
'Adodc1.Recordset!isMember = "0"
'End If
'Adodc1.Recordset.Update
'Adodc1.Recordset.Requery
End If
End If
End Sub
Private Sub autonumber()
Call BukaDB
rsuser.Open ("SELECT * FROM Member WHERE NoRegister in(select max(NoRegister) from Member)order by NoRegister desc"), Koneksi
rsuser.Requery
    Dim Urut As String * 6
    Dim Hitung As Long
    With rsuser
        If .EOF Then
            Urut = "MMB" + "001"
            Text1 = Urut
        Else
            Hitung = Right(!NoRegister, 3) + 1
            Urut = "MMB" + Right("00" & Hitung, 3)
        End If
        Text1 = Urut
    End With
End Sub

Private Sub Command2_Click()
Form_Load
End Sub

Private Sub Command3_Click()
If Text1 = "" Then
MsgBox "Data Belum Dipilih", vbCritical + vbOKOnly, "Peringatan"
Else
If Check2.Value = 1 Then
MsgBox "Masa Berlaku Habis", vbCritical + vbOKOnly, "Peringatn"
Else
If Command3.Caption = "&Ubah" Then
Command3.Caption = "&Simpan"
Command1.Enabled = False
Command2.Enabled = True
Command4.Enabled = False
DataGrid1.Enabled = False
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Else
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7 = "" Or DTPicker1.Value = "" Or DTPicker2.Value = "" Then
MsgBox "Data Belum Lengkap", vbInformation + vbOKOnly, "Peringatan"
ElseIf Check3.Value = 0 And Check4.Value = 0 Then
MsgBox "Data Belum Lengkap", vbInformation + vbOKOnly, "Peringatan"
ElseIf DTPicker1.Value = 0 And DTPicker2.Value = 0 Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
Else
'Adodc1.Recordset!NoRegister = Text1.Text
'Adodc1.Recordset!NamaTim = Text2.Text
'Adodc1.Recordset!AtasNama = Text3.Text
'Adodc1.Recordset!alamat = Text4.Text
'Adodc1.Recordset!telp = Val(Text5.Text)
'Adodc1.Recordset!TanggalDaftar = Format(DTPicker1.Value, "d/MM/yyyy")
'Adodc1.Recordset!TanggalHangus = Format(DTPicker2.Value, "d/MM/yyyy")
'Adodc1.Recordset!Biaya = Text6.Text
'Adodc1.Recordset!KodeJenisMember = Text7.Text
'Adodc1.Recordset!Berbatas = lt
'Adodc1.Recordset.Update
'Adodc1.Recordset.Requery
Dim rspl2 As String
rspl2 = "update Member set Namatim='" & Text2 & "',AtasNama='" & Text3 & "',Alamat='" & Text4 & "',Telp='" & Text5 & "',TanggalDaftar='" & Format(DTPicker1, "yyyy/mm/dd") & "',TanggalHangus='" & Format(DTPicker2, "yyyy/mm/dd") & "',Biaya='" & Text6 & "',KodeJenisMember='" & Text7 & "',Berbatas='" & lt & "' where NoRegister='" & Text1 & "'"
Koneksi.Execute rspl2
Form_Load
MsgBox "Data Berhasil Diubah", vbInformation + vbOKOnly, "Informasi"
End If
End If
End If
End If
End Sub

Private Sub Command4_Click()
If Not Adodc1.Recordset.EOF = False Then
MsgBox "Data Kosong", vbCritical + vbOKOnly, "Peringatan"
Else
If MsgBox("Apakah Anda Yakin Akan Menghapus", vbQuestion + vbYesNo, "Informasi") = vbYes Then
Dim hps As String
hps = "delete from Member where NoRegister='" & Text1 & "'"
Koneksi.Execute hps
MsgBox "Data Berhasil Dihapus", vbInformation + vbOKOnly, "Informasi"
Form_Load
End If
End If
End Sub

Private Sub Command5_Click()
If Text1 = "" Then
MsgBox "Data Belum Dipilih", vbCritical + vbOKOnly, "Peringatan"
Else
If Check1.Value = 1 Then
MsgBox "Masih Aktif", vbCritical + vbOKOnly, "Informasi"
Else
If Command5.Caption = "&Perpanjangan" Then
Command5.Caption = "&Selesai"
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
DTPicker1.Enabled = True
Check3.Enabled = True
Check4.Enabled = True
Text7.Enabled = True
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Check3 = 0
Check4 = 0
Else
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7 = "" Or DTPicker1.Value = "" Or DTPicker2.Value = "" Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
ElseIf Check3.Value = 0 And Check4.Value = 0 Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
ElseIf DTPicker1.Value = 0 And DTPicker2.Value = 0 Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
Else
Adodc1.Recordset!NoRegister = Text1.Text
Adodc1.Recordset!NamaTim = Text2.Text
Adodc1.Recordset!AtasNama = Text3.Text
Adodc1.Recordset!alamat = Text4.Text
Adodc1.Recordset!telp = Text5.Text
Adodc1.Recordset!TanggalDaftar = DTPicker1.Value
Adodc1.Recordset!TanggalHangus = DTPicker2.Value
Adodc1.Recordset!Biaya = Text6.Text
Adodc1.Recordset!KodeJenisMember = Text7.Text
Adodc1.Recordset!Berbatas = lt
If DTPicker2 > Now Then
Adodc1.Recordset!isMember = "-1"
ElseIf DTPicker2 < Now Then
Adodc1.Recordset!isMember = "0"
End If
Adodc1.Recordset.Update
Adodc1.Recordset.Requery
Form_Load
MsgBox "Data Diperpanjang", vbInformation + vbOKOnly, "Informasi"
End If
End If
End If
End If
End Sub

Private Sub DataGrid1_Click()
If Not Adodc1.Recordset.EOF = False Then
MsgBox "Data Masih Kosong", vbInformation + vbOKOnly, "Informasi"
Else
Text1.Text = Adodc1.Recordset!NoRegister
Text2.Text = Adodc1.Recordset!NamaTim
Text3.Text = Adodc1.Recordset!AtasNama
Text4.Text = Adodc1.Recordset!alamat
Text5.Text = Adodc1.Recordset!telp
Text6.Text = Adodc1.Recordset!Biaya
Text7.Text = Adodc1.Recordset!KodeJenisMember
DTPicker1.Value = Adodc1.Recordset!TanggalDaftar
DTPicker2.Value = Adodc1.Recordset!TanggalHangus

If Adodc1.Recordset!Berbatas = -1 Then
Check3.Value = 1
ElseIf Adodc1.Recordset!Berbatas = 0 Then
Check4.Value = 1
End If

If Adodc1.Recordset!isMember = -1 Then
Check1.Value = 1
Command5.Enabled = False
ElseIf Adodc1.Recordset!isMember = 0 Then
Check2.Value = 1
Command5.Enabled = True
End If

Adodc2.RecordSource = "select * from JenisMember where KodeJenisMember='" & Text7 & "'"
Adodc2.Refresh
Text8 = Adodc2.Recordset!NamaJenisMember
Text9 = Adodc2.Recordset!JumlahBulan

End If
End Sub

Private Sub DTPicker1_Change()
DTPicker2.Value = DTPicker1
End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text7.SetFocus
End If
End Sub

Private Sub Form_Load()
Call BukaDB
Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "select NoRegister,NamaTim,AtasNama,Alamat,Telp,KodeJenisMember,TanggalDaftar,TanggalHangus,Biaya,Berbatas,isMember from Member order by NoRegister"
Adodc1.Refresh

Adodc2.ConnectionString = Koneksi
Adodc2.RecordSource = "select * from JenisMember"
Adodc2.Refresh

Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = False
Command1.Caption = "&Baru"
Command3.Caption = "&Ubah"
Command5.Caption = "&Perpanjangan"
DataGrid1.Enabled = True
mati
kosong
If Not Adodc1.Recordset.EOF = False Then
Adodc1.Refresh
Else
Call seleksi1
End If
Adodc1.Refresh
End Sub
Sub seleksi1()
Call BukaDB
Dim rsseleksi As ADODB.Recordset
Set rsseleksi = New ADODB.Recordset
rsseleksi.Open "select * from Member", Koneksi
rsseleksi.MoveFirst
Do Until rsseleksi.EOF
If rsseleksi!TanggalHangus <= Now Then
Dim sel As String
sel = "Update Member set isMember='0' where NoRegister='" & rsseleksi!NoRegister & "'"
Koneksi.Execute (sel)
rsseleksi.MoveNext
Else
rsseleksi.MoveNext
End If
Loop
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
DTPicker1.Enabled = False
DTPicker2.Enabled = False
'Check1.Enabled = False
'Check2.Enabled = False
Check3.Enabled = False
Check4.Enabled = False
End Sub
Sub hidup()
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text7.Enabled = True
DTPicker1.Enabled = True
Check3.Enabled = True
Check4.Enabled = True
End Sub

Sub kosong()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
DTPicker1.Value = Now
DTPicker2.Value = DTPicker1
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
DTPicker1.SetFocus
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If Check4 = 1 Then
frmlihatjenismember.Show vbModal, Me
ElseIf Check3 = 1 Then
frmlihatjenismember2.Show vbModal, Me
Else
MsgBox "Belum Dipilih", vbCritical + vbOKOnly, "Peringatan"
End If
End Sub
