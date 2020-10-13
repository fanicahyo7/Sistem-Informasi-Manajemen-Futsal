VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.OCX"
Begin VB.Form frmlihatbarang 
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8970
   ControlBox      =   0   'False
   Icon            =   "frmlihatbarang.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   10610
      _Version        =   262144
      Locked          =   -1  'True
      PaneTree        =   "frmlihatbarang.frx":9E4A
      Begin Threed.SSPanel SSPanel3 
         Height          =   1335
         Left            =   30
         TabIndex        =   3
         Top             =   4650
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   2355
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CommandButton SSCommand2 
            Caption         =   "OK"
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
            Left            =   6000
            Picture         =   "frmlihatbarang.frx":9EBC
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton SSCommand1 
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
            Left            =   7440
            Picture         =   "frmlihatbarang.frx":BB86
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   120
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2070
         Left            =   30
         TabIndex        =   2
         Top             =   2490
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   3651
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmlihatbarang.frx":D850
            Height          =   2055
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   3625
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
         Height          =   2370
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   4180
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Frame Frame2 
            BackColor       =   &H8000000D&
            Caption         =   "Cari Data"
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
            Left            =   4560
            TabIndex        =   10
            Top             =   1080
            Width           =   4215
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
               Height          =   360
               Left            =   1560
               TabIndex        =   12
               Top             =   240
               Width           =   2295
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
               Height          =   360
               Left            =   1560
               TabIndex        =   11
               Top             =   720
               Width           =   2295
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Kode Barang :"
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
               Left            =   120
               TabIndex        =   14
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Nama Barang :"
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
               Left            =   120
               TabIndex        =   13
               Top             =   720
               Width           =   1455
            End
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
            Height          =   360
            Left            =   6120
            TabIndex        =   9
            Top             =   480
            Width           =   1815
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
            Height          =   360
            Left            =   2040
            TabIndex        =   7
            Top             =   960
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
            Height          =   360
            Left            =   2040
            TabIndex        =   6
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Stok :"
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
            Left            =   5040
            TabIndex        =   8
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Barang :"
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
            Left            =   480
            TabIndex        =   5
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Barang  :"
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
            Left            =   480
            TabIndex        =   4
            Top             =   480
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "frmlihatbarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataGrid1_Click()
If Not Adodc1.Recordset.EOF = False Then
MsgBox "Data Masih Ksosong", vbInformation + vbOKOnly, "Informasi"
Else
'Text1.Text = Adodc1.Recordset!KodeBarang
'Text2.Text = Adodc1.Recordset!NamaBarang
'Text3.Text = Adodc1.Recordset!stok
Text1 = DataGrid1.Columns(0)
Text2 = DataGrid1.Columns(1)
Text3 = DataGrid1.Columns(4)
End If
End Sub

Private Sub Form_Load()
With Me
.Top = (Screen.Height / 2) - (Me.Height / 2)
.Left = (Screen.Width / 2) - (Me.Width / 2)
End With

Call BukaDB
Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "select KodeBarang,NamaBarang,HargaBeli,HargaJual,Stok from MstBarang order by KodeBarang"
Adodc1.Refresh


Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub SSCommand1_Click()
Unload Me
End Sub

Private Sub SSCommand2_Click()
If Text1 = "" Then
MsgBox "Data Tidak Ada", vbCritical + vbOKOnly, "Peringatan"
Else
frmrevisi.Text3.Text = Adodc1.Recordset!KodeBarang
frmrevisi.Text4.Text = Adodc1.Recordset!stok
frmrevisi.Text5.Text = Adodc1.Recordset!stok
Unload Me
End If
End Sub

Private Sub Text4_Change()
Adodc1.RecordSource = "Select KodeBarang,NamaBarang,HargaBeli,HargaJual,Stok from MstBarang where KodeBarang like '%" & TandaPetik(Text4) & "%'"
Adodc1.Refresh
If Not Adodc1.Recordset.EOF Then
    With DataGrid1
    Set .DataSource = Adodc1
        .Refresh
    End With
End If
End Sub

Private Sub Text5_Change()
Adodc1.RecordSource = "Select KodeBarang,NamaBarang,HargaBeli,HargaJual,Stok from MstBarang where NamaBarang like '%" & TandaPetik(Text5) & "%'"
Adodc1.Refresh
If Not Adodc1.Recordset.EOF Then
    With DataGrid1
    Set .DataSource = Adodc1
        .Refresh
    End With
End If
End Sub
