VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.OCX"
Begin VB.Form frmlihatbaranglagi 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8790
   Icon            =   "frmlihatbaranglagi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9763
      _Version        =   262144
      Locked          =   -1  'True
      PaneTree        =   "frmlihatbaranglagi.frx":9E4A
      Begin Threed.SSPanel SSPanel2 
         Height          =   1845
         Left            =   30
         TabIndex        =   1
         Top             =   2355
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   3254
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   375
            Left            =   600
            Top             =   -360
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
            Bindings        =   "frmlihatbaranglagi.frx":9EBC
            Height          =   1815
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   3201
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   24
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   11.25
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
         Height          =   2235
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   3942
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Frame Frame2 
            BackColor       =   &H8000000D&
            Caption         =   "Cari Data"
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
            Left            =   4200
            TabIndex        =   12
            Top             =   960
            Width           =   4095
            Begin VB.TextBox Text5 
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   1800
               TabIndex        =   14
               Top             =   240
               Width           =   1695
            End
            Begin VB.TextBox Text6 
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   1800
               TabIndex        =   13
               Top             =   720
               Width           =   1695
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Kode barang :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   240
               TabIndex        =   16
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Nama barang :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   240
               TabIndex        =   15
               Top             =   720
               Width           =   1455
            End
         End
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5520
            TabIndex        =   11
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5520
            TabIndex        =   9
            Top             =   120
            Width           =   2175
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            TabIndex        =   7
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            TabIndex        =   5
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Stok    :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4560
            TabIndex        =   10
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4560
            TabIndex        =   8
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Barang :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   6
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Barang :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   4
            Top             =   120
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1215
         Left            =   30
         TabIndex        =   3
         Top             =   4290
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   2143
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CommandButton Command1 
            Caption         =   "Batal"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   7440
            Picture         =   "frmlihatbaranglagi.frx":9ED1
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   6240
            Picture         =   "frmlihatbaranglagi.frx":BB9B
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   120
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmlihatbaranglagi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "Data Tidak Ada", vbCritical + vbOKOnly, "Peringatan"
Else
frmpenjualanlangsung.Text3 = Adodc1.Recordset!KodeBarang
frmpenjualanlangsung.Text5 = Adodc1.Recordset!HargaJual
Unload Me
End If
End Sub

Private Sub DataGrid1_Click()
If Not Adodc1.Recordset.EOF = False Then
MsgBox "Data Masih Kosong", vbInformation + vbOKOnly, "Informasi"
Else
'Text1.Text = Adodc1.Recordset!KodeBarang
'Text2.Text = Adodc1.Recordset!NamaBarang
'Text3.Text = Adodc1.Recordset!HargaJual
'Text4.Text = Adodc1.Recordset!stok
Text1 = DataGrid1.Columns(0)
Text2 = DataGrid1.Columns(1)
Text3 = DataGrid1.Columns(2)
Text4 = DataGrid1.Columns(3)
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
Text4.Enabled = False
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Text5_Change()
Adodc1.RecordSource = "Select KodeBarang,NamaBarang,HargaBeli,HargaJual,Stok from MstBarang where KodeBarang like '%" & TandaPetik(Text5) & "%'"
Adodc1.Refresh
If Not Adodc1.Recordset.EOF Then
    With DataGrid1
    Set .DataSource = Adodc1
        .Refresh
    End With
End If
End Sub

Private Sub Text6_Change()
Adodc1.RecordSource = "Select KodeBarang,NamaBarang,HargaBeli,HargaJual,Stok from MstBarang where NamaBarang like '%" & TandaPetik(Text6) & "%'"
Adodc1.Refresh
If Not Adodc1.Recordset.EOF Then
    With DataGrid1
    Set .DataSource = Adodc1
        .Refresh
    End With
End If
End Sub
