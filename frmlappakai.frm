VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frmlappakai 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   LinkTopic       =   "Form2"
   ScaleHeight     =   3525
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   3510
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6191
      _Version        =   262144
      Locked          =   -1  'True
      PaneTree        =   "frmlappakai.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   2040
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   3598
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton Option2 
            BackColor       =   &H8000000D&
            Caption         =   "Berdasarkan Tanggal"
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
            Left            =   960
            TabIndex        =   3
            Top             =   1080
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H8000000D&
            Caption         =   "Semua"
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
            Left            =   960
            TabIndex        =   2
            Top             =   600
            Width           =   1095
         End
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   330
            Left            =   4560
            Top             =   2160
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
            Height          =   375
            Left            =   4320
            Top             =   2160
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
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   4200
            TabIndex        =   4
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   104529921
            CurrentDate     =   41786
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   4200
            TabIndex        =   5
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   104529921
            CurrentDate     =   41786
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Lihat :"
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
            TabIndex        =   9
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Laporan Pakai Lapangan"
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
            Left            =   600
            TabIndex        =   8
            Top             =   120
            Width           =   2895
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Sampai :"
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
            Left            =   3240
            TabIndex        =   7
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Dari :"
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
            Left            =   3240
            TabIndex        =   6
            Top             =   600
            Width           =   735
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1320
         Left            =   2115
         TabIndex        =   10
         Top             =   2160
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   2328
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
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
            Height          =   975
            Left            =   1080
            Picture         =   "frmlappakai.frx":0072
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
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
            Height          =   975
            Left            =   2400
            Picture         =   "frmlappakai.frx":1D3C
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1320
         Left            =   30
         TabIndex        =   13
         Top             =   2160
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   2328
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CommandButton Command3 
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
            Left            =   360
            Picture         =   "frmlappakai.frx":3A06
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   120
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmlappakai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim excel As New excel.Application
Private Sub Command1_Click()
If Option1 = False And Option2 = False Then
MsgBox "Pilih Yang Akan Ditampilkan", vbCritical + vbOKOnly, "Peringatan"
Else
frmMain.laptutup
frmMain.TampilkanForm "LapPakaiLapangan"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
anuexcel
End Sub
Sub anuexcel()
    Set excel = excel.Application
    excel.Workbooks.Add
    excel.Worksheets(1).Activate
    

    For i = 0 To DataEnvironment1.rsRPakailap.Fields.Count - 1
        excel.Worksheets(1).Cells(1) = "Data Penjualan"
        excel.Worksheets(1).Cells(2, i + 1) = DataEnvironment1.rsRPakailap.Fields(i).Name
    Next
    

    If DataEnvironment1.rsRPakailap.State = 0 Then DataEnvironment1.rsRPakailap.Open
    If DataEnvironment1.rsRPakailap.RecordCount > 0 Then DataEnvironment1.rsRPakailap.MoveFirst
        For i = 1 To DataEnvironment1.rsRPakailap.RecordCount
            For j = 0 To DataEnvironment1.rsRPakailap.Fields.Count - 1
                excel.Worksheets(1).Cells(i + 2, j + 1) = DataEnvironment1.rsRPakailap(j)
            Next
            DataEnvironment1.rsRPakailap.MoveNext
        Next
        
    excel.Columns.AutoFit
    excel.Visible = True
    excel.Workbooks(1).Saved = False
End Sub
Private Sub Form_Load()
With Me
.Top = (Screen.Height / 2) - (Me.Height / 2)
.Left = (Screen.Width / 2) - (Me.Width / 2)
End With
Adodc1.ConnectionString = Koneksi
Adodc2.ConnectionString = Koneksi
Option1.Value = False
Option2.Value = False
DTPicker1.Enabled = False
DTPicker2.Enabled = False
DTPicker1.Value = Format(Now, "yyyy/MM/DD")
DTPicker2.Value = Format(Now, "yyyy/MM/DD")
Command3.Visible = False
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
DTPicker1.Enabled = False
DTPicker2.Enabled = False
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
DTPicker1.Enabled = True
DTPicker2.Enabled = True
End If
End Sub
Sub mati()
Unload Me
End Sub

