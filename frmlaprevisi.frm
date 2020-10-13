VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.OCX"
Begin VB.Form frmlaprevisi 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   LinkTopic       =   "Form2"
   ScaleHeight     =   3540
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
      PaneTree        =   "frmlaprevisi.frx":0000
      Begin Threed.SSPanel SSPanel2 
         Height          =   1320
         Left            =   30
         TabIndex        =   1
         Top             =   2160
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   2328
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
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
            Left            =   4440
            Picture         =   "frmlaprevisi.frx":0052
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
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
            Height          =   975
            Left            =   3120
            Picture         =   "frmlaprevisi.frx":1D1C
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   2040
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   3598
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
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
            TabIndex        =   6
            Top             =   600
            Width           =   1095
         End
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
            TabIndex        =   5
            Top             =   1080
            Width           =   1575
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
            TabIndex        =   7
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
            Format          =   115343361
            CurrentDate     =   41786
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   4200
            TabIndex        =   8
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
            Format          =   115343361
            CurrentDate     =   41786
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
            TabIndex        =   12
            Top             =   600
            Width           =   735
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
            TabIndex        =   11
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Laporan Pembelian"
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
            Left            =   1200
            TabIndex        =   10
            Top             =   120
            Width           =   1815
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
      End
   End
End
Attribute VB_Name = "frmlaprevisi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1 = False And Option2 = False Then
MsgBox "Pilih Yang Akan Ditampilkan", vbCritical + vbOKOnly, "Peringatan"
Else
frmMain.laptutup
frmMain.TampilkanForm "laprevisistok"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
With Me
.Top = (Screen.Height / 2) - (Me.Height / 2)
.Left = (Screen.Width / 2) - (Me.Width / 2)
End With
DTPicker1.Enabled = False
DTPicker2.Enabled = False
DTPicker1.Value = Now
DTPicker2.Value = Now
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

