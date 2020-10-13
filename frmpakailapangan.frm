VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.OCX"
Begin VB.Form frmpakailapangan 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   19500
   Icon            =   "frmpakailapangan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   19500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19500
      _ExtentX        =   34396
      _ExtentY        =   16536
      _Version        =   262144
      AutoSize        =   1
      Locked          =   -1  'True
      PaneTree        =   "frmpakailapangan.frx":9E4A
      Begin Threed.SSPanel SSPanel4 
         Height          =   4080
         Left            =   12030
         TabIndex        =   1
         Top             =   30
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   7197
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox Text16 
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
            Left            =   1800
            TabIndex        =   43
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Tambah"
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
            Left            =   4560
            Picture         =   "frmpakailapangan.frx":9EDC
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox Text12 
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
            Left            =   1800
            TabIndex        =   6
            Top             =   3360
            Width           =   2535
         End
         Begin VB.TextBox Text10 
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
            Left            =   1800
            TabIndex        =   5
            Top             =   2760
            Width           =   2535
         End
         Begin VB.TextBox Text11 
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
            Left            =   1800
            TabIndex        =   4
            Top             =   2160
            Width           =   2535
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
            Height          =   405
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1560
            Width           =   2535
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
            Height          =   405
            Left            =   1800
            TabIndex        =   2
            Top             =   960
            Width           =   2535
         End
         Begin MSAdodcLib.Adodc Adodc3 
            Height          =   330
            Left            =   360
            Top             =   4560
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
            Connect         =   $"frmpakailapangan.frx":BBA6
            OLEDBString     =   $"frmpakailapangan.frx":BC32
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "select * from MstBarang"
            Caption         =   "Adodc3"
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
            Left            =   600
            Top             =   4080
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
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Jual      :"
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
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Total                :"
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
            Left            =   120
            TabIndex        =   12
            Top             =   3360
            Width           =   1695
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah           :"
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
            Left            =   120
            TabIndex        =   11
            Top             =   2760
            Width           =   1695
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga              :"
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
            Left            =   120
            TabIndex        =   10
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Barang :"
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
            Left            =   120
            TabIndex        =   9
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Urut          :"
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
            Left            =   120
            TabIndex        =   8
            Top             =   960
            Width           =   1695
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   4080
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   11910
         _ExtentX        =   21008
         _ExtentY        =   7197
         _Version        =   262144
         BackColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   495
            Left            =   2880
            TabIndex        =   41
            Top             =   1080
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
            Format          =   112721921
            CurrentDate     =   41815
         End
         Begin MSAdodcLib.Adodc Adodc5 
            Height          =   330
            Left            =   3600
            Top             =   4080
            Width           =   1215
            _ExtentX        =   2143
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
            Caption         =   "Adodc5"
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
         Begin MSAdodcLib.Adodc Adodc4 
            Height          =   330
            Left            =   3000
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
            Caption         =   "Adodc4"
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
         Begin VB.TextBox Text15 
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
            Left            =   2880
            TabIndex        =   39
            Top             =   3480
            Width           =   2655
         End
         Begin VB.TextBox Text14 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2880
            TabIndex        =   37
            Top             =   2280
            Width           =   2655
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   9000
            TabIndex        =   21
            Top             =   2880
            Width           =   2655
         End
         Begin VB.TextBox Text13 
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
            Left            =   9000
            TabIndex        =   20
            Top             =   2280
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
            Height          =   405
            Left            =   9000
            TabIndex        =   19
            Top             =   1680
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
            Height          =   405
            Left            =   9000
            TabIndex        =   18
            Top             =   1080
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
            Height          =   405
            Left            =   9000
            TabIndex        =   17
            Top             =   480
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
            Height          =   405
            Left            =   2880
            TabIndex        =   16
            Top             =   2880
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
            Height          =   405
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   1680
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
            Height          =   405
            Left            =   2880
            TabIndex        =   14
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal                       :"
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
            Left            =   240
            TabIndex        =   40
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Lapangan        :"
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
            Left            =   240
            TabIndex        =   38
            Top             =   3480
            Width           =   2535
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Atas Nama                   :"
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
            Left            =   240
            TabIndex        =   36
            Top             =   2280
            Width           =   2535
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Grand Total Harga         :"
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
            Left            =   6120
            TabIndex        =   29
            Top             =   2880
            Width           =   2775
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Pembelian             :"
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
            Left            =   6120
            TabIndex        =   28
            Top             =   2280
            Width           =   2775
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Sewa Lapangan :"
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
            Left            =   6120
            TabIndex        =   27
            Top             =   1680
            Width           =   2775
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "DP                                      :"
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
            Left            =   6120
            TabIndex        =   26
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Mulai                        :"
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
            Left            =   6120
            TabIndex        =   25
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Lapangan         :"
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
            Left            =   240
            TabIndex        =   24
            Top             =   2880
            Width           =   2535
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Booking                :"
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
            Left            =   240
            TabIndex        =   23
            Top             =   1680
            Width           =   2535
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Pakai Lapangan :"
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
            Left            =   240
            TabIndex        =   22
            Top             =   480
            Width           =   2535
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1020
         Left            =   30
         TabIndex        =   30
         Top             =   8325
         Width           =   19440
         _ExtentX        =   34290
         _ExtentY        =   1799
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
            Left            =   1680
            Picture         =   "frmpakailapangan.frx":BCBE
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   120
            Width           =   975
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
            Height          =   975
            Left            =   360
            Picture         =   "frmpakailapangan.frx":D988
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Simpan"
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
            Picture         =   "frmpakailapangan.frx":F652
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   120
            Width           =   1095
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   4035
         Left            =   30
         TabIndex        =   34
         Top             =   4200
         Width           =   19440
         _ExtentX        =   34290
         _ExtentY        =   7117
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   330
            Left            =   9360
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
         Begin MSComctlLib.ListView ListView1 
            Height          =   3975
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   7011
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
   End
End
Attribute VB_Name = "frmpakailapangan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub SetLV()
With ListView1
    .View = lvwReport
    .Gridlines = True
    .MultiSelect = True
    .FullRowSelect = True
    .HotTracking = True
    .HoverSelection = True

    .ColumnHeaders.Add 1, , "No Urut", 1000
    .ColumnHeaders.Add 2, , "Kode Jual", 2100
    .ColumnHeaders.Add 3, , "Kode Pakai Lapangan", 2100
    .ColumnHeaders.Add 4, , "Kode Barang", 1500
    .ColumnHeaders.Add 5, , "Jumlah", 1000
    .ColumnHeaders.Add 6, , "Harga", 1500
    .ColumnHeaders.Add 7, , "Total", 1500
    .Width = 10800
End With
End Sub

Sub autokojul()
Call BukaDB
Dim Kode As String

Set rsuser = New ADODB.Recordset
    rsuser.Open "Select * From ItemPenjualan Where KodePenjualan Like '%" & Format(Date, "ddMMyy") & "%' ORDER BY KodePenjualan desc", Koneksi

If rsuser.BOF Then
        Text16.Text = "JUL" & Format(Date, "ddMMyy") & "0001"
        Exit Sub
    Else
        rsuser.Requery
        
        If (rsuser.EOF Or rsuser.BOF) Then
            rsuser.MoveLast
        End If
        Kode = rsuser!KodePenjualan
        Kode = Val(Right(Kode, 4))
        Kode = Kode + 1
    End If
    
    If Val(Kode) < 10 Then
 Kode = "JUL" & Format(Date, "ddMMyy") & "000" & Kode
        Text16.Text = Kode
    ElseIf Val(Kode) < 100 Then
        Kode = "JUL" & Format(Date, "ddMMyy") & "00" & Kode
        Text16.Text = Kode
    ElseIf Val(Kode) < 1000 Then
        Kode = "JUL" & Format(Date, "ddMMyy") & "0" & Kode
        Text16.Text = Kode
    ElseIf Val(Kode) < 10000 Then
        Kode = "JUL" & Format(Date, "ddMMyy") & "" & Kode
        Text16.Text = Kode
    Else
        MsgBox "Kapasitas Tidak Memadai!", _
        vbInformation + vbOKOnly, "Perhatian"
        Kode = ""
    End If
End Sub


Private Sub autonumber()
Call BukaDB
Dim Kode As String

Set rsuser = New ADODB.Recordset
    rsuser.Open "Select * From TrPakaiLapangan Where NoPakaiLapangan Like '%" & Format(Date, "ddMMyy") & "%' ORDER BY NoPakaiLapangan desc", Koneksi

If rsuser.BOF Then
        Text1.Text = "PKL" & Format(Date, "ddMMyy") & "0001"
        Exit Sub
    Else
        rsuser.Requery
        
        If (rsuser.EOF Or rsuser.BOF) Then
            rsuser.MoveLast
        End If
        Kode = rsuser!NoPakaiLapangan
        Kode = Val(Right(Kode, 4))
        Kode = Kode + 1
    End If
    
    If Val(Kode) < 10 Then
 Kode = "PKL" & Format(Date, "ddMMyy") & "000" & Kode
        Text1.Text = Kode
    ElseIf Val(Kode) < 100 Then
        Kode = "PKL" & Format(Date, "ddMMyy") & "00" & Kode
        Text1.Text = Kode
    ElseIf Val(Kode) < 1000 Then
        Kode = "PKL" & Format(Date, "ddMMyy") & "0" & Kode
        Text1.Text = Kode
    ElseIf Val(Kode) < 10000 Then
        Kode = "PKL" & Format(Date, "ddMMyy") & "" & Kode
        Text1.Text = Kode
    Else
        MsgBox "Kapasitas Tidak Memadai!", _
        vbInformation + vbOKOnly, "Perhatian"
        Kode = ""
    End If
End Sub

Private Sub Command1_Click()
Dim a As Integer
Command1.Enabled = False
Command6.Enabled = True
Command2.Enabled = True
kosong
Text1 = ""
autonumber
bukabooking
End Sub

Private Sub Command2_Click()
Text1 = ""
Text2 = ""
aturcommand
kosong
mati
ListView1.Enabled = False
ListView1.ListItems.Clear
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command5_Click()
Set rsuser = New ADODB.Recordset
rsuser.Open "select * from MstBarang where KodeBarang = '" & Text9.Text & "'", Koneksi
If Not rsuser.EOF Then
If Val(Text10.Text) > rsuser.Fields("Stok") Then
MsgBox "Maaf Stock Barang Tidak Cukup", vbInformation, "Informasi"
Else
Dim a As Integer
Dim lst As ListItem
    Set lst = ListView1.ListItems.Add()
    lst.Text = Text8
    lst.SubItems(1) = Text16
    lst.SubItems(2) = Text1
    lst.SubItems(3) = Text9
    lst.SubItems(4) = Text10
    lst.SubItems(5) = Text11
    lst.SubItems(6) = Text12
    Text8 = ""
    Text9 = ""
    Text10 = ""
    Text11 = ""
    Text12 = ""
    
a = MsgBox("Apakah Anda Akan Menambah Barang Lagi ?", vbQuestion + vbYesNo, "Konfirmasi")
If a = vbYes Then
Text9.SetFocus
Text8 = lst.Text + 1
Else
If a = vbNo Then
Command5.Enabled = False
Text9.Enabled = False
Command6.SetFocus
End If
End If

Dim i, tot
For i = 1 To ListView1.ListItems.Count
tot = Val(tot) + Val(ListView1.ListItems(i).SubItems(6))
Next
Text13.Text = tot
End If
End If
End Sub

Private Sub Command6_Click()
If Text2.Text = "" Or Text6 = "" Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
Else
frmtotalboju.Show vbModal, Me
End If
End Sub

Private Sub Form_Load()
Call BukaDB
Adodc2.ConnectionString = Koneksi
Adodc2.RecordSource = "select * from TrBooking"
Adodc2.Refresh

Adodc3.ConnectionString = Koneksi
Adodc3.RecordSource = "select * from MstBarang"
Adodc3.Refresh

Adodc4.ConnectionString = Koneksi
Adodc4.RecordSource = "select * from Mstlapangan"
Adodc4.Refresh

'Adodc5.ConnectionString = Koneksi
'Adodc5.RecordSource = "select * from ItemPenjualan"
'Adodc5.Refresh

Call SetLV
mati
Command1.Enabled = True
Command2.Enabled = False
Command6.Enabled = False
Command5.Enabled = False
DTPicker1.Value = Now
End Sub
Sub mati()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text7.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text13.Enabled = False
Text6.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text11.Enabled = False
Text10.Enabled = False
Text12.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
Text16.Enabled = False
DTPicker1.Enabled = False
End Sub

Sub aturcommand()
Command1.Enabled = True
Command2.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
End Sub
Sub kosong()
Text3 = ""
Text7 = ""
Text4 = ""
Text5 = ""
Text13 = ""
Text6 = ""
Text8 = ""
Text9 = ""
Text10 = ""
Text11 = ""
Text12 = ""
Text14 = ""
Text15 = ""
Text16 = ""
End Sub
Sub bukabooking()
Text2.Enabled = True
DTPicker1.Enabled = True
End Sub
Sub bukapenjualan()
Text9.Enabled = True
Text10.Enabled = True
End Sub

Private Sub Text10_Change()
If Not IsNumeric(Text10) Then Text10 = "0"
Text12.Text = Val(Text11) * Val(Text10)
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
Dim stk
Call BukaDB
Adodc3.RecordSource = "select * from MstBarang where KodeBarang ='" & Text9.Text & "'"
Adodc3.Refresh
With Adodc3.Recordset
If .RecordCount > 0 Then
stk = Adodc3.Recordset!stok - (Val(Text10))
If stk < 0 Then
MsgBox "Stok Tidak memenuhi Permintaan", vbCritical + vbOKOnly, "Peringatan"
Text10.Text = ""
End If
End If
End With
End Sub

Private Sub Text13_Change()
If Not IsNumeric(Text13) Then Text13 = "0"
Text6.Text = Val(Text6.Text) + Val(Text13.Text)
End Sub

Private Sub Text2_Change()
If Text2 = "" Then
Exit Sub
Else
Adodc2.RecordSource = "select * from TrBooking where NoBooking='" & Text2 & "'"
Adodc2.Refresh
Text14.Text = Adodc2.Recordset!AtasNama
End If
End Sub

Private Sub Text3_Change()
If Text3 = "" Then
Exit Sub
Else
Adodc4.RecordSource = "select * from MstLapangan where KodeLapangan='" & Text3 & "'"
Adodc4.Refresh
Text15 = Adodc4.Recordset!NamaLapangan
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
frmlihatbooking.Show vbModal, Me
End Sub

Private Sub Text4_Change()
If Not IsNumeric(Text4) Then Text4 = "0"
Text6.Text = Val(Text5.Text) - Val(Text4.Text)
End Sub

Private Sub Text5_Change()
If Not IsNumeric(Text5) Then Text5 = "0"
Text6.Text = Val(Text5.Text) + Val(Text13.Text) - Val(Text4.Text)
End Sub

Private Sub Text6_Change()
If Not IsNumeric(Text6) Then Text6 = "0"
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
frmmstbarangjual.Show vbModal, Me
End Sub


