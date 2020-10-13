VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.OCX"
Begin VB.Form frmbooking 
   BorderStyle     =   0  'None
   Caption         =   "Booking"
   ClientHeight    =   9825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20400
   Icon            =   "frmbooking.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   20400
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9825
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20400
      _ExtentX        =   35983
      _ExtentY        =   17330
      _Version        =   262144
      AutoSize        =   1
      Locked          =   -1  'True
      PaneTree        =   "frmbooking.frx":9E4A
      Begin Threed.SSPanel SSPanel3 
         Height          =   1665
         Left            =   30
         TabIndex        =   3
         Top             =   8130
         Width           =   20340
         _ExtentX        =   35878
         _ExtentY        =   2937
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CommandButton Command5 
            Caption         =   "Cetak"
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
            Left            =   5520
            Picture         =   "frmbooking.frx":9EBC
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   240
            Width           =   1095
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
            Height          =   975
            Left            =   4080
            Picture         =   "frmbooking.frx":13D06
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   240
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
            Height          =   975
            Left            =   2760
            Picture         =   "frmbooking.frx":159D0
            Style           =   1  'Graphical
            TabIndex        =   30
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
            Left            =   1440
            Picture         =   "frmbooking.frx":1769A
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   240
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
            Height          =   975
            Left            =   120
            Picture         =   "frmbooking.frx":19364
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   240
            Width           =   1095
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   3825
         Left            =   30
         TabIndex        =   2
         Top             =   4215
         Width           =   20340
         _ExtentX        =   35878
         _ExtentY        =   6747
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSAdodcLib.Adodc Adodc3 
            Height          =   330
            Left            =   6720
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
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
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
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   375
            Left            =   5280
            Top             =   4320
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
            Height          =   330
            Left            =   9360
            Top             =   4440
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
            RecordSource    =   "select * from TrBooking"
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
            Bindings        =   "frmbooking.frx":1B02E
            Height          =   3855
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   18135
            _ExtentX        =   31988
            _ExtentY        =   6800
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
            Caption         =   "Booking"
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
         Height          =   4095
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   20340
         _ExtentX        =   35878
         _ExtentY        =   7223
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
         Begin MSAdodcLib.Adodc Adodc6 
            Height          =   330
            Left            =   6960
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
            Caption         =   "Adodc6"
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
         Begin MSAdodcLib.Adodc Adodc5 
            Height          =   330
            Left            =   7080
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
         Begin VB.TextBox Text10 
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
            Left            =   10200
            TabIndex        =   40
            Top             =   2040
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   495
            Left            =   10200
            TabIndex        =   39
            Top             =   1440
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   171048962
            CurrentDate     =   41748
         End
         Begin VB.TextBox Text9 
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
            Left            =   10200
            TabIndex        =   38
            Top             =   960
            Width           =   2295
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   10200
            TabIndex        =   37
            Top             =   480
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   3840
            TabIndex        =   36
            Top             =   1920
            Width           =   2295
            _ExtentX        =   4048
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
            Format          =   171048961
            CurrentDate     =   41670
         End
         Begin MSAdodcLib.Adodc Adodc4 
            Height          =   330
            Left            =   14400
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
         Begin VB.TextBox Text6 
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
            Left            =   10200
            TabIndex        =   34
            Top             =   2520
            Width           =   2295
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
            Left            =   3840
            TabIndex        =   32
            Top             =   960
            Width           =   2295
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H8000000D&
            Caption         =   "Tidak"
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
            Left            =   16320
            TabIndex        =   26
            Top             =   1320
            Width           =   855
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H8000000D&
            Caption         =   "Ya"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   15600
            TabIndex        =   25
            Top             =   1320
            Width           =   615
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H8000000D&
            Caption         =   "Tidak"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   16320
            TabIndex        =   23
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H8000000D&
            Caption         =   "Ya"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   15600
            TabIndex        =   22
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   15600
            TabIndex        =   20
            Top             =   1800
            Width           =   2295
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3840
            TabIndex        =   16
            Top             =   2880
            Width           =   2295
         End
         Begin VB.TextBox Text5 
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
            Left            =   3840
            TabIndex        =   15
            Top             =   1440
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   3840
            TabIndex        =   13
            Top             =   2400
            Width           =   2295
            _ExtentX        =   4048
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
            Format          =   171048961
            CurrentDate     =   41656
         End
         Begin VB.TextBox Text3 
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
            Left            =   3840
            TabIndex        =   12
            Top             =   3360
            Width           =   2295
         End
         Begin VB.TextBox Text1 
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
            Left            =   3840
            TabIndex        =   10
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Shift      :"
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
            TabIndex        =   35
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga               :"
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
            TabIndex        =   33
            Top             =   2520
            Width           =   1815
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Jam   :"
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
            TabIndex        =   27
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Pembatalan        :"
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
            Left            =   13440
            TabIndex        =   24
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Status Booking :"
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
            Left            =   13440
            TabIndex        =   21
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "DP (min 50%)     :"
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
            Left            =   13440
            TabIndex        =   19
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Selesai    :"
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
            TabIndex        =   18
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Mulai       :"
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
            TabIndex        =   17
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Atas Nama                    :"
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
            Left            =   1080
            TabIndex        =   14
            Top             =   1440
            Width           =   2775
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Lapangan         :"
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
            Left            =   1080
            TabIndex        =   11
            Top             =   3360
            Width           =   2655
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Register Member :"
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
            Left            =   1080
            TabIndex        =   9
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Booking        :"
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
            Left            =   1080
            TabIndex        =   8
            Top             =   2400
            Width           =   2655
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Transaksi     :"
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
            Left            =   1080
            TabIndex        =   7
            Top             =   1920
            Width           =   2535
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Lapangan          :"
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
            Left            =   1080
            TabIndex        =   6
            Top             =   2880
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Booking                 :"
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
            Left            =   1080
            TabIndex        =   5
            Top             =   480
            Width           =   2655
         End
      End
   End
End
Attribute VB_Name = "frmbooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sb As Boolean
Dim pb As Boolean
Dim anu As String

Private Sub Combo1_Click()
Adodc2.RecordSource = "select * from MstLapangan where KodeLapangan='" & Combo1 & "'"
Adodc2.Refresh
Text3 = Adodc2.Recordset!NamaLapangan
End Sub

Private Sub Combo2_Click()
Adodc3.RecordSource = "select * from Member where NoRegister='" & Text4 & "'"
Adodc3.Refresh
Adodc4.RecordSource = "select * from MstShift where KodeShift='" & Combo2 & "'"
Adodc4.Refresh
Text9 = Adodc4.Recordset!JamMulai
DTPicker3.Enabled = True
If Text4 = "" Then
Text6 = Adodc4.Recordset!Harga
Else
With Adodc3.Recordset
 If Not .RecordCount > 0 Then
 MsgBox "Kode Member Tidak Terdaftar", vbCritical + vbOKOnly, "Peringatan"
 Text4 = ""
Text6 = Adodc4.Recordset!Harga
Else
Text6 = Adodc4.Recordset!HargaMember
Text5 = Adodc3.Recordset!AtasNama
End If
End With
End If
End Sub

Private Sub Combo3_Click()
Adodc3.RecordSource = "select * from Member where NoRegister='" & Text4 & "'"
Adodc3.Refresh
Adodc4.RecordSource = "select * from MstShift where KodeShift='" & Combo3 & "'"
Adodc4.Refresh
Text9 = Adodc4.Recordset!JamMulai
DTPicker3.Enabled = True

With Adodc3.Recordset
 If Not .RecordCount > 0 Then
 MsgBox "Kode Member Tidak Terdaftar", vbCritical + vbOKOnly, "Peringatan"
 Text4 = ""
Text6 = Adodc4.Recordset!Harga
Else
Text6 = Adodc4.Recordset!HargaMember
End If
End With
End Sub

Private Sub Command1_Click()
If Command1.Caption = "&Baru" Then
Command1.Caption = "&Simpan"
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
hidup
kosong
DataGrid1.Enabled = False
Call autonumber
Else
If Text1 = "" Or Text3 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text9 = "" Or Text10 = "" Or Combo1 = "" Or DTPicker1 = 0 Or DTPicker2 = 0 Or DTPicker3 = 0 Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
Else
If Text7 < anu Then
MsgBox "Tidak Boleh Kurang 50% dari Harga", vbCritical + vbOKOnly, "Peringatan"
Text7.SetFocus
'Else
'Adodc3.RecordSource = "select * from Member where NoRegister='" & Text4 & "'"
'Adodc3.Refresh
'With Adodc3.Recordset
' If Not .RecordCount > 0 Then
' MsgBox "Kode Member Tidak Terdaftar", vbCritical + vbOKOnly, "Peringatan"
' Text4 = ""
Else
If Text10 <= 0 Then
MsgBox "Jumlah Jam Tidak Boleh <= 0", vbCritical + vbOKOnly, "Peringatan"
Else
Adodc5.RecordSource = "select * from TrBooking where TanggalBooking='" & Format(DTPicker2, "YYYY/MM/DD") & "' and KodeLapangan ='" & Combo1 & "'"
Adodc5.Refresh

'Adodc3.RecordSource = "select * from Member where NoRegister='" & Text4 & "'"
'Adodc3.Refresh
'With Adodc3.Recordset

If Not Adodc5.Recordset.EOF Then
Adodc5.Recordset.Filter = "KodeShift='" & Combo2 & "'"
End If
If Not Adodc5.Recordset.EOF Then
MsgBox "Data Sudah Ada", vbCritical + vbOKOnly, "Peringatan"
'Dim cari As String
'Dim cari1 As String
'Dim cari2 As String
'cari = "TanggalBooking='" & Format(DTPicker2, "YYYY/MM/DD") & "'"
'cari1 = "KodeLapangan='" & Combo1 & "'"
'cari2 = "KodeShift='" & Combo2 & "'"
'With Adodc6.Recordset
'.Find cari And cari1 And cari2
'If Not .EOF Then
'MsgBox "Data Sudah Ada", vbCritical + vbOKOnly, "Peringatan"
Else
Adodc6.Recordset.AddNew
Adodc6.Recordset!NoBooking = Text1
Adodc6.Recordset!Kodelapangan = Combo1
If Text4 = "" Then
Text4 = "-"
End If
Adodc6.Recordset!KodeShift = Combo2
Adodc6.Recordset!NoRegister = Text4
Adodc6.Recordset!AtasNama = Text5
'Else
'With Adodc3.Recordset
' If Not .RecordCount > 0 Then
' Text4 = "-"
' Adodc6.Recordset!NoRegister = Text4
'Adodc6.Recordset!KodeShift = Combo2
'Adodc6.Recordset!AtasNama = Text5
'Else
'Adodc6.Recordset!KodeShift = Combo2
'Adodc6.Recordset!NoRegister = Text4
'Text5 = Adodc6.Recordset!AtasNama
'Adodc6.Recordset!AtasNama = Text5
'End If
'End With
'End If
Adodc6.Recordset!Tanggal = Format(DTPicker1, "YYYY/MM/DD")
Adodc6.Recordset!TanggalBooking = Format(DTPicker2, "YYYY/MM/DD")
Adodc6.Recordset!DP = Text7
Adodc6.Recordset!StatusBooking = sb
Adodc6.Recordset!Pembatalan = pb
Adodc6.Recordset!JamMulai = Text9
Adodc6.Recordset!JamSelesai = Format(DTPicker3, "HH:MM:SS")
Adodc6.Recordset!Harga = Text6
Adodc6.Recordset.Update
Adodc6.Recordset.Requery
MsgBox "Data Berhasil Disimpan", vbInformation + vbOKOnly, "Informasi"

Dim strcon As String
Dim strsql As String
Dim lokasidatabase As String
On Error Resume Next
strsql = "select * from TrBooking where NoBooking='" & Text1 & "'"

With buktibooking
.ado.ConnectionString = Koneksi
.ado.Source = strsql
End With
frmMain.TampilkanForm "buktibooking"

Command1.Caption = "&Baru"
tombol
kosong
mati
DataGrid1.Enabled = True
Combo1.Clear
Combo2.Clear
Combo3.Clear
Form_Load
End If
'End With
End If
End If
'End With
End If
End If
End Sub

Private Sub Command2_Click()
kosong
Command1.Caption = "&Baru"
Command3.Caption = "&Ubah"
tombol
mati
DataGrid1.Enabled = True
End Sub

Private Sub Command3_Click()
If Text1 = "" Then
MsgBox "Pilih Dahulu Yang Akan Diubah", vbCritical + vbOKOnly, "Peringatan"
Else
If Command3.Caption = "&Ubah" Then
Command3.Caption = "&Simpan"
Command2.Enabled = True
Command1.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
hidup
DataGrid1.Enabled = False
Else
If Text1 = "" Or Text3 = "" Or Text6 = "" Or Text7 = "" Or Text9 = "" Or Text10 = "" Or Combo1 = "" Or DTPicker1 = 0 Or DTPicker2 = 0 Or DTPicker3 = 0 Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
Else
Adodc5.RecordSource = "select * from TrBooking where TanggalBooking like '" & DTPicker2 & "' and KodeLapangan like '" & Combo1 & "'"
Adodc5.Refresh

Adodc3.RecordSource = "select * from Member where NoRegister='" & Text4 & "'"
Adodc3.Refresh
With Adodc3.Recordset
If Not .EOF Then
Adodc5.Recordset.Filter = "KodeShift='" & Combo2 & "'"
End If
If Not Adodc5.Recordset.EOF Then
MsgBox "Data Sudah Ada", vbCritical + vbOKOnly, "Peringatan"
Else
Adodc6.Recordset!NoBooking = Text1
Adodc6.Recordset!Kodelapangan = Combo1
If Text4 = "" Then
Text4 = "-"
Adodc6.Recordset!KodeShift = Combo2
Adodc6.Recordset!NoRegister = Text4
Adodc6.Recordset!AtasNama = Text5
Else
With Adodc3.Recordset
 If Not .RecordCount > 0 Then
 Text4 = "-"
 Adodc6.Recordset!NoRegister = Text4
Adodc6.Recordset!KodeShift = Combo2
Adodc6.Recordset!AtasNama = Text5
Else
Adodc6.Recordset!KodeShift = Combo2
Adodc6.Recordset!NoRegister = Text4
Text5 = Adodc3.Recordset!AtasNama
Adodc6.Recordset!AtasNama = Text5
End If
End With
End If
Adodc6.Recordset!JamMulai = Text9
Adodc6.Recordset!JamSelesai = DTPicker3
Adodc6.Recordset!Harga = Text6
Adodc6.Recordset!StatusBooking = sb
Adodc6.Recordset!Pembatalan = pb
Adodc6.Recordset.Update
Adodc6.Recordset.Requery
MsgBox "Data Berhasil Diubah", vbInformation + vbOKOnly, "Informasi"
Command3.Caption = "&Ubah"
tombol
kosong
mati
DataGrid1.Enabled = True
End If
End With
End If
End If
End If
End Sub

Private Sub Command4_Click()
If Not Adodc6.Recordset.EOF = False Then
MsgBox "Data Kosong", vbCritical + vbOKOnly, "Peringatan"
Else
If MsgBox("Apakah Anda Yakin Akan Menghapus ?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
Adodc6.Recordset.Delete
MsgBox "Data Berhasil Dihapus", vbInformation + vbOKOnly, "Informasi"
tombol
kosong
End If
End If
End Sub

Private Sub Command5_Click()
If Text1 = "" Then
MsgBox "Pilih Data Terlebih Dahulu", vbCritical + vbOKOnly, "Peringatan"
Else
Dim strcon As String
Dim strsql As String
Dim lokasidatabase As String

On Error Resume Next

strsql = "select * from TrBooking where NoBooking='" & Text1 & "'"

With buktibooking
.ado.ConnectionString = Koneksi
.ado.Source = strsql
End With
frmMain.TampilkanForm "buktibooking"
End If
End Sub

Private Sub DataGrid1_Click()
If Not Adodc6.Recordset.EOF = False Then
MsgBox "Data Masih Kosong", vbInformation + vbOKOnly, "Informasi"
Else
Text1 = Adodc6.Recordset!NoBooking
Combo1 = Adodc6.Recordset!Kodelapangan
DTPicker1 = Adodc6.Recordset!Tanggal
DTPicker2 = Adodc6.Recordset!TanggalBooking
Text7 = Adodc6.Recordset!DP
Text9 = Adodc6.Recordset!JamMulai
DTPicker3 = Adodc6.Recordset!JamSelesai
Combo2 = Adodc6.Recordset!KodeShift
Text4 = Adodc6.Recordset!NoRegister

If Adodc6.Recordset!StatusBooking = -1 Then
Check1.Value = 1
ElseIf Adodc6.Recordset!StatusBooking = 0 Then
Check2 = 1
End If

If Adodc6.Recordset!Pembatalan = -1 Then
Check3.Value = 1
ElseIf Adodc6.Recordset!Pembatalan = 0 Then
Check4.Value = 1
End If

Text10 = totalWaktu(DTPicker3, Text9)


Adodc3.RecordSource = "select * from Member where NoRegister='" & Text4 & "'"
Adodc3.Refresh
With Adodc3.Recordset
 If .RecordCount > 0 Then
 Text5 = !AtasNama
 Else
 Text5 = Adodc6.Recordset!AtasNama
End If
End With

Adodc4.RecordSource = "select * from MstShift where KodeShift='" & Combo2 & "'"
Adodc4.Refresh
If Text4 = "-" Then
Text6 = Adodc4.Recordset!Harga * (Text10)
Else
Text6 = Adodc4.Recordset!HargaMember * (Text10)
End If

Adodc2.RecordSource = "select * from MstLapangan where KodeLapangan='" & Combo1 & "'"
Adodc2.Refresh
Text3 = Adodc2.Recordset!NamaLapangan

End If
End Sub
Private Sub Check1_Click()
If Check1.Value = 1 Then
Check2.Value = 0
Check4 = 1
sb = "True"
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Check1.Value = 0
sb = "False"
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
Check4.Value = 0
Check2.Value = 1
pb = "True"
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
Check3.Value = 0
pb = "False"
End If
End Sub

Private Function totalWaktu(jamAwal As Variant, jamAkhir As Variant) As String

Dim menitAkhir, menitAwal, jumlahmenit As Long

menitAwal = (Hour(jamAwal) * 3600) - (Minute(jamAwal) * 60)
menitAkhir = (Hour(jamAkhir) * 3600) - (Minute(jamAkhir) * 60)
jumlahmenit = jumlahmenit - (menitAkhir - menitAwal)

totalWaktu = Format(Str(Int((Int((jumlahmenit / 3600)) Mod 24))), "00")
End Function

Private Sub DTPicker3_Change()
Text10 = totalWaktu(DTPicker3, Text9)

Adodc3.RecordSource = "select * from Member where NoRegister='" & Text4 & "'"
Adodc3.Refresh
Adodc4.RecordSource = "select * from MstShift where KodeShift='" & Combo2 & "'"
Adodc4.Refresh
Text9 = Adodc4.Recordset!JamMulai
DTPicker3.Enabled = True
If Text4 = "" Then
Text6 = Adodc4.Recordset!Harga
Else
With Adodc3.Recordset
 If Not .RecordCount > 0 Then
Text6 = Adodc4.Recordset!Harga
Else
Text6 = Adodc4.Recordset!HargaMember
End If
End With
End If

Text6 = (Text10) * (Text6)
End Sub

Private Sub Form_Load()
Call BukaDB

Adodc6.ConnectionString = Koneksi
Adodc6.RecordSource = "select NoBooking,AtasNama,Tanggal,TanggalBooking,JamMulai,JamSelesai,KodeLapangan,DP,StatusBooking,Pembatalan,NoRegister,KodeShift,Harga from TrBooking"
Adodc6.Refresh

Adodc2.ConnectionString = Koneksi
Adodc2.RecordSource = "select * from MstBarang"
Adodc2.Refresh

Adodc3.ConnectionString = Koneksi
Adodc3.RecordSource = "select * from Member"
Adodc3.Refresh

Adodc4.ConnectionString = Koneksi
Adodc4.RecordSource = "select * from MstShift"
Adodc4.Refresh

Adodc5.ConnectionString = Koneksi
Adodc5.RecordSource = "select * from TrBooking"
Adodc5.Refresh

If Not Adodc6.Recordset.EOF Then
If Adodc6.Recordset!TanggalBooking > Now Then
Adodc6.Recordset!StatusBooking = 0
End If
End If


DTPicker1.Value = Now
DTPicker2.Value = Now
DTPicker1 = Format(DTPicker1, "MM/DD/yyyy")
DTPicker2 = Format(DTPicker2, "MM/DD/yyyy")

Adodc2.RecordSource = "select * from MstLapangan"
    Adodc2.Refresh
        Do While Not Adodc2.Recordset.EOF
            Combo1.AddItem Adodc2.Recordset!Kodelapangan
            Adodc2.Recordset.MoveNext
        Loop

Adodc4.RecordSource = "select * from MstShift"
    Adodc4.Refresh
        Do While Not Adodc4.Recordset.EOF
            Combo2.AddItem Adodc4.Recordset!KodeShift
            Adodc4.Recordset.MoveNext
            Loop
    tombol
Text5.Enabled = False
Text3.Enabled = False
Text6.Enabled = False
Text10.Enabled = False
Text9.Enabled = False
Combo2.Visible = True
Adodc6.Visible = False
mati
End Sub
Sub mati()
Text1.Enabled = False
Text4.Enabled = False
DTPicker1.Enabled = False
DTPicker2.Enabled = False
DTPicker3.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Text7.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Check4.Enabled = False
Text5.Enabled = False
End Sub
Sub hidup()
Text5.Enabled = True
Text4.Enabled = True
DTPicker1.Enabled = True
DTPicker2.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Text7.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Check4.Enabled = True
End Sub
Sub tombol()
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
End Sub
Sub kosong()
Text1 = ""
Text4 = ""
Text5 = ""
DTPicker1 = Now
DTPicker2 = Now
Text7 = ""
Combo1 = ""
Combo2 = ""
Text3 = ""
Combo3 = ""
Text9 = ""
DTPicker3 = "12:00:00 AM"
Text10 = ""
Text6 = ""
Check1 = 0
Check2 = 0
Check3 = 0
Check4 = 0
End Sub

Private Sub autonumber()
Call BukaDB
Dim Kode As String

Set rsuser = New ADODB.Recordset
    rsuser.Open "Select * From TrBooking Where NoBooking Like '%" & Format(Date, "ddMMyy") & "%' ORDER BY NoBooking desc", Koneksi

If rsuser.BOF Then
        Text1.Text = "BOK" & Format(Date, "ddMMyy") & "0001"
        Exit Sub
    Else
        rsuser.Requery
        
        If (rsuser.EOF Or rsuser.BOF) Then
            rsuser.MoveLast
        End If
        Kode = rsuser!NoBooking
        Kode = Val(Right(Kode, 4))
        Kode = Kode + 1
    End If
    
    If Val(Kode) < 10 Then
 Kode = "BOK" & Format(Date, "ddMMyy") & "000" & Kode
        Text1.Text = Kode
    ElseIf Val(Kode) < 100 Then
        Kode = "BOK" & Format(Date, "ddMMyy") & "00" & Kode
        Text1.Text = Kode
    ElseIf Val(Kode) < 1000 Then
        Kode = "BOK" & Format(Date, "ddMMyy") & "0" & Kode
        Text1.Text = Kode
    ElseIf Val(Kode) < 10000 Then
        Kode = "BOK" & Format(Date, "ddMMyy") & "" & Kode
        Text1.Text = Kode
    Else
        MsgBox "Kapasitas Tidak Memadai!", _
        vbInformation + vbOKOnly, "Perhatian"
        Kode = ""
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Adodc3.RecordSource = "select * from Member where NoRegister='" & TandaPetik(Text4) & "'"
Adodc3.Refresh
With Adodc3.Recordset
 If .RecordCount > 0 Then
 Text5 = !AtasNama
 Text9 = ""
 Text10 = ""
 Text6 = ""
 Else
 MsgBox "Kode Tidak Ada", vbCritical + vbOKOnly, "Peringatan"
 Text4 = ""
 Combo2 = ""
 Text5 = ""
 Text9 = ""
 Text10 = ""
 Text6 = ""
 End If
 End With
 End If
 
End Sub

Private Sub Text6_Change()
If Text6 = "" Then
Text6 = 0
Else
anu = Text6 * 50 / 100
Text7 = anu
End If
End Sub

Private Sub Text7_Change()
If Not IsNumeric(Text7) Then Text7 = "0"
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
anu = Text6 * 50 / 100
If Text7 < anu Then
MsgBox "Tidak Boleh Kurang 50% dari Harga", vbCritical + vbOKOnly, "Peringatan"
Text7.SetFocus
Else
Text7 = anu
End If
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
frmlihatshift.Show vbModal, Me
End Sub
