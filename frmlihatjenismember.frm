VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.OCX"
Begin VB.Form frmlihatjenismember 
   BorderStyle     =   0  'None
   Caption         =   "lihatjenismember"
   ClientHeight    =   6765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9825
   Icon            =   "frmlihatjenismember.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6765
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   6765
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   11933
      _Version        =   262144
      AutoSize        =   1
      Locked          =   -1  'True
      PaneTree        =   "frmlihatjenismember.frx":9E4A
      Begin Threed.SSPanel SSPanel3 
         Height          =   1485
         Left            =   30
         TabIndex        =   3
         Top             =   5250
         Width           =   9765
         _ExtentX        =   17224
         _ExtentY        =   2619
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   1920
            Top             =   -360
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
            Left            =   8160
            Picture         =   "frmlihatjenismember.frx":9EBC
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   1335
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
            Height          =   1095
            Left            =   6600
            Picture         =   "frmlihatjenismember.frx":BB86
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2640
         Left            =   30
         TabIndex        =   2
         Top             =   2520
         Width           =   9765
         _ExtentX        =   17224
         _ExtentY        =   4657
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmlihatjenismember.frx":D850
            Height          =   2535
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   4471
            _Version        =   393216
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
            Caption         =   "Master Jenis Member"
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
         Height          =   2400
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   9765
         _ExtentX        =   17224
         _ExtentY        =   4233
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
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
            Left            =   7200
            TabIndex        =   13
            Top             =   240
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
            Left            =   3120
            TabIndex        =   12
            Top             =   1440
            Width           =   2295
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
            Left            =   3120
            TabIndex        =   11
            Top             =   840
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
            Height          =   450
            Left            =   3120
            TabIndex        =   10
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Biaya :"
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
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Bulan             :"
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
            TabIndex        =   8
            Top             =   1440
            Width           =   2535
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
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
            Left            =   240
            TabIndex        =   7
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
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
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   2775
         End
      End
   End
End
Attribute VB_Name = "frmlihatjenismember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "" Then
MsgBox "Data Tidak Ada", vbCritical + vbOKOnly, "Peringatan"
Else
formMember.Text7 = Text1
formMember.Text8 = Text2
formMember.Text9 = Text3
formMember.Text6 = Text4

     On Error Resume Next
    formMember.DTPicker2.Value = formMember.DTPicker1.Value
        If Val(formMember.Text9.Text) + formMember.DTPicker1.Month = 2 Then
            formMember.DTPicker2.Month = 2
        ElseIf Val(formMember.Text9.Text) + formMember.DTPicker1.Month <= 12 Then
            formMember.DTPicker2.Month = formMember.DTPicker1.Month + Val(formMember.Text9.Text)
           
        ElseIf Val(formMember.Text9.Text) + formMember.DTPicker1.Month > 12 Then
            formMember.DTPicker2.Month = (formMember.DTPicker1.Month + Val(formMember.Text9.Text)) Mod 12
            formMember.DTPicker2.Year = (formMember.DTPicker2.Year + ((formMember.DTPicker1.Month + Val(formMember.Text9.Text)) + _
                                            (12 - formMember.DTPicker2.Month)) / 12) - 1
        End If
Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DataGrid1_Click()
If Not Adodc1.Recordset.EOF = False Then
MsgBox "data Masih Kosong", vbInformation + vbOKOnly, "Informasi"
Else
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
Adodc1.RecordSource = "select KodeJenisMember,NamaJenisMember,JumlahBulan,Biaya from JenisMember where JumlahBulan > 0 order by JumlahBulan"
Adodc1.Refresh

Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
End Sub

