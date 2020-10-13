VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.OCX"
Begin VB.Form frmlihatbooking 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10710
   Icon            =   "frmlihatbooking.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7920
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   13996
      _Version        =   262144
      Locked          =   -1  'True
      PaneTree        =   "frmlihatbooking.frx":9E4A
      Begin Threed.SSPanel SSPanel3 
         Height          =   1215
         Left            =   30
         TabIndex        =   3
         Top             =   6690
         Width           =   10635
         _ExtentX        =   18759
         _ExtentY        =   2143
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
            Height          =   975
            Left            =   8160
            Picture         =   "frmlihatbooking.frx":9EBC
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   1095
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
            Height          =   975
            Left            =   9360
            Picture         =   "frmlihatbooking.frx":BB86
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   120
            Width           =   1095
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2310
         Left            =   30
         TabIndex        =   2
         Top             =   4290
         Width           =   10635
         _ExtentX        =   18759
         _ExtentY        =   4075
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmlihatbooking.frx":D850
            Height          =   2295
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   4048
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   18
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
               Size            =   8.25
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
         Height          =   4170
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   10635
         _ExtentX        =   18759
         _ExtentY        =   7355
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Frame Frame1 
            BackColor       =   &H8000000D&
            Caption         =   "Pencarian"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   600
            TabIndex        =   29
            Top             =   2640
            Width           =   4335
            Begin VB.TextBox Text8 
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   1800
               TabIndex        =   33
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox Text5 
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1800
               TabIndex        =   32
               Top             =   360
               Width           =   2295
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Atas Nama :"
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
               TabIndex        =   31
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "No. Booking :"
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
               TabIndex        =   30
               Top             =   360
               Width           =   1335
            End
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   495
            Left            =   7320
            TabIndex        =   27
            Top             =   600
            Width           =   2055
            _ExtentX        =   3625
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
            Format          =   171835393
            CurrentDate     =   41755
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   7320
            TabIndex        =   26
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   171835393
            CurrentDate     =   41755
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
            Height          =   435
            Left            =   7320
            TabIndex        =   25
            Top             =   2040
            Width           =   2055
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H8000000D&
            Caption         =   "No"
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
            Left            =   8160
            TabIndex        =   23
            Top             =   1560
            Width           =   735
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H8000000D&
            Caption         =   "Yes"
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
            Left            =   7320
            TabIndex        =   22
            Top             =   1560
            Width           =   735
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H8000000D&
            Caption         =   "No"
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
            Left            =   8160
            TabIndex        =   20
            Top             =   1200
            Width           =   615
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H8000000D&
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7320
            TabIndex        =   19
            Top             =   1200
            Width           =   735
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
            Height          =   435
            Left            =   2640
            TabIndex        =   15
            Top             =   2160
            Width           =   2055
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
            Height          =   435
            Left            =   2640
            TabIndex        =   13
            Top             =   1680
            Width           =   2055
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Puma Gaffer by Barreto"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2640
            TabIndex        =   11
            Top             =   1200
            Width           =   2055
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
            Height          =   435
            Left            =   2640
            TabIndex        =   9
            Top             =   720
            Width           =   2055
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
            Left            =   2640
            TabIndex        =   7
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "DP                       :"
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
            Left            =   5280
            TabIndex        =   24
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Pembatalan         :"
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
            Left            =   5280
            TabIndex        =   21
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Status Booking    :"
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
            Left            =   5280
            TabIndex        =   18
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Booking :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5280
            TabIndex        =   17
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal               :"
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
            Left            =   5280
            TabIndex        =   16
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Atas Nama :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   480
            TabIndex        =   14
            Top             =   2160
            Width           =   2175
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Mulai                 :"
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
            Left            =   480
            TabIndex        =   12
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga                        :"
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
            TabIndex        =   10
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Lapangan         :"
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
            TabIndex        =   8
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Booking              :"
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
            TabIndex        =   6
            Top             =   240
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "frmlihatbooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
Check2.Value = 0
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

Private Sub DataGrid1_Click()
If Not Adodc1.Recordset.EOF = False Then
MsgBox "Data Masih Kosong", vbInformation + vbOKOnly, "Informasi"
Else
Text1.Text = Adodc1.Recordset!NoBooking
Text2.Text = Adodc1.Recordset!Kodelapangan
Text3.Text = Adodc1.Recordset!Harga
Text4.Text = Adodc1.Recordset!JamMulai
DTPicker1 = Adodc1.Recordset!Tanggal
DTPicker2.Value = Adodc1.Recordset!TanggalBooking
Text6.Text = Adodc1.Recordset!AtasNama
Text7.Text = Adodc1.Recordset!DP

If Adodc1.Recordset!StatusBooking = -1 Then
Check1.Value = 1
ElseIf Adodc1.Recordset!StatusBooking = 0 Then
Check2 = 1
End If

If Adodc1.Recordset!Pembatalan = -1 Then
Check3.Value = 1
ElseIf Adodc1.Recordset!Pembatalan = 0 Then
Check4.Value = 1
End If
End If
End Sub
Sub tutup()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
DTPicker1.Enabled = False
DTPicker2.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Check4.Enabled = False
End Sub

Private Sub Form_Load()
frmpakailapangan.Text2.Text = ""
With Me
.Top = (Screen.Height / 2) - (Me.Height / 2)
.Left = (Screen.Width / 2) - (Me.Width / 2)
End With
Call BukaDB
Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "select * from TrBooking where StatusBooking = -1 and Pembatalan = 0"
Adodc1.Refresh

tutup
End Sub

Private Sub SSCommand1_Click()
Unload Me
End Sub

Private Sub SSCommand2_Click()
If Text1 = "" Then
MsgBox "Data Tidak Ada", vbCritical + vbOKOnly, "Peringatan"
Else
If Check2 = 1 Then
MsgBox "Booking Sudah Tidak Berlaku", vbInformation + vbOKOnly, "Informasi"
ElseIf Check1 = 1 Then
Call BukaDB

frmpakailapangan.Text2.Text = Adodc1.Recordset!NoBooking
frmpakailapangan.Text3.Text = Adodc1.Recordset!Kodelapangan
frmpakailapangan.Text4.Text = Adodc1.Recordset!DP
frmpakailapangan.Text5.Text = Adodc1.Recordset!Harga
frmpakailapangan.Text7.Text = Adodc1.Recordset!JamMulai
frmpakailapangan.Text14.Text = Adodc1.Recordset!AtasNama
Unload Me

a = MsgBox("Apakah Anda Melakukan Pembelian ? ", vbQuestion + vbYesNo, "Konfirmasi")
If a = vbYes Then
frmpakailapangan.bukapenjualan
frmpakailapangan.autokojul
nurut
frmpakailapangan.Command5.Enabled = True
frmpakailapangan.Text9.SetFocus
ElseIf a = vbNo Then
frmpakailapangan.Text13 = "0"
frmpakailapangan.Command5.Enabled = False
End If
End If
End If
End Sub

Private Sub nurut()
Call BukaDB
rsuser.Open ("SELECT * FROM ItemPenjualan WHERE NoUrut in(select max(NoUrut) from ItemPenjualan)order by NoUrut desc"), Koneksi
rsuser.Requery
    Dim Urut As String * 4
    Dim Hitung As Long
    With rsuser
        If .EOF Then
            Urut = "01"
            frmpakailapangan.Text8 = Urut
        Else
            Hitung = Right(!NoUrut, 4) + 1
            Urut = Right("0" & Hitung, 4)
        End If
        frmpakailapangan.Text8 = Urut
    End With
End Sub

Private Sub Text5_Change()
Adodc1.RecordSource = "select * from TrBooking where nobooking like '%" & TandaPetik(Text5) & "%' and StatusBooking = -1 and Pembatalan = 0"
Adodc1.Refresh
If Not Adodc1.Recordset.EOF Then
    With DataGrid1
    Set .DataSource = Adodc1
        .Refresh
    End With
End If
End Sub

Private Sub Text8_Change()
Adodc1.RecordSource = "select * from TrBooking where atasnama like '%" & TandaPetik(Text8) & "%' and StatusBooking = -1 and Pembatalan = 0"
Adodc1.Refresh
If Not Adodc1.Recordset.EOF Then
    With DataGrid1
    Set .DataSource = Adodc1
        .Refresh
    End With
End If
End Sub
