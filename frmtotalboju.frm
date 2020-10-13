VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.OCX"
Begin VB.Form frmtotalboju 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6510
   Icon            =   "frmtotalboju.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6800
      _Version        =   262144
      Locked          =   -1  'True
      PaneTree        =   "frmtotalboju.frx":9E4A
      Begin Threed.SSPanel SSPanel2 
         Height          =   1425
         Left            =   30
         TabIndex        =   2
         Top             =   2400
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   2514
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   480
            TabIndex        =   12
            Text            =   "Text5"
            Top             =   -480
            Width           =   2295
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   1200
            TabIndex        =   11
            Text            =   "Text4"
            Top             =   -480
            Width           =   615
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   840
            Top             =   -480
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
         Begin VB.CommandButton Command2 
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
            Left            =   3720
            Picture         =   "frmtotalboju.frx":9E9C
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
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
            Left            =   5040
            Picture         =   "frmtotalboju.frx":BB66
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   120
            Width           =   1095
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   2280
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   4022
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox Text3 
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
            Left            =   2880
            TabIndex        =   8
            Top             =   1560
            Width           =   3015
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2880
            TabIndex        =   6
            Top             =   960
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2880
            TabIndex        =   4
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Kembali          :"
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
            TabIndex        =   7
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Bayar               :"
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
            TabIndex        =   5
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Grand Total :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   480
            TabIndex        =   3
            Top             =   360
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "frmtotalboju"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Val(Text1.Text) > Val(Text2.Text) Then
MsgBox "Uang Yang Dibayarkan Tidak Mencukupi", vbInformation + vbOKOnly, "Peringatan"
ElseIf Val(Text1.Text) <= Val(Text2.Text) Then
Call BukaDB
Dim sql As String
sql = "insert into TrPakaiLapangan (Tanggal,NoPakaiLapangan,NoBooking,KodeLapangan,DP,HargaSewaLapangan,TotalPembelian,GrandTotalHarga) values ('" & Format(frmpakailapangan.DTPicker1, "yyyy/MM/DD") & "','" & frmpakailapangan.Text1.Text & "','" & frmpakailapangan.Text2.Text & "','" & frmpakailapangan.Text3.Text & "','" & frmpakailapangan.Text4.Text & "','" & frmpakailapangan.Text5.Text & "','" & frmpakailapangan.Text13.Text & "','" & frmpakailapangan.Text6 & "')"
Koneksi.Execute (sql)

With frmpakailapangan.ListView1
Dim i As Integer
For i = 1 To .ListItems.Count
Koneksi.Execute "insert into ItemPenjualan (Tanggal,Nourut,KodePenjualan,NoPakaiLapangan,KodeBarang,Jumlah,Harga,Total) values ('" & Format(frmpakailapangan.DTPicker1, "YYYY/MM/DD") & "','" & .ListItems(i).Text & "','" & .ListItems(i).ListSubItems(1).Text & "','" & .ListItems(i).ListSubItems(2).Text & "','" & .ListItems(i).ListSubItems(3).Text & "','" & .ListItems(i).ListSubItems(4).Text & "','" & .ListItems(i).ListSubItems(5).Text & "','" & .ListItems(i).ListSubItems(6).Text & "')"
Next
End With

For i = 1 To frmpakailapangan.ListView1.ListItems.Count
frmpakailapangan.Adodc3.RecordSource = "select * from MstBarang where KodeBarang ='" & frmpakailapangan.ListView1.ListItems(i).SubItems(3) & "'"
frmpakailapangan.Adodc3.Refresh
With frmpakailapangan.Adodc3.Recordset
If .RecordCount > 0 Then
.Clone
 !stok = !stok - Val(frmpakailapangan.ListView1.ListItems(i).SubItems(4))
.Update
frmpakailapangan.kosong
End If
End With
Next

Dim ubah As String
ubah = "update TrBooking set StatusBooking = 0 and Pembatalan = -1 where NoBooking='" & frmpakailapangan.Text2 & "'"
Koneksi.Execute (ubah)

MsgBox "Tersimpan", vbInformation + vbOKOnly, "Informasi"

Dim strcon As String
Dim strsql As String
Dim lokasidatabase As String
On Error Resume Next
strsql = "select * from ItemPenjualan where NoPakaiLapangan='" & frmpakailapangan.Text1 & "'"
With NotaPakaiLapangan
.ado.ConnectionString = Koneksi
.ado.Source = strsql
End With
NotaPakaiLapangan.Field16 = Text5
Adodc1.ConnectionString = Koneksi
NotaPakaiLapangan.Field17 = Text2
NotaPakaiLapangan.Field18 = Text3
frmMain.TampilkanForm "NotaPakaiLapangan"
Unload Me
frmpakailapangan.aturcommand
frmpakailapangan.ListView1.ListItems.Clear
frmpakailapangan.mati
frmpakailapangan.Text1 = ""
frmpakailapangan.Text2 = ""
End If
End Sub

Private Sub Form_Load()
With Me
.Top = (Screen.Height / 2) - (Me.Height / 2)
.Left = (Screen.Width / 2) - (Me.Width / 2)
End With
Text1.Enabled = False
Text3.Enabled = False
Text1.Text = frmpakailapangan.Text6
Text5 = frmpakailapangan.Text1
End Sub

Private Sub Text2_Change()
If Not IsNumeric(Text2) Then Text2 = "0"
Text3.Text = Val(Text2.Text) - Val(Text1.Text)
If Text3 < 0 Then
Label3.Caption = "Kurang    :"
Else
If Text3 >= 0 Then
Label3.Caption = "Kembali     :"
End If
End If
End Sub
