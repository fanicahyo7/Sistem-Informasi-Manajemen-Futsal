VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.OCX"
Begin VB.Form frmtotallangsung 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   7455
   ClientTop       =   3990
   ClientWidth     =   6750
   Icon            =   "frmtotallangsung.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6800
      _Version        =   262144
      Locked          =   -1  'True
      PaneTree        =   "frmtotallangsung.frx":9E4A
      Begin Threed.SSPanel SSPanel1 
         Height          =   2295
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   4048
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
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
            Height          =   495
            Left            =   2760
            TabIndex        =   7
            Top             =   1560
            Width           =   3255
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2760
            TabIndex        =   5
            Top             =   960
            Width           =   3255
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
            Height          =   495
            Left            =   2760
            TabIndex        =   3
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Sisa               :"
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
            TabIndex        =   6
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Bayar            :"
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
            TabIndex        =   4
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Grand Total :"
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
            TabIndex        =   2
            Top             =   360
            Width           =   1695
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1410
         Left            =   30
         TabIndex        =   8
         Top             =   2415
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   2487
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   1320
            Top             =   -360
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
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   3840
            Picture         =   "frmtotallangsung.frx":9E9C
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
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
            Left            =   5160
            Picture         =   "frmtotallangsung.frx":BB66
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmtotallangsung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Val(Text1) > Val(Text2) Then
MsgBox "Uang Yang Dibayarkan Kurang", vbCritical + vbOKOnly, "Peringatan"
ElseIf Val(Text1) <= Val(Text2) Then
Call BukaDB
With frmpenjualanlangsung.ListView1
Dim i As Integer
For i = 1 To .ListItems.Count
Koneksi.Execute "insert into ItemPenjualan (Tanggal,Nourut,KodePenjualan,NoPakaiLapangan,kodeBarang,Jumlah,Harga,Total) values ('" & Format(frmpakailapangan.DTPicker1, "YYYY/MM/DD") & "','" & .ListItems(i).Text & "','" & .ListItems(i).ListSubItems(1).Text & "','" & .ListItems(i).ListSubItems(2).Text & "','" & .ListItems(i).ListSubItems(3).Text & "','" & .ListItems(i).ListSubItems(4).Text & "','" & .ListItems(i).ListSubItems(5).Text & "','" & .ListItems(i).ListSubItems(6).Text & "')"
Next
End With

For i = 1 To frmpenjualanlangsung.ListView1.ListItems.Count
frmpenjualanlangsung.Adodc1.RecordSource = "select * from MstBarang where KodeBarang ='" & frmpenjualanlangsung.ListView1.ListItems(i).SubItems(3) & "'"
frmpenjualanlangsung.Adodc1.Refresh
With frmpenjualanlangsung.Adodc1.Recordset
If .RecordCount > 0 Then
.Clone
 !stok = !stok - Val(frmpenjualanlangsung.ListView1.ListItems(i).SubItems(4))
.Update
End If
End With
Next
MsgBox "Tersimpan", vbInformation + vbOKOnly, "Informasi"
NotaPenjualan.Field2 = frmpenjualanlangsung.Text8
NotaPenjualan.Field1 = Format(frmpenjualanlangsung.DTPicker1, "DD/MM/YYYY")
NotaPenjualan.Field3 = frmpenjualanlangsung.Text7
NotaPenjualan.Field10 = Text2
NotaPenjualan.Field11 = Text3
Dim strcon As String
Dim strsql As String
Dim lokasidatabase As String
On Error Resume Next
strsql = "select * from ItemPenjualan where KodePenjualan='" & frmpenjualanlangsung.Text8 & "'"
With NotaPenjualan
.DataControl1.ConnectionString = Koneksi
.DataControl1.Source = strsql
End With
Unload Me
frmMain.TampilkanForm "NotaPenjualan"
frmpenjualanlangsung.ListView1.ListItems.Clear
frmpenjualanlangsung.mati
frmpenjualanlangsung.Text7 = ""
frmpenjualanlangsung.Text8 = ""
frmpenjualanlangsung.Command4.Enabled = False
frmpenjualanlangsung.Command1.Enabled = True
End If
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
Text1.Enabled = False
Text3.Enabled = False
Text1.Text = frmpenjualanlangsung.Text7.Text
End Sub

Private Sub Text2_Change()
If Not IsNumeric(Text2) Then Text2 = "0"
Text3.Text = Val(Text2) - Val(Text1)
If Text3 < 0 Then
Label3.Caption = "Kurang     :"
Else
If Text3 >= 0 Then
Label3.Caption = "Kembali      :"
End If
End If
End Sub

