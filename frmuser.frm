VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.OCX"
Begin VB.Form frmuser 
   BorderStyle     =   0  'None
   Caption         =   "frmuser"
   ClientHeight    =   9360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20160
   LinkTopic       =   "Form2"
   ScaleHeight     =   9360
   ScaleWidth      =   20160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20160
      _ExtentX        =   35560
      _ExtentY        =   16510
      _Version        =   262144
      AutoSize        =   1
      Locked          =   -1  'True
      PaneTree        =   "frmuser.frx":0000
      Begin Threed.SSPanel SSPanel3 
         Height          =   1530
         Left            =   30
         TabIndex        =   3
         Top             =   7800
         Width           =   20100
         _ExtentX        =   35454
         _ExtentY        =   2699
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
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
            Height          =   1095
            Left            =   5040
            Picture         =   "frmuser.frx":0072
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   360
            Width           =   1215
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
            Height          =   1095
            Left            =   3600
            Picture         =   "frmuser.frx":1D3C
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   360
            Width           =   1215
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
            Left            =   2160
            Picture         =   "frmuser.frx":3A06
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   360
            Width           =   1215
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
            Height          =   1095
            Left            =   720
            Picture         =   "frmuser.frx":56D0
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   360
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   3960
         Left            =   30
         TabIndex        =   2
         Top             =   3750
         Width           =   20100
         _ExtentX        =   35454
         _ExtentY        =   6985
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   3975
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   18135
            _ExtentX        =   31988
            _ExtentY        =   7011
            _Version        =   393216
            Cols            =   7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   3630
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   20100
         _ExtentX        =   35454
         _ExtentY        =   6403
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   6480
            TabIndex        =   22
            Text            =   "Text7"
            Top             =   3600
            Width           =   375
         End
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   330
            Left            =   2280
            Top             =   3600
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
            Height          =   450
            Left            =   1920
            TabIndex        =   20
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   4800
            TabIndex        =   19
            Text            =   "Text5"
            Top             =   3600
            Width           =   1335
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   3600
            Top             =   3600
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
            ConnectStringType=   3
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
         Begin VB.ComboBox Combo1 
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
            Left            =   7200
            TabIndex        =   14
            Top             =   2160
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
            Height          =   495
            Left            =   7200
            TabIndex        =   13
            Top             =   1440
            Width           =   2775
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
            Height          =   735
            Left            =   7200
            TabIndex        =   12
            Top             =   600
            Width           =   2775
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
            Height          =   495
            IMEMode         =   3  'DISABLE
            Left            =   1920
            PasswordChar    =   "*"
            TabIndex        =   11
            Top             =   2040
            Width           =   2535
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
            Height          =   495
            Left            =   1920
            TabIndex        =   10
            Top             =   1320
            Width           =   2535
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode User"
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
            TabIndex        =   21
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Status User"
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
            Left            =   5520
            TabIndex        =   9
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Hp"
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
            Left            =   5520
            TabIndex        =   8
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
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
            Left            =   5520
            TabIndex        =   7
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
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
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "UserName"
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
            Top             =   1320
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gridbaris As Byte
Sub AktifkanGrid()
    With MSFlexGrid1
        .RowHeightMin = 300
        .Col = 0
        .Row = 0
        .Text = "NO"
        .CellFontBold = True
        .ColWidth(0) = 400
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .RowHeightMin = 300
        .Col = 1
        .Row = 0
        .Text = "Kode User"
        .CellFontBold = True
        .ColWidth(1) = 2000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 2
        .Row = 0
        .Text = "User Name"
        .CellFontBold = True
        .ColWidth(2) = 1900
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 3
        .Row = 0
        .Text = "Password"
        .CellFontBold = True
        .ColWidth(3) = 1600
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 4
        .Row = 0
        .Text = "Alamat"
        .CellFontBold = True
        .ColWidth(4) = 3300
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 5
        .Row = 0
        .Text = "No.Hp"
        .CellFontBold = True
        .ColWidth(5) = 1800
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 6
        .Row = 0
        .Text = "UserID"
        .CellFontBold = True
        .ColWidth(6) = 1000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
    End With
End Sub

Sub TampilkanGrid()
    Dim Baris As Integer
    MSFlexGrid1.Clear
    Call AktifkanGrid
    
    MSFlexGrid1.Rows = 2
    Baris = 0
    
    Dim anu As New ADODB.Recordset
anu.Open "select * from Login where KodeUser='" & frmMain.Label1 & "'", Koneksi
frmMain.SSPanel3 = anu!UserName
    
    Adodc2.RecordSource = "select * from Login where KodeUser='" & frmMain.Label1 & "'"
    Adodc2.Refresh
    If Adodc2.Recordset!LevelID = "1" Then
    Call BukaDB
    rsuser.Open "select * from Login order by UserName", Koneksi
    Else
    rsuser.Open "select * from Login where Username='" & frmMain.SSPanel3 & "'", Koneksi
    End If
    
    If rsuser.BOF Then
        MsgBox "Tabel Barang masih kosong!", _
        vbOKOnly + vbInformation, "Perhatian"
        Exit Sub
    Else
        With rsuser
        .MoveFirst
        Do While Not .EOF
            On Error Resume Next
            Baris = Baris + 1
            MSFlexGrid1.Rows = Baris + 1
            MSFlexGrid1.TextMatrix(Baris, 0) = Baris
            MSFlexGrid1.TextMatrix(Baris, 1) = !KodeUser
            MSFlexGrid1.TextMatrix(Baris, 2) = !UserName
            MSFlexGrid1.TextMatrix(Baris, 3) = "********"
            MSFlexGrid1.TextMatrix(Baris, 4) = !alamat
            MSFlexGrid1.TextMatrix(Baris, 5) = !HPNO
            MSFlexGrid1.TextMatrix(Baris, 6) = !LevelID
        .MoveNext
        Loop
        End With
    End If
End Sub
Sub buka()
Dim aduh As New ADODB.Recordset
aduh.Open "select * from Login where Kodeuser='" & frmMain.Label1 & "'", Koneksi
If aduh!LevelID = 1 Then
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Combo1.Enabled = True
ElseIf aduh!LevelID = 2 Then
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
End If
End Sub

Sub tutup()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Combo1.Enabled = False
Text6.Enabled = False
End Sub
Sub keluar()
Unload Me
End Sub
Sub kosong()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Combo1 = ""
Text6 = ""
End Sub
Private Sub autonumber()
Call BukaDB
rsuser.Open ("SELECT * FROM Login WHERE KodeUser in(select max(KodeUser) from Login)order by KodeUser desc"), Koneksi
rsuser.Requery
    Dim Urut As String * 6
    Dim Hitung As Long
    With rsuser
        If .EOF Then
            Urut = "USR" + "001"
            Text6 = Urut
        Else
            Hitung = Right(!KodeUser, 3) + 1
            Urut = "USR" + Right("00" & Hitung, 3)
        End If
        Text6 = Urut
    End With
End Sub

Private Sub Combo1_Click()
Dim rscom As ADODB.Recordset
Set rscom = New ADODB.Recordset
rscom.Open "select * from AccessLevel where LevelName='" & Combo1 & "'", Koneksi
Text5.Text = rscom!LevelID
End Sub

Private Sub Command1_Click()
If Command1.Caption = "&Baru" Then
Command1.Caption = "&Simpan"
buka
kosong
autonumber
MSFlexGrid1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Else
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
Else
Dim cari As String
cari = "UserName='" & Text1 & "'"
With Adodc1.Recordset
.Find cari
If Not .EOF Then
MsgBox "Data Sudah Ada", vbCritical + vbOKOnly, "Peringatan"
Else
Call BukaDB
rsuser.Open "insert into Login(KodeUser,UserName,UserPsw,Alamat,HPNo,LevelID) values('" & Text6 & "','" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "')", Koneksi
MsgBox "Data Berhasil Disimpan", vbInformation + vbOKOnly, "Informasi"
Form_Load
End If
End With
End If
End If
End Sub

Private Sub Command2_Click()
Form_Load
End Sub

Private Sub Command3_Click()
If Text1 = "" Then
MsgBox "Data Belum Dipilih", vbCritical + vbOKOnly, "Peringatan"
Else
If Command3.Caption = "&Ubah" Then
Command3.Caption = "&Simpan"
buka
Text2.Enabled = False
Command1.Enabled = False
Command2.Enabled = True
Command4.Enabled = False
MSFlexGrid1.Enabled = False
Else
Call BukaDB
Dim sqlupdate As String
sqlupdate = "update Login set UserName='" & Text1 & "',Alamat='" & Text3 & "',HPNo='" & Val(Text4) & "' where KodeUser='" & Text6 & "'"
Koneksi.Execute sqlupdate
MsgBox "Data Berhasil Diubah", vbInformation + vbOKOnly, "Informasi"
Form_Load
End If
End If
End Sub

Private Sub Command4_Click()
If Text1 = "" Then
MsgBox "Data Belum Dipilih", vbCritical + vbOKOnly, "Peringatan"
Else
konfirmasi = MsgBox("Apakah Anda Yakin Menghapus " & Text1 & "?", vbQuestion + vbYesNo, "Konfirmasi")
If konfirmasi = vbYes Then
Dim sqldelete As String
sqldelete = "delete from Login where KodeUser='" & Text6 & "'"
Koneksi.Execute sqldelete
Form_Load
End If
End If
End Sub

Private Sub Form_Load()
Call BukaDB
Text5.Visible = False
Adodc1.Visible = False
Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "select * from Login"
Adodc1.Refresh
Adodc2.Visible = False
Adodc2.ConnectionString = Koneksi

Call TampilkanGrid
Command1.Caption = "&Baru"
Command1.Enabled = True
Command2.Enabled = False
Command3.Caption = "&Ubah"
Command3.Enabled = True
Command4.Enabled = True
MSFlexGrid1.Enabled = True
kosong
tutup

Dim rscom As ADODB.Recordset
Set rscom = New ADODB.Recordset
Combo1.Clear
rscom.Open "select * from AccessLevel", Koneksi
        Do While Not rscom.EOF
            Combo1.AddItem rscom!levelname
            rscom.MoveNext
        Loop
Text7.Visible = False
Dim rsanu As New ADODB.Recordset
rsanu.Open "select * from login where username='" & frmMain.SSPanel3.Caption & "'", Koneksi
Text7 = rsanu!LevelID
Dim rsiku As New ADODB.Recordset
rsiku.Open "select * from AccessLevel where Levelid='" & Text7 & "'", Koneksi
Text7 = rsiku!levelname
If Text7 = "Kasir" Then
Command1.Enabled = False
Command4.Enabled = False
ElseIf Text7 = "Admin" Then
Exit Sub
End If
End Sub

Private Sub MSFlexGrid1_DblClick()
gridbaris = MSFlexGrid1.Row
    Call BukaDB
    Dim rsini As ADODB.Recordset
    Set rsini = New ADODB.Recordset
    
    rsini.Open "SELECT * FROM Login WHERE UserName='" & MSFlexGrid1.TextMatrix(gridbaris, 2) & "'", Koneksi, adOpenDynamic, adLockBatchOptimistic
    
    If rsini.BOF Then
        MsgBox "Data Tidak Ada", vbOKOnly + vbInformation, "Perhatian"
        Exit Sub
    Else
        rsini.MoveFirst
        Do While Not rsini.EOF
            On Error Resume Next
            Text1.Text = rsini!UserName
            Text2.Text = rsini!Userpsw
            Text3.Text = rsini!alamat
            Text4.Text = rsini!HPNO
            Text5.Text = rsini!LevelID
            Text6.Text = rsini!KodeUser
            
        Call BukaDB
Dim rsitu As ADODB.Recordset
Set rsitu = New ADODB.Recordset
rsitu.Open "select * from AccessLevel where LevelID like '" & Text5.Text & "'", Koneksi
Combo1.Text = rsitu!levelname

        rsini.MoveNext
        Loop
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub Text4_Change()
If Not IsNumeric(Text4) Then Text4 = "0"
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo1.SetFocus
End If
End Sub
