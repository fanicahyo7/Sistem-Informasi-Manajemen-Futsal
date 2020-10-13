VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frmlogin 
   BorderStyle     =   0  'None
   ClientHeight    =   2985
   ClientLeft      =   6750
   ClientTop       =   3990
   ClientWidth     =   5880
   ControlBox      =   0   'False
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5318
      _Version        =   262144
      PaneTree        =   "frmlogin.frx":9E4A
      Begin Threed.SSPanel SSPanel1 
         Height          =   2955
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   5212
         _Version        =   262144
         BackColor       =   12648384
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
            Height          =   1095
            Left            =   3960
            Picture         =   "frmlogin.frx":9E7C
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1680
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Masuk"
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
            Left            =   2640
            Picture         =   "frmlogin.frx":BB46
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1680
            Width           =   1095
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
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   2520
            PasswordChar    =   "*"
            TabIndex        =   6
            Top             =   1080
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
            Height          =   375
            Left            =   2520
            TabIndex        =   5
            Top             =   480
            Width           =   2535
         End
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   330
            Left            =   1680
            Top             =   3000
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
            Height          =   330
            Left            =   1200
            Top             =   3000
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
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Password   :"
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
            Left            =   840
            TabIndex        =   4
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "User Name :"
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
            Left            =   840
            TabIndex        =   3
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Login"
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
            TabIndex        =   2
            Top             =   120
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call BukaDB
'        Adodc1.RecordSource = "Select * from Login where UserName ='" & TandaPetik(Text1) & "' and UserPsw='" & TandaPetik(Text2) & "'"
'        Adodc1.Refresh
'        If Adodc1.Recordset.EOF Then
        Dim rssitu As New ADODB.Recordset
            rssitu.Open "Select * from Login where UserName ='" & TandaPetik(Text1) & "' and UserPsw='" & TandaPetik(Text2) & "'", Koneksi
            If rssitu.EOF Then
            MsgBox "Password Salah, Coba Lagi!", vbCritical + vbOKOnly, "Peringatan"
            Text2 = ""
            Text2.SetFocus
        Else
        Call BukaDB
'        Adodc1.RecordSource = "select * from Login where Username='" & Text1 & "'"
'        Adodc1.Refresh
'        frmMain.SSPanel3.Caption = Adodc1.Recordset!LevelID
        frmMain.Label1 = rssitu!KodeUser
        frmMain.SSPanel3.Caption = rssitu!LevelID
        
'        Adodc2.RecordSource = "select * from AccessLevel where LevelID like '" & frmMain.SSPanel3 & "'"
'        Adodc2.Refresh
'            frmMain.SSPanel3.Caption = Adodc2.Recordset!LevelName
'        If Adodc2.Recordset!LevelID = 1 Then
    Dim rssana As New ADODB.Recordset
        rssana.Open "select * from AccessLevel where LevelID like '" & frmMain.SSPanel3 & "'", Koneksi
        If rssana!LevelID = 1 Then
        frmMain.TaskPanel1.Groups.Clear
        frmMain.SetMainMenu
        Else
'        If Adodc2.Recordset!LevelID = 2 Then
        If rssana!LevelID = 2 Then
        frmMain.TaskPanel1.Groups.Clear
        frmMain.SetMainMenu2
        Else
        If rssana!LevelID = 3 Then
        frmMain.TaskPanel1.Groups.Clear
        frmMain.SetMainMenu3
        End If
        End If
        End If
            frmMain.SSPanel3.Caption = Text1
            Unload Me
            frmMain.WindowState = 2
            
            
End If
End Sub

Private Sub Command2_Click()
If MsgBox("Apakah Anda Yakin Ingin keluar Aplikasi ?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
Unload Me
End
Else
Exit Sub
End If
End Sub

Private Sub Form_Load()
Call BukaDB

'Adodc1.ConnectionString = Koneksi
'Adodc1.RecordSource = "select * from login"
'Adodc1.Refresh
'
'Adodc2.ConnectionString = Koneksi
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
    Call BukaDB
'    Adodc1.RecordSource = "Select * From Login where UserName='" & TandaPetik(Text1) & "'"
'    Adodc1.Refresh
'            If Adodc1.Recordset.EOF Then
            Dim rsanu As New ADODB.Recordset
            rsanu.Open "Select * From Login where UserName='" & TandaPetik(Text1) & "'", Koneksi
                If rsanu.EOF Then
                MsgBox "User Tidak Terdeteksi, Coba lagi", vbCritical + vbOKOnly, "Peringatan"
                Text1 = ""
            Else
            Text2.SetFocus
            End If
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call BukaDB
'        Adodc1.RecordSource = "Select * from Login where UserName ='" & TandaPetik(Text1) & "' and UserPsw='" & TandaPetik(Text2) & "'"
'        Adodc1.Refresh
'        If Adodc1.Recordset.EOF Then
        Dim rssitu As New ADODB.Recordset
            rssitu.Open "Select * from Login where UserName ='" & TandaPetik(Text1) & "' and UserPsw='" & TandaPetik(Text2) & "'", Koneksi
            If rssitu.EOF Then
            MsgBox "Password Salah, Coba Lagi!", vbCritical + vbOKOnly, "Peringatan"
            Text2 = ""
            Text2.SetFocus
        Else
        Call BukaDB
'        Adodc1.RecordSource = "select * from Login where Username='" & Text1 & "'"
'        Adodc1.Refresh
'        frmMain.SSPanel3.Caption = Adodc1.Recordset!LevelID
        frmMain.Label1 = rssitu!KodeUser
        frmMain.SSPanel3.Caption = rssitu!LevelID
        
'        Adodc2.RecordSource = "select * from AccessLevel where LevelID like '" & frmMain.SSPanel3 & "'"
'        Adodc2.Refresh
'            frmMain.SSPanel3.Caption = Adodc2.Recordset!LevelName
'        If Adodc2.Recordset!LevelID = 1 Then
    Dim rssana As New ADODB.Recordset
        rssana.Open "select * from AccessLevel where LevelID like '" & frmMain.SSPanel3 & "'", Koneksi
        If rssana!LevelID = 1 Then
        frmMain.TaskPanel1.Groups.Clear
        frmMain.SetMainMenu
        Else
'        If Adodc2.Recordset!LevelID = 2 Then
        If rssana!LevelID = 2 Then
        frmMain.TaskPanel1.Groups.Clear
        frmMain.SetMainMenu2
        Else
        If rssana!LevelID = 3 Then
        frmMain.TaskPanel1.Groups.Clear
        frmMain.SetMainMenu3
        End If
        End If
        End If
            frmMain.SSPanel3.Caption = Text1
            Unload Me
            frmMain.WindowState = 2
            
            
End If
End If
End Sub
