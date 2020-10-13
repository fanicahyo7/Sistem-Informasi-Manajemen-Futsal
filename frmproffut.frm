VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.OCX"
Begin VB.Form frmproffut 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7305
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   8281
      _Version        =   262144
      AutoSize        =   1
      Locked          =   -1  'True
      PaneTree        =   "frmproffut.frx":0000
      Begin Threed.SSPanel SSPanel2 
         Height          =   3225
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   5689
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
            Height          =   495
            Left            =   2760
            TabIndex        =   10
            Top             =   2520
            Width           =   3975
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
            Height          =   495
            Left            =   2760
            TabIndex        =   9
            Top             =   1920
            Width           =   3975
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
            Left            =   2760
            TabIndex        =   8
            Top             =   1320
            Width           =   3975
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
            Left            =   2760
            TabIndex        =   7
            Top             =   720
            Width           =   3975
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Profil Futsal"
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
            Left            =   480
            TabIndex        =   13
            Top             =   120
            Width           =   2895
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Kabupaten/Kota"
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
            TabIndex        =   6
            Top             =   2520
            Width           =   2175
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Telp"
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
            TabIndex        =   5
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label2 
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
            Left            =   240
            TabIndex        =   4
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Futsal"
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
            TabIndex        =   3
            Top             =   720
            Width           =   1815
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   1320
         Left            =   30
         TabIndex        =   1
         Top             =   3345
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   2328
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CommandButton Command2 
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
            Height          =   1095
            Left            =   4200
            Picture         =   "frmproffut.frx":0052
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
            Width           =   1215
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
            Left            =   5520
            Picture         =   "frmproffut.frx":1D1C
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   120
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmproffut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
Else
WriteIniValue "C:\Setting Futsal.ini", "Setting Futsal", "Nama Futsal", Text1.Text
WriteIniValue "C:\Setting Futsal.ini", "Setting Futsal", "Alamat", Text2.Text
WriteIniValue "C:\Setting Futsal.ini", "Setting Futsal", "Telp", Text3.Text
WriteIniValue "C:\Setting Futsal.ini", "Setting Futsal", "Kabupaten/Kota", Text4.Text
Call BukaDB
'Dim nama As String
'nama = ReadIniValue("C:\Setting Futsal.ini", "Setting Futsal", "Nama Futsal")
'If nama = "" Then
'frmMain.SSPanel6(0).Caption = "Nama Futsal"
'Else
'frmMain.SSPanel6(0).Caption = ReadIniValue("C:\Setting Futsal.ini", "Setting Futsal", "Nama Futsal")
'End If
Unload Me
End If
End Sub

Private Sub Form_Load()
With Me
.Top = (Screen.Height / 2) - (Me.Height / 2)
.Left = (Screen.Width / 2) - (Me.Width / 2)
End With
Call BukaDB
Dim nm As String
nm = ReadIniValue("C:\Setting Futsal.ini", "Setting Futsal", "Nama Futsal")

If nm = "" Then
Exit Sub
Else
Unload Me
frmSplash.Show
End If
'Text1.Text = ReadIniValue("C:\Setting Futsal.ini", "Setting Futsal", "Nama Futsal")
'Text2.Text = ReadIniValue("C:\Setting Futsal.ini", "Setting Futsal", "Alamat")
'Text3.Text = ReadIniValue("C:\Setting Futsal.ini", "Setting Futsal", "Telp")
'Text4.Text = ReadIniValue("C:\Setting Futsal.ini", "Setting Futsal", "Kabupaten/Kota")
End Sub

Private Sub Text3_Change()
If Not IsNumeric(Text3) Then Text3.Text = "0"
End Sub
