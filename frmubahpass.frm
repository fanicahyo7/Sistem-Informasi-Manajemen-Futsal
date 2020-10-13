VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frmubahpass 
   ClientHeight    =   4680
   ClientLeft      =   6810
   ClientTop       =   4050
   ClientWidth     =   6840
   ControlBox      =   0   'False
   Icon            =   "frmubahpass.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   6840
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8281
      _Version        =   262144
      Locked          =   -1  'True
      PaneTree        =   "frmubahpass.frx":9E4A
      Begin Threed.SSPanel SSPanel1 
         Height          =   3165
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   5583
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
            Height          =   525
            IMEMode         =   3  'DISABLE
            Left            =   3960
            PasswordChar    =   "*"
            TabIndex        =   9
            Top             =   2400
            Width           =   2535
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
            Height          =   525
            IMEMode         =   3  'DISABLE
            Left            =   3960
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   1800
            Width           =   2535
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
            Height          =   525
            IMEMode         =   3  'DISABLE
            Left            =   3960
            PasswordChar    =   "*"
            TabIndex        =   5
            Top             =   1200
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
            Height          =   525
            Left            =   3960
            TabIndex        =   3
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Ubah Password :"
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
            TabIndex        =   13
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Konfirmasi Password Baru :"
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
            Left            =   720
            TabIndex        =   8
            Top             =   2400
            Width           =   3255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Password Baru                      :"
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
            Left            =   720
            TabIndex        =   6
            Top             =   1800
            Width           =   3135
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Password Lama                    :"
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
            Left            =   720
            TabIndex        =   4
            Top             =   1200
            Width           =   3135
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Username                               :"
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
            Left            =   720
            TabIndex        =   2
            Top             =   600
            Width           =   3135
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1380
         Left            =   30
         TabIndex        =   10
         Top             =   3285
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   2434
         _Version        =   262144
         BackColor       =   -2147483635
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CommandButton Command1 
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
            Left            =   3840
            Picture         =   "frmubahpass.frx":9E9C
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
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
            Left            =   5280
            Picture         =   "frmubahpass.frx":BB66
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   120
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmubahpass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call BukaDB
    rsuser.Open "select * from Login where UserName='" & TandaPetik(Text1) & "' and UserPsw='" & TandaPetik(Text2) & "'", Koneksi
    If rsuser.EOF Then
        MsgBox "Password Lama Salah", vbCritical + vbOKOnly, "Peringatan"
        Text2.SetFocus
        Text2 = ""
    Else
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan"
Else
If Text4 <> Text3 Then
        MsgBox "Password Konfirmasi Tidak Sama", vbCritical + vbOKOnly, "Peringatan"
        Text4.SetFocus
        Text4 = ""
    Else
       konfirmasi = MsgBox("Yakin Password Akan Diganti", vbQuestion + vbYesNo)
            If konfirmasi = vbYes Then
            Dim editpw As String
            editpw = "UPDATE Login set UserPsw='" & TandaPetik(Text4.Text) & "' WHERE UserName='" & TandaPetik(Text1.Text) & "' AND UserPsw='" & TandaPetik(Text2.Text) & "'"
            Koneksi.Execute (editpw)
            MsgBox "Ubah Password Berhasil", vbInformation + vbOKOnly, "Informasi"
            Unload Me
            frmMain.tutupform
            frmMain.laptutup
            frmMain.Label4.Caption = ""
            frmMain.SSPanel3.Caption = "LOG OFF"
            frmlogin.Show , Me
        Else
        MsgBox "Password Gagal Diubah", vbCritical + vbOKOnly, "Informasi"
        End If
    End If
End If
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
Call BukaDB
Text1.Enabled = False
Text1.Text = frmMain.SSPanel3.Caption
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
'If KeyAscii = 13 Then
'    Call BukaDB
'    rsuser.Open "select * from Login where UserName='" & TandaPetik(Text1) & "'", Koneksi
'    If Not rsuser.EOF Then
'        Text2.SetFocus
'    Else
'        MsgBox "Username Salah", vbCritical + vbOKOnly, "Peringatan"
'        Text1.SetFocus
'        Text1 = ""
'    End If
'End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    Call BukaDB
    rsuser.Open "select * from Login where UserName='" & TandaPetik(Text1) & "' and UserPsw='" & TandaPetik(Text2) & "'", Koneksi
    If Not rsuser.EOF Then
        Text3.SetFocus
    Else
        MsgBox "Password Lama Salah", vbCritical + vbOKOnly, "Peringatan"
        Text2.SetFocus
        Text2 = ""
    End If
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    If Text3 = "" Then
        MsgBox "Password Baru Belum Dibuat", vbCritical + vbOKOnly, "Peringatan"
        Text3.SetFocus
    Else
        Text4.SetFocus
    End If
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
Command1.SetFocus
End If
End Sub
