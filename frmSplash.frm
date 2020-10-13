VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3360
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":9E4A
   ScaleHeight     =   3360
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   5535
      _Version        =   720898
      _ExtentX        =   9763
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483638
   End
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   120
      Top             =   1320
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Kota Batu"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "082334349494"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   2160
      Width           =   4935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jl.Dadaptulis"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1800
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Barca Futsal"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1320
      Width           =   4695
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

With Me
.Top = (Screen.Height / 2) - (Me.Height / 2)
.Left = (Screen.Width / 2) - (Me.Width / 2)
End With

Timer1.Enabled = True

'Label1.Caption = ReadIniValue(App.Path & "\Setting.ini", "Setting Futsal", "Nama Futsal")
'Label2.Caption = ReadIniValue(App.Path & "\Setting.ini", "Setting Futsal", "Alamat")
'Label3.Caption = ReadIniValue(App.Path & "\Setting.ini", "Setting Futsal", "Telp")
'Label4.Caption = ReadIniValue(App.Path & "\Setting.ini", "Setting Futsal", "Kabupaten/Kota")
End Sub

Private Sub Timer1_Timer()

ProgressBar2.Value = ProgressBar2.Value + 1

If ProgressBar2.Value = 100 Then

Timer1.Enabled = False

Unload Me
frmMain.Show
End If

End Sub
