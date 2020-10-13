VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#11.2#0"; "Codejock.CommandBars.v11.2.2.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.OCX"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.TaskPanel.v11.2.2.ocx"
Object = "{43E0D4CB-B249-449C-AC19-6AD486137F7D}#4.0#0"; "IGTransition40.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "SIFUT (Sistem Informasi Manajemen Futsal)"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   17880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmMain.frx":9E4A
   ScaleHeight     =   737
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17880
      _ExtentX        =   31538
      _ExtentY        =   19500
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   2
      SplitterBarAppearance=   0
      BorderStyle     =   0
      Locked          =   -1  'True
      PaneTree        =   "frmMain.frx":19694
      Begin Threed.SSPanel SSPanel3 
         Height          =   750
         Left            =   0
         TabIndex        =   9
         Top             =   10305
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   1323
         _Version        =   262144
         ForeColor       =   16777215
         BackColor       =   12420637
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "LOG OFF"
         FloodColor      =   12582912
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   615
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   17880
         _ExtentX        =   31538
         _ExtentY        =   1085
         _Version        =   262144
         BackColor       =   12420637
         PictureBackgroundStyle=   2
         PictureBackground=   "frmMain.frx":19746
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel SSPanel6 
            Height          =   405
            Index           =   0
            Left            =   840
            TabIndex        =   2
            Top             =   0
            Width           =   16920
            _ExtentX        =   29845
            _ExtentY        =   714
            _Version        =   262144
            Font3D          =   5
            ForeColor       =   16777215
            BackColor       =   12420637
            BackStyle       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Barca Futsal"
            BevelOuter      =   0
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Height          =   255
               Left            =   4560
               TabIndex        =   10
               Top             =   120
               Width           =   735
            End
         End
         Begin Threed.SSPanel SSPanel6 
            Height          =   210
            Index           =   1
            Left            =   840
            TabIndex        =   3
            Top             =   360
            Width           =   5865
            _ExtentX        =   10345
            _ExtentY        =   370
            _Version        =   262144
            Font3D          =   5
            ForeColor       =   16777215
            BackColor       =   12420637
            BackStyle       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "SIFUT (Sistem Informasi Manajemen Futsal)"
            BevelOuter      =   0
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin VB.Image Image3 
            Height          =   555
            Left            =   165
            Picture         =   "frmMain.frx":D98B8
            Stretch         =   -1  'True
            Top             =   0
            Width           =   570
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   405
         Left            =   2250
         TabIndex        =   4
         Top             =   645
         Width           =   15630
         _ExtentX        =   27570
         _ExtentY        =   714
         _Version        =   262144
         BackColor       =   12632256
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   135
            TabIndex        =   5
            Top             =   90
            Width           =   60
         End
      End
      Begin ActiveTransition.SSTransition SSTransition1 
         Height          =   9975
         Left            =   2250
         TabIndex        =   6
         Top             =   1080
         Width           =   15630
         _ExtentX        =   27570
         _ExtentY        =   17595
         _Version        =   262144
         BorderStyle     =   0
         BackColor       =   12632256
         AutoSize        =   1
         Begin VB.Image Image1 
            Height          =   9975
            Left            =   0
            Picture         =   "frmMain.frx":F09E6
            Stretch         =   -1  'True
            Top             =   0
            Width           =   15630
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   9630
         Left            =   0
         TabIndex        =   7
         Top             =   645
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   16986
         _Version        =   262144
         BackColor       =   8421504
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeTaskPanel.TaskPanel TaskPanel1 
            Height          =   6015
            Left            =   480
            TabIndex        =   8
            Top             =   1560
            Width           =   1335
            _Version        =   720898
            _ExtentX        =   2355
            _ExtentY        =   10610
            _StockProps     =   64
            ItemLayout      =   1
            HotTrackStyle   =   3
         End
         Begin XtremeCommandBars.ImageManager ImageManager1 
            Left            =   0
            Top             =   0
            _Version        =   720898
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   0
            Icons           =   "frmMain.frx":10888E
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FormhWnd As Long

Private Sub Form_Load()
Label1.Visible = False
Me.WindowState = 2
Me.Show
'SetMainMenu
TaskPanel1.Groups.Clear
frmlogin.Show vbModal, Me
Image1.Width = Width
Image1.Height = Height
End Sub

Private Sub Form_Resize()
    TaskPanel1.Move 0, 0, SSPanel1.Width, SSPanel1.Height
    Image1.Width = Width
    Image1.Height = Height
End Sub

Private Sub SSTransition1_TransitionPrepare(TransitionType As ActiveTransition.Constants_TransitionType, Duration As Long, ByVal TagVariant As Variant)
    Dim frm As Form
    On Error Resume Next

    For Each frm In Forms
        If Not (frm.Name = "frmMain" _
            Or frm.Name = TagVariant.Name Or frm.Name = "frmSplash" Or frm.Name = "frmSysTray") Then
            frm.Hide
            SSTransition1.RemoveControl frm.hWnd
        End If
    Next
    SSTransition1.AddControl TagVariant.hWnd

    TagVariant.Show

    frmMain.Show
End Sub

Public Sub TampilkanForm(pNamaForm As String)
    On Error Resume Next

    Screen.MousePointer = vbHourglass

    Select Case pNamaForm
            
        Case "frmpembelian"
            SSTransition1.Transition ssDissolve, ssTransitionImmediate, frmpembelian
            FormhWnd = frmpembelian.hWnd
            Label4.Caption = "Transaksi Pembelian"
            
        Case "frmpenjualanlangsung"
            SSTransition1.Transition ssDissolve, ssTransitionImmediate, frmpenjualanlangsung
            FormhWnd = frmpenjualanlangsung.hWnd
            Label4.Caption = "Transaksi Penjualan Langsung"
            
        Case "frmpakailapangan"
            SSTransition1.Transition ssDissolve, ssTransitionImmediate, frmpakailapangan
            FormhWnd = frmpakailapangan.hWnd
            Label4.Caption = "Transaksi Pakai Lapangan Dan Penjualan"
            
        Case "formMember"
            SSTransition1.Transition ssDissolve, ssTransitionImmediate, formMember
            FormhWnd = formMember.hWnd
            Label4.Caption = "Master Member"
            
        Case "frmbarang"
            SSTransition1.Transition ssDissolve, ssTransitionImmediate, frmbarang
            FormhWnd = frmbarang.hWnd
            Label4.Caption = "Master Barang"

        Case "frmmstsupplier"
            SSTransition1.Transition ssDissolve, ssTransitionImmediate, frmmstsupplier
            FormhWnd = frmmstsupplier.hWnd
            Label4.Caption = "Master Supplier"
            
        Case "frmmstlapangan"
            SSTransition1.Transition ssDissolve, ssTransitionImmediate, frmmstlapangan
            FormhWnd = frmmstlapangan.hWnd
            Label4.Caption = "Master Lapangan"

        Case "frmrevisi"
            SSTransition1.Transition ssDissolve, ssTransitionImmediate, frmrevisi
            FormhWnd = frmrevisi.hWnd
            Label4.Caption = "Transaksi Revisi Stok"
        
        Case "frmbooking"
            SSTransition1.Transition ssDissolve, ssTransitionImmediate, frmbooking
            FormhWnd = frmbooking.hWnd
            Label4.Caption = "Transaksi Booking"

        Case "frmmstmember"
            SSTransition1.Transition ssDissolve, ssTransitionImmediate, frmmstmember
            FormhWnd = frmmstmember.hWnd
            Label4.Caption = "Master Jenis Member"
    
        Case "frmmstshift"
            SSTransition1.Transition ssDissolve, ssTransitionImmediate, frmmstshift
            FormhWnd = frmmstshift.hWnd
            Label4.Caption = "Master Shift"
        
        Case "LapBarang"
            SSTransition1.Transition ssDissolve, ssTransitionImmediate, LapBarang
            FormhWnd = LapBarang.hWnd
            Label4.Caption = "Laporan Barang"
            
        Case "LapPembelian"
            SSTransition1.Transition ssDissolve, ssTransitionImmediate, LapPembelian
            FormhWnd = LapPembelian.hWnd
            Label4.Caption = "Laporan Pembelian"
            
        Case "LapPakaiLapangan"
            SSTransition1.Transition ssDissolve, ssTransitionImmediate, LapPakaiLapangan
            Label4.Caption = "Laporan Pakai Lapangan"
            
        Case "LapPenjualan"
            SSTransition1.Transition ssDissolve, ssTransitionImmediate, LapPenjualan
            Label4.Caption = "Laporan Penjualan"
            
        Case "buktibooking"
            SSTransition1.Transition ssDissolve, ssTransitionImmediate, buktibooking
            Label4.Caption = "Cetak Bukti Booking"
            
        Case "NotaPakaiLapangan"
            SSTransition1.Transition ssDissolve, ssTransitionImmediate, NotaPakaiLapangan
            Label4.Caption = "Cetak Nota Pakai Lapangan"
            
        Case "frmuser"
            SSTransition1.Transition ssDissolve, ssTransitionImmediate, frmuser
            Label4.Caption = "User"
        
        Case "NotaPenjualan"
            SSTransition1.Transition ssDissolve, ssTransitionImmediate, NotaPenjualan
            Label4.Caption = "Cetak Nota Penjualan"
            
        Case "laprevisistok"
            SSTransition1.Transition ssDissolve, ssTransitionImmediate, laprevisistok
            Label4.Caption = "Cetak Revisi Stok"
    
    End Select
    frmMain.Refresh

    Screen.MousePointer = vbNormal
End Sub

Sub SetMainMenu()
    TaskPanel1.SetIconSize 101, 61
    TaskPanel1.SingleSelection = True
    TaskPanel1.MultiColumn = False
    TaskPanel1.RightToLeft = False
    TaskPanel1.AllowDrag = False
    TaskPanel1.ItemLayout = xtpTaskItemLayoutImages
    TaskPanel1.HotTrackStyle = xtpTaskPanelHighlightItem
    TaskPanel1.VisualTheme = xtpTaskPanelThemeNativeWinXP
    TaskPanel1.Behaviour = xtpTaskPanelBehaviourExplorer
    TaskPanel1.SetGroupOuterMargins 0, 0, 0, 0
    TaskPanel1.SetGroupInnerMargins 7, 0, 2, 0
    TaskPanel1.SetItemInnerMargins 0, 0, 0, 0
    TaskPanel1.SetItemOuterMargins 5, 5, 5, 5
    TaskPanel1.SetMargins 3, 3, 3, 3, 3
    TaskPanel1.ColorSet.BackgroundDark = CLng("&H00404040")
    TaskPanel1.ColorSet.BackgroundLight = CLng("&H00404040")
    TaskPanel1.ColorSet.NormalGroupClient = CLng("&H00404040")
    ImageManager1.Icons.MaskColor = CLng("&H00404040")
    CreateToolBox
    End Sub
    Sub SetMainMenu2()
    TaskPanel1.SetIconSize 101, 61
    TaskPanel1.SingleSelection = True
    TaskPanel1.MultiColumn = False
    TaskPanel1.RightToLeft = False
    TaskPanel1.AllowDrag = False
    TaskPanel1.ItemLayout = xtpTaskItemLayoutImages
    TaskPanel1.HotTrackStyle = xtpTaskPanelHighlightItem
    TaskPanel1.VisualTheme = xtpTaskPanelThemeNativeWinXP
    TaskPanel1.Behaviour = xtpTaskPanelBehaviourExplorer
    TaskPanel1.SetGroupOuterMargins 0, 0, 0, 0
    TaskPanel1.SetGroupInnerMargins 7, 0, 2, 0
    TaskPanel1.SetItemInnerMargins 0, 0, 0, 0
    TaskPanel1.SetItemOuterMargins 5, 5, 5, 5
    TaskPanel1.SetMargins 3, 3, 3, 3, 3
    TaskPanel1.ColorSet.BackgroundDark = CLng("&H00404040")
    TaskPanel1.ColorSet.BackgroundLight = CLng("&H00404040")
    TaskPanel1.ColorSet.NormalGroupClient = CLng("&H00404040")
    ImageManager1.Icons.MaskColor = CLng("&H00404040")
    CreateToolBox2
    End Sub
    
    Sub SetMainMenu3()
    TaskPanel1.SetIconSize 101, 61
    TaskPanel1.SingleSelection = True
    TaskPanel1.MultiColumn = False
    TaskPanel1.RightToLeft = False
    TaskPanel1.AllowDrag = False
    TaskPanel1.ItemLayout = xtpTaskItemLayoutImages
    TaskPanel1.HotTrackStyle = xtpTaskPanelHighlightItem
    TaskPanel1.VisualTheme = xtpTaskPanelThemeNativeWinXP
    TaskPanel1.Behaviour = xtpTaskPanelBehaviourExplorer
    TaskPanel1.SetGroupOuterMargins 0, 0, 0, 0
    TaskPanel1.SetGroupInnerMargins 7, 0, 2, 0
    TaskPanel1.SetItemInnerMargins 0, 0, 0, 0
    TaskPanel1.SetItemOuterMargins 5, 5, 5, 5
    TaskPanel1.SetMargins 3, 3, 3, 3, 3
    TaskPanel1.ColorSet.BackgroundDark = CLng("&H00404040")
    TaskPanel1.ColorSet.BackgroundLight = CLng("&H00404040")
    TaskPanel1.ColorSet.NormalGroupClient = CLng("&H00404040")
    ImageManager1.Icons.MaskColor = CLng("&H00404040")
    CreateToolBox3
    End Sub
    
Function CreateToolboxGroup(Caption As String) As TaskPanelGroup
    Dim Folder As TaskPanelGroup, Pointer As TaskPanelGroupItem

    Set Folder = TaskPanel1.Groups.Add(0, Caption)

    Set CreateToolboxGroup = Folder
End Function
Sub CreateToolBox()
    Dim FolderAppPanes As TaskPanelGroup
    Dim FolderMasterPanes As TaskPanelGroup
    Dim FolderTransaksiPanes As TaskPanelGroup
    Dim laporan As TaskPanelGroup
    Dim keluar As TaskPanelGroup

    Set FolderAppPanes = CreateToolboxGroup("Aplikasi")
        FolderAppPanes.Items.Add 1, "", xtpTaskItemTypeLink, 1 'Log In/Out
        FolderAppPanes.Items.Add 2, "", xtpTaskItemTypeLink, 2 'Change Password
        FolderAppPanes.Items.Add 26, "", xtpTaskItemTypeLink, 45 'User

    Set FolderMasterPanes = CreateToolboxGroup("Master Data")
        FolderMasterPanes.Items.Add 6, "", xtpTaskItemTypeLink, 100 'Master member
        FolderMasterPanes.Items.Add 7, "", xtpTaskItemTypeLink, 34  'Master barang
        FolderMasterPanes.Items.Add 8, "", xtpTaskItemTypeLink, 36  'Master supplier
        FolderMasterPanes.Items.Add 9, "", xtpTaskItemTypeLink, 35  'Master lapangan
        FolderMasterPanes.Items.Add 10, "", xtpTaskItemTypeLink, 21 'Master Jenis Member
        FolderMasterPanes.Items.Add 16, "", xtpTaskItemTypeLink, 20 'Master Shift

    Set FolderTransaksiPanes = CreateToolboxGroup("Transaksi")
        FolderTransaksiPanes.Items.Add 12, "", xtpTaskItemTypeLink, 38  'Transaksi Pembelian
        FolderTransaksiPanes.Items.Add 13, "", xtpTaskItemTypeLink, 39  'Transaksi Penjualan
        FolderTransaksiPanes.Items.Add 14, "", xtpTaskItemTypeLink, 40  'Transaksi pakailapangan
        FolderTransaksiPanes.Items.Add 15, "", xtpTaskItemTypeLink, 41  'Transaksi Revisi stok
        FolderTransaksiPanes.Items.Add 61, "", xtpTaskItemTypeLink, 101 'Transaksi booking
        
    Set laporan = CreateToolboxGroup("Laporan Data")
        laporan.Items.Add 20, "", xtpTaskItemTypeLink, 110  'Laporan Barang
        laporan.Items.Add 21, "", xtpTaskItemTypeLink, 46   'Laporan Pembelian
        laporan.Items.Add 22, "", xtpTaskItemTypeLink, 42   'Laporan PakaiLapangan
        laporan.Items.Add 23, "", xtpTaskItemTypeLink, 43   'Laporan Penjualan
        laporan.Items.Add 86, "", xtpTaskItemTypeLink, 22 'laporan revisi
        
    Set keluar = CreateToolboxGroup("Keluar")
        keluar.Items.Add 3, "", xtpTaskItemTypeLink, 3 'Exit
        

    TaskPanel1.Icons.AddIcons ImageManager1.Icons

    FolderAppPanes.Expanded = False
    FolderMasterPanes.Expanded = False
    FolderTransaksiPanes.Expanded = False
    laporan.Expanded = False
    keluar.Expanded = False
End Sub

Sub CreateToolBox2()
    Dim FolderAppPanes As TaskPanelGroup
    Dim FolderMasterPanes As TaskPanelGroup
    Dim FolderTransaksiPanes As TaskPanelGroup
    Dim laporan As TaskPanelGroup
    Dim keluar As TaskPanelGroup

    Set FolderAppPanes = CreateToolboxGroup("Aplikasi")
        FolderAppPanes.Items.Add 1, "", xtpTaskItemTypeLink, 1 'Log In/Out
        FolderAppPanes.Items.Add 2, "", xtpTaskItemTypeLink, 2 'Change Password
        FolderAppPanes.Items.Add 26, "", xtpTaskItemTypeLink, 45 'User

    Set FolderMasterPanes = CreateToolboxGroup("Master Data")
        FolderMasterPanes.Items.Add 6, "", xtpTaskItemTypeLink, 100 'Master member
        FolderMasterPanes.Items.Add 7, "", xtpTaskItemTypeLink, 34  'Master barang
        FolderMasterPanes.Items.Add 9, "", xtpTaskItemTypeLink, 35  'Master lapangan

    Set FolderTransaksiPanes = CreateToolboxGroup("Transaksi")
        FolderTransaksiPanes.Items.Add 13, "", xtpTaskItemTypeLink, 39  'Transaksi Penjualan
        FolderTransaksiPanes.Items.Add 14, "", xtpTaskItemTypeLink, 40  'Transaksi pakailapangan
        FolderTransaksiPanes.Items.Add 61, "", xtpTaskItemTypeLink, 101 'Transaksi booking
        
    Set laporan = CreateToolboxGroup("Laporan Data")
        laporan.Items.Add 20, "", xtpTaskItemTypeLink, 110  'Laporan Barang
        laporan.Items.Add 22, "", xtpTaskItemTypeLink, 42   'Laporan PakaiLapangan
        laporan.Items.Add 23, "", xtpTaskItemTypeLink, 43   'Laporan Penjualan
        
    Set keluar = CreateToolboxGroup("Keluar")
        keluar.Items.Add 3, "", xtpTaskItemTypeLink, 3 'Exit
        
    TaskPanel1.Icons.AddIcons ImageManager1.Icons

    FolderAppPanes.Expanded = False
    FolderMasterPanes.Expanded = False
    FolderTransaksiPanes.Expanded = False
    laporan.Expanded = False
    keluar.Expanded = False
    End Sub

Sub CreateToolBox3()
    Dim FolderAppPanes As TaskPanelGroup
    Dim laporan As TaskPanelGroup
    Dim keluar As TaskPanelGroup
    
    Set FolderAppPanes = CreateToolboxGroup("Aplikasi")
        FolderAppPanes.Items.Add 1, "", xtpTaskItemTypeLink, 1 'Log In/Out
        
    Set laporan = CreateToolboxGroup("Laporan Data")
        laporan.Items.Add 20, "", xtpTaskItemTypeLink, 110  'Laporan Barang
        laporan.Items.Add 21, "", xtpTaskItemTypeLink, 46   'Laporan Pembelian
        laporan.Items.Add 22, "", xtpTaskItemTypeLink, 42   'Laporan PakaiLapangan
        laporan.Items.Add 23, "", xtpTaskItemTypeLink, 43   'Laporan Penjualan
        
    Set keluar = CreateToolboxGroup("Keluar")
        keluar.Items.Add 3, "", xtpTaskItemTypeLink, 3 'Exit
        
    TaskPanel1.Icons.AddIcons ImageManager1.Icons
    
    FolderAppPanes.Expanded = False
    laporan.Expanded = False
    keluar.Expanded = False
    End Sub
Private Sub HideAllForm()
    Dim frm As Form
    Dim panel As TaskPanelGroup

    For Each frm In Forms
        If frm.Name <> "frmMain" Then frm.Hide
    Next

    For Each panel In TaskPanel1.Groups
        panel.Expanded = False
    Next

End Sub
Sub tutupform()
    Dim frm As Form
    Dim panel As TaskPanelGroup

    For Each frm In Forms
        If frm.Name <> "frmMain" Then frm.Hide
    Next

    For Each panel In TaskPanel1.Groups
        panel.Expanded = False
    Next
End Sub

Private Sub TaskPanel1_ItemClick(ByVal Item As ITaskPanelGroupItem)
    On Error Resume Next
    Select Case Item.ID
    
            '##APLIKASI##
            
        Case 1 'Login/logout

                If MsgBox("Apakah Anda Yakin Untuk LogOut?", vbYesNo + vbQuestion) = vbYes Then
                    HideAllForm
                    laptutup
                    Label4.Caption = ""
                    SSPanel3.Caption = "LOG OFF"
                    frmlogin.Show , Me
                    TaskPanel1.Groups.Clear
                    frmuser.keluar
                    End If
                
            
        Case 2 'Change Password
            frmubahpass.Show 1, Me

        Case 3 'Exit
            If MsgBox("Apakah Anda Yakin Ingin Keluar Aplikasi?", vbYesNo, "Konfirmasi") = vbYes Then
                End
            End If
            
            '##MASTER DATA##
                     
        Case 6  'Master member
            laptutup
            Screen.MousePointer = vbHourglass
            TampilkanForm "formMember"

        Case 7 'Master Barang
            laptutup
            Screen.MousePointer = vbHourglass
            TampilkanForm "frmbarang"

        Case 8 'Master supplier
            laptutup
            Screen.MousePointer = vbHourglass
            TampilkanForm "frmmstsupplier"

        Case 9 'Master lapangan
            laptutup
            Screen.MousePointer = vbHourglass
            TampilkanForm "frmmstlapangan"
        
        Case 10 'Master Jenis Member
            laptutup
            Screen.MousePointer = vbHourglass
            TampilkanForm "frmmstmember"
            
        Case 16 'Master Shift
            laptutup
            Screen.MousePointer = vbHourglass
            TampilkanForm "frmmstshift"
        
        '##TRANSAKSI##
        
        Case 12 'Transaksi Pembelian
            laptutup
            Screen.MousePointer = vbHourglass
            TampilkanForm "frmpembelian"

        Case 13 'Transaksi Penjualan
            laptutup
            Screen.MousePointer = vbHourglass
            TampilkanForm "frmpenjualanlangsung"

        Case 14 'Transaksi pakai lapangan
            laptutup
            Screen.MousePointer = vbHourglass
            TampilkanForm "frmpakailapangan"

        Case 15 'Transaksi Revisi
            laptutup
            Screen.MousePointer = vbHourglass
            TampilkanForm "frmrevisi"

        Case 61 'Transaksi booking
            laptutup
            Screen.MousePointer = vbHourglass
            TampilkanForm "frmbooking"

        
        '##LAPORAN##
        
        Case 20 ' Laporan Barang
            laptutup
            Screen.MousePointer = vbHourglass
            TampilkanForm "LapBarang"
    
        Case 21 ' laporan Pembelian
            frmlappembelian.Show 1, Me
    
        Case 22 ' laporan Pakailapangan
            frmlappakai.Show 1, Me
            
        Case 23 'Laporan Penjualan
            frmlappenjualan.Show 1, Me
            
        Case 86 ' laporan revisi stok
            frmlaprevisi.Show 1, Me
            
            
          '##KEAMANAN##
            
'        Case 24 'Setup Database
'            frmsetdatabase.Show 1, Me
'
'        Case 25 'Profil Futsal
'            frmproffut.Show 1, Me
            
        Case 26 'User
            laptutup
            Screen.MousePointer = vbHourglass
            TampilkanForm "frmuser"
            
    End Select

    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    On Error Resume Next

    If MsgBox("Apakah Anda Yakin Ingin Keluar Aplikasi?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then
        Cancel = 1
        Exit Sub
    End If

    End
End Sub

Sub laptutup()
LapBarang.mati
buktibooking.mati
LapPakaiLapangan.mati
LapPembelian.mati
LapPenjualan.mati
NotaPakaiLapangan.mati
NotaPenjualan.mati
laprevisistok.mati
End Sub
