VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmDaftarPasienLama 
   Caption         =   "Daftar Pasien"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   16965
   WindowState     =   2  'Maximized
   Begin VB.Frame fr2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   7320
      Width           =   15135
      Begin VB.CommandButton btn 
         Caption         =   "Riwayat Catatan Medis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   10440
         TabIndex        =   17
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton btn 
         Caption         =   "Riwayat Diagnosa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   8520
         TabIndex        =   16
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton btn 
         Caption         =   "Riwayat Pelayanan Obat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   6600
         TabIndex        =   15
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton btn 
         Caption         =   "Riwayat Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   4920
         TabIndex        =   14
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton btn 
         Caption         =   "Ubah Data Pasien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   3240
         TabIndex        =   13
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton btn 
         Caption         =   "Registrasi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1560
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Frame fr 
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   16095
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   3
         Left            =   5880
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   2
         Left            =   3960
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   1
         Left            =   1800
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton btn 
         Caption         =   "Cari"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   14760
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
      Begin VSFlex8LCtl.VSFlexGrid grid 
         Height          =   5895
         Left            =   0
         TabIndex        =   1
         Top             =   1080
         Width           =   15855
         _cx             =   27966
         _cy             =   10398
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDaftarPasienLama.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label1 
         Caption         =   "Telepon"
         Height          =   495
         Index           =   2
         Left            =   5880
         TabIndex        =   12
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Alamat"
         Height          =   495
         Index           =   4
         Left            =   3960
         TabIndex        =   10
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Nama Pasien"
         Height          =   495
         Index           =   1
         Left            =   1800
         TabIndex        =   8
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "NRM/NREG"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmDaftarPasienLama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

End Sub

Private Sub btn_Click(Index As Integer)
    Select Case Index
        Case 0
            Call LoadData
        Case 1
            frmRegistrasi.Show
            frmRegistrasi.ZOrder 0
            frmRegistrasi.Form_Load
            frmRegistrasi.txt(0).Text = grid.TextMatrix(grid.Row, 1)
        Case 3 'riwayat Pelayanan transaksi
            frmtransaksipelayanan.Show
            frmtransaksipelayanan.clear
            frmtransaksipelayanan.dcNRegistrasi.Enabled = True
            frmtransaksipelayanan.txtNRm = grid.TextMatrix(grid.Row, 1)
            frmtransaksipelayanan.btnTambah.visible = False
            frmtransaksipelayanan.btnUbah.visible = False
            frmtransaksipelayanan.ZOrder 0
        Case 4 'riwayat Pelayanan Obat
            frmtransaksiObat.Show
            frmtransaksiObat.clear
            frmtransaksiObat.dcNRegistrasi.Enabled = True
            frmtransaksiObat.txtNRm = grid.TextMatrix(grid.Row, 1)
            frmtransaksiObat.ZOrder 0
        Case 5 'riwayat Pelayanan Diagnosa
            frmtransaksiDiagnosa.Show
            frmtransaksiDiagnosa.clear
            frmtransaksiDiagnosa.dcNRegistrasi.Enabled = True
            frmtransaksiDiagnosa.txtNRm = grid.TextMatrix(grid.Row, 1)
            frmtransaksiDiagnosa.ZOrder 0
        Case 6 'riwayat Pelayanan Catatan Medis
            frmtransaksiCatatanMedis.Show
            frmtransaksiCatatanMedis.clear
            frmtransaksiCatatanMedis.dcNRegistrasi.Enabled = True
            frmtransaksiCatatanMedis.txtNRm = grid.TextMatrix(grid.Row, 1)
            frmtransaksiCatatanMedis.ZOrder 0
    End Select
End Sub

Public Sub LoadData()
    ReadRs "SELECT nRM as NRM, namaPasien as 'Nama Pasien', tglLahir as 'Tanggal Lahir' FROM pasien " & _
           "where " & _
           "nRM like '%" & txt(0).Text & "%' and " & _
           "namaPasien like '%" & txt(1).Text & "%'  "
    Const setColumn = "NRM=1700,Nama Pasien=4000,Tanggal Lahir=2000"
    Call isiGrid("daftarPasienLama", grid, rs, setColumn)
    
End Sub

Private Sub Form_Activate()
    fr.Move 100, 100, Me.Width - 100, Me.Height - 1000 - fr2.Height
    fr2.Move fr.Left, fr.Top + fr.Height, fr.Width, fr2.Height
    grid.Move 100, grid.Top, fr.Width, fr.Height - grid.Top - 100
    
End Sub

Private Sub Form_Load()
    
    For i = 0 To 3
        txt(i).Text = ""
    Next
    
    Call LoadData
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call LoadData
End Sub
