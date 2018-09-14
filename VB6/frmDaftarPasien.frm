VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDaftarPasien 
   Caption         =   "Daftar Pasien"
   ClientHeight    =   9225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17460
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
   ScaleHeight     =   9225
   ScaleWidth      =   17460
   WindowState     =   2  'Maximized
   Begin VB.Frame fr2 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   7080
      Width           =   17535
      Begin VB.CommandButton btn 
         Caption         =   "Tutup"
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
         Index           =   9
         Left            =   15360
         TabIndex        =   26
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton btn 
         Caption         =   "Konsul"
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
         Left            =   14040
         TabIndex        =   22
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton btn 
         Caption         =   "Cetak"
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
         Left            =   12720
         TabIndex        =   21
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton btn 
         Caption         =   "Detail Catatan Medis"
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
         Index           =   8
         Left            =   11160
         TabIndex        =   25
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton btn 
         Caption         =   "Detail Diagnosa"
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
         Index           =   7
         Left            =   9840
         TabIndex        =   24
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton btn 
         Caption         =   "Detail Obat"
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
         Left            =   8520
         TabIndex        =   23
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton btn 
         Caption         =   "Rekam Medis"
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
         Left            =   6960
         TabIndex        =   20
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton btn 
         Caption         =   "Detail Layanan"
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
         Left            =   5640
         TabIndex        =   19
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   4
         Left            =   7680
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   3
         Left            =   5760
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   2
         Left            =   3840
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   1
         Left            =   1680
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton btn 
         Caption         =   "Ubah Registrasi"
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
         Left            =   4320
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Instalasi"
         Height          =   495
         Index           =   4
         Left            =   3840
         TabIndex        =   18
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Dokter"
         Height          =   495
         Index           =   3
         Left            =   7680
         TabIndex        =   17
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Ruangan"
         Height          =   495
         Index           =   2
         Left            =   5760
         TabIndex        =   16
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Nama Pasien"
         Height          =   495
         Index           =   1
         Left            =   1680
         TabIndex        =   15
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "NRM/NREG"
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   14
         Top             =   0
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
         Left            =   6600
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin VSFlex8LCtl.VSFlexGrid grid 
         Height          =   5895
         Left            =   0
         TabIndex        =   1
         Top             =   720
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
         FormatString    =   $"frmDaftarPasien.frx":0000
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
      Begin MSComCtl2.DTPicker dt 
         Height          =   495
         Index           =   0
         Left            =   1920
         TabIndex        =   2
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   239403011
         CurrentDate     =   31262
      End
      Begin MSComCtl2.DTPicker dt 
         Height          =   495
         Index           =   1
         Left            =   4440
         TabIndex        =   3
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   239403011
         CurrentDate     =   31262
      End
      Begin VB.Label Label1 
         Caption         =   "Periode Registrasi :"
         Height          =   375
         Index           =   5
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "s/d"
         Height          =   495
         Index           =   6
         Left            =   4080
         TabIndex        =   4
         Top             =   120
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmDaftarPasien"
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
            SaveSetting "MedApp", Me.Name, "tglAwal", dt(0).Value
            SaveSetting "MedApp", Me.Name, "tglAkhir", dt(1).Value
        Case 1
            MsgBox "belum beres"
            Exit Sub
            frmRegistrasi.Show
            frmRegistrasi.ZOrder 0
            frmRegistrasi.Form_Load
            frmRegistrasi.txt(0).Text = grid.TextMatrix(grid.Row, 1)
        Case 2 'Riwayat detail Pelayanan
            frmtransaksipelayanan.Show
            frmtransaksipelayanan.clear
            frmtransaksipelayanan.dcNRegistrasi.Enabled = False
            frmtransaksipelayanan.txtNrm.Text = grid.TextMatrix(grid.Row, 3)
            frmtransaksipelayanan.dcNRegistrasi.Text = grid.TextMatrix(grid.Row, 1)
            frmtransaksipelayanan.dcRuangan.Tag = grid.TextMatrix(grid.Row, 11)
            frmtransaksipelayanan.dcRuangan.Text = grid.TextMatrix(grid.Row, 7)
            frmtransaksipelayanan.ZOrder 0
        Case 3 'Rekam Medis
            frmtransaksipelayanan.Show
            frmtransaksipelayanan.clear
            frmtransaksipelayanan.dcNRegistrasi.Enabled = True
            frmtransaksipelayanan.txtNrm.Text = grid.TextMatrix(grid.Row, 3)
            'frmtransaksipelayanan.dcNRegistrasi.Text = grid.TextMatrix(grid.Row, 1)
            'frmtransaksipelayanan.dcRuangan.Text = grid.TextMatrix(grid.Row, 7)
            frmtransaksipelayanan.btnTambah.visible = False
            frmtransaksipelayanan.btnUbah.visible = False
            frmtransaksipelayanan.ZOrder 0
        Case 5
            frmInputTransferPasien.Show
            frmInputTransferPasien.clear
            'frmInputTransferPasien.dcNRegistrasi.Enabled = False
            frmInputTransferPasien.txtNrm.Text = grid.TextMatrix(grid.Row, 3)
            frmInputTransferPasien.txtNRegistrasi.Text = grid.TextMatrix(grid.Row, 1)
            frmInputTransferPasien.txtRuangan.Tag = grid.TextMatrix(grid.Row, 11)
            frmInputTransferPasien.txtRuangan.Text = grid.TextMatrix(grid.Row, 7)
            frmInputTransferPasien.ZOrder 0
            
        
       Case 6  'Riwayat detail Obat
            frmtransaksiObat.Show
            frmtransaksiObat.clear
            frmtransaksiObat.dcNRegistrasi.Enabled = False
            frmtransaksiObat.txtNrm.Text = grid.TextMatrix(grid.Row, 3)
            frmtransaksiObat.dcNRegistrasi.Text = grid.TextMatrix(grid.Row, 1)
            frmtransaksiObat.ZOrder 0
       Case 7 'Riwayat detail Diagnosa
            frmtransaksiDiagnosa.Show
            frmtransaksiDiagnosa.clear
            frmtransaksiDiagnosa.dcNRegistrasi.Enabled = False
            frmtransaksiDiagnosa.txtNrm.Text = grid.TextMatrix(grid.Row, 3)
            frmtransaksiDiagnosa.dcNRegistrasi.Text = grid.TextMatrix(grid.Row, 1)
            frmtransaksiDiagnosa.ZOrder 0
       Case 8 'Riwayat detail catatan Medis
            frmtransaksiCatatanMedis.Show
            frmtransaksiCatatanMedis.clear
            frmtransaksiCatatanMedis.dcNRegistrasi.Enabled = False
            frmtransaksiCatatanMedis.txtNrm.Text = grid.TextMatrix(grid.Row, 3)
            frmtransaksiCatatanMedis.dcNRegistrasi.Text = grid.TextMatrix(grid.Row, 1)
            frmtransaksiCatatanMedis.ZOrder 0
        Case 9
            Unload Me
    End Select
End Sub

Public Sub LoadData()
    ReadRs "select reg.nRegistrasi as NoRegistrasi,reg.tglRegistrasi as Tanggal,reg.nRM as NRM,pas.namaPasien as NamaPasien, " & _
           "kl.namaKelompok as Jenis, " & _
           "ins.namaInstalasi as Instalasi,rg.namaRuangan as Ruangan,pg.namaPegawai  as Dokter,asl.namaAsal Asal,kls.namaKelas as Kelas,reg.nRuangan " & _
           "from registrasi reg,pasien pas,ruangan rg,pegawai pg,instalasi ins,kelas kls,asal asl,kelompok kl " & _
           "where reg.nRM=pas.nRM and reg.nRuangan=rg.nRuangan and reg.nDokter=pg.npegawai and reg.nInstalasi=ins.nInstalasi  and " & _
           "reg.nKelas=kls.nKelas and reg.nAsal=asl.nAsal and reg.nKelompok=kl.nKelompok  " & _
           "and " & _
           "(reg.nRegistrasi like '%" & txt(0).Text & "%' or reg.nRM like '%" & txt(0).Text & "%') and " & _
           "reg.tglRegistrasi between '" & Format_tgl(dt(0).Value) & " 00:00' and '" & Format_tgl(dt(1).Value) & " 23:59' and " & _
           "pas.namaPasien like '%" & txt(1).Text & "%' and " & _
           "ins.namaInstalasi like '%" & txt(2).Text & "%' and " & _
           "rg.namaRuangan like '%" & txt(3).Text & "%' and " & _
           "pg.namaPegawai  like '%" & txt(4).Text & "%' and " & _
           "reg.nRuangan = '" & MDIForm1.dcRuangan.BoundText & "'"
    Const setColumn = "NoRegistrasi=1500,Tanggal=1500,NRM=1500,NamaPasien=2500," & _
                      "Jenis=1000,Instalasi=1500,Ruangan=1500,Dokter=2000,Asal=1500,Kelas=1500,nRuangan=0"
    Call isiGrid("daftarPasien", grid, rs, setColumn)
    

    
End Sub

Private Sub Form_Activate()
    fr.Move 100, 100, Me.Width - 100, Me.Height - 1000 - fr2.Height
    fr2.Move fr.Left, fr.Top + fr.Height, fr.Width, fr2.Height
    grid.Move 100, grid.Top, fr.Width, fr.Height - grid.Top - 100
End Sub

Private Sub Form_Load()
    
    For i = 0 To 4
        txt(i).Text = ""
    Next
    
    'If dt(0).Value < DateAdd("yyyy", -1, Now()) Then
    '    dt(0).Value = Now()
    'Else
    Dim tglAwal As String
    tglAwal = (GetSetting("MedApp", Me.Name, "tglAwal"))
    Dim tglAkhir As String
    tglAkhir = (GetSetting("MedApp", Me.Name, "tglAkhir"))
    'End If
    dt(0).Value = CDate(IIf(tglAwal = "", Now(), tglAwal))
    dt(1).Value = CDate(IIf(tglAkhir = "", Now(), tglAkhir))
    
    Call LoadData
    
    grid.Move 100, grid.Top, Screen.Width - MDIForm1.tv.Width - 200, grid.Height
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call LoadData
End Sub
