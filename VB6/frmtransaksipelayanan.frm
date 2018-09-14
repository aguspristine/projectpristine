VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmtransaksipelayanan 
   Caption         =   ".: Transaksi Pelayanan"
   ClientHeight    =   9240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17670
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   17670
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   21
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
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
      Left            =   14160
      TabIndex        =   20
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton btnCetak 
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
      Left            =   12480
      TabIndex        =   19
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton btnUbah 
      Caption         =   "Ubah"
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
      Left            =   10800
      TabIndex        =   18
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton btnTambah 
      Caption         =   "Tambah"
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
      Left            =   9120
      TabIndex        =   17
      Top             =   8160
      Width           =   1575
   End
   Begin VB.TextBox txtTglMasuk 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox txtUmur 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   8
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtJenisKelamin 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtNamaPasien 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtNRm 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VSFlex8LCtl.VSFlexGrid grid 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   15495
      _cx             =   27331
      _cy             =   10186
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
      FormatString    =   $"frmtransaksipelayanan.frx":0000
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
   Begin MSDataListLib.DataCombo dcNRegistrasi 
      Height          =   405
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo dcRuangan 
      Height          =   405
      Left            =   4920
      TabIndex        =   15
      Top             =   1320
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo dcKelas 
      Height          =   405
      Left            =   8280
      TabIndex        =   16
      Top             =   1320
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label8 
      Caption         =   "Kelas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   13
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "Ruangan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Tanggal Masuk"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Umur"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   9
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Jenis Kelamin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "NRegistrasi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "NRM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmtransaksipelayanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LoadData()
       
           ReadRs "SELECT " & _
                  "transaksipelayanan.nRegistrasi, " & _
                  "transaksipelayanan.tglTransaksi, " & _
                  "tindakanpelayanan.NamaTindakan, " & _
                  "transaksipelayanan.hargaSatuan, " & _
                  "transaksipelayanan.jmlPelayanan, " & _
                  "transaksipelayanan.totalBiaya, " & _
                  "kelas.namaKelas, " & _
                  "ruangan.namaRuangan, " & _
                  "`user`.namaPegawai AS 'Dokter', " & _
                  "`user`.namaPegawai AS 'User', " & _
                  "transaksipelayanan.nStruk " & _
                  "From " & _
                  "transaksipelayanan " & _
                  "INNER JOIN tindakanpelayanan ON tindakanpelayanan.nTindakan = transaksipelayanan.nTindakan " & _
                  "INNER JOIN kelas ON kelas.nkelas = transaksipelayanan.nKelas " & _
                  "INNER JOIN pegawai ON pegawai.nPegawai = transaksipelayanan.nDokter " & _
                  "INNER JOIN pegawai AS `user` ON `user`.nPegawai = transaksipelayanan.nUser " & _
                  "INNER JOIN ruangan ON ruangan.nRuangan = transaksipelayanan.nRuangan " & _
                  "where transaksipelayanan.nRegistrasi like '%" & dcNRegistrasi.Text & "%' " & _
                  "and transaksipelayanan.nRm ='" & txtNRm.Text & "'" & _
                  "and transaksipelayanan.nRuangan like '%" & dcRuangan.BoundText & "%'" & _
                  "and transaksipelayanan.nKelas like '%" & dcKelas.BoundText & "%'"
           
           Call isiGrid("frmtransaksipelayanan", grid, rs, "nRegistrasi=1500,Tanggal=2200,NamaTindakan=3000," & _
                        "Harga=1700,Qty=800,Total=1700,Kelas=1500,Ruangan=1800,Dokter=1500,User=1500,Struk=1500")

End Sub
Private Sub btnUbah_Click()
    If grid.TextMatrix(grid.Row, 11) <> "" Then
        MsgBox "Pelayanan sudah di bayar tidak bisa di ubah !", vbInformation, "..:."
        Exit Sub
    End If
End Sub

Private Sub btn_Click(Index As Integer)

End Sub

Private Sub btnTutup_Click()

End Sub

Private Sub btnTambah_Click()
    frmInputTransaksiPelayanan.Show
    frmInputTransaksiPelayanan.txtNRm.Text = txtNRm.Text
    frmInputTransaksiPelayanan.txtNRegistrasi.Text = dcNRegistrasi.Text
    If dcRuangan.Text = dcRuangan.BoundText Then
        frmInputTransaksiPelayanan.txtRuangan.Tag = dcRuangan.Tag
    Else
        frmInputTransaksiPelayanan.txtRuangan.Tag = dcRuangan.BoundText
    End If
    frmInputTransaksiPelayanan.txtRuangan.Text = dcRuangan.Text
    
End Sub


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Call dcNRegistrasi_Change
End Sub

Private Sub dcKelas_Change()
    Call LoadData
End Sub

Private Sub dcKelas_Click(Area As Integer)
    dcKelas_Change
End Sub

Private Sub dcNRegistrasi_Change()
    If dcNRegistrasi.Text <> "" Then
        dcRuangan.Text = ""
        dcKelas.Text = ""
        ReadRs "select * from registrasi where nRegistrasi='" & dcNRegistrasi.Text & "'"
        If rs.RecordCount <> 0 Then
            txtTglMasuk.Text = rs!tglRegistrasi
            
        End If
        Call loadDataCombo(dcRuangan, rs, "SELECT distinct transaksipelayanan.nRuangan,ruangan.namaRuangan From transaksipelayanan " & _
                                          "INNER JOIN ruangan ON ruangan.nRuangan = transaksipelayanan.nRuangan " & _
                                          "where  transaksipelayanan.nRegistrasi ='" & dcNRegistrasi.Text & "'")
        Call loadDataCombo(dcKelas, rs, "SELECT distinct transaksipelayanan.nKelas,kelas.namaKelas From transaksipelayanan " & _
                                          "INNER JOIN kelas ON kelas.nKelas = transaksipelayanan.nKelas " & _
                                          "where  transaksipelayanan.nRegistrasi ='" & dcNRegistrasi.Text & "'")
    End If
    Call LoadData
End Sub



Private Sub dcNRegistrasi_Click(Area As Integer)
    dcNRegistrasi_Change
End Sub

Private Sub dcRuangan_Change()
    If dcRuangan.Text <> "" Then
        dcKelas.Text = ""
        Call loadDataCombo(dcKelas, rs, "SELECT distinct transaksipelayanan.nKelas,kelas.namaKelas From transaksipelayanan " & _
                                          "INNER JOIN kelas ON kelas.nKelas = transaksipelayanan.nKelas " & _
                                          "where  transaksipelayanan.nRegistrasi ='" & dcNRegistrasi.Text & "'  and nRuangan='" & dcRuangan.BoundText & "'")
    End If
    Call LoadData
End Sub

Private Sub dcRuangan_Click(Area As Integer)
    dcRuangan_Change
End Sub

Private Sub Form_Load()
    'Call LoadData
    Call clear
    'grid.Width = Me.Width
    grid.Move 100, grid.Top, Screen.Width - MDIForm1.tv.Width - 200, grid.Height
End Sub

Public Sub clear()
    dcNRegistrasi.Text = ""
    dcRuangan.Text = ""
    dcKelas.Text = ""
    txtTglMasuk.Text = ""
End Sub

Private Sub txtNRm_Change()
    Call clear
    ReadRs "select * from pasien where nRm ='" & txtNRm.Text & "'"
    If rs.RecordCount <> 0 Then
        txtNamaPasien.Text = rs!namaPasien
        txtJenisKelamin.Text = IIf(rs!jenisKelamin = "L", "Laki-laki", "Perempuan")
        txtUmur.Text = DateDiff("YYYY", CDate(rs!tglLahir), Now()) & " tahun " & (DateDiff("M", CDate(rs!tglLahir), Now()) Mod 12) & " bulan"
    End If
    
    Call loadDataCombo(dcNRegistrasi, rs, "select distinct nRegistrasi,nRegistrasi from transaksipelayanan where nRM ='" & txtNRm.Text & "'")
    
    Call LoadData
End Sub
