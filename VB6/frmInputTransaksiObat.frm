VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmInputTransaksiObat 
   Caption         =   "OBAT"
   ClientHeight    =   9405
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9405
   ScaleWidth      =   15225
   Begin VB.CommandButton cmdTambah 
      Caption         =   "Tambah"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12120
      TabIndex        =   25
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton btnSimpan 
      Caption         =   "Simpan"
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
      Left            =   12120
      TabIndex        =   22
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton btnTutup 
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
      Left            =   13560
      TabIndex        =   21
      Top             =   7920
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtpMasuk 
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   1440
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy hh:mm"
      Format          =   57999363
      CurrentDate     =   42875
   End
   Begin VB.TextBox txtRuangan 
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
      Left            =   4680
      TabIndex        =   11
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox txtNRegistrasi 
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
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtJk 
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
      Left            =   6120
      TabIndex        =   7
      Top             =   600
      Width           =   1575
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
      Left            =   7800
      TabIndex        =   5
      Top             =   600
      Width           =   1575
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
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   3855
   End
   Begin VB.TextBox txtNrm 
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
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VSFlex8LCtl.VSFlexGrid grid 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   14775
      _cx             =   26061
      _cy             =   7858
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
      FormatString    =   $"frmInputTransaksiObat.frx":0000
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
   Begin MSDataListLib.DataCombo dcJenisDiagnosa 
      Height          =   405
      Left            =   240
      TabIndex        =   16
      Top             =   2520
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
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
   Begin MSDataListLib.DataCombo dcIcdx 
      Height          =   405
      Left            =   3480
      TabIndex        =   18
      Top             =   2520
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSDataListLib.DataCombo dcDokter 
      Height          =   405
      Left            =   120
      TabIndex        =   20
      Top             =   8040
      Width           =   6015
      _ExtentX        =   10610
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
   Begin MSDataListLib.DataCombo dcNamaDiagnosa 
      Height          =   405
      Left            =   5040
      TabIndex        =   23
      Top             =   2520
      Width           =   6975
      _ExtentX        =   12303
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
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Nama Diagnosa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      TabIndex        =   24
      Top             =   2160
      Width           =   1665
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Dokter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   19
      Top             =   7680
      Width           =   705
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "ICD10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   17
      Top             =   2160
      Width           =   675
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Jenis Diagnosa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   15
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000000&
      Height          =   975
      Left            =   120
      Top             =   2040
      Width           =   14775
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000000&
      Height          =   1815
      Left            =   120
      Top             =   120
      Width           =   14775
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
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
      Height          =   285
      Left            =   4680
      TabIndex        =   14
      Top             =   1080
      Width           =   930
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   12
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "N Registrasi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
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
      Height          =   285
      Left            =   6120
      TabIndex        =   8
      Top             =   240
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Height          =   285
      Left            =   7800
      TabIndex        =   6
      Top             =   240
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   240
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmInputTransaksiObat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub loadCombo()
    Call loadDataCombo(dcJenisDiagnosa, rs, "SELECT nJenisDiagnosa,namaJenisDiagnosa FROM  jenisdiagnosa where visible=1")
    Call loadDataCombo(dcIcdx, rs, "SELECT nDiagnosa,kodeDiagnosa FROM  diagnosa where visible=1")
    Call loadDataCombo(dcNamaDiagnosa, rs, "SELECT nDiagnosa,namaDiagnosa FROM  diagnosa where visible=1")
    Call loadDataCombo(dcDokter, rs, "select nPegawai,namaPegawai from diagnosa where nJenisPegawai='01' and visible=1")
    'Call loadDataCombo(dcPerawat, rs, "select nPegawai,namaPegawai from pegawai where nJenisPegawai='02' and visible=1")
    
End Sub

Private Sub Command1_Click()

End Sub




Private Sub btnTutup_Click()
    Unload Me
End Sub

Private Sub dcKelas_Change()
    Call dcNamaTindakan_Change
End Sub

Private Sub dcKelas_Click(Area As Integer)
    Call dcNamaTindakan_Change
End Sub

Private Sub dcNamaTindakan_Change()
'    txtTarif.Text = 0
'    ReadRs "select * from hargatotaltindakan where nTindakan ='" & dcNamaTindakan.BoundText & "' and nKelas ='" & dcKelas.BoundText & "'"
'    If rs.RecordCount <> 0 Then
'        txtTarif.Text = rs!totalHarga
'        txtTotal.Text = Val(txtTarif.Text) * Val(txtJml.Text)
'    End If
End Sub

Private Sub dcNamaTindakan_Click(Area As Integer)
    Call dcNamaTindakan_Change
End Sub

Private Sub Form_Load()
    Call loadCombo
    dcNamaTindakan.Text = ""
    dcKelas.Text = ""
    dcDokter.Text = ""
    dcPerawat.Text = ""
    
    grid.Rows = 1
    Set rs = Nothing
    Const setColumn = "nJenisDiagnosa=2500,ICD10=2000,Nama Diagnosa=3500"
    Call captionGrid("frmInputTransaksiObat", grid, 3, setColumn)
End Sub

Private Sub cmdTambah_Click()
Dim baris As Integer

     grid.Rows = grid.Rows + 1
     baris = grid.Rows - 1
     grid.TextMatrix(baris, 0) = baris
     grid.TextMatrix(baris, 1) = dcJenisDiagnosa.BoundText
     grid.TextMatrix(baris, 2) = dcIcdx.Text
     grid.TextMatrix(baris, 3) = dcNamaDiagnosa.BoundText
     
End Sub
Private Sub btnSimpan_Click()
Dim objSave As String

    If grid.Rows = 1 Then
        MsgBox "Isi terlebih dahulu pelayanan!", vbInformation, "..:."
        Exit Sub
    End If
    objSave = ""
    For i = 1 To grid.Rows - 1
        objSave = objSave & ",('" & txtNrm.Text & "','" & txtNRegistrasi.Text & "','" & Format_Tgl_Jam(dtpMasuk.Value) & "'," & _
                "'" & txtRuangan.Tag & "','" & grid.TextMatrix(i, 1) & "','" & grid.TextMatrix(i, 5) & "'," & _
                "'" & grid.TextMatrix(i, 6) & "','" & grid.TextMatrix(i, 3) & "','" & grid.TextMatrix(i, 7) & "',0,0,0," & _
                "'" & dcDokter.BoundText & "','" & dcPerawat.BoundText & "','" & publicNPegawai & "'" & _
                ")"
    Next
    objSave = Right(objSave, Len(objSave) - 1)
    WriteRs "insert into transaksidiagnosa (nRegistrasi,tglDiagnosa,nRuangan,nJenisDiagnosa,nDiagnosa," & _
            "nDokter,nUserID) " & _
            "values " & _
            objSave
    MsgBox "Tersimpan!", vbInformation, "..:."
    Unload Me
End Sub

Private Sub txtJml_Change()
    txtTotal.Text = Val(txtTarif.Text) * Val(txtJml.Text)
End Sub
Public Sub clear()
    txtNRegistrasi.Text = ""
    txtRuangan.Text = ""
    txtNamaPasien.Text = ""
    txtJk.Text = ""
    txtUmur.Text = ""
    dtpMasuk.Value = Now()
End Sub

Private Sub txtNRm_Change()
    Call clear
    ReadRs "select * from pasien where nRm ='" & txtNrm.Text & "'"
    If rs.RecordCount <> 0 Then
        txtNamaPasien.Text = rs!namaPasien
        txtJk.Text = IIf(rs!jenisKelamin = "L", "Laki-laki", "Perempuan")
        txtUmur.Text = DateDiff("YYYY", CDate(rs!tglLahir), Now()) & " tahun " & (DateDiff("M", CDate(rs!tglLahir), Now()) Mod 12) & " bulan"
    End If
    
End Sub

