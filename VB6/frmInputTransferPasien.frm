VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmInputTransferPasien 
   Caption         =   "TRANSFER PASIEN"
   ClientHeight    =   9405
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15225
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9405
   ScaleWidth      =   15225
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
      Left            =   6480
      TabIndex        =   19
      Top             =   4320
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
      Left            =   7920
      TabIndex        =   18
      Top             =   4320
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtpMasuk 
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   1440
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
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
      CustomFormat    =   "dd/MM/yyyy hh:mm"
      Format          =   57475075
      CurrentDate     =   42875
   End
   Begin VB.TextBox txtRuangan 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   4680
      TabIndex        =   10
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox txtNRegistrasi 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtJk 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   6120
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtUmur 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   7800
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtNamaPasien 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   3855
   End
   Begin VB.TextBox txtNrm 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin MSDataListLib.DataCombo dcJenisTransfer 
      Height          =   405
      Left            =   2280
      TabIndex        =   15
      Top             =   2160
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
   Begin MSDataListLib.DataCombo dcDokter 
      Height          =   405
      Left            =   2280
      TabIndex        =   17
      Top             =   3120
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
   Begin MSDataListLib.DataCombo dcRuanganTujuan 
      Height          =   405
      Left            =   2280
      TabIndex        =   20
      Top             =   2640
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
      Caption         =   "Ruangan Tujuan"
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
      TabIndex        =   21
      Top             =   2640
      Width           =   1755
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
      Left            =   240
      TabIndex        =   16
      Top             =   3120
      Width           =   705
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Jenis Transfer"
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
      TabIndex        =   14
      Top             =   2160
      Width           =   1485
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000000&
      Height          =   1935
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
      TabIndex        =   13
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
      TabIndex        =   11
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
      TabIndex        =   9
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
      TabIndex        =   7
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
      TabIndex        =   5
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
      TabIndex        =   3
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
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmInputTransferPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub loadCombo()
    Call loadDataCombo(dcJenisTransfer, rs, "SELECT nJenisTransfer,namaJenisTransfer FROM  jenistransfer where visible=1")
    Call loadDataCombo(dcRuanganTujuan, rs, "SELECT nRuangan,namaRuangan FROM  ruangan where visible=1")
    Call loadDataCombo(dcDokter, rs, "select nPegawai,namaPegawai from pegawai where nJenisPegawai='01' and visible=1")
    
End Sub

Private Sub btnTutup_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    Call loadCombo
    dcJenisTransfer.Text = ""
    dcRuanganTujuan.Text = ""
    dcDokter.Text = ""
    'dcPerawat.Text = ""
    
'    grid.Rows = 1
'    Set rs = Nothing
'    Const setColumn = "nJenisDiagnosa=2500,ICD10=2000,Nama Diagnosa=3500"
'    Call captionGrid("frmInputTransaksiObat", grid, 3, setColumn)
End Sub

Private Sub btnSimpan_Click()
Dim objSave As String
Dim nTransfer As String

    If dcJenisTransfer.Text = "" Then MsgBox "Pilih jenis transfer!", vbInformation, "..:.": Exit Sub
    If dcRuanganTujuan.Text = "" Then MsgBox "Pilih Ruangan Tujuan!", vbInformation, "..:.": Exit Sub
    If txtNrm.Text = "" Then MsgBox "Error loading data, ulangi proses sebelumnya!", vbInformation, "..:.": Exit Sub
    
    ReadRs " select  * from transferpasien"
    nTransfer = rs.RecordCount + 1
    WriteRs "insert into transferpasien (nTransfer,tglTransfer,nRM,nRegistrasi,nRuanganAsal,nRuanganTujuan," & _
            "nJenisTransfer,nDokter,nUser) " & _
            "values " & _
            "('" & nTransfer & "','" & Format_Tgl_Jam(dtpMasuk) & "','" & txtNrm.Text & "','" & txtNRegistrasi.Text & "', " & _
            "('" & txtRuangan.Tag & "','" & dcRuanganTujuan.BoundText & "','" & dcJenisTransfer.BoundText & "', " & _
            "('" & dcDokter.BoundText & "','" & publicNPegawai & "' " & _
            ")"
    MsgBox "Tersimpan!", vbInformation, "..:."
    Unload Me
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

