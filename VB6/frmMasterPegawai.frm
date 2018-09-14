VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmMasterPegawai 
   Caption         =   "Master Pegawai"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14760
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8055
   ScaleWidth      =   14760
   Begin VB.TextBox txtTelepon 
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
      Height          =   405
      Left            =   9000
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   1800
      Width           =   2895
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
      Left            =   10920
      TabIndex        =   43
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton btnBatal 
      Caption         =   "Batal"
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
      Left            =   9480
      TabIndex        =   42
      Top             =   6960
      Width           =   1335
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
      Left            =   8040
      TabIndex        =   41
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox txtAlamat 
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
      Height          =   1005
      Left            =   9000
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   600
      Width           =   4935
   End
   Begin VB.CheckBox chkStatus 
      Caption         =   "Status"
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
      Left            =   8880
      TabIndex        =   39
      Top             =   5760
      Width           =   2295
   End
   Begin VB.TextBox txtKodePos 
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
      Height          =   405
      Left            =   9000
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox txtRT 
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
      Height          =   405
      Left            =   10440
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtRW 
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
      Height          =   405
      Left            =   9000
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtTempatLahir 
      Appearance      =   0  'Flat
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
      Left            =   2280
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2040
      Width           =   4455
   End
   Begin VB.TextBox txtNamaPegawai 
      Appearance      =   0  'Flat
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
      Left            =   2280
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1560
      Width           =   4455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Perempuan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   8
      Top             =   3120
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Laki-Laki"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   7
      Top             =   3120
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.TextBox txtNPegawai 
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
      Left            =   2280
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtNik 
      Appearance      =   0  'Flat
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
      Left            =   2280
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1080
      Width           =   4455
   End
   Begin MSComCtl2.DTPicker dtpTglLahir 
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   2520
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
      Format          =   240713731
      CurrentDate     =   42784
   End
   Begin MSComCtl2.DTPicker dtpTglMasuk 
      Height          =   495
      Left            =   2280
      TabIndex        =   26
      Top             =   5040
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
      Format          =   240713731
      CurrentDate     =   42784
   End
   Begin MSDataListLib.DataCombo dcGolonganDarah 
      Height          =   405
      Left            =   2280
      TabIndex        =   28
      Top             =   3600
      Width           =   975
      _ExtentX        =   1720
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
   Begin MSDataListLib.DataCombo dcAgama 
      Height          =   405
      Left            =   2280
      TabIndex        =   29
      Top             =   4080
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
   Begin MSDataListLib.DataCombo dcPernikahan 
      Height          =   405
      Left            =   2280
      TabIndex        =   30
      Top             =   4560
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
   Begin MSDataListLib.DataCombo dcProvinsi 
      Height          =   405
      Left            =   9000
      TabIndex        =   31
      Top             =   2280
      Width           =   4095
      _ExtentX        =   7223
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
   Begin MSDataListLib.DataCombo dcKota 
      Height          =   405
      Left            =   9000
      TabIndex        =   32
      Top             =   2760
      Width           =   4095
      _ExtentX        =   7223
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
   Begin MSDataListLib.DataCombo dcKecamatan 
      Height          =   405
      Left            =   9000
      TabIndex        =   33
      Top             =   3240
      Width           =   4095
      _ExtentX        =   7223
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
   Begin MSDataListLib.DataCombo dcKelurahan 
      Height          =   405
      Left            =   9000
      TabIndex        =   34
      Top             =   3720
      Width           =   4095
      _ExtentX        =   7223
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
   Begin MSDataListLib.DataCombo dcJenisPegawai 
      Height          =   405
      Left            =   2280
      TabIndex        =   38
      Top             =   5640
      Width           =   4095
      _ExtentX        =   7223
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
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal Masuk :"
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
      TabIndex        =   27
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "Jenis Pegawai :"
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
      TabIndex        =   25
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Caption         =   "Kode Pos :"
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
      Left            =   6840
      TabIndex        =   24
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Caption         =   "RT :"
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
      Left            =   8880
      TabIndex        =   23
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "RW :"
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
      Left            =   6840
      TabIndex        =   22
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Kelurahan :"
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
      Left            =   6840
      TabIndex        =   21
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Kecamatan :"
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
      Left            =   6840
      TabIndex        =   20
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Kota :"
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
      Left            =   6840
      TabIndex        =   19
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Provinsi :"
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
      Left            =   6840
      TabIndex        =   18
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Telepon :"
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
      Left            =   6840
      TabIndex        =   17
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Alamat :"
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
      Left            =   6840
      TabIndex        =   16
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Pernikahan :"
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
      TabIndex        =   15
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Agama :"
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
      TabIndex        =   14
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Golongan Darah :"
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
      TabIndex        =   13
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Tempat Lahir :"
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
      Left            =   360
      TabIndex        =   12
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Nama Pegawai :"
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
      Left            =   360
      TabIndex        =   10
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "NIK :"
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
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "No Pegawai :"
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
      Left            =   720
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal Lahir :"
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
      Left            =   480
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Jenis Kelamin :"
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
      Left            =   480
      TabIndex        =   2
      Top             =   3120
      Width           =   1695
   End
End
Attribute VB_Name = "frmMasterPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub loadCombo()
    Call loadDataCombo(dcGolonganDarah, rs, "SELECT nGolonganDarah,namaGolonganDarah FROM  golongandarah where visible=1")
    Call loadDataCombo(dcAgama, rs, "SELECT nAgama,namaAgama FROM  agama where visible=1")
    Call loadDataCombo(dcPernikahan, rs, "SELECT nAgama,namaAgama FROM  agama where visible=1")
End Sub

Private Sub clear()

     ReadRs "select * from pegawai"
    txtNPegawai.Text = "P" & Format(rs.RecordCount + 1, "0########")
    txtNik.Text = ""
    txtNamaPegawai.Text = ""
    txtTempatLahir.Text = ""
    dtpTglLahir.Value = Now()
    dcGolonganDarah.Text = ""
    dcAgama.Text = ""
    dcPernikahan.Text = ""
    dtpTglMasuk.Value = Now()
    dcJenisPegawai.Text = ""
    txtAlamat.Text = ""
    txtTelepon.Text = ""
    dcProvinsi.Text = ""
    dcKota.Text = ""
    dcKecamatan.Text = ""
    dcKelurahan.Text = ""
    txtRW.Text = ""
    txtRT.Text = ""
    txtKodePos.Text = ""
    chkStatus.Value = vbChecked
End Sub

Private Sub btnBatal_Click()
   Call clear
End Sub

Private Sub btnSimpan_Click()
    MsgBox "Tersimpan!", vbInformation, "..:."
End Sub

Private Sub btnTutup_Click()
    Unload Me
End Sub

Public Sub Form_Load()
    Call clear
    Call loadCombo
End Sub


