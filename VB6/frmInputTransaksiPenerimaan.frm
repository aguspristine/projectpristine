VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmInputTransaksiPenerimaan 
   Caption         =   "TRANSAKSI PENERIMAAN"
   ClientHeight    =   12615
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12615
   ScaleWidth      =   15225
   Begin MSComCtl2.DTPicker dtpTerima 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy hh:mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   38
      Top             =   720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy hh:mm"
      Format          =   174456835
      CurrentDate     =   42875
   End
   Begin VB.TextBox txtSubTotal 
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
      Left            =   11040
      TabIndex        =   36
      Top             =   9240
      Width           =   3855
   End
   Begin VB.TextBox txtTotalAll 
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
      Left            =   11040
      TabIndex        =   35
      Top             =   10680
      Width           =   3855
   End
   Begin VB.TextBox txtTotalPpn 
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
      Left            =   11040
      TabIndex        =   34
      Top             =   10200
      Width           =   3855
   End
   Begin VB.TextBox txtTotalDiskon 
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
      Left            =   11040
      TabIndex        =   33
      Top             =   9720
      Width           =   3855
   End
   Begin VB.TextBox txtTotal 
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
      Left            =   3480
      TabIndex        =   32
      Top             =   3840
      Width           =   2655
   End
   Begin VB.TextBox txtPpn 
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
      Left            =   1920
      TabIndex        =   31
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox txtDiskon 
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
      Left            =   360
      TabIndex        =   30
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox txtTarif 
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
      Left            =   8040
      TabIndex        =   29
      Top             =   3000
      Width           =   1815
   End
   Begin MSDataListLib.DataCombo dcNamaDetailJenisBarang 
      Height          =   405
      Left            =   4800
      TabIndex        =   25
      Top             =   720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   714
      _Version        =   393216
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
      TabIndex        =   18
      Top             =   11280
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
      TabIndex        =   17
      Top             =   11280
      Width           =   1335
   End
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
      Height          =   375
      Left            =   9360
      TabIndex        =   15
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox txtJml 
      Alignment       =   1  'Right Justify
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
      Left            =   9960
      TabIndex        =   12
      Text            =   "1"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtNRegistrasi 
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
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
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
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VSFlex8LCtl.VSFlexGrid grid 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   4560
      Width           =   14655
      _cx             =   25850
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
      FormatString    =   $"frmInputTransaksiPenerimaan.frx":0000
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
   Begin MSDataListLib.DataCombo dcNamaSuplier 
      Height          =   405
      Left            =   8880
      TabIndex        =   26
      Top             =   720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   714
      _Version        =   393216
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
   Begin MSDataListLib.DataCombo dcNamaBarang 
      Height          =   405
      Left            =   360
      TabIndex        =   27
      Top             =   3000
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   714
      _Version        =   393216
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
   Begin MSDataListLib.DataCombo dcSatuan 
      Height          =   405
      Left            =   5880
      TabIndex        =   28
      Top             =   3000
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   714
      _Version        =   393216
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
   Begin MSDataListLib.DataCombo dcNamaruangan 
      Height          =   405
      Left            =   4800
      TabIndex        =   37
      Top             =   1800
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   714
      _Version        =   393216
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
   Begin MSComCtl2.DTPicker dtpFaktur 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy hh:mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   39
      Top             =   1800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy hh:mm"
      Format          =   174391299
      CurrentDate     =   42875
   End
   Begin MSComCtl2.DTPicker dtpkadaluarsa 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy hh:mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   40
      Top             =   3840
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy hh:mm"
      Format          =   174391299
      CurrentDate     =   42875
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Tgl Kadaluarsa"
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
      Index           =   2
      Left            =   6240
      TabIndex        =   41
      Top             =   3480
      Width           =   1560
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Total"
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
      Index           =   3
      Left            =   9840
      TabIndex        =   24
      Top             =   10680
      Width           =   540
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "PPN"
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
      Index           =   2
      Left            =   9840
      TabIndex        =   23
      Top             =   10200
      Width           =   435
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Diskon"
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
      Index           =   1
      Left            =   9840
      TabIndex        =   22
      Top             =   9720
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "PPN"
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
      Index           =   1
      Left            =   1920
      TabIndex        =   21
      Top             =   3480
      Width           =   435
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Diskon"
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
      Index           =   1
      Left            =   360
      TabIndex        =   20
      Top             =   3480
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Tgl Terima"
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
      Index           =   1
      Left            =   2280
      TabIndex        =   19
      Top             =   360
      Width           =   1170
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Sub Total"
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
      Index           =   0
      Left            =   9840
      TabIndex        =   16
      Top             =   9240
      Width           =   1020
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Total"
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
      TabIndex        =   14
      Top             =   3480
      Width           =   540
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah"
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
      Index           =   0
      Left            =   9960
      TabIndex        =   13
      Top             =   2640
      Width           =   765
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Harga"
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
      Index           =   0
      Left            =   8160
      TabIndex        =   11
      Top             =   2640
      Width           =   630
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Satuan"
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
      Left            =   5880
      TabIndex        =   10
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Nama Barang"
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
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   1425
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000000&
      Height          =   1935
      Left            =   240
      Top             =   2520
      Width           =   14655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000000&
      Height          =   2175
      Left            =   240
      Top             =   120
      Width           =   14655
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
      Left            =   4920
      TabIndex        =   8
      Top             =   1440
      Width           =   930
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Tgl Faktur"
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
      Index           =   0
      Left            =   2280
      TabIndex        =   7
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "No Faktur"
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
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Jenis Barang"
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
      Left            =   4800
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nama Suplier"
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
      Left            =   8880
      TabIndex        =   3
      Top             =   360
      Width           =   1425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "No Terima"
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
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   1125
   End
End
Attribute VB_Name = "frmInputTransaksiPenerimaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub loadCombo()
    Call loadDataCombo(dcNamaDetailJenisBarang, rs, "SELECT nDetailJenisBarang,namaDetailJenisBarang FROM  detailjenisbarang where visible=1")
    Call loadDataCombo(dcNamaSuplier, rs, "SELECT nRekanan,namaRekanan FROM rekanan where visible=1")
    Call loadDataCombo(dcNamaBarang, rs, "SELECT nBarang,namaBarang FROM  barang where visible=1")
    Call loadDataCombo(dcSatuan, rs, "SELECT nSatuan,namaSatuan FROM  satuanbarang where visible=1")
    Call loadDataCombo(dcNamaruangan, rs, "select nRuangan,namaRuangan from ruangan where visible=1")
End Sub
Private Sub btnTutup_Click()
    Unload Me
End Sub
Private Sub dcNamaDetailJenisBarang_Change()
    dcNamaBarang.Text = ""
    Call loadDataCombo(dcNamaBarang, rs, "SELECT nBarang,namaBarang FROM barang where ndetailJenisBarang = '" & dcNamaDetailJenisBarang.BoundText & "' and visible=1")
End Sub
Private Sub dcNamaDetailJenisBarang_Click(Area As Integer)
    Call dcNamaDetailJenisBarang_Change
End Sub
Private Sub dcNamaBarang_Change()
    dcSatuan.Text = ""
    ReadRs "select br.nBarang,br.namaBarang,sb.nSatuan,sb.namaSatuan from barang as br " & _
           "inner join satuanbarang as sb on sb.nSatuan = br.nSatuan where nBarang ='" & dcNamaBarang.BoundText & "'"
    If rs.RecordCount <> 0 Then
        dcSatuan.Text = rs!namaSatuan
        dcSatuan.BoundText = rs!nSatuan
        txtTarif.Text = 0
        txtDiskon.Text = 0
        txtPpn.Text = 0
        txtTotal.Text = 0
    End If
End Sub
Private Sub dcNamaBarang_Click(Area As Integer)
    Call dcNamaBarang_Change
End Sub
Private Sub Form_Load()
    Call loadCombo
     dcNamaDetailJenisBarang.Text = ""
     dcNamaSuplier.Text = ""
     dcSatuan.Text = ""
     dcNamaBarang.Text = ""
     dcNamaruangan.Text = ""
     dtpFaktur.Value = Now()
     dtpTerima.Value = Now()
     dtpkadaluarsa.Value = Now()
    kode = getNewNumberWithDate1("transaksistruk", "nStruk", "tglStruk", "", Format_tgl(Now))
    nFaktur = getNewNumberWithDate1("transaksistruk", "noFaktur", "tglStruk", "", Format_tgl(Now))
    txtNrm.Text = "RS/" + kode
    txtNRegistrasi.Text = "F-" + nFaktur
    grid.Rows = 1
    Set rs = Nothing
    Const setColumn = "nBarang=1000,NamaBarang=3500,nSatuan=0,Satuan=2000,Harga=2000," & _
                      "Jumlah=1500,Diskon=2500,PPN=2500,Total=2500,TglKadaluarsa=1500"
    Call captionGrid("frmInputTransaksiPenerimaan", grid, 10, setColumn)
End Sub
Private Sub cmdTambah_Click()
Dim baris As Integer
Dim tgl As String
'tgl =  dtpkadaluarsa.Value
     grid.Rows = grid.Rows + 1
     baris = grid.Rows - 1
     grid.TextMatrix(baris, 0) = baris
     grid.TextMatrix(baris, 1) = dcNamaBarang.BoundText
     grid.TextMatrix(baris, 2) = dcNamaBarang.Text
     grid.TextMatrix(baris, 3) = dcSatuan.BoundText
     grid.TextMatrix(baris, 4) = dcSatuan.Text
     grid.TextMatrix(baris, 5) = txtTarif.Text
     grid.TextMatrix(baris, 6) = txtJml.Text
     grid.TextMatrix(baris, 7) = txtDiskon.Text
     grid.TextMatrix(baris, 8) = txtPpn.Text
     grid.TextMatrix(baris, 9) = txtTotal.Text
     grid.TextMatrix(baris, 10) = Format(dtpkadaluarsa.Value, "yyyy-MM-dd hh:mm")
     
     txtSubTotal.Text = 0
     txtTotalDiskon.Text = 0
     txtTotalPpn.Text = 0
     txtTotalAll.Text = 0
     For i = 1 To baris
        txtSubTotal.Text = Val(txtSubTotal.Text) + grid.TextMatrix(i, 5)
        txtTotalDiskon.Text = Val(txtTotalDiskon.Text) + grid.TextMatrix(i, 7)
        txtTotalPpn.Text = Val(txtTotalPpn.Text) + grid.TextMatrix(i, 8)
        txtTotalAll.Text = Val(txtSubTotal.Text) + Val(txtTotalPpn.Text) - Val(txtTotalDiskon.Text)
     Next
End Sub
Private Sub btnSimpan_Click()
'On Error GoTo hell
Dim objSave1, objSave2, objSave3, objSave4, keterangan, jenisTrans As String
Dim qtyAwal, qtyMasuk, qtyKeluar, qtyAkhir As Integer
    keterangan = "Penerimaan Barang Suplier  " + "No Terima : " + txtNrm.Text
    If dcNamaDetailJenisBarang.Text = "" Then
        MsgBox "Jenis Barang Tidak Boleh Kosong!", vbInformation, "..:."
        Exit Sub
    End If
    If dcNamaSuplier.Text = "" Then
        MsgBox "Nama Suplier Tidak Boleh Kosong!", vbInformation, "..:."
        Exit Sub
    End If
    If dcNamaruangan.Text = "" Then
        MsgBox "Nama Ruangan Tidak Boleh Kosong!", vbInformation, "..:."
        Exit Sub
    End If
    If grid.Rows = 1 Then
        MsgBox "Isi barang Terlebih dahulu!", vbInformation, "..:."
        Exit Sub
    End If
    cn.BeginTrans
    objSave1 = objSave1 & "('" & txtNrm.Text & "','" & txtNRegistrasi.Text & "','" & Format_Tgl_Jam(dtpTerima.Value) & "','" & Format_Tgl_Jam(dtpFaktur.Value) & "', " & _
                "'" & dcNamaruangan.BoundText & "',001,'" & dcNamaSuplier.BoundText & "'," & _
                "'" & publicNPegawai & "','1')"
    objSave2 = ""
    objSave3 = ""
    objSave4 = ""
    For i = 1 To grid.Rows - 1
    
        nStrukDetail = Format(getNewNumber("transaksistrukdetail", "nStrukDetail", ""), "0#########")
        objSave2 = objSave2 & "('" & nStrukDetail & "','" & txtNrm.Text & "','" & Format_Tgl_Jam(grid.TextMatrix(i, 10)) & "'," & _
                    "'" & grid.TextMatrix(i, 1) & "','" & grid.TextMatrix(i, 3) & "'," & _
                    "'" & grid.TextMatrix(i, 5) & "','" & grid.TextMatrix(i, 6) & "'," & _
                    "'" & grid.TextMatrix(i, 7) & "','" & grid.TextMatrix(i, 8) & "','" & publicNPegawai & "','1')"

        nTStok = Format(getNewNumber("transaksistok", "nTStok", ""), "0#########")
        objSave3 = objSave3 & "('" & nTStok & "','" & grid.TextMatrix(i, 1) & "'," & _
                   "'" & grid.TextMatrix(i, 3) & "','" & grid.TextMatrix(i, 5) & "'," & _
                   "'" & grid.TextMatrix(i, 6) & "','" & grid.TextMatrix(i, 7) & "'," & _
                   "'" & grid.TextMatrix(i, 8) & "','" & Format_Tgl_Jam(grid.TextMatrix(i, 10)) & "'," & _
                   "'" & txtNrm.Text & "','" & dcNamaruangan.BoundText & "','" & publicNPegawai & "','1')"
                   
        nKartuStok = Format(getNewNumber("transaksikartustok", "nKartuStok", ""), "0#########")
        
        ReadRs "select sum(qtystok) as qtystok from transaksistok where nBarang = '" & grid.TextMatrix(i, 1) & "' and nRuangan= '" & dcNamaruangan.BoundText & "' and visible='1'"
        If IsNull(rs!qtystok) Then
            qtyAwal = 0
            qtyMasuk = grid.TextMatrix(i, 6)
            qtyKeluar = 0
            qtyAkhir = qtyAwal + grid.TextMatrix(i, 6)
        Else
            qtyAwal = rs!qtystok
            qtyMasuk = grid.TextMatrix(i, 6)
            qtyKeluar = 0
            qtyAkhir = qtyAwal + grid.TextMatrix(i, 6)
        End If
        objSave4 = objSave4 & ",('" & nKartuStok & "','" & Format_Tgl_Jam(Now) & "','" & grid.TextMatrix(i, 1) & "','" & dcNamaruangan.BoundText & "'," & _
                 "'" & qtyAwal & "','" & qtyMasuk & "','" & qtyKeluar & "','" & qtyAkhir & "','" & keterangan & "','001')"
    Next
    objSave2 = Right(objSave2, Len(objSave2) - 1)
    objSave3 = Right(objSave3, Len(objSave3) - 1)
    objSave4 = Right(objSave4, Len(objSave4) - 1)
    WriteRs "insert into transaksistruk (nStruk,noFaktur,tglStruk,tglFaktur,nRuangan,nKelTrans," & _
            "nRekanan,nUser,visible) " & _
            "values " & _
            objSave1
    WriteRs2 "insert into transaksistrukdetail (nStrukDetail,nStruk,tglkadaluarsa,nBarang,nSatuan,hargasatuan," & _
             "qty,hargadiskon,ppn,nUser,visible) " & _
             "values (" & _
             objSave2
    WriteRs3 "insert into transaksistok (nTStok,nBarang,nSatuan,harga,qtystok,diskon,ppn," & _
             "tglkadaluarsa,nStruk,nRuangan,nUser,visible) " & _
             "values (" & _
             objSave3
    WriteRs "insert into transaksikartustok (nKartuStok,tglStok,nBarang,nRuangan," & _
             "qtyAwal,qtyMasuk,qtyKeluar,qtyAkhir,keterangan,jenisTransaksi) " & _
             "values " & _
             objSave4
    cn.CommitTrans
    MsgBox "Tersimpan!", vbInformation, "..:."
    Unload Me
hell:
    cn.RollbackTrans
End Sub
Private Sub txtDiskon_Change()
    txtTotal.Text = Val(txtJml.Text) * (Val(txtTarif.Text) + Val(txtPpn.Text) - Val(txtDiskon.Text))
End Sub
Private Sub txtJml_Change()
    txtTotal.Text = Val(txtJml.Text) * (Val(txtTarif.Text) + Val(txtPpn.Text) - Val(txtDiskon.Text))
End Sub
Private Sub txtPpn_Change()
    txtTotal.Text = Val(txtJml.Text) * (Val(txtTarif.Text) + Val(txtPpn.Text) - Val(txtDiskon.Text))
End Sub
Public Sub clear()
    txtNRegistrasi.Text = ""
    txtRuangan.Text = ""
    txtNamaPasien.Text = ""
    txtJk.Text = ""
    txtUmur.Text = ""
    dtpFaktur.Value = Now()
End Sub
Private Sub txtTarif_Change()
    txtTotal.Text = Val(txtJml.Text) * (Val(txtTarif.Text) + Val(txtPpn.Text) - Val(txtDiskon.Text))
End Sub
Private Sub txtTarif_Click()
   Call txtTarif_Change
End Sub
Private Sub txtDiskon_Click()
   Call txtDiskon_Change
End Sub
Private Sub txtJml_Click()
   Call txtJml_Change
End Sub
Private Sub txtPpn_Click()
   Call txtPpn_Change
End Sub
'Private Sub txtNRm_Change()
'    Call clear
'    ReadRs "select * from pasien where nRm ='" & txtNrm.Text & "'"
'    If rs.RecordCount <> 0 Then
'        txtNamaPasien.Text = rs!namaPasien
'        txtJk.Text = IIf(rs!jenisKelamin = "L", "Laki-laki", "Perempuan")
'        txtUmur.Text = DateDiff("YYYY", CDate(rs!tglLahir), Now()) & " tahun " & (DateDiff("M", CDate(rs!tglLahir), Now()) Mod 12) & " bulan"
'    End If
'
'End Sub


Private Sub txtTotal_Change()
    txtTotal.Text = Val(txtJml.Text) * (Val(txtTarif.Text) + Val(txtPpn.Text) - Val(txtDiskon.Text))
End Sub
