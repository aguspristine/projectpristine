VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmKartuStok 
   Caption         =   "Kartu Stok"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16080
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
   ScaleWidth      =   16080
   WindowState     =   2  'Maximized
   Begin VB.Frame fr2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   7320
      Width           =   15135
   End
   Begin VB.Frame fr 
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   16095
      Begin VB.TextBox txtNamaBarang 
         Height          =   390
         Left            =   11040
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   480
         Width           =   3615
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
         TabIndex        =   3
         Top             =   360
         Width           =   1095
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
         FormatString    =   $"frmKartuStok.frx":0000
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
         Left            =   0
         TabIndex        =   7
         Top             =   480
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
         Format          =   117833731
         CurrentDate     =   31262
      End
      Begin MSComCtl2.DTPicker dt 
         Height          =   495
         Index           =   1
         Left            =   2640
         TabIndex        =   8
         Top             =   480
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
         Format          =   117833731
         CurrentDate     =   31262
      End
      Begin MSDataListLib.DataCombo dcDetailJenis 
         Height          =   390
         Left            =   7560
         TabIndex        =   11
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcRuangan 
         Height          =   390
         Left            =   4800
         TabIndex        =   12
         Top             =   480
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Periode"
         Height          =   375
         Index           =   5
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "s/d"
         Height          =   375
         Index           =   6
         Left            =   2160
         TabIndex        =   9
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Nama Barang"
         Height          =   495
         Index           =   4
         Left            =   11040
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Jenis Barang"
         Height          =   495
         Index           =   1
         Left            =   7560
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Ruangan"
         Height          =   495
         Index           =   0
         Left            =   4800
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmKartuStok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataCombo()
    Call loadDataCombo(dcRuangan, rs, "SELECT nRuangan,NamaRuangan FROM  ruangan where visible=1")
    Call loadDataCombo(dcDetailJenis, rs, "SELECT nDetailJenisBarang,namaDetailJenisBarang FROM  detailjenisbarang where visible=1")

End Sub

Private Sub btn_Click(Index As Integer)
    If txtNamaBarang.Text = "" Then MsgBox "Nama Barang Masih Kosong", vbCritical, ".:Warning": txtNmaBarang.SetFocus: Exit Sub
    If dcRuangan.Text = "" Then MsgBox "Nama Ruangan Masih Kosong", vbCritical, ".:Warning": dcRuangan.SetFocus: Exit Sub
    Call LoadData
End Sub

Public Sub LoadData()
    Dim str1, str2, str3, str4, str5, str6 As String
    
    If dcDetailJenis.Text <> "" Then
        str1 = "and brg.ndetailJenisBarang = '" & dcDetailJenis.BoundText & "'"
    End If
    
    ReadRs "select ks.tglStok,ks.keterangan,brg.namaBarang,ks.qtyAwal,ks.qtyMasuk,qtyKeluar,qtyAkhir " & _
            "From transaksikartustok as ks " & _
            "INNER JOIN barang as brg on brg.nBarang = ks.nBarang " & _
            "INNER JOIN ruangan as ru on ru.nRuangan = ks.nRuangan " & _
            "Where ks.tglStok between '" & Format_tgl(dt(0).Value) & " 00:00' and '" & Format_tgl(dt(1).Value) & " 23:59' " & _
            "and brg.namaBarang like '%" & txtNamaBarang.Text & "%' " & _
            str1
    Call isiGrid("frmKartuStok", grid, rs, "tglStok=2500,keterangan=6500,namaBarang=3500,qtyAwal=1500,qtyMasuk=1500,qtyKeluar=1500,qtyAkhir=1500")
End Sub

Private Sub Form_Activate()
    fr.Move 100, 100, Me.Width - 100, Me.Height - 1000 - fr2.Height
    fr2.Move fr.Left, fr.Top + fr.Height, fr.Width, fr2.Height
    grid.Move 100, grid.Top, fr.Width, fr.Height - grid.Top - 100
    
End Sub

Private Sub Form_Load()
Dim tglAwal, tglAkhir As String
'    For i = 0 To 3
'        txt(i).Text = ""
'    Next
    txtNamaBarang.Text = ""
    
    tglAwal = (GetSetting("MedApp", Me.Name, "tglAwal"))
    tglAkhir = (GetSetting("MedApp", Me.Name, "tglAkhir"))
    
    dt(0).Value = CDate(IIf(tglAwal = "", Now(), tglAwal))
    dt(1).Value = CDate(IIf(tglAkhir = "", Now(), tglAkhir))
    
    Call DataCombo
    Call LoadData
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call LoadData
End Sub
