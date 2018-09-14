VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmMBarang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".: Master Barang"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13425
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   13425
   Begin VB.Frame Frame1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13335
      Begin VB.Frame Frame3 
         Height          =   2295
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   13095
         Begin VB.CheckBox chk 
            Caption         =   "Aktif"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7560
            TabIndex        =   10
            Top             =   1800
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.TextBox txtNBarang 
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
            Left            =   3360
            TabIndex        =   8
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox txtNmaBarang 
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
            Left            =   3360
            TabIndex        =   7
            Top             =   1200
            Width           =   4935
         End
         Begin MSDataListLib.DataCombo dcDetailJenis 
            Height          =   330
            Left            =   3360
            TabIndex        =   9
            Top             =   720
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dcSatuan 
            Height          =   330
            Left            =   3360
            TabIndex        =   17
            Top             =   1680
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   ":"
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
            Index           =   0
            Left            =   3120
            TabIndex        =   19
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Satuan Barang"
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
            TabIndex        =   18
            Top             =   1680
            Width           =   1530
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   ":"
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
            Index           =   3
            Left            =   3120
            TabIndex        =   16
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   ":"
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
            Index           =   4
            Left            =   3120
            TabIndex        =   15
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label2 
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
            Index           =   5
            Left            =   120
            TabIndex        =   14
            Top             =   1200
            Width           =   1425
         End
         Begin VB.Label Label3 
            Caption         =   "No Barang"
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
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   ":"
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
            Index           =   6
            Left            =   3120
            TabIndex        =   12
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Detail Jenis Barang"
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
            TabIndex        =   11
            Top             =   720
            Width           =   2010
         End
      End
      Begin VB.Frame Frame5 
         Height          =   4935
         Left            =   120
         TabIndex        =   4
         Top             =   2520
         Width           =   13095
         Begin VSFlex8LCtl.VSFlexGrid gridBarang 
            Height          =   4575
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   12735
            _cx             =   22463
            _cy             =   8070
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
            FormatString    =   $"frmMBarang.frx":0000
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
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
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
         Index           =   0
         Left            =   10320
         TabIndex        =   3
         Top             =   7560
         Width           =   1455
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Tutup"
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
         Index           =   0
         Left            =   11880
         TabIndex        =   2
         Top             =   7560
         Width           =   1215
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Default         =   -1  'True
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
         Index           =   0
         Left            =   8760
         TabIndex        =   1
         Top             =   7560
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmMBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LoadData()
       
           ReadRs "SELECT b.nBarang AS NoBarang,b.namaBarang,djb.namaDetailJenisBarang," & _
                  " sb.namaSatuan , b.visible AS status " & _
                  " FROM barang AS b " & _
                  " INNER JOIN detailjenisbarang AS djb ON djb.nDetailJenisBarang = b.ndetailJenisBarang " & _
                  " INNER JOIN satuanBarang AS sb ON sb.nSatuan = b.nSatuan"
           
           Call isiGrid("frmMBarang", gridBarang, rs, "NoBarang=1500,namaBarang=3000,namaDetailJenisBarang=3000,namaSatuan=2000,Status=500")

End Sub

Private Sub DataCombo()
    Call loadDataCombo(dcDetailJenis, rs, "SELECT ndetailJenisBarang,NamadetailJenisBarang FROM  detailJenisBarang where visible=1")
    Call loadDataCombo(dcSatuan, rs, "SELECT nSatuan,NamaSatuan FROM  SatuanBarang where visible=1")

End Sub

Private Sub loadBersih()
    txtNBarang.Text = ""
    txtNmaBarang.Text = ""
    dcDetailJenis.Text = ""
    dcSatuan.Text = ""
End Sub

Private Sub cmdBatal_Click(Index As Integer)
    Select Case Index
        Case 0
            Call loadBersih
    End Select

End Sub

Private Sub cmdSimpan_Click(Index As Integer)
    Dim status As String
    
    Select Case Index
        Case 0 'tab 2
            If dcDetailJenis.Text = "" Then MsgBox "Pilih Detail Jenis Barang", vbCritical, ".:Warning": dcDetailJenis.SetFocus: Exit Sub
            If txtNmaBarang.Text = "" Then MsgBox "Nama Barang Masih Kosong", vbCritical, ".:Warning": txtNmaBarang.SetFocus: Exit Sub
            If dcSatuan.Text = "" Then MsgBox "Pilih Satuan Barang", vbCritical, ".:Warning": dcSatuan.SetFocus: Exit Sub
            txtNBarang = Format(getNewNumber("Barang", "nBarang", ""), "0#########")
            status = IIf(chk.Value = Checked, "1", "0")
            
            WriteRs "insert into barang values ('" & txtNBarang.Text & "','" & dcDetailJenis.BoundText & "', '" & txtNmaBarang.Text & "','" & dcSatuan.BoundText & "' , '" & status & "')"
            MsgBox "Simpan berhasil !", vbOKOnly, ".:Informasi"
            
            Call LoadData
            Call loadBersih
    End Select
End Sub

Private Sub cmdTutup_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIForm1)
    Call LoadData
    Call DataCombo
    Call loadBersih
End Sub
