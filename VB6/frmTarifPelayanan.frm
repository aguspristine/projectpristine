VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmTarifPelayanan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".: Kelompok Asal Pasien"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10395
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   10395
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      Begin VB.TextBox txtCari 
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
         Index           =   1
         Left            =   1680
         TabIndex        =   22
         Top             =   5760
         Width           =   2895
      End
      Begin VB.TextBox txtCari 
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
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
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
         Left            =   6720
         TabIndex        =   18
         Top             =   5640
         Width           =   1095
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
         Left            =   7920
         TabIndex        =   17
         Top             =   5640
         Width           =   1095
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
         Left            =   9120
         TabIndex        =   16
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Frame Frame5 
         Height          =   3735
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   10095
         Begin VSFlex8LCtl.VSFlexGrid gridTarifTindakan 
            Height          =   3375
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   9855
            _cx             =   17383
            _cy             =   5953
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
            FormatString    =   $"frmTarifPelayanan.frx":0000
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
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   10095
         Begin VB.TextBox txtHargaTotal 
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
            Left            =   6720
            TabIndex        =   2
            Top             =   720
            Width           =   3255
         End
         Begin MSDataListLib.DataCombo dcTindakan 
            Height          =   330
            Left            =   6720
            TabIndex        =   3
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
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
         Begin MSDataListLib.DataCombo dcJnsTindakan 
            Height          =   330
            Left            =   2160
            TabIndex        =   10
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
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
         Begin MSDataListLib.DataCombo dcKelas 
            Height          =   330
            Left            =   2160
            TabIndex        =   11
            Top             =   720
            Width           =   2295
            _ExtentX        =   4048
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
            Index           =   1
            Left            =   6480
            TabIndex        =   13
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Total Tarif "
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
            Left            =   4680
            TabIndex        =   12
            Top             =   720
            Width           =   1185
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
            Left            =   1920
            TabIndex        =   9
            Top             =   720
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
            Left            =   1920
            TabIndex        =   8
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Tindakan"
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
            TabIndex        =   6
            Top             =   240
            Width           =   1575
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
            Left            =   6480
            TabIndex        =   5
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nama Tindakan"
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
            TabIndex        =   4
            Top             =   240
            Width           =   1665
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nama Tindakan"
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
         Left            =   1680
         TabIndex        =   21
         Top             =   5400
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   5400
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmTarifPelayanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LoadData()
    ReadRs " SELECT TP.NamaTindakan,K.namaKelas AS Kelas,HTT.totalHarga AS totalHarga FROM hargaTotalTindakan AS HTT" & _
            " INNER JOIN tindakanpelayanan AS TP ON TP.nTindakan = HTT.nTindakan" & _
            " INNER JOIN kelas AS K ON K.nkelas=HTT.nKelas where K.namaKelas like '%" & txtCari(0).Text & "%' and TP.NamaTindakan like '%" & txtCari(1).Text & "%'  "
    Call isiGrid("frmTarifPelayanan", gridTarifTindakan, rs, "NamaTindakan=5000,Kelas=1500,totalHarga=1500")
End Sub

Private Sub DataCombo()
    Call loadDataCombo(dcJnsTindakan, rs, "SELECT nJenisTindakan,NamaJenisTindakan FROM  jenistindakan where visible=1")
    Call loadDataCombo(dcKelas, rs, "SELECT nKelas,NamaKelas FROM kelas where visible=1")
    Call loadDataCombo(dcTindakan, rs, "SELECT nTindakan,NamaTindakan FROM  TindakanPelayanan where njenisTindakan ='" & dcJnsTindakan.BoundText & "' and visible=1")

End Sub

Private Sub loadBersih()
    dcTindakan.Text = ""
    dcJnsTindakan.Text = ""
    dcKelas.Text = ""
    txtHargaTotal.Text = 0
    txtCari(0).Text = ""
    txtCari(1).Text = ""
End Sub

Private Sub cmdSimpan_Click()
    Dim strQuery As String
    
    If dcJnsTindakan.Text = "" Then MsgBox "Pilih Jenis Pelayanan", vbCritical, ".:Warning": dcJnsTindakan.SetFocus: Exit Sub
    If dcTindakan.Text = "" Then MsgBox "Pilih Tindakan Pelayanan", vbCritical, ".:Warning": dcTindakan.SetFocus: Exit Sub
    If dcKelas.Text = "" Then MsgBox "Pilih Kelas Pelayanan", vbCritical, ".:Warning": dcKelas.SetFocus: Exit Sub
    If txtHargaTotal.Text = 0 Or txtHargaTotal.Text = "" Then MsgBox "Harga Total Tindakan Masih Kosong", vbCritical, ".:Warning": txtNmaTindakanPelayanan.SetFocus: Exit Sub
   
    strQuery = "select * from hargaTotalTindakan where nTindakan='" & dcTindakan.BoundText & "' and nKelas='" & dcKelas.BoundText & "'"
    Set rs2 = Nothing
    rs2.Open strQuery, cn, adOpenStatic, adLockOptimistic
    
    If rs2.RecordCount = 0 Then
        WriteRs "insert into hargaTotalTindakan values ('" & dcTindakan.BoundText & "', '" & dcKelas.BoundText & "', '" & txtHargaTotal.Text & "')"
        MsgBox "Simpan Tarif Pelayanan berhasil !", vbOKOnly, ".:Informasi"
    Else
       Dim pesan As VbMsgBoxResult
       pesan = MsgBox("Tindakan - " & dcTindakan.Text & " - dengan Kelas - " & dcKelas.Text & " - sudah ada tarif. " & vbNewLine & " Apakah akan merubah tarifnya?", vbQuestion + vbYesNo, "Konfirmasi")
       If pesan = vbYes Then
        WriteRs "update hargaTotalTindakan set totalHarga = '" & txtHargaTotal.Text & "' where ntindakan='" & dcTindakan.BoundText & "' and nkelas= '" & dcKelas.BoundText & "'"
        MsgBox "Ubah Tarif Pelayanan berhasil !", vbOKOnly, ".:Informasi"
       End If
    End If
    
    rs2.Close
     
    Call LoadData
    Call loadBersih

End Sub

Private Sub dcJnsTindakan_Click(Area As Integer)
    dcTindakan.Text = ""
    Call loadDataCombo(dcTindakan, rs, "SELECT nTindakan,NamaTindakan FROM  TindakanPelayanan where njenisTindakan ='" & dcJnsTindakan.BoundText & "' and visible=1")

End Sub

Private Sub cmdBatal_Click()
    Call loadBersih
    Call LoadData
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIForm1)
    Call LoadData
    Call DataCombo
    Call loadBersih
End Sub

Private Sub txtCari_Change(Index As Integer)
    Select Case Index
        Case 0
          Call LoadData
        Case 1
          Call LoadData
    End Select
End Sub
