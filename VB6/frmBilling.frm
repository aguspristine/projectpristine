VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmBilling 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".: Billing"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8820
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   8820
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin TabDlg.SSTab SSTab1 
         Height          =   6015
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   10610
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Cara Bayar"
         TabPicture(0)   =   "frmBilling.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmdBatal(0)"
         Tab(0).Control(1)=   "Frame4"
         Tab(0).Control(2)=   "cmdTutup(0)"
         Tab(0).Control(3)=   "cmdSimpan(0)"
         Tab(0).Control(4)=   "Frame2"
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Jenis Kartu"
         TabPicture(1)   =   "frmBilling.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame5"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "cmdSimpan(1)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "cmdTutup(1)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "cmdBatal(1)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).ControlCount=   5
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
            Index           =   1
            Left            =   4080
            TabIndex        =   23
            Top             =   5400
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
            Index           =   1
            Left            =   7200
            TabIndex        =   22
            Top             =   5400
            Width           =   1215
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
            Index           =   1
            Left            =   5640
            TabIndex        =   21
            Top             =   5400
            Width           =   1455
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
            Index           =   0
            Left            =   -70920
            TabIndex        =   20
            Top             =   5400
            Width           =   1455
         End
         Begin VB.Frame Frame5 
            Height          =   3255
            Left            =   120
            TabIndex        =   18
            Top             =   1920
            Width           =   8295
            Begin VSFlex8LCtl.VSFlexGrid gridJenisKartu 
               Height          =   2895
               Left            =   120
               TabIndex        =   19
               Top             =   240
               Width           =   8055
               _cx             =   14208
               _cy             =   5106
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
               FormatString    =   $"frmBilling.frx":0038
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
         Begin VB.Frame Frame4 
            Height          =   3255
            Left            =   -74880
            TabIndex        =   16
            Top             =   1920
            Width           =   8295
            Begin VSFlex8LCtl.VSFlexGrid gridCaraBayar 
               Height          =   2895
               Left            =   120
               TabIndex        =   17
               Top             =   240
               Width           =   8055
               _cx             =   14208
               _cy             =   5106
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
               FormatString    =   $"frmBilling.frx":0117
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
            Left            =   -67800
            TabIndex        =   15
            Top             =   5400
            Width           =   1215
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
            Left            =   -69360
            TabIndex        =   14
            Top             =   5400
            Width           =   1455
         End
         Begin VB.Frame Frame3 
            Height          =   1455
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   8295
            Begin VB.TextBox txtNmaJenisKartu 
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
               TabIndex        =   27
               Top             =   720
               Width           =   3855
            End
            Begin VB.TextBox txtNJenisKartu 
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
               TabIndex        =   26
               Top             =   240
               Width           =   2295
            End
            Begin VB.CheckBox chkTp 
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
               Left            =   7320
               TabIndex        =   9
               Top             =   840
               Value           =   1  'Checked
               Width           =   855
            End
            Begin VB.Label Label3 
               Caption         =   "No Jenis Kartu"
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
               Caption         =   "Nama Jenis Kartu"
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
               Index           =   5
               Left            =   120
               TabIndex        =   12
               Top             =   720
               Width           =   3015
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
               TabIndex        =   11
               Top             =   240
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
               Index           =   3
               Left            =   3120
               TabIndex        =   10
               Top             =   720
               Width           =   255
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1455
            Left            =   -74880
            TabIndex        =   2
            Top             =   360
            Width           =   8295
            Begin VB.TextBox txtNmCaraBayar 
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
               Left            =   2880
               TabIndex        =   25
               Top             =   840
               Width           =   3735
            End
            Begin VB.TextBox txtNCaraBayar 
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
               Left            =   2880
               TabIndex        =   24
               Top             =   360
               Width           =   3015
            End
            Begin VB.CheckBox chkJp 
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
               Left            =   6840
               TabIndex        =   7
               Top             =   960
               Value           =   1  'Checked
               Width           =   855
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
               Index           =   2
               Left            =   2640
               TabIndex        =   6
               Top             =   840
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
               Index           =   1
               Left            =   2640
               TabIndex        =   5
               Top             =   360
               Width           =   255
            End
            Begin VB.Label Label2 
               Caption         =   "Nama Cara Bayar"
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
               Left            =   120
               TabIndex        =   4
               Top             =   840
               Width           =   2535
            End
            Begin VB.Label Label1 
               Caption         =   "No Cara Bayar"
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
               Top             =   360
               Width           =   2775
            End
         End
      End
   End
End
Attribute VB_Name = "frmBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LoadData()
    
   Select Case SSTab1.Tab
        Case 0
           ReadRs "SELECT  nCaraBayar AS NoCaraBayar,NamaCaraBayar,visible as status FROM  CaraBayar"
           
           Call isiGrid("frmBilling", gridCaraBayar, rs, "NoCaraBayar=1000,Cara Bayar=2000,Status=5000")
        
        Case 1
           ReadRs "SELECT  nJenisKartu AS NoJenisKartu,NamaJenisKartu,visible as status FROM  JenisKartu"
           
           Call isiGrid("frmBilling", gridJenisKartu, rs, "NoJenisKartu=1000,NamaJenisKartu=2000,Status=1000")
   End Select
End Sub

Private Sub loadBersih()
    Select Case SSTab1.Tab
        Case 0
            txtNCaraBayar.Text = ""
            txtNmCaraBayar.Text = ""
        Case 1
            txtNJenisKartu.Text = ""
            txtNmaJenisKartu.Text = ""
            
    End Select

End Sub

Private Sub cmdBatal_Click(Index As Integer)
    Select Case Index
        Case 0
            Call loadBersih
        Case 1
            Call loadBersih
    End Select

End Sub

Private Sub cmdSimpan_Click(Index As Integer)
    Dim status As String
    
    Select Case Index
        Case 0 'tab1
             If txtNmCaraBayar.Text = "" Then MsgBox "Cara Bayar Masih Kosong", vbCritical, ".:Warning": txtNmCaraBayar.SetFocus: Exit Sub
             txtNCaraBayar = Format(getNewNumber("CaraBayar", "nCaraBayar", ""), "0#")
             status = IIf(chkJp.Value = Checked, "1", "0")
                          
             WriteRs "insert into caraBayar values ('" & txtNCaraBayar.Text & "', '" & txtNmCaraBayar.Text & "', '" & status & "')"
             MsgBox "Simpan berhasil !", vbOKOnly, ".:Informasi"
            
            Call LoadData
            Call loadBersih
        Case 1 'tab 2
            If txtNmaJenisKartu.Text = "" Then MsgBox "Jenis Kartu Masih Kosong", vbCritical, ".:Warning": txtNmaJenisKartu.SetFocus: Exit Sub
            txtNJenisKartu = Format(getNewNumber("JenisKartu", "nJenisKartu", ""), "0#")
            status = IIf(chkTp.Value = Checked, "1", "0")
            
            WriteRs "insert into jenisKartu values ('" & txtNJenisKartu.Text & "', '" & txtNmaJenisKartu.Text & "', '" & status & "')"
            MsgBox "Simpan berhasil !", vbOKOnly, ".:Informasi"
            
            Call LoadData
            Call loadBersih
    End Select
End Sub

Private Sub cmdTutup_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIForm1)
    SSTab1.Tab = 0
    Call LoadData
    Call loadBersih
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case PreviousTab
        Case 0
            Call LoadData
            Call loadBersih
        Case 1
           Call LoadData
           Call loadBersih
    End Select
End Sub

