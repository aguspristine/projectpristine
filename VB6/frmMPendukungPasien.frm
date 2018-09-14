VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMPendukungPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".: Master Pendukung"
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
         Tabs            =   6
         Tab             =   2
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
         TabCaption(0)   =   "Agama"
         TabPicture(0)   =   "frmMPendukungPasien.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmdBatal(0)"
         Tab(0).Control(1)=   "Frame4"
         Tab(0).Control(2)=   "cmdTutup(0)"
         Tab(0).Control(3)=   "cmdSimpan(0)"
         Tab(0).Control(4)=   "Frame2"
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Pendidikan"
         TabPicture(1)   =   "frmMPendukungPasien.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdBatal(1)"
         Tab(1).Control(1)=   "cmdTutup(1)"
         Tab(1).Control(2)=   "cmdSimpan(1)"
         Tab(1).Control(3)=   "Frame5"
         Tab(1).Control(4)=   "Frame3"
         Tab(1).ControlCount=   5
         TabCaption(2)   =   "Pekerjaan"
         TabPicture(2)   =   "frmMPendukungPasien.frx":0038
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "cmdBatal(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "cmdTutup(2)"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "cmdSimpan(2)"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "Frame6"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "Frame7"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).ControlCount=   5
         TabCaption(3)   =   "Hubungan Keluarga"
         TabPicture(3)   =   "frmMPendukungPasien.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame9"
         Tab(3).Control(1)=   "Frame8"
         Tab(3).Control(2)=   "cmdSimpan(3)"
         Tab(3).Control(3)=   "cmdTutup(3)"
         Tab(3).Control(4)=   "cmdBatal(3)"
         Tab(3).ControlCount=   5
         TabCaption(4)   =   "Gol. Darah"
         TabPicture(4)   =   "frmMPendukungPasien.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame11"
         Tab(4).Control(1)=   "Frame10"
         Tab(4).Control(2)=   "cmdSimpan(4)"
         Tab(4).Control(3)=   "cmdTutup(4)"
         Tab(4).Control(4)=   "cmdBatal(4)"
         Tab(4).ControlCount=   5
         TabCaption(5)   =   "Warga Negara"
         TabPicture(5)   =   "frmMPendukungPasien.frx":008C
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Frame13"
         Tab(5).Control(1)=   "Frame12"
         Tab(5).Control(2)=   "cmdSimpan(5)"
         Tab(5).Control(3)=   "cmdTutup(5)"
         Tab(5).Control(4)=   "cmdBatal(5)"
         Tab(5).ControlCount=   5
         Begin VB.Frame Frame13 
            Height          =   1455
            Left            =   -74880
            TabIndex        =   72
            Top             =   720
            Width           =   8055
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
               Index           =   5
               Left            =   6840
               TabIndex        =   75
               Top             =   960
               Value           =   1  'Checked
               Width           =   855
            End
            Begin VB.TextBox txtNmWrgNegara 
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
               TabIndex        =   74
               Top             =   840
               Width           =   3735
            End
            Begin VB.TextBox txtNWrgNegara 
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
               TabIndex        =   73
               Top             =   360
               Width           =   3015
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "No Warga Negara"
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
               TabIndex        =   79
               Top             =   360
               Width           =   1875
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Nama Warga Negara"
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
               Index           =   17
               Left            =   120
               TabIndex        =   78
               Top             =   840
               Width           =   2190
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
               Index           =   16
               Left            =   2640
               TabIndex        =   77
               Top             =   360
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
               Index           =   15
               Left            =   2640
               TabIndex        =   76
               Top             =   840
               Width           =   255
            End
         End
         Begin VB.Frame Frame12 
            Height          =   3015
            Left            =   -74880
            TabIndex        =   70
            Top             =   2280
            Width           =   8055
            Begin VSFlex8LCtl.VSFlexGrid gridWrgNegara 
               Height          =   2655
               Left            =   120
               TabIndex        =   71
               Top             =   240
               Width           =   7815
               _cx             =   13785
               _cy             =   4683
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
               FormatString    =   $"frmMPendukungPasien.frx":00A8
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
            Index           =   5
            Left            =   -69600
            TabIndex        =   69
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
            Index           =   5
            Left            =   -68040
            TabIndex        =   68
            Top             =   5400
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
            Index           =   5
            Left            =   -71160
            TabIndex        =   67
            Top             =   5400
            Width           =   1455
         End
         Begin VB.Frame Frame11 
            Height          =   1455
            Left            =   -74880
            TabIndex        =   59
            Top             =   720
            Width           =   8055
            Begin VB.TextBox txtNGolDarah 
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
               TabIndex        =   62
               Top             =   360
               Width           =   3015
            End
            Begin VB.TextBox txtNmGolDarah 
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
               TabIndex        =   61
               Top             =   840
               Width           =   3735
            End
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
               Index           =   4
               Left            =   6840
               TabIndex        =   60
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
               Index           =   14
               Left            =   2640
               TabIndex        =   66
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
               Index           =   13
               Left            =   2640
               TabIndex        =   65
               Top             =   360
               Width           =   255
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Nama Gol. Darah"
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
               Index           =   12
               Left            =   120
               TabIndex        =   64
               Top             =   840
               Width           =   1830
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "No Gol. Darah"
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
               TabIndex        =   63
               Top             =   360
               Width           =   1515
            End
         End
         Begin VB.Frame Frame10 
            Height          =   3015
            Left            =   -74880
            TabIndex        =   57
            Top             =   2280
            Width           =   8055
            Begin VSFlex8LCtl.VSFlexGrid gridGolDarah 
               Height          =   2655
               Left            =   120
               TabIndex        =   58
               Top             =   240
               Width           =   7815
               _cx             =   13785
               _cy             =   4683
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
               FormatString    =   $"frmMPendukungPasien.frx":0187
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
            Index           =   4
            Left            =   -69600
            TabIndex        =   56
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
            Index           =   4
            Left            =   -68040
            TabIndex        =   55
            Top             =   5400
            Width           =   1215
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
            Index           =   4
            Left            =   -71160
            TabIndex        =   54
            Top             =   5400
            Width           =   1455
         End
         Begin VB.Frame Frame9 
            Height          =   1455
            Left            =   -74880
            TabIndex        =   46
            Top             =   720
            Width           =   8055
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
               Index           =   3
               Left            =   6840
               TabIndex        =   49
               Top             =   960
               Value           =   1  'Checked
               Width           =   855
            End
            Begin VB.TextBox txtNmHubKeluarga 
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
               TabIndex        =   48
               Top             =   840
               Width           =   3735
            End
            Begin VB.TextBox txtNHubKeluarga 
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
               TabIndex        =   47
               Top             =   360
               Width           =   3015
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "No Hub. Keluarga"
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
               TabIndex        =   53
               Top             =   360
               Width           =   1875
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Nama Hub. Keluarga"
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
               Index           =   11
               Left            =   120
               TabIndex        =   52
               Top             =   840
               Width           =   2190
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
               Index           =   10
               Left            =   2640
               TabIndex        =   51
               Top             =   360
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
               Index           =   9
               Left            =   2640
               TabIndex        =   50
               Top             =   840
               Width           =   255
            End
         End
         Begin VB.Frame Frame8 
            Height          =   3015
            Left            =   -74880
            TabIndex        =   44
            Top             =   2280
            Width           =   8055
            Begin VSFlex8LCtl.VSFlexGrid gridHubKeluarga 
               Height          =   2655
               Left            =   120
               TabIndex        =   45
               Top             =   240
               Width           =   7815
               _cx             =   13785
               _cy             =   4683
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
               FormatString    =   $"frmMPendukungPasien.frx":0266
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
            Index           =   3
            Left            =   -69600
            TabIndex        =   43
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
            Index           =   3
            Left            =   -68040
            TabIndex        =   42
            Top             =   5400
            Width           =   1215
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
            Index           =   3
            Left            =   -71160
            TabIndex        =   41
            Top             =   5400
            Width           =   1455
         End
         Begin VB.Frame Frame7 
            Height          =   1455
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   8055
            Begin VB.TextBox txtNPekerjaan 
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
               TabIndex        =   36
               Top             =   360
               Width           =   3015
            End
            Begin VB.TextBox txtNmPekerjaan 
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
               TabIndex        =   35
               Top             =   840
               Width           =   3735
            End
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
               Index           =   2
               Left            =   6840
               TabIndex        =   34
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
               Index           =   8
               Left            =   2640
               TabIndex        =   40
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
               Index           =   7
               Left            =   2640
               TabIndex        =   39
               Top             =   360
               Width           =   255
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Nama Pekerjaan"
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
               Index           =   6
               Left            =   120
               TabIndex        =   38
               Top             =   840
               Width           =   1725
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "No Pekerjaan"
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
               TabIndex        =   37
               Top             =   360
               Width           =   1410
            End
         End
         Begin VB.Frame Frame6 
            Height          =   3015
            Left            =   120
            TabIndex        =   31
            Top             =   2280
            Width           =   8055
            Begin VSFlex8LCtl.VSFlexGrid gridPekerjaan 
               Height          =   2655
               Left            =   120
               TabIndex        =   32
               Top             =   240
               Width           =   7815
               _cx             =   13785
               _cy             =   4683
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
               FormatString    =   $"frmMPendukungPasien.frx":0345
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
            Index           =   2
            Left            =   5400
            TabIndex        =   30
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
            Index           =   2
            Left            =   6960
            TabIndex        =   29
            Top             =   5400
            Width           =   1215
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
            Index           =   2
            Left            =   3840
            TabIndex        =   28
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
            Index           =   1
            Left            =   -71160
            TabIndex        =   22
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
            Left            =   -68040
            TabIndex        =   21
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
            Left            =   -69600
            TabIndex        =   20
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
            Left            =   -71160
            TabIndex        =   19
            Top             =   5400
            Width           =   1455
         End
         Begin VB.Frame Frame5 
            Height          =   3015
            Left            =   -74880
            TabIndex        =   17
            Top             =   2280
            Width           =   8055
            Begin VSFlex8LCtl.VSFlexGrid gridPendidikan 
               Height          =   2655
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   7815
               _cx             =   13785
               _cy             =   4683
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
               FormatString    =   $"frmMPendukungPasien.frx":0424
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
            Height          =   3015
            Left            =   -74880
            TabIndex        =   15
            Top             =   2280
            Width           =   8055
            Begin VSFlex8LCtl.VSFlexGrid gridAgama 
               Height          =   2535
               Left            =   120
               TabIndex        =   16
               Top             =   240
               Width           =   7815
               _cx             =   13785
               _cy             =   4471
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
               FormatString    =   $"frmMPendukungPasien.frx":0503
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
            Left            =   -68040
            TabIndex        =   14
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
            Left            =   -69600
            TabIndex        =   13
            Top             =   5400
            Width           =   1455
         End
         Begin VB.Frame Frame3 
            Height          =   1455
            Left            =   -74880
            TabIndex        =   8
            Top             =   720
            Width           =   8055
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
               Index           =   1
               Left            =   6840
               TabIndex        =   27
               Top             =   960
               Value           =   1  'Checked
               Width           =   855
            End
            Begin VB.TextBox txtNmPendidikan 
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
               TabIndex        =   26
               Top             =   840
               Width           =   3735
            End
            Begin VB.TextBox txtNPendidikan 
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
               TabIndex        =   25
               Top             =   360
               Width           =   3015
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "No Pendidikan"
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
               TabIndex        =   12
               Top             =   360
               Width           =   1530
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Nama Pendidikan"
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
               TabIndex        =   11
               Top             =   840
               Width           =   1845
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
               Left            =   2640
               TabIndex        =   10
               Top             =   360
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
               Left            =   2640
               TabIndex        =   9
               Top             =   840
               Width           =   255
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1455
            Left            =   -74880
            TabIndex        =   2
            Top             =   720
            Width           =   8055
            Begin VB.TextBox txtNmAgama 
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
               TabIndex        =   24
               Top             =   840
               Width           =   3735
            End
            Begin VB.TextBox txtNAgama 
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
               TabIndex        =   23
               Top             =   360
               Width           =   3015
            End
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
               Index           =   0
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
               AutoSize        =   -1  'True
               Caption         =   "Nama Agama"
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
               Left            =   120
               TabIndex        =   4
               Top             =   840
               Width           =   1440
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "No Agama"
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
               TabIndex        =   3
               Top             =   360
               Width           =   1125
            End
         End
      End
   End
End
Attribute VB_Name = "frmMPendukungPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LoadData()
    
   Select Case SSTab1.Tab
        Case 0 'agama
           ReadRs "SELECT  nAgama AS NoAgama,NamaAgama,visible as status FROM  Agama"
           Call isiGrid("frmMPendukungPasien", gridAgama, rs, "nAgama=1000,Nama Agama=2000,Status=500")
        
        Case 1 'Pendidikan
           ReadRs "SELECT  nPendidikan AS NoPendidikan,NamaPendidikan,visible as status FROM  Pendidikan"
           Call isiGrid("frmMPendukungPasien", gridPendidikan, rs, "NoPendidikan=1000,NamaPendidikan=2000,status=500")
   
        Case 2 'Pekerjaan
           ReadRs "SELECT  nPekerjaan AS NoPekerjaan,NamaPekerjaan,visible as status FROM  Pekerjaan"
           Call isiGrid("frmMPendukungPasien", gridPekerjaan, rs, "NoPekerjaan=1000,NamaPekerjaan=2000,status=500")
    
        Case 3 'hub.Keluarga
           ReadRs "SELECT  nHubunganKeluarga AS NoHubunganKeluarga,NamaHubunganKeluarga,visible as status FROM  HubunganKeluarga"
           Call isiGrid("frmMPendukungPasien", gridHubKeluarga, rs, "NoHubunganKeluarga=1000,NamaHubunganKeluarga=2000,status=500")
   
        Case 4 'Gol darah
           ReadRs "SELECT  nGolonganDarah AS NoGolonganDarah,NamaGolonganDarah,visible as status FROM  GolonganDarah"
           Call isiGrid("frmMPendukungPasien", gridGolDarah, rs, "NoGolonganDarah=1000,NamaGolonganDarah=2000,status=500")
   
        Case 5 'Wrg Negara
           ReadRs "SELECT  nWargaNegara AS NoWargaNegara,NamaWargaNegara,visible as status FROM  WargaNegara"
           Call isiGrid("frmMPendukungPasien", gridWrgNegara, rs, "NoWargaNegara=1000,NamaWargaNegara=2000,status=500")
   
   End Select
End Sub

Private Sub loadBersih()
    Select Case SSTab1.Tab
        Case 0 'Agama
            txtNAgama.Text = ""
            txtNmAgama.Text = ""
        
        Case 1 'Pendidikan
            txtNPendidikan.Text = ""
            txtNmPendidikan.Text = ""
        
        Case 2 'Pekerjaan
            txtNPekerjaan.Text = ""
            txtNmPekerjaan.Text = ""
        
        Case 3 'hub.Keluarga
            txtNHubKeluarga.Text = ""
            txtNmHubKeluarga.Text = ""
        
        Case 4 'Gol darah
            txtNGolDarah.Text = ""
            txtNmGolDarah.Text = ""
        
        Case 5 'Wrg Negara
            txtNWrgNegara.Text = ""
            txtNmWrgNegara.Text = ""
    
    End Select

End Sub

Private Sub cmdBatal_Click(Index As Integer)
    Select Case Index
        Case 0
            Call loadBersih
        Case 1
            Call loadBersih
        Case 2 'Pekerjaan
            Call loadBersih
        Case 3 'hub.Keluarga
            Call loadBersih
        Case 4 'Gol darah
            Call loadBersih
        Case 5 'Wrg Negara
            Call loadBersih
  End Select

End Sub

Private Sub cmdSimpan_Click(Index As Integer)
    Dim status As String
   
    Select Case Index
        Case 0 'agama
             If txtNmAgama.Text = "" Then MsgBox "Agama Masih Kosong", vbCritical, ".:Warning": txtNmAgama.SetFocus: Exit Sub
             txtNAgama = Format(getNewNumber("agama", "nagama", ""), "0#")
             status = IIf(chk(0).Value = Checked, "1", "0")
                          
             WriteRs "insert into agama values ('" & txtNAgama.Text & "', '" & txtNmAgama.Text & "', '" & status & "')"
             MsgBox "Simpan berhasil !", vbOKOnly, ".:Informasi"
            
            Call LoadData
            Call loadBersih
        Case 1 'Pendidikan
            If txtNmPendidikan.Text = "" Then MsgBox "Pendidikan Masih Kosong", vbCritical, ".:Warning": txtNmPendidikan.SetFocus: Exit Sub
            txtNPendidikan = Format(getNewNumber("pendidikan", "npendidikan", ""), "0#")
            status = IIf(chk(1).Value = Checked, "1", "0")
            
            WriteRs "insert into Pendidikan values ('" & txtNPendidikan.Text & "', '" & txtNmPendidikan.Text & "','" & status & "')"
            MsgBox "Simpan berhasil !", vbOKOnly, ".:Informasi"
            
            Call LoadData
            Call loadBersih
        
        Case 2 'Pekerjaan
            If txtNmPekerjaan.Text = "" Then MsgBox "Pekerjaan Masih Kosong", vbCritical, ".:Warning": txtNmPekerjaan.SetFocus: Exit Sub
            txtNPekerjaan = Format(getNewNumber("Pekerjaan", "nPekerjaan", ""), "0#")
            status = IIf(chk(2).Value = Checked, "1", "0")
            
            WriteRs "insert into Pekerjaan values ('" & txtNPekerjaan.Text & "', '" & txtNmPekerjaan.Text & "','" & status & "')"
            MsgBox "Simpan berhasil !", vbOKOnly, ".:Informasi"
            
            Call LoadData
            Call loadBersih
        
        Case 3 'hub.Keluarga
            If txtNmHubKeluarga.Text = "" Then MsgBox "Hubungan Keluarga Masih Kosong", vbCritical, ".:Warning": txtNmHubKeluarga.SetFocus: Exit Sub
            txtNHubKeluarga = Format(getNewNumber("hubunganKeluarga", "nhubunganKeluarga", ""), "0#")
            status = IIf(chk(3).Value = Checked, "1", "0")
            
            WriteRs "insert into hubunganKeluarga values ('" & txtNHubKeluarga.Text & "', '" & txtNmHubKeluarga.Text & "','" & status & "')"
            MsgBox "Simpan berhasil !", vbOKOnly, ".:Informasi"
            
            Call LoadData
            Call loadBersih
        
        Case 4 'Gol darah
            If txtNmGolDarah.Text = "" Then MsgBox "Golongan Darah Masih Kosong", vbCritical, ".:Warning": txtNmGolDarah.SetFocus: Exit Sub
            txtNGolDarah = Format(getNewNumber("golonganDarah", "ngolonganDarah", ""), "0#")
            status = IIf(chk(4).Value = Checked, "1", "0")
            
            WriteRs "insert into golonganDarah values ('" & txtNGolDarah.Text & "', '" & txtNmGolDarah.Text & "','" & status & "')"
            MsgBox "Simpan berhasil !", vbOKOnly, ".:Informasi"
            
            Call LoadData
            Call loadBersih
        Case 5 'Wrg Negara
            If txtNmWrgNegara.Text = "" Then MsgBox "Warga Negara Masih Kosong", vbCritical, ".:Warning": txtNmWrgNegara.SetFocus: Exit Sub
            txtNWrgNegara = Format(getNewNumber("WargaNegara", "nWargaNegara", ""), "0#")
            status = IIf(chk(5).Value = Checked, "1", "0")
            
            WriteRs "insert into WargaNegara values ('" & txtNWrgNegara.Text & "', '" & txtNmWrgNegara.Text & "','" & status & "')"
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
        Case 2 'Pekerjaan
            Unload Me
        Case 3 'hub.Keluarga
            Unload Me
        Case 4 'Gol darah
            Unload Me
        Case 5 'Wrg Negara
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
        Case 2 'Pekerjaan
            
            Call LoadData
            Call loadBersih
        Case 3 'hub.Keluarga
            
            Call LoadData
            Call loadBersih
        Case 4 'Gol darah
            
            Call LoadData
            Call loadBersih
        Case 5 'Wrg Negara

            Call LoadData
            Call loadBersih
    End Select
End Sub

