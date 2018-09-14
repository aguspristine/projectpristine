VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMKelompokAsal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".: Kelompok Asal Pasien"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8820
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   8820
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
         TabCaption(0)   =   "Kelompok Pasien"
         TabPicture(0)   =   "frmMKelompokAsal.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmdBatal(0)"
         Tab(0).Control(1)=   "Frame4"
         Tab(0).Control(2)=   "cmdTutup(0)"
         Tab(0).Control(3)=   "cmdSimpan(0)"
         Tab(0).Control(4)=   "Frame2"
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Institusi Kelompok"
         TabPicture(1)   =   "frmMKelompokAsal.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdBatal(1)"
         Tab(1).Control(1)=   "cmdTutup(1)"
         Tab(1).Control(2)=   "cmdSimpan(1)"
         Tab(1).Control(3)=   "Frame7"
         Tab(1).Control(4)=   "Frame8"
         Tab(1).ControlCount=   5
         TabCaption(2)   =   "Asal Pasien"
         TabPicture(2)   =   "frmMKelompokAsal.frx":0038
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "cmdSimpan(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "cmdTutup(2)"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "cmdBatal(2)"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "Frame6"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "Frame5"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).ControlCount=   5
         Begin VB.Frame Frame8 
            Height          =   3255
            Left            =   -74880
            TabIndex        =   55
            Top             =   2040
            Width           =   8295
            Begin VSFlex8LCtl.VSFlexGrid gridInstitusiKelompok 
               Height          =   2895
               Left            =   120
               TabIndex        =   56
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
               FormatString    =   $"frmMKelompokAsal.frx":0054
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
         Begin VB.Frame Frame7 
            Height          =   1695
            Left            =   -74880
            TabIndex        =   42
            Top             =   360
            Width           =   8295
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
               TabIndex        =   45
               Top             =   1320
               Value           =   1  'Checked
               Width           =   855
            End
            Begin VB.TextBox txtNamaInstitusi 
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
               TabIndex        =   44
               Top             =   1200
               Width           =   3735
            End
            Begin VB.TextBox txtNoInstitusi 
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
               TabIndex        =   43
               Top             =   240
               Width           =   3015
            End
            Begin MSDataListLib.DataCombo dcKelompok 
               Height          =   330
               Left            =   2880
               TabIndex        =   54
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
               TabIndex        =   53
               Top             =   720
               Width           =   255
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Kelompok Pasien"
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
               Index           =   15
               Left            =   120
               TabIndex        =   52
               Top             =   720
               Width           =   1800
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
               Index           =   12
               Left            =   2640
               TabIndex        =   49
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
               Index           =   11
               Left            =   2640
               TabIndex        =   48
               Top             =   240
               Width           =   255
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Nama Institusi"
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
               Index           =   10
               Left            =   120
               TabIndex        =   47
               Top             =   1200
               Width           =   1530
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "No Institusi"
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
               TabIndex        =   46
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame5 
            Height          =   3255
            Left            =   120
            TabIndex        =   40
            Top             =   1920
            Width           =   8295
            Begin VSFlex8LCtl.VSFlexGrid gridAsal 
               Height          =   2895
               Left            =   120
               TabIndex        =   41
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
               FormatString    =   $"frmMKelompokAsal.frx":0133
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
         Begin VB.Frame Frame6 
            Height          =   1455
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   8295
            Begin VB.TextBox txtNAsal 
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
               TabIndex        =   35
               Top             =   360
               Width           =   3015
            End
            Begin VB.TextBox txtNmAsal 
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
               TabIndex        =   34
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
               TabIndex        =   33
               Top             =   960
               Value           =   1  'Checked
               Width           =   855
            End
            Begin VB.Label Label5 
               Caption         =   "No Asal"
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
               TabIndex        =   39
               Top             =   360
               Width           =   2775
            End
            Begin VB.Label Label2 
               Caption         =   "Nama Asal"
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
               Left            =   120
               TabIndex        =   38
               Top             =   840
               Width           =   2535
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
               TabIndex        =   37
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
               TabIndex        =   36
               Top             =   840
               Width           =   255
            End
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
            Index           =   2
            Left            =   4080
            TabIndex        =   31
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
            Left            =   7200
            TabIndex        =   30
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
            Index           =   2
            Left            =   5640
            TabIndex        =   29
            Top             =   5400
            Width           =   1455
         End
         Begin VB.Frame Frame2 
            Height          =   1455
            Left            =   -74880
            TabIndex        =   21
            Top             =   360
            Width           =   8295
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
               TabIndex        =   24
               Top             =   960
               Value           =   1  'Checked
               Width           =   855
            End
            Begin VB.TextBox txtNKelolompok 
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
            Begin VB.TextBox txtNmKelompok 
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
               TabIndex        =   22
               Top             =   840
               Width           =   3735
            End
            Begin VB.Label Label1 
               Caption         =   "No Kelompok"
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
               TabIndex        =   28
               Top             =   360
               Width           =   2415
            End
            Begin VB.Label Label2 
               Caption         =   "Nama Kelompok"
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
               TabIndex        =   27
               Top             =   840
               Width           =   2535
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
               TabIndex        =   26
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
               Index           =   2
               Left            =   2640
               TabIndex        =   25
               Top             =   840
               Width           =   255
            End
         End
         Begin VB.Frame Frame3 
            Height          =   1815
            Left            =   -74880
            TabIndex        =   10
            Top             =   360
            Width           =   8295
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
               TabIndex        =   14
               Top             =   1320
               Value           =   1  'Checked
               Width           =   855
            End
            Begin VB.TextBox txtNTindakanPelayanan 
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
               TabIndex        =   12
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox txtNmaTindakanPelayanan 
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
               TabIndex        =   11
               Top             =   1200
               Width           =   3855
            End
            Begin MSDataListLib.DataCombo dcJenisPelayanan 
               Height          =   330
               Left            =   3360
               TabIndex        =   13
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
               TabIndex        =   20
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
               TabIndex        =   19
               Top             =   240
               Width           =   255
            End
            Begin VB.Label Label2 
               Caption         =   "Nama Tindakan Pelayanan"
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
               TabIndex        =   18
               Top             =   1200
               Width           =   3015
            End
            Begin VB.Label Label3 
               Caption         =   "No Tindakan Pelayanan"
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
               TabIndex        =   17
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
               TabIndex        =   16
               Top             =   720
               Width           =   255
            End
            Begin VB.Label Label4 
               Caption         =   "Jenis Pelayanan"
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
               Top             =   720
               Width           =   2415
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
            Left            =   -69360
            TabIndex        =   9
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
            Index           =   0
            Left            =   -67800
            TabIndex        =   8
            Top             =   5400
            Width           =   1215
         End
         Begin VB.Frame Frame4 
            Height          =   3255
            Left            =   -74880
            TabIndex        =   6
            Top             =   1920
            Width           =   8295
            Begin VSFlex8LCtl.VSFlexGrid gridKelompok 
               Height          =   2895
               Left            =   120
               TabIndex        =   7
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
               FormatString    =   $"frmMKelompokAsal.frx":0212
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
            TabIndex        =   5
            Top             =   5400
            Width           =   1455
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
            Left            =   -69360
            TabIndex        =   4
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
            Left            =   -67800
            TabIndex        =   3
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
            Index           =   1
            Left            =   -70920
            TabIndex        =   2
            Top             =   5400
            Width           =   1455
         End
      End
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
      Left            =   2520
      TabIndex        =   51
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nama Institusi"
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
      Index           =   13
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   1530
   End
End
Attribute VB_Name = "frmMKelompokAsal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LoadData()
    
   Select Case SSTab1.Tab
        Case 0 'KelompokPasien
           ReadRs "SELECT  nKelompok AS NoKelompok,NamaKelompok,visible as status FROM  Kelompok"
           
           Call isiGrid("frmMKelompokAsal", gridKelompok, rs, "NoKelompok=1000,NamaKelompok=2000,status=500")
        
        Case 1 'institusi kelompok
           ReadRs "SELECT institusikelompok.nInstitusKelompok AS NoInstitusi ,kelompok.NamaKelompok," & _
           "  institusikelompok.namaInstitusikelompok, institusikelompok.visible as status FROM institusikelompok " & _
           " INNER JOIN kelompok ON institusikelompok.nKelompok=kelompok.nkelompok"
           
           Call isiGrid("frmMKelompokAsal", gridInstitusiKelompok, rs, "NoInstitusi=1000,NamaKelompok=2000,namaInstitusikelompok=2000,Status=500")
           
        Case 2 'asal pasien
           ReadRs "SELECT  nAsal AS NoAsal,NamaAsal,visible as status FROM  Asal"
           
           Call isiGrid("frmMKelompokAsal", gridAsal, rs, "NoAsal=1000,NamaAsal=2000,status=500")
   End Select
End Sub


Private Sub loadBersih()
    Select Case SSTab1.Tab
        Case 0 'KelompokPasien
            txtNKelolompok.Text = ""
            txtNmKelompok.Text = ""
        Case 1 'institusi kelompok
            txtNoInstitusi.Text = ""
            txtNamaInstitusi.Text = ""
            dcKelompok.Text = ""
        Case 2 'asal pasien
            txtNAsal.Text = ""
            txtNmAsal.Text = ""
            
    End Select

End Sub

Private Sub cmdBatal_Click(Index As Integer)
    Select Case Index
        Case 0 'KelompokPasien
            Call loadBersih
        Case 1 'institusi kelompok
            Call loadBersih
        Case 2 'asal pasien
            Call loadBersih
    End Select

End Sub

Private Sub cmdSimpan_Click(Index As Integer)
    Dim status As String
    
    Select Case Index
        Case 0 'KelompokPasien
             If txtNmKelompok.Text = "" Then MsgBox "Nama Kelompok Masih Kosong", vbCritical, ".:Warning": txtNmKelompok.SetFocus: Exit Sub
             txtNKelolompok = Format(getNewNumber("kelompok", "nKelompok", ""), "0#")
             status = IIf(chk(0).Value = Checked, "1", "0")
                          
             WriteRs "insert into kelompok values ('" & txtNKelolompok.Text & "', '" & txtNmKelompok.Text & "', '" & status & "')"
             MsgBox "Simpan berhasil !", vbOKOnly, ".:Informasi"
            
            Call LoadData
            Call loadBersih
        Case 1 'institusi kelompok
            
            If dcKelompok.Text = "" Then MsgBox "Pilih Kelompok Pasien", vbCritical, ".:Warning": dcKelompok.SetFocus: Exit Sub
            If txtNamaInstitusi.Text = "" Then MsgBox "institusi Kelompok Pasien Masih Kosong", vbCritical, ".:Warning": txtNamaInstitusi.SetFocus: Exit Sub
            txtNoInstitusi.Text = Format(getNewNumber("institusiKelompok", "nInstitusKelompok", ""), "0##")
            status = IIf(chk(1).Value = Checked, "1", "0")
            
            WriteRs "insert into institusiKelompok values ('" & txtNoInstitusi.Text & "', '" & txtNamaInstitusi.Text & "','" & dcKelompok.BoundText & "', '" & status & "')"
            MsgBox "Simpan berhasil !", vbOKOnly, ".:Informasi"
            
            Call LoadData
            Call loadBersih
            
        Case 2 'asal pasien
            If txtNmAsal.Text = "" Then MsgBox "Nama Asal Masih Kosong", vbCritical, ".:Warning": txtNmAsal.SetFocus: Exit Sub
            txtNAsal = Format(getNewNumber("asal", "nAsal", ""), "0#")
            status = IIf(chk(2).Value = Checked, "1", "0")
                          
            WriteRs "insert into asal values ('" & txtNAsal.Text & "', '" & txtNmAsal.Text & "', '" & status & "')"
            MsgBox "Simpan berhasil !", vbOKOnly, ".:Informasi"
            
            Call LoadData
            Call loadBersih
    End Select
End Sub
Private Sub DataCombo()
    Call loadDataCombo(dcKelompok, rs, "SELECT nKelompok,NamaKelompok FROM  Kelompok where visible=1")

End Sub

Private Sub cmdTutup_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            Unload Me
        Case 2
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIForm1)
    SSTab1.Tab = 0
    Call LoadData
    Call DataCombo
    Call loadBersih
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case PreviousTab
        Case 0 'KelompokPasien
            Call LoadData
            Call loadBersih
        Case 1 'institusi KelompokPasien
           Call LoadData
           Call DataCombo
           Call loadBersih
        Case 2 'asal Pasien
           Call LoadData
           Call loadBersih
    End Select
End Sub



