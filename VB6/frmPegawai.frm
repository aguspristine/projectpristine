VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserPreparation 
   Caption         =   "User Preparation"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   13935
   Begin MSComctlLib.ListView lv 
      Height          =   2415
      Left            =   6840
      TabIndex        =   14
      Top             =   1440
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4260
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
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
      Left            =   11880
      TabIndex        =   13
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox txtUserName 
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
      Top             =   1560
      Width           =   3015
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
      Left            =   8520
      TabIndex        =   10
      Top             =   4560
      Width           =   1575
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
      Left            =   10200
      TabIndex        =   9
      Top             =   4560
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo dcPegawai 
      Height          =   405
      Left            =   2280
      TabIndex        =   0
      Top             =   1080
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
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
   Begin MSDataListLib.DataCombo dcTingkat 
      Height          =   405
      Left            =   2280
      TabIndex        =   1
      Top             =   2520
      Width           =   3135
      _ExtentX        =   5530
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
   Begin MSDataListLib.DataCombo dcRuangan 
      Height          =   405
      Left            =   9480
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
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
   Begin MSComctlLib.ListView lv2 
      Height          =   2415
      Left            =   10200
      TabIndex        =   16
      Top             =   1440
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4260
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label5 
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
      Height          =   375
      Left            =   10200
      TabIndex        =   17
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Kelompok Menu"
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
      TabIndex        =   15
      Top             =   1080
      Width           =   2295
   End
   Begin MSForms.CheckBox chkStatus 
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   3120
      Width           =   1695
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2990;873"
      Value           =   "0"
      Caption         =   "Status Aktif"
      FontName        =   "Tahoma"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Tingkat :"
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
      Left            =   0
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Ruangan :"
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
      Left            =   7200
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Password :"
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
      Left            =   0
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "User Name :"
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
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "NamaPegawai :"
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
      Left            =   -240
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
End
Attribute VB_Name = "frmUserPreparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub clear()
    'dcKodePegawai.Text = ""
    'dcPegawai.Text = ""
    txtUserName.Text = ""
    txtPassword.Text = ""
    dcRuangan.Text = ""
    dcTingkat.Text = ""
    chkStatus.Value = vbUnchecked
    
    For i = 0 To lv.ListItems.Count - 1
        lv.ListItems(i + 1).Checked = False
    Next
    For i = 0 To lv2.ListItems.Count - 1
        lv2.ListItems(i + 1).Checked = False
    Next
End Sub

Private Sub btnBatal_Click()
    Call clear
End Sub

Private Sub btnSimpan_Click()
Dim arrIdMenu As String
Dim vis As String
    
    If chkStatus.Value = True Then
        vis = "1"
    Else
        vis = "0"
    End If
    
    For i = 1 To lv.ListItems.Count
        If lv.ListItems(i).Checked = True Then
            arrIdMenu = arrIdMenu & "," & Replace(lv.ListItems(i).Key, "lv", "")
        End If
    Next
    arrIdMenu = Right(arrIdMenu, Len(arrIdMenu) - 1)
    ReadRs "select * from userprep where nPegawai = '" & dcPegawai.BoundText & "'"
    If rs.RecordCount = 0 Then
        WriteRs "insert into userprep (nPegawai,namaUser,kataSandi,arrIdMenu,nTingkat,visible)" & _
                "values ('" & dcPegawai.BoundText & "','" & txtUserName.Text & "','" & txtPassword.Text & "'," & _
                "'" & arrIdMenu & "','" & dcTingkat.BoundText & "','" & vis & "')"
    Else
        WriteRs "update userprep set namaUser='" & txtUserName.Text & "',kataSandi='" & txtPassword.Text & "', " & _
                "arrIdMenu='" & arrIdMenu & "',nTingkat='" & dcTingkat.BoundText & "',visible='" & vis & "' " & _
                "where nPegawai='" & dcPegawai.BoundText & "'"
    End If
    WriteRs "delete from userruangan where nPegawai ='" & dcPegawai.BoundText & "'"
    For i = 1 To lv2.ListItems.Count
        If lv2.ListItems(i).Checked = True Then
            WriteRs "insert into userruangan (nPegawai,nRuangan) values ('" & dcPegawai.BoundText & "','" & Replace(lv2.ListItems(i).Key, "lv2", "") & "')"
        End If
    Next
    MsgBox "Tersimpan!", vbInformation, "..:."
End Sub

Private Sub btnTutup_Click()
    Unload Me
End Sub

Private Sub dcPegawai_Change()
    Call clear
    ReadRs "select * from userprep where nPegawai = '" & dcPegawai.BoundText & "'"
    If rs.RecordCount <> 0 Then
        txtUserName.Text = rs!namaUser
        txtPassword.Text = rs!kataSandi
        
        ReadRs2 "select * from ruangan where nRuangan='" & rs!nRuangan & "'"
        If rs2.RecordCount <> 0 Then
            dcRuangan.BoundText = rs2!nRuangan
            dcRuangan.Text = rs2!namaRuangan
        End If
        ReadRs2 "select * from tingkatuser where nTingkatUser='" & rs!nTingkat & "'"
        If rs2.RecordCount <> 0 Then
            dcTingkat.BoundText = rs2!nTingkatUser
            dcTingkat.Text = rs2!namaTingkatUser
        End If
        If rs!visible = "1" Then
            chkStatus.Value = vbChecked
        Else
            chkStatus.Value = vbUnchecked
        End If
        ReadRs "select * from kelompokmenu where nKelompok in (" & rs!arrIdMenu & ")"
        Dim ii As Integer
        For i = 0 To rs.RecordCount - 1
            For ii = 1 To lv.ListItems.Count
                If rs!nKelompok = Replace(lv.ListItems(ii).Key, "lv", "") Then
                    lv.ListItems(ii).Checked = True
                    Exit For
                End If
            Next
            rs.MoveNext
        Next
        
        ReadRs "select * from userruangan where nPegawai ='" & dcPegawai.BoundText & "'"
        For i = 0 To rs.RecordCount - 1
            For ii = 1 To lv2.ListItems.Count
                If rs!nRuangan = Replace(lv2.ListItems(ii).Key, "lv2", "") Then
                    lv2.ListItems(ii).Checked = True
                    Exit For
                End If
            Next
            rs.MoveNext
        Next
    End If
End Sub

Private Sub dcPegawai_Click(Area As Integer)
    Call dcPegawai_Change
End Sub

Private Sub Form_Load()
   ' Call loadDataCombo(dcKodePegawai, rs, "select nPegawai,nPegawai from pegawai where visible=1")
    Call loadDataCombo(dcPegawai, rs, "select nPegawai,namaPegawai from pegawai where  visible=1")
    Call loadDataCombo(dcRuangan, rs, "select nRuangan,namaRuangan from ruangan where visible=1")
    Call loadDataCombo(dcTingkat, rs, "select nTingkatUser,namaTingkatUser from tingkatuser where visible=1")
    ReadRs "select * from kelompokmenu where visible=1"
    lv.ListItems.clear
    For i = 0 To rs.RecordCount - 1
        lv.ListItems.Add i + 1, "lv" & rs!nKelompok, rs!namaKelompokMenu
        rs.MoveNext
    Next
    ReadRs "select * from ruangan where visible=1"
    lv2.ListItems.clear
    For i = 0 To rs.RecordCount - 1
        lv2.ListItems.Add i + 1, "lv2" & rs!nRuangan, rs!namaRuangan
        rs.MoveNext
    Next
    
    Call clear
    
End Sub
