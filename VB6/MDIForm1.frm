VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Total Project"
   ClientHeight    =   7590
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   12045
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   0
      ScaleHeight     =   6615
      ScaleWidth      =   2880
      TabIndex        =   4
      Top             =   975
      Width           =   2880
      Begin MSComctlLib.TreeView tv 
         Height          =   7455
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   13150
         _Version        =   393217
         Indentation     =   176
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   0
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
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   12045
      TabIndex        =   0
      Top             =   0
      Width           =   12045
      Begin VB.CommandButton Command3 
         Caption         =   "Desain Form"
         Height          =   855
         Left            =   6840
         TabIndex        =   3
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "frm1"
         Height          =   615
         Left            =   3720
         TabIndex        =   2
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "List Table"
         Height          =   615
         Left            =   600
         TabIndex        =   1
         Top             =   2280
         Width           =   3015
      End
      Begin MSDataListLib.DataCombo dcRuangan 
         Height          =   405
         Left            =   9000
         TabIndex        =   6
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "New Project"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -720
         TabIndex        =   8
         Top             =   120
         Width           =   3975
      End
      Begin VB.Label Label1 
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
         Left            =   7320
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
   End
   Begin VB.Menu mMaster 
      Caption         =   "Master"
      Visible         =   0   'False
      Begin VB.Menu mInstalasiRuangan 
         Caption         =   "Instalasi Ruangan"
      End
      Begin VB.Menu mKelompokAsalPasien 
         Caption         =   "Kelompok Asal Pasien"
      End
      Begin VB.Menu mDiagnosa 
         Caption         =   "Diagnosa"
      End
      Begin VB.Menu mBiling 
         Caption         =   "Billing"
      End
      Begin VB.Menu mPendukung 
         Caption         =   "Pendukung"
      End
      Begin VB.Menu mTindakan 
         Caption         =   "Tindakan"
         Begin VB.Menu mPelayananTindakan 
            Caption         =   "Pelayanan Tindakan"
         End
         Begin VB.Menu mTarifTindakan 
            Caption         =   "Tarif Tindakan"
         End
      End
      Begin VB.Menu mBarang 
         Caption         =   "Barang"
         Begin VB.Menu mJenisBarang 
            Caption         =   "Jenis Barang"
         End
         Begin VB.Menu mDataBarang 
            Caption         =   "Data Barang"
         End
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tvIndex As Integer
Private Sub Command1_Click()
    frmListTable.Show
    
End Sub

Private Sub Command2_Click()
    Form1.Show
End Sub

Private Sub Command3_Click()
    frmDesain.Show
End Sub

Private Sub mBiling_Click()
    frmBilling.Show
End Sub

Private Sub mDataBarang_Click()
    frmMBarang.Show
End Sub

Private Sub mDiagnosa_Click()
    frmDiagnosa.Show
End Sub

Private Sub MDIForm_Load()
    'frmDaftarPasien.Show
    'frmDaftarPasien.Form_Load
    Call loadMenu
    tv.Height = Screen.Height
    dcRuangan.Left = Screen.Width - dcRuangan.Width - 150
    Label1.Left = dcRuangan.Left - Label1.Width - 150
    Call loadDataCombo(dcRuangan, rs, "select distinct userruangan.nRuangan ,ruangan.namaRuangan from userruangan,ruangan where userruangan.nRuangan = ruangan.nRuangan and  visible =1")
    dcRuangan.Text = rs(1)
    dcRuangan.BoundText = rs(0)
End Sub

Private Sub mFile_Click()
    Picture2.visible = Not Picture2.visible
End Sub

Private Sub mInstalasiRuangan_Click()
    frmMRuanganPelayanan.Show
End Sub

Private Sub mJenisBarang_Click()
    frmMJenisBarang.Show
End Sub

Private Sub mKelompokAsalPasien_Click()
    frmMKelompokAsal.Show
End Sub

Private Sub mPelayananTindakan_Click()
    frmPelayananTindakan.Show
    
End Sub

Private Sub mPendukung_Click()
    frmMPendukungPasien.Show
End Sub

Private Sub mTarifTindakan_Click()
    frmTarifPelayanan.Show
End Sub


Private Sub loadMenu()
Dim txtKey0, txtKey1, txtKey2, txtKey3, txtKey4, txtKey5, txtKey6, txtKey7, txtNama, txtSatuan As String
Dim nodX As Node
    
    
    ReadRs "select * from userprep where namaUser='" & publicIdLogin & "'"
    tv.Nodes.clear
    ReadRs "select * from usermenu where nKelompokMenu in (" & rs!arrIdMenu & ") and visible=1 order by urutan,parent"
    'ReadRs "select * from usermenu where nPegawai = '" & publicNPegawai & "' and visible=1 order by urutan,parent"
    If rs.EOF = True Then Exit Sub
    Do
        txtNama = rs(2)
        Select Case rs(7)
            Case 0
                txtKey0 = "A~" & rs(0)
                Set nodX = tv.Nodes.Add(, , txtKey0, txtNama)
            Case 1
                txtKey1 = "B~" & rs(0)
                Set nodX = tv.Nodes.Add(txtKey0, tvwChild, txtKey1, txtNama)
                
            Case 2
                txtKey2 = "C~" & rs(0)
                Set nodX = tv.Nodes.Add(txtKey1, tvwChild, txtKey2, txtNama)
                
            Case 3
                txtKey3 = "D~" & rs(0)
                Set nodX = tv.Nodes.Add(txtKey2, tvwChild, txtKey3, txtNama)
                
            Case 4
                txtKey4 = "E~" & rs(0)
                Set nodX = tv.Nodes.Add(txtKey3, tvwChild, txtKey4, txtNama)
            
            Case 5
                txtKey5 = "F~" & rs(0)
                Set nodX = tv.Nodes.Add(txtKey4, tvwChild, txtKey5, txtNama)
            
            Case 6
                txtKey6 = "G~" & rs(0)
                Set nodX = tv.Nodes.Add(txtKey5, tvwChild, txtKey6, txtNama)
            
            Case 7
                txtKey7 = "H~" & rs(0)
                Set nodX = tv.Nodes.Add(txtKey6, tvwChild, txtKey7, txtNama)
                
        End Select
        
'        Set nodx = tv.Nodes.Add(, , "A~" & rs(0), rs(1))
        rs.MoveNext
    Loop Until rs.EOF
    txtKey0 = "A~9999"
    Set nodX = tv.Nodes.Add(, , txtKey0, "Log Out")
End Sub

Private Sub tv_Click()
On Error Resume Next
Dim idMenu As Integer
Dim splt() As String
Dim namaForm As Form
Dim nmForm As String
Dim loadFirst As String
    
    If tvIndex = 0 Then Exit Sub
    splt = Split(tv.Nodes(tvIndex).Key, "~")
    idMenu = Val(splt(1))
    If idMenu = 9999 Then Exit Sub 'Unload MDIForm1: frmLogin.Show: Exit Sub
    ReadRs "select * from usermenu where nUserMenu =" & idMenu & " and visible=1"
    If Trim(rs!formName) = "" Or IsNull(rs!formName) Then Exit Sub
    
    
    nmForm = IIf(IsNull(rs!formName), "", rs!formName) 'rs!formName
    loadFirst = IIf(IsNull(rs!loadFirst), "", rs!loadFirst)
    
    Set namaForm = Forms.Add(nmForm)
    namaForm.Show
    namaForm.ZOrder 0
   If loadFirst <> "" Then
        loadFirst = CallByName(namaForm, loadFirst, VbMethod)
    End If
    namaForm.WindowState = vbMaximized
    
End Sub

Private Sub tv_DblClick()
Dim idMenu As Integer
Dim splt() As String
Dim namaForm As Form
Dim nmForm As String
Dim loadFirst As String
    
    If tvIndex = 0 Then Exit Sub
    splt = Split(tv.Nodes(tvIndex).Key, "~")
    idMenu = Val(splt(1))
    If idMenu = 9999 Then Unload MDIForm1: frmLogin.Show: Exit Sub
    
End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
    tvIndex = Node.Index
End Sub
