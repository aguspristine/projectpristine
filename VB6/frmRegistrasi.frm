VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRegistrasi 
   Caption         =   "Registrasi"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10800
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
   ScaleHeight     =   7650
   ScaleWidth      =   10800
   Begin VB.CommandButton cmd 
      Caption         =   "Tutup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   7800
      TabIndex        =   24
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton cmd 
      Caption         =   "..."
      Height          =   375
      Index           =   3
      Left            =   3960
      TabIndex        =   23
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Batal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   5520
      TabIndex        =   22
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Cetak Bukti Registrasi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   960
      TabIndex        =   21
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   3240
      TabIndex        =   8
      Top             =   6600
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dt 
      Height          =   495
      Left            =   6480
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   0   'False
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
      Format          =   156237827
      CurrentDate     =   42784
   End
   Begin VB.TextBox txt 
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
      Index           =   2
      Left            =   2040
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txt 
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
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2160
      Width           =   4455
   End
   Begin VB.TextBox txt 
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
      Index           =   0
      Left            =   2400
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1680
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo dc 
      Height          =   405
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Top             =   3000
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
   Begin MSDataListLib.DataCombo dc 
      Height          =   405
      Index           =   1
      Left            =   2400
      TabIndex        =   3
      Top             =   3480
      Width           =   3855
      _ExtentX        =   6800
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
   Begin MSDataListLib.DataCombo dc 
      Height          =   405
      Index           =   2
      Left            =   2400
      TabIndex        =   4
      Top             =   3960
      Width           =   5295
      _ExtentX        =   9340
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
   Begin MSDataListLib.DataCombo dc 
      Height          =   405
      Index           =   3
      Left            =   2400
      TabIndex        =   5
      Top             =   4440
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
   Begin MSDataListLib.DataCombo dc 
      Height          =   405
      Index           =   4
      Left            =   2400
      TabIndex        =   6
      Top             =   4920
      Width           =   2655
      _ExtentX        =   4683
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
   Begin MSDataListLib.DataCombo dc 
      Height          =   405
      Index           =   5
      Left            =   2400
      TabIndex        =   7
      Top             =   5400
      Width           =   3735
      _ExtentX        =   6588
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
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Dokter : "
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
      Left            =   840
      TabIndex        =   18
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Asal Pasien : "
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
      Left            =   840
      TabIndex        =   17
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal Registrasi : "
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
      Left            =   3960
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "No Registrasi : "
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
      TabIndex        =   15
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Penjamin : "
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
      Left            =   840
      TabIndex        =   14
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Kelas Pelayanan : "
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
      TabIndex        =   13
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Ruangan : "
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
      Left            =   840
      TabIndex        =   12
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Instalasi : "
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
      Left            =   840
      TabIndex        =   11
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Nama Pasien : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   10
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "NRM : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   9
      Top             =   1680
      Width           =   1455
   End
End
Attribute VB_Name = "frmRegistrasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmd_Click(Index As Integer)
On Error GoTo hell
    Select Case Index
        Case 0 'simpan
        'nRegistrasi, tglRegistrasi, nRM, nInstalasi, nRuangan, nKelas, nStatusPasien, nStatusRawat,
        'nAsal , nStatusPulang, nKondisiPulang, tglPulang, nKelompok, nDokter, nUser
            If validasiSimpan = False Then Exit Sub
            cn.BeginTrans
            
            If cekSudahTerdaftarPasien(txt(0)) <> "Belum Pulang" Then
                Dim kode As String
'                WriteRs "INSERT INTO `db2`.`userprep` (`nPegawai`, `namaUser`, `kataSandi`, `arrIdMenu`, `nRuangan`, `nTingkat`, `visible`) VALUES ('P000000000', 'kasir3', 'qq', '3', '01', '01', '1');"

                kode = getNewNumberWithDate("registrasi", "nRegistrasi", "", Format_tgl(dt.Value))
                WriteRs "insert into registrasi values (" & _
                        "'" & kode & "','" & Format_Tgl_Jam(dt.Value) & "','" & txt(0).Text & "','" & dc(0).BoundText & "','" & dc(1).BoundText & "', " & _
                        "'" & dc(4).BoundText & "','-','-','" & dc(5).BoundText & "','00', " & _
                        "'-','" & Format_tgl(dt.Value) & "','" & dc(3).BoundText & "','" & dc(2).BoundText & "','" & publicIdLogin & "' " & _
                        ")"
                        
'                strSQL = "insert into registrasi "
'                strSQL = strSQL + "insert into registrasi "
                
                MsgBox "Registrasi berhasil !", vbOKOnly, "..:."
            Else
                MsgBox "Pasien " & cekSudahTerdaftarPasien(txt(0)) & "/Masih dirawat!", vbOKOnly, "..:."
            End If
            cn.CommitTrans
            
        Case 2
            Call Form_Load
            
        Case 3 'master pasien
            frmMasterPasien.Show
            frmMasterPasien.ZOrder 0
        Case 4
            Unload Me
    End Select
hell:
    cn.RollbackTrans
    
End Sub

Private Function validasiSimpan() As Boolean
    validasiSimpan = True
    For i = 0 To 1
        If txt(i).Text = "" Then
            validasiSimpan = False
            txt(i).SetFocus
            Exit Function
        End If
    Next
    For i = 0 To 5
        If dc(i).Text = "" Then
            validasiSimpan = False
            dc(i).SetFocus
            Exit Function
        End If
    Next
End Function

Private Sub dc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Dim arrCbo() As String
        Dim arrField() As String
        arrCbo = Split("instalasi:namaInstalasi~ruangan:namaRuangan~pegawai:namaPegawai~kelompok:namaKelompok~kelas:namaKelas~asal:namaAsal", "~")
        arrField() = Split(arrCbo(Index), ":")
        ReadRs "SELECT n" & arrField(0) & "," & arrField(1) & " FROM " & arrField(0) & " where visible='1' and (" & arrField(1) & " LIKE '%" & dc(Index).Text & "%')"
        If rs.EOF = True Then
            dc(Index).Text = ""
            Exit Sub
        End If
        dc(Index).BoundText = rs(0).Value
        dc(Index).Text = rs(1).Value
        
        If Index <> 5 Then
            Call cmd_Click(0)
        End If
    End If
End Sub

Public Sub Form_Load()
    For i = 0 To 2
        txt(i).Text = ""
    Next
    For i = 0 To 5
        dc(i).Text = ""
    Next
    dt.Value = Now()
    Call loadCombo
End Sub

Private Sub loadCombo()
    Dim arrCbo() As String
    Dim arrField() As String
    Dim ii As Integer
    
    arrCbo = Split("instalasi:namaInstalasi~ruangan:namaRuangan~pegawai:namaPegawai~kelompok:namaKelompok~kelas:namaKelas~asal:namaAsal", "~")
    For i = 0 To 5
        arrField() = Split(arrCbo(i), ":")
        ReadRs "select n" & arrField(0) & "," & arrField(1) & " from " & arrField(0) & " where visible='1'"
        
        Set dc(i).RowSource = rs
        dc(i).BoundColumn = rs(0).Name
        dc(i).ListField = rs(1).Name
    Next
End Sub

Private Sub txt_Change(Index As Integer)
    Select Case Index
        Case 0 'NRM
            ReadRs "Select * from pasien where nRM = '" & txt(0).Text & "'"
            If rs.RecordCount <> 0 Then
                txt(0).Text = rs!nRM
                txt(1).Text = rs!namaPasien
            End If
    End Select
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmd_Click (0)
End Sub
