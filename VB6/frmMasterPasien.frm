VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMasterPasien 
   Caption         =   "Master Pasien"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13185
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8055
   ScaleWidth      =   13185
   Begin VB.CommandButton cmd 
      Caption         =   "Registrasi"
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
      Left            =   6600
      TabIndex        =   11
      Top             =   6360
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Perempuan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2160
      TabIndex        =   10
      Top             =   3720
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Laki-Laki"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2160
      TabIndex        =   9
      Top             =   3360
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.TextBox txt 
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
      Index           =   0
      Left            =   2160
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1440
      Width           =   1455
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
      Left            =   2160
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1920
      Width           =   4455
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
      Left            =   4320
      TabIndex        =   1
      Top             =   6360
      Width           =   2175
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
      Index           =   1
      Left            =   2040
      TabIndex        =   0
      Top             =   6360
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dt 
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   2760
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
      Format          =   96272387
      CurrentDate     =   42784
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
      Left            =   360
      TabIndex        =   6
      Top             =   1920
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
      Left            =   600
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal Lahir : "
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
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Jenis Kelamin : "
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
      TabIndex        =   4
      Top             =   3360
      Width           =   1695
   End
End
Attribute VB_Name = "frmMasterPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0 'simpan
            Dim nRM As String
            Dim nIP As String
            Dim namaPasien As String
            Dim tempatLahir As String
            Dim tglLahir As String
            Dim jenisKelamin As String
            Dim golDarah As String
            Dim nAgama As String
            Dim nStatusNikah As String
            Dim alamat As String
            Dim telepon As String
            Dim nProvinsi As String
            Dim nKota As String
            Dim nKecamatan As String
            Dim nKelurahan As String
            Dim rw As String
            Dim rt As String
            Dim kodePos As String
            Dim visible As String
            
            If txt(1).Text = "" Then Exit Sub
            nRM = txt(0).Text
            namaPasien = txt(1).Text
            tglLahir = Format_tgl(dt.Value)
            If Option1(0).Value = True Then jenisKelamin = "L"
            If Option1(1).Value = True Then jenisKelamin = "P"
            visible = "1"
            
            WriteRs "insert into pasien values ('" & nRM & "', '" & nIP & "', '" & namaPasien & "', '" & tempatLahir & "', " & _
                "'" & tglLahir & "', '" & jenisKelamin & "', '" & golDarah & "', '" & nAgama & "', '" & nStatusNikah & "', " & _
                "'" & alamat & "', '" & telepon & "', '" & nProvinsi & "', '" & nKota & "', '" & nKecamatan & "', " & _
                "'" & nKelurahan & "', '" & rw & "', '" & rt & "', '" & kodePos & "', '" & visible & "')"
                
            MsgBox "Simpan berhasil !", vbOKOnly, "..:."
    Case 2
        If txt(1).Text = "" Then Exit Sub
        frmRegistrasi.Show
        frmRegistrasi.ZOrder 0
        frmRegistrasi.Form_Load
        frmRegistrasi.txt(0).Text = txt(0).Text
    End Select
End Sub

Public Sub Form_Load()
    For i = 0 To 1
        txt(i).Text = ""
    Next
    dt.Value = Now()
    txt(0) = Format(getNewNumber("pasien", "nRM", ""), "0#########")
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
End Function
