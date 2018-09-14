VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "..:."
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4635
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnKoneksi 
      Caption         =   ".."
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton vbButton2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton vbButton1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1320
      TabIndex        =   7
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ID               :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim pwd  As String

Private Sub btnKoneksi_Click()
    frmSettingKoneksi.Show (vbModal)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    Text1.Text = Format(Trim(Text1), "0####")
    ReadRs "select u.nPegawai, u.namaUser,u.kataSandi,p.namaPegawai,p.nPegawai from userPrep u,pegawai p " & _
           "where p.nPegawai=u.nPegawai and u.visible='1' and u.namaUser ='" & Text1.Text & "'"
    If rs.RecordCount <> 0 Then
      Text2.Text = rs!namaPegawai
      Text2.Tag = rs!nPegawai
      pwd = rs!kataSandi
      Text3.SetFocus
    Else
      MsgBox "User ini tidak aktif !", , "..:."
    End If
  End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then vbButton1_Click
End Sub

Private Sub vbButton1_Click()
  If Text1 = "" Then Text1.SetFocus: Exit Sub
  If Text2 = "" Then Text2.SetFocus: Exit Sub
  If Text3 = "" Then Text3.SetFocus: Exit Sub
  If Text3 = pwd Then
    publicNPegawai = Text2.Tag
    publicIdLogin = Text1.Text
    publicNamaLogin = Text2.Text
'    MDIForm1.Show
'    MDIForm1.Text1.Text = NAMA_LOGIN
    MDIForm1.Show
    Unload Me
  Else
    MsgBox "Password salah !", , "..:."
    Text3 = ""
    Text3.SetFocus
  End If
End Sub

Private Sub vbButton2_Click()
  End
End Sub

