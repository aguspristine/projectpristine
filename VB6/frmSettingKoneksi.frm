VERSION 5.00
Begin VB.Form frmSettingKoneksi 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "..:."
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4620
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
   Icon            =   "frmSettingKoneksi.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
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
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox Text4 
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
      Height          =   330
      Left            =   1440
      TabIndex        =   11
      Top             =   2280
      Width           =   1815
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
      Height          =   330
      Left            =   1440
      TabIndex        =   10
      Top             =   1920
      Width           =   2415
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
      Height          =   330
      Left            =   1440
      TabIndex        =   9
      Top             =   1560
      Width           =   975
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
      TabIndex        =   4
      Top             =   3480
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
      TabIndex        =   3
      Top             =   3480
      Width           =   975
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
      Height          =   330
      Left            =   1440
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
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
      Index           =   4
      Left            =   360
      TabIndex        =   8
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User :"
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
      Index           =   3
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Database :"
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
      Index           =   2
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Setting Koneksi"
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
      TabIndex        =   5
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Port :"
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
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Server :"
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
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "frmSettingKoneksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim pwd  As String

Private Sub Form_Load()
    Text1 = GetSetting("T-PRO", "Koneksi", "Server")
  Text2 = GetSetting("T-PRO", "Koneksi", "Port")
  Text3 = GetSetting("T-PRO", "Koneksi", "Database")
  Text4 = GetSetting("T-PRO", "Koneksi", "User")
  Text5 = GetSetting("T-PRO", "Koneksi", "Password")
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
  If Text4 = "" Then Text3.SetFocus: Exit Sub
'  If Text5 = "" Then Text3.SetFocus: Exit Sub
  SaveSetting "T-PRO", "Koneksi", "Server", Text1.Text
  SaveSetting "T-PRO", "Koneksi", "Port", Text2.Text
  SaveSetting "T-PRO", "Koneksi", "Database", Text3.Text
  SaveSetting "T-PRO", "Koneksi", "User", Text4.Text
  SaveSetting "T-PRO", "Koneksi", "Password", Text5.Text
  End
End Sub

Private Sub vbButton2_Click()
    Unload Me
End Sub

