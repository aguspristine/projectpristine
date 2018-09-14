VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form antrianDispenser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dispenser Antrian"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7305
   Begin MSWinsockLib.Winsock ws1 
      Left            =   360
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton call 
      Caption         =   "call"
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   11
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton call 
      Caption         =   "call"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "cetak"
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   9
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "cetak"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label v 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lbl"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   1560
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lbl"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "s"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B"
      ForeColor       =   &H80000008&
      Height          =   1335
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label v 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lbl"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lbl"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "s"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      ForeColor       =   &H80000008&
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "antrianDispenser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub call_Click(Index As Integer)
Dim noAntri As Integer
Dim idAntri As Integer

    ReadRs "select min(noAntri) as panggil from antrianregistrasipasien " & _
        "where nKelompok = '" & lbl(Index).Caption & "' and visible=1 and tglAntrian between '" & TglJam_Server(True, False) & " 00:00' and '" & TglJam_Server(True, False) & " 23:59'"
    If rs.RecordCount = 0 Then
        MsgBox "sisa antrian 0"
        Exit Sub
    Else
        noAntri = rs(0)
    End If
    ReadRs "select idAntrian from antrianregistrasipasien where " & _
          "nKelompok = '" & lbl(Index).Caption & "' and visible=1 and tglAntrian between '" & TglJam_Server(True, False) & " 00:00' and '" & TglJam_Server(True, False) & " 23:59' " & _
        "and noAntri=" & noAntri
    idAntri = rs(0)
    
    If MsgBox("call " & lbl(Index).Caption & " : " & noAntri, vbYesNo) = vbNo Then Exit Sub
    WriteRs "update antrianregistrasipasien set visible = 3 where visible=2 and nKelompok = '" & lbl(Index).Caption & "'"
    WriteRs "update antrianregistrasipasien set visible = 2 where idAntrian=" & idAntri
    
    Call LoadData
    ReadRs "select * from loketip where loket=" & Index + 1
    If ws1.State <> sckClosed Then ws1.Close
    ws1.Connect rs(0), rs(1)
End Sub

Private Sub cmd_Click(Index As Integer)
Dim idAntrian As Integer
Dim noAntri As Integer

    ReadRs "select max(idAntrian) from antrianregistrasipasien "
    idAntrian = IIf(IsNull(rs(0)), 0, rs(0))
    ReadRs "select max(noAntri) from antrianregistrasipasien where nKelompok='" & lbl(Index).Caption & "' and tglAntrian between '" & TglJam_Server(True, False) & " 00:00' and '" & TglJam_Server(True, False) & " 23:59'"
    noAntri = IIf(IsNull(rs(0)), 0, rs(0))

    WriteRs "insert into antrianregistrasipasien values (" & idAntrian + 1 & "," & noAntri + 1 & "," & _
            "'" & TglJam_Server(True, False) & "','-','-'," & Index + 1 & ",'-','" & lbl(Index).Caption & "','-',1,'-')"
    
    Call LoadData
    MsgBox "Cetak antrian " & Index + 1 & ":" & noAntri + 1
End Sub

Private Sub Command1_Click()
    Call LoadData
End Sub

Public Sub LoadData()
Dim ii As Integer
For i = 0 To 1
    s(i).Caption = 0
    v(i).Caption = 0
    t(i).Caption = 0
Next

ReadRs "select nKelompok, count(noantri) as ttl from antrianregistrasipasien  " & _
        "where visible ='1' and tglAntrian between '" & TglJam_Server(True, False) & " 00:00' and '" & TglJam_Server(True, False) & " 23:59' " & _
        "group by nKelompok"
For i = 0 To rs.RecordCount - 1
    For ii = 0 To 1
        If lbl(ii).Caption = rs!nKelompok Then
            s(ii).Caption = rs!ttl
        End If
    Next
    rs.MoveNext
Next
        
ReadRs "select nKelompok, count(noantri) as ttl from antrianregistrasipasien  " & _
        "where  tglAntrian between '" & TglJam_Server(True, False) & " 00:00' and '" & TglJam_Server(True, False) & " 23:59' " & _
        "group by nKelompok"
For i = 0 To rs.RecordCount - 1
    For ii = 0 To 1
        If lbl(ii).Caption = rs!nKelompok Then
            t(ii).Caption = rs!ttl
        End If
    Next
    rs.MoveNext
Next

ReadRs "select noAntri,nKelompok from antrianregistrasipasien where visible=2 " & _
        "and  tglAntrian between '" & TglJam_Server(True, False) & " 00:00' and '" & TglJam_Server(True, False) & " 23:59'"
For i = 0 To rs.RecordCount - 1
     For ii = 0 To 1
        If lbl(ii).Caption = rs!nKelompok Then
             v(ii).Caption = rs!noAntri
         End If
    Next
    rs.MoveNext
Next
End Sub

Private Sub Form_Load()
    Call LoadData
End Sub
