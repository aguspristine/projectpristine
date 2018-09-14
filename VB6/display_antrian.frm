VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form display_antrian 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10590
   ClientLeft      =   195
   ClientTop       =   -195
   ClientWidth     =   19740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10590
   ScaleWidth      =   19740
   Begin MSWinsockLib.Winsock WS1 
      Left            =   1560
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1000
      LocalPort       =   1000
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   1440
      Top             =   3600
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   960
      Top             =   3600
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   720
      TabIndex        =   6
      Top             =   7200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   3600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   3600
   End
   Begin MSWinsockLib.Winsock WS2 
      Left            =   2040
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1000
      LocalPort       =   1000
   End
   Begin MSWinsockLib.Winsock WS3 
      Left            =   2520
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1000
      LocalPort       =   1000
   End
   Begin MSWinsockLib.Winsock WS4 
      Left            =   1560
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1000
      LocalPort       =   1000
   End
   Begin MSWinsockLib.Winsock WS5 
      Left            =   2040
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1000
      LocalPort       =   1000
   End
   Begin MSWinsockLib.Winsock WS6 
      Left            =   2520
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1000
      LocalPort       =   1000
   End
   Begin VB.Label lblJam 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11:11:11"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   39.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   4800
      TabIndex        =   10
      Top             =   9240
      Width           =   3105
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1.999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   80.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2010
      Index           =   2
      Left            =   6105
      TabIndex        =   2
      Top             =   4665
      Width           =   4395
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   12600
      Width           =   3975
   End
   Begin VB.Label lblWs 
      BackStyle       =   0  'Transparent
      Caption         =   "#data_updating"
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Image pic 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   4095
      Left            =   12840
      Stretch         =   -1  'True
      Top             =   7440
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label runText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SELAMAT DATANG DI RUMAH SAKIT KHUSUS DAERAH DUREN SAWIT."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   13080
      Width           =   11175
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1.999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   5
      Left            =   16335
      TabIndex        =   11
      Top             =   10440
      Width           =   5055
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   35.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   4680
      TabIndex        =   7
      Top             =   12000
      Width           =   8295
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   35.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   13200
      TabIndex        =   8
      Top             =   12000
      Width           =   8055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label lblconn 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   5
      Top             =   5640
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1.999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   4
      Left            =   4800
      TabIndex        =   4
      Top             =   7680
      Width           =   5055
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1.999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   3
      Left            =   5280
      TabIndex        =   3
      Top             =   6480
      Width           =   5055
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1.999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   99.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2520
      Index           =   1
      Left            =   5565
      TabIndex        =   1
      Top             =   3000
      Width           =   5475
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1.999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   99.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2520
      Index           =   0
      Left            =   5565
      TabIndex        =   0
      Top             =   1320
      Width           =   5475
   End
End
Attribute VB_Name = "display_antrian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const SND_APPLICATION = &H80
'The sound is played using an application-specific association
Const SND_ALIAS = &H10000
'The pszSound parameter is a system-event alias in the registry or the WIN.INI file.
'Do not use with either SND_FILENAME or SND_RESOURCE.
Const SND_ALIAS_ID = &H110000
'The pszSound parameter is a predefined sound identifier.
Const SND_ASYNC = &H1
'The sound is played asynchronously and PlaySound returns immediately after beginning the sound. To terminate an asynchronously played waveform sound, call PlaySound with pszSound set to NULL.
Const SND_FILENAME = &H20000
'The pszSound parameter is a filename
Const SND_LOOP = &H8
'The sound plays repeatedly until PlaySound is called again with the pszSound parameter set to NULL. You must also specify the SND_ASYNC flag to indicate an asynchronous sound event
Const SND_MEMORY = &H4
'A sound event’s file is loaded in RAM. The parameter specified by pszSound must point to an image of a sound in memory.
Const SND_NODEFAULT = &H2
'No default sound event is used. If the sound cannot be found, PlaySound returns silently without playing the default sound.
Const SND_NOSTOP = &H10
'The specified sound event will yield to another sound event that is already playing. If a sound cannot be played because the resource needed to generate that sound is busy playing another sound, the function immediately returns FALSE without playing the requested sound. If this flag is not specified, PlaySound attempts to stop the currently playing sound so that the device can be used to play the new sound.
Const SND_NOWAIT = &H2000
'If the driver is busy, return immediately without playing the sound.
Const SND_PURGE = &H40
'Sounds are to be stopped for the calling task. If pszSound is not NULL, all instances of the specified sound are stopped. If pszSound is NULL, all sounds that are playing on behalf of the calling task are stopped. You must also specify the instance handle to stop SND_RESOURCE events.
Const SND_RESOURCE = &H40004 'The pszSound parameter is a resource identifier; hmod must identify the instance that contains the resource.
Const SND_SYNC = &H0
'Synchronous playback of a sound event. PlaySound returns after the sound event completes.

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Dim tmt As Integer
Dim tmt2 As Integer
Dim tmt3 As Integer
Dim tmt4 As Integer
Dim onload As Boolean
Dim loket As Integer
Dim KedipLoket As Integer
Dim jenisAntrian As String
Dim TimeToRefresh As Integer
Dim vdeo As Integer
Dim reload As Boolean
Dim sora As Integer
Dim HFullscrenn As Double
Dim WFullscrenn As Double

Private Sub File1_DblClick()
'    For i = 0 To File1.ListCount - 1
'        Debug.Print File1.List(i)
'    Next
End Sub

Private Sub Form_DblClick()
    
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Debug.Print KeyCode
If KeyCode = 112 Then frmSetServer.Show: Unload Me
End Sub

Private Sub Form_Load()
    Timer1.Enabled = True
    For i = 0 To 5
        lbl(i).Caption = ""
    Next
    'File1.Path = App.Path & "\video"
    tmt3 = 10
    tmt = 100
    onload = True
    'Label1.Caption = App.Path
    tmt3 = 60
    
    sora = 70
    
    'DS_top = 307
    'DS_left = 544
    'DS_width = 816
    'DS_height = 461
    'Fullscreen_Enabled = False
    'DirectShow_Load_Media App.Path & "\video\Video.mp4"
    'vdeo = 0
    'DirectShow_Load_Media App.Path & "\video\" & File1.List(0)
'    DirectShow_Loop
    'DirectShow_Play
    'DirectShow_Volume sora
'    MsgBox "2"
'    pic.Visible = False
    
    Call OpenPortWinsock
    Call loadAntrian
    
    '@IPComputer,@Port,@Loket,@StatusEnabled
    Call WriteRs("delete from loketip where loket='1'")
    Call WriteRs("delete from loketip where loket='2'")
    Call WriteRs("delete from loketip where loket='3'")
    Call WriteRs("delete from loketip where loket='4'")
    
    Call WriteRs("insert into loketip values ('" & WS1.LocalIP & "','1001','1','1')")
    Call WriteRs("insert into loketip values ('" & WS1.LocalIP & "','1002','2','1')")
    Call WriteRs("insert into loketip values ('" & WS1.LocalIP & "','1003','3','1')")
    Call WriteRs("insert into loketip values ('" & WS1.LocalIP & "','1004','4','1')")
    
End Sub

Private Sub Label5_Click()

End Sub

Private Sub lbl_Click(Index As Integer)
    Call Form_DblClick
End Sub

Private Sub lblResep_DblClick()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    tmt = tmt + 1
    tmt4 = tmt4 + 1
    If tmt4 > 5 Then
    'If Val(Format(Now(), "ss")) Mod 5 = 0 Then
        Call loadAntrian
        'tmt = 0
        tmt4 = 0
    End If
    If tmt > 20 Then
        Timer1.Enabled = False
        tmt = 0
        reload = True
        lblWs.visible = False
        Call OpenPortWinsock
    End If
    
    
    
'    lblJam.Caption = Format(Now(), "hh:nn:ss")
'    If Val(Format(Now(), "ss")) Mod 10 = 0 Then
'        'pic.Picture = File1.Path & "\File1.Tag"
'        pic.Picture = LoadPicture(File1.Path & "\" & File1.List(Val(File1.Tag)))
'        File1.Tag = Val(File1.Tag) + 1
'        If Val(File1.Tag) > File1.ListCount - 1 Then File1.Tag = 0
'    End If
'    On Error GoTo Error_Handler
'    Label3.Caption = DirectShow_Position.CurrentPosition & "/" & DirectShow_Position.StopTime
'    If DirectShow_Position.CurrentPosition >= DirectShow_Position.StopTime Then
'            'DirectShow_Position.CurrentPosition = 0
'        vdeo = vdeo + 1
'        If vdeo > File1.ListCount - 1 Then vdeo = 0
'        DirectShow_Load_Media App.Path & "\video\" & File1.List(vdeo)
''    DirectShow_Loop
'        DirectShow_Play
'        DirectShow_Volume 0
'    End If
'Error_Handler:
End Sub

Private Sub OpenPortWinsock()
'    Timer1.Enabled = False
    
    If WS1.State <> 0 Then WS1.Close
    WS1.LocalPort = 1001
    WS1.Listen
    
    If WS2.State <> 0 Then WS2.Close
    WS2.LocalPort = 1002
    WS2.Listen
    
    If WS3.State <> 0 Then WS3.Close
    WS3.LocalPort = 1003
    WS3.Listen
    
    If WS4.State <> 0 Then WS4.Close
    WS4.LocalPort = 1004
    WS4.Listen
'
'    If WS5.State <> 0 Then WS5.Close
'    WS5.LocalPort = 2005
'    WS5.Listen
'
'    If WS6.State <> 0 Then WS6.Close
'    WS6.LocalPort = 2006
'    WS6.Listen
    
    
End Sub


Private Sub ClosePortWinsock()
'    Timer1.Enabled = False
    
    If WS1.State <> 0 Then WS1.Close
    'WS1.LocalPort = 2001
    'WS1.Listen
    
    If WS2.State <> 0 Then WS2.Close
    'WS2.LocalPort = 2002
    'WS2.Listen
    
    If WS3.State <> 0 Then WS3.Close
    'WS3.LocalPort = 2003
    'WS3.Listen
    
'    If WS4.State <> 0 Then WS4.Close
'    'WS4.LocalPort = 2004
'    'WS4.Listen
'
'    If WS5.State <> 0 Then WS5.Close
'    'WS5.LocalPort = 2005
'    'WS5.Listen
'
'    If WS6.State <> 0 Then WS6.Close
'    'WS6.LocalPort = 2006
'    'WS6.Listen
    
    
End Sub


Private Sub ws1_ConnectionRequest(ByVal requestID As Long)
    If WS1.State <> sckClosed Then
        WS1.Close
    End If
'    lblWs.Visible = True
    WS1.Accept requestID
    WS1.SendData "OK"
    
    Call ClosePortWinsock
    'WS1.Close
'    WS1.LocalPort = 2001
'    WS1.Listen
    
    lblWs.Tag = "1"
    If reload = False Then Exit Sub
    Call loadAntrian
End Sub

Private Sub ws2_ConnectionRequest(ByVal requestID As Long)
    If WS2.State <> sckClosed Then
        WS2.Close
    End If
'    lblWs.Visible = True
    WS2.Accept requestID
    WS2.SendData "OK"
    
    Call ClosePortWinsock
    'WS2.Close
'    WS2.LocalPort = 2002
'    WS2.Listen
    
    lblWs.Tag = "1"
    If reload = False Then Exit Sub
    Call loadAntrian
End Sub

Private Sub ws3_ConnectionRequest(ByVal requestID As Long)
    If WS3.State <> sckClosed Then
        WS3.Close
    End If
'    lblWs.Visible = True
    WS3.Accept requestID
    WS3.SendData "OK"
    
    Call ClosePortWinsock
'    WS3.Close
'    WS3.LocalPort = 2003
'    WS3.Listen
    
    lblWs.Tag = "1"
    If reload = False Then Exit Sub
    Call loadAntrian
End Sub

Private Sub ws4_ConnectionRequest(ByVal requestID As Long)
    If WS4.State <> sckClosed Then
        WS4.Close
    End If
'    lblWs.Visible = True
    WS4.Accept requestID
    WS4.SendData "OK"
    
    Call ClosePortWinsock
'    WS4.Close
'    WS4.LocalPort = 2004
'    WS4.Listen
    
    lblWs.Tag = "1"
    If reload = False Then Exit Sub
    Call loadAntrian
End Sub

Private Sub ws5_ConnectionRequest(ByVal requestID As Long)
    If WS5.State <> sckClosed Then
        WS5.Close
    End If
'    lblWs.Visible = True
    WS5.Accept requestID
    WS5.SendData "OK"
    
    Call ClosePortWinsock
'    WS5.Close
'    WS5.LocalPort = 2005
'    WS5.Listen
    
    lblWs.Tag = "1"
    If reload = False Then Exit Sub
    Call loadAntrian
End Sub

Private Sub ws6_ConnectionRequest(ByVal requestID As Long)
    If WS6.State <> sckClosed Then
        WS6.Close
    End If
'    lblWs.Visible = True
    WS6.Accept requestID
    WS6.SendData "OK"
    
    Call ClosePortWinsock
'    WS6.Close
'    WS6.LocalPort = 2006
'    WS6.Listen
    
    lblWs.Tag = "1"
    If reload = False Then Exit Sub
    Call loadAntrian
End Sub

Private Sub lblconn_DblClick()
    End
End Sub



Private Sub loadAntrian()
On Error Resume Next
Dim disada As Boolean

'    If reload = False Then Exit Sub
'    reload = True
    'Call LIST_RESEP
    lblWs.visible = True
    Set rs = Nothing
    ReadRs ("select* from antrianregistrasipasien  where visible ='2' and " & _
             "tglAntrian between '" & Format(Now(), "yyyy-mm-dd 00:00") & "' and '" & Format(Now(), "yyyy-mm-dd 23:59") & "'")
'    For i = 0 To 4
'        lbl(i).Caption = "-"
'    Next
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        For i = 0 To rs.RecordCount - 1
'            If rs!JenisPasien = "BPJS" Then
'                jenisAntrian = 1
'            Else
                jenisAntrian = "A"
'            End If
            If rs!nLoketCounter = 1 Then
                If lbl(0).Caption <> jenisAntrian & "-" & Format(rs!noAntri, "0##") Then disada = True
                'lbl(0).Caption = rs!JenisPasien & " : " & Format(rs!NoAntrian, "0##")
                lbl(0).Caption = jenisAntrian & "-" & Format(rs!noAntri, "0##")
                loket = 1
            End If
            If rs!nLoketCounter = 2 Then
                If lbl(1).Caption <> jenisAntrian & "-" & Format(rs!noAntri, "0##") Then disada = True
                'lbl(1).Caption = rs!JenisPasien & " : " & Format(rs!NoAntrian, "0##")
                lbl(1).Caption = jenisAntrian & "-" & Format(rs!noAntri, "0##")
                loket = 2
            End If
            If rs!nLoketCounter = 3 Then
                If lbl(2).Caption <> jenisAntrian & "-" & Format(rs!noAntri, "0##") Then disada = True
                'lbl(2).Caption = rs!JenisPasien & " : " & Format(rs!NoAntrian, "0##")
                lbl(2).Caption = jenisAntrian & "-" & Format(rs!noAntri, "0##")
                loket = 3
            End If
            If rs!nLoketCounter = 4 Then
                If lbl(3).Caption <> jenisAntrian & "-" & Format(rs!noAntri, "0##") Then disada = True
                'lbl(3).Caption = rs!JenisPasien & " : " & Format(rs!NoAntrian, "0##")
                lbl(3).Caption = jenisAntrian & "-" & Format(rs!noAntri, "0##")
                loket = 4
            End If
'            If rs!NoLoketCounter = 5 Then
'                If lbl(4).Caption <> jenisAntrian & "-" & Format(rs!NoAntrian, "0##") Then disada = True
'                'lbl(4).Caption = rs!JenisPasien & " : " & Format(rs!NoAntrian, "0##")
'                lbl(4).Caption = jenisAntrian & "-" & Format(rs!NoAntrian, "0##")
'                loket = 5
'            End If
'            If rs!NoLoketCounter = 6 Then
'                If lbl(5).Caption <> jenisAntrian & "-" & Format(rs!NoAntrian, "0##") Then disada = True
'                'lbl(4).Caption = rs!JenisPasien & " : " & Format(rs!NoAntrian, "0##")
'                lbl(5).Caption = jenisAntrian & "-" & Format(rs!NoAntrian, "0##")
'                loket = 6
'            End If
            
            If disada = True Then Call playSound(rs!noAntri)
            disada = False
            rs.MoveNext
        Next
    End If
    onload = False
    If Timer1.Enabled = False Then Timer1.Enabled = True
End Sub

Private Sub playSound(angka As Integer)
Dim t As Single
Dim belas As Boolean
Dim puluh As Boolean
Dim ratus As Boolean

    If onload = True Then Exit Sub
    lbl(loket - 1).BackColor = &H8080FF
    lbl(loket - 1).BackStyle = 1
    Timer2.Enabled = True
    Call sndPlaySound(App.Path & "\sound\nomor-urut.wav", SND_ASYNC Or SND_NODEFAULT)
    
    t = Timer
    Do
        DoEvents
    Loop Until Timer - t > 2
    
    If jenisAntrian = 1 Then Call sndPlaySound(App.Path & "\sound\satu.wav", SND_ALIAS Or SND_SYNC)
    If jenisAntrian = 2 Then Call sndPlaySound(App.Path & "\sound\dua.wav", SND_ALIAS Or SND_SYNC)
    
'    t = Timer
'    Do
'        DoEvents
'    Loop Until Timer - t > 1
    
    belas = False
    puluh = False
    ratus = False
    
    If angka > 199 And angka < 1000 Then ratus = True
    If angka > 99 And angka < 200 Then Call sndPlaySound(App.Path & "\sound\seratus.wav", SND_ALIAS Or SND_SYNC): angka = angka - 100
    If angka > 19 And angka < 100 Then puluh = True
    
    If angka < 20 And angka > 11 Then angka = angka - 10: belas = True
    
    If Len(CStr(angka)) = 2 And angka = 10 Then Call sndPlaySound(App.Path & "\sound\sepuluh.wav", SND_ALIAS Or SND_SYNC): GoTo loketttt
    If Len(CStr(angka)) = 2 And angka = 11 Then Call sndPlaySound(App.Path & "\sound\sebelas.wav", SND_ALIAS Or SND_SYNC): GoTo loketttt
    If Len(CStr(angka)) = 3 And angka = 100 Then Call sndPlaySound(App.Path & "\sound\seratus.wav", SND_ALIAS Or SND_SYNC): GoTo loketttt
    If Len(CStr(angka)) = 4 And angka = 1000 Then Call sndPlaySound(App.Path & "\sound\seribu.wav", SND_ALIAS Or SND_SYNC): GoTo loketttt
    
    For i = 1 To Len(CStr(angka))
        If ratus = False And angka > 200 And Val(Mid(angka, 2, 2)) = 10 Then Call sndPlaySound(App.Path & "\sound\sepuluh.wav", SND_ALIAS Or SND_SYNC): Exit For
        If ratus = False And angka > 200 And Val(Mid(angka, 2, 2)) = 11 Then Call sndPlaySound(App.Path & "\sound\sebelas.wav", SND_ALIAS Or SND_SYNC): Exit For
        If ratus = False And angka > 200 Then
            If Val(Mid(angka, 2, 2)) > 19 And puluh = False Then
                puluh = True
            Else
                puluh = False
            End If
        End If
        If ratus = False And angka > 200 And Val(Mid(angka, 2, 2)) < 20 And Val(Mid(angka, 2, 2)) > 11 Then
            'If Val(Mid(angka, 2, 2)) < 20 And Val(Mid(angka, 2, 2)) > 11 Then belas = True:
            Select Case Val(Mid(angka, 2, 2))
                Case 12
                    Call sndPlaySound(App.Path & "\sound\dua.wav", SND_ALIAS Or SND_SYNC)
                    Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
                Case 13
                    Call sndPlaySound(App.Path & "\sound\tiga.wav", SND_ALIAS Or SND_SYNC)
                    Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
                Case 14
                    Call sndPlaySound(App.Path & "\sound\empat.wav", SND_ALIAS Or SND_SYNC)
                    Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
                Case 15
                    Call sndPlaySound(App.Path & "\sound\lima.wav", SND_ALIAS Or SND_SYNC)
                    Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
                Case 16
                    Call sndPlaySound(App.Path & "\sound\enam.wav", SND_ALIAS Or SND_SYNC)
                    Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
                Case 17
                    Call sndPlaySound(App.Path & "\sound\tujuh.wav", SND_ALIAS Or SND_SYNC)
                    Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
                Case 18
                    Call sndPlaySound(App.Path & "\sound\delapan.wav", SND_ALIAS Or SND_SYNC)
                    Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
                Case 19
                    Call sndPlaySound(App.Path & "\sound\sembilan.wav", SND_ALIAS Or SND_SYNC)
                    Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
            End Select
            Exit For
        End If
        Select Case Mid(CStr(angka), i, 1)
           Case 1
               Call sndPlaySound(App.Path & "\sound\satu.wav", SND_ALIAS Or SND_SYNC)
           Case 2
               Call sndPlaySound(App.Path & "\sound\dua.wav", SND_ALIAS Or SND_SYNC)
           Case 3
               Call sndPlaySound(App.Path & "\sound\tiga.wav", SND_ALIAS Or SND_SYNC)
           Case 4
               Call sndPlaySound(App.Path & "\sound\empat.wav", SND_ALIAS Or SND_SYNC)
           Case 5
               Call sndPlaySound(App.Path & "\sound\lima.wav", SND_ALIAS Or SND_SYNC)
           Case 6
               Call sndPlaySound(App.Path & "\sound\enam.wav", SND_ALIAS Or SND_SYNC)
           Case 7
               Call sndPlaySound(App.Path & "\sound\tujuh.wav", SND_ALIAS Or SND_SYNC)
           Case 8
               Call sndPlaySound(App.Path & "\sound\delapan.wav", SND_ALIAS Or SND_SYNC)
           Case 9
               Call sndPlaySound(App.Path & "\sound\sembilan.wav", SND_ALIAS Or SND_SYNC)
        End Select
        

'        If ratus = False And angka > 200 Then
'            If Val(Mid(angka, 2, 2)) < 20 And Val(Mid(angka, 2, 2)) > 11 Then belas = True
'        End If

        If belas = True Then Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
        If puluh = True Then Call sndPlaySound(App.Path & "\sound\puluh.wav", SND_ALIAS Or SND_SYNC) ': puluh = False
        If angka > 19 And angka < 100 Then puluh = False
        If ratus = True Then Call sndPlaySound(App.Path & "\sound\ratus.wav", SND_ALIAS Or SND_SYNC): ratus = False
belas:
    Next
    
    
loketttt:
    Call sndPlaySound(App.Path & "\sound\loket.wav", SND_ASYNC Or SND_NODEFAULT)
    
    t = Timer
    Do
        DoEvents
    Loop Until Timer - t > 1
    Select Case loket
        Case 1
            Call sndPlaySound(App.Path & "\sound\satu.wav", SND_ALIAS Or SND_SYNC)
        Case 2
            Call sndPlaySound(App.Path & "\sound\dua.wav", SND_ALIAS Or SND_SYNC)
        Case 3
            Call sndPlaySound(App.Path & "\sound\tiga.wav", SND_ALIAS Or SND_SYNC)
        Case 4
            Call sndPlaySound(App.Path & "\sound\empat.wav", SND_ALIAS Or SND_SYNC)
        Case 5
            Call sndPlaySound(App.Path & "\sound\lima.wav", SND_ALIAS Or SND_SYNC)
        Case 6
            Call sndPlaySound(App.Path & "\sound\enam.wav", SND_ALIAS Or SND_SYNC)
        Case 7
            Call sndPlaySound(App.Path & "\sound\tujuh.wav", SND_ALIAS Or SND_SYNC)
        Case 8
            Call sndPlaySound(App.Path & "\sound\delapan.wav", SND_ALIAS Or SND_SYNC)
        Case 9
            Call sndPlaySound(App.Path & "\sound\sembilan.wav", SND_ALIAS Or SND_SYNC)
        Case 10
            Call sndPlaySound(App.Path & "\sound\sepuluh.wav", SND_ALIAS Or SND_SYNC)
    End Select
End Sub

Private Sub Timer2_Timer()
'    lbl(KedipLoket).FontBold = Not lbl(KedipLoket).FontBold
    tmt2 = tmt2 + 1
    If tmt2 > 10 Then
        Timer2.Enabled = False
        For i = 0 To 5
'            lbl(i).BackColor = &H8000000F
            lbl(i).BackStyle = 0
        Next
        tmt2 = 0
'        lblWs.Visible = False
    End If
End Sub

Private Sub Timer3_Timer()
    If runText.Left < 0 - runText.Width Then runText.Left = Screen.Width
    runText.Move runText.Left - 200
End Sub

Private Sub Timer4_Timer()
    tmt3 = tmt3 + 1
    If tmt3 > 60 Then
'        strSQL = "select distinct noantrian from AntrianPasienRegistrasi where TglAntrian > '" & Format(Now(), "yyyy-mm-dd") & " 00:00' and JenisPasien = 'bpjs' " '  group by jenispasien"
'        Call msubRecFO(rsa, strSQL)
'        If rsa.RecordCount <> 0 Then
'            lbl2(0).Caption = "TOTAL BPJS : " & rsa.RecordCount
'        End If
'        strSQL = "select distinct noantrian from AntrianPasienRegistrasi where TglAntrian > '" & Format(Now(), "yyyy-mm-dd") & " 00:00' and JenisPasien = 'UMUM'   " 'group by jenispasien"
'        Call msubRecFO(rsa, strSQL)
'        If rsa.RecordCount <> 0 Then
'            lbl2(1).Caption = "TOTAL UMUM : " & rsa.RecordCount
'        End If
'        tmt3 = 0
    End If
    lblJam.Caption = Format(Now(), "hh:nn:ss")
    On Error GoTo Error_Handler
    'Label3.Caption = Round(DirectShow_Position.CurrentPosition, 0) & "/" & Round(DirectShow_Position.StopTime, 0)
'    If DirectShow_Position.CurrentPosition >= DirectShow_Position.StopTime Then
'            'DirectShow_Position.CurrentPosition = 0
'        vdeo = vdeo + 1
'        If vdeo > File1.ListCount - 1 Then vdeo = 0
'        DirectShow_Load_Media App.Path & "\video\" & File1.List(vdeo)
''    DirectShow_Loop
'        DirectShow_Play
'        DirectShow_Volume sora
'    End If
'
'    If lblResep.Height > Screen.Height - pic2.Height - lblJam.Height Then
'        lblResep.Top = lblResep.Top - 700
'        If lblResep.Top + lblResep.Height < pic2.Top Then lblResep.Top = lblJam.Top
'    End If
Error_Handler:
End Sub



