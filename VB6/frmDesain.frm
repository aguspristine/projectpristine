VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmDesain 
   Caption         =   "Desain Form"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15450
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   15450
   WindowState     =   2  'Maximized
   Begin VSFlex8LCtl.VSFlexGrid fg 
      Height          =   7035
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   15195
      _cx             =   26802
      _cy             =   12409
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AllowUserResizing=   0
      SelectionMode   =   0
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
      FormatString    =   $"frmDesain.frx":0000
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
   Begin MSForms.ComboBox ComboBox1 
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   4095
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "7223;873"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      Caption         =   "Nama Tabel :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "frmDesain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
    ReadRs "SELECT     COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH FROM         INFORMATION_SCHEMA.COLUMNS WHERE     (TABLE_NAME = '" & ComboBox1.Text & "')"
    fg.Rows = rs.RecordCount + 1
    For i = 0 To rs.RecordCount - 1
        fg.TextMatrix(i + 1, 1) = rs!COLUMN_NAME
        fg.TextMatrix(i + 1, 2) = rs!DATA_TYPE
        fg.TextMatrix(i + 1, 3) = IIf(IsNull(rs!CHARACTER_MAXIMUM_LENGTH) = True, "", rs!CHARACTER_MAXIMUM_LENGTH)
        fg.TextMatrix(i + 1, 5) = 1
        rs.MoveNext
    Next
    
End Sub

Private Sub fg_Click()
Dim str As String

    fg.ComboList = ""
    str = "-|"
    If fg.Col = 4 Then
        fg.ComboList = "-|TextBox|ComboBox|CheckBox"
    End If
    If fg.Col = 7 Then
        ReadRs "select * from INFORMATION_SCHEMA.TABLES "
        For i = 0 To rs.RecordCount - 1
            str = str & "|" & rs!TABLE_NAME
            rs.MoveNext
        Next
         fg.ComboList = str
    End If
    If fg.Col = 8 Then
        ReadRs "SELECT     COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH FROM         INFORMATION_SCHEMA.COLUMNS WHERE     (TABLE_NAME = '" & fg.TextMatrix(fg.Row, 7) & "') "
        For i = 0 To rs.RecordCount - 1
            str = str & "|" & rs!COLUMN_NAME
            rs.MoveNext
        Next
        fg.ComboList = str
    End If
    If fg.Col = 9 Then
        ReadRs "SELECT     COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH FROM         INFORMATION_SCHEMA.COLUMNS WHERE     (TABLE_NAME = '" & fg.TextMatrix(fg.Row, 7) & "') "
        For i = 0 To rs.RecordCount - 1
            str = str & "|" & rs!COLUMN_NAME
            rs.MoveNext
        Next
        fg.ComboList = str
    End If
End Sub

Private Sub Form_Load()
    ReadRs "select * from INFORMATION_SCHEMA.TABLES "
    For i = 0 To rs.RecordCount - 1
        ComboBox1.AddItem rs!TABLE_NAME
        rs.MoveNext
    Next
    fg.Editable = flexEDKbdMouse
    fg.AllowUserResizing = flexResizeColumns
End Sub


