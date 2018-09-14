VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMaster 
   Caption         =   "Form2"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   12735
   WindowState     =   2  'Maximized
   Begin MSForms.ComboBox cbo 
      Height          =   375
      Index           =   0
      Left            =   9240
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   2175
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3836;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lbl 
      Height          =   255
      Index           =   0
      Left            =   9120
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
      Caption         =   "Caption"
      Size            =   "3201;450"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txt 
      Height          =   375
      Index           =   0
      Left            =   9120
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   3015
      VariousPropertyBits=   746604571
      Size            =   "5318;661"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NameTable As String

Private Sub Form_Load()
Dim ArObj As Integer
Dim ii As Integer
Const CnvrtDigitToPx As Double = 200

    Me.Caption = "Pasien"
    ReadRs "select * from INFORMATION_SCHEMA.COLUMNS  where TABLE_NAME='" & NameTable & "'"
    For i = 0 To rs.RecordCount - 1
        ArObj = i + 1
        
        Load lbl(ArObj)
        lbl(ArObj).Move 200, ArObj * 500 ', CnvrtDigitToPx * CDbl(rs!CHARACTER_MAXIMUM_LENGTH)
        lbl(ArObj).visible = True
        lbl(ArObj).Caption = rs!COLUMN_NAME
        
'        ReadRs2 "select * from INFORMATION_SCHEMA.KEY_COLUMN_USAGE where TABLE_NAME='" & NameTable & "' and COLUMN_NAME='" & rs!COLUMN_NAME & "'"
'        If rs2.RecordCount <> 0 Then
'            'combobox
'            ReadRs2 "select * from INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS where CONSTRAINT_NAME ='" & rs2!CONSTRAINT_NAME & "'"
'            ReadRs2 "select * from INFORMATION_SCHEMA.TABLE_CONSTRAINTS      where CONSTRAINT_NAME='" & rs2!UNIQUE_CONSTRAINT_NAME & "'"
'            Load cbo(ArObj)
'            cbo(ArObj).Move 200 + lbl(ArObj).Width, ArObj * 500, CnvrtDigitToPx * CDbl(rs!CHARACTER_MAXIMUM_LENGTH)
'            cbo(ArObj).Visible = True
'
'            ReadRs2 "Select * from " & rs2!TABLE_NAME
'            For ii = 0 To rs2.RecordCount
'                cbo(ArObj).AddItem rs2(1)
'                rs2.MoveNext
'            Next
'        Else
            'textbox
            Load txt(ArObj)
            If IsNull(rs!CHARACTER_MAXIMUM_LENGTH) = False Then
                txt(ArObj).Move 200 + lbl(ArObj).Width, ArObj * 500, CnvrtDigitToPx * CDbl(rs!CHARACTER_MAXIMUM_LENGTH)
                txt(ArObj).visible = True
            End If
'        End If
        rs.MoveNext
    Next
End Sub
