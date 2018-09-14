VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListTable 
   Caption         =   "List Table"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11145
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   11145
   WindowState     =   2  'Maximized
   Begin MSComctlLib.TreeView tv 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   13150
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
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
End
Attribute VB_Name = "frmListTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub loadMenu()
Dim txtKey0, txtKey1, txtKey2, txtKey3, txtKey4, txtKey5, txtKey6, txtKey7, txtNama, txtSatuan As String
Dim nodX As Node

    tv.Nodes.Clear
    ReadRs "select * from usermenu order by urutan"
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

End Sub

Private Sub Form_Load()
    Call loadMenu
End Sub
