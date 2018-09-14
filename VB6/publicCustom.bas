Attribute VB_Name = "publicCustom"
Public Sub isiGrid(namaForm As String, namaGrid As VSFlexGrid, namaRs As Recordset, setColumn As String)
Dim ii As Integer
'On Error GoTo Hell
    
    
    
    Dim splt() As String
    Dim colAtt() As String
    
    namaGrid.Rows = 1
    splt = Split(setColumn + ",", ",")
    namaGrid.Cols = namaRs.Fields.Count + 1
    namaGrid.ColWidth(0) = 500
    namaGrid.TextMatrix(0, 1) = "No"
    'namaGrid.TextMatrix(0, 2) = namaRs.Fields.Count + 1
    For i = 0 To namaRs.Fields.Count - 1
        colAtt = Split(splt(i), "=")
        namaGrid.ColWidth(i + 1) = colAtt(1)
        namaGrid.TextMatrix(0, i + 1) = colAtt(0)
    Next
    
    If namaRs.RecordCount = 0 Then GoTo holiday
    namaGrid.Rows = namaRs.RecordCount + 1
    namaGrid.Cols = namaRs.Fields.Count + 1
    namaRs.MoveFirst
    For i = 0 To namaRs.RecordCount - 1
        namaGrid.TextMatrix(i + 1, 0) = i + 1
        For ii = 0 To namaRs.Fields.Count + 1 - 2
            namaGrid.TextMatrix(i + 1, ii + 1) = IIf(IsNull(namaRs(ii)) = True, "", namaRs(ii))
        Next
        namaRs.MoveNext
    Next
    'GoTo holiday
    
    
    
'    ReDim arrGrid(namaRs.Fields.Count + 1, 2)
'    namaRs.MoveFirst
'    namaGrid.Cols = namaRs.Fields.Count + 1
'    namaGrid.ColWidth(0) = 500
'    arrGrid(0, 0) = "500"
'    arrGrid(0, 1) = "No"
'    arrGrid(0, 2) = namaRs.Fields.Count + 1
'    For i = 0 To namaRs.Fields.Count - 1
'        namaGrid.ColWidth(i + 1) = (namaRs(i).ActualSize) * 150
'        namaGrid.TextMatrix(0, i + 1) = namaRs(i).Name
'        arrGrid(i + 1, 0) = (namaRs(i).ActualSize) * 110
'        arrGrid(i + 1, 1) = namaRs(i).Name
'    Next
'    Exit Sub
'
holiday:
    
'    namaGrid.Cols = arrGrid(0, 2)
'    For i = 0 To arrGrid(0, 2) - 1
'        namaGrid.ColWidth(i) = arrGrid(i, 0)
'        namaGrid.TextMatrix(0, i) = arrGrid(i, 1)
'    Next
    
Hell:
End Sub

Public Sub captionGrid(namaForm As String, namaGrid As VSFlexGrid, jml As Integer, setColumn As String)
Dim i As Integer

    splt = Split(setColumn + ",", ",")
    namaGrid.Cols = jml + 1
    namaGrid.ColWidth(0) = 500
    namaGrid.TextMatrix(0, 1) = "No"
    For i = 0 To jml - 1
        colAtt = Split(splt(i), "=")
        namaGrid.ColWidth(i + 1) = colAtt(1)
        namaGrid.TextMatrix(0, i + 1) = colAtt(0)
    Next
End Sub

Public Sub loadDataCombo(nmDataCombo As Object, namaRs As ADODB.Recordset, strQuery As String)
On Error GoTo errLoadDataCombo

    Set namaRs = New ADODB.Recordset
    Set namaRs = cn.Execute(strQuery)
    
    Set nmDataCombo.RowSource = namaRs
    nmDataCombo.BoundColumn = namaRs(0).Name
    nmDataCombo.ListField = namaRs(1).Name
Exit Sub
errLoadDataCombo:
    MsgBox "Load Data Combo Error ", vbExclamation, "Warning"
End Sub

Public Sub centerForm(ByRef oForm1 As Form, ByVal oForm2 As Form)
    oForm1.Left = (oForm2.Width - oForm1.Width) / 2
    oForm1.Top = (oForm2.Height - 1500 - oForm1.Height) / 2
End Sub
