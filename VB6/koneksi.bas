Attribute VB_Name = "koneksi"

'"Untuk memahami hati dan pikiran seseorang,
'jangan lihat apa yang sudah dia capai,
'tapi lihat pada apa yang dia cita-citakan.
                ' Kahlil Gibran, Sastrawan asal Lebanon

'"Jika tindakan-tindakan Anda mengilhami orang lain untuk bermimpi lebih,
'belajar lebih, bekerja lebih, dan menjadi lebih baik,
'Anda adalah seorang pemimpin."
'                                               John Quincy Adams (1767-1848),
'                                               Presiden Amerika Serikat (1825-1829)


'"everything looks so perfect from far away", Such Great Height - The Postal Service

'Public cn As New MYSQL_CONNECTION
'Public Rs As New MYSQL_RS
'Public Rs2 As New MYSQL_RS
'Public Rs3 As New MYSQL_RS
'Public RsWkt As New MYSQL_RS
Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rs2 As New ADODB.Recordset
Public rs3 As New ADODB.Recordset
Public RsWkt As New ADODB.Recordset
Public strSQL As String
Public strSQL2 As String
Public strSQL3 As String


Public qty_icon  As Long
Public OS As String

Sub Main()
  BukaKoneksi
End Sub

Public Sub BukaKoneksi()
On Error GoTo kantin
  
 ' cn.CursorLocation = adUseClient
 ' cn.Open "Provider=SQLNCLI10.1;Password=j4s4medik4;DataTypeCompatibility=80;Persist Security Info=True;User ID=sa;Initial Catalog=Poliklinik_DB;Data Source=.\SS2008R2"
  'cn.Open "DRIVER={MySQL ODBC 3.51 Driver}; " & _
          "SERVER=10.1.0.10; " & _
          "PORT=3306; " & _
          "DATABASE=dailyusageCC; " & _
          "USER=dailyusagecc; " & _
          "PASSWORD=dailyusagecc; " & _
          "OPTION=3; "
  
  'Dim Server, db, user, pwd As String

'  Server = GetTxt("Setting.txt", "DB", "server")
'  user = GetTxt("Setting.txt", "DB", "user")
'  pwd = Hex(GetTxt("Setting.txt", "DB", "pass"))
''  pwd = GetTxt("Setting.txt", "DB", "pass")
'  db = GetTxt("Setting.txt", "DB", "db")

  Server = GetSetting("T-PRO", "Koneksi", "Server")
  Port = GetSetting("T-PRO", "Koneksi", "Port")
  db = GetSetting("T-PRO", "Koneksi", "Database")
  user = GetSetting("T-PRO", "Koneksi", "User")
  pwd = GetSetting("T-PRO", "Koneksi", "Password")
  
'
  cn.CursorLocation = adUseClient
  'cn.OpenConnection "localhost", "root", "root", "db"
  cn.Open "DRIVER={MySQL ODBC 3.51 Driver}; " & _
          "SERVER=" & Server & "; " & _
          "PORT=" & Port & "; " & _
          "DATABASE=" & db & "; " & _
          "USER=" & user & "; " & _
          "PASSWORD=" & pwd & "; " & _
          "OPTION=3; "
  'cn.OpenConnection "10.1.0.10", "prgposcabang", "kagemane no jutsu", "poscabang"
           
  frmLogin.Show
  'display_antrian.Show
  'antrianDispenser.Show
  'MDIForm1.Show
  'frmListTable.Show
  
kantin:
    If Err.Number = -2147467259 Then
        frmSettingKoneksi.Show vbModal
    End If
  If Err.Number <> 0 And Err.Number <> -2147467259 Then
    SaveSetting "File_management", "err_log", Now, Err.Description
  End If
End Sub

Public Sub TutupKoneksi()
  cn.CloseConnection
End Sub

Public Function WriteRs(sql As String)
  Set rs = Nothing
  rs.Open sql, cn, adOpenStatic, adLockOptimistic
End Function

Public Function ReadRs(sql As String)
  Set rs = Nothing
  rs.Open sql, cn, adOpenStatic, adLockReadOnly
End Function

Public Function WriteRs2(sql As String)
  Set rs2 = Nothing
  rs2.Open sql, cn, adOpenStatic, adLockOptimistic
End Function

Public Function ReadRs2(sql As String)
  Set rs2 = Nothing
  rs2.Open sql, cn, adOpenStatic, adLockReadOnly
End Function

Public Function WriteRs3(sql As String)
  Set rs3 = Nothing
  rs3.Open sql, cn, adOpenStatic, adLockOptimistic
End Function

Public Function ReadRs3(sql As String)
  Set rs3 = Nothing
  rs3.Open sql, cn, adOpenStatic, adLockReadOnly
End Function

Public Function blokTxt(ttx As TextBox)
  ttx.SelStart = 0
  ttx.SelLength = Len(ttx)
  ttx.SetFocus
End Function

Public Function WriteRs4(sql As String)
  Set rs4 = Nothing
  rs4.Open sql, cn, adOpenStatic, adLockOptimistic
End Function

Public Function ReadRs4(sql As String)
  Set rs4 = Nothing
  rs4.Open sql, cn, adOpenStatic, adLockReadOnly
End Function

'Public Function Search_Text(grd As VSFlexGrid, ttx As TextBox, _
'                            jml_kolom As Double, jml_baris As Double)
'Dim brs As Double
'Dim kol As Double
'  For brs = 1 To jml_baris
'    For kol = 1 To jml_kolom
'      If InStr(1, grd.TextMatrix(brs, kol), ttx, vbTextCompare) Then
'        grd.Select brs, kol
'        grd.ShowCell brs, kol
'        grd.SetFocus
'        Exit Function
'      End If
'    Next
'  Next
'  MsgBox "Data tidak di temukan.!", vbOKOnly, "..:."
'  blokTxt ttx
'End Function

Public Function Format_tgl(txt As Date) As String
  Format_tgl = Format(txt, "yyyy-MM-dd")
End Function

Public Function Format_Rp(jml_duit As String) As String
  Format_Rp = "Rp. " & Format(jml_duit, "##,####,####") & ".-"
End Function

Public Function Format_jam(jam_brp As Date) As String
  Format_jam = Format(jam_brp, "hh:nn:ss")
End Function

Public Function Format_Tgl_Jam(kang_jam_berapa_sekarang As String) As String
  Format_Tgl_Jam = Format(kang_jam_berapa_sekarang, "yyyy-MM-dd hh:nn:ss")
End Function

Public Function TglJam_Server(tgl As Boolean, jam As Boolean) As String
  Set RsWkt = Nothing
  RsWkt.Open "select now()", cn, adOpenStatic, adLockReadOnly
  If jam = True And tgl = True Then TglJam_Server = Format_Tgl_Jam(RsWkt.Fields(0))
  If jam = False And tgl = True Then TglJam_Server = Format_tgl(RsWkt.Fields(0))
  If jam = True And tgl = False Then TglJam_Server = Format_jam(RsWkt.Fields(0))
End Function
'
'Public Function CustGrid(nama_grid As VSFlexGrid, Kolom As Integer, _
'                         Kolom_persentase As Integer, caption As String, _
'                         sembunyi As Boolean, TeksDitengah As Boolean)
'Dim lebar_persen As Integer
'
'  lebar_persen = nama_grid.Width / 100
'  nama_grid.ColWidth(Kolom) = Kolom_persentase * lebar_persen
'
'  nama_grid.TextMatrix(0, Kolom) = caption
'
'  nama_grid.ColHidden(Kolom) = sembunyi
'
'  If TeksDitengah = True Then nama_grid.ColAlignment(Kolom) = flexAlignCenterCenter
'  If TeksDitengah = False Then nama_grid.ColAlignment(Kolom) = flexAlignGeneral
'
'End Function

'Public Function rsfield(recrs As MYSQL_RS, Index As Integer) As String
'  rsfield = recrs.Fields(Index)
'End Function

Function getNewNumber(tableName As String, fieldName As String, keys As String)
Dim newKode As String
    ReadRs "select count(" & fieldName & ") from " & tableName
    If rs.RecordCount <> 0 Then
        newKode = keys & (Val(rs(0)) + 1)
    End If
    getNewNumber = newKode
End Function
Function getNewNumberWithDate(tableName As String, fieldName As String, keys As String, tgl As Date) As String
Dim newKode As String
    ReadRs "select count(" & fieldName & ") from " & tableName & " where tglRegistrasi = '" & Format_tgl(tgl) & "'"
    If rs.RecordCount <> 0 Then
        newKode = keys & (Val(rs(0)) + 1)
    End If
    getNewNumberWithDate = Format(tgl, "yyMMdd") & Format(newKode, "0###")
End Function
Function cekSudahTerdaftarPasien(nRM As String) As String
    ReadRs "select s.namaStatus from registrasi r,statusPulang s where r.nStatusPulang=s.nStatusPulang and nRm='" & nRM & "'"
    If rs.RecordCount = 0 Then
        cekSudahTerdaftarPasien = "Belum Terdaftar"
    Else
        cekSudahTerdaftarPasien = rs(0)
    End If
End Function

Function SelisihHariJam(ByVal Awal As Date, ByVal Akhir As Date) As String
Dim Detik As Double, jam As Long, HariX As String
Dim mnt, dtk As Integer
Dim JamLengkap As String
   
  If Awal > Akhir Then
    Detik = DateDiff("s", Format(Akhir, "dd/MM/yyyy hh:nn:ss"), Format(Awal, "dd/MM/yyyy hh:nn:ss"))
    jam = Detik \ 3600
    If jam > 23 Then
       hari = jam \ 24
       JamLengkap = Format((Akhir - Awal), "hh:mm:ss")
    Else
       hari = (jam \ 24) * (-1)
       JamLengkap = Format((Akhir - Awal), "hh:mm:ss")
    End If
      If hari = 0 Then
        SelisihHariJam = JamLengkap
      Else
        SelisihHariJam = hari & " hr, " & JamLengkap
      End If
    Exit Function
  Else
    Detik = DateDiff("s", Format(Awal, "dd/MM/yyyy hh:nn:ss"), Format(Akhir, "dd/MM/yyyy hh:nn:ss"))
    jam = Detik \ 3600
    If jam > 23 Then
       hari = jam \ 24
       JamLengkap = Format((Akhir - Awal), "hh:mm:ss")
    Else
       hari = (jam \ 24) * (-1)
       JamLengkap = Format((Akhir - Awal), "hh:mm:ss")
    End If
      If hari = 0 Then
        SelisihHariJam = JamLengkap
      Else
        SelisihHariJam = hari & " hr, " & JamLengkap
      End If
    Exit Function
  End If
End Function
Function getNewNumberWithDate1(tableName As String, fieldName As String, fieldTanggal As String, keys As String, tgl As Date) As String
Dim newKode As String
    ReadRs "select count(" & fieldName & ") from " & tableName & " where " & fieldTanggal & " = '" & Format_tgl(tgl) & "'"
    If rs.RecordCount <> 0 Then
        newKode = keys & (Val(rs(0)) + 1)
    End If
    getNewNumberWithDate1 = Format(tgl, "yyMMdd") & Format(newKode, "0###")
End Function
Function cekSTokRuangan(nBarang As String, nRuangan As String) As String
    ReadRs "select sum(qtystok) as qtystok from transaksistok where nBarang = '" & nBarang & "' and nRuangan= '" & nRuangan & "' and visible='1'"
    If rs.RecordCount = 0 Then
        cekSudahTerdaftarPasien = "Data Belum Teregistrasi"
    Else
        cekSudahTerdaftarPasien = rs!qtystok
    End If
End Function
