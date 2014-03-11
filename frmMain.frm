VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "YM! Answering Machine - TES BAHASA INGGRIS ver 1.0.1"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   3570
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmMain.frx":0000
      Top             =   4560
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   6120
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Timer tmrAutoReplay 
      Interval        =   500
      Left            =   6600
      Top             =   240
   End
   Begin VB.OptionButton optON 
      Caption         =   "ON"
      Height          =   195
      Left            =   1470
      TabIndex        =   3
      Top             =   150
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.OptionButton optOFF 
      Caption         =   "OFF"
      Height          =   195
      Left            =   2325
      TabIndex        =   2
      Top             =   150
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   " [ Daftar pesan yang masuk ] "
      Height          =   3570
      Left            =   120
      TabIndex        =   0
      Top             =   435
      Width           =   6855
      Begin VB.ListBox lstMessageIn 
         Height          =   3180
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Balas Otomatis"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1230
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PROGRAMMER            As String = "K4m4r82"
Private Const MY_EMIL               As String = "k4m4r82@yahoo.com"
Private Const MY_BLOG               As String = "http://coding4ever.wordpress.com"

Private Const ABOUT                 As String = "<b>YM! Answering Machine - TES BAHASA INGGRIS ver <version></b>" & vbCrLf & _
                                                "             By <#0A8EFE><b>" & PROGRAMMER & "</b></#>" & vbCrLf & _
                                                "             Email : " & MY_EMIL & vbCrLf & _
                                                "             Blog : " & MY_BLOG & vbCrLf & vbCrLf
                                
Private Const BANTUAN               As String = "<b>YM! Answering Machine - TES BAHASA INGGRIS ver <version></b>" & vbCrLf & _
                                                "             By <#0A8EFE><b>" & PROGRAMMER & "</b></#>" & vbCrLf & _
                                                "             Email : " & MY_EMIL & vbCrLf & _
                                                "             Blog : " & MY_BLOG & vbCrLf & vbCrLf & _
                                                "             Daftar keyword yang tersedia :" & vbCrLf & _
                                                "             o <b>about</b> -> informasi program" & vbCrLf & _
                                                "             o <b>mulai</b> -> untuk memulai tes bahasa inggris" & vbCrLf & _
                                                "             o <b>soal</b> -> untuk mendapatkan soal tes bahasa inggris " & vbCrLf & _
                                                "             o <b>soalterakhir</b> -> informasi soal terakhir" & vbCrLf & _
                                                "             o <b>jawab jawaban</b> -> untuk menjawab soal. contoh <b>JAWAB A</b>" & vbCrLf & _
                                                "             o <b>batal</b> -> untuk mengabaikan soal terakhir/enggak sanggup jawab :D" & vbCrLf & _
                                                "             o <b>selesai</b> -> untuk mengakhiri tes bahasa Inggris" & vbCrLf & vbCrLf
                      
Private Const SESSION_TIME_OUT      As String = "Maaf sesi tes bahasa Inggris Anda sudah selesai/belum dibuat, ketik <b>MULAI</b> untuk memulai tes." & vbCrLf & vbCrLf
Private Const SOAL_SUDAH_DIJAWAB    As String = "Maaf soal terakhir tidak ditemukan/sudah dijawab, ketik <b>SOAL</b> untuk mendapatkan soal baru atau <b>SELESAI</b> untuk melihat hasil tes." & vbCrLf & vbCrLf
Private Const SESSION_OK            As String = "Selamat !!! sesi tes bahasa Inggris Anda sudah dibuat, ketik <b>SOAL</b> untuk memulai tes." & vbCrLf & vbCrLf

Private Function getSoalNumber(ByVal ymID As String, ByVal sessionID As Long, Optional increment As Long = 1) As Long
    strSql = "SELECT COUNT(*) + " & increment & " FROM history WHERE ym_id = '" & ymID & "' AND sesi_id = " & sessionID & ""
    getSoalNumber = CLng(dbGetValue(strSql, 1))

End Function

Private Function acakSoal(ByVal ymID As String, ByVal tgl As String, ByRef id As Long) As String
    Dim rsRandom        As ADODB.Recordset
    
    Dim jmlRecord       As Long
    Dim recNumber       As Long
    
    On Error GoTo errHandle
    
    strSql = "SELECT id, soal FROM bank_soal " & _
             "WHERE id NOT IN (SELECT soal_id FROM history WHERE ym_id = '" & ymID & "' AND tanggal = #" & tgl & "#)"
    Set rsRandom = openRecordset(strSql)
    If Not rsRandom.EOF Then
        jmlRecord = getRecordCount(rsRandom)
        
        Randomize
        recNumber = Int((jmlRecord - 1) * Rnd)
        rsRandom.Move recNumber
        id = rsRandom("id").Value
        acakSoal = rsRandom("soal").Value
        
    Else
        id = 0
        acakSoal = "Maaf stok soal habis :D"
    End If
    Call closeRecordset(rsRandom)
    
    Exit Function
errHandle:
    acakSoal = "Maaf stok soal habis :D"
End Function

Private Function getLastSoalID(ByVal ymID As String, ByVal sesiID As Long, Optional cekJawaban As Boolean = True) As Long
    Dim lastID  As Long
    
    If cekJawaban Then
        strSql = "SELECT MAX(id) FROM history WHERE ym_id = '" & ymID & "' AND sesi_id = " & sesiID & " AND jawaban IS NULL AND batal = 0"
    Else
        strSql = "SELECT MAX(id) FROM history WHERE ym_id = '" & ymID & "' AND sesi_id = " & sesiID & " AND batal = 0"
    End If
    lastID = CLng(dbGetValue(strSql, 0))
    
    strSql = "SELECT soal_id FROM history WHERE id = " & lastID & ""
    getLastSoalID = CLng(dbGetValue(strSql, 0))
End Function

Private Function getLastSessionID(ByVal ymID As String) As Long
    strSql = "SELECT MAX(sesi_id) FROM sesi WHERE ym_id = '" & ymID & "' AND time_out = 0"
    getLastSessionID = CLng(dbGetValue(strSql, 0))
End Function

Private Function getLastSessionTime(ByVal sessionID As Long) As String
    Dim tmp As String
    
    strSql = "SELECT jam FROM sesi WHERE sesi_id = " & sessionID & ""
    tmp = CLng(dbGetValue(strSql, 0))
    If IsDate(tmp) Then tmp = Format(tmp, "hh:mm:ss")
    
    getLastSessionTime = tmp
End Function

Private Function rep(ByVal Kata As String) As String
    rep = Replace(Kata, "'", "''")
End Function

Private Sub pause(ByVal interval As Variant)
    Dim Current As Variant
    
    Current = Timer
    Do While Timer - Current < Val(interval)
        DoEvents
    Loop
End Sub

Private Sub optOFF_Click()
    tmrAutoReplay.Enabled = False
End Sub

Private Sub optON_Click()
    tmrAutoReplay.Enabled = True
End Sub

Private Sub tmrAutoReplay_Timer()
    Dim hwndYMMainClass     As Long
    
    Dim soalID              As Long
    Dim sessionID           As Long
    Dim menit               As Long
    
    Dim benar               As Long
    Dim salah               As Long
    Dim nilai               As Long
    Dim ret                 As Long
    
    Dim sessionTime         As String
    Dim soal                As String
    Dim msg                 As String
    Dim keyword             As String
    Dim arrKeyword()        As String
    Dim param               As String
    Dim jawaban             As String
    Dim analisa             As String
    Dim ymID                As String
    
    Dim tgl                 As String
    Dim jam                 As String
    Dim skrg                As String
    Dim lama                As String
    Dim arrJam()            As String
    
    'On Error GoTo errHandle
    
    Select Case ymVersion
        Case "8"
            hwndYMMainClass = FindWindow("YSearchMenuWndClass", vbNullString)
    
        Case "9"
            hwndYMMainClass = FindWindow("ATL:007C07F0", vbNullString)
            
        Case "10"
            hwndYMMainClass = FindWindow("CConvWndBase", vbNullString)
            
    End Select
    
    If hwndYMMainClass <> 0 Then
        tgl = Format(Now, "yyyy/MM/dd")
        jam = Format(Now, "hh:mm:ss")
        
        Call pause(1)
        
        ymID = getYMID(hwndYMMainClass)
        msg = getYMMessage(hwndYMMainClass)
        
        sessionID = getLastSessionID(ymID)
        sessionTime = getLastSessionTime(sessionID)
        
        lstMessageIn.AddItem ymID & ": " & msg
        lstMessageIn.ListIndex = lstMessageIn.ListCount - 1
        
        strSql = "INSERT INTO ym_log (ym_id, tanggal, jam, pesan) VALUES ('" & ymID & "', #" & tgl & "#, #" & jam & "#, '" & rep(msg) & "')"
        conn.Execute strSql
        
        If Len(msg) > 0 Then
            If InStr(1, msg, " ") > 0 Then
                arrKeyword = Split(msg, " ")
                keyword = arrKeyword(0)
                
            Else
                keyword = msg
                ReDim arrKeyword(0)
            End If
            
        Else
            keyword = ""
            ReDim arrKeyword(0)
        End If
        
        If UBound(arrKeyword) > 0 Then param = UCase$(arrKeyword(1))
        
        Select Case UCase$(keyword)
            Case "ABOUT"
                msg = Replace(ABOUT, "<version>", appVersion)
                
            Case "BANTUAN"
                msg = Replace(BANTUAN, "<version>", appVersion)
                
            Case "MULAI"
                If Not (sessionID > 0) Then
                    strSql = "INSERT INTO sesi (ym_id, tanggal, jam) VALUES ('" & ymID & "', #" & tgl & "#, #" & jam & "#)"
                    conn.Execute strSql
        
                    msg = SESSION_OK
                    
                Else
                    skrg = Format(Now, "hh:mm:ss")
                    If IsDate(skrg) And IsDate(sessionTime) Then
                        lama = Format(TimeValue(skrg) - TimeValue(sessionTime), "hh:mm:ss")
                        
                        arrJam = Split(lama, ":")
                        menit = (arrJam(0) * 60) + arrJam(1)
                    Else
                        menit = 0
                    End If
                    
                    If menit >= 30 Then 'session time out max 30 menit
                        strSql = "INSERT INTO sesi (ym_id, tanggal, jam) VALUES ('" & ymID & "', #" & tgl & "#, #" & jam & "#)"
                        conn.Execute strSql
                        
                        msg = SESSION_OK
                    
                    Else
                        msg = "Sesi Anda belum habis, silahkan melanjutkan menjawab soal atau <b>SELESAI</b> untuk melihat hasil tes." & vbCrLf & vbCrLf
                    End If
                End If
                
            Case "SOAL"
                If Not (sessionID > 0) Then
                    msg = SESSION_TIME_OUT
                Else
                    soalID = getLastSoalID(ymID, sessionID)
        
                    'cek soal terakhir sudah dijawab atau belum
                    strSql = "SELECT MAX(id) FROM history WHERE ym_id = '" & ymID & "' AND sesi_id = " & sessionID & " AND soal_id = " & soalID & " AND batal = 0"
                    ret = CLng(dbGetValue(strSql, 0))
                    If ret > 0 Then
                        strSql = "SELECT jawaban FROM history WHERE ym_id = '" & ymID & "' AND sesi_id = " & sessionID & " AND batal = 0 " & _
                                 "AND id = (SELECT MAX(id) FROM history WHERE ym_id = '" & ymID & "' AND sesi_id = " & sessionID & " AND soal_id = " & soalID & " AND batal = 0)"
                        jawaban = CStr(dbGetValue(strSql, ""))
                        If Len(jawaban) > 0 Then 'soal terakhir sudah dijawab
                            msg = acakSoal(ymID, tgl, soalID) & vbCrLf & _
                                  "Ketik: <B>JAWAB JAWABAN</B> untuk menjawab. Contoh: <B>JAWAB A</B>" & vbCrLf & vbCrLf
                            msg = Replace(msg, "[A]", "<b><red>[</red>A<red>]</red></b>")
                            msg = Replace(msg, "[B]", "<b><red>[</red>B<red>]</red></b>")
                            msg = Replace(msg, "[C]", "<b><red>[</red>C<red>]</red></b>")
                            msg = Replace(msg, "[D]", "<b><red>[</red>D<red>]</red></b>")
                            msg = Replace(msg, "[E]", "<b><red>[</red>E<red>]</red></b>")
                            
                            msg = "<b>Soal Nomor " & getSoalNumber(ymID, sessionID) & "</b>" & vbCrLf & _
                                  msg
                            
                            strSql = "INSERT INTO history (ym_id, sesi_id, tanggal, jam, soal_id) VALUES ('" & _
                                     ymID & "', " & sessionID & ", #" & tgl & "#, #" & jam & "#, " & soalID & ")"
                            conn.Execute strSql
                            
                        Else
                            msg = "Maaf soal terakhir belum dijawab, ketik <b>SOALTERAKHIR</b> untuk melihat soal terakhir." & vbCrLf & vbCrLf
                        End If
                        
                    Else
                        msg = acakSoal(ymID, tgl, soalID) & vbCrLf & _
                              "Ketik: <B>JAWAB JAWABAN</B> untuk menjawab. Contoh: <B>JAWAB A</B>" & vbCrLf & vbCrLf
                            
                        msg = Replace(msg, "[A]", "<b><red>[</red>A<red>]</red></b>")
                        msg = Replace(msg, "[B]", "<b><red>[</red>B<red>]</red></b>")
                        msg = Replace(msg, "[C]", "<b><red>[</red>C<red>]</red></b>")
                        msg = Replace(msg, "[D]", "<b><red>[</red>D<red>]</red></b>")
                        msg = Replace(msg, "[E]", "<b><red>[</red>E<red>]</red></b>")
                        
                        msg = "<b>Soal Nomor " & getSoalNumber(ymID, sessionID) & "</b>" & vbCrLf & _
                              msg
                              
                        strSql = "INSERT INTO history (ym_id, sesi_id, tanggal, jam, soal_id) VALUES ('" & _
                                 ymID & "', " & sessionID & ", #" & tgl & "#, #" & jam & "#, " & soalID & ")"
                        conn.Execute strSql
                    End If
                End If
            
            Case "SOALTERAKHIR"
                If Not (sessionID > 0) Then
                    msg = SESSION_TIME_OUT
                    
                Else
                    soalID = getLastSoalID(ymID, sessionID, False)
                    
                    strSql = "SELECT soal FROM bank_soal WHERE id = " & soalID & ""
                    soal = CStr(dbGetValue(strSql, ""))
                    If Len(soal) > 0 Then
                        msg = soal & vbCrLf & _
                              "Ketik: <B>JAWAB JAWABAN</B> untuk menjawab. Contoh: <B>JAWAB A</B>" & vbCrLf & vbCrLf
                        msg = Replace(msg, "[A]", "<b><red>[</red>A<red>]</red></b>")
                        msg = Replace(msg, "[B]", "<b><red>[</red>B<red>]</red></b>")
                        msg = Replace(msg, "[C]", "<b><red>[</red>C<red>]</red></b>")
                        msg = Replace(msg, "[D]", "<b><red>[</red>D<red>]</red></b>")
                        msg = Replace(msg, "[E]", "<b><red>[</red>E<red>]</red></b>")
                        msg = "<b>Soal Nomor " & getSoalNumber(ymID, sessionID, 0) & "</b>" & vbCrLf & _
                              msg
                    Else
                        msg = "Maaf soal terakhir tidak ditemukan, ketik <b>SOAL</b> untuk mendapatkan soal baru atau <b>SELESAI</b> untuk melihat hasil tes." & vbCrLf & vbCrLf
                    End If
                End If
                
            Case "JAWAB"
                If Not (sessionID > 0) Then
                    msg = SESSION_TIME_OUT
                Else
                    If Not (Len(param) > 0) Then
                        msg = "Maaf jawaban kosong, ketik <b>JAWAB JAWABAN</b> untuk menjawab soal. Contoh: JAWAB A" & vbCrLf & vbCrLf
                        
                    Else
                        soalID = getLastSoalID(ymID, sessionID)
                        
                        If soalID > 0 Then
                            'ambil jawaban soal di master bank soal
                            strSql = "SELECT jawaban FROM bank_soal WHERE id = " & soalID & ""
                            jawaban = CStr(dbGetValue(strSql, ""))
                            jawaban = UCase$(jawaban)
                            
                            If jawaban = param Then 'bandingkan dengan jawaban peserta
                                msg = "Jawaban Anda benar =D>"
                                strSql = "UPDATE history SET jawaban = '" & jawaban & "', hasil_jawaban = 1 " & _
                                         "WHERE ym_id = '" & ymID & "' AND sesi_id = " & sessionID & " AND soal_id = " & soalID & ""
                            Else
                                msg = "Jawaban Anda salah =)) jawaban yang benar: <b>" & jawaban & "</b>"
                                strSql = "UPDATE history SET jawaban = '" & jawaban & "', hasil_jawaban = 0 " & _
                                         "WHERE ym_id = '" & ymID & "' AND sesi_id = " & sessionID & " AND soal_id = " & soalID & ""
                            End If
                            conn.Execute strSql
                            
                            strSql = "SELECT analisa FROM bank_soal WHERE id = " & soalID & ""
                            analisa = CStr(dbGetValue(strSql, ""))
                            
                            msg = msg & vbCrLf & "Analisa/Keterangan:" & vbCrLf & analisa & vbCrLf & vbCrLf
                            
                        Else
                            msg = SOAL_SUDAH_DIJAWAB
                        End If
                    End If
                End If
                
            Case "BATAL"
                If Not (sessionID > 0) Then
                    msg = SESSION_TIME_OUT
                Else
                    soalID = getLastSoalID(ymID, sessionID)
                    
                    If soalID > 0 Then
                        strSql = "UPDATE history SET batal = 1 " & _
                                 "WHERE id = (SELECT MAX(id) FROM history WHERE ym_id = '" & ymID & "' AND sesi_id = " & sessionID & " AND soal_id = " & soalID & " AND batal = 0)"
                        conn.Execute strSql
                        
                        msg = "Soal terakhir sudah dibatalkan, ketik <b>SOAL</b> untuk mendapatkan soal baru atau <b>SELESAI</b> untuk melihat hasil tes." & vbCrLf & vbCrLf
                        
                    Else
                        msg = SOAL_SUDAH_DIJAWAB
                    End If
                End If
                
            Case "SELESAI"
                If Not (sessionID > 0) Then
                    msg = SESSION_TIME_OUT
                Else
                    strSql = "SELECT COUNT(*) FROM history " & _
                             "WHERE ym_id = '" & ymID & "' AND sesi_id = " & sessionID & " AND hasil_jawaban = 1 AND batal = 0"
                    benar = CInt(dbGetValue(strSql, 0))
    
                    strSql = "SELECT COUNT(*) FROM history " & _
                             "WHERE ym_id = '" & ymID & "' AND sesi_id = " & sessionID & " AND hasil_jawaban = 0 AND batal = 0"
                    salah = CInt(dbGetValue(strSql, 0))
    
                    nilai = benar * 5
                    
                    msg = "Hasil tes Anda: benar = " & benar & ", salah = " & salah & ", nilai = " & nilai & "" & vbCrLf & _
                          "                    ooO Terima kasih sudah mendownload sample program ini :) Ooo" & vbCrLf & vbCrLf
                          
                    strSql = "UPDATE sesi SET time_out = 1 WHERE sesi_id = " & sessionID & ""
                    conn.Execute strSql
                End If
                
            Case Else
                msg = "Keyword <b>" & msg & "</b> tidak terdaftar. Ketik <B>BANTUAN</B> untuk informasi lebih lanjut" & vbCrLf & vbCrLf
        End Select
        
        Call ymChatSend(hwndYMMainClass, msg)
        Call closeYM(hwndYMMainClass)
        
    End If
    
    Exit Sub
errHandle:
    msg = "Keyword <b>" & msg & "</b> tidak terdaftar. Ketik <B>BANTUAN</B> untuk informasi lebih lanjut" & vbCrLf & vbCrLf
    Call ymChatSend(hwndYMMainClass, msg)
    Call closeYM(hwndYMMainClass)
End Sub
