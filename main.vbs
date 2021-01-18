Function Include(vbsFile)
    Dim fso , f ,s
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f   = fso.OpenTextFile(vbsFile)
    s       = f.ReadAll()
    f.close
    ExecuteGlobal s
End Function
set fso = CreateObject("Scripting.FileSystemObject")
Dirfile = fso.GetParentFolderName(WScript.ScriptFullName)

Include Dirfile & "\koneksi.vbs"
' Include Dirfile & "\common.vbs"
wscript.echo ""
wscript.echo "=========================== DATA CI_SCRIP ========================================"
wscript.echo ""

Dim result_ins_scrip
Dim result_del_scrip

' Delete CI_SCRIP 
COUNT_DEL_CI_SCRIP = "SELECT COUNT(*) FROM CI_SCRIP WHERE SNAP_DAT='02-OCT-20'"
set result_del_scrip = oConDW.Execute(COUNT_DEL_CI_SCRIP)
wscript.echo "Total Row Delete : " & result_del_scrip(0)

DEL_CI_SCRIP = "DELETE FROM CI_SCRIP WHERE SNAP_DAT='02-OCT-20'"
' wscript.echo DEL_CI_SCRIP
oConDW.Execute(DEL_CI_SCRIP)
wscript.echo "Data CI_SCRIP berhasil Di Hapus"
wscript.echo "--------------------------"

' INSERT CI_SCRIP
INSERT_CI_SCRIP = "INSERT INTO CI_SCRIP " &_
"select " &_
"code_base_sec, sec_dsc,SEC_MODAL_DASAR sec_jml_dasar, blnc_lokal_scrip, blnc_asing_scrip," &_
"acct_lokal_scrip, acct_asing_scrip, blnc_lokal_scripless, blnc_asing_scripless," &_
"acct_lokal_scripless, acct_asing_scripless, SNAP_DAT from CI_SCRIP_SCRIPLESS_BCKP WHERE SNAP_DAT = '02-OCT-20'"
COUNT_INS_CI_SCRIP = "SELECT COUNT(*) FROM CI_SCRIP WHERE SNAP_DAT='02-OCT-20'"

oConDW.Execute(INSERT_CI_SCRIP)
set result_ins_scrip = oConDW.Execute(COUNT_INS_CI_SCRIP)
wscript.echo "Total Row Insert : " & result_ins_scrip(0)
wscript.echo "Data CI_SCRIP berhasil Di Tambah"

wscript.echo ""
wscript.echo "=========================== DATA CI_SCRIPLESS ===================================="
wscript.echo ""

Dim result_ins_scripless
Dim result_del_scripless

' Delete CI_SCRIPLESS 
COUNT_DEL_CI_SCRIPLESS = "SELECT COUNT(*) FROM CI_SCRIPLESS WHERE SNAP_DAT='02-OCT-20'"
set result_del_scripless = oConDW.Execute(COUNT_DEL_CI_SCRIPLESS)
wscript.echo "Total Row Delete : " & result_del_scripless(0)

DEL_CI_SCRIPLESS = "DELETE FROM CI_SCRIPLESS WHERE SNAP_DAT='02-OCT-20'"
' wscript.echo DEL_CI_SCRIPLESS
oConDW.Execute(DEL_CI_SCRIPLESS)
wscript.echo "Data CI_SCRIPLESS berhasil Di Hapus"
wscript.echo "--------------------------"

' INSERT CI_SCRIPLESS 
INSERT_CI_SCRIPLESS = "INSERT INTO CI_SCRIPLESS " &_ 
"select " &_
"code_base_sec, sec_dsc, sec_modal_dasar, blnc_lokal_scripless, " &_
"blnc_asing_scripless, acct_lokal_scripless, acct_asing_scripless, " &_
"SNAP_DAT from CI_SCRIP_SCRIPLESS_BCKP WHERE SNAP_DAT = '02-OCT-20'"
oConDW.Execute(INSERT_CI_SCRIPLESS)

COUNT_CI_SCRIPLESS = "SELECT COUNT(*) FROM CI_SCRIPLESS WHERE SNAP_DAT='02-OCT-20'"
set result_ins_scripless = oConDW.Execute(COUNT_CI_SCRIPLESS)
wscript.echo "Total Row Insert : " & result_ins_scripless(0)
wscript.echo "Data CI_SCRIPLESS berhasil Di Tambah"

wscript.echo ""
wscript.echo "=========================== DATA CI_SCRIP_SCRIPLESS =============================="
wscript.echo ""
Dim result_ins_scrip_scripless
Dim result_del_scrip_scripless

' Delete CI_SCRIP_SCRIPLESS 
COUNT_DEL_CI_SCRIP_SCRIPLESS = "SELECT COUNT(*) FROM CI_SCRIP_SCRIPLESS WHERE SNAP_DAT='02-OCT-20'"
set result_del_scrip_scripless = oConDW.Execute(COUNT_DEL_CI_SCRIP_SCRIPLESS)
wscript.echo "Total Row Delete : " & result_del_scrip_scripless(0)

DEL_CI_SCRIP_SCRIPLESS = "DELETE FROM CI_SCRIP_SCRIPLESS WHERE SNAP_DAT='02-OCT-20'"
' wscript.echo DEL_CI_SCRIP_SCRIPLESS
oConDW.Execute(DEL_CI_SCRIP_SCRIPLESS)
wscript.echo "Data CI_SCRIP_SCRIPLESS berhasil Di Hapus"
wscript.echo "--------------------------"

' INSERT CI_SCRIP_SCRIPLESS 
INSERT_CI_SCRIP_SCRIPLESS = "INSERT INTO CI_SCRIP_SCRIPLESS " &_
"select " &_
"SEC_ISSUER_ID, SEC_ISSUER, CLO_PRI, SEKTOR, SUB_SEKTOR, " &_
"code_base_sec, sec_dsc, sec_modal_dasar, blnc_lokal_scrip, " &_
"blnc_asing_scrip, total_blnc_scrip, acct_lokal_scrip, acct_asing_scrip, " &_
"total_acct_scrip, blnc_lokal_scripless, blnc_asing_scripless, " &_
"total_blnc_scripless, acct_lokal_scripless, acct_asing_scripless, " &_
"total_acct_scripless, SNAP_DAT from CI_SCRIP_SCRIPLESS_BCKP WHERE SNAP_DAT = '02-OCT-20'"

On Error Resume Next
    INSERT_EX = oConDW.Execute(INSERT_CI_SCRIP_SCRIPLESS) 'Insert Ke table CI_SCRIP_SCRIPLESS
If Err.Number = 0 Then
    
    'JIKA INSERT BERHASIL
    WScript.Echo "Notice : Query Insert Sukses"
    date_now    = now()
    START_END   = Replace(date_now,"/","-")

    lst_upd     = Replace(date_now,"/","")
    lst_upd1    = Replace(lst_upd,":","")
    LST_UPD_TS  = Replace(lst_upd1," ","")

    LOG_SUKSES = "INSERT INTO LOG_SCRIP_SCRIPLESS VALUES('"& START_END &"','"& START_END &"','INSERT SCRIP_SCRIPLESS','INSERT BERHASIL','SUCCESS','0',NULL,'"& LST_UPD_TS &"')"
    oConDW.Execute(LOG_SUKSES)
    wscript.echo "Notice : Insert LOG SUKSES"
    wscript.echo "Notice : Data berhasil di simpan ke Table CI_SCRIP_SCRIPLESS"
    
    ' ' Create Exel File
    '     Dir_File_Exel = "D:\GIT3\SCRIP_SCRIPLESS_"& LST_UPD_TS &"_.xlsx"
    '     Set objExcel = CreateObject("Excel.Application") 
    '     objExcel.Visible = True 
    '     Set objWorkbook = objExcel.Workbooks.Add 
    '     objExcel.Cells(1,1).Value = "SEC_ISSUER_ID"
    '     objExcel.Cells(1,2).Value = "SEC_ISSUER"
    '     objExcel.Cells(1,3).Value = "CLO_PRI"
    '     objExcel.Cells(1,4).Value = "SEKTOR"
    '     objExcel.Cells(1,5).Value = "SUB_SEKTOR"
    '     objExcel.Cells(1,6).Value = "CODE_BASE_SEC"
    '     objExcel.Cells(1,7).Value = "SEC_DSC"
    '     objExcel.Cells(1,8).Value = "SEC_MODAL_DASAR"
    '     objExcel.Cells(1,9).Value = "BLNC_LOKAL_SCRIP"
    '     objExcel.Cells(1,10).Value = "BLNC_ASING_SCRIP"
    '     objExcel.Cells(1,11).Value = "TOTAL_BLNC_SCRIP"
    '     objExcel.Cells(1,12).Value = "ACCT_LOKAL_SCRIP"
    '     objExcel.Cells(1,13).Value = "ACCT_ASING_SCRIP"
    '     objExcel.Cells(1,14).Value = "TOTAL_ACCT_SCRIP"

    '     For cell = 1 To 14
    '         objExcel.Cells(1,cell).Font.Bold = True
    '         ' objExcel.Cells(1,cell).Borders(1).Weight = 2
    '         ' objExcel.Cells(1,cell).Borders(2).Weight = 2
    '         ' objExcel.Cells(1,cell).Borders(3).Weight = 2
    '         ' objExcel.Cells(1,cell).Borders(4).Weight = 2
    '     Next
        
    '     Dim RS_CI_SCRIP_SCRIPLESS 
    '     Set RS_CI_SCRIP_SCRIPLESS = o.ConDW.Execute("SELECT * FROM CI_SCRIP_SCRIPLESS")
    '     Dim Countx 
    '     Set Countx = oConDW.Execute("SELECT COUNT(*) FROM CI_SCRIP_SCRIPLESS WHERE SNAP_DAT='02-OCT-20'")
        
    '     For i = 2 To Countx(0)

    '         ' For cells = 1 To 14
    '         '     objExcel.Cells(i,cells).Borders(1).Weight = 2
    '         '     objExcel.Cells(i,cells).Borders(2).Weight = 2
    '         '     objExcel.Cells(i,cells).Borders(3).Weight = 2
    '         '     objExcel.Cells(i,cells).Borders(4).Weight = 2
    '         ' Next
    '         ' For cells = 0 To 14
    '         '     cellsobj = cells + 1
    '         '     objExcel.Cells(i,cellsobj).Value = RS_CI_SCRIP_SCRIPLESS(cells)
    '         ' Next

    '         objExcel.Cells(i,1).Value = RS_CI_SCRIP_SCRIPLESS(0)
    '         objExcel.Cells(i,2).Value = RS_CI_SCRIP_SCRIPLESS(1)
    '         objExcel.Cells(i,3).Value = RS_CI_SCRIP_SCRIPLESS(2)
    '         objExcel.Cells(i,4).Value = RS_CI_SCRIP_SCRIPLESS(3)
    '         objExcel.Cells(i,5).Value = RS_CI_SCRIP_SCRIPLESS(4)
    '         objExcel.Cells(i,6).Value = RS_CI_SCRIP_SCRIPLESS(5)
    '         objExcel.Cells(i,7).Value = RS_CI_SCRIP_SCRIPLESS(6)
    '         objExcel.Cells(i,8).Value = RS_CI_SCRIP_SCRIPLESS(7)
    '         objExcel.Cells(i,9).Value = RS_CI_SCRIP_SCRIPLESS(8)
    '         objExcel.Cells(i,10).Value = RS_CI_SCRIP_SCRIPLESS(9)
    '         objExcel.Cells(i,11).Value = RS_CI_SCRIP_SCRIPLESS(10)
    '         objExcel.Cells(i,12).Value = RS_CI_SCRIP_SCRIPLESS(11)
    '         objExcel.Cells(i,13).Value = RS_CI_SCRIP_SCRIPLESS(12)
    '         objExcel.Cells(i,14).Value = RS_CI_SCRIP_SCRIPLESS(13)

    '         RS_CI_SCRIP_SCRIPLESS.MoveNext
    '         ' wscript.echo RS_CI_SCRIP_SCRIPLESS(0)
    '     Next
    '     objWorkbook.SaveAs Dir_File_Exel
    '     objWorkbook.Close 
    '     objExcel.Quit
    '     Set objExcel = Nothing
    '     Set objWorkbook = Nothing
    ' 'End exel

    'Send Email Berhasil
    Dim email_notif_berhasil
    Set email_notif_berhasil = CreateObject("CDO.Message") 
            
    'This section provides the configuration information for the remote SMTP server.
    email_notif_berhasil.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'Send the message using the network (SMTP over the network).
    email_notif_berhasil.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
    email_notif_berhasil.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465 'or 587
    email_notif_berhasil.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    email_notif_berhasil.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
            
    ' Google apps mail servers require outgoing authentication. Use a valid email address and password registered with Google Apps.
    email_notif_berhasil.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
    email_notif_berhasil.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") ="gunawanprasetyo313@gmail.com" 'your Google apps mailbox address
    email_notif_berhasil.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") ="gunawan12345" 'Google apps password for that mailbox
    email_notif_berhasil.Configuration.Fields.Update
            
    email_notif_berhasil.From    = "gunawanprasetyo313@gmail.com"
    email_notif_berhasil.To      = "reclosher@gmail.com"
    ' email_notif_berhasil.AddAttachment Dir_File_Exel

    ' Single CC
    email_notif_berhasil.Cc      = "gunn@gmail.com"

    ' Multi CC
    ' Arrcc               = Array("gunn1@gmail.com;","gunn2@gmail.com;","alif@gmail.com;")
    ' email_cc            = ""
    ' For each recivement in Arrcc
    '     email_cc = email_cc & recivement
    ' Next
    ' email_notif_berhasil.Cc      = email_cc
            
    'we are sending a text email.. simply switch the comments around to send an html email instead
    email_notif_berhasil.Subject  = "Scrip Scrippless Notifikasi"

    Arrb    = Array("","Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember")
    Tgl     = Date
    Tglspl  = Split(Tgl,"/")
    Tglnow  = Tglspl(0) & " " & Arrb(Tglspl(1)) & " " & Tglspl(2)
    
    email_notif_berhasil.HTMLBody = "<div>" &_
                                        "<p>Kepada <b>Unit Penelitian</b>,</p>" &_ 
                                        "<div>" &_ 
                                            "<span>Dengan ini kami informasikan bahwa pada tanggal </span><b>"& Tglnow &"</b>" & _ 
                                            "<span> Terakait eksekusi procedure untuk data perbandingan script dan scriptless dinyatakan </span>" &_
                                            "<b>BERHASIL</b><span>"&_
                                        "</div><br>" &_
                                        "<div>" &_ 
                                            "<span>Demikian informasi disampaikan,</span><br>" &_
                                            "<span>Terima Kasih.</span>" &_
                                        "</div>" &_
                                    "</div>"
    email_notif_berhasil.Send  
    Set email_notif_berhasil = Nothing 

    Set ObjFSO = CreateObject("Scripting.FileSystemObject")
        Set objLog = objFSO.CreateTextFile("D:\GIT3\LOG_"& LST_UPD_TS &".txt")

        objLog.WriteLine now() & " - " & INSERT_CI_SCRIP_SCRIPLESS
        objLog.WriteLine now() & " - Sukses" & vbCrlf

        objLog.WriteLine now() & " - " & LOG_SUKSES
        objLog.WriteLine now() & " - Sukses" & vbCrlf

        objLog.WriteLine now() & " - " & "Send Email Notifikasi"
        objLog.WriteLine now() & " - " & "Sukses"
        
        objLog.CLose
        
Else
    'JIKA INSERT GAGAL
    WScript.Echo "Err : Query Insert Gagal"
    err_ins     = "Error : " & Err.Number & " Src: " & Err.Source & " Dsc: " &  Err.Description
    date_now    = now()
    START_END   = Replace(date_now,"/","-")

    lst_upd     = Replace(date_now,"/","")
    lst_upd1    = Replace(lst_upd,":","")
    LST_UPD_TS  = Replace(lst_upd1," ","")

    ' Kirim email notifikasi
    On Error Resume Next
        Dim email_notif
        Set email_notif = CreateObject("CDO.Message") 
             
        'This section provides the configuration information for the remote SMTP server.
        email_notif.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'Send the message using the network (SMTP over the network).
        email_notif.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
        email_notif.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465 'or 587
        email_notif.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        email_notif.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
             
        ' Google apps mail servers require outgoing authentication. Use a valid email address and password registered with Google Apps.
        email_notif.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
        email_notif.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") ="gunawanprasetyo313@gmail.com" 'your Google apps mailbox address
        email_notif.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") ="gunawan12345" 'Google apps password for that mailbox
             
        email_notif.Configuration.Fields.Update
             
        email_notif.From    = "gunawanprasetyo313@gmail.com"
        email_notif.To      = "reclosher@gmail.com"

        ' Single CC
        ' email_notif.Cc      = ""

        ' Multi CC
        ' Arrcc               = Array("gunn1@gmail.com;","gunn2@gmail.com;","alif@gmail.com;")
        ' email_cc            = ""
        ' For each recivement in Arrcc
        '     email_cc = email_cc & recivement
        ' Next
        ' email_notif.Cc      = email_cc
             
        'we are sending a text email.. simply switch the comments around to send an html email instead
        email_notif.Subject  = "Scrip Scrippless Notifikasi"

        Arrb    = Array("","Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember")
        Tgl     = Date
        Tglspl  = Split(Tgl,"/")
        Tglnow  = Tglspl(0) & " " & Arrb(Tglspl(1)) & " " & Tglspl(2)
        
        email_notif.HTMLBody = "<div>" &_
                                    "<p>Kepada <b>Unit Penelitian</b>,</p>" &_ 
                                    "<div>" &_ 
                                        "<span>Dengan ini kami informasikan bahwa pada tanggal </span><b>"& Tglnow &"</b>" & _ 
                                        "<span> Terakait eksekusi procedure untuk data perbandingan script dan scriptless dinyatakan </span>" &_
                                        "<b>GAGAL</b><span>, dengan pesan kesalahan sebagai berikut :</span><br>" &_
                                        "<h4>"& err_ins &"</h4>" &_
                                    "</div>" &_
                                    "<div>" &_ 
                                        "<span>Demikian informasi disampaikan,</span><br>" &_
                                        "<span>Terima Kasih.</span>" &_
                                    "</div>" &_
                                "</div>"
        email_notif.Send  
        Set email_notif = Nothing 
    If Err.Number = 0 Then
        ' Jika email sukses dikirim
        email     = "SUKSESS"

    else
        ' Jika email gagal dikirim
        email     = "GAGAL"
        err_email = "Error : " & Err.Number & " Src: " & Err.Source & " Dsc: " &  Err.Description
        
        Dim email_gagal
        Set email_gagal = CreateObject("CDO.Message") 
             
        'This section provides the configuration information for the remote SMTP server.
        email_gagal.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'Send the message using the network (SMTP over the network).
        email_gagal.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
        email_gagal.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465 'or 587
        email_gagal.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        email_gagal.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
             
        ' Google apps mail servers require outgoing authentication. Use a valid email address and password registered with Google Apps.
        email_gagal.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
        email_gagal.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") ="gunawanprasetyo313@gmail.com" 'your Google apps mailbox address
        email_gagal.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") ="gunawan12345" 'Google apps password for that mailbox
             
        email_gagal.Configuration.Fields.Update
             
        email_gagal.From    = "gunawanprasetyo313@gmail.com"
        email_gagal.To      = "reclosher@gmail.com"
             
        'we are sending a text email.. simply switch the comments around to send an html email instead
        'email_gagal.HTMLBody = "this is the body"
        email_gagal.Subject  = "Scrip Scrippless Notifikasi"

        Arrb    = Array("","Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember")
        Tgl     = Date
        Tglspl  = Split(Tgl,"/")
        Tglnow  = Tglspl(0) & " " & Arrb(Tglspl(1)) & " " & Tglspl(2)

        email_gagal.HTMLBody = "<div>" &_  
                                    "<p>Kepada <b>Unit Penelitian</b>,</p>" &_ 
                                    "<div>" &_ 
                                        "<span>Dengan ini kami informasikan bahwa pada tanggal </span><b>"& Tglnow &"</b>" & _ 
                                        "<span> Terakait eksekusi procedure untuk data perbandingan script dan scriptless dinyatakan </span>" &_
                                        "<b>mengalami gangguan</b><span>, dengan pesan kesalahan sebagai berikut :</span><br>" &_
                                        "<h4>"& err_email &"</h4>" &_
                                    "</div>" &_
                                    "<div>" &_ 
                                        "<span>Demikian informasi disampaikan,</span><br>" &_
                                        "<span>Terima Kasih.</span>" &_
                                    "</div>" &_
                                "</div>"        
        email_gagal.Send  
        Set email_gagal = Nothing 

        wscript.echo err_email
        wscript.echo "Akan dikirimkan proses notif email gagal"
    end if
    wscript.echo "Err : Send Email " & email
    
    ' Cek pengiriman email
    if email = "SUKSESS" Then 
        PROC_FAILED_DSC   = err_ins
        EMAIL_FLG         = "1"
        EMAIL_FAILED_DSC  = NULL
        PCT_LOG           = "Pencatatan Log email Suksess"
        
    else 
        PROC_FAILED_DSC   = err_ins
        EMAIL_FLG         = "2"
        EMAIL_FAILED_DSC  = err_email
        PCT_LOG           = "Pencatatan Log email Gagal" 
    end if

    ' Insert Jika Log Gagal
    wscript.echo "--------------------------"
    LOG_GAGAL = "INSERT INTO LOG_SCRIP_SCRIPLESS VALUES('"& START_END &"','"& START_END &"','INSERT SCRIP_SCRIPLESS','INSERT GAGAL','"& PROC_FAILED_DSC &"','"& EMAIL_FLG &"','"& EMAIL_FAILED_DSC &"','"& LST_UPD_TS &"')"
    oConDW.Execute(LOG_GAGAL)
    wscript.echo "Err : Insert LOG GAGAL"
    wscript.echo pct_log
    Err.Clear
    
    ' Pencatatan log
    if email = "SUKSESS" Then
        ' Jika email sukses
        Set ObjFSO = CreateObject("Scripting.FileSystemObject")
        Set objLog = objFSO.CreateTextFile("D:\GIT3\Log_Gagal_01_"& LST_UPD_TS &".txt")

        objLog.WriteLine now() & " - " & INSERT_CI_SCRIP_SCRIPLESS
        objLog.WriteLine now() & " - Gagal"
        objLog.WriteLine now() & " - " & err_ins & vbcrlf

        objLog.WriteLine now() & " - " & LOG_GAGAL
        objLog.WriteLine now() & " - Sukses" & vbcrlf

        objLog.WriteLine now() & " - " & "Send Email Notifikasi"
        objLog.WriteLine now() & " - " & "Sukses"
        
        objLog.CLose
    else 
        ' Jika email gagal
        Set ObjFSO = CreateObject("Scripting.FileSystemObject")
        Set objLog = objFSO.CreateTextFile("D:\GIT3\Log_Gagal_02_"& LST_UPD_TS &".txt")

        objLog.WriteLine now() & " - " & INSERT_CI_SCRIP_SCRIPLESS
        objLog.WriteLine now() & " - Gagal"
        objLog.WriteLine now() & " - " & err_ins & vbcrlf

        objLog.WriteLine now() & " - " & LOG_GAGAL
        objLog.WriteLine now() & " - Sukses" & vbcrlf

        objLog.WriteLine now() & " - " & "Send Email Notifikasi"
        objLog.WriteLine now() & " - " & "Gagal" 
        objLog.WriteLine now() & " - " & err_email & vbcrlf
        
        objLog.CLose
    end if
End If

COUNT_CI_SCRIP_SCRIPLESS = "SELECT COUNT(*) FROM CI_SCRIP_SCRIPLESS WHERE SNAP_DAT='02-OCT-20'"
set result_ins_scrip_scripless = oConDW.Execute(COUNT_CI_SCRIP_SCRIPLESS)
wscript.echo "Total Row Insert : " & result_ins_scrip_scripless(0)

wscript.echo ""
wscript.echo "=========================== DATA LOG_SCRIP_SCRIPLESS =============================="
wscript.echo ""

COUNT_LOG  = "SELECT COUNT(*) FROM LOG_SCRIP_SCRIPLESS"
LOG_SUKSES = "SELECT COUNT(*) FROM LOG_SCRIP_SCRIPLESS WHERE PROC_NAME = 'INSERT BERHASIL'"
LOG_GAGAL  = "SELECT COUNT(*) FROM LOG_SCRIP_SCRIPLESS WHERE PROC_NAME = 'INSERT GAGAL'"

set result_log  = oConDW.Execute(COUNT_LOG)
set result_logs = oConDW.Execute(LOG_SUKSES)
set result_logg = oConDW.Execute(LOG_GAGAL)
' wscript.echo result_log 

' if result_log = true then
'     wscript.echo "Suksess"
' end if

wscript.echo "Log Sukses : " & result_logs(0)
wscript.echo "Log Gagal  : " & result_logg(0)
wscript.echo "Total Log  : " & result_log(0)