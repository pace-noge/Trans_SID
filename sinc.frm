VERSION 5.00
Begin VB.Form sinc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sinkronisasi"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8745
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame syncframe 
      Caption         =   "Sync Database"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      Begin VB.CommandButton cmdsync 
         Caption         =   "Mulai Sinkronisasi"
         Height          =   495
         Left            =   2640
         TabIndex        =   2
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   2415
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   8055
      End
   End
End
Attribute VB_Name = "sinc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdsync_Click()

    Dim updatedin As ADODB.Recordset
    Dim rssa As ADODB.Recordset
    Dim rsansbi As ADODB.Recordset
    Dim rsdebbi As ADODB.Recordset
    Dim rsdebans As ADODB.Recordset
    Dim dataskr As Integer
    Dim dataagunanans As ADODB.Recordset
    Dim dataAgunan As ADODB.Recordset
    Dim rsnorek As ADODB.Recordset
    Dim rsmaksid As ADODB.Recordset
    Dim maksid As String
    Dim jenis_agunan As String
    Dim jenis_pengikatan As String
    Dim penilai_independen As String
    Dim asuransi As String
    Dim paripasu As Long
    Dim tgl_penilaian As String
    Dim pengikatan As String
    Dim NAMA_ALIAS As String
    Dim rsnorekbi As ADODB.Recordset
    Dim rsidbi As ADODB.Recordset
    
    dataskr = 0
    
    If statsid = 0 Then
        MsgBox "Koneksi Dengan Database SID Gagal, Cek Database SID ", vbCritical + vbOKOnly, "Info"
        Exit Sub
    End If
    
    
    'sqlupdatedin = "select t_din.*, t_debitur.id_debitur from t_din, t_debitur where t_din.din = t_debitur.din and t_din.cif_bank is not null and t_din.cif_bank <> ''"
    sqlupdatedin = "select t_debitur.*, t_kredit.*, r_debitur_fasilitas.*  from t_kredit, r_debitur_fasilitas, t_debitur where t_kredit.id_fasilitas = r_debitur_fasilitas.id_fasilitas and r_debitur_fasilitas.id_debitur = t_debitur.id_debitur"

    Set updatedin = New ADODB.Recordset
    Set updatedin = dbsid.Execute(sqlupdatedin)
    
        
    headerstatus = "Sinkronisasi data DIN " & vbNewLine & "Jumlah data DIN : " & updatedin.RecordCount & vbNewLine & String(20, "=") & vbNewLine
    
    While Not updatedin.EOF
        
        DoEvents


        'stransbi = "select idnama, norek from rekredit where idnama = '" & updatedin!cif_bank & "' and dinrequest <> 2"
        stransbi = "select idnama, norek from rekredit where norek = '" & updatedin!no_rekening & "' and dinrequest <> 2"
        
        Set rsansbi = New ADODB.Recordset
        Set rsansbi = dbbank.Execute(stransbi)
        
        dataskr = dataskr + 1
        
        If updatedin!no_rekening = "135.15692" Then
        
        nasa = cek
        End If
        
        If Not rsansbi.EOF Then
            
            If din = "06582042115119000114" Then
                nasa = cek
            End If
            
            'kadang ga update kalo langsung banyak yg di update
        
'            strupdatedapok = "update datapokok set din = '" & updatedin!din & "', nama = '" & FormatText(updatedin!nama_debitur) & "', alias = '" & updatedin!NAMA_ALIAS & "', alamat = '" & updatedin!ALAMAT_DEBITUR & "', kode_pos = '" & updatedin!KODE_POS & "', kelurahan = '" & updatedin!kelurahan & "', kecamatan = '" & updatedin!kecamatan & "', idlokasi = '" & updatedin!DATI2_DEBITUR & "', status = '" & updatedin!status & "', ket_status = '" & updatedin!ket_status & "' where idnama = '" & rsansbi!idnama & "'"
'            dbbank.Execute (strupdatedapok)
            
            idnamas = FormatText(rsansbi!idnama)
            dbbank.Execute ("update datapokok set din = '" & updatedin!din & "' where idnama = '" & idnamas & "'")
            dbbank.Execute ("update datapokok set nama = '" & FormatText(IIf(IsNull(updatedin!nama_debitur), " ", updatedin!nama_debitur)) & "' where idnama = '" & idnamas & "'")
            dbbank.Execute ("update datapokok set alias = '" & FormatText(IIf(IsNull(updatedin!NAMA_ALIAS), " ", updatedin!NAMA_ALIAS)) & "' where idnama = '" & idnamas & "'")
            dbbank.Execute ("update datapokok set alamat = '" & FormatText(updatedin!ALAMAT_DEBITUR) & "' where idnama = '" & idnamas & "'")
            dbbank.Execute ("update datapokok set kode_pos = '" & updatedin!KODE_POS & "' where idnama = '" & idnamas & "'")
            dbbank.Execute ("update datapokok set kelurahan = '" & FormatText(updatedin!kelurahan) & "' where idnama = '" & idnamas & "'")
            dbbank.Execute ("update datapokok set kecamatan = '" & FormatText(updatedin!kecamatan) & "' where idnama = '" & idnamas & "'")
            dbbank.Execute ("update datapokok set idlokasi = '" & updatedin!DATI2_DEBITUR & "' where idnama = '" & idnamas & "'")
            dbbank.Execute ("update datapokok set status = '" & updatedin!status & "' where idnama = '" & idnamas & "'")
            dbbank.Execute ("update datapokok set ket_status = '" & FormatText(IIf(IsNull(updatedin!ket_status), " ", updatedin!ket_status)) & "' where idnama = '" & idnamas & "'")

            'kadang ga update kalo langsung banyak yg di update
'            strupdaterekredit = "update rekredit set din='" & updatedin!din & "', cif_bi = '" & updatedin!ID_DEBITUR & "', dinrequest = 2, statusSID = 0 where idnama = '" & rsansbi!idnama & "'"
'            dbbank.Execute (strupdaterekredit)
            
            dbbank.Execute ("update rekredit set din='" & updatedin!din & "' where idnama = '" & idnamas & "'")
'            dbbank.Execute ("update rekredit set no_akad_awal ='" & updatedin!no_pk_awal & "' where idnama = '" & idnamas & "'")
'            dbbank.Execute ("update rekredit set no_akad_akhir ='" & updatedin!no_pk_akhir & "' where idnama = '" & idnamas & "'")

            
            
            strrsnorekbi = "select * from rekredit where idnama = '" & idnamas & "'"
            Set rsnorekbi = dbbank.Execute(strrsnorekbi)
            
            If Not rsnorekbi.EOF Or Not rsnorekbi.BOF Then
            
                While Not rsnorekbi.EOF
                    
                    stridbi = "select "
                    dbbank.Execute ("update rekredit set cif_bi = '" & updatedin!id_debitur & "' where norek = '" & updatedin!no_rekening & "'")
                
                rsnorekbi.MoveNext
                Wend
                
            End If
            
            'dbbank.Execute ("update rekredit set cif_bi = '" & updatedin!id_debitur & "' where idnama = '" & idnamas & "'")
            dbbank.Execute ("update datapokok set dinrequest = 2 where idnama = '" & idnamas & "'")
            dbbank.Execute ("update rekredit set statusSID = 0 where idnama = '" & idnamas & "'")
            
            
            'txtstatus.Text = txtstatus.Text & "update cif :" & rsansbi!idnama & "No Rekening :" & rsansbi!norek & vbNewLine
           
            Text1.Text = headerstatus & "data yg d proses " & dataskr & "/" & updatedin.RecordCount & vbNewLine
             Text1.Refresh
            rsansbi.MoveNext

        Else



        End If
        
       Text1.Text = headerstatus & "data yg d proses " & dataskr & "/" & updatedin.RecordCount & vbNewLine
       Text1.Refresh
              
    updatedin.MoveNext
    Wend
    
    
    
  
   
    
    strdeb = "select T_DIN.CIF_BANK, T_DEBITUR.* from T_DIN, T_DEBITUR where T_DIN.DIN = T_DEBITUR.DIN order by DIN"
    Set rsdebbi = New ADODB.Recordset
    Set rsdebbi = dbsid.Execute(strdeb)
    
      headerdebitur = "Sinkronisasi Data DEBITUR " & vbNewLine & "Jumlah data Debitur : " & rsdebbi.RecordCount & vbNewLine & String(20, "=") & vbNewLine
    
    dataskr = 0
    
    headerstatus = Text1.Text
    
    While Not rsdebbi.EOF
        
        DoEvents
        
        Text1.Text = headerstatus & vbNewLine & vbNewLine & headerdebitur & "data yg d proses " & dataskr & "/" & rsdebbi.RecordCount & vbNewLine
        Text1.Refresh
        
        strdebans = "select * from rekredit where dinrequest = 2 and din = '" & rsdebbi!din & "' "
        Set rsdebans = New ADODB.Recordset
        Set rsdebans = dbbank.Execute(strdebans)
        
        If Not rsdebans.EOF Then
            
            If Not IsNull(rsdebbi!cif_bank) Or IsEmpty(rsdebbi!cif_bank) Or rsdebbi!cif_bank = "" Then
            
                updatedebitur = "update rekredit set cif_bi = '" & rsdebbi!id_debitur & "' where idnama = '" & FormatText(rsdebbi!cif_bank) & "'"
                dbbank.Execute (updatedebitur)
                
                If IsNull(rsdebbi!NAMA_ALIAS) Then
                    NAMA_ALIAS = " "
                Else
                    NAMA_ALIAS = rsdebbi!NAMA_ALIAS
                End If
                
                dbbank.Execute ("update datapokok set alias = '" & FormatText(NAMA_ALIAS) & "' where idnama = '" & FormatText(rsdebbi!cif_bank) & "'")
                dbbank.Execute ("update datapokok set alamat = '" & FormatText(rsdebbi!ALAMAT_DEBITUR) & "' where idnama = '" & FormatText(rsdebbi!cif_bank) & "'")
                dbbank.Execute ("update datapokok set status = '" & rsdebbi!status & "' where idnama = '" & FormatText(rsdebbi!cif_bank) & "'")
                dbbank.Execute ("update datapokok set t_lahir='" & rsdebbi!TEMPAT_LAHIR & "' where idnama = '" & FormatText(rsdebbi!cif_bank) & "'")
                dbbank.Execute ("update datapokok set tgl_lahir = '" & Format(rsdebbi!TGL_AKTE_AWAL, "yyyy-mm-dd") & "' where idnama = '" & FormatText(rsdebbi!cif_bank) & "'")
                dbbank.Execute ("update datapokok set idlokasi = '" & rsdebbi!DATI2_DEBITUR & "' where idnama = '" & FormatText(rsdebbi!cif_bank) & "'")
                dbbank.Execute ("update datapokok set kelurahan = '" & rsdebbi!kelurahan & "' where idnama = '" & FormatText(rsdebbi!cif_bank) & "'")
                dbbank.Execute ("update datapokok set kecamatan = '" & rsdebbi!kecamatan & "' where idnama = '" & FormatText(rsdebbi!cif_bank) & "'")
                dbbank.Execute ("update datapokok set kode_pos = '" & rsdebbi!KODE_POS & "' where idnama = '" & FormatText(rsdebbi!cif_bank) & "'")
                dbbank.Execute ("update datapokok set sandi_pekerjaan = '" & rsdebbi!SANDI_PEKERJAAN & "' where idnama = '" & FormatText(rsdebbi!cif_bank) & "'")
                dbbank.Execute ("update datapokok set tempat_bekerja = '" & FormatText(rsdebbi!TEMPAT_BEKERJA) & "' where idnama = '" & FormatText(rsdebbi!cif_bank) & "'")
                dbbank.Execute ("update datapokok set ibu_debitur = '" & FormatText(rsdebbi!IBU_DEBITUR) & "' where idnama = '" & FormatText(rsdebbi!cif_bank) & "'")
                dbbank.Execute ("update datapokok set din = '" & rsdebbi!din & "' where idnama = '" & FormatText(rsdebbi!cif_bank) & "'")
                dbbank.Execute ("update datapokok set hub_dgn_bank = '" & rsdebbi!HUB_DGN_BANK & "' where idnama = '" & FormatText(rsdebbi!cif_bank) & "'")
                dbbank.Execute ("update datapokok set LANGGAR_BMPK = '" & rsdebbi!LANGGAR_BMPK & "' where idnama = '" & FormatText(rsdebbi!cif_bank) & "'")
                dbbank.Execute ("update datapokok set LAMPAU_BMPK = '" & rsdebbi!LAMPAU_BMPK & "' where idnama = '" & FormatText(rsdebbi!cif_bank) & "'")
                dbbank.Execute ("update datapokok set status = '" & rsdebbi!status & "' where idnama = '" & FormatText(rsdebbi!cif_bank) & "'")
            
            End If
            'dbbank.Execute (updatedatapokok)
            
            dataskr = dataskr + 1
            Text1.Text = headerstatus & vbNewLine & vbNewLine & headerdebitur & "data yg d proses " & dataskr & "/" & rsdebbi.RecordCount & vbNewLine
            Text1.Refresh
            
            
            rsdebbi.MoveNext

        Else
            
            
        
        End If
        
     dataskr = dataskr + 1
    Text1.Text = headerstatus & vbNewLine & vbNewLine & headerdebitur & "data yg d proses " & dataskr & "/" & rsdebbi.RecordCount & vbNewLine
    Text1.Refresh

    If rsdebbi.EOF Or rsdebbi.BOF Then
        nasa = cek
    Else
        rsdebbi.MoveNext
        
    End If
    Wend
    
'    qdataaguanan = "SELECT rekredit.CIF_BI, rekredit.norek, data_agunan.* FROM rekredit, data_agunan WHERE rekredit.norek = data_agunan.norek and rekredit.cif_bi = '" & dataagunan!ID_DEBITUR & ""
    strdataagunanans = "SELECT rekredit.CIF_BI, rekredit.norek, datapokok.nama, rekredit.jaminan, datapokok.IDLOKASI, datapokok.alamat, rekredit.NILAIJAMINAN, rekredit.nilaijaminan1, rekredit.TGLMASUK FROM rekredit, datapokok WHERE datapokok.idnama = rekredit.idnama and jaminan <> 'Asuransi jiwa' ORDER BY norek asc"
    Set datagunanans = New ADODB.Recordset
    Set dataagunanans = dbbank.Execute(strdataagunanans)
    
    strCreateDataAgunan = "CREATE TABLE IF NOT EXISTS data_agunan (ID_AGUNAN VARCHAR(20) NOT NULL, norek VARCHAR(20) DEFAULT '-', jenis_agunan VARCHAR(2) DEFAULT '-', pemilik_agunan VARCHAR(100) DEFAULT '-', `bukti_milik` VARCHAR(255) DEFAULT '-', idlokasi VARCHAR(4) DEFAULT '-', alamat_agunan VARCHAR(100) DEFAULT '-', `nilai_agunan` DOUBLE DEFAULT '0', nilai_agunan_bank DOUBLE DEFAULT '0', nilai_agunan_penilai DOUBLE DEFAULT '0', penilai_independen VARCHAR(50) DEFAULT '0', tgl_penilaian DATE DEFAULT NULL, paripasu DOUBLE DEFAULT '0', asuransi VARCHAR(1) DEFAULT '0', id_agunan_bi VARCHAR(32), pengikatan varchar(32) DEFAULT '-', PRIMARY KEY (ID_AGUNAN))"
    dbbank.Execute (strCreateDataAgunan)
    
    ' SELECT rekredit.CIF_BI, rekredit.norek, data_agunan.* FROM rekredit, data_agunan WHERE rekredit.norek = data_agunan.norek and rekredit.cif_bi = '" & dataagunan!ID_DEBITUR & "'
    
    strkosong = "delete from data_agunan"
    dbbank.Execute (strkosong)
    
    header = "Sinkronisasi data agunan" & vbNewLine & vbNewLine & "====================================" & vbNewLine & "Jumlah Data = " & dataagunanans.RecordCount
    
    i = 0
    
    While Not dataagunanans.EOF
        
        DoEvents
        
        i = i + 1
        
        strdataagunan = "select * from T_AGUNAN where id_debitur = '" & dataagunanans!cif_bi & "'"
        Set dataAgunan = New ADODB.Recordset
        Set dataAgunan = dbsid.Execute(strdataagunan)
        
        If Not dataAgunan.EOF Then
            
'            strnorek = "select norek from rekredit where rekredit.CIF_BI = '" & dataagunan!ID_DEBITUR & "'"
'            Set rsnorek = New ADODB.Recordset
'            Set rsnorek = dbbank.Execute(strnorek)
'
            id_agunan_bi = dataAgunan!ID_Agunan
            jenis_agunan = dataAgunan!jenis_agunan
            jenis_pengikatan = dataAgunan!jenis_pengikatan
            If IsNull(dataAgunan!penilai_independen) Then
                penilai_independen = "-"
            Else
                penilai_independen = dataAgunan!penilai_independen
            End If
            
            
            
            asuransi = dataAgunan!asuransi
            
            If IsNull(dataAgunan!paripasu) Then
                paripasu = "0"
            Else
                paripasu = dataAgunan!paripasu
            End If
            
            If IsNull(dataAgunan!tgl_penilaian) Then
                tgl_penilaian = "1900-01-01"
            Else
                tgl_penilaian = dataAgunan!tgl_penilaian
            End If
        Else
            id_agunan_bi = "-"
            
            If InStr(dataagunanans!jaminan, " BPKB ") <> 0 Then
                jenis_agunan = "02"
                jenis_pengikatan = "03"
                penilai_independen = "-"
                asuransi = "-"
                paripasu = 0
                tgl_penilaian = dataagunanans!tglmasuk
            ElseIf InStr(dataagunanans!jaminan, " SHM ") <> 0 Then
                jenis_agunan = "03"
                jenis_pengikatan = "04"
                penilai_independen = "-"
                asuransi = "-"
                paripasu = 0
                tgl_penilaian = dataagunanans!tglmasuk
            ElseIf InStr(dataagunanans!jaminan, " Barang Dagangan ") <> 0 Then
                jenis_agunan = "04"
                jenis_pengikatan = "99"
                penilai_independen = "-"
                asuransi = "-"
                paripasu = 0
                tgl_penilaian = dataagunanans!tglmasuk
            Else
                jenis_agunan = "06"
                jenis_pengikatan = "99"
                penilai_independen = "-"
                asuransi = "-"
                paripasu = 0
                tgl_penilaian = dataagunanans!tglmasuk
            End If
            
            
            
                
            
        End If
            
            Set rsmaksid = New ADODB.Recordset
            Set rsmaksid = dbbank.Execute("select max(id_agunan) as idmaks from data_agunan")
            
            If IsNull(rsmaksid!idmaks) Then
                idbaru = "000001"
            Else
                maksid = Val(rsmaksid!idmaks) + 1
                idbaru = String(6 - Len(maksid), "0") & Val(maksid)
                End If
            
            
            
            dbbank.Execute ("insert into data_agunan values('" & idbaru & "', '" & dataagunanans!norek & "', '" & FormatText(jenis_agunan) & "', '" & FormatText(dataagunanans!nama) & "', '" & FormatText(dataagunanans!jaminan) & "', '" & dataagunanans!idlokasi & "', '" & FormatText(dataagunanans!alamat) & "', " & dataagunanans!nilaijaminan & ", " & IIf(IsNull(dataagunanans!nilaijaminan1), 0, dataagunanans!nilaijaminan1) & ", " & 0 & ", '" & penilai_independen & "', '" & Format(tgl_penilaian, "yyyy-mm-dd") & "', " & paripasu & ", '" & asuransi & "', '" & id_agunan_bi & "', '" & jenis_pengikatan & "')")
        
        
        
        Text1.Refresh
        Text1.Text = header & vbNewLine
        Text1.Text = "Sync Data Agunan" & vbNewLine & "==============================="
        Text1.Text = "Data yang di proses: " & i & "/" & dataagunanans.RecordCount
    
    dataagunanans.MoveNext
    Wend
    
    dbbank.Execute ("update rekredit set din='-' where isnull(din)")
    dbbank.Execute ("update rekredit set din='-' where din = ''")
    dbbank.Execute ("update rekredit set din='-' where din = '0'")
    dbbank.Execute ("UPDATE datapokok SET alamat = REPLACE(alamat, CHAR(13), '')")
    dbbank.Execute ("UPDATE datapokok SET nama = REPLACE(nama, CHAR(13), '')")
    
    MsgBox ("Sinkronisasi selesai")

End Sub

Private Sub Command1_Click()
    rekbank.Show
End Sub

