VERSION 5.00
Object = "{AE19B085-7851-4724-8240-EC49EA45E455}#3.0#0"; "pbmasone1.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export SID"
   ClientHeight    =   5970
   ClientLeft      =   150
   ClientTop       =   240
   ClientWidth     =   7950
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin PBMasOne.MasOnePB ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   5400
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   450
   End
   Begin VB.Frame Frame1 
      Caption         =   "Update DIN "
      Height          =   2295
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton Command2 
         Caption         =   "Update DIN"
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Lakukan Proses Ini Bila debitur sudah mendapat DIN dari BI !!!!"
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   3375
      End
   End
   Begin VB.TextBox txtstatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2640
      Width           =   7575
   End
   Begin VB.CommandButton command1 
      Caption         =   "Request Din"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Caption         =   "Proses Export "
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton Command3 
         Caption         =   "Update Status Permintaan din"
         Height          =   495
         Left            =   480
         TabIndex        =   8
         ToolTipText     =   "Tekan Tombol ini bila data berhasil di import di SID"
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Permintaan DIN Baru"
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Label lblstatus 
      Caption         =   "status"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4680
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public Function data_standar()
    skrg = Format(Now, "yyyymmddhhnnss")
    
    strid = "select * from idbpr"
    Set rsid = New ADODB.Recordset
    Set rsid = dbbank.Execute(strid)
    
    idbank = Left(rsid!sandibpr, 6)
    idkancab = Right(rsid!sandibpr, 3)
    
    strtgl = "select settanggal from settanggal"
    Set rstgl = New ADODB.Recordset
    Set rstgl = dbbank.Execute(strtgl)
    
    bln = Format(rstgl!settanggal, "mm")
    thn = Format(rstgl!settanggal, "yyyy")
    
    
End Function

Private Sub Command1_Click()
    
   req_din (baru)

    
End Sub

Private Sub Command2_Click()
    Dim updatedin As ADODB.Recordset
    Dim rssa As ADODB.Recordset
    Dim rsansbi As ADODB.Recordset
    Dim rsdebbi As ADODB.Recordset
    Dim rsdebans As ADODB.Recordset
    
    
    
    If statsid = 0 Then
        MsgBox "Koneksi Dengan Database SID Gagal, Cek Database SID ", vbCritical + vbOKOnly, "Info"
        Exit Sub
    End If
    
    
    sqlupdatedin = "select * from T_DIN order by CIF_BANK"
    Set updatedin = New ADODB.Recordset
    Set updatedin = dbsid.Execute(sqlupdatedin)
    

    ProgressBar1.Value = 0
    ProgressBar1.Max = updatedin.RecordCount + 1
    
    While Not updatedin.EOF
    
        stransbi = "select idnama, norek from rekredit where idnama = '" & updatedin!cif_bank & "' and dinrequest <> 2"
        Set rsansbi = New ADODB.Recordset
        Set rsansbi = dbbank.Execute(stransbi)
        
        If Not rsansbi.EOF Then
        
            strupdatedapok = "update datapokok set din = '" & updatedin!din & "', alias = '" & updatedin!NAMA_ALIAS & "', alamat = '" & updatedin!ALAMAT_DEBITUR & "', kode_pos = '" & updatedin!KODE_POS & "', kelurahan = '" & updatedin!kelurahan & "', kecamatan = '" & updatedin!kecamatan & "', idlokasi = '" & updatedin!DATI2_DEBITUR & "' where idnama = '" & updatedin!cif_bank & "'"
            dbbank.Execute (strupdatedapok)
            
            strupdaterekredit = "update rekredit set din='" & updatedin!din & "', dinrequest = 2, statusSID = 0 where idnama = '" & updatedin!cif_bank & "'"
            dbbank.Execute (strupdaterekredit)
            
            txtstatus.Text = txtstatus.Text & "update cif :" & rsansbi!idnama & "No Rekening :" & rsansbi!norek & vbNewLine
            
             updatedin.MoveNext
            
        Else
        
            ProgressBar1 = ProgressBar1.Value + 1
           
            
            
        End If
        
           
              
    updatedin.MoveNext
    Wend
    
    
       
    MsgBox ("Update data DIN database ANS selesai")
            
End Sub

Private Sub dataAgunan_Click()
    data_agunan.Show 1
End Sub

Private Sub exportdata_Click()
    frmDataSid.Show 1
End Sub

Private Sub Command3_Click()
    
    Dim dinrequestsid As ADODB.Recordset
    
    
    strdinrequestsid = "select cif_bank from t_din_request"
    Set dinrequestsid = New ADODB.Recordset
    Set dinrequestsid = dbsid.Execute(strdinrequestsid)
    
    ProgressBar1.Max = dinrequestsid.RecordCount
    ProgressBar1.Value = 0
    
    While Not dinrequestsid.EOF
        DoEvents
        ProgressBar1.Value = ProgressBar1.Value + 1
        dbbank.Execute ("update datapokok set dinrequest = 1 where idnama = '" & FormatText(dinrequestsid!cif_bank) & "'")
    dinrequestsid.MoveNext
    Wend
    
    MsgBox ("selesai")
    
End Sub

Private Sub Form_Load()
        
    
        
    
    fileini = "SIDANS.ini"
    idfile = FindFirstFile("C:\WINDOWS\SYSTEM\" & fileini, finfo)
    If idfile = -1 Then
       MsgBox "File konfigurasi tidak ditemukan", vbInformation, "File Config Tidak Ada !"
       Konfigurasi.Show vbModal
    Else
       'path_database = ReadINI(FILEini, "DATABASE PATH", "C:\WINDOWS\SYSTEM\" & FILEini)\
        server = ReadIniValue("C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "server")
        user = ReadIniValue("C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "user")
        port = ReadIniValue("C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "port")
        driver = ReadIniValue("C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "driver")
        siduser = ReadIniValue("C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "siduser")
        dbs = ReadIniValue("C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "database")
        dbpwd = ReadIniValue("C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "password")
        usersql = ReadIniValue("C:\WINDOWS\SYSTEM\SIDANS.ini", "dbsid", "userdbsid")
        pwdsql = ReadIniValue("C:\WINDOWS\SYSTEM\SIDANS.ini", "dbsid", "dbsidpwd")
        
        Call buka_koneksi
        Call data_standar
        Call bukasid
    End If
    
    
    'fulll screen
    
End Sub

Public Function req_din(aksi As String)

    Dim fileBaru
    Dim fso
    Dim strslash As String
    Dim bukaFile As String
    Dim jmlrec As Integer
    Dim X As Integer
    Dim rssandi As ADODB.Recordset
    Dim sandibpr As String
    Dim rs As ADODB.Recordset
    Dim isi As String
    Dim rec As Integer
    Dim ttl_record As String
    Dim ket_status As String
    Dim status As String
    Dim npwp As String
    Dim rsusersid As ADODB.Recordset
    Dim nama As String
    Dim namadeb As String
    Dim rec1 As Long
    Dim rsdatabank As ADODB.Recordset
    
    
    strdatabank = "select * from T_INFO_BANK"
    Set rsdatabank = New ADODB.Recordset
    Set rsdatabank = dbsid.Execute(strdatabank)
    
    
    idlembaga = rsdatabank!ID_LEMBAGA
    idbank = rsdatabank!id_bank
    idktrcab = rsdatabank!ID_KTR_CABANG
    bln = rsdatabank!bulan
    thn = rsdatabank!TAHUN
    
    s = "select * from T_SYS_USERS where userid = '" & siduser & "'"
    Set rsusersid = New ADODB.Recordset
    Set rsusersid = dbsid.Execute(s)
    
    namalengkap = rsusersid!namalengkap
    namalengkap = namalengkap & String(20 - Len(namalengkap), " ")
    rec1 = 0
    
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Set fileBaru = fso.CreateTextFile("create-din.txt", True, False)

    If aksi = "perbarui" Then
        id_data_isi = "DEB050"
        operation = "U"
    Else
        operation = "C"
        id_data_isi = "DEB010"
    End If
    
    sqltgl = "select settanggal from settanggal"
    Set rstgl = New ADODB.Recordset
    Set rstgl = dbbank.Execute(sqltgl)
        
        
        
    jns_data = Format(Now, "yyyymm") & "SIDA100"
    id_data = "DEB000"
    ID_LEMBAGA = "002"
    id_bank = "600023"
    id_ktr_cab = "001"
    create_date = Format(Now, "yyyymmddhhnnss")
    'create_user = txtnmuser.Text
    
    regist = ""
    tglbuat = Format(Now, "yyyymmddhhnnss")
    
    
    
    
    
    strslash = IIf(Right$(App.Path, 1) = "\", "", "\")
    bukaFile = App.Path & strBackSlash & "\export\" & Format(Now, "ddmmyy") & "create-din.txt"
    jmlrec = FreeFile
    
    Strsql = "select sandibpr from idbpr"
    Set rssandi = New ADODB.Recordset
    Set rssandi = dbbank.Execute(Strsql)
    sandibpr = rssandi!sandibpr
    
    
    
    strdata = "SELECT datapokok.nama, datapokok.alias, datapokok.alamat, datapokok.kode_pos, datapokok.kelurahan, datapokok.kecamatan, datapokok.IDLOKASI, datapokok.status, datapokok.KET_STATUS, datapokok.NPWP_DEBITUR, datapokok.BUKTIDIRI, datapokok.T_LAHIR, datapokok.TGL_LAHIR, datapokok.IBU_DEBITUR,datapokok.KELAMIN, datapokok.NO_PASPOR, rekredit.JNS_DEBITUR, datapokok.idnama, rekredit.TGLMASUK FROM datapokok, rekredit WHERE datapokok.idnama = rekredit.idnama AND datapokok.dinrequest = '0' AND rekredit.din = '-' AND rekredit.saldoenc <> 0 LIMIT 100"

    Set rs = New ADODB.Recordset
    Set rs = dbbank.Execute(strdata)
    
    ttl_record = rs.RecordCount
    ttl_record = String(10 - Len(ttl_record), "0") & ttl_record
    
    txtstatus.Text = "Export DIN request dari Database " & vbNewLine
    txtstatus.Text = txtstatus.Text & "File " & bukaFile & vbNewLine
    txtstatus.Text = txtstatus.Text & String(80, "=") & vbNewLine
    
    
    
    If ttl_record = 0 Then
    
        MsgBox ("Tidak Ada data baru"), vbInformation, "info"
        
        Exit Function
        
    End If
        
    
    header = jns_data & id_data & idlembaga & idbank & idktrcab & tglbuat & namalengkap & ttl_record & String(34, " ")
    footer = jns_data & "DEB090" & tglbuat & String(7, "0") & "100"
    create_user = namalengkap
    
    no = 1
    
    Open bukaFile For Output As #1
    
        Print #1, header
    
        rec = 0
        
        ProgressBar1.Value = 0
        ProgressBar1.Max = rs.RecordCount
        
        While Not rs.EOF
            DoEvents
            ProgressBar1.Value = ProgressBar1.Value + 1
            rec1 = rec1 + 1
            
            If rec1 = 100 Then
                nasa = cek
            End If
            
            
            namadeb = Replace(rs!nama, vbCrLf, "")
            namadeb = Replace(namadeb, ".", "")
            namadeb = Trim(namadeb)
            namadeb = bersihString(namadeb)
            namadeb = namadeb & String(100 - Len(namadeb), " ")
            
            'namadeb = GenerateRandomString(100)
            If IsNull(rs!alias) Then
                alias = String(50, " ")
            Else
                alias = rs!alias & String(50 - Len(rs!alias), " ")
            End If
            
            alamat = Trim(rs!alamat)
            
            alamat = Replace(alamat, vbCrLf, "")
            alamat = bersihString(alamat)
            If Len(alamat) > 100 Then
            
                alamat = Left(alamat, 100)
            Else
                 alamat = alamat & String(100 - Len(alamat), " ")
            End If
            
            If IsNull(rs!KODE_POS) Or rs!KODE_POS = "" Or rs!KODE_POS = "-" Then
                kopos = String(5, " ")
            Else
                kopos = Replace(rs!KODE_POS, vbCrLf, "") & String(5 - Len(Replace(rs!KODE_POS, vbCrLf, "")), " ")
            End If
            
            
            kelurahan = Replace((IIf(IsNull(rs!kelurahan), "kelurahan", rs!kelurahan)), "  ", " ")
            kelurahan = bersihString(kelurahan)
            If kelurahan = "" Or kelurahan = "-" Then
                kelurahan = String(50, " ")
            Else
                kelurahan = kelurahan & String(50 - Len(kelurahan), " ")
            End If
            
            If (IsNull(rs!kecamatan)) Or rs!kecamatan = "-" Then
                kec = " "
            Else
                kec = rs!kecamatan
            End If
            
            kec = Replace(kec, "  ", " ")
            kec = bersihString(kec)
            kec = kec & String(50 - Len(kec), " ")
            
            dati2 = Left(rs!idlokasi, 4)
                        
            If IsNull(rs!status) Or rs!status = "-" Then
                status = "0100"
            Else
                status = rs!status
                status = status & String(4 - Len(status), " ")
            End If
            
            
            If IsNull(rs!ket_status) Or rs!ket_status = "" Or rs!ket_status = "-" Then
                ket_status = " "
               
            Else
                ket_status = rs!ket_status
               
            End If
            ket_status = bersihString(ket_status)
            ket_status = ket_status & String(50 - Len(ket_status), " ")
            
'            npwp = IIf(IsNull(rs!npwp_debitur), String(20, " "), rs!npwp_debitur)
            
            If rs!npwp_debitur = "" Or rs!npwp_debitur = "-" Or rs!npwp_debitur = "00.000.000.0.000.000" Then
                npwp = String(20, " ")
            ElseIf IsNull(rs!npwp_debitur) Then
                npwp = String(20, " ")
            Else
                npwp = rs!npwp_debitur
                
            End If
            
            npwp = String(20 - Len(npwp), " ") & npwp
            
                        
            If rs!buktidiri = "" Or rs!buktidiri = "-" Or IsNull(rs!buktidiri) Then
                no_ktp = String(30, " ")
            Else
                no_ktp = rs!buktidiri & String(30 - Len(rs!buktidiri), " ")
            End If
            
            t_lahir = Replace(rs!t_lahir, ".", "")
            t_lahir = t_lahir & String(50 - Len(t_lahir), " ")
            
            If rs!tgl_lahir = "" Or IsNull(rs!tgl_lahir) Then
                tgl_lahir = String(8, " ")
            Else
                tgl_lahir = Format(rs!tgl_lahir, "yyyymmdd")
                tgl_lahir = tgl_lahir & String(8 - Len(tgl_lahir), " ")
            End If
            
            If IsNull(rs!IBU_DEBITUR) Then
                MsgBox ("Nama Ibu debitur dari debitur " & rs!nama & " belum ada silahkan di lengkapi dulu")
                Close #1
                Exit Function
            End If

            ibu = Replace(rs!IBU_DEBITUR, ".", "")
            ibu = bersihString(ibu)
            ibu = ibu & String(50 - Len(ibu), " ")
            kelamin = IIf(rs!kelamin = "Laki", "1", "2")
            
            If IsNull(rs!no_paspor) Or rs!no_paspor = "-" Then
                paspor = String(30, " ")
            ElseIf rs!no_paspor = "" Then
                paspor = String(30, " ")
            Else
                paspor = rs!no_paspor
            End If
            
          
            paspor = paspor & String(30 - Len(paspor), " ")
            jenis = IIf(IsNull(rs!JNS_debitur) Or rs!JNS_debitur = "-" Or rs!JNS_debitur = "", "0", rs!JNS_debitur)
            
            cif = rs!idnama & String(30 - Len(rs!idnama), " ")
            register = String(34, " ")
            regserver = String(32, " ")
            mirip = String(3, " ")
            nm_deb_req = String(100, " ")
            din = String(20, " ")
                          
            If rs!idnama = "Sya018" Then
                nasa = cek
            End If

            isi = id_data_isi & id_bank & id_ktr_cab & din & namadeb & alias & alamat & kopos & kelurahan & kec & dati2 & status & ket_status & String(50 - Len(ket_status), " ") & npwp & no_ktp & Replace(t_lahir, ".", "") & tgl_lahir & ibu & kelamin & paspor & jenis & operation & cif & register & tglbuat & regserver & mirip & nm_deb_req & create_user
            
            
           
            
'            strupdate = "update rekredit set dinrequest = 1 where idnama = '" & FormatText(rs!idnama) & "'"
'            dbbank.Execute (strupdate)
            
            Print #1, isi
            
                txtstatus.Refresh
                txtstatus.Text = txtstatus.Text & rec1 & ". " & "Nama : " & namadeb & vbNewLine
             
             
            
        rs.MoveNext
        Wend
    
        Print #1, footer
        'txtstatus.Text = "Jmlh data = " & rec
    
    Close #1
    
    MsgBox ("Request DIN Success!!"), vbInformation, "Selesai"
    

    
End Function


Private Function ktrl_lbbpr(aksi As String)

    Dim rsid As ADODB.Recordset
    'Dim idbank As String
    Dim idkancab As String
    Dim rstgl As ADODB.Recordset
    Dim rsttlkred As ADODB.Recordset
    
    strslash = IIf(Right$(App.Path, 1) = "\", "", "\")
    bukaFile = App.Path & strBackSlash & "form-05-data-kontrol.txt"
    
    skrg = Format(Now, "yyyymmddhhnnss")
    
    strid = "select * from idbpr"
    Set rsid = New ADODB.Recordset
    Set rsid = dbbank.Execute(strid)
    
    strtgl = "select settanggal from settanggal"
    Set rstgl = New ADODB.Recordset
    Set rstgl = dbbank.Execute(strtgl)
    
    strttl = "select sum(plafon) as ttlkred from rekredit"
    Set rsttlkred = New ADODB.Recordset
    Set rsttlkred = dbbank.Execute(strttl)
    
    
    
    iddata = "SID050"
    
    If aksi = "Create" Then
        operation = "C"
    ElseIf aksi = "Update" Then
        operation = "U"
    ElseIf aksi = "Delete" Then
        operation = "D"
    Else
        operation = "S"
    End If
    
    
    'wajib d isi
    idlembaga = " "
    idbank = Left(rsid!sandibpr, 6)
    idkancab = Right(rsid!sandibpr, 3)
    bulan = Format(rstgl!settanggal, "mm")
    TAHUN = Format(rstgl!settanggal, "yyyy")
    penempatanbank = " "
    penempatanbank = penempatanbank & String(18 - Len(penempatanbank), " ")
    srtberharga = String(18, " ")
    kreDiberikan = rsttlkred!ttlkred
    kreDiberikan = kredDiberikan & String(18 - Len(kreDiberikan), " ")
    tagihanlain = String(18, " ")
    penyertaan = String(18, " ")
    irrevocable = String(18, " ")
    garansi = String(18, " ")
    penerusankredit = " " ' 18 char
    createdate = skrg
    createuser = siduser
    updatedate = skrg
    statuskirim = " "
    versi = String(8, " ")
    filler = String(1, " ")
    
    isi = iddata & operation & idLembga & idbank & idkancab & bulan & TAHUN & penempatanbank & srtberharga & kreDiberikan & tagihanlain & penyertaan & _
    irrevocable & garansi & penerusankredit & createdate & createuser & updatedate & statuskirim & versi & filler
    
    Open bukaFile For Output As #1
    
        Print #1, isi
    
    Close #1
    
End Function

Private Sub Form_Unload(Cancel As Integer)
'    Dim i As Long
'    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
'    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
'
'    For i = Me.Left To (Screen.Width / 2) Step 10
'    Me.Height = Me.Height - 15
'    Me.Width = Me.Width - 20
'    Me.Left = Me.Left + 100
'    DoEvents
'    Next
    Unload Me
End Sub

Private Sub ktrlBpr_Click()
    
    ktrl_lbbpr (Create)
    
End Sub



Private Function keuDeb(aksi As String)

    Dim rsid As ADODB.Recordset
    'Dim idbank As String
    'Dim idkancab As String
    Dim rstgl As ADODB.Recordset
    Dim rsttlkred As ADODB.Recordset
    
    strslash = IIf(Right$(App.Path, 1) = "\", "", "\")
    bukaFile = App.Path & strBackSlash & "form-05-data-kontrol.txt"
    
    skrg = Format(Now, "yyyymmddhhnnss")
    
    strid = "select * from idbpr"
    Set rsid = New ADODB.Recordset
    Set rsid = dbbank.Execute(strid)
    
    strtgl = "select settanggal from settanggal"
    Set rstgl = New ADODB.Recordset
    Set rstgl = dbbank.Execute(strtgl)

    iddata = "SID060"
    
    If aksi = "Create" Then
        operation = "C"
        iddeb = String(43, " ")
    ElseIf aksi = "Update" Then
        operation = "U"
    ElseIf aksi = "Delete" Then
        operation = "D"
    Else
        operation = "S"
    End If
    
    'id lembaga
    idlembaga = String(3, " ")
    idbank = Left(rsid!sandibpr, 6)
    idkancab = Right(rsid!sandibpr, 3)
    bln = Format(rstgl!settanggal, "mm")
    thn = Format(rstgl!settanggal, "yyyy")
    
    
    
    
    

End Function


Private Function agunan(aksi As String)

    iddata = "SID041"
    
    If aksi = "Update" Then
        operation = "U"
    ElseIf aksi = "Create" Then
        operation = "C"
    ElseIf aksi = "Delete" Then
        operation = "D"
    Else
        operation = "S"
    End If
    
        
    
    

End Function

Private Sub rekbank_Click()
    frmrekbank.Show 1
End Sub

Private Sub setting_Click()
    Konfigurasi.Show 1
      
End Sub


Private Function dataSID(aksi As String)
    
    Dim ttlDataDebitur As Long
    
    
    id_data_debitur = "SID010"
    id_data_kredit = "SID030"
    
    qdatadebitur = "SELECT rekredit.DIN, datapokok.nama, datapokok.alias, datapokok.status, datapokok.KET_STATUS, rekredit.GOLDEB, datapokok.KELAMIN, datapokok.BUKTIDIRI, datapokok.no_paspor, datapokok.NO_AKTE_AWAL, datapokok.T_LAHIR, datapokok.TGL_LAHIR, datapokok.TGL_AKTE_AWAL, datapokok.NPWP_DEBITUR, datapokok.ALAMAT, datapokok.IDLOKASI, datapokok.KODE_POS, datapokok.KELURAHAN, datapokok.KECAMATAN, datapokok.KODE_AREA, datapokok.TELP, datapokok.NEGARA_DOMISILI, datapokok.IBU_DEBITUR, datapokok.SANDI_PEKERJAAN, datapokok.TEMPAT_BEKERJA, datapokok.BIDANG_USAHA, datapokok.GIN, datapokok.HUB_DGN_BANK, datapokok.LANGGAR_BMPK, datapokok.LAMPAU_BMPK, datapokok.RATING_DEBITUR, datapokok.LEMBAGA_RATING, datapokok.GO_PUBLIC FROM datapokok, rekredit WHERE datapokok.idnama = rekredit.idnama AND NOT ISNULL(rekredit.din) AND LENGTH(rekredit.din) > 10 AND rekredit.DINREQUEST <> '1'"
    Set rsdataDebitur = New ADODB.Recordset
    Set rsdataDebitur = dbbank.Execute(qdatadebitur)
    
    ttlDataDebitur = rsdataDebitur.RecordCount
    
    qdatakredit = ""
    
    
    
    jenis_fasilitas = "0605"
    id_fasilitas = String(52, " ")
    sifat = ""
    
    If aksi = "Create" Then
        operation = "C"
    ElseIf aksi = "Update" Then
        operation = "U"
    ElseIf aksi = "Delete" Then
        operation = "D"
    Else
        operation = "S"
    End If
    
    

    
    
End Function



Private Sub sinkronisasi_Click()
    
    sinc.Show 1
    
End Sub

Function GenerateRandomString(ByRef length As Integer) As String
    Randomize
    Dim allowableChars As String
    allowableChars = "abcdefghijklmnopqrstuvwxyz0123456789"

    Dim i As Integer
    For i = 1 To length
        GenerateRandomString = GenerateRandomString & Mid$(allowableChars, Int(Rnd() * Len(allowableChars) + 1), 1)
    Next
End Function

Public Function bersihString(ByVal isiText As String) As String
    isiText = Replace(isiText, vbCrLf, "")
    isiText = Replace(isiText, vbCr, "")
    isiText = Replace(isiText, vbLf, "")
    isiText = Trim(isiText)
    bersihString = isiText
End Function
