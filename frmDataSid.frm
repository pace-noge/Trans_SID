VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{AE19B085-7851-4724-8240-EC49EA45E455}#3.0#0"; "pbmasone1.ocx"
Object = "{239FEE44-9B7E-4746-85AB-019C3C126243}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmDataSid 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Data SID"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10980
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin OsenXPCntrl.OsenXPButton cmdTandai 
      Height          =   375
      Left            =   3120
      TabIndex        =   22
      Top             =   2280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Tandai Semua"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmDataSid.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   7680
      TabIndex        =   17
      Top             =   1200
      Width           =   2775
      Begin OsenXPCntrl.OsenXPButton cmdHapus 
         Height          =   615
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Hapus Agunan dan Jaminan"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmDataSid.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   240
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   5805
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   13
      Top             =   3720
      Width           =   6735
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   7335
      Begin VB.CheckBox Check2 
         Caption         =   "Data Penjamin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Relasi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Kontrol LBPR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Data Agunan (FM-4A)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Kredit yang diberikan (FM-3B)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   240
         Width           =   2655
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Data Penempatan (FM-3A)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Data Debitur (fm-01)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Proses"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   10335
      Begin OsenXPCntrl.OsenXPButton cmdMulai 
         Height          =   375
         Left            =   4320
         TabIndex        =   21
         Top             =   2040
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Start"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmDataSid.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PBMasOne.MasOnePB pgbar 
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Top             =   1560
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   450
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   11
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Nama File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Data yang sedang di proses"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Progres"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Pemrosesan Data :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      Caption         =   "TRANSFER DATA KE SID BI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   16
      Top             =   360
      Width           =   6135
   End
End
Attribute VB_Name = "frmDataSid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public footer As String
Dim sttbtn As Integer
Dim rscekrek As ADODB.Recordset
Dim cek_jml_data As Double
Public nmpengirim As String
Public blndata As String
Public thndata As String
Public bukaFile As String
Dim rscekpenempatan As ADODB.Recordset
Dim strcekpenempatan As String



Private Sub cmdHapus_Click()
    Dim strhapuagunanlunas As String
    Dim rshapusagunan As ADODB.Recordset
    Dim strhapuspenjamin As String
    Dim rshapuspenjamin As ADODB.Recordset
    
    strhapusagunanlunas = "select t_kredit.id_fasilitas, t_debitur.id_debitur, t_agunan.id_agunan, t_debitur.din from t_kredit, t_debitur, t_agunan, r_debitur_fasilitas where t_kredit.id_fasilitas = r_debitur_fasilitas.id_fasilitas and t_debitur.id_debitur = r_debitur_fasilitas.id_debitur and t_agunan.id_debitur = t_debitur.id_debitur and t_kredit.kondisi = '02'"
    Set rshapusagunan = New ADODB.Recordset
    Set rshapusagunan = dbsid.Execute(strhapusagunanlunas)
    'hpusagunan.Enabled = False
    While Not rshapusagunan.EOF
        DoEvents
        
        dbsid.Execute ("delete from t_agunan where id_agunan = '" & rshapusagunan!ID_Agunan & "'")
    
    
    rshapusagunan.MoveNext
    Wend
    
    strhapuspenjamin = "select t_kredit.id_fasilitas, t_debitur.id_debitur, t_penjamin.id_penjamin, t_debitur.din from t_kredit, t_debitur, t_penjamin, r_debitur_fasilitas where t_kredit.id_fasilitas = r_debitur_fasilitas.id_fasilitas and t_debitur.id_debitur = r_debitur_fasilitas.id_debitur and t_penjamin.id_debitur = t_debitur.id_debitur and t_kredit.kondisi = '02'"
    Set rshapuspenjamin = New ADODB.Recordset
    Set rshapuspenjamin = dbsid.Execute(strhapuspenjamin)
    
    'hpusPenjaminlunas.Enabled = False
    
    While Not rshapuspenjamin.EOF
        DoEvents
        
        dbsid.Execute ("delete from t_penjamin where id_penjamin = '" & rshapuspenjamin!Id_Penjamin & "'")
    
    rshapuspenjamin.MoveNext
    Wend
    
    MsgBox ("Data agunan dan jaminan untuk debitur yang telah lunas selesai di hapus")
    cmdHapus.Enabled = True
End Sub

Private Sub cmdMulai_Click()
    cmdMulai.Enabled = False
    createHeader
    cetakData
End Sub

Private Sub cmdTandai_Click()
     
    If sttbtn = 1 Then
        
        Check1 = vbChecked
        Check2 = vbChecked
        Check3 = vbChecked
        Check4 = vbChecked
        Check5 = vbChecked
        Check6 = vbChecked
        Check7 = vbChecked
        Check8 = vbChecked
        Check9 = vbChecked
        Check10 = vbChecked
        cmdTandai.Caption = "Hilangkan Tanda"
        sttbtn = 2
        
    Else
    
        Check1 = Unchecked
        Check2 = Unchecked
        Check3 = Unchecked
        Check4 = Unchecked
        Check5 = Unchecked
        Check6 = Unchecked
        Check7 = Unchecked
        Check8 = Unchecked
        Check9 = Unchecked
        Check10 = Unchecked
        cmdTandai.Caption = "Tandai semua"
        sttbtn = 1
        
    End If
End Sub

'Private Sub Command1_Click()
'
'    If sttbtn = 1 Then
'
'        Check1 = vbChecked
'        Check2 = vbChecked
'        Check3 = vbChecked
'        Check4 = vbChecked
'        Check5 = vbChecked
'        Check6 = vbChecked
'        Check7 = vbChecked
'        Check8 = vbChecked
'        Check9 = vbChecked
'        Check10 = vbChecked
'        Command1.Caption = "Hilangkan Tanda"
'        sttbtn = 2
'
'    Else
'
'        Check1 = Unchecked
'        Check2 = Unchecked
'        Check3 = Unchecked
'        Check4 = Unchecked
'        Check5 = Unchecked
'        Check6 = Unchecked
'        Check7 = Unchecked
'        Check8 = Unchecked
'        Check9 = Unchecked
'        Check10 = Unchecked
'        Command1.Caption = "Tandai semua"
'        sttbtn = 1
'
'    End If
'
'End Sub
Private Function totalRecord(s As String)
    
    Dim rshitung As New ADODB.Recordset
    Dim ttlrecord As String
    
    Set rshitung = dbbank.Execute(s)
    ttlrecord = rshitung.RecordCount
    totalRecord = ttlrecord
    
End Function


Private Function createHeader()
    
    Dim waktuSkrg As String
    Dim StatDeb As Integer
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim ID_Agunan As String
    Dim Wkt_agunan As Date
    Dim Cari_wkt_agunan As Date
    Dim Wkt_agunan_lalu As Date
    Dim Ulg As Double
    Dim StatDebitur As Integer
    Dim Id_fas_deb As String
    Dim rsdataDebitur As ADODB.Recordset
    Dim rsdatapengurus As ADODB.Recordset
    Dim rsdatapenempatan As ADODB.Recordset
    Dim rsDataKredit As ADODB.Recordset
    Dim rsdatapenjamin As ADODB.Recordset
    Dim rsdatakontrol As ADODB.Recordset
    Dim rsAngsuran As ADODB.Recordset
    Dim jmlhBadanUsaha As Integer
    Dim ttldeb As String
    Dim rsdatabank As ADODB.Recordset
    Dim rsdataagunan As ADODB.Recordset
    Dim rsdatakeu As ADODB.Recordset
    Dim rsdatarelasi As ADODB.Recordset
    Dim ttlrecord As String
    Dim nama As String
    Dim plafonInduk As String
    Dim plafon As String
    Dim bakidebet As String
    Dim tunggakanpokok As String
    Dim frektunggakanpkk As String
    Dim tunggBungaIntra As String
    Dim runggBungaEkstra As String
    Dim denda As String
    Dim agunan As String
    Dim ppap As String
    Dim ttlkredit As String
    Dim header As String
    Dim rsusersid As ADODB.Recordset
    Dim ttldatarelasi As String
    Dim noakadawal As String
    Dim statuskredit As String
    Dim frektunggakanbunga As String
    
    
    
    s = "select * from T_SYS_USERS where userid = '" & siduser & "'"
    Set rsusersid = New ADODB.Recordset
    Set rsusersid = dbsid.Execute(s)
    
    namalengkap = rsusersid!namalengkap
    
    strdatabank = "select * from T_INFO_BANK"
    Set rsdatabank = New ADODB.Recordset
    Set rsdatabank = dbsid.Execute(strdatabank)
    
    jenisdata = Format(Now, "yyyymm") & "SIDA100"
    iddata = "SID000"
    idlembaga = rsdatabank!ID_LEMBAGA
    idbank = rsdatabank!id_bank
    idktrcab = rsdatabank!ID_KTR_CABANG
    bln = rsdatabank!bulan
    thn = rsdatabank!TAHUN
    
    nmpengirim = namalengkap & String(30 - Len(namalengkap), " ")
    versiAplikasi = String(4, " ")
    versiref = String(4, " ")
    versival = String(4, " ")
    statusData = String(4, "0")
   
    
    ttlpengurus = String(7 - Len(ttlpengurus), "0") & ttlpengurus
    ttlpenempatan = String(7 - Len(ttlpenempatan), "0") & ttlpenempatan
    ttlsuratberharga = String(7, "0")
    ttlkredit = String(7 - Len(ttlkredit), "0") & ttlkredit
    ttltagihanlainnya = String(7, "0")
    ttlpenyertaan = String(7, "0")
    ttlirrecovable = String(7, "0")
    ttlgaransi = String(7, "0")
    ttlkreditkelolaan = String(7, "0")
    ttlpenjamin = String(7, "0")
    ttlktrllbu = String(7, "0")
    ttldatakeu = String(7 - Len(ttldatakeu), "0") & ttldatakeu
    ttlrelasi = String(7, "0")
    
    namabank = rsdatabank!NAMA_BANK
    namabank = namabank & String(50 - Len(namabank), " ")
    alamatbank = rsdatabank!alamat_bank
    alamatbank = alamatbank & String(50 - Len(alamat_bank), " ")
    kodearea = rsdatabank!kode_area
    notlpbank = rsdatabank!NO_TELEPON
    notlpbank = notlpbank & String(9 - Len(notlpbank), " ")
    sttuskantor = String(1, " ")
    sttusbank = rsdatabank!STATUS_BANK
    
    
    blndata = String(2 - Len(bln), "0") & bln
    thndata = thn
    waktucreate = thndata & blndata & Format(Now, "ddhhnnss")
    
    If Check1 = Checked Then
        ttldeb = totalRecord("SELECT rekredit.din, rekredit.CIF_BI, datapokok.nama, datapokok.alias, datapokok.status, datapokok.ket_status, rekredit.goldeb, datapokok.KELAMIN, datapokok.BUKTIDIRI, datapokok.NO_PASPOR, datapokok.T_LAHIR, datapokok.TGL_LAHIR, datapokok.TGL_AKTE_AWAL, datapokok.NPWP_DEBITUR, datapokok.ALAMAT, datapokok.IDLOKASI, datapokok.KODE_POS, datapokok.KELURAHAN, datapokok.KECAMATAN, datapokok.KODE_AREA, datapokok.telp, datapokok.NEGARA_DOMISILI, datapokok.IBU_DEBITUR, datapokok.SANDI_PEKERJAAN, datapokok.TEMPAT_BEKERJA, datapokok.BIDANG_USAHA, rekredit.sektor, datapokok.HUB_DGN_BANK, datapokok.LANGGAR_BMPK, datapokok.LAMPAU_BMPK, datapokok.RATING_DEBITUR, datapokok.LEMBAGA_RATING, datapokok.GO_PUBLIC FROM datapokok, rekredit WHERE datapokok.idnama = rekredit.idnama AND rekredit.din <> '-' and saldoenc > 0")
        ttldeb = String(7 - Len(ttldeb), "0") & ttldeb
    Else
    
        ttldeb = String(7, "0")
    End If
    
    If Check2 = Checked Then
        ttlpenjamin = totalRecord("SELECT rekredit.norek, rekredit.cif_bi, datapokok.idnama, datapokok.din, t_penjamin.* FROM datapokok, rekredit, t_penjamin WHERE rekredit.idnama = datapokok.idnama AND t_penjamin.norek = rekredit.norek")
        ttlpenjamin = String(7 - Len(ttlpenjamin), "0") & ttlpenjamin
    Else

        ttlpenjamin = String(7, "0")
    End If
    
    If Check3 = Checked Then
        ttlpenempatan = totalRecord("SELECT `datapokok`.`nama`, (`bngab`.`pro1` * 100) AS bunga, `rekbank`.`KODESLD`, `rekbank`.`sandibank`, `rekbank`.`jenis`, `rekbank`.`jkw`, `rekbank`.`ppap`, `rekbank`.`coll`, ROUND(SUM(`rekbank`.`saldoenc`)) AS nilai_penempatan FROM datapokok, rekbank, bngab WHERE `datapokok`.`idnama` = `rekbank`.`idnama` AND `rekbank`.`KODEBNG` = `bngab`.`Kodebng` AND rekbank.`KODESLD` = 'D' GROUP BY `rekbank`.`jenis`, `rekbank`.`sandibank` ORDER BY `rekbank`.`sandibank`")
        ttlpenempatan = String(7 - Len(ttlpenempatan), "0") & ttlpenempatan
    Else
        ttpenempatan = String(7, "0")
    End If
    
    If Check4 = Checked Then
        ttlkredit = totalRecord("select rekredit.norek, rekredit.tglmasuk , rekredit.tgljt, rekredit.tglmasuk as tgl_awal_kredit, rekredit.kali, rekredit.guna, rekredit.sektor, datapokok.idlokasi, bngkredit.pro1, rekredit.carahitung, rekredit.plafon, rekredit.saldoenc, rekredit.gp, rekredit.coll,  rekredit.denda, rekredit.jaminan, rekredit.ppap, rekredit.tglr, rekredit.sifat, rekredit.jaminan, rekredit.nilaijaminan from rekredit, bngkredit, datapokok where rekredit.kodebng = bngkredit.kodebng and rekredit.idnama = datapokok.idnama and rekredit.cif_bi <> '-'")
        ttlkredit2 = totalRecord("SELECT DISTINCT logkredit.norek, mutasi.tglmut FROM logkredit, mutasi WHERE logkredit.norek = mutasi.norek AND MONTH(mutasi.`TGLMUT`) = MONTH((SELECT settanggal FROM settanggal)) AND YEAR(mutasi.tglmut) =  YEAR((SELECT settanggal FROM settanggal)) GROUP BY logkredit.norek ORDER BY mutasi.`TGLMUT` DESC")
        ttlkredit = Int(ttlkredit) + Int(ttlkredit2)
        ttlkredit = String(7 - Len(ttlkredit), "0") & ttlkredit
    Else
        ttlkredit = String(7, "0")
    End If
    
    
    If Check6 = Checked Then
        ttldataagunan = totalRecord("SELECT data_agunan.*, datapokok.din,rekredit.cif_bi, rekredit.pengikatan FROM data_agunan, datapokok, rekredit WHERE datapokok.idnama = rekredit.idnama AND rekredit.norek = data_agunan.norek AND id_agunan_bi <> '-'")
        ttldataagunan = String(7 - Len(ttldataagunan), "0") & ttldataagunan
    Else
        ttldataagunan = String(7, "0")
    End If
    
    If Check8 = Checked Then
      ttlktrllbpr = 1
    End If
    
    If Check10 = Checked Then
        
        Set rsdatrelasi = New ADODB.Recordset
        'strrelasi = "select id_fasilitas, no_rekening, jenis_fasilitas from t_kredit where id_fasilitas not in(select id_deb_fas from r_debitur_fasilitas )"
        strrelasi = "select r_debitur_fasilitas.id_deb_fas as id_deb_fas, t_debitur.id_debitur as id_debitur, t_kredit.id_fasilitas as id_fasilitas, t_din.din as din, t_kredit.no_rekening as norek, t_kredit.operation as operation from r_debitur_fasilitas, t_debitur, t_kredit, t_din where r_debitur_fasilitas.id_fasilitas = t_kredit.id_fasilitas and r_debitur_fasilitas.id_debitur = t_debitur.id_debitur and t_din.din = t_debitur.din"
        Set rsdatarelasi = dbsid.Execute(strrelasi)
        ttldatarelasi = rsdatarelasi.RecordCount
        ttldatarelasi = String(7 - Len(ttldatarelasi), "0") & ttldatarelasi
    Else
    
        ttldatarelasi = String(7, "0")
    End If
    
    versiAplikasi = String(4, " ")
    versiReferensi = String(4, " ")
    versiValidasi = String(4, " ")
    statusData = String(4, "0")
    
    ttlrecord = Val(ttldeb) + Val(ttlpengurus) + Val(ttlpenempatan) + Val(ttlsuratberharga) + Val(ttlkredit) + Val(ttltagihanlainnya) + Val(ttlpenyertaan) + Val(ttlirrecovable) + Val(ttlgaransi) + Val(ttlkreditkelolaan) + Val(ttldatagunan) + Val(ttlpenjamin) + Val(ttlktrllbpr) + Val(ttldatakeu) + Val(ttldataagunan) + Val(ttldatarelasi)
    ttlrecord = String(12 - Len(ttlrecord), "0") & ttlrecord
    
    header = jenisdata & iddata & idlembaga & idbank & idktrcab & blndata & thndata & waktucreate & nmpengirim & versiAplikasi & versiReferensi & versiValidasi & statusData & ttldeb & ttlpengurus & ttlpenempatan & ttlsuratberharga & ttlkredit & ttltagihanlainnya & ttlpenyertaan & ttlirrecovable & ttlgaransi & String(7, "0") & ttldataagunan & ttlpenjamin & ttlktrllbpr & ttldatakeu & ttldatarelasi & ttlrecord '& namabank & alamatbank & kodearea & notlpbank & statuskantor & statusbank
    
    strslash = IIf(Right$(App.Path, 1) = "\", "", "\")
    bukaFile = App.Path & strBackSlash & "\export\" & Format(Now, "ddmmyy") & "data-sid.txt"
    Text3.Text = bukaFile
    
    footer = jenisdata & "SID090" & idlembaga & idbank & idktrcab & blndata & thndata & String(12 - Len(ttlrecord), "0") & ttlrecord
       
    Close #1
    Open bukaFile For Output As #1
    Print #1, header
    
    
End Function

Private Function cetakData()

    Dim waktuSkrg As String
    Dim StatDeb As Integer
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim ID_Agunan As String
    Dim Wkt_agunan As Date
    Dim Cari_wkt_agunan As Date
    Dim Wkt_agunan_lalu As Date
    Dim Ulg As Double
    Dim StatDebitur As Integer
    Dim Id_fas_deb As String
    Dim rsdataDebitur As ADODB.Recordset
    Dim rsdatapengurus As ADODB.Recordset
    Dim rsdatapenempatan As ADODB.Recordset
    Dim rsDataKredit As ADODB.Recordset
    Dim rsdatapenjamin As ADODB.Recordset
    Dim rsdatakontrol As ADODB.Recordset
    Dim rsAngsuran As ADODB.Recordset
    Dim jmlhBadanUsaha As Integer
    Dim ttldeb As String
    Dim rsdatabank As ADODB.Recordset
    Dim rsdataagunan As ADODB.Recordset
    Dim rsdatakeu As ADODB.Recordset
    Dim ttlrecord As String
    Dim nama As String
    Dim plafonInduk As String
    Dim plafon As String
    Dim bakidebet As String
    Dim tunggakanpokok As String
    Dim frektunggakanpkk As String
    Dim tunggBungaIntra As String
    Dim runggBungaEkstra As String
    Dim denda As String
    Dim agunan As String
    Dim ppap As String
    Dim ttlkredit As String
    Dim rscekdebetpenempatan As ADODB.Recordset
    Dim sukubunga As String
    Dim rscekrelasi As ADODB.Recordset
    Dim cekdin As ADODB.Recordset
    Dim rsmulai As ADODB.Recordset
    Dim rshapusagunan As ADODB.Recordset
    Dim rstglkondisi As ADODB.Recordset
    Dim rscekagunanlunas As ADODB.Recordset
    Dim strdatakred2 As String
    Dim rsdatakred2 As ADODB.Recordset
    Dim str_tkredit As String
    Dim rs_tkredit As ADODB.Recordset
    
    cek_jml_data = 0
    
   
    Text3.Text = bukaFile
    
   
    'strdatadeb = "SELECT rekredit.din, rekredit.CIF_BI, datapokok.nama, datapokok.alias, datapokok.status, datapokok.ket_status, rekredit.goldeb, datapokok.KELAMIN, datapokok.BUKTIDIRI, datapokok.NO_PASPOR, datapokok.T_LAHIR, datapokok.TGL_LAHIR, datapokok.TGL_AKTE_AWAL, datapokok.NPWP_DEBITUR, datapokok.ALAMAT, datapokok.IDLOKASI, datapokok.KODE_POS, datapokok.KELURAHAN, datapokok.KECAMATAN, datapokok.KODE_AREA, datapokok.telp, datapokok.NEGARA_DOMISILI, datapokok.IBU_DEBITUR, datapokok.SANDI_PEKERJAAN, datapokok.TEMPAT_BEKERJA, datapokok.BIDANG_USAHA, rekredit.sektor, datapokok.HUB_DGN_BANK, datapokok.LANGGAR_BMPK, datapokok.LAMPAU_BMPK, datapokok.RATING_DEBITUR, datapokok.LEMBAGA_RATING, datapokok.GO_PUBLIC FROM datapokok, rekredit WHERE datapokok.idnama = rekredit.idnama AND rekredit.saldoenc <> 0 and rekredit.din <> '-'"
    strdatadeb = "SELECT rekredit.din, rekredit.CIF_BI, datapokok.nama, datapokok.alias, datapokok.status, datapokok.ket_status, rekredit.goldeb, datapokok.KELAMIN, datapokok.BUKTIDIRI, datapokok.NO_PASPOR, datapokok.T_LAHIR, datapokok.TGL_LAHIR, datapokok.TGL_AKTE_AWAL, datapokok.NPWP_DEBITUR, datapokok.ALAMAT, datapokok.IDLOKASI, datapokok.KODE_POS, datapokok.KELURAHAN, datapokok.KECAMATAN, datapokok.KODE_AREA, datapokok.telp, datapokok.NEGARA_DOMISILI, datapokok.IBU_DEBITUR, datapokok.SANDI_PEKERJAAN, datapokok.TEMPAT_BEKERJA, datapokok.BIDANG_USAHA, rekredit.sektor, datapokok.HUB_DGN_BANK, datapokok.LANGGAR_BMPK, datapokok.LAMPAU_BMPK, datapokok.RATING_DEBITUR, datapokok.LEMBAGA_RATING, datapokok.GO_PUBLIC FROM datapokok, rekredit WHERE datapokok.idnama = rekredit.idnama AND rekredit.din <> '-' and saldoenc > 0"
    Set rsdataDebitur = New ADODB.Recordset
    Set rsdataDebitur = dbbank.Execute(strdatadeb)
    
    
    
    If jenis_debitur = 1 Then
    
        jmlhBadanUsaha = jmlhBadanUsaha + 1
    
    End If
    
    
    
    '=======================================================================================
    
    '2. data pengurus
    
    
    
    
    
    
    
    
    
    
    '===================================================================================================
    
    '3. data penempatan
    
    
    strdatapenempatan = "SELECT `datapokok`.`nama`, (`bngab`.`pro1` * 100) AS bunga, `rekbank`.`KODESLD`, `rekbank`.`sandibank`, `rekbank`.`jenis`, `rekbank`.`jkw`, `rekbank`.`ppap`, `rekbank`.`coll`, ROUND(SUM(`rekbank`.`saldoenc`)) AS nilai_penempatan FROM datapokok, rekbank, bngab WHERE `datapokok`.`idnama` = `rekbank`.`idnama` AND `rekbank`.`KODEBNG` = `bngab`.`Kodebng` AND rekbank.`KODESLD` = 'D' GROUP BY `rekbank`.`jenis`, `rekbank`.`sandibank` ORDER BY `rekbank`.`sandibank`"
    
    Set rsdatapenempatan = New ADODB.Recordset
    Set rsdatapenempatan = dbbank.Execute(strdatapenempatan)
    
    
    
    '===================================================================================================
    '5. data kredit
    'akad awal blom ada
    
    strdatakredit = "select * from data_kredit"
    Set rsDataKredit = New ADODB.Recordset
    Set rsDataKredit = dbbank.Execute(strdatakredit)
    
    ttldatakredit = rsDataKredit.RecordCount
    
    
    '===================================================================================================
    'data agunan
    '==================================================================================================
    
    'strdataagunan = "SELECT rekredit.CIF_BI, rekredit.jagunan AS jnisAgunan, rekredit.pengikatan , datapokok.nama, rekredit.jaminan AS buktimilik, datapokok.idlokasi AS lokasi, rekredit.jaminan AS alamat, rekredit.nilaijaminan, rekredit.NILAIJAMINAN1, rekredit.NILAIJAMINAN1 AS nilai_agunan_penilai, rekredit.TGLMASUK AS tglnilai FROM datapokok, rekredit WHERE datapokok.idnama = rekredit.idnama "
    strdataagunan = "SELECT data_agunan.*, datapokok.din,rekredit.cif_bi FROM data_agunan, datapokok, rekredit WHERE datapokok.idnama = rekredit.idnama AND rekredit.norek = data_agunan.norek AND id_agunan_bi <> '-'"
    Set rsdataagunan = New ADODB.Recordset
    Set rsdataagunan = dbbank.Execute(strdataagunan)
    ttldataagunan = rsdataagunan.RecordCount
    
    
    '=================================================================================================
    'data penjamin
    '=================================================================================================
    strdatapenjamin = "SELECT datapokok.din, rekredit.cif_bi, rekredit.jamin AS gol_jamin, rekredit.prosenjam AS prosenjam FROM datapokok, rekredit WHERE datapokok.idnama = rekredit.idnama AND rekredit.cif_bi <> '-'"
    Set rsdatapenjamin = New ADODB.Recordset
    Set rsdatapenjamin = dbbank.Execute(strdatapenjamin)
    
    
    
    'kontrol lbu
    '-----------------------------------------------------------------------------------------------
    
    
    '================================================================================================
    'data keuangan form06
    
    strdatakeu = "select rekredit.cif_bi from rekredit"
    Set rsdatakeu = New ADODB.Recordset
    Set rsdatakeu = dbbank.Execute(strdatakeu)
    
    
    
    '=================================================================================================
    'data bank
    
    strdatabank = "select * from T_INFO_BANK"
    Set rsdatabank = New ADODB.Recordset
    Set rsdatabank = dbsid.Execute(strdatabank)
    
    
    '===================================================================================================================
    'data kredit
    strdatakred = "SELECT rekredit.norek, rekredit.tglmasuk , rekredit.tgljt, rekredit.tglmasuk AS tgl_awal_kredit, rekredit.tpokok, rekredit.kali, rekredit.guna, rekredit.no_akad_awal, rekredit.no_akad_akhir, rekredit.sektor, rekredit.tbunga, datapokok.idlokasi, bngkredit.pro1, rekredit.carahitung, rekredit.plafon, rekredit.saldoenc, rekredit.gp, rekredit.coll,  rekredit.denda, rekredit.jaminan, rekredit.ppap, rekredit.tglr, rekredit.sifat, rekredit.jaminan, rekredit.nilaijaminan, (rekredit.tpokok/rekredit.pokok) AS xtpokok, (rekredit.tbunga/rekredit.bunga) AS xtbunga FROM rekredit, bngkredit, datapokok WHERE rekredit.kodebng = bngkredit.kodebng AND rekredit.idnama = datapokok.idnama AND rekredit.cif_bi <> '-' "
    Set rsdatakred = New ADODB.Recordset
    Set rsdatakred = dbbank.Execute(strdatakred)
    
    strdatakred2 = "SELECT DISTINCT logkredit.norek as norek , mutasi.tglmut as tglmut FROM logkredit, mutasi WHERE logkredit.norek = mutasi.norek AND MONTH(mutasi.`TGLMUT`) = MONTH((SELECT settanggal FROM settanggal)) AND YEAR(mutasi.tglmut) =  YEAR((SELECT settanggal FROM settanggal)) GROUP BY logkredit.norek ORDER BY mutasi.`TGLMUT` DESC"
    Set rsdatakred2 = New ADODB.Recordset
    Set rsdatakred2 = dbbank.Execute(strdatakred2)
    
         
    '==================================================================================================================
    
    'file target
    
    strslash = IIf(Right$(App.Path, 1) = "\", "", "\")
    bukaFile = App.Path & strBackSlash & "\export\" & Format(Now, "ddmmyy") & "data-sid.txt"
    
'    bukaFile = App.Path & "\export\" & Format(Now, "ddmmyy") & "data-sid.txt"
    
   
    
    
    
    '=======================================================================================================================
    'cetak header

    
    '=======================================================================================================================
    '2. cetak data debitur
    
'    txtstatus.Text = "Export data Debitur" & vbNewLine
'    txtstatus.Refresh
'    txtstatus.Text = txtstatus.Text & "Total debitur :" & rsdataDebitur.RecordCount & vbNewLine
'     txtstatus.Refresh
'    statusheaddatadeb = txtstatus.Text
'    jlhdata = 0

    If Check1 = Checked Then
    
    pgbar.Value = 0
    pgbar.Max = rsdataDebitur.RecordCount

    
    While Not rsdataDebitur.EOF
    
        DoEvents
        
        Text2.Refresh
        Text2.Text = " Data Debitur"
        Text1.Refresh
        jlhdata = jlhdata + 1
        Text1.Alignment = vbRightJustify
        Text1.Text = jlhdata & " / " & rsdataDebitur.RecordCount
        Text1.Refresh
        id_data_deb = "SID010"
        pgbar.Value = pgbar.Value + 1
        
        If IsNull(rsdataDebitur!cif_bi) Or rsdataDebitur!cif_bi = "" Or rsdataDebitur!cif_bi = "-" Then
              operation = "C"
        Else
            operation = "U"
        End If
            
        
        ID_Debitur = rsdataDebitur!cif_bi
        If ID_Debitur = "-" Then
            ID_Debitur = String(43, " ")
        Else
            ID_Debitur = ID_Debitur & String(43 - Len(ID_Debitur), " ")
        End If
        
        din = rsdataDebitur!din
        din = din & String(20 - Len(din), " ")
        
        strcekdin = "select * from t_debitur where din = '" & din & "'"
        Set cekdin = New ADODB.Recordset
        Set cekdin = dbsid.Execute(strcekdin)
        
        If Not cekdin.EOF Then
            operation = "S"
        Else
            operation = operation
        End If
        
        nama = rsdataDebitur!nama
        nama = nama & String(100 - Len(nama), " ")
        
      
        
        If IsNull(rsdataDebitur!alias) Then
            alias = String(50, " ")
        Else
            alias = alias & String(50 - Len(alias), " ")
        End If
        
        If IsNull(rsdataDebitur!status) Then
            MsgBox ("status debitur  " & rsdataDebitur!nama & " Masih Kosong ")
            status = String(4, " ")
        Else
            status = rsdataDebitur!status & String(4 - Len(rsdataDebitur!status), " ")
        End If
        
        If IsNull(rsdataDebitur!ket_status) Then
            MsgBox ("keterangan status" & rsdataDebitur!nama & " MASIH kosong")
            ket_status = String(50, " ")
        Else
            ket_status = rsdataDebitur!ket_status & String(50 - Len(rsdataDebitur!ket_status), " ")
        End If
        
        
        goldeb = rsdataDebitur!goldeb
        
        kelamin = rsdataDebitur!kelamin
        
        If kelamin = "Laki" Then
            kelamin = "1"
        Else
            kelamin = "2"
        End If
        
        no_ktp = rsdataDebitur!buktidiri
        no_ktp = no_ktp & String(30 - Len(no_ktp), " ")

        
        no_paspor = rsdataDebitur!no_paspor
        
        If IsNull(no_paspor) Then
            no_paspor = String(30, " ")
        Else
            no_paspor = no_paspor & String(30 - Len(no_paspor), " ")
        End If
        
        
        no_akte_akhir = " "
        no_akte_akhir = no_akte_akhir & String(30 - Len(no_akte_akhir), " ")
        t_lahir = rsdataDebitur!t_lahir
        t_lahir = t_lahir & String(50 - Len(t_lahir), " ")
        
        
        If IsNull(rsdataDebitur!tgl_lahir) Or rsdataDebitur!tgl_lahir = "" Then
            tgl_lahir = String(10, " ")
        Else
            tgl_lahir = rsdataDebitur!tgl_lahir
            tgl_lahir = Format(tgl_lahir, "yyyymmdd")
        End If
        
      
        tgl_akte_akhir = String(8, " ")
             
        If IsNull(rsdataDebitur!npwp_debitur) Or rsdataDebitur!npwp_debitur = "" Or rsdataDebitur!npwp_debitur = "-" Then
            npwp = "00.000.000.0.000.000"
        Else
            npwp = rsdataDebitur!npwp_debitur
            npwp = npwp & String(20 - Len(npwp), "0")
        End If
        
        Alamat = rsdataDebitur!Alamat
        Alamat = bersihString(Alamat)
        Alamat = Alamat & String(100 - Len(Alamat), " ")
        dati2 = rsdataDebitur!idlokasi
        
        If IsNull(rsdataDebitur!KODE_POS) Or rsdataDebitur!KODE_POS = "" Or rsdataDebitur!KODE_POS = "0" Then
            kopos = String(5, "0")
        Else
            kopos = rsdataDebitur!KODE_POS
            kopos = kopos & String(5 - Len(kopos), "0")
        End If
        
        If IsNull(rsdataDebitur!kelurahan) Or rsdataDebitur!kelurahan = "" Or rsdataDebitur!kelurahan = "0" Then
            kelurahan = String(50, " ")
        Else
            kelurahan = rsdataDebitur!kelurahan
            kelurahan = kelurahan & String(50 - Len(kelurahan), " ")
        End If
    
        If IsNull(rsdataDebitur!kecamatan) Or rsdataDebitur!kecamatan = "" Then
            kecamatan = String(50, " ")
        Else
            kecamatan = rsdataDebitur!kecamatan
            kecamatan = kecamatan & String(50 - Len(kecamatan), " ")
        End If
        
        If IsNull(rsdataDebitur!kode_area) Or rsdataDebitur!kode_area = "" Then
            kode_area = String(4, "0")
        Else
            kode_area = rsdataDebitur!kode_area
            kode_area = String(4 - Len(kode_area), "0") & kode_area
        End If
        
        If IsNull(rsdataDebitur!telp) Or rsdataDebitur!telp = "-" Or rsdataDebitur!telp = "" Then
            telp = String(8, "0")
        Else
            telp = Right(rsdataDebitur!telp, 8)
            telp = String(8 - Len(telp), "0") & telp
        End If
        
        If IsNull(rsdataDebitur!negara_domisili) Or rsdataDebitur!negara_domisili = "" Then
            negara = "ID "
        Else
            negara = rsdataDebitur!negara_domisili
        End If
            
        If IsNull(rsdataDebitur!IBU_DEBITUR) Or rsdataDebitur!IBU_DEBITUR = "" Then
            ibu = String(50, " ")
        Else
            ibu = rsdataDebitur!IBU_DEBITUR
            ibu = ibu & String(50 - Len(ibu), " ")
        End If
        
        If IsNull(rsdataDebitur!SANDI_PEKERJAAN) Or rsdataDebitur!SANDI_PEKERJAAN = "" Or rsdataDebitur!SANDI_PEKERJAAN = "-" Then
            sandi_kerja = "099"
        Else
            sandi_kerja = rsdataDebitur!SANDI_PEKERJAAN
            sandi_kerja = sandi_kerja & String(3 - Len(sandi_kerja), " ")
        End If
        
        If IsNull(rsdataDebitur!TEMPAT_BEKERJA) Or rsdataDebitur!TEMPAT_BEKERJA = "" Then
            tempat_kerja = String(50, " ")
        Else
            tempat_kerja = rsdataDebitur!TEMPAT_BEKERJA
            tempat_kerja = tempat_kerja & String(50 - Len(tempat_kerja), " ")
        End If
            
        If IsNull(rsdataDebitur!bidang_usaha) Or rsdataDebitur!bidang_usaha = "" Or rsdataDebitur!bidang_usaha = "-" Then
            bidang_usaha = rsdataDebitur!sektor
            bidang_usaha = bidang_usaha & String(5 - Len(bidang_usaha), " ")
            If IsEmpty(bidang_usaha) Or IsNull(bidang_usaha) Or bidang_usaha = "-    " Then
                 bidang_usaha = String(5, " ")
            End If
        Else
            bidang_usaha = rsdataDebitur!bidang_usaha
            bidang_usaha = bidang_usaha & String(5 - Len(bidang_usaha), " ")
        End If
        
        gin = String(6, " ")
        nama_group = String(20, " ")
        
        If IsNull(rsdataDebitur!HUB_DGN_BANK) Or rsdataDebitur!HUB_DGN_BANK = "" Or rsdataDebitur!HUB_DGN_BANK = "-" Then
            HUB_DGN_BANK = String(4, " ")
        Else
            HUB_DGN_BANK = rsdataDebitur!HUB_DGN_BANK
            HUB_DGN_BANK = HUB_DGN_BANK & String(4 - Len(HUB_DGN_BANK), " ")
        End If
        
        If IsNull(rsdataDebitur!LANGGAR_BMPK) Or rsdataDebitur!LANGGAR_BMPK = "" Or rsdataDebitur!LANGGAR_BMPK = "-" Then
            lgr_bmpk = "T"
        Else
            lgr_bmpk = rsdataDebitur!LANGGAR_BMPK
        End If
            
        If IsNull(rsdataDebitur!LAMPAU_BMPK) Or rsdataDebitur!LAMPAU_BMPK = "" Or rsdataDebitur!LAMPAU_BMPK = "-" Then
            LAMPAU_BMPK = "T"
        Else
            LAMPAU_BMPK = rsdataDebitur!LAMPAU_BMPK
        End If
        
        rating = String(5, " ")
        lbg_rating = String(50, " ")
        
        If IsNull(rsdataDebitur!go_public) Or rsdataDebitur!go_public = "" Or rsdataDebitur!go_public = "-" Then
            go_public = " "
        Else
            go_public = rsdataDebitur!go_public
        End If
        
        If IsEmpty(tgl_llahir) Then
            tgl_llahir = String(8, " ")
        End If
        
        If IsEmpty(negara_domisili) Then
            negara_domisili = "ID "
        End If
        
            
        If Len(paspor) = 0 Then
            paspor = String(30, " ")
        Else
            paspor = paspor & String(30 - Len(paspor), " ")
        End If
        
        If Not IsEmpty(tgl_akte_akhir) Then
            tgl_akte_akhir = tgl_akte_akhir & String(8 - Len(tgl_akte_akhir), " ")
        Else
            tgl_akte_akhir = String(8, " ")
        End If
        
       
        
        If IsEmpty(TEMPAT_BEKERJA) Then
            TEMPAT_BEKERJA = String(50, " ")
        Else
            TEMPAT_BEKERJA = TEMPAT_BEKERJA & String(50 - Len(TEMPAT_BEKERJA), " ")
        End If
        
        waktucreate = Format(Now, "yyyymmddhhnnss")
        Update_Date = Format(Now, "yyyymmddhhnnss")
        
        'cek lagi
        If din = "06582042115119000114" Then
            nasa = cek
        End If
        
        If operation = "U" Then
        
            nasa = cek
        End If
        
               
        isidatadeb = id_data_deb & operation & idlembaga & idbank & idkancab & blndata & thndata & ID_Debitur & din & nama & alias & status & ket_status & goldeb & kelamin & no_ktp & paspor & no_akte_akhir & t_lahir & tgl_llahir & tgl_akte_akhir & npwp & Alamat & dati2 & kopos & kelurahan & kecamatan & kode_area & telp & negara_domisili & ibu & sandi_kerja & tempat_kerja & bidang_usaha & gin & nama_group & HUB_DGN_BANK & lgr_bmpk & LAMPAU_BMPK & rating & lbg_rating & go_public & waktucreate & nmpengirim & Update_Date
        
        If Len(isidatadeb) < 940 Then
         nasa = cek
        End If
        
        Print #1, isidatadeb
        
    rsdataDebitur.MoveNext
    Wend
    
    End If
    
    '======================================================================================================================
    'cetak data penjamin
    '======================================================================================================================
    
    If Check2 = Checked Then

        Dim rspenjamin As ADODB.Recordset
        Dim strpenjamin As String
        Dim cekpenjamin As ADODB.Recordset
        Dim strcekpenjamin As String
        Dim debpenjamin As ADODB.Recordset
        Dim strdebpenjamin As String
        Dim iddeb As String
        Dim ID_Fasilitas As String
        Dim Id_Penjamin As String
        Dim nama_Penjamin As String
        Dim Golongan_Penjamin As String
        Dim Bagian_Dijamin As String
        Dim Identitas_Penjamin As String
        
        
        Dim Prime As String
        Dim Create As String
        Dim Create_user As String
        
       
        Dim status_Kirim As String
        Dim isipenjamin As String

        id_data = "SID042"


        strnoreksid = "SELECT rekredit.norek, rekredit.cif_bi, datapokok.idnama, datapokok.din, t_penjamin.* FROM datapokok, rekredit, t_penjamin WHERE rekredit.idnama = datapokok.idnama AND t_penjamin.norek = rekredit.norek"
        Set rspenjamin = New ADODB.Recordset
        Set rspenjamin = dbbank.Execute(strnoreksid)

        While Not rspenjamin.EOF
        
            strcekpenjamin = "select * from T_PENJAMIN where NAMA_PENJAMIN = '" & rspenjamin!nama_Penjamin & "' and ID_PENJAMIN = '" & rspenjamin!Identitas_Penjamin & "'"
            Set cekpenjamin = New ADODB.Recordset
            Set cekpenjamin = dbsid.Execute(strcekpenjamin)
            
            If Not cekpenjamin.EOF Then
                strdebpenjamin = "select * from T_PENJAMIN where ID_DEBITUR = '" & rspenjamin!cif_bi & "'"
                Set debpenjamin = New ADODB.Recordset
                Set debpenjamin = dbsid.Execute(strdebpenjamin)
                
                If Not debpenjamin.EOF Then
                    operation = "S"
                    iddeb = debpenjamin!ID_Debitur
                Else
                    operation = "C"
                    iddeb = rspenjamin!cif_bi
                End If
                iddeb = iddeb & String(43 - Len(iddeb), " ")
                ID_Fasilitas = String(52, " ")
                Id_Penjamin = debpenjamin!Id_Penjamin
                Id_Penjamin = Id_Penjamin & String(29 - Len(Id_Penjamin), " ")
                nama_Penjamin = debpenjamin!nama_Penjamin
                nama_Penjamin = nama_Penjamin & String(50 - Len(nama_Penjamin), " ")
                Golongan_Penjamin = debpenjamin!GOL_PENJAMIN
                Bagian_Dijamin = Val(debpenjamin!bag_dijamin)
                Bagian_Dijamin = String(5 - Len(Bagian_Dijamin), "0") & Bagian_Dijamin
                Identitas_Penjamin = debpenjamin!Identitas_Penjamin
                Identitas_Penjamin = Identitas_Penjamin & String(25 - Len(Identitas_Penjamin), " ")
                npwp = debpenjamin!npwp_penjamin
                npwp = npwp & String(20 - Len(npwp), " ")
                Alamat = debpenjamin!ALAMAT_PENJAMIN
                Alamat = Alamat & String(100 - Len(Alamat), " ")
                Prime = debpenjamin!PRIME_BANK
                din = rspenjamin!din
                din = din & String(20 - Len(din), " ")
                status_Kirim = "1"
                
            Else
                operation = "C"
                iddeb = iddeb = rspenjamin!cif_bi
                id_deb = iddeb & String(43 - Len(iddeb), " ")
                ID_Fasilitas = String(52, " ")
                Id_Penjamin = String(29, " ")
                nama_Penjamin = rspenjamin!nama_Penjamin
                nama_Penjamin = nama_Penjamin & String(50 - Len(nama_Penjamin), " ")
                Golongan_Penjamin = rspenjamin!Golongan_Penjamin
                Bagian_Dijamin = Val(rspenjamin!bag_dijamin) * 100
                Bagian_Dijamin = String(5 - Len(Bagian_Dijamin), "0") & Bagian_Dijamin
                Identitas_Penjamin = rspenjamin!Identitas_Penjamin
                Identitas_Penjamin = Identitas_Penjamin & String(25 - Len(Identitas_Penjamin), " ")
                npwp = rspenjamin!npwp_penjamin
                npwp = npwp & String(20 - Len(npwp), " ")
                Alamat = rspenjamin!ALAMAT_PENJAMIN
                Alamat = Alamat & String(100 - Len(Alamat), " ")
                Prime = rspenjamin!jenis_penjamin
                din = rspenjamin!din
                din = din & String(20 - Len(din), " ")
                status_Kirim = "1"
            
            End If
                Create = Format(Now, "yyyymmddhhnnss")
                If Len(nmpengirim) > 20 Then
                    createuser = Left(nmpengirim, 20)
                Else
                    createuser = nmpengirim & String(20 - Len(nmpengirim), " ")
                End If
                    
                Update_Date = Format(Now, "yyyymmddhhnnss")
                
                isipenjamin = id_data & operation & idlembaga & idbank & idkancab & blndata & thndata & id_deb & ID_Fasilitas & Id_Penjamin & Golongan_Penjamin & Bagian_Dijamin & Identitas_Penjamin & npwp & Alamat & Prime & Create & createuser & Update_Date & din & status_Kirim
                
                Print #1, isipenjamin
                
            rspenjamin.MoveNext
            Wend
       End If
            
            
'
'
'
'            strpenjamin = "SELECT T_DEBITUR.ID_DEBITUR, T_KREDIT.ID_FASILITAS, T_PENJAMIN.ID_PENJAMIN, T_PENJAMIN.GOL_PENJAMIN, T_PENJAMIN.NAMA_PENJAMIN, T_PENJAMIN.IDENTITAS_PENJAMIN, T_PENJAMIN.ALAMAT_PENJAMIN, T_KREDIT.KONDISI FROM T_KREDIT INNER JOIN R_DEBITUR_FASILITAS INNER JOIN T_DEBITUR INNER JOIN T_PENJAMIN ON T_DEBITUR.ID_DEBITUR = T_PENJAMIN.ID_DEBITUR ON R_DEBITUR_FASILITAS.ID_DEBITUR = T_DEBITUR.ID_DEBITUR ON T_KREDIT.ID_FASILITAS = R_DEBITUR_FASILITAS.ID_FASILITAS where t_debitur.id_debitur = '" & rsdatapenjamin!cif_bi & "'"
'            Set rspenjamin = New ADODB.Recordset
'            Set rspenjamin = dbsid.Execute(strpenjamin)
'
'            If rspenjamin!kondisi = "02" Then
'                operation = "D"
'            Else
'                operation = "S"
'            End If
'
'            id_debitur = rspenjamin!id_debitur
'            id_fasilitas = String(52, " ")
'            id_penjamin = rspenjamin!id_penjamin
'            nama_penjamin = rspenjamin!nama_penjamin
'            GOL_PENJAMIN = rspenjamin!GOL_PENJAMIN
'            bag_jamin = "10000"
'            identitas_penjamin = rspenjamin!identitas_penjamin
'            npwp_penjamin = String(20, " ")
'            alamat_penjamin = rspenjamin!alamat_penjamin
'            prime_bank = "T"
'            din = rspenjamin!din
'
'        Else
'
'            strpenjamin = "SELECT T_DEBITUR.ID_DEBITUR, T_KREDIT.ID_FASILITAS, T_PENJAMIN.ID_PENJAMIN, T_PENJAMIN.GOL_PENJAMIN, T_PENJAMIN.NAMA_PENJAMIN, T_PENJAMIN.IDENTITAS_PENJAMIN, T_PENJAMIN.ALAMAT_PENJAMIN, T_KREDIT.KONDISI FROM T_KREDIT INNER JOIN R_DEBITUR_FASILITAS INNER JOIN T_DEBITUR INNER JOIN T_PENJAMIN ON T_DEBITUR.ID_DEBITUR = T_PENJAMIN.ID_DEBITUR ON R_DEBITUR_FASILITAS.ID_DEBITUR = T_DEBITUR.ID_DEBITUR ON T_KREDIT.ID_FASILITAS = R_DEBITUR_FASILITAS.ID_FASILITAS where t_debitur.id_debitur = '" & rsdatapenjamin!cif_bi & "'"
'            Set rspenjamin = New ADODB.Recordset
'            Set rspenjamin = dbsid.Execute(strpenjamin)
'
'            operation = "C"
'
'            id_debitur = rspenjamin!id_debitur
'            id_fasilitas = String(52, " ")
'            id_penjamin = rspenjamin!id_penjamin
'            nama_penjamin = rspenjamin!nama_penjamin
'            GOL_PENJAMIN = rspenjamin!GOL_PENJAMIN
'            bag_jamin = "10000"
'            identitas_penjamin = rspenjamin!identitas_penjamin
'            npwp_penjamin = String(20, " ")
'            alamat_penjamin = rspenjamin!alamat_penjamin
'            prime_bank = "T"
'            din = rspenjamin!din
'
'
'        End If
'
'        alamat_penjamin = alamat_penjamin & String(100 - Len(alamat_penjamin), " ")
'        identitas_penjamin = identitas_penjamin & String(25 - Len(identitas_penjamin), " ")
'        nama_penjamin = nama_penjamin & String(50 - Len(nama_penjamin), " ")
'        id_penjamin = id_penjamin & String(29 - Len(id_penjamin), " ")
'        id_debitur = id_debitur & String(43 - Len(id_debitur), " ")
'        din = din & String(20 - Len(din), " ")
'        waktucreate = Format(Now, "yyyymmddhhnnss")
'        update_date = Format(Now, "yyyymmddhhnnss")
'
'
'    End If
    
    '=======================================================================================================================
    ' cetak data penempatan
    
    If Check3 = Checked Then
        
        id_data = "SID031"
        pgbar.Value = 0
        pgbar.Max = rsdatapenempatan.RecordCount
        Text2.Text = "Data Penempatan"
        
        While Not rsdatapenempatan.EOF
            DoEvents
            
            strcekpenempatan = "select * from t_penempatan where sandi_bank = '" & rsdatapenempatan!sandibank & "' and jns_penempatan = '" & rsdatapenempatan!jenis & "'"
            Set rscekpenempatan = New ADODB.Recordset
            Set rscekpenempatan = dbsid.Execute(strcekpenempatan)
            
            pgbar.Value = pgbar.Value + 1
            Text1.Refresh
            Text1.Text = pgbar.Value & " / " & rsdatapenempatan.RecordCount

            
            If Not rscekpenempatan.EOF Then
                
'                Set rscekdebetpenempatan = New ADODB.Recordset
'                strcekbakidebet = "select nilai_penempatan, id_fasilitas from t_penempatan where id_fasilitas = '" & rsdatapenempatan!bi_fasilitas & "'"
'                Set rscekdebetpenempatan = dbsid.Execute(strcekbakidebet)
                operation = "U"
                bi_fasilitas = rscekpenempatan!ID_Fasilitas
            
            Else
                operation = "C"
                bi_fasilitas = String(32, " ")
                
            
            End If
            
            bi_failitas = bi_fasilitas & String(32 - Len(bi_fasilitas), " ")
            
                If rsdatapenempatan!nilai_penempatan <> rscekpenempatan!nilai_penempatan Then

                    If rsdatapenempatan!nilai_penempatan = 0 Then
                        kondisi = "02"
                        
                        Dim rscektglkondisi  As New ADODB.Recordset
                        Dim strtglkondisi As String
                        
                        strtglkondisi = "select * from mutasi where norek = '" & rsdatapenempatan!norek & "' order by no desc"
                        Set rscektglkondisi = New ADODB.Recordset
                        Set rscektglkondisi = dbbank.Execute(strtglkondisi)
                        
                        tgl_kondisi = Format(rscektglkondisi!tglmut, "yyyymmdd")
                    Else
                        kondisi = "  "
                        tgl_kondisi = String(8, " ")
                        
                    End If
                    
                Else
                    operation = "S"
                    'bi_fasilitas = rsdatapenempatan!bi_fasilitas
                    bi_fasilitas = bi_fasilitas & String(32 - Len(bi_fasilitas), " ")
                    kondisi = "  "
                    tgl_kondisi = String(8, " ")
                End If
            
            
            
            bln = Format(Now, "mm")
            thn = Format(Now, "yyyy")
            jnis_fasilitas = "0400"
            

            
            
            jenis = rsdatapenempatan!jenis
            jenis = jenis & String(2 - Len(jenis), " ")
            sandibanktempat = rsdatapenempatan!sandibank
            sandibanktempat = sandibanktempat & String(3 - Len(sandibanktempat), " ")
            negara = "ID "
            valuta = "IDR"
            
            If jenis = "20" Or jenis = "21" Or jenis = "12" Or jenis = "10" Or jenis = "11" Then
                jkwbln = "000"
                jkwhari = "00"
            Else
                jkwbln = rsdatapenempatan!jkw
                jkwbln = String(3 - Len(jkwbln), "0") & jkwbln
                jkwhari = jkwbln * 30
            End If
            
            koll = rsdatapenempatan!coll
            
            '--- tambah field suku bunga rekbank
        
            sukubunga = rsdatapenempatan!bunga
            sukubunga = String(5 - Len(sukubunga), "0") & sukubunga
            nilaipenempatan = rsdatapenempatan!nilai_penempatan
            nilaipenempatan = String(15 - Len(nilaipenempatan), "0") & nilaipenempatan
            orig_currency = String(15, "0")
'            kondisi = "00"
'            tgl_kondisi = String(8, " ")
            agunan = String(15, "0")
            ppap = rsdatapenempatan!ppap
            ppap = String(15 - Len(ppap), "0") & ppap
            ket = String(100, "0")
            createdate = Format(Now, "yyyymmddhhnnss")
            coll = rsdatapenempatan!coll
            
            If Len(nmpengirim) > 20 Then
                createuser = Left(nmpengirim, 20)
            Else
                createuser = nmpengirim & String(20 - Len(nmpengirim), " ")
            End If
            
            updatedate = Format(Now, "yyyymmddhhnnss")
            
            cetakdatapenempatan = id_data & operation & idlembaga & idbank & idkancab & blndata & thndata & jnis_fasilitas & bi_fasilitas & jenis & sandibanktempat & negara & valuta & jkwbln & jkwhari & coll & sukubunga & nilaipenempatan & String(15, "0") & kondisi & tgl_kondisi & String(15, "0") & ppap & String(100, " ") & Format(Now, "yyyymmddhhnnss") & createuser & Format(Now, "yyyymmddhhnnss")
            
            If Len(cetakdatapenempatan) < 301 Then
                nasa = cek
            End If
            
            
            Print #1, cetakdatapenempatan
            
            rsdatapenempatan.MoveNext
        Wend
    
    End If
    
    'yogotak hubuluk motok hanorogo
    
    '=======================================================================================================================
    'cetak data kredit
    
    If Check4 = Checked Then
    
    id_data = "SID033"
    jenisFasilitas = "0605"
    golKredit = "99"
    orientasi = "9"
    
    pgbar.Value = 0
    pgbar.Max = rsdatakred.RecordCount
    jlhdata = 0
    
    no = 1
    recordke = 0
    
    While Not rsdatakred.EOF
        
        no = no + 1
        
        DoEvents
    
        Text2.Refresh
        Text2.Text = "Data kredit"
        Text1.Refresh
        jlhdata = jlhdata + 1
        Text1.Text = jlhdata & " / " & rsdatakred.RecordCount
        Text1.Refresh
        id_data_deb = "SID010"
        pgbar.Value = pgbar.Value + 1
    
    
    'cek rek
    
    If rsdatakred!norek = "130.00341" Then
        nasa = nasa
    End If
    
    strnoreksid = "select no_rekening, baki_debet, id_fasilitas, sebab_macet, ket_sebab_macet, sektor_ekonomi, no_pk_awal, no_pk_akhir, nilai_proyek, kolektibilitas, tgl_macet, PPAP from t_kredit where no_rekening = '" & rsdatakred!norek & "'"
    Set rscekrek = New ADODB.Recordset
    Set rscekrek = dbsid.Execute(strnoreksid)
    
    If rscekrek.EOF And rscekrek.BOF Then
        operation = "C"
        ID_Fasilitas = String(52, " ")
        
        sektor = rsdatakred!sektor
        
        If sektor = "1007" Then
            sektor = "6500"
        ElseIf sektor = "1009" Then
            sektor = "7000"
        ElseIf sektor = "1011" Then
            sektor = "8190"
        ElseIf sektor = "1013" Then
            sektor = "9390"
        ElseIf sektor = "1014" Then
            sektor = "9210"
        ElseIf sektor = "1015" Then
            sektor = "9100"
        Else
            sektor = "9990"
        End If
        
        nilaiproyek = String(15, " ")
        sebabmacet = " "
        ketsebabmacet = String(100, " ")
        kolektibilitas = rsdatakred!coll
        kondisi = "  "
        tglkondisi = String(8, " ")
    Else
        If rsdatakred!saldoenc = rscekrek!baki_debet And rsdatakred!coll <> rscekrek!kolektibilitas And rsdatakred!ppap = rscekrek!ppap Then
            ' asli operation = "S"
            operation = "U"
            ID_Fasilitas = rscekrek!ID_Fasilitas
            ID_Fasilitas = ID_Fasilitas & String(52 - Len(ID_Fasilitas), " ")
            sektor = rscekrek!SEKTOR_EKONOMI
            
            If rscekrek!kolektibilitas = 5 And rsdatakred!coll = 4 Then
                kolektibilitas = rscekrek!kolektibilitas
                sebabmacet = rscekrek!sebab_macet
                ketsebabmacet = rscekrek!ket_sebab_macet & String(100 - Len(rscekrek!ket_sebab_macet), " ")
            Else
                kolektibilitas = rsdatakred!coll
            End If
            
'        ElseIf rscekrek!kolektibilitas = 5 And Format(rscekrek!tgl_macet, "yyyy") < thndata Then
'
'            operation = "S"
'            id_fasilitas = rscekrek!id_fasilitas
'            id_fasilitas = id_fasilitas & String(52 - Len(id_fasilitas), " ")
'            sektor = rscekrek!sektor_ekonomi
'            noakadawal = rscekrek!no_pk_awal
'            noakadakhir = rscekrek!no_pk_akhir
'            kolektibilitas = "5"
'            sebabmacet = rscekrek!sebab_macet
'            ketsebabmacet = rscekrek!ket_sebab_macet & String(100 - Len(rscekrek!ket_sebab_macet), " ")
            kondisi = "  "
            tglkondisi = String(8, " ")
            
        ElseIf rscekrek!sebab_macet <> "" Then
        
            operation = "U"
            ID_Fasilitas = rscekrek!ID_Fasilitas
            ID_Fasilitas = ID_Fasilitas & String(52 - Len(ID_Fasilitas), " ")
            sektor = rscekrek!SEKTOR_EKONOMI
            noakadawal = rscekrek!NO_PK_AWAL
            noakadakhir = rscekrek!NO_PK_AKHIR
            kolektibilitas = "5"
            sebabmacet = rscekrek!sebab_macet
            ketsebabmacet = rscekrek!ket_sebab_macet & String(100 - Len(rscekrek!ket_sebab_macet), " ")
            
'            strtgllunas = "select * from mutasi where norek = '" & rsdatakred!norek & "' order by tglmut DESC"
'            Set rstglkondisi = New ADODB.Recordset
'            Set rstglkondisi = dbbank.Execute(strtgllunas)
            kondisi = " "
            tglkondisi = String(8, " ")
            If IsNull(rscekrek!NILAI_PROYEK) Then
                nilaiproyek = String(15, " ")
            Else
                nilaiproyek = String(15 - Len(rscekrek!NILAI_PROYEK), "0") & rscekrek!NILAI_PROYEK
            End If
        
            
        ElseIf rsdatakred!saldoenc <> rscekrek!baki_debet Or rsdatakred!coll <> rscekrek!kolektibilitas Or rsdatakred!ppap <> rscekrek!ppap Then
            If rsdatakred!saldoenc = 0 Then
                operation = "U"
                kondisi = "02"
                
                strtgllunas = "select * from mutasi where norek = '" & rsdatakred!norek & "' order by tglmut DESC"
                Set rstglkondisi = New ADODB.Recordset
                Set rstglkondisi = dbbank.Execute(strtgllunas)
                kolektibilitas = "1"
                tglkondisi = Format(rstglkondisi!tglmut, "yyyymmdd")
                
                
                
            Else
                operation = "U"
                noakadawal = rscekrek!NO_PK_AWAL
                noakadakhir = rscekrek!NO_PK_AKHIR
                kolektibilitas = rsdatakred!coll
                kondisi = "  "
                tglkondisi = String(8, " ")
                
                
            End If
            
            ID_Fasilitas = rscekrek!ID_Fasilitas
            ID_Fasilitas = ID_Fasilitas & String(52 - Len(ID_Fasilitas), " ")
            sektor = rscekrek!SEKTOR_EKONOMI
            noakadawal = rscekrek!NO_PK_AWAL
            noakadakhir = rscekrek!NO_PK_AKHIR
            
            If IsNull(rscekrek!NILAI_PROYEK) Then
                nilaiproyek = String(15, " ")
            Else
                nilaiproyek = String(15 - Len(rscekrek!NILAI_PROYEK), "0") & rscekrek!NILAI_PROYEK
            End If
            
            
            
            sebabmacet = IIf(IsNull(rscekrek!sebab_macet), "  ", rscekrek!sebab_macet)
            
            If IsNull(rscekrek!ket_sebab_macet) Or IsEmpty(rscekrek!ket_sebab_macet) Or rscekrek!ket_sebab_macet = "" Then
                ketsebabmacet = String(100, " ")
            Else
                ketsebabmacet = rscekrek!ket_sebab_macet
            End If
            
            
        Else
            operation = "S"
            ID_Fasilitas = String(52, " ")
            sektor = rscekrek!SEKTOR_EKONOMI
            nilaiproyek = String(15, " ")
            kolektibilitas = rsdatakred!coll
        End If
    End If
    
    If rsdatakred!norek = "130.02173" Then
        nasa = cek
    End If
    
    If Len(ID_Fasilitas) < 52 Then
        nasa = cek
    End If
    
    strangsuran = "select * from angsuran where norek = '" & rsdatakred!norek & "' and status = 'Tertunggak' and tgl < (select settanggal from settanggal) order by angsuran limit 1"
    Set rsAngsuran = New ADODB.Recordset
    Set rsAngsuran = dbbank.Execute(strangsuran)
    
    strmulai = "select * from angsuran where norek = '" & rsdatakred!norek & "' and angsuran = 1"
    Set rsmulai = New ADODB.Recordset
    Set rsmulai = dbbank.Execute(strmulai)
    
    tglmulai = Format(rsmulai!tgl, "yyyymmdd")
    
    If Not rsAngsuran.EOF Then
        tglmacet = rsAngsuran!tgl
    Else
        tglmacet = String(8, " ")
    End If
    
    If tglmacet <> String(8, " ") Then
        strangsuran = "select * from angsuran where norek = '" & rsdatakred!norek & "' and status='tertunggak' AND tgl < (SELECT settanggal FROM settanggal) ORDER BY angsuran LIMIT 1  "
        Set rsAngsuran = New ADODB.Recordset
        Set rsAngsuran = dbbank.Execute(strangsuran)
        
        If Not rsAngsuran.EOF Then
            tunggakanpokok = rsdatakred!tpokok
            tunggakanpokok = String(15 - Len(tunggakanpokok), "0") & tunggakanpokok
            tgltgk = rsAngsuran!tgl
        
            If IsNull(rsdatakred!xtpokok) Then
                frektunggakanpkk = "0"
            Else
                frektunggakanpkk = Round(rsdatakred!xtpokok)
            End If
        
            frektunggakanpkk = String(3 - Len(frektunggakanpkk), "0") & frektunggakanpkk
            tunggBungaIntra = rsdatakred!tbunga
            tunggBungaIntra = String(15 - Len(tunggBungaIntra), "0") & tunggBungaIntra
            tunggBungaEksta = "0"
            tunggBungaEkstra = String(15 - Len(tunggBungaEkstra), "0") & tunggBungaEkstra
            sebabmacet = sebabmacet
            ketsebabmacet = ketsebabmacet
            End If
    Else
        tunggakanpokok = String(15, "0")
        frektunggakanpkk = String(3, "0")
        tunggBungaIntra = String(15, "0")
        tunggBungaEkstra = String(15, "0")
        sebabmacet = "  "
        ketsebabmacet = String(100, " ")
    End If
        
        If rsdatakred!kali = "-" Then
            statuskredit = "0"
        ElseIf rsdatakred!kali = "R" Then
            statuskredit = "1"
        ElseIf rsdatakred!kali > 0 Then
            statuskredit = rsdatakred!kali
        Else
            statuskredit = "0"
        End If
        
        statuskredit = String(2 - Len(statuskredit), "0") & statuskredit
        guna = rsdatakred!guna
        If guna = "20" Then
            guna = "79"
        End If
        
        sektor = rsdatakred!sektor
        
        If sektor = "1007" Then
            sektor = "6500"
        ElseIf sektor = "1009" Then
            sektor = "7000"
        ElseIf sektor = "1011" Then
            sektor = "8190"
        ElseIf sektor = "1013" Then
            sektor = "9390"
        ElseIf sektor = "1014" Then
            sektor = "9210"
        ElseIf sektor = "1015" Then
            sektor = "9100"
        Else
            sektor = "9990"
        End If
        
        sektor = sektor & String(5 - Len(sektor), " ")
        dati2 = rsdatakred!idlokasi
        valuta = "IDR"
        sukubunga = rsdatakred!pro1
        sukubunga = sukubunga * 100
        sukubunga = String(5 - Len(sukubunga), "0") & sukubunga
        sifat = rsdatakred!sifat
           
        
    
        sifat_kredit = "79"
        sifat_sukubunga = "1"
        norek = rsdatakred!norek
        
        If norek = "132.51761" Then
            nasa = cek
        End If
        
        norek = norek & String(25 - Len(norek), " ")
        
'        If operation = "U" Then
'            norek = String(25, " ")
'        Else
'            norek = norek & String(25 - Len(norek), " ")
'        End If
        
        norek = norek & String(25 - Len(norek), " ")
        
        If Not IsEmpty(noakadawal) Or noakadawal <> "" Then
            noakadawal = noakadawal & String(25 - Len(noakadawal), " ")
        Else
            noakadawal = String(25, " ")
        End If
        
        If Not IsEmpty(noakadakhir) Or noakadakhir <> "" Then
            noakadakhir = noakadakhir & String(25 - Len(noakadakhir), " ")
        Else
            noakadakhir = String(25, " ")
        End If
        
        tglakadawal = Format(rsdatakred!tglmasuk, "yyyymmdd")
        tglakadakhir = tglakadawal
        tglAwalKred = Format(rsdatakred!tglmasuk, "yyyymmdd")
        tglJtuhTempo = Format(rsdatakred!tgljt, "yyyymmdd")
        plafonInduk = String(15, "0")
        plafon = rsdatakred!plafon
        plafon = String(15 - Len(plafon), "0") & plafon
        
        
        oriValuta = "IDR"
        kelonggaranTarik = rsdatakred!gp
        denda = rsdatakred!denda
        denda = String(15 - Len(denda), "0") & denda
        agunan = rsdatakred!nilaijaminan
        agunan = String(15 - Len(agunan), "0") & agunan
        ppap = rsdatakred!ppap
        ppap = String(15 - Len(ppap), "0") & ppap
        
        iddatakred = "SID033"
         
        tglmacet = Format(tglmacet, "yyyymmdd")
         
         If kolektibilitas = 1 Then
            tglmacet = String(8, " ")
            
            If tunggakanpokok = String(15, "0") And tunggBungaEkstra = String(15, "0") Then
                tgltgk = String(8, " ")
            Else
                tgltgk = Format(tgltgk, "yyyymmdd")
            End If
        Else
            tglmacet = tglmacet
            tgltgk = tglmacet
        End If
        
        
         
        If kolektibilitas = 2 Then
            kolektibilitas = 3
        Else
            kolektibilitas = kolektibilitas
        End If
        
        
         
         If kolektibilitas < 5 Then
            sebabmacet = "  "
            tglmacet = String(8, " ")
        Else
            sebabmacet = sebabmacet
            ketsebabmacet = ketsebabmacet
            tglmacet = tglmacet
        End If
         
        waktucreate = Format(Now, "yyyymmddhhnnss")
         
                
        If Len(nmpengirim) > 20 Then
            createuser = Left(nmpengirim, 20)
        Else
            createuser = nmpengirim & String(20 - Len(nmpengirim), " ")
        End If
        
        recordke = recordke + 1
        
        If recordke = 286 Or recordke = 303 Then
            nasa = cek
        End If
        
        If Val(tunggakanpokok) > 0 And Val(frektunggakanpkk) = 0 Then
            frektunggakanpkk = 1
        ElseIf Val(tunggakanpokok) = 0 And Val(frektunggakanpkk) = 0 Then
            frektunggakanpkk = 0
        Else
            frektunggakanpkk = frektunggakanpkk
        End If
        
        If Val(tunggBungaIntra) > 0 Or Val(tunggBungaEkstra) > 0 Then
            
            daputunggakan = IIf(IsNull(rsdatakred!xtbunga), "1", rsdatakred!xtbunga)
            
            If daputunggakan > 0 And daputunggakan < 1 Then
                frektunggakanbunga = "1"
            Else
                frektunggakanbunga = daputunggakan
            End If
            
            frektunggakanbunga = Round(frektunggakanbunga)
            
            frektunggakanbunga = String(3 - Len(frektunggakanbunga), "0") & frektunggakanbunga
            
            
        Else
            frektunggakanbunga = String(3, "0")
        End If
        
        nilaiproyek = IIf(IsEmpty(nilaiproyek), String(15, " "), nilaiproyek)
        kondisi = IIf(kondisi = " ", "  ", kondisi)
        frektunggakanpkk = String(3 - Len(frektunggakanpkk), "0") & frektunggakanpkk
        
        If Len(frektunggakanbunga) <> 3 Then
        
            nasa = cek
        End If
        
        If kondisi = "02" Then
          bakidebet = 0
          tunggakanpokok = String(15, "0")
          frektunggakanpkk = String(3, "0")
          tunggBungaIntra = String(15, "0")
          tunggBungaEkstra = String(15, "0")
          frektunggakanbunga = String(3, "0")
        Else
          bakidebet = rsdatakred!saldoenc
          
        End If
        
        bakidebet = String(15 - Len(bakidebet), "0") & bakidebet
        
        
        
        isidatakredit = iddatakred & operation & idlembaga & idbank & idkancab & blndata & thndata & jenisFasilitas & ID_Fasilitas & sifat_kredit & norek & noakadawal & tglakadawal & noakadakhir & tglakadakhir & tglAwalKred & tglmulai & tglJtuhTempo & statuskredit & golKredit & guna & orientasi & sektor & dati2 & nilaiproyek & "IDR" & sukubunga & sifat_sukubunga & plafonInduk & plafon & bakidebet & String(15, "0") & String(15, "0") & String(15, "0") & kolektibilitas & tglmacet & sebabmacet & ketsebabmacet & tgltgk & tunggakanpokok & frektunggakanpkk & tunggBungaIntra & tunggBungaEkstra & frektunggakanbunga & denda & String(15, "0") & kondisi & tglkondisi & agunan & ppap & String(15, "0") & String(8, " ") & String(2, "0") & String(8, " ") & String(100, " ") & String(100, " ") & String(100, " ") & waktucreate & createuser & waktucreate & 1
         
         
         
         If no = 198 Then
            nasa = cek
        End If
         
         
         Print #1, isidatakredit
         
         rsdatakred.MoveNext
         Wend
         
         While Not rsdatakred2.EOF
            
           str_tkredit = "select * from T_KREDIT where NO_REKENING = '" & rsdatakred2!norek & "'"
           Set rs_tkredit = New ADODB.Recordset
           Set rs_tkredit = dbsid.Execute(str_tkredit)
           
           operation = "U"
           jenisFasilitas = "0605"
           idfasilitas = rs_tkredit!ID_Fasilitas & String(52 - Len(rs_tkredit!ID_Fasilitas), " ")
           sifat_kredit = rs_tkredit!sifat
           norek = rsdatakred2!norek
           norek = rsdatakred2!norek & String(25 - Len(norek), " ")
           noakadawal = rs_tkredit!NO_PK_AWAL & String(25 - Len(rs_tkredit!NO_PK_AWAL), " ")
           tglakadawal = Format(rs_tkredit!TGL_PK_AWAL, "yyyymmdd")
           noakadakhir = rs_tkredit!NO_PK_AKHIR & String(25 - Len(rs_tkredit!NO_PK_AWAL), " ")
           tglakadakhir = Format(rs_tkredit!TGL_PK_AKHIR, "yyyymmdd")
           tglAwalKred = Format(rs_tkredit!TGL_AWAL_KREDIT, "yyyymmdd")
           tglmulai = Format(rs_tkredit!TGL_MULAI, "yyyymmdd")
           tglJtuhTempo = Format(rs_tkredit!TGL_JT_TEMPO, "yyyymmdd")
           statuskredit = String(2 - Len(rs_tkredit!BARU_PERPANJANGAN), "0") & rs_tkredit!BARU_PERPANJANGAN
           golKredit = rs_tkredit!GOL_KREDIT & String(2 - Len(rs_tkredit!GOL_KREDIT), " ")
           guna = rs_tkredit!JEN_PENGGUNAAN
           orientasi = rs_tkredit!OR_PENGGUNAAN
           sektor = rs_tkredit!SEKTOR_EKONOMI & String(5 - Len(rs_tkredit!SEKTOR_EKONOMI), " ")
           dati2 = rs_tkredit!DATI2_LOKASI_PROYEK
           nilaiproyek = IIf(IsNull(rs_tkredit!NILAI_PROYEK) Or rs_tkredit!NILAI_PROYEK = "", "0", rs_tkredit!NILAI_PROYEK)
           nilaiproyek = String(15 - Len(nilaiproyek), "0") & nilaiproyek
           sukubunga = String(5 - Len(rs_tkredit!SUKU_BUNGA), "0") & rs_tkredit!SUKU_BUNGA
           sifat_sukubunga = rs_tkredit!SIFAT_Sk_BUNGA
           plafonInduk = String(15, "0")
           plafon = String(15 - Len(rs_tkredit!plafon), "0") & rs_tkredit!plafon
           bakidebet = String(15, "0")
           kolektibilitas = "1"
           tglmacet = String(8, " ")
           ketsebabmacet = String(100, " ")
           tgltgk = String(8, " ")
           tunggakanpokok = String(15, "0")
           frektunggakanpkk = String(3, "0")
           tunggBungaIntra = String(15, "0")
           tunggBungaEkstra = String(15, "0")
           frektunggakanbunga = String(3, "0")
           denda = String(15, "0")
           kondisi = "02"
           tglkondisi = Format(rsdatakred2!tglmut, "yyyymmdd")
           agunan = String(15 - Len(rs_tkredit!agunan), "0") & rs_tkredit!agunan
           ppap = String(15, "0")
           
           
           isidatakredit = iddatakred & operation & idlembaga & idbank & idkancab & blndata & thndata & jenisFasilitas & ID_Fasilitas & sifat_kredit & norek & noakadawal & tglakadawal & noakadakhir & tglakadakhir & tglAwalKred & tglmulai & tglJtuhTempo & statuskredit & golKredit & guna & orientasi & sektor & dati2 & nilaiproyek & "IDR" & sukubunga & sifat_sukubunga & plafonInduk & plafon & bakidebet & String(15, "0") & String(15, "0") & String(15, "0") & kolektibilitas & tglmacet & sebabmacet & ketsebabmacet & tgltgk & tunggakanpokok & frektunggakanpkk & tunggBungaIntra & tunggBungaEkstra & frektunggakanbunga & denda & String(15, "0") & kondisi & tglkondisi & agunan & ppap & String(15, "0") & String(8, " ") & String(2, "0") & String(8, " ") & String(100, " ") & String(100, " ") & String(100, " ") & waktucreate & createuser & waktucreate & 1
           Print #1, isidatakredit
         rsdatakred2.MoveNext
         Wend
         
    End If
   
    '=======================================================================================================================
    ' Data Agunan
    
    If Check6 = Checked Then
    
    iddataagunan = "SID041"
    
    pgbar.Value = 0
    pgbar.Max = rsdataagunan.RecordCount
    Text2.Text = "Data Agunan"
    
    rec = 0
    
    While Not rsdataagunan.EOF
    
    DoEvents
    
    pgbar.Value = pgbar.Value + 1
    Text1.Refresh
    Text1.Text = pgbar.Value & " / " & rsdataagunan.RecordCount
    
    rec = rec + 1
    
    If rec = 103 Then
        check = nasa
    End If
    
    If rsdataagunan!id_agunan_bi = "-" And Not IsNull(rsdataagunan!din) Then
        ID_Debitur = String(43, " ")
        operation = "C"
    Else
    
        strcekagunanlunas = "select saldoenc from rekredit where norek = '" & rsdataagunan!norek & "'"
        Set rscekagunanlunas = New ADODB.Recordset
        Set rscekagunanlunas = dbbank.Execute(strcekagunanlunas)
        
        If Not rscekagunanlunas.EOF Or Not rscekagunanlunas.BOF Then
        
            strhapus = "SELECT T_AGUNAN.ID_AGUNAN, T_DEBITUR.DIN FROM T_KREDIT INNER JOIN R_DEBITUR_FASILITAS ON T_KREDIT.ID_FASILITAS = R_DEBITUR_FASILITAS.ID_FASILITAS INNER JOIN T_DEBITUR INNER JOIN T_AGUNAN ON T_DEBITUR.ID_DEBITUR = T_AGUNAN.ID_DEBITUR ON R_DEBITUR_FASILITAS.ID_DEBITUR = T_DEBITUR.ID_DEBITUR WHERE     (T_KREDIT.BAKI_DEBET = 0) AND (T_KREDIT.T_POKOK = 0) AND (T_KREDIT.T_BUNGA_INTRA = 0) AND (T_KREDIT.T_BUNGA_EKSTRA = 0) and (T_AGUNAN.ID_AGUNAN = '" & rsdataagunan!id_agunan_bi & "')"
            Set rshapusagunan = New ADODB.Recordset
            Set rshapusagunan = dbsid.Execute(strhapus)
            
            If Not rshapusagunan.EOF Or Not rshapusagunan.BOF Then
                operation = "D"
                idagunan = rsdataagunan!id_agunan_bi
                idagunan = idagunan & String(32 - Len(idagunan), " ")
            Else
                operation = "S"
            End If
            
            ID_Debitur = rsdataagunan!cif_bi
            ID_Debitur = ID_Debitur & String(43 - Len(ID_Debitur), " ")
        
        End If
        
    End If
    
    If rsdataagunan!id_agunan_bi = "-" Then
        idagunan = String(32, " ")
    Else
        idagunan = rsdataagunan!id_agunan_bi
        idagunan = idagunan & String(32 - Len(idagunan), " ")
    End If
    
    jenis_agunan = rsdataagunan!jenis_agunan
    peringkat = String(8, " ")
    
    
    pengikatan = rsdataagunan!pengikatan
    pemilik_agunan = rsdataagunan!pemilik_agunan
    pemilik_agunan = pemilik_agunan & String(50 - Len(pemilik_agunan), " ")
    bukti_milik = rsdataagunan!bukti_milik
    
    If Len(bukti_milik) > 50 Then
        bukti_milik = Left(bukti_milik, 50)
    Else
        bukti_milik = bukti_milik & String(50 - Len(bukti_milik), " ")
    End If
    
    idlokasi = rsdataagunan!idlokasi
    idlokasi = idlokasi & String(4 - Len(idlokasi), " ")
    Alamat = rsdataagunan!alamat_agunan
    Alamat = Alamat & String(100 - Len(Alamat), " ")
    nilai_agunan = rsdataagunan!nilai_agunan
    nilai_agunan = String(15 - Len(nilai_agunan), "0") & nilai_agunan
    nilai_bank = rsdataagunan!nilai_agunan_bank
    nilai_bank = String(15 - Len(nilai_bank), "0") & nilai_bank
    nilai_penilai = rsdataagunan!nilai_agunan_penilai
    nilai_penilai = String(15 - Len(nilai_penilai), "0") & nilai_penilai
    penilai_independen = rsdataagunan!penilai_independen
    penilai_independen = penilai_independen & String(50 - Len(penilai_independen), " ")
    tgl_penilaian = Format(rsdataagunan!tgl_penilaian, "yyyymmdd")
    paripasu = rsdataagunan!paripasu
    paripasu = String(5 - Len(paripasu), "0") & paripasu
    asuransi = rsdataagunan!asuransi
    din = rsdataagunan!din & String(20 - Len(rsdataagunan!din), " ")
    
    ' jenis pengikatan ...
    
    If pengikatan = "SKMHT" Then
        pengikatan = "04"
    ElseIf pengikatan = "BAWAH TANGAN" Then
        pengikatan = "  "
    Else
        pengikatan = "  "
    End If
    
    isidataagunan = iddataagunan & operation & idlembaga & idbank & idkancab & blndata & thndata & ID_Debitur & idagunan & jenis_agunan & peringkat & pengikatan & pemilik_agunan & bukti_milik & idlokasi & Alamat & nilai_agunan & nilai_bank & nilai_penilai & nilai_independen & tgl_penilaian & paripasu & asuransi & Format(Now, "yyyymmddhhnnss") & din & 1
    
    If Len(isidataagunan) < 489 Then
        nasa = cek
    End If
    
    Print #1, isidataagunan
    
    rsdataagunan.MoveNext
    Wend
    
    End If
    
    '==========================================================================================================
    'data penjamin
    '=========================================================================================================
    
'    If Check2 = Checked Then
'
'
'        'Dim rspenjamin As ADODB.Recordset
'        'Dim strpenjamin As String
'
'        id_data = "SID042"
'
'        While Not rsdatapenjamin.EOF
'
'            DoEvents
'
'            strnoreksid = "select * from t_penjamin where id_debitur = '" & rsdatapenjamin!cif_bi & "'"
'            Set rscekrek = New ADODB.Recordset
'            Set rscekrek = dbsid.Execute(strnoreksid)
'
'            If Not rscekrek.EOF Then
'
'
'
'                strpenjmain = "SELECT T_DEBITUR.ID_DEBITUR, T_KREDIT.ID_FASILITAS, T_PENJAMIN.ID_PENJAMIN, T_PENJAMIN.GOL_PENJAMIN, T_PENJAMIN.NAMA_PENJAMIN, T_PENJAMIN.IDENTITAS_PENJAMIN, T_PENJAMIN.ALAMAT_PENJAMIN, T_KREDIT.KONDISI FROM T_KREDIT INNER JOIN R_DEBITUR_FASILITAS INNER JOIN T_DEBITUR INNER JOIN T_PENJAMIN ON T_DEBITUR.ID_DEBITUR = T_PENJAMIN.ID_DEBITUR ON R_DEBITUR_FASILITAS.ID_DEBITUR = T_DEBITUR.ID_DEBITUR ON T_KREDIT.ID_FASILITAS = R_DEBITUR_FASILITAS.ID_FASILITAS where t_debitur.id_debitur = '" & rsdatapenjamin!cif_bi & "'"
'                Set rspenjamin = New ADODB.Recordset
'                Set rspenjamin = dbsid.Execute(strpenjamin)
'
'                If rspenjamin!kondisi = "02" Then
'                    operation = "D"
'                Else
'                    operation = "S"
'                End If
'
'                id_debitur = rspenjamin!id_debitur
'                id_fasilitas = String(52, " ")
'                id_penjamin = rspenjamin!id_penjamin
'                nama_penjamin = rspenjamin!nama_penjamin
'                GOL_PENJAMIN = rspenjamin!GOL_PENJAMIN
'                bag_jamin = "10000"
'                identitas_penjamin = rspenjamin!identitas_penjamin
'                npwp_penjamin = String(20, " ")
'                alamat_penjamin = rspenjamin!alamat_penjamin
'                prime_bank = "T"
'                din = rspenjamin!din
'
'            Else
'
'                operation = "C"
'
'                id_debitur = rspenjamin!id_debitur
'                id_fasilitas = String(52, " ")
'                id_penjamin = rspenjamin!id_penjamin
'                nama_penjamin = rspenjamin!nama_penjamin
'                GOL_PENJAMIN = rspenjamin!GOL_PENJAMIN
'                bag_jamin = "10000"
'                identitas_penjamin = rspenjamin!identitas_penjamin
'                npwp_penjamin = String(20, " ")
'                alamat_penjamin = rspenjamin!alamat_penjamin
'                prime_bank = "T"
'                din = rspenjamin!din
'
'
'            End If
'
'            alamat_penjamin = alamat_penjamin & String(100 - Len(alamat_penjamin), " ")
'            identitas_penjamin = identitas_penjamin & String(25 - Len(identitas_penjamin), " ")
'            nama_penjamin = nama_penjamin & String(50 - Len(nama_penjamin), " ")
'            id_penjamin = id_penjamin & String(29 - Len(id_penjamin), " ")
'            id_debitur = id_debitur & String(43 - Len(id_debitur), " ")
'            din = din & String(20 - Len(din), " ")
'            waktucreate = Format(Now, "yyyymmddhhnnss")
'            update_date = Format(Now, "yyyymmddhhnnss")
'
'
'        rsdatapenjamin.MoveNext
'        Wend
'
'
'    End If
    
    
    
    ' ktrl lbpr
    
    If Check8 = Checked Then
        
        iddata = "SID050"
        operation = "U"
        strpenempatan = "select ROUND(sum(saldoenc)) as saldo from rekbank where kodesld = 'D'"
        Set rspenempatan = New ADODB.Recordset
        Set rspenempatan = dbbank.Execute(strpenempatan)
        
        penempatan = rspenempatan!saldo
        penempatan = String(18 - Len(penempatan), "0") & penempatan
        
        strkredit = "SELECT SUM(saldoenc) as saldo FROM rekredit WHERE cif_bi <> '-' AND din <> '-'"
        Set rskredit = New ADODB.Recordset
        Set rskredit = dbbank.Execute(strkredit)
        
        
        
        kredit = rskredit!saldo
        kredit = String(18 - Len(kredit), "0") & kredit
        
        cetakkontrol = iddata & operation & idlembaga & idbank & idkancab & blndata & thndata & penempatan & String(18, "0") & kredit & String(18, "0") & String(18, "0") & String(18, "0") & String(18, "0") & String(18, "0") & Format(Now, "yyyymmddhhnnss") & createuser & Format(Now, "yyyymmddhhnnss") & "1"
        
        Print #1, cetakkontrol
    End If
    
    '======================================================================================================
    ' relasi debitur-fasilitas
        
    If Check10 = Checked Then
        
        pgbar.Value = 0
        nomer = 1
        
        'strdatarelasi = "select id_fasilitas, no_rekening, jenis_fasilitas from t_kredit where id_fasilitas not in(select id_deb_fas from r_debitur_fasilitas )"
        strdatarelasi = "select r_debitur_fasilitas.id_deb_fas as id_deb_fas, t_kredit.jenis_fasilitas as jenis_fasilitas , t_debitur.id_debitur as id_debitur, t_kredit.id_fasilitas as id_fasilitas, t_din.din as din, t_kredit.no_rekening as norek, t_kredit.operation as operation,  t_kredit.status_kirim as status_kirim from r_debitur_fasilitas, t_debitur, t_kredit, t_din where r_debitur_fasilitas.id_fasilitas = t_kredit.id_fasilitas and r_debitur_fasilitas.id_debitur = t_debitur.id_debitur and t_din.din = t_debitur.din"
        Set rs = New ADODB.Recordset
        Set rs = dbsid.Execute(strdatarelasi)
        
        pgbar.Max = rs.RecordCount
        
        While Not rs.EOF
            
            DoEvents
            
            pgbar.Value = pgbar.Value + 1
                      
            iddata = "SID070"
            operation = rs!operation
            
            If operation = "C" Then
            
                nasa = cek
            End If
            
            id_deb_fas = rs!id_deb_fas & String(30 - Len(rs!id_deb_fas), " ")
            id_joint_acc = "0"
            jenis = rs!jenis_fasilitas
            ID_Debitur = rs!ID_Debitur & String(43 - Len(rs!ID_Debitur), " ")
            ID_Fasilitas = rs!ID_Fasilitas
            ID_Fasilitas = ID_Fasilitas & String(52 - Len(ID_Fasilitas), " ")
            no_rekening = rs!norek
            no_rekening = no_rekening & String(25 - Len(no_rekening), " ")
            nmpengirim = Left(nmpengirim, 20)
            din = rs!din
            din = din & String(20 - Len(din), " ")
            status_Kirim = rs!status_Kirim
            'datapenempatan = iddata & operation & idlembaga & idbank & idkancab & format(now,"mm") & format(now,"yyyy") & string(30, " ") & id_join_acc & jenis & id_debitur & id_fasilitas & format(now,"yyyymmddhhnnss") & createuser & format(now,""yyyymmddhhnnss)
            datafasilitas = iddata & operation & idlembaga & idbank & idkancab & blndata & thndata & id_deb_fas & id_joint_acc & jenis & ID_Debitur & ID_Fasilitas & Format(Now, "yyyymmddhhnnss") & nmpengirim & Format(Now, "yyyymmddhhnnss") & din & no_rekening & status_Kirim
            
            If Len(datafasilitas) < 248 Then
                nasa = cek
            End If
            
            
            Print #1, datafasilitas
            
           rs.MoveNext
        Wend
        
    End If
    
    
    
    '=======================================================================================================================
    'cetak footer
    
    Print #1, footer
    
    Close #1
        MsgBox ("selesai"), vbInformation, "selesai"
        cmdMulai.Enabled = True

    Text3.Text = bukaFile
    

End Function


Private Sub Form_Load()

    sttbtn = 1
    Check2.Enabled = True

End Sub


Private Sub OsenXPButton1_Click()

End Sub

'Private Sub hpusPenjaminlunas_Click()
'    Dim strhapuspenjamin As String
'    Dim rshapuspenjamin As ADODB.Recordset
'
'    strhapuspenjamin = "select t_kredit.id_fasilitas, t_debitur.id_debitur, t_penjamin.id_penjamin, t_debitur.din from t_kredit, t_debitur, t_penjamin, r_debitur_fasilitas where t_kredit.id_fasilitas = r_debitur_fasilitas.id_fasilitas and t_debitur.id_debitur = r_debitur_fasilitas.id_debitur and t_penjamin.id_debitur = t_debitur.id_debitur and t_kredit.kondisi = '02'"
'    Set rshapuspenjamin = New ADODB.Recordset
'    Set rshapuspenjamin = dbsid.Execute(strhapuspenjamin)
'
'    hpusPenjaminlunas.Enabled = False
'
'    While Not rshapuspenjamin.EOF
'        DoEvents
'
'        dbsid.Execute ("delete from t_penjamin where id_penjamin = '" & rshapuspenjamin!id_penjamin & "'")
'
'    rshapuspenjamin.MoveNext
'    Wend
'
'    MsgBox ("Hapus data penjamin berhasil")
'    hpusPenjaminlunas.Enabled = True
'End Sub



Private Sub Timer1_Timer()
    Select Case FormatJam
    Case DuabelasJam
        StatusBar1.Panels(1).Text = Format(Time, "hh:mm:ss AM/PM")
    Case Else
        StatusBar1.Panels(1).Text = Format(Time, "hh:mm:ss")
End Select
End Sub

Private Sub Command2_Click()
    Dim rsrekredit As ADODB.Recordset
    Dim rskreditsid As ADODB.Recordset
End Sub

Public Function bersihString(ByVal isiText As String) As String
    isiText = Replace(isiText, vbCrLf, "")
    isiText = Replace(isiText, vbCr, "")
    isiText = Replace(isiText, vbLf, "")
    isiText = Trim(isiText)
    bersihString = isiText
End Function

